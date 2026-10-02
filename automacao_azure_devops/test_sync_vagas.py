import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from sync_vagas import (
    AzureDevOpsClient,
    ExistingWorkItem,
    FieldDefinition,
    ItemRecord,
    PARENT_RELATION_TYPE,
    RecordPair,
    SyncError,
    WorkbookPlan,
    _values_equal,
    parse_field_overrides,
    read_workbook_plan,
    synchronize,
)


HEADERS = [
    ("Position", "Id MyScheduling"),
    ("Position", "RoleTitle"),
    ("Project", "OppID"),
    ("Project", "PROJETO"),
    ("Project", "Estimated Close Date"),
]


def create_workbook(rows: list[list[object]]) -> Path:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Planilha1"
    worksheet.append(["Tipo de Item", *(item_type for item_type, _ in HEADERS)])
    worksheet.append(["Campos", *(field_name for _, field_name in HEADERS)])
    for row in rows:
        worksheet.append(["valor", *row])
    handle = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    handle.close()
    path = Path(handle.name)
    workbook.save(path)
    workbook.close()
    return path


class WorkbookPlanTests(unittest.TestCase):
    def tearDown(self) -> None:
        for path in getattr(self, "paths", []):
            path.unlink(missing_ok=True)

    def workbook(self, rows: list[list[object]]) -> Path:
        path = create_workbook(rows)
        self.paths = getattr(self, "paths", []) + [path]
        return path

    def test_groups_fields_and_preserves_types(self) -> None:
        close_date = datetime(2026, 4, 20)
        plan = read_workbook_plan(
            self.workbook([[6439485, "Dev", "A0001", "Projeto", close_date]])
        )

        self.assertEqual(plan.projects[0].fields["OppID"], "A0001")
        self.assertEqual(
            plan.projects[0].fields["Estimated Close Date"], close_date
        )
        self.assertEqual(plan.positions[0].fields["Id MyScheduling"], 6439485)
        self.assertEqual(plan.pairs[0].project, plan.projects[0])

    def test_deduplicates_identical_project_across_positions(self) -> None:
        plan = read_workbook_plan(
            self.workbook(
                [
                    [1, "Dev 1", "A0001", "Projeto", None],
                    [2, "Dev 2", "A0001", "Projeto", None],
                ]
            )
        )

        self.assertEqual(len(plan.projects), 1)
        self.assertEqual(len(plan.positions), 2)

    def test_rejects_conflicting_duplicate_project(self) -> None:
        path = self.workbook(
            [
                [1, "Dev 1", "A0001", "Projeto A", None],
                [2, "Dev 2", "A0001", "Projeto B", None],
            ]
        )

        with self.assertRaisesRegex(SyncError, "valores conflitantes"):
            read_workbook_plan(path)

    def test_supports_distinct_projects_and_positions(self) -> None:
        plan = read_workbook_plan(
            self.workbook(
                [
                    [1, "Dev 1", "A0001", "Projeto A", None],
                    [2, "Dev 2", "A0002", "Projeto B", None],
                ]
            )
        )

        self.assertEqual(len(plan.projects), 2)
        self.assertEqual(len(plan.positions), 2)
        self.assertEqual(len(plan.pairs), 2)

    def test_rejects_same_position_linked_to_different_projects(self) -> None:
        path = self.workbook(
            [
                [1, "Dev", "A0001", "Projeto A", None],
                [1, "Dev", "A0002", "Projeto B", None],
            ]
        )

        with self.assertRaisesRegex(
            SyncError, "mesma Position.*Projects diferentes"
        ):
            read_workbook_plan(path)

    def test_rejects_conflicting_duplicate_position(self) -> None:
        path = self.workbook(
            [
                [1, "Dev A", "A0001", "Projeto", None],
                [1, "Dev B", "A0001", "Projeto", None],
            ]
        )

        with self.assertRaisesRegex(SyncError, "valores conflitantes.*Position"):
            read_workbook_plan(path)

    def test_rejects_formula(self) -> None:
        path = self.workbook([[1, "=1+1", "A0001", "Projeto", None]])

        with self.assertRaisesRegex(SyncError, "Fórmulas não são permitidas"):
            read_workbook_plan(path)

    def test_requires_both_natural_keys(self) -> None:
        path = self.workbook([[None, "Dev", "A0001", "Projeto", None]])

        with self.assertRaisesRegex(SyncError, "Id MyScheduling"):
            read_workbook_plan(path)


class FieldResolutionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = AzureDevOpsClient("org", "project", "not-a-real-token")
        self.definitions = [
            FieldDefinition("Id MyScheduling", "Custom.Id", "integer", False),
            FieldDefinition("RoleTitle", "Custom.RoleTitle", "string", False),
            FieldDefinition("Title", "System.Title", "string", False),
        ]

    def test_resolves_fields_and_adds_system_title(self) -> None:
        record = ItemRecord(
            "Position",
            3,
            {"Id MyScheduling": 123, "RoleTitle": "Developer"},
        )

        resolved = self.client.resolve_record(record, self.definitions)

        self.assertEqual(resolved.natural_key_reference_name, "Custom.Id")
        self.assertEqual(resolved.fields["System.Title"], "Developer")

    def test_uses_explicit_default_business_mapping(self) -> None:
        definitions = [
            FieldDefinition("Opp ID_", "Custom.OppID_", "string", False),
            FieldDefinition("Account", "Custom.Account", "string", False),
            FieldDefinition("Title", "System.Title", "string", False),
            FieldDefinition("Win Prob", "Custom.WinProb", "double", False),
            FieldDefinition(
                "Start Date",
                "Microsoft.VSTS.Scheduling.StartDate",
                "dateTime",
                False,
            ),
            FieldDefinition(
                "Finish Date",
                "Microsoft.VSTS.Scheduling.FinishDate",
                "dateTime",
                False,
            ),
        ]
        record = ItemRecord(
            "Project",
            3,
            {
                "OppID": "A1",
                "Account.": "Cliente",
                "PROJETO": "Projeto",
                "Win Probability": 100,
                "Estimated Project Start Date": datetime(2026, 4, 1),
                "Estimated Project End Date": datetime(2027, 3, 2),
            },
        )

        resolved = self.client.resolve_record(record, definitions)

        self.assertEqual(resolved.fields["System.Title"], "Projeto")
        self.assertEqual(resolved.fields["Custom.Account"], "Cliente")
        self.assertEqual(resolved.fields["Custom.WinProb"], 100)
        self.assertEqual(
            resolved.natural_key_reference_name, "Custom.OppID_"
        )

    def test_environment_override_precedes_default_business_mapping(self) -> None:
        client = AzureDevOpsClient(
            "org",
            "project",
            "token",
            {"Project.PROJETO": "Custom.ProjectName"},
        )
        definitions = [
            FieldDefinition("Opp ID_", "Custom.OppID_", "string", False),
            FieldDefinition("Title", "System.Title", "string", False),
            FieldDefinition("Project Name", "Custom.ProjectName", "string", False),
        ]
        record = ItemRecord(
            "Project",
            3,
            {"OppID": "A1", "PROJETO": "Projeto"},
        )

        resolved = client.resolve_record(record, definitions)

        self.assertEqual(resolved.fields["Custom.ProjectName"], "Projeto")
        self.assertEqual(resolved.fields["System.Title"], "Projeto")

    def test_normalizes_spaces_punctuation_case_and_accents(self) -> None:
        definitions = self.definitions + [
            FieldDefinition("Opp ID", "Custom.OpportunityId", "string", False),
            FieldDefinition("Prática", "Custom.Practice", "string", False),
        ]
        project = ItemRecord(
            "Project",
            3,
            {"OppID": "A1", "PROJETO": "Projeto"},
        )
        project_definitions = definitions + [
            FieldDefinition("PROJETO", "Custom.Project", "string", False),
        ]
        position = ItemRecord(
            "Position",
            3,
            {
                "Id MyScheduling": 123,
                "RoleTitle": "Developer",
                "PRATICA": "Applications",
            },
        )

        resolved_project = self.client.resolve_record(project, project_definitions)
        resolved_position = self.client.resolve_record(position, definitions)

        self.assertEqual(
            resolved_project.natural_key_reference_name, "Custom.OpportunityId"
        )
        self.assertEqual(
            resolved_position.fields["Custom.Practice"], "Applications"
        )

    def test_matches_reference_name_leaf_after_canonical_normalization(self) -> None:
        definitions = self.definitions + [
            FieldDefinition(
                "Competências da vaga", "Custom.Role-Skills", "string", False
            )
        ]
        record = ItemRecord(
            "Position",
            3,
            {
                "Id MyScheduling": 123,
                "RoleTitle": "Developer",
                "role_skills": "Python",
            },
        )

        resolved = self.client.resolve_record(record, definitions)

        self.assertEqual(resolved.fields["Custom.Role-Skills"], "Python")

    def test_rejects_canonical_ambiguity_and_lists_diagnostics(self) -> None:
        definitions = self.definitions + [
            FieldDefinition("Opp ID", "Custom.OppIdOne", "string", False),
            FieldDefinition("Opp-ID", "Custom.OppIdTwo", "string", False),
            FieldDefinition("Somente leitura", "Custom.ReadOnly", "string", True),
        ]
        project = ItemRecord(
            "Project",
            3,
            {"OppID": "A1", "PROJETO": "Projeto"},
        )
        definitions.append(
            FieldDefinition("PROJETO", "Custom.Project", "string", False)
        )

        with self.assertRaises(SyncError) as raised:
            self.client.resolve_record(project, definitions)

        message = str(raised.exception)
        self.assertIn("ambíguo após normalização canônica", message)
        self.assertIn("Opp ID (Custom.OppIdOne)", message)
        self.assertIn("Opp-ID (Custom.OppIdTwo)", message)
        self.assertIn("Campos graváveis disponíveis", message)
        self.assertNotIn("Custom.ReadOnly", message)

    def test_rejects_unknown_field(self) -> None:
        record = ItemRecord(
            "Position",
            3,
            {"Id MyScheduling": 123, "RoleTitle": "Developer", "Unknown": "x"},
        )

        with self.assertRaises(SyncError) as raised:
            self.client.resolve_record(record, self.definitions)

        message = str(raised.exception)
        self.assertIn("Unknown", message)
        self.assertIn("não foi encontrado", message)
        self.assertIn("Campos graváveis disponíveis", message)
        self.assertIn("RoleTitle (Custom.RoleTitle)", message)
        self.assertEqual(message.count("Campos graváveis disponíveis"), 1)

    def test_reports_all_unresolved_fields_in_one_error(self) -> None:
        record = ItemRecord(
            "Position",
            3,
            {
                "Id MyScheduling": 123,
                "RoleTitle": "Developer",
                "Unknown One": "x",
                "Unknown Two": "y",
            },
        )
        definitions = self.definitions + [
            FieldDefinition("Title", "System.Title", "string", False)
        ]

        with self.assertRaises(SyncError) as raised:
            self.client.resolve_record(record, definitions)

        message = str(raised.exception)
        self.assertIn("Unknown One", message)
        self.assertIn("Unknown Two", message)
        self.assertIn("Falha ao validar os campos de Position", message)

    def test_uses_explicit_reference_override(self) -> None:
        client = AzureDevOpsClient(
            "org",
            "project",
            "token",
            {"Position.RoleTitle": "Custom.RoleTitle"},
        )
        duplicated = self.definitions + [
            FieldDefinition("RoleTitle", "Other.RoleTitle", "string", False)
        ]
        record = ItemRecord(
            "Position",
            3,
            {"Id MyScheduling": 123, "RoleTitle": "Developer"},
        )

        resolved = client.resolve_record(record, duplicated)

        self.assertIn("Custom.RoleTitle", resolved.fields)


class RelationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = AzureDevOpsClient("org", "project", "token")
        self.requests: list[dict[str, object]] = []

        def request_json(method, url, action, **kwargs):
            self.requests.append(
                {"method": method, "url": url, "action": action, **kwargs}
            )
            return {"id": 20, "fields": {}}

        self.client._request_json = request_json  # type: ignore[method-assign]

    def test_existing_parent_relation_is_idempotent(self) -> None:
        relations = (
            {
                "rel": PARENT_RELATION_TYPE,
                "url": "https://dev.azure.com/org/project/_apis/wit/workItems/10",
            },
        )

        changed = self.client.ensure_parent_relation(20, 10, relations)

        self.assertFalse(changed)
        self.assertEqual(self.requests, [])

    def test_adds_parent_relation_to_child(self) -> None:
        changed = self.client.ensure_parent_relation(20, 10, ())

        self.assertTrue(changed)
        patch = self.requests[0]["json"]
        self.assertEqual(patch[0]["path"], "/relations/-")
        self.assertEqual(patch[0]["value"]["rel"], PARENT_RELATION_TYPE)

    def test_rejects_reparenting(self) -> None:
        relations = (
            {
                "rel": PARENT_RELATION_TYPE,
                "url": "https://dev.azure.com/org/project/_apis/wit/workItems/99",
            },
        )

        with self.assertRaisesRegex(SyncError, "já possui outro pai"):
            self.client.ensure_parent_relation(20, 10, relations)


class UtilityTests(unittest.TestCase):
    def test_compares_azure_datetime_by_date(self) -> None:
        self.assertTrue(
            _values_equal("2026-04-20T00:00:00Z", datetime(2026, 4, 20))
        )

    def test_parses_field_overrides(self) -> None:
        self.assertEqual(
            parse_field_overrides('{"Position.Skillls": "Custom.Skills"}'),
            {"Position.Skillls": "Custom.Skills"},
        )

    def test_rejects_invalid_field_overrides(self) -> None:
        with self.assertRaisesRegex(SyncError, "objeto JSON válido"):
            parse_field_overrides("not-json")


class SynchronizeTests(unittest.TestCase):
    @staticmethod
    def definitions():
        return [
            FieldDefinition("Opp ID_", "Custom.OppID_", "string", False),
            FieldDefinition("Title", "System.Title", "string", False),
            FieldDefinition(
                "Id MyScheduling", "Custom.SchedulingId", "integer", False
            ),
        ]

    @staticmethod
    def multi_line_plan():
        project = ItemRecord("Project", 3, {"OppID": "A1", "PROJETO": "Projeto"})
        position_one = ItemRecord(
            "Position", 3, {"Id MyScheduling": 10, "RoleTitle": "Developer 1"}
        )
        position_two = ItemRecord(
            "Position", 4, {"Id MyScheduling": 20, "RoleTitle": "Developer 2"}
        )
        return WorkbookPlan(
            (
                RecordPair(project, position_one),
                RecordPair(project, position_two),
            )
        )

    def test_reuses_one_project_for_multiple_positions(self) -> None:
        plan = self.multi_line_plan()
        events: list[str] = []
        next_ids = iter((100, 200, 201))

        class CreatingClient:
            def get_field_definitions(self, item_type):
                return SynchronizeTests.definitions()

            def resolve_record(self, record, definitions):
                return AzureDevOpsClient(
                    "org", "project", "token"
                ).resolve_record(record, definitions)

            def find_existing(self, record):
                events.append(f"find:{record.source.item_type}")
                return None

            def create_work_item(self, record):
                work_item_id = next(next_ids)
                events.append(f"create:{record.source.item_type}:{work_item_id}")
                return work_item_id

            def update_work_item(self, existing, record):
                events.append("unexpected-update")
                return True

            def ensure_parent_relation(self, child_id, parent_id, relations):
                events.append(f"relation:{parent_id}->{child_id}")
                return True

        synchronize(CreatingClient(), plan)

        self.assertEqual(events.count("create:Project:100"), 1)
        self.assertEqual(
            [event for event in events if event.startswith("create:Position")],
            ["create:Position:200", "create:Position:201"],
        )
        self.assertIn("relation:100->200", events)
        self.assertIn("relation:100->201", events)

    def test_creates_distinct_projects_for_distinct_opportunity_ids(self) -> None:
        project_one = ItemRecord(
            "Project", 3, {"OppID": "A1", "PROJETO": "Projeto 1"}
        )
        project_two = ItemRecord(
            "Project", 4, {"OppID": "A2", "PROJETO": "Projeto 2"}
        )
        position_one = ItemRecord(
            "Position", 3, {"Id MyScheduling": 10, "RoleTitle": "Developer 1"}
        )
        position_two = ItemRecord(
            "Position", 4, {"Id MyScheduling": 20, "RoleTitle": "Developer 2"}
        )
        plan = WorkbookPlan(
            (
                RecordPair(project_one, position_one),
                RecordPair(project_two, position_two),
            )
        )
        next_ids = iter((100, 101, 200, 201))
        events: list[str] = []

        class CreatingClient:
            def get_field_definitions(self, item_type):
                return SynchronizeTests.definitions()

            def resolve_record(self, record, definitions):
                return AzureDevOpsClient(
                    "org", "project", "token"
                ).resolve_record(record, definitions)

            def find_existing(self, record):
                return None

            def create_work_item(self, record):
                work_item_id = next(next_ids)
                events.append(f"create:{record.source.item_type}:{work_item_id}")
                return work_item_id

            def update_work_item(self, existing, record):
                return False

            def ensure_parent_relation(self, child_id, parent_id, relations):
                events.append(f"relation:{parent_id}->{child_id}")
                return True

        synchronize(CreatingClient(), plan)

        self.assertIn("create:Project:100", events)
        self.assertIn("create:Project:101", events)
        self.assertIn("relation:100->200", events)
        self.assertIn("relation:101->201", events)

    def test_duplicate_identical_rows_create_relation_once(self) -> None:
        pair = self.multi_line_plan().pairs[0]
        plan = WorkbookPlan((pair, pair))
        relation_events: list[tuple[int, int]] = []
        next_ids = iter((100, 200))

        class CreatingClient:
            def get_field_definitions(self, item_type):
                return SynchronizeTests.definitions()

            def resolve_record(self, record, definitions):
                return AzureDevOpsClient(
                    "org", "project", "token"
                ).resolve_record(record, definitions)

            def find_existing(self, record):
                return None

            def create_work_item(self, record):
                return next(next_ids)

            def update_work_item(self, existing, record):
                return False

            def ensure_parent_relation(self, child_id, parent_id, relations):
                relation_events.append((parent_id, child_id))
                return True

        synchronize(CreatingClient(), plan)

        self.assertEqual(relation_events, [(100, 200)])

    def test_reexecution_is_idempotent_for_items_and_relations(self) -> None:
        plan = self.multi_line_plan()
        events: list[str] = []
        ids = {
            ("Project", "A1"): 100,
            ("Position", "10"): 200,
            ("Position", "20"): 201,
        }

        class ExistingClient:
            def get_field_definitions(self, item_type):
                return SynchronizeTests.definitions()

            def resolve_record(self, record, definitions):
                return AzureDevOpsClient(
                    "org", "project", "token"
                ).resolve_record(record, definitions)

            def find_existing(self, record):
                work_item_id = ids[
                    (record.source.item_type, str(record.source.natural_key_value))
                ]
                fields = {
                    reference: value
                    for reference, value in record.fields.items()
                }
                relations = ()
                if record.source.item_type == "Position":
                    relations = (
                        {
                            "rel": PARENT_RELATION_TYPE,
                            "url": (
                                "https://dev.azure.com/org/project/_apis/wit/"
                                "workItems/100"
                            ),
                        },
                    )
                return ExistingWorkItem(work_item_id, fields, relations)

            def update_work_item(self, existing, record):
                events.append(f"update-check:{existing.work_item_id}")
                return False

            def create_work_item(self, record):
                events.append("unexpected-create")
                return 999

            def ensure_parent_relation(self, child_id, parent_id, relations):
                changed = AzureDevOpsClient(
                    "org", "project", "token"
                ).ensure_parent_relation(child_id, parent_id, relations)
                events.append(f"relation-check:{parent_id}->{child_id}:{changed}")
                return changed

        synchronize(ExistingClient(), plan)

        self.assertNotIn("unexpected-create", events)
        self.assertEqual(
            [event for event in events if event.startswith("update-check")],
            ["update-check:100", "update-check:200", "update-check:201"],
        )
        self.assertIn("relation-check:100->200:False", events)
        self.assertIn("relation-check:100->201:False", events)

    def test_aggregates_project_and_position_errors_before_searches(self) -> None:
        project = ItemRecord(
            "Project", 3, {"OppID": "A1", "PROJETO": "Projeto"}
        )
        position = ItemRecord(
            "Position", 3, {"Id MyScheduling": 10, "RoleTitle": "Developer"}
        )
        plan = WorkbookPlan((RecordPair(project, position),))
        events: list[str] = []

        class InvalidClient:
            def get_field_definitions(self, item_type):
                return []

            def resolve_record(self, record, definitions):
                events.append(f"resolve:{record.item_type}")
                raise SyncError(f"erro de {record.item_type}")

            def find_existing(self, record):
                events.append("unexpected-search")

            def create_work_item(self, record):
                events.append("unexpected-create")

            def update_work_item(self, existing, record):
                events.append("unexpected-update")

            def ensure_parent_relation(self, child_id, parent_id, relations):
                events.append("unexpected-relation")

        with self.assertRaises(SyncError) as raised:
            synchronize(InvalidClient(), plan)

        message = str(raised.exception)
        self.assertIn("erro de Project", message)
        self.assertIn("erro de Position", message)
        self.assertIn("nenhuma busca ou escrita", message)
        self.assertEqual(events, ["resolve:Project", "resolve:Position"])

    def test_completes_all_lookups_before_writes_and_links_items(self) -> None:
        project = ItemRecord(
            "Project", 3, {"OppID": "A1", "PROJETO": "Projeto"}
        )
        position = ItemRecord(
            "Position", 3, {"Id MyScheduling": 10, "RoleTitle": "Developer"}
        )
        plan = WorkbookPlan((RecordPair(project, position),))
        events: list[str] = []

        class FakeClient:
            def get_field_definitions(self, item_type):
                return [
                    FieldDefinition("OppID", "Custom.OppID", "string", False),
                    FieldDefinition("PROJETO", "Custom.Project", "string", False),
                    FieldDefinition("Id MyScheduling", "Custom.SchedulingId", "integer", False),
                    FieldDefinition("RoleTitle", "Custom.RoleTitle", "string", False),
                    FieldDefinition("Title", "System.Title", "string", False),
                ]

            def resolve_record(self, record, definitions):
                return AzureDevOpsClient(
                    "org", "project", "token"
                ).resolve_record(record, definitions)

            def find_existing(self, record):
                events.append(f"find:{record.source.item_type}")
                if record.source.item_type == "Position":
                    return ExistingWorkItem(
                        20,
                        {
                            "Custom.SchedulingId": 10,
                            "Custom.RoleTitle": "Developer",
                            "System.Title": "Developer",
                        },
                        (),
                    )
                return None

            def create_work_item(self, record):
                events.append(f"create:{record.source.item_type}")
                return 10

            def update_work_item(self, existing, record):
                events.append(f"update:{record.source.item_type}")
                return False

            def ensure_parent_relation(self, child_id, parent_id, relations):
                events.append(f"relation:{parent_id}->{child_id}")
                return True

        synchronize(FakeClient(), plan)

        self.assertEqual(
            events,
            [
                "find:Project",
                "find:Position",
                "create:Project",
                "update:Position",
                "relation:10->20",
            ],
        )


if __name__ == "__main__":
    unittest.main()
