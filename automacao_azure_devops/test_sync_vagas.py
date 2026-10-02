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

    def test_rejects_unknown_field(self) -> None:
        record = ItemRecord(
            "Position",
            3,
            {"Id MyScheduling": 123, "RoleTitle": "Developer", "Unknown": "x"},
        )

        with self.assertRaisesRegex(SyncError, "Unknown.*não foi encontrado"):
            self.client.resolve_record(record, self.definitions)

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
