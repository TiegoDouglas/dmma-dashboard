import unittest

from sync_vagas import SyncError, resolve_state_name


STATES = [
    {"name": "Novo", "category": "Proposed"},
    {"name": "Ativo", "category": "InProgress"},
    {"name": "Resolvido", "category": "Resolved"},
    {"name": "Fechado", "category": "Completed"},
]


class ResolveStateNameTests(unittest.TestCase):
    def test_accepts_exact_case_insensitive_match(self) -> None:
        self.assertEqual(resolve_state_name("aTiVo", STATES), "Ativo")

    def test_maps_em_andamento_to_in_progress_category(self) -> None:
        self.assertEqual(resolve_state_name("Em andamento", STATES), "Ativo")

    def test_maps_new_aliases_to_proposed_category(self) -> None:
        self.assertEqual(resolve_state_name("New", STATES), "Novo")
        self.assertEqual(resolve_state_name("Nova", STATES), "Novo")

    def test_maps_closed_aliases_to_completed_category(self) -> None:
        self.assertEqual(resolve_state_name("Concluído", STATES), "Fechado")
        self.assertEqual(resolve_state_name("Closed", STATES), "Fechado")

    def test_rejects_ambiguous_category_and_lists_valid_states(self) -> None:
        states = STATES + [{"name": "Em execução", "category": "InProgress"}]

        with self.assertRaisesRegex(
            SyncError,
            r"ambíguo.*Estados válidos: Ativo, Em execução, Fechado, Novo, Resolvido",
        ):
            resolve_state_name("Em andamento", states)

    def test_rejects_unknown_state_and_lists_valid_states(self) -> None:
        with self.assertRaisesRegex(
            SyncError,
            r"não é suportado.*Estados válidos: Ativo, Fechado, Novo, Resolvido",
        ):
            resolve_state_name("Bloqueado", STATES)


if __name__ == "__main__":
    unittest.main()
