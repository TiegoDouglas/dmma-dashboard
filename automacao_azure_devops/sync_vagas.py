#!/usr/bin/env python3
"""Synchronize the single row in vagas.xlsx with an Azure DevOps work item."""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import quote

import requests
from openpyxl import load_workbook
from requests.auth import HTTPBasicAuth

EXPECTED_HEADERS = ("Description", "state", "Skills")
DEFAULT_ORGANIZATION = "AccountMSFT"
DEFAULT_PROJECT = "Esteira de Vagas"
DEFAULT_WORK_ITEM_ID = 3133
API_VERSION = "7.1"
REQUEST_TIMEOUT_SECONDS = 30


class SyncError(RuntimeError):
    """Raised when workbook validation or Azure DevOps synchronization fails."""


@dataclass(frozen=True)
class Vacancy:
    description: str
    state: str
    skills: str


def _cell_text(value: Any, field_name: str) -> str:
    if value is None:
        return ""
    if not isinstance(value, (str, int, float, bool)):
        raise SyncError(f"O campo {field_name!r} possui um tipo de valor não suportado.")
    return str(value).strip()


def read_vacancy(workbook_path: Path) -> Vacancy:
    if not workbook_path.is_file():
        raise SyncError(f"Planilha não encontrada: {workbook_path}")

    try:
        workbook = load_workbook(
            workbook_path,
            read_only=True,
            data_only=False,
            keep_links=False,
        )
    except Exception as exc:
        raise SyncError(f"Não foi possível abrir a planilha XLSX: {exc}") from exc

    try:
        if "Planilha1" not in workbook.sheetnames:
            raise SyncError("A planilha deve conter a aba 'Planilha1'.")

        worksheet = workbook["Planilha1"]
        headers = tuple(
            _cell_text(worksheet.cell(row=1, column=column).value, f"cabeçalho {column}")
            for column in range(1, 4)
        )
        if headers != EXPECTED_HEADERS:
            raise SyncError(
                "Cabeçalhos inválidos em Planilha1. "
                f"Esperado: {EXPECTED_HEADERS}; encontrado: {headers}."
            )

        extra_headers = [
            worksheet.cell(row=1, column=column).value
            for column in range(4, worksheet.max_column + 1)
            if worksheet.cell(row=1, column=column).value not in (None, "")
        ]
        if extra_headers:
            raise SyncError("Planilha1 possui cabeçalhos adicionais não permitidos.")

        populated_rows: list[tuple[int, tuple[str, str, str]]] = []
        for row_number in range(2, worksheet.max_row + 1):
            cells = tuple(worksheet.cell(row=row_number, column=column) for column in range(1, 4))
            if any(cell.data_type == "f" for cell in cells):
                raise SyncError(f"Fórmulas não são permitidas na linha {row_number}.")

            values = tuple(
                _cell_text(cell.value, EXPECTED_HEADERS[index])
                for index, cell in enumerate(cells)
            )
            extra_values = [
                worksheet.cell(row=row_number, column=column).value
                for column in range(4, worksheet.max_column + 1)
                if worksheet.cell(row=row_number, column=column).value not in (None, "")
            ]
            if extra_values:
                raise SyncError(
                    f"A linha {row_number} possui valores fora das três colunas esperadas."
                )
            if any(values):
                populated_rows.append((row_number, values))

        if len(populated_rows) != 1:
            raise SyncError(
                "Planilha1 deve conter exatamente uma linha de dados preenchida; "
                f"foram encontradas {len(populated_rows)}."
            )

        row_number, values = populated_rows[0]
        missing_fields = [
            EXPECTED_HEADERS[index] for index, value in enumerate(values) if not value
        ]
        if missing_fields:
            raise SyncError(
                f"A linha {row_number} possui campos obrigatórios vazios: "
                + ", ".join(missing_fields)
            )

        return Vacancy(description=values[0], state=values[1], skills=values[2])
    finally:
        workbook.close()


class AzureDevOpsClient:
    def __init__(self, organization: str, project: str, pat: str) -> None:
        if not pat:
            raise SyncError(
                "A variável de ambiente ADO_PAT é obrigatória. "
                "Configure-a somente como secret do GitHub Actions."
            )

        self.base_url = (
            f"https://dev.azure.com/{quote(organization, safe='')}/"
            f"{quote(project, safe='')}/_apis/wit"
        )
        self.session = requests.Session()
        self.session.auth = HTTPBasicAuth("", pat)
        self.session.headers.update({"Accept": "application/json"})

    def _response_json(self, response: requests.Response, action: str) -> dict[str, Any]:
        if not response.ok:
            message = response.text.strip().replace("\n", " ")[:500]
            raise SyncError(
                f"Azure DevOps retornou HTTP {response.status_code} ao {action}: {message}"
            )
        try:
            payload = response.json()
        except requests.JSONDecodeError as exc:
            raise SyncError(f"Resposta inválida do Azure DevOps ao {action}.") from exc
        if not isinstance(payload, dict):
            raise SyncError(f"Resposta inesperada do Azure DevOps ao {action}.")
        return payload

    def resolve_skills_reference_name(self, override: str | None) -> str:
        if override:
            return override.strip()

        try:
            response = self.session.get(
                f"{self.base_url}/fields",
                params={"api-version": API_VERSION},
                timeout=REQUEST_TIMEOUT_SECONDS,
            )
        except requests.RequestException as exc:
            raise SyncError(f"Falha ao consultar os campos do Azure DevOps: {exc}") from exc

        payload = self._response_json(response, "consultar os campos")
        fields = payload.get("value")
        if not isinstance(fields, list):
            raise SyncError("A resposta de campos do Azure DevOps não contém a lista 'value'.")

        matches = [
            field.get("referenceName")
            for field in fields
            if isinstance(field, dict)
            and str(field.get("name", "")).casefold() == "skills".casefold()
            and isinstance(field.get("referenceName"), str)
            and field["referenceName"].strip()
        ]
        unique_matches = sorted(set(matches))
        if len(unique_matches) != 1:
            raise SyncError(
                "Não foi possível identificar unicamente o referenceName do campo exibido "
                "como 'Skills'. Defina AZURE_SKILLS_FIELD_REFERENCE_NAME."
            )
        return unique_matches[0]

    def update_work_item(
        self,
        work_item_id: int,
        vacancy: Vacancy,
        skills_reference_name: str,
    ) -> None:
        fields = (
            ("System.Description", vacancy.description),
            ("System.State", vacancy.state),
            (skills_reference_name, vacancy.skills),
        )
        patch = [
            {
                "op": "add",
                "path": f"/fields/{_escape_json_pointer(reference_name)}",
                "value": value,
            }
            for reference_name, value in fields
        ]

        try:
            response = self.session.patch(
                f"{self.base_url}/workitems/{work_item_id}",
                params={"api-version": API_VERSION},
                headers={"Content-Type": "application/json-patch+json"},
                json=patch,
                timeout=REQUEST_TIMEOUT_SECONDS,
            )
        except requests.RequestException as exc:
            raise SyncError(f"Falha ao atualizar o work item {work_item_id}: {exc}") from exc

        self._response_json(response, f"atualizar o work item {work_item_id}")


def _escape_json_pointer(value: str) -> str:
    return value.replace("~", "~0").replace("/", "~1")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sincroniza automacao_azure_devops/vagas.xlsx com um work item."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path(__file__).with_name("vagas.xlsx"),
        help="Caminho da planilha XLSX.",
    )
    parser.add_argument("--organization", default=DEFAULT_ORGANIZATION)
    parser.add_argument("--project", default=DEFAULT_PROJECT)
    parser.add_argument("--work-item-id", type=int, default=DEFAULT_WORK_ITEM_ID)
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Valida e exibe um resumo sem acessar ou alterar o Azure DevOps.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        vacancy = read_vacancy(args.workbook)
        skills_override = os.getenv("AZURE_SKILLS_FIELD_REFERENCE_NAME", "").strip() or None

        if args.dry_run:
            print("Dry-run concluído: planilha válida; nenhuma chamada ao Azure DevOps foi feita.")
            print(f"Work item: {args.organization}/{args.project}#{args.work_item_id}")
            print("Campos validados: System.Description, System.State e Skills")
            if skills_override:
                print(f"ReferenceName de Skills configurado: {skills_override}")
            return 0

        client = AzureDevOpsClient(
            organization=args.organization,
            project=args.project,
            pat=os.getenv("ADO_PAT", ""),
        )
        skills_reference_name = client.resolve_skills_reference_name(skills_override)
        client.update_work_item(
            work_item_id=args.work_item_id,
            vacancy=vacancy,
            skills_reference_name=skills_reference_name,
        )
        print(
            f"Work item {args.work_item_id} atualizado com sucesso "
            f"(campo Skills: {skills_reference_name})."
        )
        return 0
    except SyncError as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
