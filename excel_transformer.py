"""Enterprise-grade Excel transformation utility.

This module extracts rows from a `Data` sheet, applies transformation rules from a
`Criteria` sheet, and writes records to a `Result` sheet in a target format.

Expected workbook sheets
------------------------
- Data: Input records.
- Criteria: Rule configuration table.
- Result (optional): Template sheet for desired output column order.

Criteria schema (case-insensitive)
----------------------------------
Required:
- target_column: Output column name.

Optional:
- rule_order: Integer priority (ascending).
- source_column: Source column name or comma-separated names.
- operation: COPY | STATIC | CONCAT | UPPER | LOWER | TITLE | DATE_FORMAT
- operation_arg: Extra argument for operation
    - CONCAT: separator string (default: " ")
    - DATE_FORMAT: strftime format (default: "%Y-%m-%d")
- default_value: Fallback value when result is empty/null.
- required: true/false. Raises validation error if output missing.
- condition: Boolean expression evaluated per row.
    Example: "Country == 'US' and Score >= 75"
"""

from __future__ import annotations

import argparse
import ast
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd

LOGGER = logging.getLogger("excel_transformer")


class TransformationError(Exception):
    """Base exception for transformation failures."""


class WorkbookValidationError(TransformationError):
    """Raised when workbook or sheet schema is invalid."""


@dataclass(frozen=True)
class Rule:
    """A normalized rule loaded from the Criteria sheet."""

    rule_order: int
    target_column: str
    source_columns: List[str]
    operation: str
    operation_arg: Optional[str]
    default_value: Any
    required: bool
    condition: Optional[str]


class SafeExpressionEvaluator(ast.NodeVisitor):
    """Safe evaluator for basic boolean criteria conditions."""

    ALLOWED_NODES = {
        ast.Expression,
        ast.BoolOp,
        ast.And,
        ast.Or,
        ast.Compare,
        ast.Name,
        ast.Load,
        ast.Constant,
        ast.UnaryOp,
        ast.Not,
        ast.Eq,
        ast.NotEq,
        ast.Gt,
        ast.GtE,
        ast.Lt,
        ast.LtE,
        ast.In,
        ast.NotIn,
        ast.Is,
        ast.IsNot,
        ast.BinOp,
        ast.Add,
        ast.Sub,
        ast.Mult,
        ast.Div,
        ast.Mod,
    }

    def evaluate(self, expression: str, context: Dict[str, Any]) -> bool:
        try:
            tree = ast.parse(expression, mode="eval")
            self.visit(tree)
            result = eval(compile(tree, "<condition>", "eval"), {"__builtins__": {}}, context)
            return bool(result)
        except Exception as exc:  # noqa: BLE001
            raise TransformationError(f"Invalid condition expression '{expression}': {exc}") from exc

    def generic_visit(self, node: ast.AST) -> None:
        if type(node) not in self.ALLOWED_NODES:  # noqa: E721
            raise TransformationError(f"Unsupported token in condition: {type(node).__name__}")
        super().generic_visit(node)


class ExcelTransformationService:
    """Service layer for workbook transformation."""

    REQUIRED_SHEETS = {"Data", "Criteria"}
    OPTIONAL_TEMPLATE_SHEET = "Result"

    def __init__(self) -> None:
        self._condition_evaluator = SafeExpressionEvaluator()

    def transform_workbook(self, input_path: Path, output_path: Path) -> pd.DataFrame:
        LOGGER.info("Loading workbook: %s", input_path)
        excel_file = pd.ExcelFile(input_path)
        self._validate_workbook_sheets(excel_file.sheet_names)

        data_df = pd.read_excel(excel_file, sheet_name="Data")
        criteria_df = pd.read_excel(excel_file, sheet_name="Criteria")
        result_template_df = (
            pd.read_excel(excel_file, sheet_name=self.OPTIONAL_TEMPLATE_SHEET)
            if self.OPTIONAL_TEMPLATE_SHEET in excel_file.sheet_names
            else None
        )

        rules = self._load_rules(criteria_df)
        result_df = self._apply_rules(data_df, rules, result_template_df)

        self._write_result_sheet(input_path, output_path, result_df)
        LOGGER.info("Transformation complete. Output written to: %s", output_path)
        return result_df

    def _validate_workbook_sheets(self, sheet_names: Iterable[str]) -> None:
        missing = self.REQUIRED_SHEETS - set(sheet_names)
        if missing:
            raise WorkbookValidationError(f"Missing required sheet(s): {', '.join(sorted(missing))}")

    def _load_rules(self, criteria_df: pd.DataFrame) -> List[Rule]:
        criteria_df = criteria_df.rename(columns=lambda c: str(c).strip().lower())

        if "target_column" not in criteria_df.columns:
            raise WorkbookValidationError("Criteria sheet must include 'target_column' column")

        defaults = {
            "rule_order": 1000,
            "source_column": "",
            "operation": "COPY",
            "operation_arg": None,
            "default_value": None,
            "required": False,
            "condition": None,
        }

        for col, default_value in defaults.items():
            if col not in criteria_df.columns:
                criteria_df[col] = default_value

        rules: List[Rule] = []
        for idx, row in criteria_df.iterrows():
            target_column = str(row["target_column"]).strip()
            if not target_column:
                raise WorkbookValidationError(f"Empty target_column at Criteria row {idx + 2}")

            source_raw = "" if pd.isna(row["source_column"]) else str(row["source_column"])
            source_columns = [col.strip() for col in source_raw.split(",") if col.strip()]
            operation = str(row["operation"]).strip().upper() if not pd.isna(row["operation"]) else "COPY"

            rules.append(
                Rule(
                    rule_order=int(row["rule_order"]),
                    target_column=target_column,
                    source_columns=source_columns,
                    operation=operation,
                    operation_arg=None if pd.isna(row["operation_arg"]) else str(row["operation_arg"]),
                    default_value=None if pd.isna(row["default_value"]) else row["default_value"],
                    required=str(row["required"]).strip().lower() in {"true", "1", "yes", "y"},
                    condition=None if pd.isna(row["condition"]) else str(row["condition"]),
                )
            )

        return sorted(rules, key=lambda r: r.rule_order)

    def _apply_rules(
        self,
        data_df: pd.DataFrame,
        rules: List[Rule],
        result_template_df: Optional[pd.DataFrame],
    ) -> pd.DataFrame:
        records: List[Dict[str, Any]] = []

        for row_index, row in data_df.iterrows():
            row_context = row.to_dict()
            transformed: Dict[str, Any] = {}

            for rule in rules:
                if rule.condition and not self._condition_evaluator.evaluate(rule.condition, row_context):
                    continue

                value = self._execute_rule(rule, row_context)
                if self._is_missing(value):
                    value = rule.default_value

                if rule.required and self._is_missing(value):
                    raise TransformationError(
                        f"Required field '{rule.target_column}' is empty for Data row {row_index + 2}"
                    )

                transformed[rule.target_column] = value

            records.append(transformed)

        result_df = pd.DataFrame(records)
        if result_template_df is not None and not result_template_df.empty:
            target_order = [c for c in result_template_df.columns if c in result_df.columns]
            remainder = [c for c in result_df.columns if c not in target_order]
            result_df = result_df[target_order + remainder]

        return result_df

    def _execute_rule(self, rule: Rule, row_context: Dict[str, Any]) -> Any:
        op = rule.operation

        if op == "STATIC":
            return rule.operation_arg

        if op == "COPY":
            return self._get_single_value(rule, row_context)

        if op == "CONCAT":
            separator = rule.operation_arg if rule.operation_arg is not None else " "
            values = [self._safe_to_string(row_context.get(column)) for column in rule.source_columns]
            return separator.join(v for v in values if v)

        source_value = self._get_single_value(rule, row_context)
        if self._is_missing(source_value):
            return source_value

        if op == "UPPER":
            return self._safe_to_string(source_value).upper()
        if op == "LOWER":
            return self._safe_to_string(source_value).lower()
        if op == "TITLE":
            return self._safe_to_string(source_value).title()
        if op == "DATE_FORMAT":
            fmt = rule.operation_arg or "%Y-%m-%d"
            dt = pd.to_datetime(source_value, errors="coerce")
            return None if pd.isna(dt) else dt.strftime(fmt)

        raise TransformationError(f"Unsupported operation '{op}' for target column '{rule.target_column}'")

    @staticmethod
    def _get_single_value(rule: Rule, row_context: Dict[str, Any]) -> Any:
        if not rule.source_columns:
            return None
        return row_context.get(rule.source_columns[0])

    @staticmethod
    def _is_missing(value: Any) -> bool:
        return value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == ""

    @staticmethod
    def _safe_to_string(value: Any) -> str:
        return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)

    def _write_result_sheet(self, input_path: Path, output_path: Path, result_df: pd.DataFrame) -> None:
        if output_path.resolve() != input_path.resolve():
            # Keep all original sheets by copying file content first.
            with open(input_path, "rb") as src, open(output_path, "wb") as dst:
                dst.write(src.read())

        with pd.ExcelWriter(output_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
            result_df.to_excel(writer, sheet_name="Result", index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Transform Excel workbook using Data + Criteria sheets")
    parser.add_argument("input", type=Path, help="Path to source Excel workbook")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Path to output workbook (default: overwrite input)",
    )
    parser.add_argument("--log-level", default="INFO", help="Logging level (e.g., INFO, DEBUG)")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(
        level=getattr(logging, str(args.log_level).upper(), logging.INFO),
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    )

    output_path = args.output or args.input
    service = ExcelTransformationService()
    service.transform_workbook(args.input, output_path)


if __name__ == "__main__":
    main()
