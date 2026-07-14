"""Local Streamlit interface for the SheetProof review service."""

from __future__ import annotations

import json
import logging
import shutil
import tempfile
from pathlib import Path
from typing import Protocol

import streamlit as st

from sheetproof.exceptions import SheetProofError
from sheetproof.models import ReviewResult, RiskLevel, RuleStatus, Severity
from sheetproof.reports.html_report import render_html
from sheetproof.reports.json_report import render_json
from sheetproof.services.review_service import ReviewService

MAX_UPLOAD_BYTES = 25 * 1024 * 1024
COPY_CHUNK_BYTES = 1024 * 1024
WORKBOOK_SUFFIXES = {".xlsx", ".xlsm"}
CONFIG_SUFFIXES = {".yml", ".yaml"}
LOGGER = logging.getLogger(__name__)


class UploadedFileLike(Protocol):
    """Small upload contract used to keep file handling testable."""

    name: str
    size: int

    def read(self, size: int = -1) -> bytes: ...

    def seek(self, offset: int, whence: int = 0) -> int: ...


class UploadValidationError(ValueError):
    """Raised when an uploaded file violates local safety limits."""


def _validated_suffix(upload: UploadedFileLike, allowed: set[str], label: str) -> str:
    suffix = Path(upload.name).suffix.lower()
    if suffix not in allowed:
        allowed_text = ", ".join(sorted(allowed))
        raise UploadValidationError(f"{label} must use one of: {allowed_text}")
    if upload.size > MAX_UPLOAD_BYTES:
        raise UploadValidationError(f"{label} exceeds the 25 MB upload limit")
    return suffix


def _save_upload(upload: UploadedFileLike, target: Path) -> None:
    """Copy an upload in bounded chunks, enforcing the limit while streaming."""

    upload.seek(0)
    written = 0
    with target.open("xb") as destination:
        while True:
            chunk = upload.read(COPY_CHUNK_BYTES)
            if not chunk:
                break
            written += len(chunk)
            if written > MAX_UPLOAD_BYTES:
                raise UploadValidationError(f"{upload.name} exceeds the 25 MB upload limit")
            destination.write(chunk)
    upload.seek(0)


def _json_text(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)


def _high_risk_rows(result: ReviewResult) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    high_levels = {RiskLevel.HIGH, RiskLevel.CRITICAL}
    for category, changes in (
        ("structure", result.structure_changes),
        ("cell", result.cell_changes),
        ("formula", result.formula_changes),
    ):
        for change in changes:
            if change.risk_level in high_levels:
                rows.append(
                    {
                        "level": change.risk_level.value,
                        "category": category,
                        "location": change.location,
                        "finding": change.description,
                        "evidence": _json_text(change.evidence),
                    }
                )
    for rule in result.rule_results:
        if rule.status in {
            RuleStatus.FAILED,
            RuleStatus.WARNING,
            RuleStatus.ERROR,
        } and rule.severity in {
            Severity.HIGH,
            Severity.CRITICAL,
        }:
            rows.append(
                {
                    "level": rule.severity.value.upper(),
                    "category": "rule",
                    "location": rule.location or "—",
                    "finding": f"{rule.name}: {rule.reason}",
                    "evidence": _json_text(rule.evidence),
                }
            )
    return rows


def _cell_change_rows(result: ReviewResult) -> list[dict[str, object]]:
    return [
        {
            "risk": change.risk_level.value,
            "type": change.change_type,
            "location": change.location,
            "before": change.before.value if change.before else None,
            "after": change.after.value if change.after else None,
            "description": change.description,
        }
        for change in result.cell_changes
    ]


def _formula_overwrite_rows(result: ReviewResult) -> list[dict[str, object]]:
    return [
        {
            "risk": change.risk_level.value,
            "location": change.location,
            "original_formula": change.before_formula,
            "replacement_value": change.replacement_value,
            "replacement_kind": (
                change.replacement_kind.value if change.replacement_kind is not None else None
            ),
            "manual_review_recommendation": change.manual_review_recommendation,
            "description": change.description,
            "neighbor_pattern": _json_text(change.neighboring_formula_pattern),
        }
        for change in result.formula_changes
        if change.before_formula is not None and change.after_formula is None
    ]


def _formula_change_rows(result: ReviewResult) -> list[dict[str, object]]:
    return [
        {
            "risk": change.risk_level.value,
            "type": change.change_type,
            "location": change.location,
            "before": change.before_formula,
            "after": change.after_formula,
            "high_impact": change.high_impact,
            "description": change.description,
        }
        for change in result.formula_changes
    ]


def _structure_change_rows(result: ReviewResult) -> list[dict[str, object]]:
    return [
        {
            "risk": change.risk_level.value,
            "type": change.change_type,
            "location": change.location,
            "before": _json_text(change.before),
            "after": _json_text(change.after),
            "description": change.description,
        }
        for change in result.structure_changes
    ]


def _rule_rows(result: ReviewResult) -> list[dict[str, object]]:
    return [
        {
            "status": rule.status.value,
            "name": rule.name,
            "type": rule.rule_type,
            "severity": rule.severity.value,
            "location": rule.location,
            "reason": rule.reason,
            "evidence": _json_text(rule.evidence),
        }
        for rule in result.rule_results
    ]


def _show_table(rows: list[dict[str, object]], empty_message: str) -> None:
    if rows:
        st.dataframe(rows, hide_index=True, use_container_width=True)
    else:
        st.info(empty_message)


def _run_review(
    before_upload: UploadedFileLike,
    after_upload: UploadedFileLike,
    config_upload: UploadedFileLike | None,
) -> None:
    before_suffix = _validated_suffix(before_upload, WORKBOOK_SUFFIXES, "Before workbook")
    after_suffix = _validated_suffix(after_upload, WORKBOOK_SUFFIXES, "After workbook")
    config_suffix = (
        _validated_suffix(config_upload, CONFIG_SUFFIXES, "Rule configuration")
        if config_upload is not None
        else None
    )

    temp_root = Path(tempfile.mkdtemp(prefix="sheetproof-web-"))
    try:
        before_path = temp_root / f"before{before_suffix}"
        after_path = temp_root / f"after{after_suffix}"
        config_path = temp_root / f"sheetproof{config_suffix}" if config_suffix else None
        _save_upload(before_upload, before_path)
        _save_upload(after_upload, after_path)
        if config_upload is not None and config_path is not None:
            _save_upload(config_upload, config_path)

        result = ReviewService().review(
            before_path,
            after_path,
            config_path=config_path,
        )
        json_report = render_json(result)
        html_report = render_html(result)
        st.session_state["sheetproof_result"] = result
        st.session_state["sheetproof_json"] = json_report
        st.session_state["sheetproof_html"] = html_report
    finally:
        try:
            shutil.rmtree(temp_root)
        except OSError as exc:
            LOGGER.exception("Unable to remove SheetProof upload directory: %s", temp_root)
            raise UploadValidationError(
                f"审查临时文件未能删除, 请手动移除: {temp_root}"
            ) from exc


def _render_result(result: ReviewResult, json_report: str, html_report: str) -> None:
    summary = result.summary
    st.subheader("审查结果")
    metric_columns = st.columns(5)
    metric_columns[0].metric("风险评分", f"{summary.risk_score}/100")
    metric_columns[1].metric("风险等级", summary.risk_level.value)
    metric_columns[2].metric("变更单元格", summary.changed_cells)
    metric_columns[3].metric("公式变更", summary.changed_formulas)
    metric_columns[4].metric("公式覆盖", summary.formula_overwrites)

    with st.expander("完整修改摘要", expanded=True):
        st.json(summary.model_dump(mode="json"))

    high_risk_tab, cells_tab, overwrites_tab, rules_tab, more_tab = st.tabs(
        ["高风险发现", "单元格差异", "公式覆盖", "规则结果", "更多差异"]
    )
    with high_risk_tab:
        _show_table(_high_risk_rows(result), "未发现 HIGH 或 CRITICAL 级别项目。")
    with cells_tab:
        _show_table(_cell_change_rows(result), "没有单元格语义差异。")
    with overwrites_tab:
        _show_table(_formula_overwrite_rows(result), "没有公式被固定值、文本或空白覆盖。")
    with rules_tab:
        _show_table(_rule_rows(result), "未配置业务规则。")
    with more_tab:
        st.caption("工作簿结构变化")
        _show_table(_structure_change_rows(result), "没有工作簿结构变化。")
        st.caption("全部公式变化")
        _show_table(_formula_change_rows(result), "没有公式变化。")

    if result.errors:
        st.subheader("错误")
        for error in result.errors:
            st.text(error)
    if result.limitations:
        with st.expander("工具限制"):
            for limitation in result.limitations:
                st.text(f"• {limitation}")

    download_json, download_html = st.columns(2)
    download_json.download_button(
        "下载 JSON 报告",
        data=json_report.encode("utf-8"),
        file_name="sheetproof-report.json",
        mime="application/json",
        use_container_width=True,
    )
    download_html.download_button(
        "下载 HTML 报告",
        data=html_report.encode("utf-8"),
        file_name="sheetproof-report.html",
        mime="text/html",
        use_container_width=True,
    )


def main() -> None:
    st.set_page_config(page_title="SheetProof", page_icon="🔎", layout="wide")
    st.title("SheetProof")
    st.caption("本地、确定性的 Excel 变更审查。文件不会上传到第三方, 也不会执行宏或外部链接。")

    before_column, after_column = st.columns(2)
    with before_column:
        before_upload = st.file_uploader(
            "修改前工作簿",
            type=["xlsx", "xlsm"],
            key="before_workbook",
        )
    with after_column:
        after_upload = st.file_uploader(
            "修改后工作簿",
            type=["xlsx", "xlsm"],
            key="after_workbook",
        )
    config_upload = st.file_uploader(
        "可选规则配置 (sheetproof.yml)",
        type=["yml", "yaml"],
        key="rule_config",
    )
    st.caption("每个上传文件最大 25 MB; 仅接受 .xlsx、.xlsm、.yml 和 .yaml。")

    start = st.button(
        "开始审查",
        type="primary",
        disabled=before_upload is None or after_upload is None,
    )
    if start and before_upload is not None and after_upload is not None:
        for key in ("sheetproof_result", "sheetproof_json", "sheetproof_html"):
            st.session_state.pop(key, None)
        try:
            with st.spinner("正在进行确定性审查……"):
                _run_review(before_upload, after_upload, config_upload)
            st.success("审查完成, 临时文件已清理。")
        except (UploadValidationError, SheetProofError) as exc:
            st.error(str(exc))
        except Exception:
            LOGGER.exception("Unexpected SheetProof web review failure")
            st.error("审查遇到内部错误。请查看运行 Streamlit 的终端日志。")

    result = st.session_state.get("sheetproof_result")
    json_report = st.session_state.get("sheetproof_json")
    html_report = st.session_state.get("sheetproof_html")
    if isinstance(result, ReviewResult) and isinstance(json_report, str) and isinstance(
        html_report, str
    ):
        _render_result(result, json_report, html_report)


if __name__ == "__main__":
    main()
