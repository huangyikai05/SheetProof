"""Load and validate SheetProof's YAML configuration."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import yaml
from pydantic import ValidationError
from yaml.constructor import ConstructorError
from yaml.nodes import MappingNode
from yaml.resolver import BaseResolver

from sheetproof.exceptions import ConfigurationError
from sheetproof.models import RuleType, SheetProofConfig
from sheetproof.rules.builtin_rules import RuleEvaluationError, parse_range


class _UniqueKeySafeLoader(yaml.SafeLoader):
    """Safe YAML loader that rejects duplicate keys at every mapping depth."""


def _construct_unique_mapping(
    loader: _UniqueKeySafeLoader,
    node: MappingNode,
    deep: bool = False,
) -> dict[Any, Any]:
    loader.flatten_mapping(node)
    mapping: dict[Any, Any] = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        try:
            duplicate = key in mapping
        except TypeError as exc:
            raise ConstructorError(
                "while constructing a mapping",
                node.start_mark,
                "found an unhashable mapping key",
                key_node.start_mark,
            ) from exc
        if duplicate:
            raise ConstructorError(
                "while constructing a mapping",
                node.start_mark,
                f"found duplicate key {key!r}",
                key_node.start_mark,
            )
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


_UniqueKeySafeLoader.add_constructor(
    BaseResolver.DEFAULT_MAPPING_TAG,
    _construct_unique_mapping,
)


def validate_config_text(text: str) -> SheetProofConfig:
    """Validate YAML text and return the canonical typed configuration.

    YAML is parsed with :func:`yaml.safe_load`; constructors capable of creating
    arbitrary Python objects are deliberately unavailable.
    """

    if not isinstance(text, str):
        raise ConfigurationError("Configuration content must be text")

    try:
        document: Any = yaml.load(text, Loader=_UniqueKeySafeLoader)
    except yaml.YAMLError as exc:
        detail = _yaml_error_detail(exc)
        raise ConfigurationError(f"Invalid YAML configuration: {detail}") from exc

    if document is None:
        document = {}
    if not isinstance(document, dict):
        raise ConfigurationError("Configuration root must be a YAML mapping")

    try:
        config = SheetProofConfig.model_validate(document)
    except ValidationError as exc:
        detail = _validation_detail(exc)
        raise ConfigurationError(f"Invalid SheetProof configuration: {detail}") from exc
    _validate_semantics(config)
    return config


def load_config(path: str | Path | None = None) -> SheetProofConfig:
    """Load a configuration file, or return safe defaults when ``path`` is ``None``."""

    if path is None:
        return SheetProofConfig()

    config_path = Path(path).expanduser()
    if not config_path.exists():
        raise ConfigurationError(f"Configuration file does not exist: {config_path}")
    if not config_path.is_file():
        raise ConfigurationError(f"Configuration path is not a file: {config_path}")

    try:
        text = config_path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeError) as exc:
        raise ConfigurationError(f"Unable to read configuration '{config_path}': {exc}") from exc
    return validate_config_text(text)


def _yaml_error_detail(exc: yaml.YAMLError) -> str:
    problem = getattr(exc, "problem", None)
    mark = getattr(exc, "problem_mark", None)
    if problem and mark is not None:
        return f"{problem} at line {mark.line + 1}, column {mark.column + 1}"
    return str(exc)


def _validation_detail(exc: ValidationError) -> str:
    messages: list[str] = []
    for error in exc.errors(include_url=False):
        location = ".".join(str(part) for part in error["loc"]) or "configuration"
        messages.append(f"{location}: {error['msg']}")
    return "; ".join(messages)


def _validate_semantics(config: SheetProofConfig) -> None:
    for index, rule in enumerate(config.rules, start=1):
        try:
            if rule.type is RuleType.FORMULA_REQUIRED and rule.range is not None:
                parse_range(rule.range)
            elif rule.type is RuleType.ALLOWED_CHANGE_RANGE:
                for value in rule.ranges:
                    parse_range(value)
            elif rule.type is RuleType.NUMERIC_RANGE and rule.target is not None:
                parse_range(rule.target, require_single_cell=True)
                limits = [value for value in (rule.min, rule.max) if value is not None]
                if not all(math.isfinite(value) for value in limits):
                    raise RuleEvaluationError("numeric_range limits must be finite")
                if rule.min is not None and rule.max is not None and rule.min > rule.max:
                    raise RuleEvaluationError("numeric_range min cannot be greater than max")
        except RuleEvaluationError as exc:
            raise ConfigurationError(
                f"Invalid rule {index} ({rule.name!r}): {exc}"
            ) from exc

    for index, value in enumerate(config.critical_cells, start=1):
        try:
            parse_range(value, require_single_cell=True)
        except RuleEvaluationError as exc:
            raise ConfigurationError(f"Invalid critical_cells item {index}: {exc}") from exc


__all__ = ["load_config", "validate_config_text"]
