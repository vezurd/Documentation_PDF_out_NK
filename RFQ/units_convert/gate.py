"""Build and apply validated conversion plans."""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from base.base_classes import CheckElement
from RFQ.ds_compare.ds_units_normalize import join_unique_units_text, normalize_units_text
from RFQ.units_convert.matrix import (
    MatrixPair,
    MatrixRow,
    _document_semantic_key,
    _file_signature,
    format_source_name_comment,
    load_matrix,
    reconcile_matrix,
    resolved_conversion_target,
    save_matrix_atomic,
)
from RFQ.units_convert.models import (
    RESIDUAL_TOLERANCE,
    STATUS_CONVERTED,
    STATUS_IDENTITY,
    STATUS_NO_GOOGLE,
    ConversionAction,
    ConversionDependency,
    ConversionInvariant,
    ConversionPlan,
    ConversionRequest,
    FractionalConversionIssue,
    GoogleUnitsIndex,
    RowStdBinding,
    UnitsConversionError,
    normalize_code,
    parse_decimal_quantity,
)


def _active_units_by_code(
    requests: Sequence[ConversionRequest],
) -> tuple[dict[str, dict[str, str]], dict[str, str], dict[str, dict[str, str]]]:
    active: dict[str, dict[str, str]] = {}
    code_display_by_code: dict[str, str] = {}
    names_by_code: dict[str, dict[str, list[str]]] = {}
    for request in requests:
        code_norm = normalize_code(request.code)
        source_norm = normalize_units_text(request.source_unit)
        if not code_norm or not source_norm:
            continue
        code_display_by_code.setdefault(
            code_norm,
            str(request.code or "").strip() or code_norm,
        )
        bucket = active.setdefault(code_norm, {})
        bucket.setdefault(source_norm, str(request.source_unit or "").strip() or source_norm)
        item_name = " ".join(str(request.item_name or "").split())
        if item_name:
            name_bucket = names_by_code.setdefault(code_norm, {}).setdefault(source_norm, [])
            if item_name not in name_bucket:
                name_bucket.append(item_name)
    names_flat = {
        code: {
            unit: format_source_name_comment(names)
            for unit, names in units.items()
        }
        for code, units in names_by_code.items()
    }
    return active, code_display_by_code, names_flat


def _build_trace(
    *,
    code_raw: str,
    source_unit: str,
    target_unit: str,
    original_quantity: Decimal,
    result_quantity: Decimal,
    coefficient: Decimal,
    status: str,
) -> str:
    return (
        f"code={code_raw}; src={source_unit}; tgt={target_unit}; "
        f"qty={original_quantity}; coef={coefficient}; result={result_quantity}; status={status}"
    )


def _quantity_cell_value(quantity: Decimal) -> int | str:
    """Write exact ints as ``int``; keep genuine fractions as decimal text."""
    integral = quantity.to_integral_value()
    if quantity == integral:
        return int(integral)
    return format(quantity.normalize(), "f")


def _apply_residual_policy(
    *,
    result: Decimal,
    request: ConversionRequest,
    code_raw: str,
    source_unit: str,
    target_unit: str,
    coefficient: Decimal,
    original_quantity: Decimal,
) -> tuple[Decimal, bool, FractionalConversionIssue | None]:
    nearest = result.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    residual = abs(result - nearest)
    if residual <= RESIDUAL_TOLERANCE:
        adjusted = residual != 0
        return Decimal(int(nearest)), adjusted, None
    issue = FractionalConversionIssue(
        contour=request.contour,
        location=request.location,
        request_id=request.request_id,
        code=code_raw,
        original_quantity=original_quantity,
        source_unit=source_unit,
        target_unit=target_unit,
        coefficient=coefficient,
        result_quantity=result,
        nearest_int=nearest,
        residual=residual,
        tags_count=request.tags_count,
        invariant_skipped=request.invariant != ConversionInvariant.NONE,
    )
    return result, False, issue


def _enforce_invariant(
    *,
    request: ConversionRequest,
    result_quantity: Decimal,
    changed: bool,
) -> None:
    if not changed or request.invariant == ConversionInvariant.NONE:
        return
    if request.invariant == ConversionInvariant.TAGS_EQUAL:
        if result_quantity != Decimal(request.tags_count):
            raise UnitsConversionError(
                "Нарушен инвариант TAGS_EQUAL для "
                f"request_id={request.request_id} источник={request.location}: "
                f"tags_count={request.tags_count} результат={result_quantity}"
            )
    elif request.invariant == ConversionInvariant.TAGS_NOT_EXCEED_QUANTITY:
        if Decimal(request.tags_count) > result_quantity:
            raise UnitsConversionError(
                "Нарушен инвариант TAGS_NOT_EXCEED_QUANTITY для "
                f"request_id={request.request_id} источник={request.location}: "
                f"tags_count={request.tags_count} результат={result_quantity}"
            )


def _append_dependency(
    dependencies: list[ConversionDependency],
    *,
    code: str,
    source_unit: str,
    target_unit: str,
    coefficient: Decimal,
) -> None:
    dependencies.append(
        ConversionDependency(
            code=code,
            source_unit=source_unit,
            target_unit=target_unit,
            coefficient=coefficient,
        )
    )


def _dedupe_dependencies(dependencies: Sequence[ConversionDependency]) -> tuple[ConversionDependency, ...]:
    unique = {
        (item.code, item.source_unit, item.target_unit, item.coefficient): item
        for item in dependencies
    }
    return tuple(
        sorted(
            unique.values(),
            key=lambda item: (item.code, item.source_unit, item.target_unit, item.coefficient),
        )
    )


def _resolve_binding_map(
    plan: ConversionPlan,
    bindings: Sequence[RowStdBinding],
) -> dict[str, RowStdBinding]:
    action_ids = {action.request_id for action in plan.actions}
    binding_map: dict[str, RowStdBinding] = {}
    for binding in bindings:
        if binding.request_id in binding_map:
            raise UnitsConversionError(
                f"Дублирующая привязка для request_id {binding.request_id!r}"
            )
        binding_map[binding.request_id] = binding
    if set(binding_map) != action_ids:
        missing = sorted(action_ids - set(binding_map))
        extra = sorted(set(binding_map) - action_ids)
        details: list[str] = []
        if missing:
            details.append(f"missing={missing}")
        if extra:
            details.append(f"extra={extra}")
        raise UnitsConversionError("Неполные или лишние привязки: " + ", ".join(details))
    return binding_map


def _snapshot_cell_value(row, column: str) -> object | None:
    element = row.el.get(column)
    if element is None:
        return None
    return element.value


def _matrix_action_hint(path: Path) -> str:
    """Return human-readable steps for resolving matrix coefficient errors."""
    return (
        "\n\nЧто нужно сделать:\n"
        "1. Откройте файл матрицы единиц измерения:\n"
        f"   {path}\n"
        "2. В жёлтых ячейках «Коэффициент N» замените «?» на коэффициент по правилу:\n"
        "   количество в ЕИ Google = исходное количество × коэффициент.\n"
        "   Серый столбец «Формула N» покажет результат, например "
        "«100/шт = 1/упак».\n"
        "   Точный вариант для дробного коэффициента: 1/3 (вводить без знака "
        "«=»). Краткая десятичная запись 0,33 также распознаётся как 1/3.\n"
        "3. Если статус «collision»: в столбце «Ед. изм. Google» укажите эталонную "
        "единицу (локально, без заведения кода в Google-базу). Сохраните, закройте "
        "Excel и повторите запуск — остальные найденные ЕИ станут «?».\n"
        "4. Сохраните и закройте файл Excel.\n"
        "5. Повторите запуск."
    )


def _build_matrix_lookup_indexes(
    document,
) -> tuple[dict[str, MatrixRow], dict[tuple[str, str], MatrixPair]]:
    """Build O(1) lookup maps for matrix rows and source-unit pairs."""
    rows_by_code: dict[str, MatrixRow] = {}
    pairs_by_key: dict[tuple[str, str], MatrixPair] = {}
    for row in document.rows:
        if row.code_normalized in rows_by_code:
            raise UnitsConversionError(
                f"Дублирующийся нормализованный код матрицы {row.code_normalized!r}"
            )
        rows_by_code[row.code_normalized] = row
        for pair in row.pairs:
            key = (row.code_normalized, pair.source_normalized)
            if key in pairs_by_key:
                continue
            pairs_by_key[key] = pair
    return rows_by_code, pairs_by_key


def build_conversion_plan(
    requests: Sequence[ConversionRequest],
    google_index: GoogleUnitsIndex,
    matrix_path: str | Path,
) -> ConversionPlan:
    """Validate requests, reconcile/write matrix xlsx, and return an immutable plan.

    Args:
        requests: Conversion requests keyed by unique ``request_id``.
        google_index: Google units lookup built from base rows.
        matrix_path: Path to conversion matrix xlsx.

    Returns:
        Immutable ``ConversionPlan`` when all active mappings are valid.

    Raises:
        UnitsConversionError: On duplicate request ids, fatal mapping/quantity errors,
            or when any active request cannot be converted.
    """
    path = Path(matrix_path)
    request_ids = [request.request_id for request in requests]
    if len(set(request_ids)) != len(request_ids):
        raise UnitsConversionError(
            "Дублирующиеся значения request_id в запросах конвертации не допускаются"
        )

    active_by_code, code_display_by_code, names_by_code = _active_units_by_code(requests)
    previous_signature = _file_signature(path) if path.exists() else None
    document = load_matrix(path)
    previous_semantic = _document_semantic_key(document)
    document = reconcile_matrix(
        document,
        google_index=google_index,
        active_by_code=active_by_code,
        code_display_by_code=code_display_by_code,
        names_by_code=names_by_code,
    )
    rows_by_code, pairs_by_key = _build_matrix_lookup_indexes(document)

    fatals: list[str] = []
    warnings: list[str] = []
    actions: list[ConversionAction] = []
    dependencies: list[ConversionDependency] = []
    fractional_issues: list[FractionalConversionIssue] = []

    for code_norm, units in active_by_code.items():
        if len(units) <= 1:
            continue
        matrix_row = rows_by_code.get(code_norm)
        target_norm, _target_display = resolved_conversion_target(
            code_norm,
            google_index=google_index,
            matrix_row=matrix_row,
        )
        if not target_norm:
            fatals.append(
                f"Несколько активных исходных ЕИ для кода {code_norm}: "
                f"{', '.join(sorted(units.values()))}"
            )

    for request in requests:
        code_raw = str(request.code or "").strip()
        code_norm = normalize_code(request.code)
        source_display = str(request.source_unit or "").strip()
        source_norm = normalize_units_text(request.source_unit)
        location = request.location or request.request_id

        if not code_norm:
            fatals.append(f"Пустой нормализованный код, источник={location}")
            continue
        if not source_norm:
            fatals.append(f"Пустая нормализованная исходная ЕИ, источник={location}")
            continue

        try:
            original_quantity = parse_decimal_quantity(request.quantity, location=location)
        except UnitsConversionError as exc:
            fatals.append(str(exc))
            continue

        google_norm, google_display = resolved_conversion_target(
            code_norm,
            google_index=google_index,
            matrix_row=rows_by_code.get(code_norm),
        )

        if not google_norm:
            if len(active_by_code.get(code_norm, {})) > 1:
                continue
            target_unit = source_display or source_norm
            status = STATUS_NO_GOOGLE
            coefficient = Decimal("1")
            result_quantity = original_quantity
            trace = _build_trace(
                code_raw=code_raw,
                source_unit=source_display,
                target_unit=target_unit,
                original_quantity=original_quantity,
                result_quantity=result_quantity,
                coefficient=coefficient,
                status=status,
            )
            warnings.append(
                f"Нет единицы Google для кода {code_raw}, источник={location}; "
                "количество не изменено"
            )
            actions.append(
                ConversionAction(
                    request_id=request.request_id,
                    code_raw=code_raw,
                    code_normalized=code_norm,
                    original_unit=source_display,
                    source_unit=source_display,
                    target_unit=target_unit,
                    original_quantity=original_quantity,
                    coefficient=coefficient,
                    result_quantity=result_quantity,
                    status=status,
                    trace=trace,
                    quantity_changed=False,
                    units_changed=False,
                    machine_residual_adjusted=False,
                )
            )
            _append_dependency(
                dependencies,
                code=code_norm,
                source_unit=source_norm,
                target_unit=source_norm,
                coefficient=coefficient,
            )
            continue

        target_unit = google_display or google_norm
        if source_norm == google_norm:
            coefficient = Decimal("1")
            result_quantity = original_quantity
            units_changed = target_unit != source_display
            changed = units_changed
            status = STATUS_IDENTITY if not changed else STATUS_CONVERTED
            try:
                _enforce_invariant(
                    request=request,
                    result_quantity=result_quantity,
                    changed=changed,
                )
            except UnitsConversionError as exc:
                fatals.append(str(exc))
                continue
            actions.append(
                ConversionAction(
                    request_id=request.request_id,
                    code_raw=code_raw,
                    code_normalized=code_norm,
                    original_unit=source_display,
                    source_unit=source_display,
                    target_unit=target_unit,
                    original_quantity=original_quantity,
                    coefficient=coefficient,
                    result_quantity=result_quantity,
                    status=status,
                    trace=_build_trace(
                        code_raw=code_raw,
                        source_unit=source_display,
                        target_unit=target_unit,
                        original_quantity=original_quantity,
                        result_quantity=result_quantity,
                        coefficient=coefficient,
                        status=status,
                    ),
                    quantity_changed=False,
                    units_changed=units_changed,
                    machine_residual_adjusted=False,
                )
            )
            _append_dependency(
                dependencies,
                code=code_norm,
                source_unit=source_norm,
                target_unit=google_norm,
                coefficient=coefficient,
            )
            continue

        pair = pairs_by_key.get((code_norm, source_norm))
        if pair is None:
            fatals.append(
                f"Отсутствует соответствие в матрице для кода {code_raw} "
                f"исходная ЕИ={source_display} источник={location}"
            )
            continue
        if pair.is_placeholder or pair.coefficient is None:
            fatals.append(
                f"Требуется коэффициент «?» для кода {code_raw} "
                f"исходная ЕИ={source_display} источник={location}"
            )
            continue
        coefficient = pair.coefficient
        if coefficient <= 0:
            fatals.append(
                f"Недопустимый неположительный коэффициент {coefficient} для кода {code_raw} "
                f"исходная ЕИ={source_display} источник={location}"
            )
            continue

        result_quantity = original_quantity * coefficient
        machine_adjusted = False
        units_changed = target_unit != source_display
        fractional_issue: FractionalConversionIssue | None = None
        if coefficient == Decimal("1"):
            result_quantity = original_quantity
            quantity_changed = False
            changed = units_changed
        else:
            result_quantity, machine_adjusted, fractional_issue = _apply_residual_policy(
                result=result_quantity,
                request=request,
                code_raw=code_raw,
                source_unit=source_display,
                target_unit=target_unit,
                coefficient=coefficient,
                original_quantity=original_quantity,
            )
            quantity_changed = True
            changed = True
            if fractional_issue is not None:
                fractional_issues.append(fractional_issue)

        status = STATUS_CONVERTED if changed else STATUS_IDENTITY
        if fractional_issue is None:
            try:
                _enforce_invariant(
                    request=request,
                    result_quantity=result_quantity,
                    changed=changed,
                )
            except UnitsConversionError as exc:
                fatals.append(str(exc))
                continue

        trace = _build_trace(
            code_raw=code_raw,
            source_unit=source_display,
            target_unit=target_unit,
            original_quantity=original_quantity,
            result_quantity=result_quantity,
            coefficient=coefficient,
            status=status,
        )
        if fractional_issue is not None:
            trace = f"{trace}; fractional=1"

        actions.append(
            ConversionAction(
                request_id=request.request_id,
                code_raw=code_raw,
                code_normalized=code_norm,
                original_unit=source_display,
                source_unit=source_display,
                target_unit=target_unit,
                original_quantity=original_quantity,
                coefficient=coefficient,
                result_quantity=result_quantity,
                status=status,
                trace=trace,
                quantity_changed=quantity_changed,
                units_changed=units_changed,
                machine_residual_adjusted=machine_adjusted,
                fractional_result=fractional_issue is not None,
            )
        )
        _append_dependency(
            dependencies,
            code=code_norm,
            source_unit=source_norm,
            target_unit=google_norm,
            coefficient=coefficient,
        )

    save_matrix_atomic(
        path,
        document,
        google_index=google_index,
        active_by_code=active_by_code,
        code_display_by_code=code_display_by_code,
        names_by_code=names_by_code,
        previous_signature=previous_signature,
        previous_semantic=previous_semantic,
    )

    if fatals:
        message = "\n".join(fatals)
        matrix_error_prefixes = (
            "Отсутствует соответствие в матрице",
            "Требуется коэффициент",
            "Недопустимый неположительный коэффициент",
            "Несколько активных исходных ЕИ",
        )
        if any(item.startswith(matrix_error_prefixes) for item in fatals):
            message += _matrix_action_hint(path)
        raise UnitsConversionError(message)

    if fractional_issues:
        warnings.append(
            "Дробный результат преобразования ЕИ: "
            f"{len(fractional_issues)} строк записаны в лог и пропущены дальше "
            "без округления."
        )

    return ConversionPlan(
        actions=tuple(actions),
        warnings=tuple(warnings),
        dependencies=_dedupe_dependencies(dependencies),
        fractional_issues=tuple(fractional_issues),
    )


def apply_conversion_plan(
    plan: ConversionPlan,
    bindings: Iterable[RowStdBinding],
) -> None:
    """Apply a validated plan to bound ``RowStd`` rows without partial mutation.

    Args:
        plan: Validated conversion plan from ``build_conversion_plan``.
        bindings: One binding per plan action, keyed by explicit ``request_id``.
            Multiple bindings may reference the same ``RowStd`` (for example
            ``VALUES`` and ``VALUES_2``).

    Raises:
        UnitsConversionError: When bindings are incomplete, duplicated, invalid,
            or required quantity/units columns are missing.
    """
    binding_list = list(bindings)
    binding_map = _resolve_binding_map(plan, binding_list)

    snapshots: list[tuple[RowStdBinding, dict[str, tuple[bool, object | None]]]] = []
    for binding in binding_map.values():
        row = binding.row
        for column in (binding.quantity_column, binding.units_column):
            if column not in row.el:
                raise UnitsConversionError(
                    f"Отсутствует обязательный столбец {column!r} в привязанной строке "
                    f"для request_id={binding.request_id!r}"
                )
        snapshot: dict[str, tuple[bool, object | None]] = {}
        for column in (
            binding.quantity_column,
            binding.units_column,
            binding.status_column,
            binding.trace_column,
        ):
            existed = column in row.el
            snapshot[column] = (existed, _snapshot_cell_value(row, column) if existed else None)
        snapshots.append((binding, snapshot))

    created_columns: list[tuple[RowStdBinding, str]] = []
    try:
        for action in plan.actions:
            binding = binding_map[action.request_id]
            row = binding.row
            for column in (binding.status_column, binding.trace_column):
                if column not in row.el:
                    row.el[column] = CheckElement(None)
                    created_columns.append((binding, column))

            if action.quantity_changed:
                row.el[binding.quantity_column].value = _quantity_cell_value(
                    action.result_quantity
                )
            if action.units_changed:
                row.el[binding.units_column].value = action.target_unit

            merged_status = join_unique_units_text(
                [row.el[binding.status_column].value, action.status]
            )
            merged_trace = join_unique_units_text(
                [row.el[binding.trace_column].value, action.trace]
            )
            row.el[binding.status_column].value = merged_status or None
            row.el[binding.trace_column].value = merged_trace or None
    except Exception:
        for binding, snapshot in snapshots:
            row = binding.row
            for column, (existed, previous) in snapshot.items():
                if existed:
                    row.el[column].value = previous
                elif column in row.el:
                    del row.el[column]
        raise
