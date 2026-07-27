"""Formules compatibles Office 16 pour l'onglet Proposition."""

from __future__ import annotations


def _column_letter(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _data_lookup_formula(
    row: int,
    header_row: int,
    column: int,
    data_last_row: int,
) -> str:
    header_cell = f"{_column_letter(column)}${header_row}"
    return (
        f'=IF($A{row}="","",IFERROR(INDEX(\'DATA\'!$A$2:$EC${data_last_row},'
        f'MATCH($A{row},\'DATA\'!$A$2:$A${data_last_row},0),'
        f'MATCH({header_cell},\'DATA\'!$A$1:$EC$1,0)),""))'
    )


def _fixed_data_lookup_formula(
    row: int,
    source_column: str,
    data_last_row: int,
) -> str:
    return (
        f'=IF($A{row}="","",IFERROR(INDEX(\'DATA\'!${source_column}$2:'
        f'${source_column}${data_last_row},MATCH($A{row},'
        f'\'DATA\'!$A$2:$A${data_last_row},0)),""))'
    )


def _result_matrix(
    start_row: int,
    settings_row: int,
    header_row: int,
    helper_metric_column: str,
    helper_rank_column: str,
    helper_isin_column: str,
    helper_last_row: int,
    data_last_row: int,
) -> list[list[str]]:
    matrix = []
    top_n_cell = f"$B${settings_row}"
    first_result = f"$A${start_row}"

    for row in range(start_row, start_row + 30):
        position = f"ROWS({first_result}:A{row})"
        formulas = [
            (
                f'=IF({position}>MIN({top_n_cell},COUNT(${helper_metric_column}$6:'
                f'${helper_metric_column}${helper_last_row})),"",IFERROR('
                f'INDEX(${helper_isin_column}$6:${helper_isin_column}${helper_last_row},'
                f'MATCH({position},${helper_rank_column}$6:'
                f'${helper_rank_column}${helper_last_row},0)),""))'
            ),
            _data_lookup_formula(row, header_row, 2, data_last_row),
            _fixed_data_lookup_formula(row, "B", data_last_row),
            _fixed_data_lookup_formula(row, "C", data_last_row),
            _fixed_data_lookup_formula(row, "D", data_last_row),
            (
                f'=IF($A{row}="","",IFERROR(INDEX(\'DATA\'!$A$2:$EC${data_last_row},'
                f'MATCH($A{row},\'DATA\'!$A$2:$A${data_last_row},0),'
                f'MATCH("Weight PTF",\'DATA\'!$A$1:$EC$1,0)),""))'
            ),
            (
                f'=IF($A{row}="","",IFERROR(INDEX(\'DATA\'!$A$2:$EC${data_last_row},'
                f'MATCH($A{row},\'DATA\'!$A$2:$A${data_last_row},0),'
                f'MATCH("Weight Bench",\'DATA\'!$A$1:$EC$1,0)),""))'
            ),
        ]
        formulas.extend(
            _data_lookup_formula(row, header_row, column, data_last_row)
            for column in range(8, 20)
        )
        matrix.append(formulas)
    return matrix


def write_proposition_formulas(
    worksheet,
    data_rows: int,
) -> None:
    """Réécrit les résultats avec des formules recalculables liées à DATA."""
    data_rows = max(1, data_rows)
    data_last_row = data_rows + 1
    helper_last_row = data_rows + 5

    helper_matrix = []
    for row in range(6, helper_last_row + 1):
        source_position = f"ROWS($U$6:U{row})"
        top_metric = (
            f'=IFERROR(IF(INDEX(\'DATA\'!$A$2:$EC${data_last_row},{source_position},'
            f'MATCH("Weight PTF",\'DATA\'!$A$1:$EC$1,0))<>0,'
            f'INDEX(\'DATA\'!$A$2:$EC${data_last_row},{source_position},'
            f'MATCH($B$2,\'DATA\'!$A$1:$EC$1,0)),""),"")'
        )
        top_rank = (
            f'=IF($U{row}="","",RANK($U{row},$U$6:$U${helper_last_row},0)+'
            f'COUNTIF($U$6:$U{row},$U{row})-1)'
        )
        top_isin = (
            f'=IFERROR(INDEX(\'DATA\'!$A$2:$A${data_last_row},'
            f'ROWS($W$6:W{row})),"")'
        )
        benchmark_metric = (
            f'=IFERROR(IF(INDEX(\'DATA\'!$A$2:$EC${data_last_row},{source_position},'
            f'MATCH("Weight Bench",\'DATA\'!$A$1:$EC$1,0))<>0,'
            f'INDEX(\'DATA\'!$A$2:$EC${data_last_row},{source_position},'
            f'MATCH($B$39,\'DATA\'!$A$1:$EC$1,0)),""),"")'
        )
        benchmark_rank = (
            f'=IF($X{row}="","",RANK($X{row},$X$6:$X${helper_last_row},1)+'
            f'COUNTIF($X$6:$X{row},$X{row})-1)'
        )
        benchmark_isin = (
            f'=IFERROR(INDEX(\'DATA\'!$A$2:$A${data_last_row},'
            f'ROWS($Z$6:Z{row})),"")'
        )
        helper_matrix.append(
            [
                top_metric,
                top_rank,
                top_isin,
                benchmark_metric,
                benchmark_rank,
                benchmark_isin,
            ]
        )

    top_matrix = _result_matrix(
        start_row=6,
        settings_row=3,
        header_row=5,
        helper_metric_column="U",
        helper_rank_column="V",
        helper_isin_column="W",
        helper_last_row=helper_last_row,
        data_last_row=data_last_row,
    )
    benchmark_matrix = _result_matrix(
        start_row=43,
        settings_row=40,
        header_row=42,
        helper_metric_column="X",
        helper_rank_column="Y",
        helper_isin_column="Z",
        helper_last_row=helper_last_row,
        data_last_row=data_last_row,
    )

    worksheet.range("A6").resize(30, 19).formula = top_matrix
    worksheet.range("A43").resize(30, 19).formula = benchmark_matrix
    worksheet.range("U6").resize(data_rows, 6).formula = helper_matrix

    if hasattr(worksheet, "hide_columns"):
        worksheet.hide_columns(21, 26)
    else:
        worksheet.range("U:Z").api.EntireColumn.Hidden = True
