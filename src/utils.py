import re
from docx.table import Table


def delete_empty_rows(table: Table) -> None:
    """
    Remove linhas vazias de uma tabela.

    Args:
        table (Table): A tabela do qual as linhas vazias serão removidas.
    """
    rows_to_delete = []

    for row in table.rows:
        if all(cell.text.strip() == "" for cell in row.cells):
            rows_to_delete.append(row)

    for row in reversed(rows_to_delete):
        tbl = row._element.getparent()
        tbl.remove(row._element)


def camel_case_to_capitalized(camel_case_str: str) -> str:
    """
    Converte uma string em formato camel case para uma string com palavras capitalizadas.

    Args:
        camel_case_str (str): A string em formato camel case.

    Returns:
        str: A string resultante com cada palavra capitalizada.

    Example:
        >>> camel_case_to_capitalized("camelCaseExample")
        'Camel Case Example'
    """
    words = re.findall(r"[A-Z][a-z]*|[a-z]+", camel_case_str)
    return " ".join(word.capitalize() for word in words)
