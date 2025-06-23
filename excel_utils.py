# --- Funciones utilitarias ---
def excel_column_to_number(col_str):
    col_str = col_str.upper().strip()
    num = 0
    for c in col_str:
        if c < 'A' or c > 'Z':
            return 0
        num = num * 26 + (ord(c) - ord('A'))
    return num

def number_to_excel_column(n):
    n += 1  # Convertir a 1-indexed
    col_str = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return col_str

def deduplicate_columns(columns):
    seen = {}
    deduped = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            deduped.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            deduped.append(col)
    return deduped

