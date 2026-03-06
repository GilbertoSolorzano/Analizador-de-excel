from pathlib import Path
def build_save_path(original_path: str, suffix: str = '_filtrado', out_ext: str = '.xlsx') -> str:
    p = Path(original_path)
    stem = p.stem  # nombre sin extensión
    parent = p.parent
    new_name = f"{stem}{suffix}{out_ext}"
    full = parent / new_name
    i = 1
    # Si ya existe, agrega un contador: nombre_filtrado(1).xlsx, etc.
    while full.exists():
        full = parent / f"{stem}{suffix}({i}){out_ext}"
        i += 1
    return str(full)