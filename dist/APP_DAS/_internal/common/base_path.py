import sys
from pathlib import Path

def get_base_dir(max_levels_up=5):
    if getattr(sys, 'frozen', False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).resolve().parent

    for _ in range(max_levels_up + 1):
        if (base / 'main.py').exists():
            # Modo script: retorna a pasta onde está main.py
            return base

        if (base / 'APP_DAS.exe').exists():
            # Modo exe: retorna o base com '_internal' anexado
            return base / '_internal'

        base = base.parent

    raise FileNotFoundError(
        "Não foi possível encontrar a raiz do projeto com 'main.py' "
        f"ou 'APP_DAS.exe' subindo até {max_levels_up} níveis a partir de {Path(__file__).parent}"
    )