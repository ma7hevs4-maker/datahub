import os
import shutil
from pathlib import Path


def sincronizar(arquivos: list[Path], pasta_sharepoint: str, log_fn=print) -> None:
    """Envia apenas os arquivos listados para o SharePoint."""
    destino = Path(pasta_sharepoint)

    if not destino.exists():
        log_fn(f"  ❌ Pasta SharePoint não encontrada: {destino}")
        return

    if not arquivos:
        log_fn(f"  ℹ️  Nenhum arquivo para sincronizar")
        return

    copiados = 0

    for src in arquivos:
        src = Path(src)
        if not src.exists():
            continue

        # Calcula destino mantendo estrutura relativa
        if "bases_mensais" in str(src):
            parts = src.parts
            idx = parts.index("bases_mensais")
            rel = Path(*parts[idx:])
            dst = destino / rel
        else:
            dst = destino / src.name

        dst.parent.mkdir(parents=True, exist_ok=True)

        if not dst.exists() or os.path.getmtime(str(src)) >= os.path.getmtime(str(dst)) - 1:
            shutil.copy2(str(src), str(dst))
            copiados += 1

            # Mostra caminho relativo para log
            if "bases_mensais" in str(src):
                parts = src.parts
                idx = parts.index("bases_mensais")
                log_fn(f"  📤 {'/'.join(parts[idx:])}")
            else:
                log_fn(f"  📤 {src.name}")

    if copiados > 0:
        log_fn(f"  ✅ {copiados} arquivo(s) novo(s) ou atualizado(s) enviado(s)")
    else:
        log_fn(f"  ℹ️  Nenhum arquivo alterado — SharePoint já está atualizado")
