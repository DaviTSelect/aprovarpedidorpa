"""Verificação e instalação automática de atualizações via GitHub Releases.

Configuração:
  Defina REPO_GITHUB com "usuario/repositorio" após criar o repositório no GitHub.
  O arquivo .exe publicado em cada release deve se chamar NOME_EXE.

Fluxo:
  1. Consulta a API do GitHub para a release mais recente.
  2. Compara o tag da release com __version__ do executável.
  3. Se há versão nova e o programa está rodando como .exe (frozen),
     baixa o novo .exe e agenda a substituição via batch após fechar.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

# ── Configure aqui após criar o repositório no GitHub ──────────────────────
# Exemplo: REPO_GITHUB = "minhaempresa/aprovar-pedidos-rpa"
REPO_GITHUB: str = "DaviTSelect/aprovarpedidorpa"

NOME_EXE: str = "AprovarPedidosPF.exe"

# Cores de status (usadas pela UI tkinter)
COR_OK      = "#53b88c"   # verde — atualizado
COR_BUSY    = "#e9b44c"   # amarelo — verificando / baixando
COR_UNKNOWN = "#85929c"   # cinza — sem configuração ou erro de rede


def buscar_release(repo: str) -> dict:
    """Consulta a API do GitHub e retorna os dados da release mais recente."""
    import requests
    r = requests.get(
        f"https://api.github.com/repos/{repo}/releases/latest",
        timeout=15,
        headers={"Accept": "application/vnd.github+json"},
    )
    r.raise_for_status()
    return r.json()


def tem_versao_nova(release: dict, versao_atual: str) -> bool:
    """Retorna True se a release do GitHub é mais recente que versao_atual."""
    tag = release.get("tag_name", "").lstrip("v")
    atual = versao_atual.lstrip("v")

    def _parse(v: str) -> tuple[int, ...]:
        try:
            return tuple(int(x) for x in v.split("."))
        except Exception:
            return (0,)

    return _parse(tag) > _parse(atual)


def pode_instalar() -> bool:
    """Atualização automática só faz sentido no executável compilado."""
    return getattr(sys, "frozen", False)


def instalar(release: dict, nome_exe: str) -> None:
    """Baixa o novo .exe e cria um batch que o substitui após este processo fechar."""
    import requests

    assets = release.get("assets", [])
    asset = next((a for a in assets if a["name"] == nome_exe), None)
    if not asset:
        raise RuntimeError(
            f"Asset '{nome_exe}' não encontrado na release {release.get('tag_name', '?')}.\n"
            f"Assets disponíveis: {[a['name'] for a in assets]}"
        )

    url = asset["browser_download_url"]
    exe_atual = Path(sys.executable)
    exe_novo = exe_atual.with_name(exe_atual.stem + ".novo.exe")

    r = requests.get(url, stream=True, timeout=180)
    r.raise_for_status()
    with open(exe_novo, "wb") as f:
        for chunk in r.iter_content(chunk_size=65536):
            if chunk:
                f.write(chunk)

    # Batch: espera o processo sair, substitui o exe e relança
    bat = exe_atual.parent / "_atualizar_aprovar.bat"
    bat.write_text(
        "@echo off\n"
        "timeout /t 3 /nobreak >nul\n"
        f'move /y "{exe_novo}" "{exe_atual}"\n'
        f'start "" "{exe_atual}"\n'
        'del "%~f0"\n',
        encoding="utf-8",
    )
    subprocess.Popen(
        ["cmd", "/c", str(bat)],
        creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS,
    )
