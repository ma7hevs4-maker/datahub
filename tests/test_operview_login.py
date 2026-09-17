"""Teste local do auto-login Microsoft/Entra do Operview (sem tocar no sistema real).

Roda Selenium (Chrome) contra uma página fake com a MESMA estrutura da
Microsoft/Entra e valida que o operview.py detecta a página, preenche
e-mail/senha das configurações e chega no app.

Executar:
    python tests/test_operview_login.py

Precisa de Google Chrome instalado (o chromedriver é baixado pelo webdriver_manager).
"""
import os
import socket
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import OperviewConfig
from core.downloader.operview import (
    _achar_por_texto,
    _aguardar_estado_login,
    _criar_driver,
    _esta_ms_login,
    _precisa_login,
    _restaurar_sessao,
    _salvar_sessao,
    tentar_login_automatico,
)
from fake_ms_server import _MsHandler, criar_servidor


def _porta_livre():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _novo_driver(base_dl, headless=True, perfil=None):
    pasta = tempfile.mkdtemp(prefix="opv_test_dl_")
    return _criar_driver(pasta, headless=headless, perfil=perfil)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    porta = _porta_livre()
    srv = criar_servidor(porta)
    srv.start()
    base = f"http://127.0.0.1:{porta}"
    drivers = []
    try:
        # ── Cenário 1: login automático completo (sem credenciais reais) ──
        d = _novo_driver(base)
        drivers.append(d)
        cfg = OperviewConfig(url=base, login="usuario.teste@enel.com", senha="SenhaFake#123")
        d.get(base + "/")
        time.sleep(2)

        assert d.current_url.rstrip("/").endswith("/ms/login"), \
            f"não redirecionou para login MS: {d.current_url}"
        assert _esta_ms_login(d), "não detectou a página MS"
        assert _precisa_login(d), "não reconheceu necessidade de login"

        ok = tentar_login_automatico(d, cfg)
        print("  login automático retornou:", ok)
        assert ok, "login automático falhou"
        assert _achar_por_texto(d, "Consulta Incidência") is not None, \
            "não chegou ao app do Operview (fake)"
        print("[PASS] 1. login automático Microsoft/Entra completo até o app")

        # ── Cenário 2: sem credenciais → fallback para manual ──
        d2 = _novo_driver(base)
        drivers.append(d2)
        d2.get(base + "/")
        time.sleep(2)
        ok2 = tentar_login_automatico(d2, OperviewConfig(url=base))
        print("  sem credenciais retornou:", ok2)
        assert ok2 is False, "sem credenciais deveria voltar para o fluxo manual"
        print("[PASS] 2. sem credenciais → retorna False (login manual)")

        # ── Cenário 3: app já aberto (sessão ativa) → não é página MS ──
        d3 = _novo_driver(base)
        drivers.append(d3)
        d3.get(base + "/app")
        time.sleep(1)
        d3.add_cookie({"name": "sessao_fake", "value": "1", "path": "/"})
        d3.get(base + "/app")
        time.sleep(1.5)
        assert _achar_por_texto(d3, "Consulta Incidência") is not None, \
            "app deveria abrir direto com sessão ativa"
        assert _precisa_login(d3) is False, "com sessão ativa não pede login"
        ok3 = tentar_login_automatico(d3, OperviewConfig(url=base, login="x", senha="y"))
        print("  sessão ativa retornou:", ok3)
        assert ok3 is True, "página do app (já autenticada) → segue sem login manual"
        print("[PASS] 3. sessão ativa → app entra direto, sem login automático")

        # ── Cenário 4: MFA exigido → retorna False (volta ao fluxo manual) ──
        _MsHandler.mode = "mfa"
        d4 = _novo_driver(base)
        drivers.append(d4)
        d4.get(base + "/")
        time.sleep(2)
        ok4 = tentar_login_automatico(
            d4, OperviewConfig(url=base, login="u@enel.com", senha="SenhaFake#123"))
        print("  com MFA retornou:", ok4)
        assert ok4 is False, "com MFA não deve confirmar login automático"
        _MsHandler.mode = "kmsi"
        print("[PASS] 4. MFA exigido → retorna False (login manual)")

        # ── Cenário 5: loop (60 min) — sessão salva/restaurada entre fluxos ──
        import json
        import shutil

        from core.downloader import operview as ov

        sessao_real = ov._SESSAO_OPERViEW
        ov._SESSAO_OPERViEW = os.path.join(
            tempfile.mkdtemp(prefix="opv_test_sessao_"), "sessao.json")
        try:
            d5 = _novo_driver(base)
            drivers.append(d5)
            d5.get(base + "/")
            time.sleep(2)
            assert tentar_login_automatico(
                d5, OperviewConfig(url=base, login="u@enel.com", senha="SenhaFake#123"))
            _salvar_sessao(d5)
            d5.quit()
            drivers.pop()

            assert os.path.exists(ov._SESSAO_OPERViEW), "sessão não foi salva em disco"

            # 2ª execução (simula o próximo fluxo de 60min) → restaura sem login
            d6 = _novo_driver(base)
            drivers.append(d6)
            restaurado = _restaurar_sessao(d6, base + "/")
            print("  restauração de sessão retornou:", restaurado)
            assert restaurado is True, "sessão deveria ser restaurada (loop sem login)"
            assert _achar_por_texto(d6, "Consulta Incidência") is not None
            assert _precisa_login(d6) is False
            assert _esta_ms_login(d6) is False
            print("[PASS] 5. loop 60min → sessão salva/restaurada, sem pedir login")
        finally:
            ov._SESSAO_OPERViEW = sessao_real

        # ── Cenário 6: shell Angular antes do redirect (bug real) ──
        # Replica o que o usuário reportou: a página abre um shell sem form de
        # login e o check antigo (~/lento: sem password, sem MS, sem "Entrar")
        # achava que já estava logado. O novo _aguardar_estado_login responde
        # "indefinido" → o fluxo trata como login necessário (não pula).
        d7 = _novo_driver(base)
        drivers.append(d7)
        d7.get(base + "/lento")
        time.sleep(1)
        assert _precisa_login(d7) is False, \
            "pré-condição: shell sem form não é detectado por _precisa_login"
        est = _aguardar_estado_login(d7, timeout=6)
        print("  estado da página shell (sem redirect):", est)
        assert est == "indefinido", "shell deveria ser 'indefinido' (não logado!)"
        print("[PASS] 6. shell Angular sem redirect → 'indefinido' (não pula login)")

        # ── Cenário 7: estado 'ok' no app autenticado ──
        d8 = _novo_driver(base)
        drivers.append(d8)
        d8.get(base + "/app")
        time.sleep(1)
        d8.add_cookie({"name": "sessao_fake", "value": "1", "path": "/"})
        d8.get(base + "/app")
        time.sleep(1.5)
        est8 = _aguardar_estado_login(d8, timeout=6)
        print("  estado do app autenticado:", est8)
        assert est8 == "ok", "app autenticado deveria ser 'ok'"
        print("[PASS] 7. app autenticado → estado 'ok'")

        # ── Cenário 8: estado 'login' na página da Microsoft ──
        d9 = _novo_driver(base)
        drivers.append(d9)
        d9.get(base + "/")
        time.sleep(2)
        est9 = _aguardar_estado_login(d9, timeout=6)
        print("  estado da página MS:", est9)
        assert est9 == "login", "página MS deveria ser 'login'"
        print("[PASS] 8. página da Microsoft → estado 'login'")

        print("\n✅ TODOS OS TESTES PASSARAM")
    finally:
        for d in drivers:
            try:
                d.quit()
            except Exception:
                pass
        srv.join(timeout=5)


if __name__ == "__main__":
    main()