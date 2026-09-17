"""Simulador local da página de login da Microsoft/Entra (para testes).

Serve um fluxo idêntico em estrutura ao real (mesmos seletores que o
operview.py usa):
  GET  /            → sem sessão: 302 → /ms/login ; com sessão: app
  GET  /ms/login    → passo 1: input[name=loginfmt] + #idSIButton9
  POST /ms/next     → passo 2: input[name=passwd]   + #idSIButton9
  POST /ms/signin   → passo 3: #KmsiCheckboxField   + #idSIButton9 (Sim)
  POST /ms/kmsi     → grava cookie e 302 → /
  GET  /app         → página do app (contém 'Consulta Incidência')

Uso:
    from fake_ms_server import criar_servidor
    srv = criar_servidor(porta)
    srv.daemon = True
    srv.start()
    ...
    srv.shutdown()
"""
import http.server
import socketserver
import threading

LOGIN_PAGE = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Entrar</title></head>
<body>
<div id="login">
  <h2>Entrar</h2>
  <form method="post" action="/ms/next">
    <label for="i0116">E-mail corporativo</label>
    <input type="email" id="i0116" name="loginfmt" placeholder="nome@enel.com">
    <input type="submit" id="idSIButton9" value="Próxima">
  </form>
</div>
</body></html>"""

PASSWORD_PAGE = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Entrar</title></head>
<body>
<div id="login">
  <h2>Digite a senha</h2>
  <form method="post" action="/ms/signin">
    <input type="password" id="i0118" name="passwd">
    <input type="submit" id="idSIButton9" value="Entrar">
  </form>
</div>
</body></html>"""

KMSI_PAGE = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Entrar</title></head>
<body>
<div id="login">
  <h2>Continuar conectado?</h2>
  <form method="post" action="/ms/kmsi">
    <input type="checkbox" id="KmsiCheckboxField" name="kmsi" checked>
    <span>Não mostrar novamente</span>
    <input type="submit" id="idSIButton9" value="Sim">
  </form>
</div>
</body></html>"""

MFA_PAGE = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Entrar</title></head>
<body>
<div id="login">
  <h2>Aprovação necessária</h2>
  <p>Mais informações necessárias. Aprovar a notificação no aplicativo Microsoft Authenticator.</p>
  <div id="mfa-ok">MFA em andamento</div>
</div>
</body></html>"""

APP_PAGE = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>OPerview RJ</title></head>
<body>
  <span>Consulta Incidência</span>
  <div id="app-ok">app autenticado</div>
</body></html>"""

LENTO_PAGE = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>OPerview RJ</title></head>
<body>
  <div id="shell">Carregando aplicação...</div>
</body></html>"""


class _MsHandler(http.server.BaseHTTPRequestHandler):
    # "kmsi" = fluxo normal (manter conectado); "mfa" = exige MFA (fallback manual)
    mode = "kmsi"

    def _send(self, code, body="", ctype="text/html; charset=utf-8", headers=None):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _tem_sessao(self):
        return "sessao_fake=1" in (self.headers.get("Cookie") or "")

    def do_GET(self):
        if self.path in ("/", "/app"):
            if self._tem_sessao():
                self._send(200, APP_PAGE)
            else:
                self._send(302, "", headers={"Location": "/ms/login"})
        elif self.path == "/ms/login":
            self._send(200, LOGIN_PAGE)
        elif self.path == "/lento":
            self._send(200, LENTO_PAGE)
        else:
            self._send(404, "<h1>404</h1>")

    def do_POST(self):
        if self.path == "/ms/next":
            self._send(200, PASSWORD_PAGE)
        elif self.path == "/ms/signin":
            body = MFA_PAGE if self.mode == "mfa" else KMSI_PAGE
            self._send(200, body)
        elif self.path == "/ms/kmsi":
            self._send(302, "", headers={
                "Location": "/app",
                "Set-Cookie": "sessao_fake=1; Path=/",
            })
        else:
            self._send(404, "<h1>404</h1>")

    def log_message(self, *a):
        pass


class _MsServer(socketserver.ThreadingTCPServer):
    daemon_threads = True

    def handle_error(self, request, client_address):
        pass  # ignora o ConnectionResetError do Chrome ao fechar conexões


def criar_servidor(porta):
    srv = _MsServer(("127.0.0.1", porta), _MsHandler)
    return threading.Thread(target=srv.serve_forever, daemon=True)


if __name__ == "__main__":
    srv = criar_servidor(8791)
    srv.start()
    print("fake MS login em http://127.0.0.1:8791/  (Ctrl+C p/ parar)")
    try:
        threading.Event().wait()
    except KeyboardInterrupt:
        pass