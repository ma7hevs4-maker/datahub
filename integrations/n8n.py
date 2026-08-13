import json
import urllib.request


def enviar(payload: dict, webhook_url: str, log_fn=print) -> None:
    if not webhook_url:
        return
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=data, headers={"Content-Type": "application/json"}
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status == 200:
                log_fn("  ✅ Payload enviado ao n8n")
            else:
                log_fn(f"  ⚠️  n8n respondeu {resp.status}")
    except Exception as e:
        log_fn(f"  ❌ Erro ao enviar ao n8n: {e}")
