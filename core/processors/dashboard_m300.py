# -*- coding: utf-8 -*-
"""Gera o dashboard M300 (Scanner 4.0 - RJ) embutindo os dados tratados.

O dashboard é um app React (Lovable) que, originalmente, lê o .xlsx via
upload manual de arquivo. Para comportar-se "como o Operview" (sem upload
manual), embutimos o .xlsx como base64 e injetamos um efeito de auto-load
que o decodifica e alimenta o estado interno (wbRef/parseXLSX/applyParsed).

Template: dashboard_m300_template.html (cópia de dashboard m300/public/dashboard.html)
"""
import base64
import io
import time
import os
import shutil

from pathlib import Path


def _write_retry(saida, content, encoding="utf-8", tries=6, delay=1.0):
    saida = Path(saida)
    last = None
    for _ in range(tries):
        try:
            tmp = saida.with_suffix(saida.suffix + ".tmp")
            tmp.write_text(content, encoding=encoding)
            try:
                os.replace(str(tmp), str(saida))
            except OSError:
                shutil.copy2(str(tmp), str(saida))
            if tmp.exists():
                try:
                    tmp.unlink()
                except OSError:
                    pass
            return
        except (OSError, PermissionError) as e:
            last = e
            time.sleep(delay)
    raise last


# Efeito React injetado: decodifica o base64 e alimenta o dashboard no mount.
_AUTOLOAD_EFFECT = (
    "useEffect(()=>{try{const __b64=window.__M300_B64__;if(!__b64)return;"
    "const __bin=atob(__b64);const __bytes=new Uint8Array(__bin.length);"
    "for(let __i=0;__i<__bin.length;__i++)__bytes[__i]=__bin.charCodeAt(__i);"
    "const __wb=XLSX.read(__bytes,{type:'array',cellDates:true});"
    "wbRef.current=__wb;setSelDays([]);"
    "const __parsed=parseXLSX(__wb);"
    "if(__parsed&&Object.keys(__parsed.kpi||{}).length){"
    "applyParsed(__parsed);setAllDates(__parsed.allDates||[]);"
    "const __ms=Object.keys(__parsed.kpi||{}).sort((a,b)=>parseInt(a)-parseInt(b));"
    "setSelectedMes(__ms[__ms.length-1]);"
    "setUploadMsg('✅ Base embutida — '+__ms.length+' mês(es)');"
    "}else{setUploadMsg('⚠️ Base embutida sem dados reconhecíveis.');}"
    "}catch(__e){console.error('Auto-load M300 falhou',__e);"
    "setUploadMsg('⚠️ Falha ao carregar base embutida: '+__e.message);}},[]);\n"
)

_ANCHOR = "// re-processa a base quando o filtro de dias muda"


def gerar_dashboard_m300(df, saida=None, log_fn=print):
    if saida is None:
        saida = Path.cwd() / "m300_dashboard.html"
    saida = Path(saida)
    saida.parent.mkdir(parents=True, exist_ok=True)

    tpl = Path(__file__).parent / "dashboard_m300_template.html"
    if not tpl.exists():
        raise FileNotFoundError(f"Template do dashboard M300 não encontrado: {tpl}")

    html = tpl.read_text(encoding="utf-8")

    # Serializa o DataFrame tratado para .xlsx em memória -> base64
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")

    if _ANCHOR not in html:
        raise RuntimeError("Anchor do template M300 não encontrado (dashboard alterado?).")
    html = html.replace(_ANCHOR, _AUTOLOAD_EFFECT + _ANCHOR, 1)

    html = html.replace(
        "</head>",
        f'<script>window.__M300_B64__="{b64}";</script>\n</head>',
        1,
    )

    _write_retry(saida, html)
    log_fn(f"  💾 Dashboard M300 gerado: {saida} ({len(df)} linhas)")
    return str(saida)
