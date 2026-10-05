import json
import os
from pathlib import Path
from dataclasses import dataclass, field, asdict

CONFIG_DIR = Path(os.environ["APPDATA"]) / "DataHub"
CONFIG_FILE = CONFIG_DIR / "config.json"


@dataclass
class GeoOnlineConfig:
    # Mantido dormente: código de download existe mas a UI não expõe mais.
    url: str = "http://geonline.enelint.global/#/"
    conta: str = ""
    login: str = ""
    senha: str = ""


@dataclass
class OperviewConfig:
    url: str = "https://operview-rj.enel.com/"
    login: str = ""
    senha: str = ""


@dataclass
class SpotfireConfig:
    # M300 reside na mesma análise (/M300/Scanner 4.0 - RJ); uma só URL.
    url: str = ""
    login: str = ""
    senha: str = ""


@dataclass
class SharePointConfig:
    enabled: bool = False
    pasta: str = ""


@dataclass
class N8nConfig:
    enabled: bool = False
    webhook: str = ""


@dataclass
class AppConfig:
    geonline: GeoOnlineConfig = field(default_factory=GeoOnlineConfig)
    operview: OperviewConfig = field(default_factory=OperviewConfig)
    spotfire: SpotfireConfig = field(default_factory=SpotfireConfig)
    sharepoint: SharePointConfig = field(default_factory=SharePointConfig)
    n8n: N8nConfig = field(default_factory=N8nConfig)
    # Fontes selecionáveis no fluxo (checkbox). Apenas Operview está ativa
    # (M300 removido das configurações do DataHub).
    fontes: list[str] = field(default_factory=lambda: ["operview"])
    origem_relatorio: str = "geonline"  # OBSOLETO: migrado p/ `fontes`
    pasta_local: str = ""
    webhook_n8n: str = ""
    base_mensal_enabled: bool = True
    tema_claro: bool = False


def _from_dict(cls, data: dict):
    fields = {}
    for f in cls.__dataclass_fields__:
        val = data.get(f)
        if val is None:
            continue
        ftype = cls.__dataclass_fields__[f].type
        if hasattr(ftype, "__dataclass_fields__"):
            fields[f] = _from_dict(ftype, val)
        else:
            fields[f] = val
    return cls(**fields)


def load_config() -> AppConfig:
    if not CONFIG_FILE.exists():
        return AppConfig()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        cfg = AppConfig()
        cfg.geonline = _from_dict(GeoOnlineConfig, data.get("geonline", {}))
        cfg.operview = _from_dict(OperviewConfig, data.get("operview", {}))
        cfg.spotfire = _from_dict(SpotfireConfig, data.get("spotfire", {}))

        # Migração: campo antigo `origem_relatorio` -> `fontes`
        if "fontes" in data and isinstance(data["fontes"], list):
            cfg.fontes = [str(f) for f in data["fontes"] if f == "operview"]
        elif data.get("origem_relatorio") == "operview":
            cfg.fontes = ["operview"]
        elif data.get("origem_relatorio") == "geonline":
            cfg.fontes = ["operview"]  # geonline está dormente; cai no operview
        cfg.origem_relatorio = data.get("origem_relatorio", "geonline")

        cfg.sharepoint = _from_dict(SharePointConfig, data.get("sharepoint", {}))
        cfg.n8n = _from_dict(N8nConfig, data.get("n8n", {}))
        cfg.pasta_local = data.get("pasta_local", "")
        cfg.webhook_n8n = data.get("webhook_n8n", "")
        cfg.base_mensal_enabled = data.get("base_mensal_enabled", True)
        cfg.tema_claro = data.get("tema_claro", False)
        return cfg
    except Exception:
        return AppConfig()


def save_config(cfg: AppConfig) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(asdict(cfg), f, ensure_ascii=False, indent=2)
