import threading
import time
from typing import Callable


class Scheduler:
    """
    Executa uma função em loop com intervalo configurável.
    Suporta: uma vez / 30 min / 60 min.
    Thread-safe: pode ser parado a qualquer momento via stop().
    Cada iteração é completamente independente — sem estado residual.
    """

    INTERVALOS = {
        "uma_vez": None,
        "30min": 30 * 60,
        "60min": 60 * 60,
    }

    def __init__(self, fn: Callable, modo: str, log_fn: Callable = print):
        if modo not in self.INTERVALOS:
            raise ValueError(f"Modo inválido: {modo}. Use: {list(self.INTERVALOS)}")
        self._fn = fn
        self._modo = modo
        self._intervalo = self.INTERVALOS[modo]
        self._log = log_fn
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._iteracao = 0

    def start(self) -> None:
        self._stop_event.clear()
        self._iteracao = 0
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()

    @property
    def rodando(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def _loop(self) -> None:
        while not self._stop_event.is_set():
            self._iteracao += 1
            if self._iteracao > 1:
                self._log(f"\n🔄 ══ Iteração #{self._iteracao} ══")
            try:
                self._fn()
            except Exception as e:
                self._log(f"❌ Erro no fluxo: {e}")

            if self._intervalo is None:
                break

            self._log(f"⏳ Próxima execução em {self._intervalo // 60} minutos...")
            # Dorme em blocos de 5s para responder ao stop() rapidamente
            for _ in range(self._intervalo // 5):
                if self._stop_event.is_set():
                    break
                time.sleep(5)

        self._log("🏁 Scheduler encerrado.")
