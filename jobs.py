"""
Registro en memoria de trabajos en segundo plano -- pensado para las cargas
de varios PDF a la vez que tardan bastante (Reporte Diario es el caso
puntual que motivó esto, ver CLAUDE.md). Pedido explícito del usuario
(2026-09-14): "ver si ese proceso podria hacerse por aparte mientras se
trabaja en otra cosa" (no pidió que sea más rápido, solo que no obligue al
usuario a quedarse parado en la pestaña) + "que muestre el progreso real...
que vaya moviendo el porcentaje".

Un trabajo corre en un `threading.Thread` propio, separado del hilo que
recibió el POST -- la request original devuelve un `job_id` de inmediato
(sin esperar a que termine el procesamiento) y el cliente sondea el estado
por su cuenta, así que cerrar la pestaña o navegar a otro lado no corta el
trabajo real (a diferencia del mecanismo `ajax-process-form` de siempre,
donde la conexión XHR abierta ES el procesamiento). Vive solo en memoria
(se pierde si el servidor se reinicia) -- alcanza para esto: no es una cola
persistente, es nada más "no bloquees la pestaña del usuario".
"""

import threading
import time
import uuid

_jobs = {}
_lock = threading.Lock()


def create_job(total):
    job_id = uuid.uuid4().hex
    with _lock:
        _jobs[job_id] = {
            "id": job_id,
            "total": total,
            "done": 0,
            "status": "running",  # running | done | error
            # "parsing" (leyendo los PDF, la parte con progreso real) ->
            # "saving" (escribiendo el Excel + guardando en la página, sin
            # progreso propio) -- separado para que la barra no se quede
            # pegada en "100%" mostrando el mismo texto varios minutos
            # mientras el guardado-espejo todavía sigue trabajando (pedido
            # explícito del usuario, ver CLAUDE.md "sexta tanda": la barra
            # tiene que reflejar lo que realmente está pasando).
            "phase": "parsing",
            "notice": None,
            "notice_level": "warning",
            "error": None,
            "result_path": None,
            "result_filename": None,
            "created_at": time.time(),
        }
    return job_id


def update_job(job_id, **fields):
    with _lock:
        if job_id in _jobs:
            _jobs[job_id].update(fields)


def increment_done(job_id, done, total):
    with _lock:
        if job_id in _jobs:
            _jobs[job_id]["done"] = done
            _jobs[job_id]["total"] = total


def get_job(job_id):
    with _lock:
        job = _jobs.get(job_id)
        return dict(job) if job is not None else None
