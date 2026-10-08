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

Bug real corregido (2026-09-21, pedido explícito del usuario -- "apenas te
sales de esa pantalla se cancela la tarea... cuando salgo y vuelvo a entrar
al modulo donde se estaba cargando aparece como vacío sin ningún proceso"):
el trabajo en sí SIEMPRE siguió corriendo bien en su hilo (nada de esto
cambió) -- lo que se perdía era el `job_id`, que vivía solo en una variable
de JS de la pestaña que hizo el POST (ver el IIFE de background-job-form en
base.html). Al navegar a otro lado esa pestaña se destruye junto con su
`job_id` y no hay forma de "reengancharse" -- el trabajo sigue vivo en este
módulo, simplemente nadie lo está mirando más. `get_active_job(kind)` +
`acknowledge_job` resuelven esto: cada job se crea con un `kind` (ej.
"combustible", ver los `create_job(..., kind=...)` en webapp.py), y la
página de carga de ese `kind` consulta si hay un job activo (corriendo, o
terminado pero sin mostrar todavía) y se lo pasa al JS para que retome el
sondeo apenas carga la página, sin que el usuario tenga que quedarse ni
volver a subir nada.
"""

import threading
import time
import uuid

_jobs = {}
_lock = threading.Lock()


class JobCancelled(BaseException):
    """
    La carga se canceló desde la página (cancel_job). La levanta update_job
    (o increment_done) en el hilo de la carga, así corta en su próximo aviso
    de avance sin que cada _run_*_job tenga que chequear nada. Es
    BaseException, como KeyboardInterrupt, para que los `except Exception`
    que aíslan cada archivo no la frenen; el hilo termina con ella y
    _ignore_cancelled_jobs la deja pasar en silencio.
    """

    def __init__(self, job_id):
        super().__init__(job_id)
        self.job_id = job_id


_previous_excepthook = threading.excepthook


def _ignore_cancelled_jobs(args):
    if isinstance(args.exc_value, JobCancelled):
        return
    _previous_excepthook(args)


threading.excepthook = _ignore_cancelled_jobs

# Los jobs terminados (done/error) y ya reconocidos (acknowledged, ver
# acknowledge_job) se podan pasada esta antigüedad -- nada más para no
# crecer sin límite en un server de mucho uptime, no hace falta más
# precisión que esto.
_MAX_JOB_AGE_SECONDS = 6 * 3600


def _prune_old_jobs_locked():
    """Se llama con _lock ya tomado -- ver _MAX_JOB_AGE_SECONDS arriba."""
    cutoff = time.time() - _MAX_JOB_AGE_SECONDS
    stale = [
        job_id for job_id, job in _jobs.items()
        if job["status"] != "running" and job.get("acknowledged") and job["created_at"] < cutoff
    ]
    for job_id in stale:
        del _jobs[job_id]


def create_job(total, kind=None):
    """
    `kind` (ej. "combustible", "gettel_pagos") identifica de qué página de
    carga es este job -- lo usa get_active_job para poder "reengancharse"
    a un job en progreso cuando el usuario vuelve a entrar a esa página
    (ver el docstring del módulo). Puede quedar None para un job que no
    necesita esto.
    """
    job_id = uuid.uuid4().hex
    with _lock:
        _prune_old_jobs_locked()
        _jobs[job_id] = {
            "id": job_id,
            "kind": kind,
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
            # Ver acknowledge_job -- true una vez que el cliente ya mostró
            # el resultado final (notice/redirect) de este job, para que
            # get_active_job deje de devolverlo en la próxima carga de
            # página (si no, un "listo" viejo reaparecería cada vez que el
            # usuario reabre la página, para siempre).
            "acknowledged": False,
            "created_at": time.time(),
            # Último aviso de avance (has_running_jobs) y cancelación pedida
            # desde la página (cancel_job): `stopping` sigue en True hasta
            # que el hilo efectivamente corta.
            "updated_at": time.time(),
            "cancel_requested": False,
            "stopping": False,
        }
    return job_id


def _touch_locked(job, fields=None):
    """
    Se llama con _lock tomado: corta la carga si se canceló (ver
    JobCancelled). Antes de cortar deja en `cancel_note` lo que ya había
    quedado guardado, para que la página lo avise (revisión 2026-10-08: con
    un solo archivo, la carga se guardaba entera y la página decía
    "cancelada" sin más).
    """
    if job["cancel_requested"]:
        fields = fields or {}
        if fields.get("status") in ("done", "error"):
            final = fields.get("notice") or fields.get("error")
            job["cancel_note"] = ("La carga ya había terminado cuando se canceló" +
                                  (": " + final if final else "; lo leído quedó guardado."))
        else:
            saved = fields.get("done", job.get("done") or 0)
            if saved:
                job["cancel_note"] = (f"Se canceló: {saved} de {job.get('total') or saved} archivo(s) ya habían "
                                      "quedado cargados.")
        job["stopping"] = False
        raise JobCancelled(job["id"])
    job["updated_at"] = time.time()


def update_job(job_id, **fields):
    with _lock:
        if job_id in _jobs:
            _touch_locked(_jobs[job_id], fields)
            _jobs[job_id].update(fields)


def increment_done(job_id, done, total):
    with _lock:
        if job_id in _jobs:
            _touch_locked(_jobs[job_id], {"done": done})
            _jobs[job_id]["done"] = done
            _jobs[job_id]["total"] = total


def cancel_job(job_id):
    """
    Cancela una carga en curso (pedido del usuario, 2026-10-06: "un botón de
    cancelar carga para así por si te faltó algo o te arrepentís no tenés que
    esperar"). Para la página queda terminada al toque (status "cancelled",
    ya reconocida); el hilo corta en su próximo aviso de avance, o sea al
    terminar el archivo que está leyendo. Lo que ya se guardó queda
    guardado. Devuelve una copia del job, o None si no existe.
    """
    with _lock:
        job = _jobs.get(job_id)
        if job is None:
            return None
        if job["status"] == "running":
            job.update(status="cancelled", cancel_requested=True, stopping=True, acknowledged=True)
        return dict(job)


# Una carga sin avisar avance en este tiempo se da por trabada: no frena el
# reinicio del servidor (has_running_jobs).
_MAX_IDLE_SECONDS = 30 * 60


def has_running_jobs():
    """
    True si alguna carga está trabajando (o terminando el archivo que leía
    antes de cancelarse). Con el reloader de debug, webapp.py espera a que
    esto dé False antes de reiniciar el servidor por un cambio de código.
    """
    now = time.time()
    with _lock:
        return any(
            (job["status"] == "running" or job.get("stopping"))
            and now - job.get("updated_at", job["created_at"]) < _MAX_IDLE_SECONDS
            for job in _jobs.values()
        )


def get_job(job_id):
    with _lock:
        job = _jobs.get(job_id)
        return dict(job) if job is not None else None


def get_active_job(kind):
    """
    Job más reciente de este `kind` que todavía tiene algo para mostrar --
    corriendo, o terminado pero sin acknowledge (ver el docstring del
    módulo). None si no hay ninguno. Pensado para que la ruta GET de cada
    página de carga se lo pase al template (`resume_job_id`) y el JS
    retome el sondeo apenas carga la página.
    """
    with _lock:
        candidates = [
            job for job in _jobs.values()
            if job.get("kind") == kind and (job["status"] == "running" or not job.get("acknowledged"))
        ]
        if not candidates:
            return None
        latest = max(candidates, key=lambda j: j["created_at"])
        return dict(latest)


def acknowledge_job(job_id):
    """El cliente ya mostró el resultado final de este job -- ver get_active_job."""
    with _lock:
        if job_id in _jobs:
            _jobs[job_id]["acknowledged"] = True
