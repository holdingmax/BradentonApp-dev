# Plan: pasar BradentonApp de SQLite a PostgreSQL

**16/10/2026 · chat 34 (el usuario lo llamó "chat 28") · plan, sin tocar código.** Se trabaja solo en el repo dev (`holdingmax/BradentonApp-dev`); producción no se toca. Cada módulo se hace recién después de que se confirme, sin commit ni push hasta que se pida.

**Revisado contra la guía de Leti** ("Guía general: cómo programamos acá con Claude Code", secciones 4, 5, 8, 9, 10 y 11) **y su guía de Render** ("cómo crear el servicio web y la base de datos"). La guía está pensada para Next.js + Prisma; acá la app es Python + Flask, así que Prisma pasa a ser Alembic (sección 9) y el build de Render es el `Dockerfile` (sección 10).

**Paso 1 de la guía, hecho:** respaldo completo en `C:\BradentonApp_respaldos\2026-10-16_antes_postgres\` (`reportes_data` entero y los 6 archivos de configuración: 603 archivos, 406 MB; las 19 bases pasan `integrity_check`).

**Paso 3 de la guía, verificado:** `.env` está en `.gitignore` y en `.dockerignore`, y nunca se subió a git. También están ignorados `reportes_data/`, `users.json` y `.flask_secret_key`. Sugerencia: agregar `.env.*`, por si se crea un `.env.testing`.

## Avance

| Paso | Estado |
|---|---|
| Paso 0: `db.py`, Alembic, Postgres local, copia y comparador | Hecho el 16/10, commit `c35300a` |
| Módulo 1: jh_facturas | Hecho el 16/10, commit `d515efc`: copia idéntica (10 filas) y 15 comparaciones idénticas |
| Módulo 2: ice_machine | Hecho el 16/10: copia idéntica (8 Payment Summary) y comparaciones idénticas (Depósitos jul-oct, antes y después de alta / reemplazo / baja). El `INSERT OR REPLACE` pasó a `ON CONFLICT … DO UPDATE` con todas las columnas |
| Módulo 3: reporte_mensual | Hecho el 16/10: copia idéntica (2 reportes, 59 departamentos) y 57 comparaciones idénticas (Reporte C-Store, Reporte Diario, historiales, Cierre y su PDF, Control CMV; antes y después de guardar, corregir y borrar un reporte). Los 2 `INSERT OR REPLACE` pasaron a `ON CONFLICT`; los departamentos se ordenan por importe y, si empatan, por nombre |
| Módulo 4: cheques | Hecho el 16/10: copia idéntica (31 cheques) y comparaciones idénticas (cuadro de Cheques, antes y después de cargar a mano, corregir, completar lo automático y borrar). Los 2 `lastrowid` pasaron a `RETURNING id` |
| Módulo 5: control_cmv | Hecho el 16/10: copia idéntica (4 reportes de Elistar) y 9 comparaciones idénticas (Control CMV jul-oct, antes y después de guardar, reemplazar y borrar). El `INSERT OR REPLACE` pasó a `ON CONFLICT` |
| Módulo 6: depositos | Hecho el 16/10: copia idéntica (84 recibos) y 25 comparaciones idénticas (Depósitos, Control Depósitos y Chase jul-oct, antes y después de cargar, detectar duplicado, cambiar tipo, editar y borrar). `lastrowid` → `RETURNING id`; los recibos sin N° van primero en el día (`NULLS FIRST`, como SQLite) |
| Módulo 7: horas_trabajo | Hecho el 16/10: copia idéntica (semanas y empleados) y 11 comparaciones idénticas (Horas e historial jul-oct, antes y después de cargar, volver a subir, agregar, editar y borrar). Los 2 `lastrowid` → `RETURNING id`; los nombres conocidos para el OCR salen en el orden de la primera aparición, como el DISTINCT de SQLite |
| Módulo 8: jh_mensual | Hecho el 16/10: copia idéntica (9 reportes) y 61 comparaciones idénticas (J.H., historiales de EFT y Cupones, PDF de EFT, Tarjetas, Cierre, botonera; antes y después de sumar renglones y borrar). `INSERT OR REPLACE` → `ON CONFLICT`; los renglones de todos los meses salen de mes en mes, como en SQLite |
| Módulo 9: documents | Hecho el 16/10: migrada igual (43 filas, guardado apagado) para que no quede ningún módulo en SQLite. Comparaciones idénticas del buscador y de un escenario con el guardado prendido solo en la copia (listar, buscar, guardar, leer, borrar). `lastrowid` → `RETURNING id`; la búsqueda por nombre no distingue mayúsculas en los dos (`LOWER … LIKE LOWER`) |
| Módulo 10: fisico | Hecho el 16/10: copia idéntica (físico, facturas de combustible y renglones) y 29 comparaciones idénticas (Físico y sus exportaciones jul-oct, Próximos EFT, botonera EFT; antes y después de cargar, buscar, editar y borrar una factura y fijar los ajustes del mes). `last_insert_rowid()` → `RETURNING id`; `COLLATE NOCASE` → `LOWER() = LOWER()` |
| Módulo 11: caja | Hecho el 16/10: copia idéntica y 41 comparaciones idénticas (Caja ene-oct, sus exportaciones y su PDF, Control Caja viejo, botonera; antes y después de gastos con N° de ticket como número, repetidos, pendientes y saldos del mes). Los 2 `lastrowid` → `RETURNING id`; el gasto repetido se busca con `ORDER BY id` |
| Módulo 12: gettel_toyota | Hecho el 16/10: copia idéntica y 49 comparaciones idénticas (Gettel, historial, su Excel y PDF, Kia y Toyota, botonera; antes y después de cargar días, pagos y el ajuste del mes; julio redirige igual en los dos porque no tiene datos). `last_insert_rowid()` → `RETURNING id`, `COLLATE NOCASE` → `LOWER()`, `printf` → `MIN(year * 100 + month)` |
| Módulo 13: lottery | Hecho el 16/10: copia idéntica y 81 comparaciones idénticas en los dos motores **y contra el código original** (Lottery, días, historial, exportaciones, mensual, resumen, PDF, Control Lottery, Caja; antes y después de cargar un día por sus tres caminos, fechas de pago con reajuste y reporte mensual) |
| Módulo 14: reportes_diarios | Hecho el 16/10: copia idéntica (1.924 filas) y 135 comparaciones idénticas en los dos motores y contra el código original (días, Reporte C-Store, historiales, Store Info y sus exportaciones, PDF, Tarjetas, Cierre y sus exportaciones, Control CMV, Kia y Toyota, Caja, botonera; antes y después de cargar un día completo, corregir, márgenes y bajas). La suma de decimales de Postgres ahora es la de SQLite (migración `0015`) |
| Módulo 15: cuenta_kia_toyota | Hecho el 16/10: copia idéntica (2.173 filas) y 27 comparaciones idénticas en los dos motores y contra el código original (Kia y Toyota jul-oct, botonera, Gettel; antes y después de un Excel que reemplaza días, pagos nuevos y repetidos, POS y asignaciones a mano). Los `INSERT OR REPLACE` pasaron a borrar + insertar, para conservar el orden de carga de SQLite que recorre la cuenta corriente |
| Módulo 16: eft | Hecho el 16/10: copia idéntica (3.432 filas) y 79 comparaciones idénticas en los dos motores y contra el código original (EFT y Cupones, historiales, J.H., Facturas varias, PDF de EFT, Próximos EFT, Tarjetas, Cierre, botonera; antes y después de un EFT copiado de uno real, DDC a mano, cupones nuevos y repetidos y detalle de cupones). Arreglado un `MAX()` con columnas sueltas que Postgres no acepta |
| Módulo 17: chase | Hecho el 16/10: copia idéntica (636 movimientos, con su rowid) y 89 comparaciones idénticas en los dos motores y contra el código original (Chase, historial, exportación, PDF, Cheques, Proveedores, Control Depósitos, Caja, Lottery, Tarjetas, Cierre, botonera; antes y después de **volver a subir el extracto real de octubre por la página**, correcciones a mano y recategorizar). La tabla lleva en Postgres una columna `rowid` con el mismo valor que en SQLite, así los 13 `rowid` del código no cambian |
| Módulo 18: cmv | Hecho el 16/10: copia idéntica (14.539 filas) y 27 comparaciones idénticas en los dos motores y contra el código original (CMV, historiales de costo y ventas, comparación de fotos, Control CMV, Ventas y costos, lista de productos; antes y después de recargar el CMV entero con precios cambiados, foto del día y ventas de un departamento cargadas dos veces) |
| Módulo 19: proveedores | Hecho el 16/10: copia idéntica (12.717 filas) y 59 comparaciones idénticas en los dos motores y contra el código original (proveedores guardados y la cuenta corriente de 10 proveedores, historial y PDF mensual, Productos, una factura de H.T. y su reporte de precios en PDF y Excel, lista, Ventas y costos, Cheques; antes y después de facturas nuevas y repetidas, renglones, ocultar, configuración, credit memos y pagos). La configuración inicial de proveedores la carga la migración y la copia la reemplaza por la de SQLite |
| Revisión de órdenes | Hecha el 16/10: todo ORDER BY que podía empatar desempata como lo hacía SQLite (por el índice que usaba o por el orden de carga). Corrida final de los 19 módulos juntos: 523 de 523 comparaciones idénticas entre SQLite y Postgres y, con los últimos arreglos, iguales al código de antes del pase |
| Respaldo (punto 8) | Hecho el 16/10: `respaldo_postgres.py` (pg_dump a `backups/`, verificado con pg_restore; restaurado de prueba en otra base) |
| Dockerfile (punto 9, parte) | Hecho el 16/10: zona horaria de Florida (J6), salida sin buffer, `.dockerignore` sin `Claude outputs/` ni respaldos |
| Punto 5: usuarios, clave de sesión y reglas | Hecho el 16/10: `config_store.py` y las tablas del esquema `config` (migración `0023_config`); la copia de datos también pasa la configuración (`migrar_sqlite_a_postgres.py config`); 37 comparaciones idénticas en los dos motores y contra el original |
| Pendientes | Prueba en Linux limpio con la base vacía (punto 9); A1/A2 adaptados; copia a testing |

**Cómo se prueba cada módulo** (PC, Postgres local en Docker: contenedor `bradenton-pg`, puerto 55432):

```
docker start bradenton-pg
set DATABASE_URL=   (la URL local está en .env.docker-local, que no se sube)
python -m alembic upgrade head
python comparar_bases.py jh_facturas
```

- Con `127.0.0.1`, no `localhost`: en esta PC `localhost` busca primero la dirección IPv6 y se cuelga.
- El comparador copia el módulo a Postgres (`migrar_sqlite_a_postgres.py --vaciar`, que solo acepta un Postgres de esta PC), abre sus páginas sobre SQLite y sobre Postgres, hace sus pruebas de escritura y compara todo.
- Para el módulo siguiente: agregar sus páginas en `PAGES` y su escenario de escritura en `SCENARIOS` (`comparar_bases.py`).

**Para tener en cuenta en los módulos que siguen:**
- Postgres pasa a minúsculas los alias de columna (`AS Total` llega como `total`). La lectura `row["Total"]` anda igual, pero `dict(row)` da la clave en minúscula: revisar los alias con mayúsculas.
- `lastrowid` no existe en Postgres: hay que usar `INSERT … RETURNING id`, que SQLite 3.49 también entiende.

## 1. Qué hay hoy

- **19 bases SQLite**, no 18. La número 19, `documents.db`, es del guardado de documentos, que está apagado desde el 28/09.
- 52 tablas, 438 columnas y 35.983 filas: unos 6 MB.
- Todas las tablas tienen clave primaria y ningún nombre de tabla se repite entre bases.
- Ningún valor guardado es de otro tipo que el declarado: se revisó columna por columna. No hay sorpresas de tipos para Postgres.
- **20 módulos `*_db.py`**: unas 8.800 líneas y unas 430 consultas. Fuera de ellos no hay SQL; `webapp.py` solo captura `sqlite3.Error` en un lugar.
- **Configuración en archivos:**
  - `users.json` (1 usuario) y `.flask_secret_key`.
  - Las reglas: `chase_master_rules.json` (53 reglas), `chase_rules.json` (59), `proveedores_pago_rules.json` (30) y `proveedores_dynamic_suppliers.json` (vacío).
  - Dos más que se crean al usarse y hoy no existen: `controles_rapidos.json` (límite de Caja) y `control_cierre_sale_tax.json` (marcas de Sale Tax).
- **Unos 360 MB de PDF y fotos** guardados (pdfs, lottery_pdfs, cheques, depositos, documents, caja_attachments). El guardado está apagado; no se migran (ver punto 7).

## 2. Cómo queda armado

- **La app (Render, Docker) y la base (Postgres en Render o Neon) quedan separadas.** La app no guarda en su disco nada que tenga que durar, así que un deploy no borra nada.
- **Una capa nueva, `db.py`, como única puerta a la base:** `db.connect("chase")` en lugar del `_connect()` de cada módulo.
  - Con `DATABASE_URL` va a Postgres (psycopg 3 con un pool de conexiones); sin la variable, al SQLite de hoy.
  - Durante la migración conviven los dos: la PC sigue trabajando con SQLite y cada módulo pasado se prueba contra Postgres. Después de la copia a producción se saca SQLite.
  - La capa resuelve lo que se escribe distinto en cada motor:
    - los `?` pasan a `%s`;
    - las filas se leen por nombre y por posición, como hoy con `sqlite3.Row`;
    - los errores quedan como `db.IntegrityError` y `db.Error`;
    - `with conn:` imita a sqlite3, porque en psycopg ese `with` cierra la conexión y en sqlite3 solo confirma.
- **Un esquema de Postgres por cada base de hoy** (`chase`, `eft`, `proveedores`, ...) dentro de una misma base de datos. Se mantiene la misma división que hoy, y se puede copiar, verificar y restaurar por módulo.
- **Tipos:**
  - `REAL` → `DOUBLE PRECISION`: mismos números que hoy. `NUMERIC` devolvería `Decimal` y rompería las cuentas en Python.
  - `TEXT` → `TEXT COLLATE "C"`: ordena igual que SQLite. El orden por defecto de Postgres pone distinto las mayúsculas, los espacios y los signos.
  - `BLOB` → `BYTEA`.
  - **Las fechas siguen como texto `AAAA-MM-DD`** en esta etapa, así `WHERE date LIKE '2026-09%'` da exactamente lo mismo. Pasarlas a `DATE` es una mejora para después, no parte de la migración.
- **Migraciones con Alembic**, la herramienta estándar de Python (sección 9 de la guía):
  - una migración por módulo, con sus tablas en su forma de hoy;
  - se aplica igual en local, testing y producción con `alembic upgrade head`, que hace lo mismo que `npx prisma migrate deploy` en la guía: se corre una vez contra testing y, ya probado, otra vez contra producción, cambiando `DATABASE_URL`;
  - en Render corre sola antes de arrancar la app;
  - mientras conviva, el lado SQLite sigue creando sus tablas como hoy.
- **Dependencias nuevas, con versión fija** (eso cierra A11 de la auditoría): `psycopg[binary]`, `psycopg-pool`, `alembic` (con `SQLAlchemy`, que Alembic necesita; la app no usa el ORM) y `python-dotenv` para leer el `.env`.

## 3. Lo propio de SQLite que hay que reescribir

Las cantidades cuentan solo lo que está dentro de las consultas SQL; `lastrowid` se cuenta en Python.

| Qué | Cuántas | Cómo queda | Riesgo |
|---|---|---|---|
| Fechas con `strftime` | 0 en SQL | Las fechas se arman en Python (no cambia) y en SQL se comparan como texto con `LIKE`/`substr`, que funcionan igual | Ninguno |
| `rowid` | 13, todas en `chase_db` | La tabla de Chase no tiene id propio (su clave es fecha + descripción + monto + ocurrencia). En Postgres lleva una columna `id` autonumérica; en SQLite sigue `rowid` (la capa da el nombre) | Medio: se prueba subiendo los extractos reales |
| `PRAGMA` y `ALTER TABLE` | 31 y 11 | Crean o completan las tablas al conectar. En Postgres no corren: el esquema lo arma Alembic con la forma final | Bajo |
| `AUTOINCREMENT` | 22 | `GENERATED BY DEFAULT AS IDENTITY` (así la copia conserva los ids) y al final se ajusta el contador | Bajo: el script de copia controla el contador |
| `lastrowid` | 14 | `INSERT … RETURNING id`, que el SQLite de la PC (3.49) también entiende | Bajo |
| `INSERT OR REPLACE` | 9 | `INSERT … ON CONFLICT (clave) DO UPDATE SET …`, que ya se usa en 20 lugares y anda en los dos | Medio: `OR REPLACE` borra y vuelve a crear la fila, y lo que no está en el INSERT vuelve a su valor por defecto. Se revisa uno por uno |
| `INSERT OR IGNORE` | 4 | `ON CONFLICT DO NOTHING` | Bajo |
| `COLLATE NOCASE` | 4 (cmv, fisico, gettel) | `LOWER(col) = LOWER(?)` | Bajo |
| `LIKE` del buscador por nombre de archivo | 3 (documents, lottery, reportes) | En SQLite no distingue mayúsculas y en Postgres sí: `LOWER(col) LIKE LOWER(?)`. Los otros 24 `LIKE` son sobre fechas y no cambian | Bajo |
| `printf('%04d-%02d', …)` | 1 (gettel) | `MIN(year * 100 + month)` y el formato en Python | Bajo |
| `GROUP_CONCAT` | 1 (reportes) | `string_agg`, que SQLite 3.49 también tiene | Bajo |
| `GROUP BY` con una columna suelta | 1 (proveedores) | `supplier_label` no está en el `GROUP BY`: SQLite lo acepta y Postgres da error. Pasa a `MAX(supplier_label)` | Bajo |
| `executescript` | 1 (kia_toyota) | Crea tablas: lo pasa a hacer Alembic | Bajo |
| Vacíos al ordenar | sin contar | SQLite pone los `NULL` primero y Postgres al final. Se agrega `NULLS FIRST` donde haga falta; lo detecta el comparador (punto 6) | Medio: cambia el orden de una tabla, no los números |
| Una conexión por consulta | todos | Contra Postgres por red, abrir una conexión cuesta 50-200 ms. Lo resuelve el pool de `db.py` | Lentitud si no se hace |

## 4. Módulo por módulo, de la base más chica a la más grande

**Paso 0, antes del primer módulo.** No cambia nada de lo que ya funciona, porque todavía nadie lo usa:
- `db.py`;
- Alembic con su configuración y la carpeta de migraciones;
- leer el `.env`;
- un Postgres local en Docker para probar rápido;
- el comparador del punto 6.

Archivos: nuevos `db.py`, `alembic.ini` y `migraciones/`; se modifica `requirements.txt`.

**Cada módulo sigue siempre los mismos pasos:**
1. su migración de Alembic;
2. en su `*_db.py`, `_connect()` pasa a `db.connect("<esquema>")` y se reescribe lo de la tabla de arriba que le toque;
3. se copian sus datos reales a Postgres;
4. se comparan contra SQLite sus páginas, PDF y Excel: **tienen que dar idénticos**;
5. commit chico al repo dev, y recién ahí el siguiente.

Cada módulo usa solo su propio esquema. Por eso se puede probar con unos en Postgres y otros en SQLite aunque se lean entre sí; por ejemplo, Caja lee Chase, Lottery y Store Info.

| # | Base (tamaño · tablas · filas) | Archivo | Qué cambia | Riesgo | Qué se compara |
|---|---|---|---|---|---|
| 1 | jh_facturas (12 KB · 1 · 10) | `jh_facturas_db.py` | Solo la conexión y su `PRAGMA` | Bajo | Facturas varias, Próximos EFT |
| 2 | ice_machine (16 KB · 1 · 8) | `ice_machine_db.py` | 1 `OR REPLACE` | Bajo | Depósitos (Payment Summary), Control Depósitos |
| 3 | reporte_mensual (20 KB · 2 · 61) | `reporte_mensual_db.py` | 2 `OR REPLACE`, 2 `ALTER` | Bajo-medio: la edición a mano | Reporte C-Store mensual y sus cruces, Control CMV |
| 4 | cheques (24 KB · 1 · 31) | `cheques_db.py` | `AUTOINCREMENT`, 2 `lastrowid`, `ALTER` | Bajo | Chase → Cheques y su cuadro |
| 5 | control_cmv (28 KB · 1 · 4) | `control_cmv_db.py` | 1 `OR REPLACE` | Bajo | Control CMV |
| 6 | depositos (28 KB · 1 · 84) | `depositos_db.py` | `AUTOINCREMENT`, `lastrowid` | Bajo | Depósitos, Control Depósitos, Caja |
| 7 | horas_trabajo (28 KB · 2 · 42) | `horas_trabajo_db.py` | 2 `AUTOINCREMENT`, 2 `lastrowid` | Bajo | Horas y su reporte |
| 8 | jh_mensual (28 KB · 1 · 9) | `jh_mensual_db.py` | 1 `OR REPLACE` | Bajo | Tarjetas y Cupones, EFT |
| 9 | documents (32 KB · 1 · 43) | `documents_db.py` | Guardado apagado: migrar la tabla o dejarla afuera (decisión) | Bajo | — |
| 10 | fisico (32 KB · 3 · 33) | `fisico_db.py` | `NOCASE`, `ALTER`, 2 `AUTOINCREMENT` | Bajo | Físico, facturas de combustible, Próximos EFT |
| 11 | caja (44 KB · 5 · 18) | `caja_db.py` | 2 `AUTOINCREMENT`, 2 `lastrowid`, `ALTER` | Medio: encadena saldos de mes a mes | Caja de cada mes del año, botonera, PDF |
| 12 | gettel_toyota (48 KB · 3 · 131) | `gettel_db.py` | `NOCASE`, `printf`, 4 `PRAGMA`, `ALTER` | Bajo-medio | Gettel/Toyota, pagos, reportes |
| 13 | lottery (56 KB · 4 · 88) | `lottery_db.py` (1.651 líneas) | 1 `OR REPLACE`, 4 `ON CONFLICT`, buscador | Medio: bloques semanales y fechas de pago encadenadas | Lottery por día y bloques, Control Lottery, PDF |
| 14 | reportes_diarios (216 KB · 3 · 1.924) | `reportes_db.py` | `GROUP_CONCAT`, 6 `ON CONFLICT`, buscador, `ALTER` | Medio-alto: casi todo sale de acá | Reporte Diario, Store Info, Departamentos, Cierre, Tarjetas, Caja; además, una carga real de un Close Store |
| 15 | cuenta_kia_toyota (260 KB · 4 · 2.173) | `cuenta_kia_toyota_db.py` | 3 `OR REPLACE`, `executescript` | Medio: pagos asignados a mano | Kia y Toyota (saldos por mes), botonera |
| 16 | eft (388 KB · 6 · 3.432) | `eft_db.py` (1.849 líneas, 63 consultas) | `OR IGNORE`, 4 `AUTOINCREMENT`, 3 `lastrowid`, claves foráneas | Medio-alto | EFT y Cupones, Cupones por mes, Próximos EFT, Tarjetas, PDF de EFT |
| 17 | chase (460 KB · 1 · 636) | `chase_db.py` | los 13 `rowid`, 4 `ALTER` | Alto: identidad de cada movimiento, ocurrencias, recategorizar | Chase de cada mes y sus categorías, Caja, Control Depósitos; además, volver a subir un extracto real |
| 18 | cmv (1,4 MB · 5 · 14.539) | `cmv_db.py` | 2 `NOCASE`, 2 `AUTOINCREMENT` | Medio | CMV costo y ventas, fotos y comparación, Ventas y costos, Precios a revisar |
| 19 | proveedores (2,6 MB · 7 · 12.717) | `proveedores_db.py` | 3 `OR IGNORE`, 5 `AUTOINCREMENT`, `GROUP BY` suelto, `IntegrityError` | Medio-alto: cuenta corriente con pagos de Chase e históricos | Planilla y cuenta corriente de cada proveedor, Productos, reporte de precios; además, una carga real de facturas |

**Al final:**
- `webapp.py`: el `sqlite3.Error` pasa a `db.Error`.
- La configuración (punto 5).
- Sacar el lado SQLite.

## 5. Usuarios, clave de sesión y reglas editables: propuesta para revisar antes de implementar

Todo va a un esquema `config`. Las páginas no cambian; cambia por dentro cómo se lee y se guarda.

| Hoy | Tabla propuesta | Columnas |
|---|---|---|
| `users.json` | `config.usuarios` | `username` (clave), `password_hash`, `is_admin`, `must_change_password`, `creado`, `actualizado` |
| `.flask_secret_key` | `config.secretos` | `nombre` (clave), `valor`. La clave de sesión se genera la primera vez y queda guardada: no se desloguea a todos en cada deploy y no hace falta una tercera variable de entorno |
| `chase_master_rules.json`, `chase_rules.json` | `config.reglas_chase` | `id`, `tipo` (maestra / personalizada), `keyword`, `detail`, `orden` |
| `proveedores_pago_rules.json` | `config.reglas_pago_proveedores` | `id`, `keyword`, `sheet_name`, `orden` |
| `proveedores_dynamic_suppliers.json` | `config.proveedores_dinamicos` | `key` (clave), `definicion` (JSONB) |
| `controles_rapidos.json` | `config.ajustes` | `nombre` (clave), `valor` (JSONB), `actualizado`. Por ahora, el límite de la alerta de Caja |
| `control_cierre_sale_tax.json` | `config.sale_tax_registrado` | `year`, `month` (clave), `asiento`, `registrado_en` |

- **El primer admin con la clave de instalación se mantiene igual:** sin usuarios, `/primer-admin` exige `BRADENTON_SETUP_KEY`.
- Si la base no responde, la app da error y no deja entrar: nunca toma "no pude leer" como "no hay usuarios". Eso cierra A1 de la auditoría.
- Los JSON del repo siguen siendo la semilla: la migración de Alembic los carga solo si la tabla está vacía.

## 6. Cómo seguir validando contra los datos reales de Cande

- **Comparador (`comparar_bases.py`):** la misma app con los mismos datos, una vez sobre SQLite y otra sobre Postgres.
  - Recorre todas las páginas de jul-oct 2026: Carga de Datos, Controles, Reportes y la botonera.
  - Hace todas las descargas.
  - Compara el HTML (sin la hora ni los tokens), el texto de los PDF y las celdas de los Excel.
  - **Tienen que dar idénticos.** Cualquier diferencia se explica o se arregla antes del commit.
- **Por cada módulo, una carga real del Drive en los dos motores** (un Close Store, un extracto de Chase, facturas de un proveedor), y se compara lo que quedó guardado.
- **Los números que ya se validaron quedan como prueba fija:**
  - agosto: Store Info = Excel Cierre al centavo;
  - Lottery, semana 27/07-02/08: $4.637,19 = Chase 05/08;
  - los saldos de Kia/Toyota al 30/09;
  - FPL, Manatee y Airgas al 15/10;
  - los EFT de octubre.

## 7. Pasar los datos reales

Script `migrar_sqlite_a_postgres.py`:
1. Lee una **copia** de cada base SQLite, en solo lectura; nunca la original.
2. Exige que la base de destino tenga el esquema al día (`alembic upgrade head`) y esté vacía; si no, se niega.
3. Copia tabla por tabla, con una transacción por base (si algo falla, no queda a medias) y respetando las claves foráneas (en EFT, primero los depósitos y después sus renglones).
4. Ajusta los contadores de ids.
5. Verifica que sean idénticos:
   - la cantidad de filas de cada tabla;
   - la suma de cada columna numérica;
   - un hash de cada fila, ordenada por su clave.

   Además compara los totales de negocio de cada mes contra lo validado: Chase por categoría, neto de los EFT, facturas por proveedor, Store Info y Caja.
6. Carga los usuarios y las reglas desde los archivos.

**Orden:**
1. Postgres local.
2. Testing.
3. Cande valida en testing.
4. Producción.

Mientras se trabaja en testing, la PC sigue cargando en SQLite. La copia a producción se hace con los datos de ese día; desde ahí producción manda y la PC deja de cargar en SQLite.

**Lo que no se migra:**
- Los PDF y fotos guardados (unos 360 MB). En Render, los "Ver PDF" viejos de Lottery y Cheques van a decir "archivo no disponible".
- Los archivos temporales y las cargas en segundo plano, que no necesitan durar.

## 8. Respaldo de Postgres

Script `respaldo_postgres.py`, como pide la sección 9 de la guía:
- `pg_dump` contra la `DATABASE_URL` del ambiente que se le indique (testing o producción);
- el archivo lleva fecha y hora en el nombre (`bradenton-dev_AAAAMMDD_HHMM.dump`) y va a una carpeta `backups/` que está en `.gitignore`;
- avisa con un mensaje claro si salió bien o si falló;
- guarda los últimos N respaldos;
- `pg_restore` para volver a uno.

Hay que instalar `pg_dump` en la PC, de la misma versión que la base de Render o más nueva. Dos formas: las "Command Line Tools" del instalador oficial de PostgreSQL para Windows, o la imagen de Docker de Postgres (Docker ya está instalado). Antes de escribirlo, se muestra cómo va a funcionar (lo pide la guía). Reemplaza las copias de `.db` que se hacen hoy antes de corregir datos reales. Se suma a los respaldos propios de Render o Neon.

## 9. Linux limpio y deploy

- **Prueba en Linux limpio:** imagen construida desde cero con Docker y la base vacía:
  1. `alembic upgrade head`;
  2. `/primer-admin` con la clave;
  3. cargar un día de cada módulo;
  4. aviso para que lo pruebe la jefa.
- **`Dockerfile`:**
  - arranque con `alembic upgrade head && python webapp.py`;
  - `TZ=America/New_York` (J6 de la auditoría: hoy "hoy" es UTC);
  - `PYTHONUNBUFFERED=1`;
  - `Claude outputs/` en `.dockerignore` (A10).
- **Variables de entorno:** `DATABASE_URL` y `BRADENTON_SETUP_KEY`. `BRADENTON_OCR_PROCESOS` es opcional. `BRADENTON_DATA_DIR` y el disco persistente quedan obsoletos.
- **Antes del deploy, la lista de la sección 11 de la guía:** sin claves en el código (se revisa con grep), build limpio y las variables listadas.

## 10. Lo que hace falta

- **Render, según la guía de Leti:** todo en Oregon y con el plan que indique Jorge.

  | Ambiente | Servicio web | Base | Quién |
  |---|---|---|---|
  | Testing | `bradenton-dev` (repo BradentonApp-dev, rama main, Docker) | `bradenton-dev-db` | Leti, cuando el código esté listo y probado |
  | Producción | `bradenton` | `bradenton-db` | Leti |

- **Las dos URL de la base de testing:**
  - la **interna** va en el servicio `bradenton-dev` → Environment, como `DATABASE_URL`;
  - la **externa** va en el `.env` de la PC, para el script de copia, el respaldo y las pruebas desde la PC.

  Ninguna va en GitHub, el chat ni WhatsApp.
- **Abrir Docker Desktop** en la PC, para el Postgres local.
- **Cambio del 16/10 (Leti, ya hablado con Jorge):** el servicio, las bases y las variables de Render los crea Leti. Claude termina todo el código y lo prueba con el Postgres local en Docker. Cuando esté listo, avisa, y recién ahí Leti crea `bradenton-dev` y `bradenton-dev-db` y carga las variables. Alfonso no necesita acceso a Render.
- **Para la copia de los datos reales (punto 7):** hace falta la URL externa de `bradenton-dev-db`. O la pone Leti directamente en el `.env` de la PC de Alfonso, o corre ella el script. La URL no se manda por WhatsApp.
- **Decisiones:**
  - las tablas del punto 5;
  - `documents.db`: migrarla o dejarla afuera;
  - después del pase, la PC apunta a testing (para probar) y la carga diaria va a producción.

**Tiempo estimado:**

| Etapa | Chats |
|---|---|
| Paso 0 y módulos 1-8 | 1 |
| Módulos 9-13 | 1 |
| Módulos 14-17 | 1-2 |
| Módulos 18-19, configuración, copia y Linux | 1-2 |
| **Total** | **4-6** |
