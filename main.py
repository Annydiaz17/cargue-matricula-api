import io
import os
import hashlib
from typing import Optional
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import oracledb
import datetime

# Activar Oracle Client en modo Thick
try:
    oracledb.init_oracle_client(lib_dir="/opt/oracle/instantclient")
    print("Oracle Client inicializado exitosamente.")
except Exception as e:
    print(f"Error al inicializar Oracle Client: {e}")

app = FastAPI(title="API REST - Carga de Excel/CSV a Oracle APEX")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_SCHEMA = os.getenv("DB_SCHEMA")
DB_DSN = os.getenv("DB_DSN")

TIPOS_ARCHIVO_VALIDOS = ['ANEXO_6A', 'DETALLADO', 'PLANTAGEN1', 'ESTABLECIMIENTOS', 'SEDES']


def leer_archivo(file_bytes: bytes, filename: str):
    """Devuelve (df, filas_omitidas). Todo se lee como texto: si pandas infiere
    numeros, los codigos DANE se convierten en decimales y dejan de cruzar."""
    filas_omitidas = 0

    if filename.lower().endswith('.csv'):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8', sep=';',
                             dtype=str, on_bad_lines='skip')
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding='latin1', sep=';',
                             dtype=str, on_bad_lines='skip')

        lineas = file_bytes.count(b'\n')
        if file_bytes and not file_bytes.endswith(b'\n'):
            lineas += 1
        filas_omitidas = max(lineas - 1 - len(df), 0)
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl', dtype=str)

    return df.fillna(''), filas_omitidas


def procesar_y_llamar_paquete(
    file_bytes: bytes, 
    filename: str, 
    anio_cargue: Optional[int], 
    calendario: Optional[str], 
    usuario_apex: Optional[str], 
    tipo_archivo: str,
    exigir_municipio: str = 'S'
):
    connection = None
    cursor = None
    id_cargue = None
    try:
        huella_archivo = hashlib.sha256(file_bytes).hexdigest()

        df, filas_omitidas = leer_archivo(file_bytes, filename)
        total_filas = len(df)

        if filas_omitidas:
            print(f"[{filename}] {datetime.datetime.now()} "
                  f"ATENCION: {filas_omitidas} linea(s) del archivo no se pudieron leer.")

        connection = oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=DB_DSN
        )
        cursor = connection.cursor()

        db_usuario = usuario_apex if usuario_apex and usuario_apex.strip() != "" else None

        # Validación si el archivo está vacío
        if total_filas == 0:
            mensaje_error = f"El archivo '{filename}' está vacío."
            print(f"[{filename}] {datetime.datetime.now()} {mensaje_error}")
            
            sql_cargue_fail = """
                INSERT INTO smc_core.cargue_archivo (
                    tipo_archivo, nombre_archivo, huella_archivo, usuario, ruta_archivo,
                    registros_leidos, registros_aceptados, registros_rechazados,
                    estado, mensaje
                ) VALUES (:1, :2, :3, :4, :5, 0, 0, 0, 'FALLIDO', :6)
            """
            cursor.execute(sql_cargue_fail, [tipo_archivo, filename, huella_archivo, db_usuario, f"api_upload/{filename}", mensaje_error])
            connection.commit()
            return

        # RUTA SEGÚN TIPO DE ARCHIVO
        if tipo_archivo == 'ANEXO_6A':
            mensaje_inicial = 'Iniciando procesamiento ANEXO_6A'
            if filas_omitidas:
                mensaje_inicial += (f' | ATENCION: {filas_omitidas} linea(s) del archivo '
                                    f'no se pudieron leer y no entraron al cargue')

            sql_cargue_archivo = """
                INSERT INTO smc_core.cargue_archivo (
                    tipo_archivo, nombre_archivo, huella_archivo, usuario, ruta_archivo,
                    registros_leidos, registros_aceptados, registros_rechazados,
                    registros_insertados, registros_actualizados, registros_sin_cambio,
                    estado, mensaje
                ) VALUES (
                    :1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13
                ) RETURNING id_cargue INTO :out_id
            """
            out_id = cursor.var(oracledb.NUMBER)
            cursor.execute(sql_cargue_archivo, [
                tipo_archivo, filename, huella_archivo, db_usuario, f"api_upload/{filename}",
                total_filas, 0, 0, 0, 0, 0, 'EN_PROCESO', mensaje_inicial, out_id
            ])
            id_cargue = out_id.getvalue()[0]
            connection.commit()

            registros_staging = []
            for idx, row_df in df.iterrows():
                num_linea = idx + 1
                contenido_original = f"Fila {num_linea} Archivo: {filename}"
                registros_staging.append((
                    id_cargue, num_linea,
                    str(row_df.get('CONTENIDO_ORIGINAL', contenido_original))[:4000],
                    str(row_df.get('PER_ID', '')), str(row_df.get('COD_TIPO_DOC', '')),
                    str(row_df.get('NRO_IDENTIDAD', '')), str(row_df.get('PRI_NOM', '')),
                    str(row_df.get('SEG_NOM', '')), str(row_df.get('PRI_APE', '')),
                    str(row_df.get('SEG_APE', '')), str(row_df.get('GENERO', '')),
                    str(row_df.get('FECHA_NACIMIENTO', '')), str(row_df.get('COD_ETNIA', '')),
                    str(row_df.get('COD_TIPO_DISCA', '')), str(row_df.get('COD_CAP_EXCEP', '')),
                    str(row_df.get('TRASTORNOS_ESPECIFICOS', '')), str(row_df.get('APOYO_ACADEMICO_ESPECIAL', '')),
                    str(row_df.get('COD_DANE', '')), str(row_df.get('CONS_SEDE', '')),
                    str(row_df.get('DANE_SEDE', '')), str(row_df.get('COD_ZONA', '')),
                    str(row_df.get('NOMBRE_SEDES', '')), str(row_df.get('AÑO', '')),
                    str(row_df.get('COD_GRADO', '')), str(row_df.get('COD_GRUPO', '')),
                    str(row_df.get('COD_JOR', '')), str(row_df.get('COD_MET', '')),
                    str(row_df.get('REPITENTE', '')), str(row_df.get('COD_FTE_REC', '')), 'PENDIENTE'
                ))

            sql_insert = """
                INSERT INTO smc_core.stg_anexo6a (
                    id_cargue, num_linea, contenido_original, per_id, tipo_documento, nro_documento,
                    primer_nombre, segundo_nombre, primer_apellido, segundo_apellido, genero,
                    fecha_nacimiento_txt, etnia, tipo_discapacidad, cap_excepcional, trastorno_especifico,
                    apoyo_academico_especial, codigo_dane, cons_sede, dane_sede_txt, zona_txt,
                    nombre_sede_txt, ano_inf, grado, codigo_grupo_txt, tipo_jornada,
                    metodologia, repitente, fuente_recursos_txt, estado_validacion
                ) VALUES (
                    :1, :2, :3, :4, :5, :6, :7, :8, :9, :10,
                    :11, :12, :13, :14, :15, :16, :17, :18, :19, :20,
                    :21, :22, :23, :24, :25, :26, :27, :28, :29, :30
                )
            """
            for i in range(0, total_filas, 5000):
                cursor.executemany(sql_insert, registros_staging[i:i + 5000])
                connection.commit()

            cursor.callproc(
                "smc_core.pkg_cargue_matricula_optimizado.procesar_cargue", 
                [id_cargue, anio_cargue, calendario, db_usuario]
            )
            print(f"[{filename}] {datetime.datetime.now()} Procesamiento ANEXO_6A finalizado con éxito.")

        elif tipo_archivo == 'PLANTAGEN1':
            print(f"[{filename}] {datetime.datetime.now()} Limpiando tabla stg_plantagen1...")
            cursor.execute("DELETE FROM smc_core.stg_plantagen1")
            connection.commit()

            registros_staging = []
            for idx, row_df in df.iterrows():
                registros_staging.append((
                    str(row_df.get('NUMVINCULACION', '')),
                    str(row_df.get('CODEMPLEADO', '')),
                    str(row_df.get('EMPLEADO', '')),
                    str(row_df.get('TIPOCARGO', '')),
                    str(row_df.get('RETIROFECHA', '')),
                    str(row_df.get('CODCARGOEMPRESA', '')),
                    str(row_df.get('CARGOEMPRESA', '')),
                    str(row_df.get('REGIMEN', '')),
                    str(row_df.get('SITUACIONLABORAL', '')),
                    str(row_df.get('VINCULACIONNOVEDAD', '')),
                    str(row_df.get('CODAREAEDUCATIVA', '')),
                    str(row_df.get('AREAEDUCATIVA', '')),
                    str(row_df.get('CODNIVELDICTA', '')),
                    str(row_df.get('NIVELDICTA', '')),
                    str(row_df.get('CODESTABLECIMIENTO', '')),
                    str(row_df.get('DEPENDENCIA', '')),
                    str(row_df.get('DANE_SEDE', ''))
                ))

            sql_insert_planta = """
                INSERT INTO smc_core.stg_plantagen1 (
                    numvinculacion, codempleado, empleado, tipocargo, retirofecha,
                    codcargoempresa, cargoempresa, regimen, situacionlaboral, vinculacionnovedad,
                    codareaeducativa, areaeducativa, codniveldicta, niveldicta,
                    codestablecimiento, dependencia, dane_sede
                ) VALUES (
                    :1, :2, :3, :4, :5, :6, :7, :8, :9, :10,
                    :11, :12, :13, :14, :15, :16, :17
                )
            """
            for i in range(0, total_filas, 5000):
                cursor.executemany(sql_insert_planta, registros_staging[i:i + 5000])
                connection.commit()

            print(f"[{filename}] {datetime.datetime.now()} Inserción en staging PLANTAGEN1 completada. Ejecutando PKG_CARGUE_PLANTA...")

            out_id_cargue = cursor.var(oracledb.NUMBER)
            cursor.callproc(
                "smc_core.pkg_cargue_planta.cargar",
                [filename, huella_archivo, exigir_municipio, out_id_cargue, db_usuario]
            )
            print(f"[{filename}] {datetime.datetime.now()} ¡Procesamiento de PLANTAGEN1 finalizado con éxito! ID Cargue: {out_id_cargue.getvalue()}")

        else:
            print(f"[{filename}] El tipo de archivo '{tipo_archivo}' no tiene un flujo definido.")

    except Exception as e:
        if connection:
            connection.rollback()
        print(f"[{filename}] {datetime.datetime.now()} Error procesando el cargue: {str(e)}")
        if connection and id_cargue:
            try:
                cursor_err = connection.cursor()
                cursor_err.execute("""
                    UPDATE smc_core.cargue_archivo 
                       SET estado = 'FALLIDO', mensaje = :1 
                     WHERE id_cargue = :2
                """, [str(e)[:4000], id_cargue])
                connection.commit()
            except Exception:
                pass
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()

@app.post("/api/upload-excel")
async def upload_excel(
    background_tasks: BackgroundTasks, 
    file: UploadFile = File(...),
    tipo_archivo: str = Form(...),
    anio_cargue: Optional[int] = Form(default=None),
    calendario: Optional[str] = Form(default="A"),
    usuario_apex: Optional[str] = Form(default=None),
    exigir_municipio: Optional[str] = Form(default="S")
):
    """Endpoint unificado para recibir archivos (.xlsx, .xls, .csv) y procesarlos en segundo plano."""
    
    extensiones_validas = ('.xlsx', '.xls', '.csv')
    if not file.filename.lower().endswith(extensiones_validas):
        raise HTTPException(
            status_code=400, 
            detail="Formato inválido. Sube un archivo con extensión .xlsx, .xls o .csv"
        )

    if tipo_archivo not in TIPOS_ARCHIVO_VALIDOS:
        raise HTTPException(
            status_code=400, 
            detail=f"tipo_archivo inválido. Debe ser uno de: {', '.join(TIPOS_ARCHIVO_VALIDOS)}"
        )

    content = await file.read()

    background_tasks.add_task(
        procesar_y_llamar_paquete, 
        content, 
        file.filename, 
        anio_cargue, 
        calendario, 
        usuario_apex, 
        tipo_archivo,
        exigir_municipio
    )

    return {
        "status": "success",
        "message": f"Archivo '{file.filename}' de tipo '{tipo_archivo}' recibido. Procesando en segundo plano.",
        "file_size_bytes": len(content)
    }