import io
import os
import hashlib
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import oracledb

import datetime


# Activar Oracle Client
try:
    oracledb.init_oracle_client(lib_dir="/opt/oracle/instantclient")
    print("Oracle Client inicializado exitosamente.")
except Exception as e:
    print(f"Error al inicializar Oracle Client en Modo Thick: {e}")

app = FastAPI(title="API REST - Carga de Excel a Oracle APEX")

# Configuración CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Credenciales de conexión Oracle
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_SCHEMA = os.getenv("DB_SCHEMA")
DB_DSN = os.getenv("DB_DSN")

def procesar_y_llamar_paquete(file_bytes: bytes, filename: str, anio_cargue: int, calendario: str, usuario_apex: str):
    """Registra el cargue en CARGUE_ARCHIVO usando el usuario real de APEX, mapea los nuevos nombres del Excel, inserta en lotes y ejecuta el PL/SQL."""
    connection = None
    try:
        # Calcular el Hash SHA-256
        huella_archivo = hashlib.sha256(file_bytes).hexdigest()

        # Lectura del Excel en memoria 
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        df = df.fillna('')  # Reemplazar nulos por vacíos
        total_filas = len(df)

        if total_filas == 0:
            print(f"[{filename}] El archivo está vacío.")
            return

        usuario_conexion = f"{DB_USER}[{DB_SCHEMA}]"

        # Conexión a la Base de Datos Oracle
        connection = oracledb.connect(
            user=usuario_conexion,
            password=DB_PASSWORD,
            dsn=DB_DSN
        )
        cursor = connection.cursor()

        # Usar el usuario que viene desde APEX
        db_usuario = usuario_apex if usuario_apex else "APEX_USER_UNKNOWN"

        # Registrar el encabezado CARGUE_ARCHIVO con el usuario real de APEX
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
            'ANEXO_6A',
            filename,
            huella_archivo,
            db_usuario,
            f"api_upload/{filename}",
            total_filas,
            0, 0, 0, 0, 0,
            'EN_PROCESO',
            'Iniciando procesamiento desde API REST',
            out_id
        ])
        
        id_cargue = out_id.getvalue()[0]
        connection.commit()
        print(f"[{filename}] {datetime.datetime.now()} Cabecera registrada en CARGUE_ARCHIVO para el usuario APEX: {db_usuario} (ID: {id_cargue})")

        # Preparar las filas de la tabla temporal asociadas al id_cargue con los nuevos nombres del Excel
        registros_staging = []
        for idx, row_df in df.iterrows():
            num_linea = idx + 1
            contenido_original = f"Fila {num_linea} - Archivo: {filename}"

            registros_staging.append((
                id_cargue,
                num_linea,
                str(row_df.get('CONTENIDO_ORIGINAL', contenido_original))[:4000],
                str(row_df.get('PER_ID', '')),
                str(row_df.get('COD_TIPO_DOC', '')),
                str(row_df.get('NRO_IDENTIDAD', '')),
                str(row_df.get('PRI_NOM', '')),
                str(row_df.get('SEG_NOM', '')),
                str(row_df.get('PRI_APE', '')),
                str(row_df.get('SEG_APE', '')),
                str(row_df.get('GENERO', '')),
                str(row_df.get('FECHA_NACIMIENTO', '')),
                str(row_df.get('COD_ETNIA', '')),
                str(row_df.get('COD_TIPO_DISCA', '')),
                str(row_df.get('COD_CAP_EXCEP', '')),
                str(row_df.get('TRASTORNOS_ESPECIFICOS', '')),
                str(row_df.get('APOYO_ACADEMICO_ESPECIAL', '')),
                str(row_df.get('COD_DANE', '')),
                str(row_df.get('CONS_SEDE', '')),
                str(row_df.get('DANE_SEDE', '')),
                str(row_df.get('COD_ZONA', '')),
                str(row_df.get('NOMBRE_SEDES', '')),
                str(row_df.get('AÑO', '')),
                str(row_df.get('COD_GRADO', '')),
                str(row_df.get('COD_GRUPO', '')),
                str(row_df.get('COD_JOR', '')),
                str(row_df.get('COD_MET', '')),
                str(row_df.get('REPITENTE', '')),
                str(row_df.get('COD_FTE_REC', '')),
                'PENDIENTE'
            ))

        # Inserción en lotes para stg_anexo6a
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

        tamano_lote = 5000
        print(f"[{filename}] {datetime.datetime.now()} Iniciando inserción de {total_filas} filas en stg_anexo6a...")

        for i in range(0, total_filas, tamano_lote):
            lote = registros_staging[i:i + tamano_lote]
            cursor.executemany(sql_insert, lote)
            connection.commit()
            print(f"[{filename}] Progreso: {min(i + tamano_lote, total_filas)} / {total_filas} filas insertadas...")

        print(f"[{filename}] {datetime.datetime.now()} Inserción temporal completada. Llamando al paquete PL/SQL...")

        # Llamar al procedimiento
        cursor.callproc("smc_core.pkg_cargue_matricula_optimizado.procesar_cargue", [id_cargue, anio_cargue, calendario])
        
        print(f"[{filename}] {datetime.datetime.now()} ¡Procesamiento del paquete PL/SQL finalizado con éxito!")

    except Exception as e:
        if connection:
            connection.rollback()
        print(f"[{filename}] {datetime.datetime.now()} Error procesando el cargue: {str(e)}")
    finally:
        if connection:
            connection.close()

@app.post("/api/upload-excel")
async def upload_excel(
    background_tasks: BackgroundTasks, 
    file: UploadFile = File(...),
    anio_cargue: int = Form(...),
    calendario: str = Form(default="A"),
    usuario_apex: str = Form(default="SMC_API")
):
    """Endpoint para recibir el Excel con las nuevas columnas y procesarlo en segundo plano."""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Formato inválido. Sube un archivo .xlsx o .xls")

    content = await file.read()

    background_tasks.add_task(procesar_y_llamar_paquete, content, file.filename, anio_cargue, calendario, usuario_apex)

    return {
        "status": "success",
        "message": f"Archivo '{file.filename}' recibido para el usuario '{usuario_apex}'. Procesando en segundo plano.",
        "file_size_bytes": len(content)
    }