from importlib.resources import path
from random import choice as ch
from string import ascii_lowercase as asc
import win32com.client as com
import mariadb as sql
import BOMtoSQL_library as bsql
import os
import sys
import boto3
from botocore.client import ClientError
import json

config_file = json.load(open("./config.json", "r"))
sql_config = config_file["sql_db_credentials"]
s3_config = config_file["s3_bucket_credentials"]

if len(sys.argv) < 1:
    print("Faltan argumentos. Ejecutar: python3 ./export_BOM.py [Ruta de ensamble principal]")
    sys.exit(1)

main_assembly_path = sys.argv[1]

print(f"Iniciando exportación con ruta de ensamble principal: {main_assembly_path}")

print("Conectando con base de datos")

try:
    conn = sql.connect(
        user=sql_config["user"],
        password=sql_config["password"],
        host=sql_config["host"],
        port=sql_config["port"],
        database=sql_config["database"]
    )
    print("Conexión a DB exitosa")
except sql.Error as e:
    print(f"Error conectandose al servidor de MariaDB: {e}")
    sys.exit(2)

cur = conn.cursor()

s3 = boto3.resource(service_name='s3', aws_access_key_id=s3_config["aws_access_key_id"], aws_secret_access_key=s3_config["aws_secret_access_key"])

print("Conectando a kahl-bom-thumbnails bucket")

try:
    bucket = s3.Bucket('kahl-bom-thumbnails')
    s3.meta.client.head_bucket(Bucket=bucket.name)
    print("Conexión a kahl-bom-thumbnails bucket exitosa")
except ClientError as e:
    code = int(e.response['Error']['Code'])
    if code == 404: 
        print(f"Error {code} al conectarse al bucket kahl-bom-thumbnails. No se encontró. Comprobar el nombre del bucket.")
    if int(e.response['Error']['Code']) == 403: 
        print(f"Error {code} al conectarse al bucket kahl-bom-thumbnails. Acceso prohibido. Comprobar credenciales de acceso.")
    sys.exit(2)

print("Iniciando Inventor Server")

invApp = com.Dispatch('{B6B5DC40-96E3-11d2-B774-0060B0F159EF}')
mod = com.gencache.EnsureModule('{D98A091D-3A0F-4C3E-B36E-61F62068D488}', 0, 1, 0)
invApp = mod.Application(invApp)

if invApp.Type != 50331904: 
    print("Error al iniciar el Inventor Server")
    sys.exit(2)

print("Inventor Server iniciado con exito")    

oAssDoc = invApp.Documents.Open(main_assembly_path, False)
oAssDoc = mod.AssemblyDocument(oAssDoc)

print(f"Documento base {oAssDoc.FullDocumentName} abierto")

oDesignProp = oAssDoc.PropertySets.Item('{32853F0F-3444-11d1-9E93-0060B03C1CA6}')
oDesignProp = mod.PropertySet(oDesignProp)

project = oDesignProp.ItemByPropId(7).Value

if project == '':
    project = ''.join(ch(asc) for i in range(10))

print("Iniciando extracción de BOM")

data = [("0/", project, oDesignProp.ItemByPropId(5).Value, "Normal", 1)]
paths = [(oDesignProp.ItemByPropId(5).Value, oAssDoc.FullDocumentName)]

print(f"Elemento: {data} con ruta {paths[0][1]}")

oBOM =  oAssDoc.ComponentDefinition.BOM
oBOM = mod.BOM(oBOM)
oBOMRows = oBOM.BOMViews.Item(2).BOMRows
oBOMRows = mod.BOMRowsEnumerator(oBOMRows)

oBOM.StructuredViewEnabled = True
oBOM.StructuredViewFirstLevelOnly = False

print("BOM estructurada habilitada y visible.")

new_data, new_paths = bsql.BOMRowsToArray(prev_node="0/", project=project, Rows=oBOMRows)

data.extend(new_data)
paths.extend(new_paths)

bsql.savePathsArrayToFile(paths)

print("CSV con rutas de archivos correctamente guardado.")

paths_path = config_file["VB_script_connection"]["component_paths_list_location"]
dest_path = config_file["VB_script_connection"]["thumbnails_temp_path"]

print("Ejecutando script de exportación de Thumbnails")

os.system(".\\VB_scripts\\bin\\x64\\Release\\VB_scripts" + " " + paths_path + " " + project + " " + dest_path + " ")

print("Actualizando BD...")

bsql.verifyProjectEntry(project, cur)

cur.executemany("INSERT INTO BOM VALUES (%s, %s, %s, %s, %s)", data)

conn.commit()

print("Base de datos actualizada.")

print("Subiendo archivos a S3 bucket")

bsql.uploadDirectory(dest_path + project, bucket, project)

print("Thumbnails subidos exitosamente")

invApp.Quit
