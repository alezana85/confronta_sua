import pandas as pd
import numpy as np
import pyfiglet

# Generar el texto en ASCII art
ascii_art = pyfiglet.figlet_format("SMEXYKAL")

# Dividir el ASCII art en líneas
ascii_lines = ascii_art.split('\n')

# Calcular el ancho del cuadro
max_length = max(len(line) for line in ascii_lines)
frame_width = max_length + 4  # Añadir espacio para los bordes

# Crear el cuadro
print('*' * frame_width)
for line in ascii_lines:
    print(f"* {line.ljust(max_length)} *")
print('*' * frame_width)
print('-' * 80)
print("Este programa compara los archivos de la cedula del SUA y la emision (EMA) para \n identificar diferencias en los montos de las cuotas obrero-patronales.")
print('\n')
print('Si tienes dudas lee el archivo README.md o contacta a tu analista de sistemas.')
print('\n')
print('Puedes copiar y pegar las rutas de las carpetas para evitar errores.')
print('\n')
print('Puedes encontrar el código fuente en el github de smexykal, solo googlea "Alejandro Lezana Github".')
print('-' * 80)

# Solicitar la ruta de la carpeta mediante input
ruta_carpeta_sua_m = input("Por favor, ingrese la ruta de la carpeta de la cedula del SUA: ")
ruta_carpeta_ema = input("Por favor, ingrese la ruta de la carpeta de la emision (EMA): ")

# Solicitar el nombre del archivo (sin la extensión) mediante input
nombre_archivo_ema = input("Por favor, ingrese el nombre del archivo (sin la extensión .xls): ")

# Nombre del archivo
nombre_archivo_sua_m = 'cedula oportuno obr-pat_gbl.xls'

# Solicitar la ruta de la carpeta para guardar los archivos limpios
ruta_guardar = input("Por favor, ingrese la ruta de la carpeta para guardar los archivos limpios: ")

# Concatenar la ruta de la carpeta con el nombre del archivo
ruta_completa_sua_m = f"{ruta_carpeta_sua_m}\\{nombre_archivo_sua_m}"

# Leer el archivo Excel sin encabezados
sua_mensual = pd.read_excel(
    ruta_completa_sua_m, 
    engine='xlrd',  # Especificar el motor para archivos .xls
    header=None,  # Indicar que no hay encabezados en el archivo  
)

# Asignar nombres a las columnas
sua_mensual.columns = ['nss', 
                       'nombre', 
                       'dias', 
                       'sdi', 
                       'licencia', 
                       'incapacidades', 
                       'ausentismos', 
                       'cuota_fija', 
                       'excedente_patronal', 
                       'excedente_obrero',  
                       'prestaciones_patronal', 
                       'prestaciones_obrero', 
                       'gastos_medicos_patronal', 
                       'gastos_medicos_obrero',
                       'riesgo_trabajo',
                       'invalidez_vida_patronal',
                       'invalidez_vida_obrero',
                       'guarderia',
                       'total_patronal',
                       'total_obrero',
                       'total']

# Identificar valores no numéricos en la columna 'licencia'
non_numeric_mask = sua_mensual['incapacidades'].apply(lambda x: not str(x).isdigit())

# Mover valores no numéricos a la columna 'nombre'
sua_mensual.loc[non_numeric_mask, 'nombre'] = sua_mensual.loc[non_numeric_mask, 'incapacidades']

# Reemplazar los valores movidos en la columna 'licencia' con NaN
sua_mensual.loc[non_numeric_mask, 'incapacidades'] = np.nan

# Eliminar las primeras dos filas de las columnas desde 'dias' hasta la última columna y desplazar los datos hacia arriba
cols_to_shift = sua_mensual.columns[2:]  # Seleccionar columnas desde 'dias' hasta la última
sua_mensual[cols_to_shift] = sua_mensual[cols_to_shift].shift(-2)

# Rellenar las últimas dos filas con NaN para mantener el tamaño del DataFrame
sua_mensual.iloc[-2:, 2:] = np.nan

# Definir la expresión regular para el formato xx-xx-xx-xxxx-x
regex_pattern = r'^\d{2}-\d{2}-\d{2}-\d{4}-\d$'

# Aplicar la expresión regular a la columna 'nss' para identificar las filas que cumplen con el formato
valid_format_mask = sua_mensual['nss'].astype(str).str.match(regex_pattern)

# Identificar filas que no cumplen con el formato del regex pero contienen datos en la columna 'dias'
invalid_format_mask = ~valid_format_mask & sua_mensual['dias'].notna() & (sua_mensual['dias'] < 32)

# Asignar el valor de 'nss' y 'nombre' de la fila más cercana hacia arriba que cumple con el formato del regex
for idx in sua_mensual[invalid_format_mask].index:
    closest_valid_idx = sua_mensual.loc[:idx-1][valid_format_mask].last_valid_index()
    if closest_valid_idx is not None:
        sua_mensual.at[idx, 'nss'] = sua_mensual.at[closest_valid_idx, 'nss']
        sua_mensual.at[idx, 'nombre'] = sua_mensual.at[closest_valid_idx, 'nombre']

# Aplicar la expresión regular a la columna 'nss' para identificar las filas que cumplen con el formato
valid_format_mask = sua_mensual['nss'].astype(str).str.match(regex_pattern)

# Filtrar el DataFrame para mantener solo las filas que cumplen con el formato
sua_mensual = sua_mensual[valid_format_mask]

# Rellenar valores NaN con 0 antes de convertir a int
#sua_mensual[['dias', 'licencia', 'incapacidades', 'ausentismos']] = sua_mensual[['dias', 'licencia', 'incapacidades', 'ausentismos']].fillna(0)

# Convertir columna 'dias', 'licencia', 'incapacidades', 'ausentismos' a int
sua_mensual[['dias', 'licencia', 'incapacidades', 'ausentismos']] = sua_mensual[['dias', 'licencia', 'incapacidades', 'ausentismos']].astype(int)

# Convertir columnas restantes a float
sua_mensual[sua_mensual.columns[6:]] = sua_mensual[sua_mensual.columns[6:]].astype(float)

# Agrupar por 'nss' y sumar los valores de las columnas
sua_mensual = sua_mensual.groupby('nss', as_index=False).sum()

# Ordenar el DataFrame por nombre y resetear los índices
sua_mensual = sua_mensual.sort_values('nombre').reset_index(drop=True)

# Eliminar guiones de la columna 'nss' y convertir a uint64
sua_mensual['nss'] = sua_mensual['nss'].str.replace('-', '').astype('uint64')

# Remplazar de la columna 'nombre' los '/' por 'Ñ'
sua_mensual['nombre'] = sua_mensual['nombre'].str.replace('/', 'Ñ')

# Concatenar la ruta de la carpeta con el nombre del archivo y la extensión .xls
ruta_completa_ema = f"{ruta_carpeta_ema}\\{nombre_archivo_ema}.xls"

ema = pd.read_excel(ruta_completa_ema,
                    sheet_name=1,
                    header=4,
                    dtype={'nss': str}
                    )

# Asignar nombres a las columnas
ema.columns = ['nss',
               'nombre',
               'origen',
               'tipo',
               'fecha',
               'dias',
               'sdi',
               'cuota_fija',
               'excedente_patronal',
               'excedente_obrero',
               'prestaciones_patronal',
               'prestaciones_obrero',
               'gastos_medicos_patronal',
               'gastos_medicos_obrero',
               'riesgo_trabajo',
               'invalidez_vida_patronal',
               'invalidez_vida_obrero',
               'guarderia',
               'total']

# Verificar que la columna 'nss' se mantenga como texto
ema['nss'] = ema['nss'].astype(str)

# Eliminar filas que en la columna 'tipo' tengan el numero 2
ema = ema[~ema['tipo'].eq(2)]

# Eliminar columnas innecesarias
ema.drop(columns=['origen', 'tipo', 'fecha'], inplace=True)

# Agrupar el DataFrame por 'nss' y 'nombre' y sumar los valores
ema = ema.groupby(['nss', 'nombre'], as_index=False).sum()

# Convertir la columna 'nss' a uint64
ema['nss'] = ema['nss'].astype('uint64')

# Convertir las columnas 'excedente_patronal' y 'excedente_obrero' a float64
ema[['excedente_patronal', 'excedente_obrero']] = ema[['excedente_patronal', 'excedente_obrero']].astype('float64')

# Eliminar espacios extra en la columna 'nombre'
ema['nombre'] = ema['nombre'].str.strip()

# Remplazar de la columna 'nombre' los '#' por 'Ñ'
ema['nombre'] = ema['nombre'].str.replace('#', 'Ñ')

# Ordenar el DataFrame por 'nombre' y resetear los índices
ema.sort_values(by='nombre', inplace=True)
ema.reset_index(drop=True, inplace=True)

# Copiar el DataFrame de SUA mensual
sua_vs_ema = sua_mensual.copy()

# Eliminar columnas 'total_patronal', 'total_obrero' de sua_vs_ema
sua_vs_ema.drop(columns=['total_patronal', 'total_obrero'], inplace=True)

# Iterar sobre cada fila de sua_vs_ema
for index, row in sua_vs_ema.iterrows():
    nss = row['nss']
    
    # Buscar el nss correspondiente en ema
    ema_row = ema[ema['nss'] == nss]
    
    if not ema_row.empty:
        # Realizar la resta de las columnas especificadas
        sua_vs_ema.at[index, 'dias'] = row['dias'] - ema_row.iloc[0]['dias']
        sua_vs_ema.at[index, 'sdi'] = row['sdi'] - ema_row.iloc[0]['sdi']
        sua_vs_ema.at[index, 'cuota_fija'] = row['cuota_fija'] - ema_row.iloc[0]['cuota_fija']
        sua_vs_ema.at[index, 'excedente_patronal'] = row['excedente_patronal'] - ema_row.iloc[0]['excedente_patronal']
        sua_vs_ema.at[index, 'excedente_obrero'] = row['excedente_obrero'] - ema_row.iloc[0]['excedente_obrero']
        sua_vs_ema.at[index, 'prestaciones_patronal'] = row['prestaciones_patronal'] - ema_row.iloc[0]['prestaciones_patronal']
        sua_vs_ema.at[index, 'prestaciones_obrero'] = row['prestaciones_obrero'] - ema_row.iloc[0]['prestaciones_obrero']
        sua_vs_ema.at[index, 'gastos_medicos_patronal'] = row['gastos_medicos_patronal'] - ema_row.iloc[0]['gastos_medicos_patronal']
        sua_vs_ema.at[index, 'gastos_medicos_obrero'] = row['gastos_medicos_obrero'] - ema_row.iloc[0]['gastos_medicos_obrero']
        sua_vs_ema.at[index, 'riesgo_trabajo'] = row['riesgo_trabajo'] - ema_row.iloc[0]['riesgo_trabajo']
        sua_vs_ema.at[index, 'invalidez_vida_patronal'] = row['invalidez_vida_patronal'] - ema_row.iloc[0]['invalidez_vida_patronal']
        sua_vs_ema.at[index, 'invalidez_vida_obrero'] = row['invalidez_vida_obrero'] - ema_row.iloc[0]['invalidez_vida_obrero']
        sua_vs_ema.at[index, 'guarderia'] = row['guarderia'] - ema_row.iloc[0]['guarderia']
        sua_vs_ema.at[index, 'total'] = row['total'] - ema_row.iloc[0]['total']
    else:
        # Asignar NaN si no se encuentra el nss en ema
        sua_vs_ema.at[index, 'dias'] = np.nan
        sua_vs_ema.at[index, 'sdi'] = np.nan
        sua_vs_ema.at[index, 'cuota_fija'] = np.nan
        sua_vs_ema.at[index, 'excedente_patronal'] = np.nan
        sua_vs_ema.at[index, 'excedente_obrero'] = np.nan
        sua_vs_ema.at[index, 'prestaciones_patronal'] = np.nan
        sua_vs_ema.at[index, 'prestaciones_obrero'] = np.nan
        sua_vs_ema.at[index, 'gastos_medicos_patronal'] = np.nan
        sua_vs_ema.at[index, 'gastos_medicos_obrero'] = np.nan
        sua_vs_ema.at[index, 'riesgo_trabajo'] = np.nan
        sua_vs_ema.at[index, 'invalidez_vida_patronal'] = np.nan
        sua_vs_ema.at[index, 'invalidez_vida_obrero'] = np.nan
        sua_vs_ema.at[index, 'guarderia'] = np.nan
        sua_vs_ema.at[index, 'total'] = np.nan

# Crear columna 'observaciones' en sua_vs_ema
sua_vs_ema['observacion_sistema'] = ''

for index, row in sua_vs_ema.iterrows():
    if row['total'] == 0:
        sua_vs_ema.at[index, 'observacion_sistema'] = 'SIN DIFERENCIAS'
    elif row['total'] != 0 and row['incapacidades'] != 0:
        sua_vs_ema.at[index, 'observacion_sistema'] = 'INCAPACIDAD'
    elif row['total'] != 0 and row['dias'] != 0:
        sua_vs_ema.at[index, 'observacion_sistema'] = 'DIFERENCIA EN DIAS'
    elif row['total'] != 0 and row[['cuota_fija', 'excedente_patronal', 'excedente_obrero', 'prestaciones_patronal', 'prestaciones_obrero', 'gastos_medicos_patronal', 'gastos_medicos_obrero', 'riesgo_trabajo', 'guarderia']].sum() == 0:
        sua_vs_ema.at[index, 'observacion_sistema'] = 'PENSIONADO'
    elif np.isnan(row['total']):
        sua_vs_ema.at[index, 'observacion_sistema'] = 'NO ESTA EN LA EMISION'
    else:
        sua_vs_ema.at[index, 'observacion_sistema'] = 'OTRA DIFERENCIA'

# Añadir columna de 'observacion_usuario' en sua_vs_ema vacia
sua_vs_ema['observacion_usuario'] = ''

# Copiar el DataFrame de EMA
ema_vs_sua = ema.copy()

# Iterar sobre cada fila de ema_vs_sua
for index, row in ema_vs_sua.iterrows():
    nss = row['nss']
    
    # Buscar el nss correspondiente en sua_vs_ema
    sua_row = sua_mensual[sua_mensual['nss'] == nss]
    
    if not sua_row.empty:
        # Realizar la resta de las columnas especificadas
        ema_vs_sua.at[index, 'dias'] = row['dias'] - sua_row.iloc[0]['dias']
        ema_vs_sua.at[index, 'sdi'] = row['sdi'] - sua_row.iloc[0]['sdi']
        ema_vs_sua.at[index, 'cuota_fija'] = row['cuota_fija'] - sua_row.iloc[0]['cuota_fija']
        ema_vs_sua.at[index, 'excedente_patronal'] = row['excedente_patronal'] - sua_row.iloc[0]['excedente_patronal']
        ema_vs_sua.at[index, 'excedente_obrero'] = row['excedente_obrero'] - sua_row.iloc[0]['excedente_obrero']
        ema_vs_sua.at[index, 'prestaciones_patronal'] = row['prestaciones_patronal'] - sua_row.iloc[0]['prestaciones_patronal']
        ema_vs_sua.at[index, 'prestaciones_obrero'] = row['prestaciones_obrero'] - sua_row.iloc[0]['prestaciones_obrero']
        ema_vs_sua.at[index, 'gastos_medicos_patronal'] = row['gastos_medicos_patronal'] - sua_row.iloc[0]['gastos_medicos_patronal']
        ema_vs_sua.at[index, 'gastos_medicos_obrero'] = row['gastos_medicos_obrero'] - sua_row.iloc[0]['gastos_medicos_obrero']
        ema_vs_sua.at[index, 'riesgo_trabajo'] = row['riesgo_trabajo'] - sua_row.iloc[0]['riesgo_trabajo']
        ema_vs_sua.at[index, 'invalidez_vida_patronal'] = row['invalidez_vida_patronal'] - sua_row.iloc[0]['invalidez_vida_patronal']
        ema_vs_sua.at[index, 'invalidez_vida_obrero'] = row['invalidez_vida_obrero'] - sua_row.iloc[0]['invalidez_vida_obrero']
        ema_vs_sua.at[index, 'guarderia'] = row['guarderia'] - sua_row.iloc[0]['guarderia']
        ema_vs_sua.at[index, 'total'] = row['total'] - sua_row.iloc[0]['total']
    else:
        # Asignar NaN si no se encuentra el nss en sua_vs_ema
        ema_vs_sua.at[index, 'dias'] = np.nan
        ema_vs_sua.at[index, 'sdi'] = np.nan
        ema_vs_sua.at[index, 'cuota_fija'] = np.nan
        ema_vs_sua.at[index, 'excedente_patronal'] = np.nan
        ema_vs_sua.at[index, 'excedente_obrero'] = np.nan
        ema_vs_sua.at[index, 'prestaciones_patronal'] = np.nan
        ema_vs_sua.at[index, 'prestaciones_obrero'] = np.nan
        ema_vs_sua.at[index, 'gastos_medicos_patronal'] = np.nan
        ema_vs_sua.at[index, 'gastos_medicos_obrero'] = np.nan
        ema_vs_sua.at[index, 'riesgo_trabajo'] = np.nan
        ema_vs_sua.at[index, 'invalidez_vida_patronal'] = np.nan
        ema_vs_sua.at[index, 'invalidez_vida_obrero'] = np.nan
        ema_vs_sua.at[index, 'guarderia'] = np.nan
        ema_vs_sua.at[index, 'total'] = np.nan

# Crear columna 'observaciones' en ema_vs_sua
ema_vs_sua['observacion_sistema'] = ''

for index, row in ema_vs_sua.iterrows():
    if row['total'] == 0:
        ema_vs_sua.at[index, 'observacion_sistema'] = 'SIN DIFERENCIAS'
    elif row['total'] != 0 and row['dias'] != 0:
        ema_vs_sua.at[index, 'observacion_sistema'] = 'DIFERENCIA EN DIAS'
    elif row['total'] != 0 and row[['cuota_fija', 'excedente_patronal', 'excedente_obrero', 'prestaciones_patronal', 'prestaciones_obrero', 'gastos_medicos_patronal', 'gastos_medicos_obrero', 'riesgo_trabajo', 'guarderia']].sum() == 0:
        ema_vs_sua.at[index, 'observacion_sistema'] = 'PENSIONADO'
    elif np.isnan(row['total']):
        ema_vs_sua.at[index, 'observacion_sistema'] = 'NO ESTA EN SUA'
    else:
        ema_vs_sua.at[index, 'observacion_sistema'] = 'DIFERENCIA'

# Si en ema_vs_sua en la columna 'observacion_sistema' es 'DIFERENCIA' y en sua_vs_ema en la columna 'observacion_sistema' es 'INCAPACIDAD' cambiar en ema_vs_sua a 'INCAPACIDAD'
for index, row in ema_vs_sua.iterrows():
    if row['observacion_sistema'] == 'DIFERENCIA':
        nss = row['nss']
        sua_vs_ema_row = sua_vs_ema[sua_vs_ema['nss'] == nss]
        if not sua_vs_ema_row.empty and sua_vs_ema_row.iloc[0]['observacion_sistema'] == 'INCAPACIDAD':
            ema_vs_sua.at[index, 'observacion_sistema'] = 'INCAPACIDAD'

'''En el DataFrame sua_vs_ema hay que colocar en la columna 'observacion_usuario' lo siguiente:
   hay que iterar en el DataFrame ema y buscar el nss que le corresponde para dividir su 'total' entre 'dias' y dividirlo
   entre 'incapacidades' del DataFrame sua_mensual iterando tambien en este DataFrame para localizar el 'nss '''

for index, row in sua_vs_ema.iterrows():
    nss = row['nss']
    
    # Buscar el nss correspondiente en ema
    ema_row = ema[ema['nss'] == nss]
    
    if not ema_row.empty:
        # Buscar el nss correspondiente en sua_mensual
        sua_row = sua_mensual[sua_mensual['nss'] == nss]
        
        if not sua_row.empty:
            # Calcular el valor de 'observacion_usuario'
            total_ema = ema_row.iloc[0]['total']
            dias_ema = ema_row.iloc[0]['dias']
            incapacidades_sua = sua_row.iloc[0]['incapacidades']
            
            if dias_ema != 0 and incapacidades_sua != 0:
                observacion_usuario = total_ema / dias_ema * incapacidades_sua
                sua_vs_ema.at[index, 'observacion_usuario'] = observacion_usuario
            else:
                sua_vs_ema.at[index, 'observacion_usuario'] = np.nan
        else:
            sua_vs_ema.at[index, 'observacion_usuario'] = np.nan
    else:
        sua_vs_ema.at[index, 'observacion_usuario'] = np.nan

# Colocar una columna 'validacion_nombre' en sua_vs_ema entre las columnas 'nombre' y 'dias'
sua_vs_ema.insert(2, 'validacion_nombre', '')

# Iterar sobre cada fila de sua_vs_ema y comparar el nombre con el DataFrame ema_vs_sua, si el nombre es igual asignar 'OK' en 'validacion_nombre', si no es igual asignar 'NOMBRE DIFERENTE', si no esta en ema_vs_sua asignar NaN
for index, row in sua_vs_ema.iterrows():
    nss = row['nss']
    
    # Buscar el nss correspondiente en ema_vs_sua
    ema_vs_sua_row = ema_vs_sua[ema_vs_sua['nss'] == nss]
    
    if not ema_vs_sua_row.empty:
        if row['nombre'] == ema_vs_sua_row.iloc[0]['nombre']:
            sua_vs_ema.at[index, 'validacion_nombre'] = 'OK'
        else:
            sua_vs_ema.at[index, 'validacion_nombre'] = 'NOMBRE DIFERENTE'
    else:
        sua_vs_ema.at[index, 'validacion_nombre'] = np.nan

'''Exportar DataFrames au solo archivo Excel como pestañas, el DataFrame sua_mensual_limpia con el nombre
   SUA, el DataFrame ema_limpia con el nombre EMA, el DataFrame ema_vs_sua como CONFRONTA EMA y el DataFrame
   sua_vs_ema como CONFRONTA SUA'''
with pd.ExcelWriter(f"{ruta_guardar}\\CONFRONTA_MENSUAL.xlsx") as writer:
    sua_mensual.to_excel(writer, sheet_name='SUA', index=False)
    ema.to_excel(writer, sheet_name='EMA', index=False)
    ema_vs_sua.to_excel(writer, sheet_name='CONFRONTA EMA', index=False)
    sua_vs_ema.to_excel(writer, sheet_name='CONFRONTA SUA', index=False)

print("Archivo guardado exitosamente.")