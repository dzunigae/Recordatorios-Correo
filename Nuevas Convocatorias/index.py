import pandas as pd

ADMINISTRADORES = 'assets/Administradores PoP.xlsx'
NUEVAS_CONVOCATORIAS = 'assets/Hoja de Nuevas Convocatorias.xlsx'
YA_RESPONDIDAS = 'assets/Ya respondidas.xlsx'
YA_RESPONDIDAS_OUT = 'out/Ya respondidas out.xlsx'

REPORTE = 'out/reporte.xlsx'

def reporte_correos_convocatorias(ADMINISTRADORES,NUEVAS_CONVOCATORIAS,YA_RESPONDIDAS,REPORTE,YA_RESPONDIDAS_OUT):
    #Variables
    LISTA_INDICES_ELIMINAR = []
    CONVOCATORIAS_YA_RESPONDIDAS = set()

    #Abrir el archivo Hoja de nuevas convocatorias, el cual contiene la información de las convocatorias publicadas
    # a través del formulario
    HOJA_NUEVAS_CONVOCATORIAS_DF = pd.read_excel(NUEVAS_CONVOCATORIAS)
    HOJA_NUEVAS_CONVOCATORIAS_DF.index = HOJA_NUEVAS_CONVOCATORIAS_DF.index + 2
    HOJA_NUEVAS_CONVOCATORIAS_DF = HOJA_NUEVAS_CONVOCATORIAS_DF.reset_index(drop=False)

    #Abrir el archivo de Ya respondidas, para evitar la redundancia en la lista de convocatorias que debemos
    # atender
    YA_RESPONDIDAS_DF = pd.read_excel(YA_RESPONDIDAS)

    #Eliminar aquellas convocatorias que ya cuentan con todos sus programas aprobados o rechazados, tando del
    # DataFrame de Hoja de Nuevas Convocatorias como de Ya respondidas
    for i in range(len(HOJA_NUEVAS_CONVOCATORIAS_DF)):
        SELECCIONADO = HOJA_NUEVAS_CONVOCATORIAS_DF.loc[i]
        if SELECCIONADO['Columna del administrador'] == 'x':
            LISTA_INDICES_ELIMINAR.append(i)
            CONVOCATORIAS_YA_RESPONDIDAS.add(SELECCIONADO['index'])

    HOJA_NUEVAS_CONVOCATORIAS_DF = HOJA_NUEVAS_CONVOCATORIAS_DF.drop(LISTA_INDICES_ELIMINAR)
    YA_RESPONDIDAS_DF = YA_RESPONDIDAS_DF[~YA_RESPONDIDAS_DF['Convocatoria'].isin(CONVOCATORIAS_YA_RESPONDIDAS)]

    #Filtrar únicamente las columnas de la Hoja de nuevas Convocatorias que nos sirven
    HOJA_NUEVAS_CONVOCATORIAS_DF = HOJA_NUEVAS_CONVOCATORIAS_DF[['index','Dirección de correo electrónico','Nombre de la entidad','Título de la convocatoria','Programas académicos requeridos','Decisión de aprobación o rechazo']].copy()
    
    #Crear el Data Frame que sirve de soporte
    REPORTE_DF = pd.DataFrame(columns=['index','Programa Académico','Administrador PoP','Dirección de correo electrónico','Nombre de la entidad','Título de la convocatoria','Decisión de aprobación o rechazo'])

    #Rellenar el Data Frame del reporte con todas las convocatorias aún pendientes por respuesta
    

    #print(YA_RESPONDIDAS_DF)
    REPORTE_DF.to_excel(REPORTE, index=False)
    YA_RESPONDIDAS_DF.to_excel(YA_RESPONDIDAS_OUT, index=False)

reporte_correos_convocatorias(ADMINISTRADORES,NUEVAS_CONVOCATORIAS,YA_RESPONDIDAS,REPORTE,YA_RESPONDIDAS_OUT)