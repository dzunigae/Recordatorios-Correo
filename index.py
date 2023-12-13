import pandas as pd

ADMINISTRADORES = 'assets/Administradores PoP.xlsx'
NUEVAS_CONVOCATORIAS = 'assets/Hoja de Nuevas Convocatorias.xlsx'

REPORTE = 'out/reporte.xlsx'

def reporte_correos_convocatorias(ADMINISTRADORES,NUEVAS_CONVOCATORIAS,REPORTE):
    #Abrir el archivo Hoja de nuevas convocatorias, el cual contiene la información de las convocatorias publicadas
    # a través del formulario
    HOJA_NUEVAS_CONVOCATORIAS_DF = pd.read_excel(NUEVAS_CONVOCATORIAS)
    HOJA_NUEVAS_CONVOCATORIAS_DF.index = HOJA_NUEVAS_CONVOCATORIAS_DF.index + 2
    HOJA_NUEVAS_CONVOCATORIAS_DF = HOJA_NUEVAS_CONVOCATORIAS_DF.reset_index(drop=False)

    #Eliminar aquellas que ya cuentan con todos sus programas aprobados o rechazados
    HOJA_NUEVAS_CONVOCATORIAS_DF.drop(HOJA_NUEVAS_CONVOCATORIAS_DF[HOJA_NUEVAS_CONVOCATORIAS_DF['Columna del administrador'] == 'x'].index, inplace=True)

    #Filtrar únicamente las columnas de la Hoja de nuevas Convocatorias que nos sirven
    HOJA_NUEVAS_CONVOCATORIAS_DF = HOJA_NUEVAS_CONVOCATORIAS_DF[['index','Dirección de correo electrónico','Nombre de la entidad','Título de la convocatoria','Programas académicos requeridos','Decisión de aprobación o rechazo']].copy()

    

    #print(HOJA_NUEVAS_CONVOCATORIAS_DF)
    HOJA_NUEVAS_CONVOCATORIAS_DF.to_excel(REPORTE, index=False)

reporte_correos_convocatorias(ADMINISTRADORES,NUEVAS_CONVOCATORIAS,REPORTE)