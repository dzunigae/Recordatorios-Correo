import pandas as pd

CURRICULUMS = './Curriculums/assets/Hoja de Curriculums.xlsx'

REPORTE = './Curriculums/out/Reporte.xlsx'

def tratamiento_Hoja_de_Curriculums(CURRICULUMS,REPORTE):

    #Variables
    SET_REDUNDANCIA = {
        "Convocatorias a las que desea aplicar",
        "Adjunte la hoja de vida en formato PDF marcada con su nombre completo y programa curricular.",
        "¿Desea terminar?"
    }

    #Abrir los archivos necesarios
    CURRICULUMS_DF = pd.read_excel(CURRICULUMS)

    #Creación de nuevo Data Frame sin redundancia
    REPORTE_DF = pd.DataFrame(columns=['Marca temporal',
                                       'Dirección de correo electrónico',
                                       'Unnamed: 2',
                                       'Nombres y apellidos',
                                       'Correo electrónico de la universidad',
                                       'Número de documento de identidad',
                                       'Tipo de admisión',
                                       'Programa al que pertenece',
                                       'Convocatorias a las que desea aplicar',
                                       'Adjunte la hoja de vida en formato PDF marcada con su nombre completo y programa curricular.',
                                       '¿Desea terminar?'])
    
    #Rellenar Data Frame anterior
    for i in range(len(CURRICULUMS_DF)):
        REGISTRO_ACTUAL = CURRICULUMS_DF.loc[i]
        REGISTRO_ACTUAL = REGISTRO_ACTUAL.dropna()
        REGISTRO_ACTUAL_LISTA_COLUMNAS = REGISTRO_ACTUAL.index.tolist()
        NUEVA_ENTRADA = {}

        for j in REGISTRO_ACTUAL_LISTA_COLUMNAS:
            PARTE_DE_SET = False
            #Mirar si el título de la columna coincide con SET_REDUNDANCIA
            for k in SET_REDUNDANCIA:
                if k in j:
                    PARTE_DE_SET = True
            if PARTE_DE_SET:
                t = j[:len(j)-2]
                NUEVA_ENTRADA[t] = REGISTRO_ACTUAL[j]
            else:
                NUEVA_ENTRADA[j] = REGISTRO_ACTUAL[j]

        REPORTE_DF = REPORTE_DF._append(NUEVA_ENTRADA,ignore_index=True)

    #Convertir el reporte a Excel
    REPORTE_DF.to_excel(REPORTE,index=False)

tratamiento_Hoja_de_Curriculums(CURRICULUMS,REPORTE)