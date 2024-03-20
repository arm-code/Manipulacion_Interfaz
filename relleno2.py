import openpyxl
import pyautogui
import time

portada = """
 __________________________________
|              ____                |
|   ________  / / /__  ____  ____  |
|  / ___/ _ \/ / / _ \/ __ \/ __ \ |
| / /  /  __/ / /  __/ / / / /_/ / |
|/_/   \___/_/_/\___/_/ /_/\____/  |
|                                  |
|  by: arm-code (GPL License)      |
|__________________________________|
"""


print(portada)
# MATERIAS PARA EL PLAN 22, VAN DESDE LA MATERIA 01 HASTA LA 26 (DE LA 21 SE SALTA A LA 26)
number_asignatura = ['01', '02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','26']

# DEBEMOS CAMBIAR LA ETAPA, FASE Y EL PLAN
name_asignatura = '2403A22'

# DEBEMOS COLOCAR EL NOMBRE DE ARCHIVO QUE CONTIENE LAS PLANTILLAS
excel = '2403-A.xlsx'
workbook = openpyxl.load_workbook(excel, data_only=True)

print("COLOQUE EL CURSOR EN LA VENTANA DEL SIOSAD PLANTILLAS.")
print('(EN EL ESPACIO DE LA ASIGNATURA)')

print('\nLEYENDO LOS DATOS DEL EXCEL...')
time.sleep(5)

n = 0

# Iterar sobre cada hoja del libro
for sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]
    
    num_cols = sheet.max_column
    num_rows = sheet.max_row

    print('\nLOS SIGUIENTES DATOS SERAN GUARDADOS:')
    print(f'Plantilla: materia {sheet_name}')    

    # Definir el rango de columnas desde B hasta AE
    columnas_a_leer = list(sheet.iter_cols(min_col=2, max_col=31, min_row=1, max_row=2))

    # Iterar sobre cada fila de la hoja
    for row_index in range(2, 3):  # Solo filas 1 y 2
        data_to_enter = []

        # Iterar sobre cada celda en el rango de columnas definido
        for col_index in range(len(columnas_a_leer)):
            cell_value = columnas_a_leer[col_index][row_index - 1].value
            data_to_enter.append(str(cell_value))

        # Imprimir los datos de la fila actual
        #print(data_to_enter)
        for i in range(len(data_to_enter)):
            if (i+1) % 5 == 0:
                print('|',data_to_enter[i], '|')
            else:
                print('|',data_to_enter[i], end=' ')
        print('\nINGRESANDO LOS DATOS EN EL SIOSAD...')
        # se deben verificar las coordenadas del click, de lo contrario se va ir a otro lado la captura
        pyautogui.click(x=150, y=150)

        # NUMERO DE ASIGNATURA
        pyautogui.write(number_asignatura[n])
        pyautogui.press('tab')

        # CLAVE DE PLANTILLA
        pyautogui.write(number_asignatura[n] + name_asignatura)
        pyautogui.press('tab')

        # NUMERO DE RESPUESTAS
        pyautogui.write('30')
        pyautogui.press('tab')

        #BASE DE CODIFICACION
        pyautogui.write('1')
        pyautogui.press('tab')
        
        # NORMAS DE CODIFICACION
        # LAS NORMAS DE CODIFICACION VARIAN PARA MATERIAS DE 30 Y 40 PREGUNTAS.
        # PARA MATERIAS DE 30 PREGUNTAS USAMOS [17,19,21,24,27]
        # PARA MATERIAS DE 40 PREGUNTAS USAMOS [23,27,31,34,37]

        pyautogui.write('16')   # NORMA 1
        pyautogui.press('tab')  
        pyautogui.write('18')   # NORMA 2
        pyautogui.press('tab')  
        pyautogui.write('20')   # NORMA 3
        pyautogui.press('tab')
        pyautogui.write('23')   # NORMA 4
        pyautogui.press('tab')
        pyautogui.write('26')   # NORMA 5
        pyautogui.press('tab')

        # RESPUESTAS
        for i in range(1,31):
            if i % 5 == 0 :
                pyautogui.write(data_to_enter[i-1])
                pyautogui.press('tab')
                
            else:
                pyautogui.write(data_to_enter[i-1])
        
        input('\nREVISE CUIDADOSAMENTE LA CAPTURA \nPRESIONE ENTER PARA CONTINUAR...\n>')
        print('enter')
        print('EN SEGUIDA VUELVA A COLOCAR EL CURSOR EN LA VENTANA DEL SIOSAD PLANTILLAS.')        
        time.sleep(3)

        # GUARDAR PLANTILLA
        pyautogui.press('f2')
        pyautogui.press('enter')
        pyautogui.press('enter')
        print('\nCAPTURA EXITOSA!\n')
        try:  
            n = n + 1
            print('\tSIGUIENTE MATERIA: ', number_asignatura[n])
            print("\t|--------------------------|")
            input('\t|PARA DETENER EL PROGRAMA: |\n\t|->PRESIONE CTRL + C       |\n\t|PARA AGREGAR SIG MATERIA? |\n\t|->PRESIONE ENTER          |\n\t>')
            
            print('enter')
            time.sleep(5)
        except:
            print('Al parecer se ha terminado de leer el EXCEL.')
    workbook.close()