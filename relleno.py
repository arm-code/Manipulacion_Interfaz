import openpyxl
import pyautogui
import time

#number_asignatura = ['01', '02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','26']
number_asignatura = ['20', '21','30','41','53','60']
name_asignatura = '2308B33'

excel = 'relleno.xlsx'

workbook = openpyxl.load_workbook(excel)
sheet = workbook.active

num_cols = sheet.max_column
num_rows = sheet.max_row

time.sleep(3)
print('ESPERE MIENTRAS SE CARGAN LOS DATOS...')
n = 0
for row_index in range(1, num_rows + 1):    
    data_to_enter = []
    for col_index in range(1, num_cols + 1):
        cell_value = sheet.cell(row = row_index, column=col_index).value
        #print(cell_value)
        data_to_enter.append(str(cell_value))
    
    pyautogui.click(x=100, y= 100)

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
    pyautogui.write('17')   # NORMA 1
    pyautogui.press('tab')  
    pyautogui.write('19')   # NORMA 2
    pyautogui.press('tab')  
    pyautogui.write('21')   # NORMA 3
    pyautogui.press('tab')
    pyautogui.write('24')   # NORMA 4
    pyautogui.press('tab')
    pyautogui.write('27')   # NORMA 5
    pyautogui.press('tab')

    # RESPUESTAS
    for i in range(1,31):
        if i % 5 == 0 :
            pyautogui.write(data_to_enter[i-1])
            pyautogui.press('tab')
            
        else:
            pyautogui.write(data_to_enter[i-1])
    
    input('REVISE CUIDADOSAMENTE LA CAPTURA Y PRESIONE ENTER PARA CONTINUAR...')
    time.sleep(3)

    # GUARDAR PLANTILLA
    pyautogui.press('f2')
    pyautogui.press('enter')
    pyautogui.press('enter')
    print('CAPTURA EXITOSA!')    
    n = n + 1
    print('VALOR DE n = ', n)
    input('DESEA SALIR? PRESIONE CTRL + C, DE LO CONTRARROP PRESIONE ENTER...')
    time.sleep(5)
workbook.close()
