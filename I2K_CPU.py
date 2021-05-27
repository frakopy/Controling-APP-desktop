import pyautogui as pg
import os, time, datetime
from subprocess import Popen

#Obtenemos el direciotio acutal y lo guardamos en una vairable para poder regresar con os.chdir
DIR_ACTUAL = os.getcwd()

# #Necesitamos cambiarnos de directorio para que funcione la ejecucion de run.bat ya que usa algunas
#dependencias como jre que se encuentran en el directorio client
os.chdir(r'path where the app is')
Popen('run.bat')

# #Rotornamos a nuestros directorio de trabajo
os.chdir(DIR_ACTUAL)

def calcular_fechas():

    hoy = datetime.datetime.now()#Obtenemos la fecha actual para calcular las fechas
    ayer = hoy - datetime.timedelta(days=1)#Obtenemos la fecha de un dia anterior
    dia = str(ayer)[8:10]
    mes = str(ayer)[5:7]
    año = str(ayer)[0:4]

    ayer = dia + mes + año

    siete_dias_atras = hoy - datetime.timedelta(days=7)#Obtenemos la fecha de hace 7 dias
    dia = str(siete_dias_atras)[8:10]
    mes = str(siete_dias_atras)[5:7]
    año = str(siete_dias_atras)[0:4]

    siete_dias_atras = dia + mes + año

    return siete_dias_atras, ayer


def esperar_imagen(imagen):
    #Utilizamos el siguiente bucle para asegurarnos que el codigo continue hasta que aparezca en pantalla
    #la ventana donde relizaremos las operaciones
    while True:
        path = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes'
        path_imagen = path + imagen
        pos_imagen = pg.locateOnScreen(path_imagen, confidence=0.8)
        if pos_imagen != None: #Si ubica la imagen entonces login debe ser distinto de un valor None
            print('imagen localizda: ', imagen)
            break
        else:
            print('Intentando ubicar imagen: ', imagen)

esperar_imagen(r'\login.png')

#Localizamos el boton para seleccionar el nodo y obtenemos la posicion de donde esta
im1 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\selec_pps.png'
p1_cursor = pg.locateCenterOnScreen(im1, confidence=0.8)

#Teniendo la posicion del boton nos movemos hacia el mismo y hacemos click
pg.click(pg.moveTo(p1_cursor))
time.sleep(0.5)

#Con la funcion move logramos mover el mouse de forma realativa es decir en referencia la posicion del boton
#actual en la que se encuentra el cursor, en este caso lo movemos a partir de la posicion en la que se encuentra el cursor
#5 pixeles hacia la izquerda(-5 en x) y 15 pixeles hacia abajo(15 en y) y luego hacemos click
pg.move(-5,15)
pg.click()

#Localizamos la imagen donde se escbrie el passwor, damos click y escribimos el password
time.sleep(0.5)
im2 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\password.png'
p2_cursor = pg.locateCenterOnScreen(im2, confidence=0.8)
pg.click(pg.moveTo(p2_cursor))
pg.typewrite('xxxxx', interval=0.01)

#Localizamos la imagen donde se encuentra el boton de login y damos click
time.sleep(0.5)
im3 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\login.png'
p3_cursor = pg.locateCenterOnScreen(im3, confidence=0.8)
pg.click(pg.moveTo(p3_cursor))

esperar_imagen(r'\filter.png')

#Localizamos la imagen donde se encuentra la opcion de performance y damos click
time.sleep(0.5)
im4 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\performance.png'
p4_cursor = pg.locateCenterOnScreen(im4, confidence=0.8)
pg.click(pg.moveTo(p4_cursor))

#Localizamos la imagen donde se encuentra la opcion de performance monitoring y damos click
time.sleep(0.5)
im5 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\per_monitoring.png'
p5_cursor = pg.locateCenterOnScreen(im5, confidence=0.8)
pg.click(pg.moveTo(p5_cursor))

#Esperamos hasta que ubique la imagen donde esta la opcion de linux
esperar_imagen(r'\linux.png')

#Localizamos la imagen donde se encuentra el KPI, nos movemos a la derecha y 
#damos click izquierdo para seleccionarlo y luego click derecho para seleccionar
#la opcion de link to monitor
time.sleep(0.5)
im6 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\linux.png'
p6_cursor = pg.locateCenterOnScreen(im6, confidence=0.8)
pg.moveTo(p6_cursor)
time.sleep(0.5)
pg.move(80,0)
pg.click()
pg.click(button='right')
pg.move(20,112)
pg.click()

#Esperar por la imagen donde aparece el campo current para proceder con la ejecucion del codigo
esperar_imagen(r'\query_history.png')

#Ubicamos el boton query history data para buscar el reporte que necesitamos
time.sleep(0.5)
im7 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\query_history.png'
p7_cursor = pg.locateCenterOnScreen(im7, confidence=0.8)
pg.moveTo(p7_cursor)
pg.click()

#Esperar por la imagen donde aparece el campo fecha para proceder con la ejecucion del codigo
esperar_imagen(r'\date.png')

#Ubicamos el label date, nos movemos a la derecha, damos click, control+a y luego 
#escribimos la fecha inicial que tendra nuestro reporte, luego la fecha final y por ultimo
#damos click en el boton de query para obtener el reporte
time.sleep(0.5)

#Este script se corre los lunes por tanto obtenemos como fecha inicial la fecha de hace 7 siete_dias_atras
#y como fecha final la fecha de un dia anterior al Lunes es decir el Domingo.
fecha_inicial, fecha_final = calcular_fechas()

#Le concatemaos la hora que escribiremos en el campo de fechas del I2000
fecha_inicial = fecha_inicial + '0000'
fecha_final = fecha_final + '2359'

im8 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\date.png'
p8_cursor = pg.locateCenterOnScreen(im8, confidence=0.8)
pg.moveTo(p8_cursor)
pg.move(50,0)
pg.click()
pg.hotkey('ctrl', 'a')
pg.typewrite(fecha_inicial, interval=0.02)#Escribimos la fecha inicial del reporte
pg.move(0,50)
pg.click()
pg.hotkey('ctrl', 'a')
pg.typewrite(fecha_final, interval=0.02)#Escribimos la fecha final del reporte

#Ubicamos el boton query para obtener el reporte 
time.sleep(0.5)
im9 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\query.png'
p9_cursor = pg.locateCenterOnScreen(im9, confidence=0.8)
pg.moveTo(p9_cursor)
pg.click()

#Esperamos despues de dar click en el boton de query para que se desactiven los iconos que luego 
#estaremos validando hasta que esten activos de lo contrario el codigo seguira ejecutandose single
#esperar a que esten listos los datos de nuestro reporte
time.sleep(3)

#Esperar por la imagen donde aparece el campo average, para tener seguridad de que ya 
#estan listos nuestros datos del reporte
esperar_imagen(r'\iconos_activos.png')

#Ubicamos el label index, bajamos 15 pixeles, damos click derecho y seleccionamos la opcion
#vertical, debe ser verticl y no horizontal para poder tener el formato que espara el otro
#script que hace el calculo del CPU ademas de que debemos guardarlo en formato html
time.sleep(0.5)
im10 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\index.png'
p10_cursor = pg.locateCenterOnScreen(im10, confidence=0.8)
pg.moveTo(p10_cursor)
pg.move(0,15)
pg.click()
pg.click(button='right')
pg.move(10,120)
pg.click()

#Ubicamos el boton query nuevamente para volver a tener una referencia de movimiento ya que
#las ventanas cambian al seleccionar lo opcion de vertical 
time.sleep(0.5)
im11 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\query.png'
p11_cursor = pg.locateCenterOnScreen(im11, confidence=0.8)
pg.moveTo(p11_cursor)

#Ubicamos el label que dice CPU para luego bajar y seleccionar la opcion de save
time.sleep(1)
im12 = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\imagenes\cpu.png'
p12_cursor = pg.locateCenterOnScreen(im12, confidence=0.8)
pg.moveTo(p12_cursor)
pg.move(0,15)
pg.click(button='right')
pg.move(15,50)
pg.click()


#Ubicamos el label que dice filetype, nos movemos hacia abajo y seleccionamos .html
time.sleep(0.5)
im13 = r'D:/A_PYTHON/ProgramasPython/Control_NodosCA/Reporte_CPU_VPN-PPS/CPU_PPS_PA\imagenes/filetype.png'
p13_cursor = pg.locateCenterOnScreen(im13, confidence=0.8)
pg.moveTo(p13_cursor)
pg.move(45,0)
pg.click()
pg.move(0,35)
pg.click()

#Ubicamos el label que dice file name y nos movemos a la derecha, escribimos el nombre del reprote
# y damos enter, este reprote se guarda en el directorio client de donde esta el ejecutable del pps
time.sleep(0.5)
im14 = r'D:/A_PYTHON/ProgramasPython/Control_NodosCA/Reporte_CPU_VPN-PPS/CPU_PPS_PA\imagenes/filename.png'
p14_cursor = pg.locateCenterOnScreen(im14, confidence=0.8)
pg.moveTo(p14_cursor)
pg.move(45,0)
pg.click()
pg.typewrite('reporte_cpu_pps', interval=0.02)
pg.press('enter')

#Damos un tiempo para que se guarde el reporte y luego continamos con la ejecucion del programa
time.sleep(3)

#------------Llamamos al script que nos hace el calcluo del CPU--------------------------------------------------------------------------------

FILE_REPORTE = r'D:\Soporte_Core_CA\VPN_NGIN_PPS_HUAWEI__TRINING_EL SALVADOR_EL MAS NUEVO\PPS-NGIN-VPN-DOCUMENTACION\I2000_EJECUTABLE_PORTABLE\I2000\client\reporte_cpu_pps.html'

#Si el archivo reporte se obtuvo y se guardo en el directorio client entonces llamamos al programa que hace el calculo del CUP
if os.path.isfile(FILE_REPORTE):
    Popen(r'start D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\CPU_PPS_PA\calcular_cpu.py', shell=True)
else:
    print(f'El archivo de reporte de CPU PPS no se encuentra en el path:\n{FILE_REPORTE}')

#-----------------------------------------------------------------------------------------------------------------------------------------------
