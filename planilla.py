from openpyxl import Workbook
import datetime, csv

#Nota: para acceder a las celdas, primero toca crearlas, 
#cuando una hoja de trabajo es creada no contiene celdas.
# Formatear la fecha y la hora para el nombre del archivo
year=datetime.date.today().strftime("%Y")
mes=datetime.date.today().strftime("%B")
libro = Workbook() #instanciar un libro de excel
hoja = libro.get_active_sheet() #activar una hoja de calculo
hoja.title ="Disponibilidad"

Nombre = hoja.cell("A1")
Nombre.value="Host"

Apellidos = hoja.cell("B1")
Apellidos.value="IdServicio"

Apellidos = hoja.cell("B1")
Apellidos.value="IdVivienda"

Apellidos = hoja.cell("B2")
Apellidos.value="RDB"
#celda.value = "hola hola"


# Cargar archivo CSV
reader = csv.reader(open('prueba.csv','rb'),delimiter=",")
for index,row in enumerate(reader):

	Host = hoja.cell("A"+str(index+2))
	Host.value=row[0]

	IdServicio=hoja.cell("B"+str(index+2))
	Apellidos.value=row[1]

	Apellidos=hoja.cell("B"+str(index+2))
	Apellidos.value=row[1]
	#print 'personas'+str(index+1)
 	#print 'Nombre: ' + row[0] + ', Apellidos: ' + row[1] + ', Edad: ' + row[2] + '\r'

#hoja2 = libro.create_sheet(0) #Insertar una segunda hoja
#hoja = libro.create_sheet(title="Nueva hoja")
#Guardar la hoja de calculo
libro.save(mes+'_'+year+'.xlsx') 