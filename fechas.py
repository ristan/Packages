from paquetes import myframe
import datetime

def fecha():
	detalle_vel_bajada = 'DETALLE_VEL_BAJADA_CTR'
	detalle_vel_subida = 'DETALLE_VEL_SUBIDA_CTR'
	detalle_Disponibilidad = 'DETALLE_DISPO_CTR'
	myframe.cls()

	year=datetime.date.today().strftime("%Y")
	mes=datetime.date.today().strftime("%B")
	dia=datetime.date.today().strftime("%A")
	fecha_actual = datetime.date.today().strftime("%Y%m%d")
	print mes,year

	detalle_vel_bajada=fecha_actual +'_'+ detalle_vel_bajada +'.1'
	detalle_vel_subida=fecha_actual +'_'+ detalle_vel_subida +'.1'
	detalle_Disponibilidad=fecha_actual +'_'+ detalle_Disponibilidad +'.1'

	print detalle_vel_bajada
	print detalle_vel_subida
	print detalle_Disponibilidad