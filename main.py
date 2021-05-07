# Sofia Munera Medina
# Python
# POO con pandas
# Fuentes: https://bioinf.comav.upv.es/courses/linux/python/pandas.html#:~:text=Pandas%20proporciona%20herramientas%20que%20permiten,fusionar%20y%20unir%20datos 

# EJEMPLO 1

# Importamos las librerias que estaremos usando 

import pandas as pd
from pandas import ExcelWriter

# Creamos la clase BDD

class BDD:

  # Documentacion

  ''' Representa una base de datos (BDD) que contiene las las ordenes de diferentes productos de un supermercado '''

  # Vamos a crear una funcion de la clase BDD llamada l

  def pedidos(self):

    # Documentacion 

    ''' Lee la BDD contenida en Libro1, clasifica sus datos, los ordena y filtra. Guarda los resultados de este proceso en Copia1 '''

    # Primero mediante la funcion de pandas read_excel vamos a leer la informacion de Libro1 (archivo .xlsx) especificamente la hoja llamada Ventas 

    ventas=pd.read_excel('Libro1.xlsx',sheet_name='Ventas')

    # Vemos el tamaño del dataframe que acabamos de leer por consola usando la funcion de pandas llamada shape, obtenemos una tupla:  (filas, columnas)

    print('Tamaño Dataframe original: ', ventas.shape)

    # Ahora al dateframe obtenido y almacenado en ventas (posee 5 columnas) le vamos a agregar una nueva llamada 'Costo total' que va realizar la siguiente operacion entre columnas: Precio unitario*Orden - Descuento. La columna nueva se inserta al final 

    ventas['Costo total']=ventas['Precio unitario'] * ventas['Orden '] - ventas['Descuento']

    # Verificamos con shape que acabamos de agregar una columna a ventas 

    print('Tamaño Dataframe modificado: ', ventas.shape)

    # El dateframe modificado (con columna Costo Total) lo vamos a ordenar de menor a mayor segun los datos de la columna Precio unitario, esto se hace mediante la funcion sort_values

    ventas_orden=ventas.sort_values(by='Precio unitario',ascending=True)

    # Ahora mediante groupby vamos a agrupar las columnas #Orden y Costo total y vamos a agregar 

    cobros=ventas.groupby('#Orden')['Costo total'].agg([sum,max])
    

    orden_1234=ventas[ventas['#Orden']==1234]
    

    # Haciendo uso de la funcion ExcelWriter de pandas vamos a crear un nuevo archivo .xlsx llamado Copia1 

    nuevo=ExcelWriter('Copia1.xlsx')

    # Enviamos el dateframe que acabamos de crear (matriz1) a el archivo nuevo (Copia1) 

    ventas_orden.to_excel(nuevo,'Ventas',index=False)
    cobros.to_excel(nuevo,'Cobros',index=False)
    orden_1234.to_excel(nuevo,'Orden 1234',index=False)


    # Guardamos lo que acabamos de introducir al archivo nuevo 

    nuevo.save()





a=BDD()
a.pedidos()
