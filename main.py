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

    # Ahora mediante groupby vamos a agrupar las columnas #Orden y Costo total, luego vamos a agregar tres columnas; la primera contiene la  funcion sum (suma los costos totales asociados a #Orden), la segunda columna contiene la funcion max (Imprime el Costo total mas alto asociado a #Orden) finalmente está la columna min (Costo total mas bajo asociado a #Orden)

    cobros=ventas.groupby('#Orden')['Costo total'].agg([sum,max, min])

    # imprimimos los resultados de esta funcion groupby

    print('\nAgrupacion de columnas #Orden y Costo total\n')
    print(cobros)
    
    # Ahora vamos a filtrar la hoja ventas por la columna #Orden y seleccionaremos las filas o entradas que corresponden al numero de orden 1234

    orden_1234=ventas[ventas['#Orden']==1234]

    print('\nFiltrado Orden 1234: \n')
    print(orden_1234)

    # Haciendo uso de la funcion ExcelWriter de pandas vamos a crear un nuevo archivo .xlsx llamado Copia1 

    nuevo=ExcelWriter('Copia1.xlsx')

    # El nuevo archivo que acabamos de crear (Copia1) va a tener 3 hojas; 

    # Ventas(va a contener el Dateframe modificado 6x6 y ordenado segun precio unitario de menor a mayor ) 

    ventas_orden.to_excel(nuevo,'Ventas',index=False)

    # Cobros(contiene la suma de los pedidos realizados por cada #Orden, el precio mas alto y el mas bajo de la orden 3*3)

    cobros.to_excel(nuevo,'Cobros',index=False)

    # Orden 1234 contiene los productos pedidos por este numero de orden 

    orden_1234.to_excel(nuevo,'Orden 1234',index=False)

    # Guardamos lo que acabamos de introducir al archivo nuevo 

    nuevo.save()

# Llamamos la clase BDD

a=BDD()

# La funcion pedidos de BDD va a leer la informacion de Libro1, la analiza y modifica e imprime los resultados en el archivo nuevo Copia1

a.pedidos()

# FINALIZA EJEMPLO 1  


# EJEMPLO 2

# Creamos una clase llamada Gimnasio 

class Gimnasio:

  # Documentacion

  ''' Registra cedula, nombre, detalles de salud y pagos de los usuarios de un gimnasio '''

  # Creamos una funcion llamada clientes

  def clientes(self):

    # Documentacion 

    ''' Extrae elmentos de dos hojas de Libro2, las une y guarda la tabla resultante en Copia2 '''

    # Primero mediante la funcion de pandas read_excel vamos a leer la informacion de Libro2 (archivo .xlsx) especificamente la hoja llamada Registrados (corresponde a la cedula, nombre, edad y eps) 

    registrados=pd.read_excel('Libro2.xlsx',sheet_name='Registrados')
    
    # Tambien del mismo archivo Libro2 extraemos los datod de la hoja llamada plan

    plan=pd.read_excel('Libro2.xlsx',sheet_name='Plan')

    # Si verificamos el contenido de plan con (print(plan)) podemos ver que tiene una columna llamada vencimiento sin datos. Mediate fillna vamos a reemplazar los valores numero por un valor especifico (fecha del vencimiento del plan)

    plan_mod=plan.fillna('07/06/2021') 
    
    # Ahora mediante merge vamos a combinar las dos hojas de libro dos y las vamos a almacenar en combinacion. Para hacer esto debemos de tener una columna en comun en ambas hojas, en este caso es Cedula
    
    combinacion=registrados.merge(plan_mod,left_on='Cedula ',right_on='Cedula ')

    # Creamos una nueva Hoja llamada Copia dos mediante ExcelWriter

    nuevo=ExcelWriter('Copia2.xlsx')

    # Enviamos la combinacion de las dos hojas a este archivo nuevo, ademas mediante freeze_panes vamos a inmobilizar la primera fila (encabezados) y la primera columna (numeros de cedulas)

    combinacion.to_excel(nuevo,'Usuarios',index=False,freeze_panes=(1,1))

    # Guarda los datos ingresados al archivo nuevo 

    nuevo.save()

  # definimos otra funcion de la clase Gimansio 

  def salud(self):

    # Documentacion 

    ''' Analiza el estado de salud de los usuarios con respecto a su IMC, extrae los datos de Libro2 y los almacena en una hoja nueva del archivo Copia2 '''

    # Leemos la hoja IMC del archivo Libro2 

    imc=pd.read_excel('Libro2.xlsx',sheet_name='IMC')
      
    # Calculamos el IMC de cada usuario mediante la formula peso*altura**2 

    IMC=imc['Peso']/imc['Altura']**2

    # Insertamos el calculo del IMC (indice de masa corporal) en una nueva columna al final del datedrame imc leido anteriormente

    imc.insert(3,'IMC',IMC,True)

    # Ahora analizaremos los resultados obtenidos del IMC de los usuarios 

    condicion=[]

    # Creamos un ciclo for que itera con i = 0,1,2 ya que tenemos 3 usuarios en la base de datos del gimnasio hoja imc

    for i in range(0,3):

      # Si el imc calculados se encuentra en este rango diremos que el peso es saludable 

      if 18.5<=IMC[i]<=24.9:

        a='Peso saludable'

        # Agregamos el resultado a la lista condicion 

        condicion.append(a)

      # Si el peso no se encuentra en este rango este por encima o por debajo diremos que no es saludanle 
      
      else:

        a='Peso NO saludable'

        # Agregamos el resultado a la lista condicion 

        condicion.append(a)
    
    # Ahora mediante insert agregamos al final de la tabla la condicion de cada usuario segun su imc 

    imc.insert(4,'Condicion',condicion,True)
    
    # Usando rename vamos a cambiar el nombre de dos columnas (Peso y Altura) con el fin de ser mas especifico con respecto a los datos que se muestran (diremos las unidades de los datos)

    imc_mod=imc.rename(columns={'Peso':'Peso (kg)', 'Altura':'Estatura (m)'})

    # Creamos una tabla dinamica mediante .pivot_table con el indice en la columna Cedula que nos permitirá filtrar por la cedula seleccionada y ver solo los datos correspondinetes a esta (la tabla contiene las columnas Peso, estatura, imc y condicion)

    tabla_dinamica=imc_mod.pivot_table(index='Cedula ', values=['Peso (kg)','Estatura (m)','IMC','Condicion']) 
    print(tabla_dinamica)

    # Creamos una archivo nuevo llamado copia3

    nuevo=ExcelWriter('Copia3.xlsx')

    tabla_dinamica.to_excel(nuevo,'Salud',index=False)

    nuevo.save()


  

a=Gimnasio()
a.clientes()
a.salud()
    


