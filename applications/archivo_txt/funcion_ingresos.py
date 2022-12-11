
def registros_500(datos):
    
    ''' Con esta funcion podemos separar un dataset con mas de 500 registros en varios datasets separados por una lista de datos 
        Este Desarrollo se implementa para la carga de informacion en el sistema SIIGO que solo permite importar informacion de a 500 
        registros'''
       
    lista_datos = []
    init = 0
    end = 498
    rango = round((len(datos)/498) + 0.5)

   
    
    for i in range(rango):
        lista_datos.append([])
                
        for j in range(1):          
                lista_datos[i] = datos[init:end].reset_index()

        init = end
        end += 498


    return lista_datos 