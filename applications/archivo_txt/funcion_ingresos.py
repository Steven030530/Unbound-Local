
def registros_500(datos):
    
    ''' Con esta funcion podemos separar un dataset con mas de 500 registros en varios datasets separados por una lista de datos 
        Este Desarrollo se implementa para la carga de informacion en el sistema SIIGO que solo permite importar informacion de a 500 
        registros'''

    datos1 = []
    datos2 = []
    datos3 = []
    datos4 = []
    datos5 = []
    datos6 = []
    datos7 = []
    datos8 = []
    datos9 = []
    datos10 = []
    datos11 = []
    datos12 = []
    datos13 = []
    datos14 = []
    datos15 = []


    if len(datos) <= 497:
        datos1 = datos[0:497].reset_index()

    elif len(datos) <= 990:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()

    elif len(datos) <= 1480:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()

    elif len(datos) <= 1975:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()

    elif len(datos) <= 2470:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()

    elif len(datos) <= 2965:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()
        datos6 = datos[2470:2965].reset_index()

    elif len(datos) <= 3460:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()
        datos6 = datos[2470:2965].reset_index()
        datos7 = datos[2965:3460].reset_index() 

    elif len(datos) <= 3955:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()
        datos6 = datos[2470:2965].reset_index()
        datos7 = datos[2965:3460].reset_index() 
        datos8 = datos[3460:3955].reset_index()

    elif len(datos) <= 4450:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()
        datos6 = datos[2470:2965].reset_index()
        datos7 = datos[2965:3460].reset_index() 
        datos8 = datos[3460:3955].reset_index()
        datos9 = datos[3955:4450].reset_index()

    elif len(datos) <= 4945:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()
        datos6 = datos[2470:2965].reset_index()
        datos7 = datos[2965:3460].reset_index() 
        datos8 = datos[3460:3955].reset_index()
        datos9 = datos[3955:4450].reset_index()
        datos10 = datos[4450:4945].reset_index()

    elif len(datos) <= 5440:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()
        datos6 = datos[2470:2965].reset_index()
        datos7 = datos[2965:3460].reset_index() 
        datos8 = datos[3460:3955].reset_index()
        datos9 = datos[3955:4450].reset_index()
        datos10 = datos[4450:4945].reset_index()
        datos11 = datos[4945:5440].reset_index()


    elif len(datos) <= 5935:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()
        datos6 = datos[2470:2965].reset_index()
        datos7 = datos[2965:3460].reset_index() 
        datos8 = datos[3460:3955].reset_index()
        datos9 = datos[3955:4450].reset_index()
        datos10 = datos[4450:4945].reset_index()
        datos11 = datos[4945:5440].reset_index()
        datos12 = datos[5440:5935].reset_index()

    elif len(datos) <= 6430:
        datos1 = datos[0:497].reset_index()
        datos2 = datos[497:990].reset_index()
        datos3 = datos[990:1480].reset_index()
        datos4 = datos[1480:1975].reset_index()
        datos5 = datos[1975:2470].reset_index()
        datos6 = datos[2470:2965].reset_index()
        datos7 = datos[2965:3460].reset_index() 
        datos8 = datos[3460:3955].reset_index()
        datos9 = datos[3955:4450].reset_index()
        datos10 = datos[4450:4945].reset_index()
        datos11 = datos[4945:5440].reset_index()
        datos12 = datos[5440:5935].reset_index()
        datos13 = datos[5935:6435].reset_index()

    if len(datos13) > 0: 
        lista_datos = list([datos1,datos2,datos3,datos4,datos5,datos6,datos7,datos8,datos9,datos10,datos11,datos12,datos13])
    elif len(datos12) > 0: 
        lista_datos = list([datos1,datos2,datos3,datos4,datos5,datos6,datos7,datos8,datos9,datos10,datos11,datos12]) 
    elif len(datos11) > 0:
        lista_datos = list([datos1,datos2,datos3,datos4,datos5,datos6,datos7,datos8,datos9,datos10,datos11])
    elif len(datos10) > 0:
        lista_datos = list([datos1,datos2,datos3,datos4,datos5,datos6,datos7,datos8,datos9,datos10])
    elif len(datos9) > 0:
        lista_datos = list([datos1,datos2,datos3,datos4,datos5,datos6,datos7,datos8,datos9])
    elif len(datos8) > 0:
        lista_datos = list([datos1,datos2,datos3,datos4,datos5,datos6,datos7,datos8])
    elif len(datos7) > 0:
        lista_datos = list([datos1,datos2,datos3,datos4,datos5,datos6,datos7])
    elif len(datos6) > 0:
        lista_datos = list([datos1,datos2,datos3,datos4,datos5,datos6])
    elif len(datos5) > 0:
        lista_datos = list([datos1,datos2,datos3,datos4,datos5])
    elif len(datos4) > 0:
        lista_datos = list([datos1,datos2,datos3,datos4])
    elif len(datos3) > 0:
        lista_datos = list([datos1,datos2,datos3])
    elif len(datos2) > 0:
        lista_datos = list([datos1,datos2])
    elif len(datos1) > 0:
        lista_datos = list([datos1])

    return lista_datos