######################################### DESARROLLO INGRESO ##############################################################

class Ingresos_Abila:
    
    def aporte_general(colectivo,fecha,consecutivo,path):
        ''' colectivo = NNJ-AM-V
            fecha = Formato (Mes/Dia/Año)
            consecutivo = Numero
            path = Ubicacion del Archivo Base de Ingreso '''
        
        import pandas as pd
        import numpy as np
        
        
        dict_calendar={"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO",
                    "08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}

        # Incluimos la informacion necesaria para leer los datos relevantes:

        if colectivo == "NNJ": 
            list_subp = ["COAN-FCC","COAN-FG","COAN-FJ","COAN-MAE","COAN-MAI","COAN-MT","COAN-MV","COAN-NH", 
                        "COAN-NJ","COAN-PN"]

        elif colectivo == "AM":
            list_subp = ["COAN-AC","COAN-FZ","COAN-MAE","COAN-MAI","COAN-NJ","COAN-VP","COAN-HD"]

        elif colectivo == "V":
            list_subp = ["COAN-CS","COAN-SC"]
        
        
        fecha = fecha
        fecha = pd.to_datetime(pd.Series(fecha))
        fecha = fecha.dt.strftime("%m/%d/%Y")[0]
        consecutivo = consecutivo
        archivo = path
        sheet = colectivo
        data = pd.read_excel(archivo,sheet_name=sheet)
        data["SUBP"] = data["SUBP"].str.strip()

        try:
            # Creamos el archivo para importar en abila
            abila_ingreso = pd.DataFrame()
            abila_ingreso["XXXXXXXX"] = data["SUBP"] + "000" + consecutivo
            abila_ingreso["SESSION"] = "CR-COAN-99-" + dict_calendar.get(fecha[0:2])+fecha[6:] +"-DC"
            abila_ingreso["DESCRIPTION"] = "INGRESO RECIBIDO DE UNBOUND KANSAS"
            abila_ingreso["DATE"] = fecha
            abila_ingreso["DOCUMENT"] = data["SUBP"] + "000" + consecutivo
            abila_ingreso["DESCRIPTION DOC"] = "APORTE DEL MES DE " + dict_calendar.get(fecha[0:2])
            abila_ingreso["DATE2"] = fecha
            abila_ingreso["SUBP"] = "COAN-" + data["SUBP"]
            if colectivo == "NNJ":
                abila_ingreso["FOUND"] = "10"
            elif colectivo == "AM":
                abila_ingreso["FOUND"] = "15"
            elif colectivo == "V":
                abila_ingreso["FOUND"] = "19"
            abila_ingreso["GL"] = "20110"
            abila_ingreso["DPTO"] = "0"
            abila_ingreso["CH"] = data["CH"]
            abila_ingreso["DEBIT"] = ""
            abila_ingreso["CREDIT"] = data["Aporte Mes Actual"]
            abila_ingreso.drop(columns="XXXXXXXX",inplace=True)
            abila_ingreso.drop(abila_ingreso.loc[abila_ingreso['CREDIT']==0].index, inplace=True) 
            abila_ingreso.reset_index(drop=True,inplace=True)

            for i in list_subp:
                banco_subp = abila_ingreso.loc[abila_ingreso['SUBP'] == i]
                banco = [abila_ingreso["SESSION"][0],abila_ingreso["DESCRIPTION"][0],fecha,
                        i[5:] + "000" + consecutivo,abila_ingreso["DESCRIPTION DOC"][0],fecha,
                        i,abila_ingreso["FOUND"][0],"10105","0","",sum(banco_subp["CREDIT"]),0]
                abila_ingreso.loc[len(abila_ingreso.index)] = banco


            if colectivo == "AM":
                abila_ingreso["DOCUMENT"] = abila_ingreso["DOCUMENT"].replace(["MAI000"+ consecutivo,
                "MAE000" + consecutivo,"NJ000" + consecutivo],["MAIA000" + consecutivo,"MAEA000" + consecutivo,"NJA000"
                + consecutivo])
                

            # organizamos los datos para que queden en orden de subproyecto

            abila_ingreso.sort_values(["SUBP","FOUND"],inplace=True)
            
            #Eliminamos registros innecesarios
            
            abila_ingreso.drop(abila_ingreso.loc[(abila_ingreso['CREDIT']==0) & (abila_ingreso['DEBIT']==0) ].index,
            inplace=True)

            # Generamos el Archivo de Importacion
        
            abila_ingreso.to_excel("INGRESO APORTE GENERAL " + sheet + ".xlsx",index=False)

            print("\n*************** FELICITACIONES SE GENERO CON EXITO ***************\n")
        except:
            print("No hay datos")


    def aporte_cumpleanios(colectivo,fecha,consecutivo,path):
        
        ''' colectivo = NNJ-AM-V
            fecha = Formato (Mes/Dia/Año)
            consecutivo = Numero
            path = Ubicacion del Archivo Base de Ingreso '''        
        

        import pandas as pd
        import numpy as np

        
        dict_calendar={"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO",
                    "08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}

        # Incluimos la informacion necesaria para leer los datos relevantes:

        if colectivo == "NNJ": 
            list_subp = ["COAN-FCC","COAN-FG","COAN-FJ","COAN-MAE","COAN-MAI","COAN-MT","COAN-MV","COAN-NH", 
                        "COAN-NJ","COAN-PN"]

        elif colectivo == "AM":
            list_subp = ["COAN-AC","COAN-FZ","COAN-MAE","COAN-MAI","COAN-NJ","COAN-VP","COAN-HD"]

        elif colectivo == "V":
            list_subp = ["COAN-CS","COAN-SC"]

        fecha = fecha
        fecha = pd.to_datetime(pd.Series(fecha))
        fecha = fecha.dt.strftime("%m/%d/%Y")[0]
        consecutivo = consecutivo
        archivo = path
        sheet = colectivo
        data = pd.read_excel(archivo,sheet_name=sheet)
        data["SUBP"] = data["SUBP"].str.strip()

        try:
            # Creamos el archivo para importar en abila
            abila_ingreso = pd.DataFrame()
            abila_ingreso["XXXXXXXX"] = data["SUBP"] + "000" + consecutivo
            abila_ingreso["SESSION"] = "CR-COAN-99-" + dict_calendar.get(fecha[0:2]) + fecha[6:] +"-DC"
            abila_ingreso["DESCRIPTION"] = "INGRESO RECIBIDO DE UNBOUND KANSAS"
            abila_ingreso["DATE"] = fecha
            abila_ingreso["DOCUMENT"] = data["SUBP"] + "000" + consecutivo
            abila_ingreso["DESCRIPTION DOC"] = "APORTE CUMPLEANIOS DEL MES DE " + dict_calendar.get(fecha[0:2])
            abila_ingreso["DATE2"] = fecha
            abila_ingreso["SUBP"] = "COAN-" + data["SUBP"]
            abila_ingreso["FOUND"] = "20"
            abila_ingreso["GL"] = "20110"
            abila_ingreso["DPTO"] = "0"
            abila_ingreso["CH"] = data["CH"]
            abila_ingreso["DEBIT"] = ""
            abila_ingreso["CREDIT"] = data["Cumpleaños"]
            abila_ingreso.drop(columns="XXXXXXXX",inplace=True)
            abila_ingreso.drop(abila_ingreso.loc[abila_ingreso['CREDIT']==0].index, inplace=True) 
            abila_ingreso.reset_index(drop=True,inplace=True)

            for i in list_subp:
                banco_subp = abila_ingreso.loc[abila_ingreso['SUBP'] == i]
                banco = [abila_ingreso["SESSION"][0],abila_ingreso["DESCRIPTION"][0],fecha,
                        i[5:] + "000" + consecutivo,abila_ingreso["DESCRIPTION DOC"][0],fecha,
                        i,abila_ingreso["FOUND"][0],"10105","0","",sum(banco_subp["CREDIT"]),0]
                abila_ingreso.loc[len(abila_ingreso.index)] = banco

            if colectivo == "AM":
                abila_ingreso["DOCUMENT"] = abila_ingreso["DOCUMENT"].replace(["MAI000"+ consecutivo,"MAE000"+
                consecutivo,"NJ000" + consecutivo],["MAIA000" + consecutivo,"MAEA000" + consecutivo,"NJA000" + consecutivo])

            # organizamos los datos para que queden en orden de subproyecto

            abila_ingreso.sort_values(["SUBP","FOUND"],inplace=True)
            
            # Eliminar registros innecesarios
            
            abila_ingreso.drop(abila_ingreso.loc[(abila_ingreso['CREDIT']==0) & (abila_ingreso['DEBIT']==0) ].index,
            inplace=True)

        # Generamos el Archivo de Importacion
        
            abila_ingreso.to_excel("INGRESO CUMPLEAÑOS " + sheet + ".xlsx",index=False)

            print("\n*************** FELICITACIONES SE GENERO CON EXITO ***************\n")

        except:
            print("No hay datos")


    def aporte_regalo(colectivo,fecha,consecutivo,path):
        
        ''' colectivo = NNJ-AM-V
            fecha = Formato (Mes/Dia/Año)
            consecutivo = Numero
            path = Ubicacion del Archivo Base de Ingreso '''

        import pandas as pd
        import numpy as np

        
        dict_calendar={"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO",
        "08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}

        # Incluimos la informacion necesaria para leer los datos relevantes:

        if colectivo == "NNJ": 
            list_subp = ["COAN-FCC","COAN-FG","COAN-FJ","COAN-MAE","COAN-MAI","COAN-MT","COAN-MV","COAN-NH", 
                        "COAN-NJ","COAN-PN"]

        elif colectivo == "AM":
            list_subp = ["COAN-AC","COAN-FZ","COAN-MAE","COAN-MAI","COAN-NJ","COAN-VP","COAN-HD"]

        elif colectivo == "V":
            list_subp = ["COAN-CS","COAN-SC"]

        fecha = fecha
        fecha = pd.to_datetime(pd.Series(fecha))
        fecha = fecha.dt.strftime("%m/%d/%Y")[0]
        consecutivo = consecutivo
        archivo = path
        sheet = colectivo
        data = pd.read_excel(archivo,sheet_name=sheet)
        data["SUBP"] = data["SUBP"].str.strip()
        
        try:
            # Creamos el archivo para importar en abila
            abila_ingreso = pd.DataFrame()
            abila_ingreso["XXXXXXXX"] = data["SUBP"] + "000" + consecutivo
            abila_ingreso["SESSION"] = "CR-COAN-99-" + dict_calendar.get(fecha[0:2]) + fecha[6:] +"-DC"
            abila_ingreso["DESCRIPTION"] = "INGRESO RECIBIDO DE UNBOUND KANSAS"
            abila_ingreso["DATE"] = fecha
            abila_ingreso["DOCUMENT"] = "RE000" + consecutivo
            abila_ingreso["DESCRIPTION DOC"] = "APORTE REGALO ESPECIAL DEL MES DE " + dict_calendar.get(fecha[0:2])
            abila_ingreso["DATE2"] = fecha
            abila_ingreso["SUBP"] = "COAN-" + data["SUBP"]
            abila_ingreso["FOUND"] = "51"
            abila_ingreso["GL"] = "20110"
            abila_ingreso["DPTO"] = "0"
            abila_ingreso["CH"] = data["CH"]
            abila_ingreso["DEBIT"] = ""
            abila_ingreso["CREDIT"] = data["Regalo Especial"]
            abila_ingreso.drop(columns="XXXXXXXX",inplace=True)
            abila_ingreso.drop(abila_ingreso.loc[abila_ingreso['CREDIT']==0].index, inplace=True) 
            abila_ingreso.reset_index(drop=True,inplace=True)

            for i in list_subp:
                banco_subp = abila_ingreso.loc[abila_ingreso['SUBP'] == i]
                banco = [abila_ingreso["SESSION"][0],abila_ingreso["DESCRIPTION"][0],fecha,
                        "RE000" + consecutivo,abila_ingreso["DESCRIPTION DOC"][0],fecha,
                        i,abila_ingreso["FOUND"][0],"10105","0","",sum(banco_subp["CREDIT"]),0]
                abila_ingreso.loc[len(abila_ingreso.index)] = banco



            # organizamos los datos para que queden en orden de subproyecto

            abila_ingreso.sort_values(["SUBP","FOUND"],inplace=True)
            
            # Eliminamos registros innecesarios
            
            abila_ingreso.drop(abila_ingreso.loc[(abila_ingreso['CREDIT']==0) & (abila_ingreso['DEBIT']==0) ].index,
            inplace=True)

            # Generamos el Archivo de Importacion
            
            abila_ingreso.to_excel("INGRESO REGALO ESPECIAL " + sheet + ".xlsx",index=False)

            print("\n*************** FELICITACIONES SE GENERO CON EXITO ***************\n")

        except:
            print("No hay datos")

######################################### DESARROLLO CONSOLIDACION ########################################################
        
        
class Consolidar:
    
    def consolidacion_archivos(tipo):
        
        ''' Con esta funcion vamos a consolidar todos los
        archivos generados en uno solo.
        
        tipo = INGRESO o EGRESO puede estar acompañado de 
        otro valor para diferenciar'''
        
        from os import listdir,path,remove
        import pandas as pd

        directorio = []
        archivos = []

        for i in listdir("."):
            c=path.splitext(i)
            if c[1] == ".xlsx":
                directorio.append(i)

        for d in directorio:
            a = pd.read_excel(d)
            archivos.append(a)
            consolidacion = pd.concat(archivos,ignore_index=False)

        consolidacion.to_excel("CONSOLIDACION " + tipo + ".xlsx",index=False)

        for d in directorio:
            remove(d)
        
        print("\n*************** FELICITACIONES SE GENERO CON EXITO ***************\n")

        
        

        
        
######################################### DESARROLLO EGRESO ###############################################################        
        
        
        
class Egreso_Abila:
    
    def egreso_general(colectivo,empresa,fecha,consecutivo,path):
        
        ''' colectivo = NNJ-AM-V
            empresa = BANCOLOMBIA o EFECTY
            fecha = Formato (Mes/Dia/Año)
            consecutivo = Numero
            path = Ubicacion del Archivo Unificado de Dispersion '''

        import pandas as pd
        import numpy as np

        dict_calendar={"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO",
         "08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}
        
        if empresa == "BANCOLOMBIA":
            tipo_entrega = "Cuenta Bancaria"
        
        elif empresa == "EFECTY":
            tipo_entrega = "Efecty"

        # Incluimos la informacion necesaria para leer los datos relevantes:

        if colectivo == "NNJ": 
            list_subp = ["COAN-FCC","COAN-FG","COAN-FJ","COAN-MAE","COAN-MAI","COAN-MT","COAN-MV","COAN-NH", 
                        "COAN-NJ","COAN-PN"]

        elif colectivo == "AM":
            list_subp = ["COAN-AC","COAN-FZ","COAN-MAE","COAN-MAI","COAN-NJ","COAN-VP","COAN-HD"]

        elif colectivo == "V":
            list_subp = ["COAN-CS","COAN-SC"]        
        

        fecha = fecha
        fecha = pd.to_datetime(pd.Series(fecha))
        fecha = fecha.dt.strftime("%m/%d/%Y")[0]
        consecutivo = consecutivo
        archivo = path
        sheet = colectivo
        data = pd.read_excel(archivo,sheet_name=sheet)
        data["SUBP"] = data["SUBP"].str.strip()
        
        
        data_ok = (data["Estado de Entrega"] == "OKA") & (data["Tipo Entrega Aporte"] == tipo_entrega)
        data = data[data_ok]

        try:
            # Creamos el archivo para importar en abila
            abila_egreso = pd.DataFrame()
            abila_egreso["XXXXXXXX"] = data["SUBP"] + "000" + consecutivo
            abila_egreso["SESSION"] = "CD-COAN-99-" + dict_calendar.get(fecha[0:2]) + fecha[6:] +"-DC"
            abila_egreso["DESCRIPTION"] = "DISPERSION MES DE " + dict_calendar.get(fecha[0:2]) 
            abila_egreso["DATE"] = fecha
            abila_egreso["DOCUMENT"] = consecutivo
            abila_egreso["DESCRIPTION DOC"] = "APORTE GENERAL MES DE " + dict_calendar.get(fecha[0:2])
            abila_egreso["DATE2"] = fecha
            abila_egreso["EMPTY"] = ""
            abila_egreso["SUBP"] = "COAN-" + data["SUBP"]
            if colectivo == "NNJ":
                abila_egreso["FOUND"] = "10"
            elif colectivo == "AM":
                abila_egreso["FOUND"] = "15"
            elif colectivo == "V":
                abila_egreso["FOUND"] = "19"
                    
            abila_egreso["GL"] = "20110"
            abila_egreso["DPTO"] = "0"
            abila_egreso["CH"] = data["CH"]
            abila_egreso["DEBIT"] = data["Aporte Mes Actual"] + data["Solicita"]
            abila_egreso["CREDIT"] = ""
            abila_egreso.drop(columns="XXXXXXXX",inplace=True)
            abila_egreso.drop(abila_egreso.loc[abila_egreso['CREDIT']==0].index, inplace=True) 
            abila_egreso.reset_index(drop=True,inplace=True)

            for i in list_subp:
                banco_subp = abila_egreso.loc[abila_egreso['SUBP'] == i]
                banco = [abila_egreso["SESSION"][0],abila_egreso["DESCRIPTION"][0],fecha,
                        consecutivo,abila_egreso["DESCRIPTION DOC"][0],fecha,"",
                        i,abila_egreso["FOUND"][0],"10105","0","",0,sum(banco_subp["DEBIT"])]
                abila_egreso.loc[len(abila_egreso.index)] = banco


            # organizamos los datos para que queden en orden de subproyecto

            abila_egreso.sort_values(["SUBP","FOUND"],inplace=True)
            
            #Eliminamos registros innecesarios
            
            abila_egreso.drop(abila_egreso.loc[(abila_egreso['CREDIT']=="") & (abila_egreso['DEBIT']==0) ].index, inplace=True)
            abila_egreso.drop(abila_egreso.loc[(abila_egreso['CREDIT']==0) & (abila_egreso['DEBIT']==0) ].index, inplace=True)

            # Generamos el Archivo de Importacion

            abila_egreso.to_excel("/webapps/project/proyectounbound/media/DISPERSION APORTE " + empresa + " " + sheet + ".xlsx",index=False)

            print("\n*************** FELICITACIONES SE GENERO CON EXITO ***************\n")
        except:
            print("No hay datos")
    


    def egreso_cumple(colectivo,empresa,fecha,consecutivo,path):
        
        ''' colectivo = NNJ-AM-V
            empresa = BANCOLOMBIA o EFECTY
            fecha = Formato (Mes/Dia/Año)
            consecutivo = Numero
            path = Ubicacion del Archivo Unificado de Dispersion '''

        import pandas as pd
        import numpy as np

        dict_calendar={"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO",
         "08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}

        if empresa == "BANCOLOMBIA":
            tipo_entrega = "Cuenta Bancaria"

        elif empresa == "EFECTY":
            tipo_entrega = "Efecty"

        # Incluimos la informacion necesaria para leer los datos relevantes:

        if colectivo == "NNJ": 
            list_subp = ["COAN-FCC","COAN-FG","COAN-FJ","COAN-MAE","COAN-MAI","COAN-MT","COAN-MV","COAN-NH", 
                        "COAN-NJ","COAN-PN"]

        elif colectivo == "AM":
            list_subp = ["COAN-AC","COAN-FZ","COAN-MAE","COAN-MAI","COAN-NJ","COAN-VP","COAN-HD"]

        elif colectivo == "V":
            list_subp = ["COAN-CS","COAN-SC"]
        
       
        fecha = fecha
        fecha = pd.to_datetime(pd.Series(fecha))
        fecha = fecha.dt.strftime("%m/%d/%Y")[0]
        consecutivo = consecutivo
        archivo = path
        sheet = colectivo
        data = pd.read_excel(archivo,sheet_name=sheet)
        data["SUBP"] = data["SUBP"].str.strip()

        data_ok = (data["Estado de Entrega"] == "OKA") & (data["Tipo Entrega Aporte"] == tipo_entrega)
        data = data[data_ok]

        try:

            # Creamos el archivo para importar en abila
            abila_egreso = pd.DataFrame()
            abila_egreso["XXXXXXXX"] = data["SUBP"] + "000" + consecutivo
            abila_egreso["SESSION"] = "CD-COAN-99-" + dict_calendar.get(fecha[0:2]) + fecha[6:] +"-DC"
            abila_egreso["DESCRIPTION"] = "DISPERSION MES DE " + dict_calendar.get(fecha[0:2])
            abila_egreso["DATE"] = fecha
            abila_egreso["DOCUMENT"] = consecutivo
            abila_egreso["DESCRIPTION DOC"] = "CUMPLEANIOS MES DE " + dict_calendar.get(fecha[0:2])
            abila_egreso["DATE2"] = fecha
            abila_egreso["EMPTY"] = ""
            abila_egreso["SUBP"] = "COAN-" + data["SUBP"]
            abila_egreso["FOUND"] = "20"
            abila_egreso["GL"] = "20110"
            abila_egreso["DPTO"] = "0"
            abila_egreso["CH"] = data["CH"]
            abila_egreso["DEBIT"] = data["Cumpleaños"]
            abila_egreso["CREDIT"] = ""
            abila_egreso.drop(columns="XXXXXXXX",inplace=True)
            abila_egreso.drop(abila_egreso.loc[abila_egreso['CREDIT']==0].index, inplace=True) 
            abila_egreso.reset_index(drop=True,inplace=True)

            for i in list_subp:
                banco_subp = abila_egreso.loc[abila_egreso['SUBP'] == i]
                banco = [abila_egreso["SESSION"][0],abila_egreso["DESCRIPTION"][0],fecha,
                        consecutivo,abila_egreso["DESCRIPTION DOC"][0],fecha,"",
                        i,abila_egreso["FOUND"][0],"10105","0","",0,sum(banco_subp["DEBIT"])]
                abila_egreso.loc[len(abila_egreso.index)] = banco


            # organizamos los datos para que queden en orden de subproyecto

            abila_egreso.sort_values(["SUBP","FOUND"],inplace=True)

            #Eliminamos registros innecesarios

            abila_egreso.drop(abila_egreso.loc[(abila_egreso['CREDIT']=="") & (abila_egreso['DEBIT']==0) ].index, inplace=True)

            abila_egreso.drop(abila_egreso.loc[(abila_egreso['CREDIT']==0) & (abila_egreso['DEBIT']==0) ].index, inplace=True)


            # Generamos el Archivo de Importacion

            abila_egreso.to_excel("/webapps/project/proyectounbound/media/DISPERSION CUMPLEANIOS " + empresa + " " + sheet + ".xlsx",index=False)

            print("\n*************** FELICITACIONES SE GENERO CON EXITO ***************\n")
        except:
            print("No hay datos")
            
            

            
            
            
    def egreso_regalo(colectivo,empresa,fecha,consecutivo,path):
        
        ''' colectivo = NNJ-AM-V
            empresa = BANCOLOMBIA o EFECTY
            fecha = Formato (Mes/Dia/Año)
            consecutivo = Numero
            path = Ubicacion del Archivo Unificado de Dispersion '''
            
        import pandas as pd
        import numpy as np

        dict_calendar={"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO",
         "08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}

        if empresa == "BANCOLOMBIA":
            tipo_entrega = "Cuenta Bancaria"

        elif empresa == "EFECTY":
            tipo_entrega = "Efecty"

        # Incluimos la informacion necesaria para leer los datos relevantes:

        if colectivo == "NNJ": 
            list_subp = ["COAN-FCC","COAN-FG","COAN-FJ","COAN-MAE","COAN-MAI","COAN-MT","COAN-MV","COAN-NH", 
                        "COAN-NJ","COAN-PN"]

        elif colectivo == "AM":
            list_subp = ["COAN-AC","COAN-FZ","COAN-MAE","COAN-MAI","COAN-NJ","COAN-VP","COAN-HD"]

        elif colectivo == "V":
            list_subp = ["COAN-CS","COAN-SC"]
            
        
        fecha = fecha
        fecha = pd.to_datetime(pd.Series(fecha))
        fecha = fecha.dt.strftime("%m/%d/%Y")[0]
        consecutivo = consecutivo
        archivo = path
        sheet = colectivo
        data = pd.read_excel(archivo,sheet_name=sheet)
        data["SUBP"] = data["SUBP"].str.strip()

        data_ok = (data["Estado de Entrega"] == "OKA") & (data["Tipo Entrega Aporte"] == tipo_entrega)
        data = data[data_ok]

        try:

            # Creamos el archivo para importar en abila
            abila_egreso = pd.DataFrame()
            abila_egreso["XXXXXXXX"] = data["SUBP"] + "000" + consecutivo
            abila_egreso["SESSION"] = "CD-COAN-99-" + dict_calendar.get(fecha[0:2]) + fecha[6:] +"-DC"
            abila_egreso["DESCRIPTION"] = "DISPERSION MES DE " + dict_calendar.get(fecha[0:2])
            abila_egreso["DATE"] = fecha
            abila_egreso["DOCUMENT"] = consecutivo
            abila_egreso["DESCRIPTION DOC"] = "REGALO ESPECIAL MES DE " + dict_calendar.get(fecha[0:2])
            abila_egreso["DATE2"] = fecha
            abila_egreso["EMPTY"] = ""
            abila_egreso["SUBP"] = "COAN-" + data["SUBP"]
            abila_egreso["FOUND"] = "51"
            abila_egreso["GL"] = "20110"
            abila_egreso["DPTO"] = "0"
            abila_egreso["CH"] = data["CH"]
            abila_egreso["DEBIT"] = data["Regalo Especial"]
            abila_egreso["CREDIT"] = ""
            abila_egreso.drop(columns="XXXXXXXX",inplace=True)
            abila_egreso.drop(abila_egreso.loc[abila_egreso['CREDIT']==0].index, inplace=True) 
            abila_egreso.reset_index(drop=True,inplace=True)

            for i in list_subp:
                banco_subp = abila_egreso.loc[abila_egreso['SUBP'] == i]
                banco = [abila_egreso["SESSION"][0],abila_egreso["DESCRIPTION"][0],fecha,
                        consecutivo,abila_egreso["DESCRIPTION DOC"][0],fecha,"",
                        i,abila_egreso["FOUND"][0],"10105","0","",0,sum(banco_subp["DEBIT"])]
                abila_egreso.loc[len(abila_egreso.index)] = banco


            # organizamos los datos para que queden en orden de subproyecto

            abila_egreso.sort_values(["SUBP","FOUND"],inplace=True)

            #Eliminamos registros innecesarios

            abila_egreso.drop(abila_egreso.loc[(abila_egreso['CREDIT']=="") & (abila_egreso['DEBIT']==0) ].index, inplace=True)

            abila_egreso.drop(abila_egreso.loc[(abila_egreso['CREDIT']==0) & (abila_egreso['DEBIT']==0) ].index, inplace=True)

            # Generamos el Archivo de Importacion

            abila_egreso.to_excel("/webapps/project/proyectounbound/media/DISPERSION REGALOS " + empresa + " " + sheet + ".xlsx",index=False)


            print("\n*************** FELICITACIONES SE GENERO CON EXITO ***************\n")
        except:
            print("No hay datos")
        
        
        
        
        
        
        
    def egreso_fomentado(colectivo,empresa,fecha,consecutivo,path):
        
        ''' colectivo = FC
            empresa = BANCOLOMBIA o EFECTY
            fecha = Formato (Mes/Dia/Año)
            consecutivo = Numero
            path = Ubicacion del Archivo Unificado de Dispersion '''
            
        import pandas as pd
        import numpy as np

        dict_calendar={"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO",
         "08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}

        if empresa == "BANCOLOMBIA":
            tipo_entrega = "Cuenta Bancaria"

        elif empresa == "EFECTY":
            tipo_entrega = "Efecty"

        # Incluimos la informacion necesaria para leer los datos relevantes:

        list_subp = ["COAN-FCC","COAN-FG","COAN-FJ","COAN-MAE","COAN-MAI","COAN-MT","COAN-MV","COAN-NH", 
                    "COAN-NJ","COAN-PN"]

           
        fecha = fecha
        fecha = pd.to_datetime(pd.Series(fecha))
        fecha = fecha.dt.strftime("%m/%d/%Y")[0]
        consecutivo = consecutivo
        archivo = path
        sheet = colectivo
        data = pd.read_excel(archivo,sheet_name=sheet)
        data["SUBP"] = data["SUBP"].str.strip()

        data_ok = (data["Estado de Entrega"] == "OKA") & (data["Tipo Entrega Aporte"] == tipo_entrega)
        data = data[data_ok]

        try:
            # Creamos el archivo para importar en abila
            abila_egreso = pd.DataFrame()
            abila_egreso["XXXXXXXX"] = data["SUBP"] + "000" + consecutivo
            abila_egreso["SESSION"] = "CD-COAN-99-" + dict_calendar.get(fecha[0:2]) + fecha[6:] +"-DC"
            abila_egreso["DESCRIPTION"] = "DISPERSION MES DE " + dict_calendar.get(fecha[0:2])
            abila_egreso["DATE"] = fecha
            abila_egreso["DOCUMENT"] = consecutivo
            abila_egreso["DESCRIPTION DOC"] = "FOMENTANDO CAPACIDADES MES DE " + dict_calendar.get(fecha[0:2])
            abila_egreso["DATE2"] = fecha
            abila_egreso["EMPTY"] = ""
            abila_egreso["SUBP"] = "COAN-" + data["SUBP"]
            abila_egreso["FOUND"] = "10"
            abila_egreso["GL"] = "20110"
            abila_egreso["DPTO"] = "0"
            abila_egreso["CH"] = data["CH"]
            abila_egreso["DEBIT"] = data["Aporte Mes Actual"] + data["Solicita"]
            abila_egreso["CREDIT"] = ""
            abila_egreso.drop(columns="XXXXXXXX",inplace=True)
            abila_egreso.drop(abila_egreso.loc[abila_egreso['CREDIT']==0].index, inplace=True) 
            abila_egreso.reset_index(drop=True,inplace=True)

            for i in list_subp:
                banco_subp = abila_egreso.loc[abila_egreso['SUBP'] == i]
                banco = [abila_egreso["SESSION"][0],abila_egreso["DESCRIPTION"][0],fecha,
                        consecutivo,abila_egreso["DESCRIPTION DOC"][0],fecha,"",
                        i,abila_egreso["FOUND"][0],"10105","0","",0,sum(banco_subp["DEBIT"])]
                abila_egreso.loc[len(abila_egreso.index)] = banco


            # organizamos los datos para que queden en orden de subproyecto

            abila_egreso.sort_values(["SUBP","FOUND"],inplace=True)

            #Eliminamos registros innecesarios

            abila_egreso.drop(abila_egreso.loc[(abila_egreso['CREDIT']=="") & (abila_egreso['DEBIT']==0) ].index, inplace=True)

            abila_egreso.drop(abila_egreso.loc[(abila_egreso['CREDIT']==0) & (abila_egreso['DEBIT']==0) ].index, inplace=True)

            # Generamos el Archivo de Importacion

            abila_egreso.to_excel("/webapps/project/proyectounbound/media/DISPERSION FOMENTANDO CAPACIDADES " + empresa + " " + sheet + ".xlsx",index=False)


            print("\n*************** FELICITACIONES SE GENERO CON EXITO ***************\n")
        except:
            print("No hay datos")