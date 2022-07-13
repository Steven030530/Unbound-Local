def archivo_txt(name,fecha):

    from django.shortcuts import render
    from requests import request
    from django.shortcuts import redirect
    import pandas as pd
    import numpy as np
    import warnings
    warnings.simplefilter("ignore")
    from os import remove

    # SELECCIONAMOS LA HOJA DEL EXCEL QUE VAMOS A TRABAJAR Y LA FECHA DEL ARCHIVO TXT

    
    nombre_archivo = name
    fecha = fecha
    hojas = pd.read_excel(nombre_archivo,sheet_name=None)
    lista_sheets = list(hojas.keys())
    lista_sheets.remove("LISTAS")
    lista_archivos = []

    for i in lista_sheets:

        data = pd.read_excel(nombre_archivo,sheet_name=i)



    # ORGANIZAMOS EL ARCHIVO DE LOS DATOS PARA EXTRAER LOS VALORES QUE EFECTIVAMENTE SE ENTREGARON Y QUE SEAN DE EFECTY, ADEMAS CREAMOS UN VALOR
    # GENERALIZADO DE LOS VALORES RECIBIDOS POR LOS ACUDIENTES POR CADA UNO DE LOS FONDOS
    #ok = data["Estado de Entrega"] == "OKA"
        efecty = data["Tipo Entrega Aporte"] == "Efecty"
        data_ok = data[efecty]
        data_ok["Valor Dispersion"] = data_ok["Aporte Mes Actual"] + data_ok["Cumpleaños"] + data_ok["Regalo Especial"] + data_ok["Solicita"]
        #LLENAMOS ESPACIOS VACIOS CON ESPACIO
        data_ok["Otros Nombre/s Acudiente"].fillna(value=" ",inplace=True)
        

    # CREAMOS UN DATAFRAME CON LOS REQUERIMIENTOS ESTABLECIDOS POR EFECTY 

        efecty = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Modelo TXT Efecty.xlsx")

        efecty["Documento"] = data_ok["CC Acudiente"]
        efecty["co"] = '"'
        efecty["02"] = "02"
        efecty["cpc"] = '"|"'
        efecty["cpc1"] = '"|"'
        efecty["Tipo Documento"] = data_ok["Tipo de Documento"]
        efecty["cp"] = '"|'
        efecty.drop_duplicates(["Documento"],inplace=True,keep="last")
        efecty["Valor"] = data_ok.groupby("CC Acudiente").cumsum()[["Valor Dispersion"]]
        efecty["p"] = "|"
        efecty["Fecha y Hora"] = fecha + " 00:00:00"
        efecty["pc"] = '|"'
        efecty["Nombres"] = data_ok["Primer nombre Acudiente"] + " " + data_ok["Otros Nombre/s Acudiente"]
        efecty["cpc2"] = '"|"'
        efecty["Apellido1"] = data_ok["Primer apellido Acudiente"]
        efecty["cpc3"] = '"|"'
        efecty["Apellido2"] = data_ok["Segundo apellido Acudiente"]
        efecty["cpc4"] = '"|"'
        efecty["Telefono"] = data_ok["Telefono de Contacto"]
        efecty["cpc5"] = '"|"'
        efecty["Comentarios"] = data_ok["SUBP"]
        efecty["cpc6"] = '"|"'
        efecty["Codigo PS"] = "060410"
        efecty["cpc7"] = '"|"'
        efecty["PIN"] = "N.A"
        efecty["co1"] = '"'
        
    # CAMBIAMOS EL TIPO DE DOCUMENTO POR EL REQUERIDO POR EFECTY ADEMAS CAMBIAMOS LOS CAMPOS VACIOS EN EL APELLIDO POR NO APLICA
        efecty.replace(to_replace="Cedula de Ciudadania",value="CC",inplace=True,regex=True)
        efecty.replace(to_replace="Cedula Venezolana",value="CV",inplace=True,regex=True)
        efecty.replace(to_replace="Cedula Extranjeria",value="CE",inplace=True,regex=True)
        efecty["Apellido2"].replace(to_replace= np.nan ,value="N/A",inplace=True)
        
    # GENERAMOS EL ARCHIVO DE EXCEL PARA SUBIR A LA PLATAFORMA DE EFECTY
        lista_archivos.append(efecty)  
    
    consol = pd.concat(lista_archivos)
    consol.to_excel("archivos_txt\EfectyTXTConsolidado.xlsx",index=False)

    ################################## BANCOLOMBIA  ##########################################################################

# ORGANIZAMOS EL ARCHIVO DE LOS DATOS PARA EXTRAER LOS VALORES QUE EFECTIVAMENTE SE ENTREGARON Y QUE SEAN DE BANCOLOMBIA, ADEMAS CREAMOS UN VALOR
# GENERALIZADO DE LOS VALORES RECIBIDOS POR LOS ACUDIENTES POR CADA UNO DE LOS FONDOS, LIMPIAMOS LOS DATOS Y ADEMAS UNIFICAMOS LOS NOMBRES Y APELLIDOS DEL ACUDIENTE
    lista_archivos_banco = []
    for i in lista_sheets:
        dato = pd.read_excel(nombre_archivo,sheet_name=i)
        #ok = dato["Estado de Entrega"] == "OKA"
        bancolombia = dato["Tipo Entrega Aporte"] == "Cuenta Bancaria"
        dato_ok = dato[bancolombia]
        dato_ok["Valor Dispercion"] = dato_ok["Aporte Mes Actual"] + dato_ok["Cumpleaños"] + dato_ok["Regalo Especial"] + dato_ok["Solicita"]
        dato_ok["Otros Nombre/s Acudiente"].fillna(value=" ",inplace=True)
        dato_ok["Segundo apellido Acudiente"].fillna(value=" ",inplace=True)
        dato_ok["Nombre Del Acudiente (nombres - apellidos)"] = dato_ok["Primer nombre Acudiente"] + " " + dato_ok["Otros Nombre/s Acudiente"] + " " + dato_ok["Primer apellido Acudiente"] + " " + dato_ok["Segundo apellido Acudiente"]
        dato_ok.reset_index(inplace=True)
        
        
    # ESTABLECEMOS EL LIMITE DEL NOMBRE DEL ACUDIENTE A 30 CARACTERES

        for j in range(len(dato_ok["Nombre Del Acudiente (nombres - apellidos)"])):
            if len(dato_ok["Nombre Del Acudiente (nombres - apellidos)"][j])>=30:
                dato_ok["Nombre Del Acudiente (nombres - apellidos)"][j] = dato_ok["Nombre Del Acudiente (nombres - apellidos)"][j][0:30]
            else:
                dato_ok["Nombre Del Acudiente (nombres - apellidos)"][j]
                
    # CREAMOS UN NUEVO DATAFRAME CON LOS REQUERIMIENTO QUE EXIGE LA PLATAFORMA DE BANCOLOMBIA PARA SUBIR LOS DATOS DE PAGO
        nueva_tabla = pd.DataFrame()
        nueva_tabla["Tipo de Documento"] = dato_ok["Tipo de Documento"]
        nueva_tabla["Identificacion"] = dato_ok["CC Acudiente"]
        nueva_tabla["Nombre"] = dato_ok["Nombre Del Acudiente (nombres - apellidos)"]
        nueva_tabla["Tipo de Transaccion"] = 37
        nueva_tabla["Codigo Bancolombia"] = 1007
        nueva_tabla["Numero de la Cuenta"] = dato_ok["Numero Cuenta Bancaria"]
        nueva_tabla["Email"] = " "
        nueva_tabla["Documento"] = " "
        nueva_tabla["Referencia"] = dato_ok["SUBP"]
        nueva_tabla.drop_duplicates(["Identificacion"],inplace=True,keep="last")
        nueva_tabla["Oficina de entrega"] = " "
        #nueva_tabla["Tipo de Documento"] = 1
        nueva_tabla["Valor Transacciones"] = dato_ok.groupby("CC Acudiente").cumsum()[["Valor Dispercion"]]
        nueva_tabla.replace(to_replace="ñ",value="n",inplace=True,regex=True)
        nueva_tabla.replace(to_replace="á",value="a",inplace=True,regex=True)
        nueva_tabla.replace(to_replace="é",value="e",inplace=True,regex=True)
        nueva_tabla.replace(to_replace="í",value="i",inplace=True,regex=True)
        nueva_tabla.replace(to_replace="ó",value="o",inplace=True,regex=True)
        nueva_tabla.replace(to_replace="ú",value="u",inplace=True,regex=True)
        
    # CAMBIAMOS EL TIPO DE DOCUMENTO POR EL REQUERIDO POR EFECTY ADEMAS CAMBIAMOS LOS CAMPOS VACIOS EN EL APELLIDO POR NO APLICA
        nueva_tabla.replace(to_replace="Cedula de Ciudadania",value="1",inplace=True,regex=True)
        nueva_tabla.replace(to_replace="Pasaporte",value="5",inplace=True,regex=True)
        nueva_tabla.replace(to_replace="Cedula Extranjeria",value="2",inplace=True,regex=True)
        
    # GEBERAMOS EL ARCHIVO EXCEL PARA SUBIR A LA PLATAFORMA DE BANCOLOMBIA   
        lista_archivos_banco.append(nueva_tabla)
    consolban = pd.concat(lista_archivos_banco)
    consolban.to_excel("archivos_txt\BancolombiaTXTConsolidado.xlsx",index=False)

###########################################################################################################################################################


#################################################################################################################################################################################

def ingreso_41(name,date,consecutivo):
    import pandas as pd
    import warnings
    import calendar
    warnings.simplefilter("ignore")
    import datetime 
   

    dict_calendar = {"January":"ENERO","February":"FEBRERO","March":"MARZO","April":"ABRIL","May":"MAYO","June":"JUNIO","July":"JULIO","August":"AGOSTO",
                "September":"SEPTIEMBRE","October":"OCTUBRE","November":"NOVIEMBRE","December":"DICIEMBRE"}

    cuenta_ingreso = 41800102
    cuenta_pasivo = 28150505
    consecutivo = consecutivo
    fecha = pd.to_datetime(pd.Series(date))
    fecha = fecha.dt.strftime("%d/%m/%Y")[0]
    nombre_archivo = name
    monthinteger = int(fecha[4]) 
    month = datetime.date(1900, monthinteger, 1).strftime('%B')
    month = dict_calendar.get(month) 
    # ORGANIZAR TABLA DE CONSOLIDACION

    datos = pd.read_excel(nombre_archivo,sheet_name="Consolidacion",index_col="SUBPROYECTO")
    consolidacion = datos[1:29]
    consolidacion.drop(columns=["Unnamed: 1"],inplace=True)

    # SEPARAMOS LOS VALORES QUE VAN AL INGRESO Y LOS QUE VAN AL PASIVO

    ingreso = consolidacion.drop(index=["Refrigerios GU","Actividades Recreativas(Gimnasia)","Fomentando Capacidades"])
    pasivo = consolidacion.drop(index= ingreso.index )

    # CREAMOS UNA TABLA QUE CONTENGA LOS REGISTROS DEL INGRESO
    tabla1 =  pd.DataFrame(ingreso.sum(),columns=["Valor"])
    tabla1["Subproyecto"] = tabla1.index
    tabla1["Cuenta_Contable"] = cuenta_ingreso
    tabla1.reset_index(drop=True,inplace=True)


    #CREAMOS UNA TABLA QUE CONTENGA LOS REGISTROS DEL PASIVO
    tabla2 = pd.DataFrame(pasivo.sum(),columns=["Valor"])
    tabla2["Subproyecto"] = tabla2.index
    tabla2["Cuenta_Contable"] = cuenta_pasivo
    tabla2.reset_index(drop=True,inplace=True)

    # REGISTRAMOS LOS VALORES EN EL ARCHIVO MODELO PARA LA IMPORTACION 

    modelo = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Comprobante Modelo.xlsx")
    modelo["Código centro/subcentro de costos"] = tabla1["Subproyecto"]
    modelo["Código cuenta contable"] = cuenta_ingreso
    modelo["Consecutivo comprobante"] = consecutivo
    modelo["Identificación tercero"] = 2000
    modelo["Tipo de comprobante"] = 11
    modelo["Fecha de elaboración "]= fecha
    modelo["Crédito"] = tabla1["Valor"]
    modelo["Descripción"] = "INGRESO GENERAL" + " " + month
    banco = [11,modelo["Consecutivo comprobante"][1], modelo["Fecha de elaboración "][1],
            "","",11100511,890903938,"","","","","","","","","","","","",
            modelo["Descripción"][1],"ADMINISTRATIVO",sum(modelo["Crédito"]),"","","","",""]
    modelo.loc[len(modelo.index)] = banco
    modelo.reset_index(inplace=True)

    # ORGANIZAMOS LOS CENTROS DE COSTOS PARA LA IMPORTACION

    for i in range(len(modelo["Código centro/subcentro de costos"])):
        if len(modelo["Código centro/subcentro de costos"][i]) == 2:
            modelo["Código centro/subcentro de costos"][i] = "0000" + modelo["Código centro/subcentro de costos"][i] + "-1"
        elif len(modelo["Código centro/subcentro de costos"][i]) == 3 and modelo["Código centro/subcentro de costos"][i] == "MAI":
            modelo["Código centro/subcentro de costos"][i] = "00MAIN-1"
        elif len(modelo["Código centro/subcentro de costos"][i]) == 3 and modelo["Código centro/subcentro de costos"][i] == "MAE":
            modelo["Código centro/subcentro de costos"][i] = "00MAEN-1"
        
        elif modelo["Código centro/subcentro de costos"][i] == "ADMINISTRATIVO (Abuelos)":
            modelo["Código centro/subcentro de costos"][i] = "000100-1"
            
        elif  modelo["Código centro/subcentro de costos"][i] == "ADMINISTRATIVO":
            modelo["Código centro/subcentro de costos"][i] = "000100-1"
        
        elif modelo["Código centro/subcentro de costos"][i] == "MAE(Abuelos)":
            modelo["Código centro/subcentro de costos"][i] = "00MAEA-1"
            
        elif  modelo["Código centro/subcentro de costos"][i] == "MAI (Abuelos)":
            modelo["Código centro/subcentro de costos"][i] = "00MAIA-1"
        
        
        else:
            modelo["Código centro/subcentro de costos"][i] = "000" + modelo["Código centro/subcentro de costos"][i] + "-1"
            
            
    # EXPORTAMOS EL ARCHIVO DE EXCEL PARA IMPORTAR EN SIIGO
        
    modelo.drop(modelo.loc[modelo['Crédito']==0].index, inplace=True) 
    modelo.to_excel("Ingreso Cuenta 41" + ".xlsx",index=False)

#####################################################################################################################################################################


######################################################################################################################################################################

def ingreso_28(name,date,consecutivo):
    import pandas as pd
    import warnings
    import calendar
    from applications.archivo_txt import funcion_ingresos as FD
    import datetime 
    warnings.simplefilter("ignore")

    dict_calendar = {"January":"ENERO","February":"FEBRERO","March":"MARZO","April":"ABRIL","May":"MAYO","June":"JUNIO","July":"JULIO","August":"AGOSTO",
                "September":"SEPTIEMBRE","October":"OCTUBRE","November":"NOVIEMBRE","December":"DICIEMBRE"}

    consecutivo = int(consecutivo)
    fecha = pd.to_datetime(pd.Series(date))
    fecha = fecha.dt.strftime("%d/%m/%Y")[0]
    monthinteger = int(fecha[4]) 
    month = datetime.date(1900, monthinteger, 1).strftime('%B')
    month = dict_calendar.get(month) 
    nombre_archivo = name
    hojas = pd.read_excel(nombre_archivo,sheet_name=None)
    lista_sheets = list(hojas.keys())
    lista_sheets.remove("LISTAS")
    lista_sheets.remove("FC")

    for k in lista_sheets:

            hoja_actual = k
            print(k)

            datos = pd.read_excel(nombre_archivo,sheet_name=k)
            datos["SUBP"] = datos["SUBP"].str.strip() 

            try:
                lista_datos = FD.registros_500(datos)
            except:
                lista_datos = []


            if hoja_actual != "AM":

                for j in lista_datos:

                    for i in range(len(j["SUBP"])):
                        if len(j["SUBP"][i]) == 2:
                            j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                            j["SUBP"][i] = "00MAIN-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                            j["SUBP"][i] = "00MAEN-1"
                        else:
                            j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"

            elif hoja_actual == "AM":

                for j in lista_datos:

                    for i in range(len(j["SUBP"])):

                        if len(j["SUBP"][i]) == 2 and j["SUBP"][i] == "NJ":
                            j["SUBP"][i] = "000NJA-1"
                        elif len(j["SUBP"][i]) == 2:
                            j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                            j["SUBP"][i] = "00MAIA-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                            j["SUBP"][i] = "00MAEA-1"
                        else:
                            j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"

        
        
           
                
                
            # APADRINAMIENTO GENERAL


            numero_egreso = 1
            
            for y in lista_datos:
                aporte_general = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Comprobante Modelo.xlsx")
                aporte_general["Código centro/subcentro de costos"] = y["SUBP"]
                if k == "V":
                    aporte_general["Código cuenta contable"] = 28150520
                else:
                    aporte_general["Código cuenta contable"] = 28150505

                aporte_general["Consecutivo comprobante"] = consecutivo
                aporte_general["Identificación tercero"] = y["CC Acudiente"]
                aporte_general["Tipo de comprobante"] = 11
                aporte_general["Fecha de elaboración "]= fecha
                aporte_general["Crédito"] = y["Aporte Mes Actual"]
                aporte_general["Descripción"] = "INGRESO GENERAL APADRINAMIENTO" + " " + month
                banco = [11,aporte_general["Consecutivo comprobante"][1], aporte_general["Fecha de elaboración "][1],"","",11100511,890903938,"","","","","","","","","","","","",aporte_general["Descripción"][1],"000100-1",sum(aporte_general["Crédito"]),"","","","",""]
                aporte_general.loc[len(aporte_general.index)] = banco
                consecutivo = consecutivo + 1
                aporte_general.to_excel("INGRESO GENERAL AP " + hoja_actual  + str(numero_egreso) + ".xlsx",index=False)
                numero_egreso = numero_egreso + 1
                print("No Consecutivo: " + str(consecutivo))

            # APORTE CUMPLEAÑOS

            aporte_cumple = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Comprobante Modelo.xlsx")
            aporte_cumple["Código centro/subcentro de costos"] = datos["SUBP"]
            aporte_cumple["Código cuenta contable"] = 28150530
            aporte_cumple["Consecutivo comprobante"] = consecutivo
            aporte_cumple["Identificación tercero"] = datos["CC Acudiente"]
            aporte_cumple["Tipo de comprobante"] = 11
            aporte_cumple["Fecha de elaboración "]= fecha
            aporte_cumple["Crédito"] = datos["Cumpleaños"]
            aporte_cumple["Descripción"] = "INGRESO GENERAL CUMPLEAÑOS" + " " + month
            banco = [11,aporte_cumple["Consecutivo comprobante"][1], aporte_cumple["Fecha de elaboración "][1],"","",11100511,890903938,"","","","","","","","","","","","",aporte_cumple["Descripción"][1],"000100-1",sum(aporte_cumple["Crédito"]),"","","","",""]
            aporte_cumple.loc[len(aporte_cumple.index)] = banco
            aporte_cumple = aporte_cumple[aporte_cumple["Crédito"] != 0]

            consecutivo = consecutivo + 1
            numero_egreso_cumple = 1
            aporte_cumple.to_excel("INGRESO GENERAL CUMPLEAÑOS " + hoja_actual + str(numero_egreso_cumple) + ".xlsx",index=False)
            print("No Consecutivo: " + str(consecutivo))



            # APORTE REGALOS ESPECIALES              

            aporte_regalo = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Comprobante Modelo.xlsx")
            aporte_regalo["Código centro/subcentro de costos"] = datos["SUBP"]
            aporte_regalo["Código cuenta contable"] = 28150510
            aporte_regalo["Consecutivo comprobante"] = consecutivo
            aporte_regalo["Identificación tercero"] = datos["CC Acudiente"]
            aporte_regalo["Tipo de comprobante"] = 11
            aporte_regalo["Fecha de elaboración "]= fecha
            aporte_regalo["Crédito"] = datos["Regalo Especial"]
            aporte_regalo["Descripción"] = "INGRESO REGALOS ESPECIALES" + " " +dict_calendar.get(calendar.month_name[pd.to_datetime(aporte_regalo["Fecha de elaboración "][0]).month])
            banco = [11,aporte_regalo["Consecutivo comprobante"][1], aporte_regalo["Fecha de elaboración "][1],"","",11100511,890903938,"","","","","","","","","","","","",aporte_regalo["Descripción"][1],"000100-1",sum(aporte_regalo["Crédito"]),"","","","",""]
            aporte_regalo.loc[len(aporte_regalo.index)] = banco
            aporte_regalo = aporte_regalo[aporte_regalo["Crédito"] != 0]
            
            
            numero_egreso_re = 1
            aporte_regalo.to_excel("INGRESO GENERAL REGALO ESPECIAL " + hoja_actual + str(numero_egreso_re) + ".xlsx",index=False)
            print("No Consecutivo: " + str(consecutivo))
                    
###################################################################################################################################################################################################################################################################################################



###################################################################################################################################################################################################################################################################################################


def egreso_general_siigo(name,date,consecutivo,entrega):

    import pandas as pd
    import warnings
    import datetime 
    warnings.simplefilter("ignore")
    from applications.archivo_txt import funcion_ingresos as FD

    dict_calendar = {"January":"ENERO","February":"FEBRERO","March":"MARZO","April":"ABRIL","May":"MAYO","June":"JUNIO","July":"JULIO","August":"AGOSTO",
                "September":"SEPTIEMBRE","October":"OCTUBRE","November":"NOVIEMBRE","December":"DICIEMBRE"}

    nombre_archivo = name

    ## Establecemos las variables para ejecutar el bucle
    hojas = pd.read_excel(nombre_archivo,sheet_name=None)
    lista_sheets = list(hojas.keys())
    lista_sheets.remove("LISTAS")
    fecha = pd.to_datetime(pd.Series(date))
    fecha = fecha.dt.strftime("%d/%m/%Y")[0]
    monthinteger = int(fecha[4]) 
    month = datetime.date(1900, monthinteger, 1).strftime('%B')
    month = dict_calendar.get(month)
    comprobante = consecutivo
    print(entrega)

    for i in lista_sheets:
    
            # leemos el archivo con pandas
            datos_p = pd.read_excel(nombre_archivo,sheet_name=i)
            datos_p["SUBP"] = datos_p["SUBP"].str.strip()       

            datos_ok = (datos_p["Estado de Entrega"] == "OKA") & (datos_p["Tipo Entrega Aporte"] == entrega)
            datos = datos_p[datos_ok]


            numero_para_guardar_egreso = 1

            hoja_actual = i
            print(hoja_actual)

            #" Subdividimos los archivos de a 500 Registros"

            try:
                lista_datos = FD.registros_500(datos)
            except:
                lista_datos = []
                
            #"Organizamos los subproyectos para que sean leidos por el sistema siigo"

            if hoja_actual != "AM":

                for j in lista_datos:

                    for i in range(len(j["SUBP"])):
                        if len(j["SUBP"][i]) == 2:
                            j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                            j["SUBP"][i] = "00MAIN-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                            j["SUBP"][i] = "00MAEN-1"
                        else:
                            j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"

            elif hoja_actual == "AM":

                for j in lista_datos:

                    for i in range(len(j["SUBP"])):

                        if len(j["SUBP"][i]) == 2 and j["SUBP"][i] == "NJ":
                            j["SUBP"][i] = "000NJA-1"
                        elif len(j["SUBP"][i]) == 2:
                            j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                            j["SUBP"][i] = "00MAIA-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                            j["SUBP"][i] = "00MAEA-1"
                        else:
                            j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"

            # Generamos Modelo para Cargar al sistema contable            

            for j in lista_datos:

                print("Consecutivo Nro: " + str(comprobante))
                modelo = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Comprobante Modelo.xlsx")
                modelo["Identificación tercero"] = j["CC Acudiente"] 
                modelo["Tipo de comprobante"] = 1
                modelo["Consecutivo comprobante"] = comprobante
                modelo["Fecha de elaboración "] = fecha
                modelo["Débito"] = j["Aporte Mes Actual"]
                    
                if hoja_actual == "FC":
                    modelo["Código cuenta contable"] = 28150560
                
                elif hoja_actual == "FNC":
                    modelo["Código cuenta contable"] = 28150585

                elif hoja_actual == "V":
                    modelo["Código cuenta contable"] = 28150520
                
                else:
                    modelo["Código cuenta contable"] = 28150505
                
                modelo["Código centro/subcentro de costos"] = j["SUBP"]
                modelo["Descripción"] = ("APORTE " + month + " " + hoja_actual)
                if entrega == "Cuenta Bancaria":
                    banco = [1,modelo["Consecutivo comprobante"][0], modelo["Fecha de elaboración "][0],"","",11100511,890903938,"","","","","","","","","","","","",modelo["Descripción"][0],"000100-1","",sum(modelo["Débito"]),"","","",""]
                    modelo.loc[len(modelo.index)] = banco
                elif entrega == "Efecty":
                    gasto = [1,modelo["Consecutivo comprobante"][0], modelo["Fecha de elaboración "][0],"","",53051501,830131993,"","","","","","","","","","","","",modelo["Descripción"][0],"000100-1",sum(modelo["Débito"]*0.014),"","","","",""]
                    modelo.loc[len(modelo.index)] = gasto
                    banco = [1,modelo["Consecutivo comprobante"][0], modelo["Fecha de elaboración "][0],"","",11100511,890903938,"","","","","","","","","","","","",modelo["Descripción"][0],"000100-1","",sum(modelo["Débito"]),"","","",""]
                    modelo.loc[len(modelo.index)] = banco

            # le asignamos el siguiente valor al comprobante

                comprobante = int(comprobante) + 1

            #"Exportamos el Archivo a Excel"

                modelo.drop(modelo.loc[modelo['Débito']==0].index, inplace=True)
                if entrega == "Cuenta Bancaria":
                    modelo.to_excel("APORTE APADRINAMIENTO BANCOLOMBIA " + str(hoja_actual) + " " + str(numero_para_guardar_egreso) + ".xlsx",index=False)
                elif entrega == "Efecty":
                    modelo.to_excel("APORTE APADRINAMIENTO EFECTY " + str(hoja_actual) + " " + str(numero_para_guardar_egreso) + ".xlsx",index=False)
                numero_para_guardar_egreso = numero_para_guardar_egreso + 1


    ############################### CUMPLEAÑOS ###############################################################################3

    for i in lista_sheets:
    # leemos el archivo con pandas
        datos_p = pd.read_excel(nombre_archivo,sheet_name=i)
        datos_p["SUBP"] = datos_p["SUBP"].str.strip()           

        datos_ok = (datos_p["Estado de Entrega"] == "OKA") & (datos_p["Tipo Entrega Aporte"] == entrega)
        datos = datos_p[datos_ok]
        datos = datos[datos["Cumpleaños"] != 0]        


        numero_para_guardar_egreso = 1

        hoja_actual = i
        
        print(i)

        #" Subdividimos los archivos de a 500 Registros"
        try:
            lista_datos = FD.registros_500(datos)
        except:
            lista_datos = []
            
        #"Organizamos los subproyectos para que sean leidos por el sistema siigo"

        if hoja_actual != "AM":

            for j in lista_datos:

                for i in range(len(j["SUBP"])):
                    if len(j["SUBP"][i]) == 2:
                        j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                    elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                        j["SUBP"][i] = "00MAIN-1"
                    elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                        j["SUBP"][i] = "00MAEN-1"
                    else:
                        j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"

        elif hoja_actual == "AM":

            for j in lista_datos:

                for i in range(len(j["SUBP"])):

                    if len(j["SUBP"][i]) == 2 and j["SUBP"][i] == "NJ":
                        j["SUBP"][i] = "000NJA-1"
                    elif len(j["SUBP"][i]) == 2:
                        j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                    elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                        j["SUBP"][i] = "00MAIA-1"
                    elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                        j["SUBP"][i] = "00MAEA-1"
                    else:
                        j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"


        #"Alimentamos el archivo modelo para importar en contabilidad" 

        for j in lista_datos:

            print("Consecutivo Nro: " + str(comprobante))
            modelo = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Comprobante Modelo.xlsx")
            modelo["Identificación tercero"] = j["CC Acudiente"] 
            modelo["Tipo de comprobante"] = 1
            modelo["Consecutivo comprobante"] = comprobante
            modelo["Fecha de elaboración "] = fecha
            modelo["Débito"] = j["Cumpleaños"]
            modelo["Código cuenta contable"] = 28150530
            modelo["Código centro/subcentro de costos"] = j["SUBP"]
            modelo["Descripción"] = ("APORTE CUMPLEAÑOS " + month + " " + hoja_actual)
            banco = [1,modelo["Consecutivo comprobante"][0], modelo["Fecha de elaboración "][0],"","",11100511,890903938,"","","","","","","","","","","","",modelo["Descripción"][0],"000100-1","",sum(modelo["Débito"]),"","","",""]
            modelo.loc[len(modelo.index)] = banco

        # le asignamos el siguiente valor al comprobante

            comprobante = comprobante + 1

        #"Exportamos el Archivo a Excel"

            modelo.drop(modelo.loc[modelo['Débito']==0].index, inplace=True)
            if entrega == "Cuenta Bancaria": 
                modelo.to_excel("APORTE CUMPLEAÑOS BANCOLOMBIA " + str(hoja_actual) + " " + str(numero_para_guardar_egreso) + ".xlsx",index=False)
            elif entrega == "Efecty":
                modelo.to_excel("APORTE CUMPLEAÑOS EFECTY " + str(hoja_actual) + " " + str(numero_para_guardar_egreso) + ".xlsx",index=False)
            numero_para_guardar_egreso = numero_para_guardar_egreso + 1

        ################################### REGALOS ESPECIALES ##################################################################
    for i in lista_sheets:
        
            # leemos el archivo con pandas
            datos_p = pd.read_excel(nombre_archivo,sheet_name=i)
            datos_p["SUBP"] = datos_p["SUBP"].str.strip()           

            datos_ok = (datos_p["Estado de Entrega"] == "OKA") & (datos_p["Tipo Entrega Aporte"] == entrega)
            datos = datos_p[datos_ok]
            datos = datos[datos["Regalo Especial"] != 0]        


            numero_para_guardar_egreso = 1

            hoja_actual = i
            print(i)

            #" Subdividimos los archivos de a 500 Registros"

            try:
                lista_datos = FD.registros_500(datos)
            except:
                lista_datos = []
                
            #"Organizamos los subproyectos para que sean leidos por el sistema siigo"

            if hoja_actual != "AM":

                for j in lista_datos:

                    for i in range(len(j["SUBP"])):
                        if len(j["SUBP"][i]) == 2:
                            j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                            j["SUBP"][i] = "00MAIN-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                            j["SUBP"][i] = "00MAEN-1"
                        else:
                            j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"

            elif hoja_actual == "AM":

                for j in lista_datos:

                    for i in range(len(j["SUBP"])):

                        if len(j["SUBP"][i]) == 2 and j["SUBP"][i] == "NJ":
                            j["SUBP"][i] = "000NJA-1"
                        elif len(j["SUBP"][i]) == 2:
                            j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                            j["SUBP"][i] = "00MAIA-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                            j["SUBP"][i] = "00MAEA-1"
                        else:
                            j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"              

            #"Alimentamos el archivo modelo para importar en contabilidad" 

            for j in lista_datos:

                print("Consecutivo Nro: " + str(comprobante))
                modelo = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Comprobante Modelo.xlsx")
                modelo["Identificación tercero"] = j["CC Acudiente"] 
                modelo["Tipo de comprobante"] = 1
                modelo["Consecutivo comprobante"] = comprobante
                modelo["Fecha de elaboración "] = fecha
                modelo["Débito"] = j["Regalo Especial"]
                modelo["Código cuenta contable"] = 28150510
                modelo["Código centro/subcentro de costos"] = j["SUBP"]
                modelo["Descripción"] = ("APORTE REGALO ESPECIAL " + month + " "  + hoja_actual)
                banco = [1,modelo["Consecutivo comprobante"][0], modelo["Fecha de elaboración "][0],"","",11100511,890903938,"","","","","","","","","","","","",modelo["Descripción"][0],"000100-1","",sum(modelo["Débito"]),"","","",""]
                modelo.loc[len(modelo.index)] = banco           

            # le asignamos el siguiente valor al comprobante

                comprobante = comprobante + 1

            #"Exportamos el Archivo a Excel"

                modelo.drop(modelo.loc[modelo['Débito']==0].index, inplace=True)
                if entrega == "Cuenta Bancaria":
                    modelo.to_excel("APORTE REGALO ESPECIAL BANCOLOMBIA " + str(hoja_actual) + " " + str(numero_para_guardar_egreso) + ".xlsx",index=False)
                elif entrega == "Efecty":
                    modelo.to_excel("APORTE REGALO ESPECIAL EFECTY " + str(hoja_actual) + " " + str(numero_para_guardar_egreso) + ".xlsx",index=False)
                numero_para_guardar_egreso = numero_para_guardar_egreso + 1
    
        ############################################# APORTES OTROS MESES ##########################################################################

    for i in lista_sheets:    

            # leemos el archivo con pandas
            datos_p = pd.read_excel(nombre_archivo,sheet_name=i)
            datos_p["SUBP"] = datos_p["SUBP"].str.strip()           

            datos_ok = (datos_p["Estado de Entrega"] == "OKA") & (datos_p["Tipo Entrega Aporte"] == entrega)
            datos = datos_p[datos_ok]
            datos = datos[datos["Solicita"] != 0]        


            numero_para_guardar_egreso = 1

            hoja_actual = i
            print(i)

            #" Subdividimos los archivos de a 500 Registros"
            try:
                lista_datos = FD.registros_500(datos)
            except:
                lista_datos = []
            #"Organizamos los subproyectos para que sean leidos por el sistema siigo"

            if hoja_actual != "AM":

                for j in lista_datos:

                    for i in range(len(j["SUBP"])):
                        if len(j["SUBP"][i]) == 2:
                            j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                            j["SUBP"][i] = "00MAIN-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                            j["SUBP"][i] = "00MAEN-1"
                        else:
                            j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"

            elif hoja_actual == "AM":

                for j in lista_datos:

                    for i in range(len(j["SUBP"])):

                        if len(j["SUBP"][i]) == 2 and j["SUBP"][i] == "NJ":
                            j["SUBP"][i] = "000NJA-1"
                        elif len(j["SUBP"][i]) == 2:
                            j["SUBP"][i] = "0000" + j["SUBP"][i] + "-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAI":
                            j["SUBP"][i] = "00MAIA-1"
                        elif len(j["SUBP"][i]) == 3 and j["SUBP"][i] == "MAE":
                            j["SUBP"][i] = "00MAEA-1"
                        else:
                            j["SUBP"][i] = "000" + j["SUBP"][i] + "-1"


            #"Alimentamos el archivo modelo para importar en contabilidad"     

            for j in lista_datos:

                print("Consecutivo Nro: " + str(comprobante))
                modelo = pd.read_excel(r"C:\Users\darwi\Desktop\unbound\static\Comprobante Modelo.xlsx")
                modelo["Identificación tercero"] = j["CC Acudiente"] 
                modelo["Tipo de comprobante"] = 1
                modelo["Consecutivo comprobante"] = comprobante
                modelo["Fecha de elaboración "] = fecha
                modelo["Débito"] = j["Solicita"]
                modelo["Código cuenta contable"] = 28150505
                modelo["Código centro/subcentro de costos"] = j["SUBP"]
                modelo["Descripción"] = ("APORTE MESES ANTERIORES " + month + " " +hoja_actual)
                banco = [1,modelo["Consecutivo comprobante"][0], modelo["Fecha de elaboración "][0],"","",11100511,890903938,"","","","","","","","","","","","",modelo["Descripción"][0],"000100-1","",sum(modelo["Débito"]),"","","",""]
                modelo.loc[len(modelo.index)] = banco



            # le asignamos el siguiente valor al comprobante

                comprobante = comprobante + 1

            #"Exportamos el Archivo a Excel"

                modelo.drop(modelo.loc[modelo['Débito']==0].index, inplace=True) 
                if entrega == "Cuenta Bancaria":
                    modelo.to_excel("APORTE MESES ANTERIORES BANCOLOMBIA " + str(hoja_actual) + " " + str(numero_para_guardar_egreso) + ".xlsx",index=False)
                elif entrega == "Efecty":
                    modelo.to_excel("APORTE MESES ANTERIORES EFECTY " + str(hoja_actual) + " " + str(numero_para_guardar_egreso) + ".xlsx",index=False)
                numero_para_guardar_egreso = numero_para_guardar_egreso + 1


                



