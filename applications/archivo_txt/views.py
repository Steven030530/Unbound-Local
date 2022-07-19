from calendar import c
from email.errors import NoBoundaryInMultipartDefect
from email.headerregistry import ContentDispositionHeader
from importlib.resources import path
from re import X
from urllib import response
from django.http import HttpResponse
from django.shortcuts import render
from django.views.generic import TemplateView,ListView
from pandas import concat
from applications.archivo_txt.funciones import archivo_txt, ingreso_28 , ingreso_41 , egreso_general_siigo
from django.core.files.storage import FileSystemStorage
from applications.archivo_txt.funciones_abila import *
from openpyxl import Workbook ,load_workbook

# Create your views here.

class InicioUnbound(TemplateView):
    template_name = 'archivo_txt/inicio.html'


class RegistroInicial(TemplateView):
    template_name = 'archivo_txt/registro_inicial.html'


class Abila(TemplateView):
    template_name = 'archivo_txt/seccion_abila.html'

class ReporteExcel(TemplateView):

    def get(self,request):
        
        fs = FileSystemStorage()
        nombre_archivo = "Archivo_Plano_Excel.xlsx"
        response = HttpResponse(content_type='application/vnd.ms-excel')
        contenido = 'attachment;filename = {0}'.format(nombre_archivo)
        response["Content-Disposition"] = contenido 
        print(fs.listdir(r"C:\Users\darwi\Desktop\proyectounbound\media")[1])
        print(len(fs.listdir(r"C:\Users\darwi\Desktop\proyectounbound\media")[1]))
        print(fs.path(fs.listdir(r"C:\Users\darwi\Desktop\proyectounbound\media")[1][0]))
        wb = load_workbook(str(fs.path(fs.listdir(r"C:\Users\darwi\Desktop\proyectounbound\media")[1][0])))
        wb.save(response)
        fs.delete(fs.path(fs.listdir(r"C:\Users\darwi\Desktop\proyectounbound\media")[1][0]))
        return response
        
        
            

def upload(request):
    

    context = {'var1': 'Archivo con error','error': 'Archivo con error'}
    if  request.method == 'POST':
        uploaded_file = request.FILES["ArchivoTxt"]
        uploaded_date = request.POST["date"]
        uploaded_emp = request.POST["empresa"]
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name,uploaded_file)
        fs.delete(fs.path(fs.listdir(r"C:\Users\darwi\Desktop\proyectounbound\media")[1][0]))
        
        try:
             archivo=archivo_txt(uploaded_file,uploaded_date,uploaded_emp)
             context['url'] = fs.url(name)
             
             
             
        except Exception as e:
            context.update({'error': e})
            print(repr(e))

    mi_page = render(request,"archivo_txt/upload.html",context)
    return mi_page

def cuenta_28(request):
    
    context = {'var1': 'Archivo con error','error': 'Archivo con error'}
    if  request.method == 'POST':
        uploaded_file = request.FILES["Dispersion"]
        uploaded_date = request.POST["date"]
        uploaded_conse = request.POST["consecutivo"]
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name,uploaded_file)
        url = fs.url(name)

        try:
             ingreso_28(uploaded_file,uploaded_date,uploaded_conse)
             context['url'] = fs.url(name)
             
        except Exception as e:
            context.update({'error': e})
            print(repr(e))
  
    return render(request,"archivo_txt/ingreso_28.html",context)

def cuenta_41(request):
    
    context = {'var1': 'Archivo con error','error': 'Archivo con error'}
    if  request.method == 'POST':
        uploaded_file = request.FILES["Dispersion"]
        uploaded_date = request.POST["date"]
        uploaded_conse = request.POST["consecutivo"]
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name,uploaded_file)
        url = fs.url(name)

        try:
             ingreso_41(uploaded_file,uploaded_date,uploaded_conse)
             context['url'] = fs.url(name)
             
        except Exception as e:
            context.update({'error': e})
            print(repr(e))
        
    return render(request,"archivo_txt/ingreso_41.html",context)

def egreso_siigo(request):
    
    context = {'var1': 'Archivo con error','error': 'Archivo con error'}
    if  request.method == 'POST':
        uploaded_file = request.FILES["Dispersion"]
        uploaded_date = request.POST["date"]
        uploaded_conse = request.POST["consecutivo"]
        uploaded_entre = request.POST["entrega"]
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name,uploaded_file)
        url = fs.url(name)

        try:
             egreso_general_siigo(uploaded_file,uploaded_date,uploaded_conse,uploaded_entre)
             context['url'] = fs.url(name)
             
        except Exception as e:
            context.update({'error': e})
            print(repr(e))
        
    return render(request,"archivo_txt/egreso_siigo.html",context)

def ingreso_abila(request):
    
    context = {'var1': 'Archivo con error','error': 'Archivo con error'}
    if  request.method == 'POST':
        uploaded_file = request.FILES["Dispersion"]
        uploaded_date = request.POST["date"]
        uploaded_conse = request.POST["consecutivo"]
        uploaded_entre = ["NNJ","AM","V"]
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name,uploaded_file)
        url = fs.url(name)
        for i in uploaded_entre:
            try:
                Ingresos_Abila.aporte_general(i,uploaded_date,uploaded_conse,uploaded_file)
                context['url'] = fs.url(name)
                Ingresos_Abila.aporte_cumpleanios(i,uploaded_date,uploaded_conse,uploaded_file)
                context['url'] = fs.url(name)
                Ingresos_Abila.aporte_regalo(i,uploaded_date,uploaded_conse,uploaded_file)
                context['url'] = fs.url(name)
                
                
            except Exception as e:
                context.update({'error': e})
                print(repr(e))
        try:
            Consolidar.consolidacion_archivos("INGRESO")
            context['url'] = fs.url(name)
        except Exception as e:
            context.update({'error': e})
            print(repr(e))

    return render(request,"archivo_txt/ingreso_abila.html",context)


def egreso_abila(request):
    
    context = {'var1': 'Archivo con error','error': 'Archivo con error'}
    if  request.method == 'POST':
        uploaded_file = request.FILES["Dispersion"]
        uploaded_date = request.POST["date"]
        uploaded_conse = request.POST["consecutivo"]
        uploaded_empresa = request.POST["empresa"]
        uploaded_entre = ["NNJ","AM","V","FC"]
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name,uploaded_file)
        url = fs.url(name)
        for i in uploaded_entre:
            try:
                Egreso_Abila.egreso_general(i,uploaded_empresa,uploaded_date,uploaded_conse,uploaded_file)
                context['url'] = fs.url(name)
                Egreso_Abila.egreso_cumple(i,uploaded_empresa,uploaded_date,uploaded_conse,uploaded_file)
                context['url'] = fs.url(name)
                Egreso_Abila.egreso_regalo(i,uploaded_empresa,uploaded_date,uploaded_conse,uploaded_file)
                context['url'] = fs.url(name)
                Egreso_Abila.egreso_fomentado(i,uploaded_empresa,uploaded_date,uploaded_conse,uploaded_file)
                context['url'] = fs.url(name)
                
                
            except Exception as e:
                context.update({'error': e})
                print(repr(e))
        try:
            Consolidar.consolidacion_archivos("EGRESO")
            context['url'] = fs.url(name)
        except Exception as e:
            context.update({'error': e})
            print(repr(e))
    return render(request,"archivo_txt/egreso_abila.html",context)