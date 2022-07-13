from django.contrib import admin
from django.shortcuts import render
from django.urls import path
from requests import request
from . import views
from applications.archivo_txt.funciones import archivo_txt



app_name = "unbound_app"

urlpatterns = [
    path('',views.InicioUnbound.as_view(),name="inicio"),
    path("siigo/",views.RegistroInicial.as_view(),name="registro"),
    path("txt/",views.upload,name="upload"),
    path("siigo/cuenta_28/",views.cuenta_28,name="cuenta_28"),
    path("siigo/cuenta_41/",views.cuenta_41,name="cuenta_41"),
    path("siigo/egreso/",views.egreso_siigo,name="egreso_siigo"),
    path("abila/",views.Abila.as_view(),name="abila"),
    path("abila/ingreso/",views.ingreso_abila,name="abila_ingreso"),
    path("abila/egreso/",views.egreso_abila,name="egreso_ingreso"),
]

