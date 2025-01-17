import tkinter as tk
from tkinter import messagebox
import requests
from datetime import datetime
import openpyxl

# Função para capturar dados meteorológicos
def buscar_dados():
    try:
        API_KEY = "efead565e9cc8e21d50060bd1a0cfb2b"
        LOCATION = "São Paulo"
        URL = f"http://api.openweathermap.org/data/2.5/weather?q={LOCATION}&appid={API_KEY}&units=metric&lang=pt_br"

        response = requests.get(URL)
        data = response.json()

        if response.status_code == 200:
            temperatura = data["main"]["temp"]
            umidade = data["main"]["humidity"]

            salvar_dados(temperatura, umidade)
            messagebox.showinfo("Sucesso", "Dados capturados e salvos com sucesso!")
        else:
            messagebox.showerror("Erro", f"Erro na API: {data.get('message', 'Desconhecido')}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao buscar dados: {e}")




