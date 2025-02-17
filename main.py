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

# Função para salvar dados na planilha
def salvar_dados(temperatura, umidade):
    try:
        
        nome_arquivo = "dados_climaticos.xlsx"

        # Obter data e hora atual
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        try:
            workbook = openpyxl.load_workbook(nome_arquivo)
            sheet = workbook.active
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["Data/Hora", "Temperatura (°C)", "Umidade (%)"])

        # Adicionar dados na planilha
        sheet.append([data_hora, temperatura, umidade])

        # Salvar arquivo
        workbook.save(nome_arquivo)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar dados: {e}")

# Interface 
def criar_interface():
    janela = tk.Tk()
    janela.title("Previsão do Tempo de São Paulo")
    janela.geometry("300x100")

    # Botão para buscar previsão
    botao_buscar = tk.Button(janela, text="Buscar Previsão", command=buscar_dados, font=("Arial", 12), bg="lightblue")
    botao_buscar.pack(pady=30)

    janela.mainloop()


if __name__ == "__main__":
    criar_interface()
