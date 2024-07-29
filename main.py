import subprocess
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import messagebox
import locale
import os
from datetime import datetime
# Define o local para português do Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def chamar_script():
    if os.path.exists(r"modelos_jfa.xlsx"):
        os.remove(r"modelos_jfa.xlsx")
    if os.path.exists(r"modelos_stetson.xlsx"):
        os.remove(r"modelos_stetson.xlsx")
    if os.path.exists(r"modelos_taramps.xlsx"):
        os.remove(r"modelos_taramps.xlsx")
    if os.path.exists(r"modelos_usina.xlsx"):
        os.remove(r"modelos_usina.xlsx")
    if os.path.exists(r"produtos.xlsx"):
        os.remove(r"produtos.xlsx")
    dia_inicial = cal_inicial.get_date().strftime('%Y-%m-%d')
    dia_final = cal_final.get_date().strftime('%Y-%m-%d')
    
    # Fecha a janela do Tkinter antes de executar os scripts
    janela.destroy()

    # Lista de scripts para executar
    scripts = ['jfa.py','stetson.py', 'taramps.py', 'usina.py']#'jfa.py','stetson.py', 'taramps.py', 'usina.py'
    
    for script in scripts:
        # Monta o comando para chamar o script com os argumentos necessários
        comando = [
            'python',
            script,
            '--dia_inicial', dia_inicial,
            '--dia_final', dia_final
        ]
        
        # Executa o comando e captura a saída
        resultado = subprocess.run(comando, capture_output=True, text=True)
        print(f"Saída de {script}:")
        print(resultado.stdout)
        print(resultado.stderr)
    
    # Mostra uma mensagem de conclusão após a execução dos scripts
    messagebox.showinfo("Conclusão", "Todos os scripts foram executados com sucesso!")

# Cria a janela principal
janela = tk.Tk()
janela.title('Market Share')
data_atual = datetime.now()
# Cria e posiciona os elementos da interface
ttk.Label(janela, text='Data Inicial:').grid(column=0, row=0, padx=10, pady=10)
cal_inicial = DateEntry(janela, width=22, background='darkblue', foreground='white', borderwidth=2, locale='pt_BR', day=data_atual.day - 1)
cal_inicial.grid(column=1, row=0, padx=10, pady=10)

ttk.Label(janela, text='Data Final:').grid(column=0, row=1, padx=10, pady=10)
cal_final = DateEntry(janela, width=22, background='darkblue', foreground='white', borderwidth=2, locale='pt_BR', day=data_atual.day - 1)
cal_final.grid(column=1, row=1, padx=10, pady=10)

ttk.Button(janela, text='Executar', command=chamar_script).grid(column=0, row=2, columnspan=2, pady=10)

janela.mainloop()
