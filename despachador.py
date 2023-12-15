import tkinter as tk
from tkinter import filedialog
from tkcalendar import DateEntry
import xlrd
import datetime

def lerArquivoExcel(caminhoArquivo, dataInicio, dataFim, opcaoSelecionada):
    try:
        workbook = xlrd.open_workbook(caminhoArquivo)
        planilha = workbook.sheet_by_index(0)

        colunaData = 7

        dadosFiltrados = []
        datasFiltradas = []
        for indiceLinha in range(1, planilha.nrows):
            dataLinha = planilha.cell_value(indiceLinha, colunaData)

            if isinstance(dataLinha, float):
                dataTupla = xlrd.xldate_as_tuple(dataLinha, workbook.datemode)
                dataLinha = datetime.datetime(*dataTupla)
            elif isinstance(dataLinha, str):
                try:
                    dataLinha = datetime.datetime.strptime(dataLinha, "%d/%m/%Y %H:%M:%S")
                except ValueError:
                    continue  # Ignora a linha se a data não estiver no formato esperado

            # Verifica se a data está no intervalo desejado
            if dataInicio <= dataLinha.date() <= dataFim:
                dadosFiltrados.append(planilha.row_values(indiceLinha))
                datasFiltradas.append(dataLinha.strftime("%d/%m/%Y %H:%M:%S"))

        if dadosFiltrados:
            nomeArquivo = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Arquivo de Texto", "*.txt")])
            if nomeArquivo:
                nomeArquivoComOpcao = f"{nomeArquivo[:-4]}_{opcaoSelecionada.get()}.txt"

                with open(nomeArquivoComOpcao, "w") as arquivoSaida:
                    for linha, data in zip(dadosFiltrados, datasFiltradas):
                        arquivoSaida.write("<despacho>\n")
                        arquivoSaida.write(f"{opcaoSelecionada.get()}\n")
                        for indice, valor in enumerate(linha):
                            if isinstance(valor, float):
                                # Evitar a escrita de números float indesejados
                                continue
                            arquivoSaida.write(f"{valor}\n")

                        arquivoSaida.write(f"{data}\n")  # Incluir a data e hora após percorrer os valores da linha
                        arquivoSaida.write("</despacho>\n")

        else:
            labelErro.config(text="Nenhuma linha correspondente ao intervalo selecionado.")

    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")

def selecionarArquivo():
    caminhoArquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsm")])
    if caminhoArquivo:
        lerArquivoExcel(caminhoArquivo, dataInicioEntry.get_date(), dataFimEntry.get_date(), opcaoSelecionada)

def exportarParaTxt():
    pass

root = tk.Tk()
root.title("Selecione o Arquivo Excel")

labelArquivo = tk.Label(root, text="Selecione o Arquivo Excel:")
labelArquivo.pack()

labelErro = tk.Label(root, text="", fg="red")
labelErro.pack()

botaoSelecionar = tk.Button(root, text="Selecionar Arquivo", command=selecionarArquivo)
botaoSelecionar.pack()

labelDataInicio = tk.Label(root, text="Data de início:")
labelDataInicio.pack()
dataInicioEntry = DateEntry(root, locale='pt_BR', date_pattern='dd/mm/yyyy')
dataInicioEntry.pack()

labelDataFim = tk.Label(root, text="Data de término:")
labelDataFim.pack()
dataFimEntry = DateEntry(root, locale='pt_BR', date_pattern='dd/mm/yyyy')
dataFimEntry.pack()

opcoes = [
    "E-PROC JF-RJ", "E-PROC TRF2", "E-PROC JF-PR", "E-PROC JF-RS", "E-PROC JF-SC",
    "E-PROC TRF4", "E-PROC TNU", "E-PROC TNU 2", "E-PROC TJ-SC 1", "E-PROC TJ-SC 2", "E-PROC TJ-TO 1", "E-PROC TJ-TO 2"
]

opcaoSelecionada = tk.StringVar(root)
opcaoSelecionada.set(opcoes[0])

labelOpcao = tk.Label(root, text="Selecione uma opção:")
labelOpcao.pack()
menuOpcao = tk.OptionMenu(root, opcaoSelecionada, *opcoes)
menuOpcao.pack()

root.mainloop()