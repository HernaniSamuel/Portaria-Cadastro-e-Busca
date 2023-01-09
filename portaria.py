import customtkinter as ctk
import pandas as pd
from datetime import date

ctk.set_appearance_mode('light', )
ctk.set_default_color_theme('green')

root = ctk.CTk()
root.geometry("500x450")


def salvar(df):
    try:
        df.to_excel(r'C:\Users\User\Desktop\Portaria.xlsx')
    except:
        root = ctk.CTkToplevel()
        root.geometry("800x250")
        root.title('Falha de salvamento')

        frame = ctk.CTkFrame(master=root)
        frame.pack(pady=20, padx=60, fill="both", expand=True)

        label = ctk.CTkLabel(master=frame, text="Falha ao salvar arquivo, planilha aberta em outro local!", font=('Arial', 20))
        label.pack(pady=12, padx=10)

        frame.mainloop()


def resgatar():
    try:
        backup = pd.read_excel(r'C:\Users\User\Desktop\Portaria.xlsx')
        df = pd.DataFrame(columns=['Data', 'Nome', 'Identidade', 'Placa', 'Motivo/Destino', 'Entrada', 'Saida'])
        df['Data'] = backup['Data']
        df['Nome'] = backup['Nome']
        df['Identidade'] = backup['Identidade']
        df['Placa'] = backup['Placa']
        df['Motivo/Destino'] = backup['Motivo/Destino']
        df['Entrada'] = backup['Entrada']
        df['Saida'] = backup['Saida']

    except:
        df = pd.DataFrame(columns=['Data', 'Nome', 'Identidade', 'Placa', 'Motivo/Destino', 'Entrada', 'Saida'])

    return df


def formatar(placa):
    try:
        traco = 0
        for i in range(0, len(placa)):
            if placa[i] == '-':
                traco = 1
        if traco == 0:
            texto = f'{placa[0]}{placa[1]}{placa[2]}-{placa[3]}{placa[4]}{placa[5]}{placa[6]}'
        else:
            texto = f'{placa[0]}{placa[1]}{placa[2]}{placa[3]}{placa[4]}{placa[5]}{placa[6]}{placa[7]}'
        return texto.upper()
    except:
        texto = placa
        return texto.upper()


def menu():
    root = ctk.CTk()
    root.geometry("500x250")
    root.title("Menu")

    frame = ctk.CTkFrame(master=root)
    frame.pack(pady=20, padx=60, fill="both", expand=True)

    label = ctk.CTkLabel(master=frame, text="Menu", font=('Roboto', 40))
    label.pack(pady=12, padx=10)

    button = ctk.CTkButton(master=frame, text="Cadastrar", command=cadastro)
    button.pack(pady=12, padx=10)

    button = ctk.CTkButton(master=frame, text="Buscar", command=busca)
    button.pack(pady=12, padx=10)

    root.mainloop()


def busca():
    tela_busca = ctk.CTkToplevel()
    tela_busca.geometry("500x450")
    tela_busca.title('Busca')

    frame = ctk.CTkFrame(master=tela_busca)
    frame.pack(pady=20, padx=60, fill="both", expand=True)

    label = ctk.CTkLabel(master=frame, text="Busca", font=('Roboto', 30))
    label.pack(pady=12, padx=10)

    placa = ctk.CTkEntry(master=frame, placeholder_text="Placa (sem espaço)")
    placa.pack(pady=12, padx=10)
    saida = ctk.CTkEntry(master=frame, placeholder_text="Saída")
    saida.pack(pady=12, padx=10)

    def buscar():
        df = resgatar()
        Placa_carro = placa.get()
        Placa = formatar(Placa_carro)
        for i in range(0, len(df)):
            if Placa == df['Placa'][i] and df['Saida'][i] == '-':
                Saida = saida.get()
                df['Saida'][i] = Saida
                salvar(df)
                label_encontrada = ctk.CTkLabel(master=frame, text="Saída registrada!")
                label_encontrada.pack(pady=12, padx=10)
                frame.update()
                break

    botao_buscar = ctk.CTkButton(master=frame, text="Pesquisar", command=buscar)
    botao_buscar.pack(pady=12, padx=10)


def cadastro():
    tela_cadastro = ctk.CTkToplevel()
    tela_cadastro.geometry("500x450")
    tela_cadastro.title("Cadastro")

    frame = ctk.CTkFrame(master=tela_cadastro)
    frame.pack(pady=20, padx=60, fill="both", expand=True)

    label = ctk.CTkLabel(master=frame, text="Cadastro", font=('Roboto', 30))
    label.pack(pady=12, padx=10)

    nome = ctk.CTkEntry(master=frame, placeholder_text="Nome")
    nome.pack(pady=12, padx=10)

    identidade = ctk.CTkEntry(master=frame, placeholder_text="Identidade")
    identidade.pack(pady=12, padx=10)

    placa = ctk.CTkEntry(master=frame, placeholder_text="Placa (sem espaço)")
    placa.pack(pady=12, padx=10)

    motivo_destino = ctk.CTkEntry(master=frame, placeholder_text="Motivo/Destino")
    motivo_destino.pack(pady=12, padx=10)

    entrada = ctk.CTkEntry(master=frame, placeholder_text="Entrada")
    entrada.pack(pady=12, padx=10)

    def cadastrar():
        df = resgatar()
        Nome = nome.get()
        Identidade = identidade.get()
        Placa_carro = placa.get()
        Placa = formatar(Placa_carro)
        Motivo = motivo_destino.get()
        Entrada = entrada.get()
        data_atual = date.today()
        data = '{}/{}'.format(data_atual.day, data_atual.month)
        dados = [data, Nome, Identidade, Placa, Motivo, Entrada, '-']
        df.loc[len(df) + 1] = dados
        salvar(df)
        del df
        label = ctk.CTkLabel(master=frame, text="Cadastrado!")
        label.pack(pady=12, padx=10)
        frame.update()

    button = ctk.CTkButton(master=frame, text="Cadastrar", command=cadastrar)
    button.pack(pady=12, padx=10)


if __name__ == '__main__':
    menu()
