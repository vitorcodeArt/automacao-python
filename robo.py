import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from ttkthemes import ThemedTk
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
import urllib
import pandas as pd


contatos_df = pd.DataFrame()


navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com/")

while len(navegador.find_elements(By.ID, "side")) < 1:
    time.sleep(3)

def enviar_mensagem(numero, texto):
    link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
    navegador.get(link)
    while len(navegador.find_elements(By.ID, "side")) < 1:
        time.sleep(3)

    while True:
        try:
            enviarMsg = navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span')
            enviarMsg.send_keys(Keys.ENTER)
            time.sleep(6)
            break
        except NoSuchElementException:
            pass

        # Verifica se o alerta de número inválido apareceu
        try:
            alerta = navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button')
            if alerta.is_displayed():
                navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button').click()
                break
        except NoSuchElementException:
            pass    
        
def abrir_arquivo():
    global contatos_df
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivo Excel", "*.xlsx")])
    if arquivo:
        contatos_df = pd.read_excel(arquivo)
        contatos_df['Mensagem'] = ''
        contatos_df["Telefone"] = contatos_df.apply(lambda row: ';'.join(str(num).split('.')[0] for num in [row['Telefone 1'], row['Telefone 2'], row['Telefone 3']] if pd.notnull(num)), axis=1)
        while 'Telefone' not in contatos_df.columns:
            print('Aguardando criação da coluna "Telefone"...')
            time.sleep(1)  # aguarda 1 segundo
        time.sleep(4)
            
        nome_arquivo = arquivo.split("/")[-1]  # obtém o nome do arquivo a partir do caminho completo
        rotulo_arquivo.configure(text=f"Arquivo selecionado: {nome_arquivo}")          

def enviar():
    data = entrada_data.get()
    for i, mensagem in enumerate(contatos_df['Mensagem']):
        cliente = contatos_df.loc[i, "Cliente"]
        endereco = contatos_df.loc[i, "Endereço"]
        numero = contatos_df.loc[i, "Número"]
        complemento = str(contatos_df.loc[i, "Complemento"])
        ocorrencia = contatos_df.loc[i, "Ultimo Motivo/Ocorrencia"]
        if complemento == 'nan':
            complemento = ''
        os = contatos_df.loc[i, "BA"]
        telefones = str(contatos_df.loc[i, "Telefone"]).split(";") # separa os números de telefone por ";" em uma lista
        cep = contatos_df.loc[i, "Cep"]
        opcoes_texto = {
            "Notas Novas": 
        f"Recentemente recebemos o cancelamento de serviços da Vivo. Por isso, precisamos agendar a coleta dos equipamentos Vivo que você não utiliza mais. Para isso, precisamos que você confirme algumas informações:\n\n •  *{cliente}*\n •  *{endereco}, {numero} -{ complemento}.*\n •  *CEP: {cep}*\n\nA retirada sempre acontece em horário comercial. Deixamos marcado a retirada dos seus equipamentos para *{data}* (desde que seja feita a confirmação).\n\nSomente maiores de 18 anos poderão realizar esse procedimento, ok? Então caso não tenha ninguém para fazer a entrega dos equipamentos para a nossa equipe, pedimos que você nos responda com a melhor data/horário.\n\nPara sua segurança, você pode confirmar se realmente é a nossa equipe que irá fazer a retirada, conferindo o nº da sua coleta: *{os}*.\n\n*Importante:* Como nossa equipe não está autorizada a entrar na sua residência, pedimos que você já deixe os equipamentos e acessórios (ex: fonte, controle e cabos) em uma sacola ou em uma caixa.\n\nO serviço de coleta é totalmente gratuito e é realizado em parceria com LOCALTEC, que é nosso fornecedor autorizado.\n\nSe quiser mais informações, acesse: www.vivo.com.br/devolverequipamento \nSegue telefone para contato: (11) 3949-7557. \nAté breve,\n\nEquipe Vivo",
    "Cliente AUSENTE/MUDOU-SE": 
        f"Olá *{cliente}*\nOS: *{os}*\nEndereço: *{endereco}, {numero} /{ complemento}*\nNosso técnico foi ao local para a coleta dos equipamentos da Vivo e em laudo constou *{ocorrencia}* seria possível Agendar a coleta para dia *{data}*? Caso preferir podemos alterar o endereço da coleta!\n(Informe CEP, NÚMERO DO LOCAL, E COMPLEMENTOS)\n\nEstamos a disposição!\n\nAgradecemos a atenção, aguardamos um retorno",
    "Agendamento LOCALTEC": 
        f"Olá:*{cliente}*\nOS:*{os}*\nRecebemos um agendamento para retirada dos aparelhos da VIVO para o dia *{data}*, podemos confirmar?\nCaso seja apartamento, pode deixar na portaria se preferir.\nMe confirme o endereço por favor:*{endereco}, {numero} / {complemento}*\n*CEP:{cep}*\nPode ser retirado em horário comercial das 8h as 18h?\nAtenciosamente,\nLocaltec."
        }
        opcao = combo_texto.get()
        texto = urllib.parse.quote(opcoes_texto[opcao])
        numeros_enviados = []
        for telefone in telefones:
            for telefone in telefones: 
                telefone_sem_virgula = telefone.split('.')[0]
                if isinstance(telefone, str) and telefone.strip() and telefone not in numeros_enviados:       
                    enviar_mensagem(telefone_sem_virgula, texto)
                    numeros_enviados.append(telefone_sem_virgula)
            else:
                print(f"Telefone inválido: {telefone} OS:{os}")


janela = ThemedTk(theme='arc')
janela.geometry("400x300")
janela.configure(width='400px', height='300px', background='#fff')

# criando um frame para centralizar os widgets
frame_central = tk.Frame(janela, bg='#fff')
frame_central.pack(expand=True)



opcoes = ["Notas Novas", 
          "Cliente AUSENTE/MUDOU-SE",
          "Agendamento LOCALTEC"
          ]

combo_texto = ttk.Combobox(frame_central, values=opcoes)
combo_texto.current(0)  # define a opção padrão como a primeira da lista
combo_texto.pack(padx=10, pady=10)


botao_arquivo = tk.Button(frame_central, text="Selecionar arquivo", command=abrir_arquivo)
botao_arquivo.pack(pady=10)

# adiciona um rótulo para mostrar o nome do arquivo selecionado
rotulo_arquivo = tk.Label(frame_central, text="Nenhum arquivo selecionado")
rotulo_arquivo.pack(pady=5)

rotulo_data = tk.Label(frame_central, text="Data (dia/mês):")
rotulo_data.configure(highlightthickness='0', foreground='#fff', background='#4DB79F')
rotulo_data.pack(pady=10)

entrada_data = tk.Entry(frame_central)
entrada_data.configure(bd=1)
entrada_data.pack(padx=10, pady=10)

botao_enviar = tk.Button(frame_central, text="Enviar Mensagens", command=enviar)
botao_enviar.configure(background='#4DB79F', foreground='#fff')
botao_enviar.pack(pady=10) 

paragrafo = tk.Label(janela, text="5% dos agendamentos feitos deverão ir para o Vitor", anchor="se")
paragrafo.pack(side="bottom", padx=10, pady=10)
paragrafo.configure(foreground="#000", background="#fff", font=('Arial', 4))

# centralizando o frame
janela.eval('tk::PlaceWindow %s center' % janela.winfo_toplevel())
janela.mainloop()