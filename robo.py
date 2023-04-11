import tkinter as tk
from ttkthemes import ThemedTk
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
import urllib
import pandas as pd


contatos_df = pd.read_excel("enviar11-04(3).xlsx")

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
            time.sleep(3)
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

def enviar():
    data = entrada_data.get()
    for i, mensagem in enumerate(contatos_df['Mensagem']):
        cliente = contatos_df.loc[i, "Cliente"]
        endereco = contatos_df.loc[i, "Endereço"]
        numero = contatos_df.loc[i, "Número"]
        complemento = str(contatos_df.loc[i, "Complemento"])
        if complemento == 'nan':
            complemento = ''
        os = contatos_df.loc[i, "BA"]
        telefones = str(contatos_df.loc[i, "Telefone"]).split(";") # separa os números de telefone por ";" em uma lista
        cep = contatos_df.loc[i, "Cep"]
        texto = urllib.parse.quote(f"Recentemente recebemos o cancelamento de serviços da Vivo. Por isso, precisamos agendar a coleta dos equipamentos Vivo que você não utiliza mais. Para isso, precisamos que você confirme algumas informações:\n\n •  *{cliente}*\n •  *{endereco}, {numero} -{ complemento}*\n •  *CEP: {cep}*\n\nA retirada sempre acontece em horário comercial. Deixamos marcado a retirada dos seus equipamentos para *{data}* (desde que seja feita a confirmação).\n\nSomente maiores de 18 anos poderão realizar esse procedimento, ok? Então caso não tenha ninguém para fazer a entrega dos equipamentos para a nossa equipe, pedimos que você nos responda com a melhor data/horário.\n\nPara sua segurança, você pode confirmar se realmente é a nossa equipe que irá fazer a retirada, conferindo o nº da sua coleta: *{os}*.\n\n*Importante:* Como nossa equipe não está autorizada a entrar na sua residência, pedimos que você já deixe os equipamentos e acessórios (ex: fonte, controle e cabos) em uma sacola ou em uma caixa.\n\nO serviço de coleta é totalmente gratuito e é realizado em parceria com LOCALTEC, que é nosso fornecedor autorizado.\n\nSe quiser mais informações, acesse: www.vivo.com.br/devolverequipamento \nSegue telefone para contato: (11) 3949-7557. \nAté breve,\n\nEquipe Vivo")
        for telefone in telefones: 
            if telefone.strip():
                enviar_mensagem(telefone, texto)

janela = ThemedTk(theme='arc')
janela.geometry("400x300")
janela.configure(width='400px', height='300px', background='#fff')

# criando um frame para centralizar os widgets
frame_central = tk.Frame(janela, bg='#fff')
frame_central.pack(expand=True)

rotulo_data = tk.Label(frame_central, text="Data (dia/mês):")
rotulo_data.configure(highlightthickness='0', foreground='#fff', background='#4DB79F')
rotulo_data.pack(pady=10)

entrada_data = tk.Entry(frame_central)
entrada_data.configure(bd=1)
entrada_data.pack(padx=10, pady=10)

botao_enviar = tk.Button(frame_central, text="Enviar Mensagens", command=enviar)
botao_enviar.configure(background='#4DB79F', foreground='#fff')
botao_enviar.pack(pady=10)

# centralizando o frame
janela.eval('tk::PlaceWindow %s center' % janela.winfo_toplevel())
janela.mainloop()