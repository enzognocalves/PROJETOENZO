import pandas as pd
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuração do logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Função para ler os dados da planilha
def ler_dados_excel(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)

    nomes_alunos = df["NOME DO ALUNO"].tolist()
    nomes_professores = df["NOME DO PROFESSOR"].tolist()
    atividades = df["ATIVIDADE"].tolist()
    disciplinas = df["DISCIPLINA"].tolist()
    emails = df["EMAIL DO ALUNO"].tolist()
    datas_limite = df["DATA LIMITE"].tolist()
    entregues = df["ENTREGUE"].tolist()

    return nomes_alunos, nomes_professores, atividades, disciplinas, emails, datas_limite, entregues

# Função para enviar e-mails
def enviar_email(destinatario, assunto, corpo, remetente, senha):
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    msg.attach(MIMEText(corpo, 'plain'))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as servidor:
            servidor.login(remetente, senha)
            servidor.send_message(msg)
        logging.info(f"E-mail enviado para {destinatario}")
    except Exception as e:
        logging.error(f"Erro ao enviar e-mail para {destinatario}: {e}")

# Função principal que utiliza os dados e envia os e-mails
def usar_dados():
    caminho = r"C:\Users\enzog\OneDrive\Área de Trabalho\Projeto\validacao_atividades.xlsx"
    remetente = "enzogville8@gmail.com"
    senha = "zmjigrqcufahvxul" # Substitua por sua senha de app do Gmail

    nomes_alunos, nomes_professores, atividades, disciplinas, emails, datas_limite, entregues = ler_dados_excel(caminho)

    for i in range(len(nomes_alunos)):
        if str(entregues[i]).strip().upper() == "FALSE":
            corpo_email = (
                f"Olá {nomes_alunos[i]},\n\n"
                f"Você ainda não entregou a atividade: {atividades[i]}\n"
                f"Disciplina: {disciplinas[i]}\n"
                f"Professor: {nomes_professores[i]}\n"
                f"Data limite para entrega: {datas_limite[i]}\n\n"
                "Por favor, realize a entrega o quanto antes até a hora 23:59 da data vigente.\n\n"
                "Atenciosamente,\n"
                "Coordenação"
            )

            enviar_email(
                destinatario=emails[i],
                assunto="Pendência de Entrega de Atividade",
                corpo=corpo_email,
                remetente=remetente,
                senha=senha
            )

# Executar
usar_dados()
