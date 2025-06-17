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

    nomes_alunos = df["Nome_Completo"].tolist()
    emails = df["Email"].tolist()
    inicio_atividade = df["Iniciado_em"].tolist()
    completo = df["Completo"].tolist()
    nota = df["Nota"].tolist()
    estado = df["Estado"].tolist()

    return nomes_alunos, emails, inicio_atividade, estado, completo, nota

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

    nomes_alunos, emails, inicio_atividade, estado, completo, nota = ler_dados_excel(caminho)

    for i in range(len(nomes_alunos)):
        if str(estado[i]).strip().upper() != "Finalizada":
            corpo_email = (
                f"Olá {nomes_alunos[i]},\n\n"
                f"Você ainda não entregou a atividade"
                f"Materia se deu inicio em: {inicio_atividade[i]}\n\n"
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
