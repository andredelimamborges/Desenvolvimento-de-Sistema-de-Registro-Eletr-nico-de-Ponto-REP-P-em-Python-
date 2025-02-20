# -*- coding: utf-8 -*-
import psutil
import os
import time
from datetime import datetime, timedelta
import pandas as pd
import logging
from logging.handlers import RotatingFileHandler
from cryptography.fernet import Fernet
# type: ignore
import ntplib


# =============================================
# CONFIGURAÇÕES 
# =============================================
PASTA_BASE = r'G:\Meu Drive\MS Avaliação_Cloud\Sistema Operacional Consultores' o
CAMINHO_GATILHO = os.path.join(PASTA_BASE, 'SISTEMA OPERACIONAL_MS_FILIAL MGI.xlsx')  
CAMINHO_SAIDA = os.path.join(PASTA_BASE, 'controle_horas.xlsx')  
TEMPO_VERIFICACAO = 10  # 
CHAVE_CRIPTOGRAFIA = os.path.join(PASTA_BASE, 'chave_secreta.key') 

# =============================================
# VERIFICAÇÃO E CRIAÇÃO DA PASTA BASE
# =============================================
if not os.path.exists(PASTA_BASE):
    try:
        os.makedirs(PASTA_BASE)  
        print(f"Pasta base criada em: {PASTA_BASE}")
    except Exception as e:
        print(f"Erro ao criar pasta base: {e}")
        exit(1)

# =============================================
# CONFIGURAÇÃO DE LOGS
# =============================================
def configurar_logs():
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    log_handler = RotatingFileHandler(os.path.join(PASTA_BASE, 'controle_horas.log'), maxBytes=5*1024*1024, backupCount=2)
    log_handler.setFormatter(log_formatter)

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.addHandler(log_handler)
    return logger

logger = configurar_logs()

# =============================================
# CRIPTOGRAFIA
# =============================================
def gerar_ou_carregar_chave():
    if os.path.exists(CHAVE_CRIPTOGRAFIA):
        with open(CHAVE_CRIPTOGRAFIA, 'rb') as arquivo_chave:
            return arquivo_chave.read()
    else:
        chave = Fernet.generate_key()
        with open(CHAVE_CRIPTOGRAFIA, 'wb') as arquivo_chave:
            arquivo_chave.write(chave)
        return chave

chave = gerar_ou_carregar_chave()
fernet = Fernet(chave)

def criptografar_dados(dados):
    return fernet.encrypt(dados.encode())

def descriptografar_dados(dados_criptografados):
    return fernet.decrypt(dados_criptografados).decode()

# =============================================
# SINCRONIZAÇÃO COM SERVIDOR NTP
# =============================================
def obter_tempo_ntp():
    cliente_ntp = ntplib.NTPClient()
    try:
        resposta = cliente_ntp.request('pool.ntp.org')
        return datetime.fromtimestamp(resposta.tx_time)
    except ntplib.NTPException as e:
        logger.error(f"Erro ao obter tempo do servidor NTP: {e}")
        return datetime.now() 

# =============================================
# FUNÇÕES AUXILIARES (REVISADAS)
# =============================================
def planilha_aberta(caminho):
    """Verifica se a planilha está aberta no Excel."""
    try:
        caminho_absoluto = os.path.abspath(caminho).lower()
        for proc in psutil.process_iter(['pid', 'name', 'open_files']):
            if proc.info['name'] == 'EXCEL.EXE':
                for arquivo in proc.info.get('open_files', []):
                    if os.path.abspath(arquivo.path).lower() == caminho_absoluto:
                        return True
        return False
    except Exception as e:
        logger.error(f"Falha ao verificar planilha: {e}")
        return False

def calcular_horas(inicio, fim):
    """Calcula horas trabalhadas e extras."""
    total = fim - inicio
    if total >= timedelta(hours=6):
        total -= timedelta(hours=1)
    horas_normais = min(total, timedelta(hours=8))
    horas_extras = max(total - timedelta(hours=8), timedelta(0))
    return horas_normais, horas_extras

def formatar_tempo(tempo):
    """Converte timedelta para formato HH:MM:SS."""
    horas, resto = divmod(tempo.seconds, 3600)
    minutos, segundos = divmod(resto, 60)
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"

def salvar_horas(inicio, fim):
    """Salva os dados na planilha de saída."""
    try:
       
        horas_normais, horas_extras = calcular_horas(inicio, fim)
        
 
        novo_registro = {
            'Data': inicio.date(),
            'Entrada': inicio.strftime('%H:%M:%S'),
            'Saída': fim.strftime('%H:%M:%S'),
            'Horas Normais': formatar_tempo(horas_normais),
            'Horas Extras': formatar_tempo(horas_extras)
        }
        
      
        if os.path.exists(CAMINHO_SAIDA):
            df_existente = pd.read_excel(CAMINHO_SAIDA)
            df_novo = pd.DataFrame([novo_registro])
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        else:
            df_final = pd.DataFrame([novo_registro])
        
      
        tentativas = 0
        while tentativas < 3:
            try:
                df_final.to_excel(CAMINHO_SAIDA, index=False)
                logger.info(f"Dados salvos em {CAMINHO_SAIDA}!")
                print(f"Dados salvos em {CAMINHO_SAIDA}!")
                return
            except PermissionError:
                logger.warning("Feche a planilha de saída para salvar! Tentando novamente...")
                print("[AVISO] Feche a planilha de saída para salvar! Tentando novamente...")
                time.sleep(5)
                tentativas += 1
        logger.error("Não foi possível salvar após 3 tentativas.")
        print("[ERRO] Não foi possível salvar após 3 tentativas.")
    except Exception as e:
        logger.critical(f"Falha ao salvar: {e}")
        print(f"[ERRO CRÍTICO] Falha ao salvar: {e}")

# =============================================
# LÓGICA PRINCIPAL (COM DEBUG)
# =============================================
def main():
    print(f"""
    ====================================
    CONTROLE DE HORAS - INICIADO
    Gatilho: {CAMINHO_GATILHO}
    Saída: {CAMINHO_SAIDA}
    ====================================
    """)
    logger.info("Controle de horas iniciado.")
    
    try:
        while True:
            while not planilha_aberta(CAMINHO_GATILHO):
                print("Aguardando abertura do gatilho...")
                logger.info("Aguardando abertura do gatilho...")
                time.sleep(TEMPO_VERIFICACAO)
            
            inicio = obter_tempo_ntp()
            print(f"\n▶ Contagem INICIADA em {inicio.strftime('%d/%m/%Y %H:%M:%S')}")
            logger.info(f"Contagem INICIADA em {inicio.strftime('%d/%m/%Y %H:%M:%S')}")
            
            while planilha_aberta(CAMINHO_GATILHO):
                print(f"Tempo atual: {datetime.now().strftime('%H:%M:%S')} (Mantenha o gatilho aberto)")
                time.sleep(TEMPO_VERIFICACAO)
            
            fim = obter_tempo_ntp()
            print(f"⏹ Contagem FINALIZADA em {fim.strftime('%H:%M:%S')}")
            logger.info(f"Contagem FINALIZADA em {fim.strftime('%H:%M:%S')}")
            salvar_horas(inicio, fim)
            
    except KeyboardInterrupt:
        print("\nPrograma encerrado pelo usuário.")
        logger.info("Programa encerrado pelo usuário.")
    except Exception as e:
        print(f"\n[ERRO] Ocorreu um erro inesperado: {e}")
        logger.error(f"Erro inesperado: {e}")

if __name__ == "__main__":
    main()
