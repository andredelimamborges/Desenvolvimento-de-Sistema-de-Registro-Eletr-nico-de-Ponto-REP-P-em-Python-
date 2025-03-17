# ⏰ Controle de Horas Automatizado 📊

Um script Python que rastreia horas trabalhadas usando uma planilha Excel como gatilho, com segurança e precisão!

---

## ✨ Funcionalidades Principais
- **🎯 Detecção Inteligente**: Monitora quando uma planilha específica é aberta/fechada
- **🔒 Segurança Avançada**: Criptografia de dados com Fernet
- **⏱ Precisão Temporal**: Sincronização com servidor NTP global
- **📈 Exportação Automática**: Salva relatórios em Excel formatados
- **📝 Logs Detalhados**: Registro de todas as atividades em arquivo rotativo

---

## 🛠 Como Funciona
1. **🔎 Monitoramento Constante**: Verifica a cada 10 segundos se a planilha gatilho está aberta
2. **⏲ Cronômetro Integrado**: 
   - Inicia quando o arquivo é aberto
   - Para quando o arquivo é fechado
3. **🧮 Cálculo Inteligente**:
   - Horas normais (até 8h)
   - Horas extras (acima de 8h)
   - Pausa automática de 1h após 6h de trabalho
4. **💾 Salvamento Seguro**: Dados criptografados e armazenados em nova planilha

---

## 🚀 Configuração Rápida
1. **Instale as dependências**:
   ```bash
   pip install psutil pandas cryptography ntplib openpyxl

   Configure o caminho:

python
Copy
PASTA_BASE = r'Insira o caminho da sua pasta'
Execute:

bash
Copy
python controle_horas.py
🛡️ Recursos de Segurança
Recurso	Descrição
Criptografia	Dados protegidos com chave AES-128
Logs Rotativos	Arquivos de log limitados a 5MB
Verificação	3 tentativas de salvamento em caso de erro
⚠️ Importante!
Mantenha a planilha gatilho SISTEMA OPERACIONAL_MS_FILIAL MGI.xlsx aberta durante o trabalho

Não altere o nome do arquivo gatilho

Backup automático mantém até 2 versões anteriores de logs

🔍 Troubleshooting
Problema	Solução
"Erro ao criar pasta"	Verifique permissões de rede/drive
"Planilha travada"	Feche o Excel e tente novamente
"Horas incorretas"	Verifique sincronização NTP
📝 Notas de Desenvolvimento
python
Copy
# Estrutura principal
if planilha_aberta(GATILHO):
    inicia_contagem()
while planilha_aberta(GATILHO):
    atualiza_contagem()
salva_dados()
