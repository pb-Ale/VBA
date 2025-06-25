# VBA
Arquivos de excel que rodam códigos em VBA. 

Automação de vários emails

# Objetivo
Automatizar emails que diariamente eram enviados dentro da operação. 

Estes códigos não manipulam diretamente colunas de planilhas, mas utilizam variáveis para compor o nome e o caminho de um arquivo Excel que será anexado a um e-mail.

Todos os scripts automatizam o envio de um e-mail via Outlook, alguns contendo um arquivo Excel como anexo. 
O nome e o caminho do arquivo são gerados dinamicamente com base na data atual. O e-mail é direcionado a um destinatário específico, com uma saudação personalizada ("Bom dia" ou "Boa tarde") conforme o horário do dia, e uma mensagem solicitando conferência de acordo com o processo, alterando o destinatário conforme cada processo.

O acionamento de cada um dos códigos depende do botão acionado. Botão adicionado separadamente em cada aba para evitar envio errôneo.
Esta automação eliminou horas de trabalho diário apenas criando emails.

# Variáveis
Descrição
Tipo de dado

# dataAtual
Data atual formatada como dd-mm-yyyy
Texto (Data formatada)

# mesAno
Ano e mês atual no formato yyyy-mm
Texto (Data formatada)

# pastaMesAno
Mês e nome do mês no formato mm. mmmm (ex: 06. junho)
Texto (Data formatada)

# nomeArquivo
# Nome do arquivo Excel no formato dd mm yy.xlsx
Texto

# caminhoArquivo
Caminho completo do arquivo a ser anexado
Texto

# saudacao
Saudação baseada na hora atual (Bom dia ou Boa tarde)
Texto

# corpoEmail
Texto do corpo do e-mail
Texto

# email
Endereço de e-mail do destinatário
Texto
