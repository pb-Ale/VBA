# VBA
Arquivos de excel que rodam códigos em VBA. 

Este código cria arquivos a partir da tabela que consegue os dados dentro do excel.
A tabela possui o código identificador da 'concessionária', que é utilizado para filtrar todas as 'propostas' que esta 'concessionária' precisa atuar.
O código cria os arquivos com base neste filtro, onde todas as propostas estão em um único arquivo de excel e são salvos em um diretório.

Estes arquivos posteriormente são utilizados para serem anexados em emails.
Cada email possui o descritivo da concessionária, processo interno a ser seguido e o anexo.

# Nome da coluna	
Descrição	
Tipo de dado
# Data de Rejeição	
Data em que houve fato gerador referente ao processo que está sendo enviado o email atual	
Data

# Tipo de borderô
Classificação interna que separa em condições cada caso. 
Texto

# Cód. Assistência
Número interno de identificação do remetente
Número

# Concessionária
Nome do remetente
Texto

# Data da venda
Data em que houve a venda. Esta data é posterior à data de rejeição, esta última é o fato gerador.
Data

# Proposta
Chave primária.
Texto

# Nº do borderô
Número interno de identificação de um grupo de chaves primárias, conforme o contexto de negócio.
Número

# Valor da Proposta
Valor referente à compra.
Número decimal

# Motivo de Rejeição
Motivo do fato gerador.
Texto

# Assunto
Assunto que será preenchido no email, referente ao nome da atividade. 
Texto

# e-mail
Email do destinatário
Texto

