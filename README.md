# VBA
Arquivos de excel que rodam códigos em VBA. 

# Resumo do código e seu objetivo
Este script VBA automatiza o preenchimento de uma planilha de controle de ocorrências a partir de dados extraídos de outra aba chamada "PreencherBase". 
Ele localiza informações específicas conforme as colunas abaixo e insere esses dados na próxima linha disponível da planilha "Controle ocorrências". 

O responsável humano da atividade copia as informações do sistema, cola na tablea aba 'PreencherBase' que então copia as informações após localizar o título de cada linha
colando posteriormente na aba de 'controle de ocorrências'.

_Formatação
Ao colar a informação vem do sistema formatada através de linhas e não com os títulos nas colunas.

Sendo
Título da coluna: Informação buscada

_Status
A seguir há a descrição da coluna 'status', recomendo lê-la antes da seguinte instrução.
Ao copiar do sistema e colar a informação na aba 'PreencherBase' o código em VBA executa a verificação a seguir:

        If Not celulaStatus Is Nothing Then 'status
            celulaStatus.Offset(1, 0).Copy
            wsPreenchedor.Range("A22").PasteSpecial Paste:=xlPasteValues
            wsPreenchedor.Range("C22").Copy
            wsBase.Range("I" & lastRow).PasteSpecial Paste:=xlPasteValues
        End If

Se a céula de status naõ está vazia, copia a informação que encontrou a partir do título da linha. Cola na célula especificada "A22".
E dentro do excel há uma fórmula que busca o status por extenso e converte para sigla, que é o necessário para este processo interno.

_tratamento de erros
O código também trata casos especiais como ausência de dados bancários, ausência de borderô, e registros com status "DG", aplicando formatações visuais para facilitar a identificação. Ao final, limpa os dados temporários da aba de preenchimento.

# Linhas existentes
Descrição da coluna
Tipo de dado

# Proposta/dígito
Código identificador de um processo
Texto

# Contrato
Segundo código identificador para casos que não avançaram na jornada de um cliente.
Exemplo: o cliente tem uma proposta que é o identificador, e o contrato seria apenas atribuído a este cliente caso efetivamente aceitasse a proposta.
Texto

# Consorciado
Nome do cliente
Texto

# Valor Financeiro
Valor da primeira parcela a ser paga.
Número decimal

# Valor Informado
Valor pago da primeira parcela.
Que pode ser menor ou maior ao valor acordado na proposta, por erro no pagamento ou recebimento do valor.
Número decimal

# Borderô
Número identificador de várias propostas.
Número

# Status
Status do processo em sistema que é uma sigla.
Texto

# Tipo de Venda
Descritivo do canal de vendas.

