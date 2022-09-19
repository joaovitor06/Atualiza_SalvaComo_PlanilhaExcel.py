# Menu_atualiza_envia_email_relatorio_excel-em-PYTHON

Interface gráfica que atualiza arquivo em Excel, envia e-mail e importa histórico no SQL SERVER.

main.py: 

-Cria janela que redireciona para as demais janelas.

relatorios_carrefour.py/ relatorios_mapfre.py/ relatorios_unisa.py: 

-Define o diretório raíz da carteira.

-Cria a condição de considerar o mês anterior como referência caso seja dia 1.

-Cria 1 ou mais dicionários (dependendo de quantos relatórios contenham na carteira) contendo informações que vão ser utilizadas na localização, atualização do arquivo, envio de e-mail e histórico.

-Cria uma janela contendo o nome do relatório, botão de atualização e de enviar e-mail, além de dois campos de texto ocultos que após a conclusão da função de atualização, um desses campos é preenchido com o tempo de execução, já quando a função de enviar e-mail é concluída, o outro campo de texto é preenchido informado que o e-mail foi enviado, utilizando com referência o assunto do e-mail.

-Cria as condições de retorno em relação a cada clique na janela. Em caso do botão de atualização ser acionado, ele irá construir o diretório apartir das variáveis do dicionário; depois irá verificar qual o mês referência do diretório (criar se necessário) e localizar o arquivo, após isso atualizar o Excell e calcular o tempo de execução. Em caso do botão de envio de e-mail ser acionado, seguirá o mesmo procedimento de localização do diretorio e arquivo (será o anexo do e-mail), após isso irá enviar o e-mail utilizando algumas variáveis do dicionário como destinatários, cópia, assunto, anexo e informações do corpo do e-mail.
verificador_pasta_mes_vigente.py: 

-Irá verificar se o diretorio existe, caso não exista, irá criá-lo, então irá verificar se o arquivo existe no diretório, caso não exista, irá copiá-lo do diretório do mês anterior, renomeando o arquivo na pasta vigente.
pega_arquivo.py:

-Módulo importado verificador_pasta_mes_vigente.py para realizar a etapa de buscar/copiar/colar e renomear o arquivo do mês anterior caso seja necessário.
atualizador_de_arquivo.py: 

-Irá realizar a atualização do arquivo e salva-lo, após a conclusão, irá realizar a importação do histórico no banco de dados informando o diretorio, data e hora da atualização e uma string 'atualização', finalizando retornando o tempo de execução.
disparador_de_email.py:

-Irá realizar um condicional para a saudação (bom dia, boa tarde e boa noite) utilizando como referência a hora do dia.

-Cria o e-mail utilizando a variável de saudação, além das variáveis importadas do respectivo dicionário, incluindo o arquivo anexo.

-Após isso irá realizar a importação no banco de dados da mesma forma realizada no atualizador, porém mudando a string para 'email'.
historico_de_eventos.py:

-Realiza a conexão com o SQL SERVER.

-Realiza o insert na tabela.
