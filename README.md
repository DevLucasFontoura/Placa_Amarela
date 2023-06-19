# Placa Amarela - Script de Cadastramento de Veículos na BIN.

O Placa Amarela é um script desenvolvido para otimizar o processo de cadastramento de veículos na BIN, devido à alta demanda existente.

BIN (Base de dados informatizada e centralizada): A BIN é responsável por armazenar as informações oficiais do DENATRAN (Departamento Nacional de Trânsito).

O script possui as seguintes etapas:

01 - Leitura do Checklist em formato PDF: O script é capaz de ler um arquivo PDF contendo todas as informações necessárias para o cadastramento do veículo. Essas informações são armazenadas em uma lista para posterior utilização.

02 - Consulta de informações: Após a leitura do PDF, o script realiza três tipos de consultas: Consulta do Chassi, Consulta do Motor e Consulta do Chassi na Base Estadual. Cada consulta gera uma imagem que é salva como registro do processo.

03 - Cadastro na BIN: Após a fase de consulta, o script realiza o cadastro do veículo na BIN. Um novo comando é digitado, direcionando para a tela de cadastro, onde a maioria das informações presentes no Checklist (PDF) são utilizadas.

04 - Confirmação e justificativa do cadastro: Após o preenchimento dos dados de cadastro, é realizada a confirmação do mesmo. Além disso, é possível adicionar uma justificativa relacionada ao cadastro.

05 - Verificação do cadastro e anexo de telas: Após a conclusão do cadastro, é realizada uma última consulta para verificar se o veículo foi cadastrado corretamente. Todas as telas capturadas durante o processo são anexadas para registro.

Com a implementação do Placa Amarela, o processo de cadastramento de veículos na BIN se torna mais eficiente e automatizado, agilizando o atendimento à alta demanda existente.

A imagem abaixo ilustra o processo completo no terminal do VSCode. Por questões de privacidade, algumas informações pessoais foram censuradas.

![resumo_placa_amarela2](https://github.com/DevLucasFontoura/Placa_Amarela/assets/129316526/1250a71d-c954-42cd-a2e0-aeb2bee6f922)


