# CadastroDeMateriais
Automação para realizar o cadastro de materiais no SAP.

## Objetivo
- Automatizar tarefas de entrada de dados no SAP Gui.

## Softwares Utilizados
- SAP Gui
- Microsoft Excel

## Linguagens de Programação Utilizadas
- Gui Scripting
- Visual Basic for Application (VBA)

# Como o código funciona?
O objetivo do código é criar uma ligação em um primeiro momento utilizando o SAP Scripting dentro do VBA, logo no início do código. </br>
A partir daqui ocorre uma mescla de VBA e SAP Scipting.

### O SAP Scripting
Responsável por servir como orientação/ localização dos inputs presentes dentro da transação de criação de materiais (MM01 - Criação de Materiais). Irá servir para passagem entre guias, retroceder etc. </br></br>
### Visual Basic for Application
É o responsável for realizar tratamentos de dados, seleção de células dentro de uma tabela. Além disso, poderá receber (extração) dados ou enviar (inserção) no SAP. </br></br>

## Como o VBA Atua para Este Caso?
O código VBA irá ser responsável por criar condições e laços de repetição para cada linha presente em uma planilha, sendo cada linha um novo item.
- Estabelece conexões com os objetos Application, connection, e session. Conecta esses objetos ao objeto WScript.
- Verifica se há valores na tabela para serem enseridos, caso não existam, saíra da Sub.
- Caso existam, é acessada a trasação MM01.
- Serão selecionadas as visões (guias) que serão criadas para aquele tipo de material e o tipo do material.
- Definirá depósitos.
- Exportará o código para o cadastro do SAP Gui para a tabela Excel junto com outras informações.
- O código VBA define uma célula para ser fixada como "target", a partir dessa célula é utilizado o comando "Offset" para a partir do target, selecionar uma a uma das informações do cadastro para cada linha, ao final de cada linha, a célula fixa se move para a próxima abaixo até que nenhum valor seja encontrado.
- Durante o processo mencionado acima, o código irá passar por uma sequência de condições e loopings. Para cada tipo de material, um escopo de código é acessado, assim diferenciando os campos de cada visão que este tipo de item irá possuir.




  
