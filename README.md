# Leitura_Produtos_XML
Código para leitura dos produtos de um arquivo XML 

Utilizei as bibliotecas os, tkinter, pandas, xmltodict e tqdm para criar um código que pede para selecionar uma pasta que contenha arquivos com a extensão 
".xml" (utilizei uma janela do "tkinter" e "os" para listar os arquivos) e faça a leitura dos produtos (utilizei o "xmltodict" para transformar o XML em um dicionário)
que contem naquela nota fiscal.

Encontrei algumas diferenças na estrutura dos arquivos XML das notas fiscais e precisei separar o código em 3 funções, 
as duas primeiras para leitura de NFE (Nota Fiscal Eletrônica) e a terceira para leitura de NFCE (Nota Fiscal de Consumidor Eletrônica):
A primeira função lê arquivos XML de NFE solicitados na Receita do Estado do PR.
A segunda função lê os arquivos XML emitidos pelas empresas ou que foram baixados dos sites Portal da NFE (Site do Governo Federal necessita de certificado digital)
ou Fsist (algumas notas não necessitam de certificado digital para baixar o XML).
A terceira função faz a leitura dos arquivos XML de notas fiscais emitidas para consumidor final.

As três funções fazem basicamente a mesma leitura do XML, respeitando as diferenças que os XML's possuem em sua estrutura.
As funções irão retornar uma lista de dicionários com os produtos do XML da nota fiscal lida, a partir dessa lista de dicionários irá criar um DataFrame da biblioteca
pandas e a partir do DF do pandas irá gerar um arquivo Excel com as informações mais relevantes lidas do XML, assim como outro Excel com as informações agrupadas por
número da NF, chave de acesso e CFOP.
