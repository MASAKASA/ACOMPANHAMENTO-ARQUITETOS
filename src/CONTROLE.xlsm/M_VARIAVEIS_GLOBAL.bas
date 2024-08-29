Attribute VB_Name = "M_VARIAVEIS_GLOBAL"
Option Explicit

' Variaveis para manipulação de banco de dados
Global CONEXAO_BD As New ADODB.Connection
Global LISTA_ARQUITETOS As Collection
Global COD_SELECIONADO As String
Global IF_SELECAO As Boolean
Global IF_SELECAO_II As Boolean
Global IF_LIMPAR_MENU_PRINCIPAL As Boolean
Global COLUNA_SELECIONADA As Long
Global LINHA_SELECIONADA As Long
