Attribute VB_Name = "M_AUX_MENU_PRINCIPAL"
Option Explicit

Public Sub tiraSelecao()
Attribute tiraSelecao.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim linha As String
    Dim coluna As String
    Dim celula As String
    
    ' Condição para não trocar cor das linhas
    IF_SELECAO = True
    IF_SELECAO_II = True
    
    If LISTA_ARQUITETOS.Count > 0 Then
        ' Troca todos os nomes para 'selecionar'
        Sheets("MENU PRINCIPAL").Select
        Range("L6").Select
        ActiveCell.FormulaR1C1 = "=SELECAO[SELEÇÃO]"
        Selection.AutoFill Destination:=Range("MENU_PRINCIPAL[SELEÇÃO]")
        Range("MENU_PRINCIPAL[SELEÇÃO]").Select
    End If
    
    ' Captura a linha e coluna clicada
    linha = CStr(LINHA_SELECIONADA)
    
    If COLUNA_SELECIONADA = 12 Then
        coluna = "L"
    Else
        coluna = "A"
    End If
    
    celula = coluna & linha
    
    ' Condição para trocar cor das linhas
    IF_SELECAO = False
    
    Range(celula).Select

End Sub
