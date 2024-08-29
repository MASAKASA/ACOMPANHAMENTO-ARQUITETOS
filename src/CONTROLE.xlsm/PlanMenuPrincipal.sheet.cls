Option Explicit

' Observa as mudacas na planilha
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If IF_LIMPAR_MENU_PRINCIPAL = False Then
        If IF_SELECAO = False And IF_SELECAO_II = False Then
            
            LINHA_SELECIONADA = Target.Row
            COLUNA_SELECIONADA = Target.Column
        
            COD_SELECIONADO = Cells(LINHA_SELECIONADA, 2)
            
            If COD_SELECIONADO <> "" And LINHA_SELECIONADA > 5 _
                        And COLUNA_SELECIONADA = 12 Then
                
                Call tiraSelecao
            Else
                Range("A1").Select
            End If
        End If
        
        If IF_SELECAO = False Then
            If COD_SELECIONADO <> "" And Cells(LINHA_SELECIONADA, 12) <> "SELECIONADO" Then
        
                Cells(LINHA_SELECIONADA, 12) = "SELECIONADO"
            End If
            
            IF_SELECAO = False
            IF_SELECAO_II = False
        End If
    End If
End Sub

