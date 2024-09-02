Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Call Build.exportarCodigo
End Sub

'Private Sub Workbook_BeforeClose(Cancel As Boolean)
''    Dim resposta As Integer
''
''    ' Perguntar ao usuário se deseja fechar sem salvar
''    resposta = MsgBox("Você deseja fechar sem salvar?", vbYesNo + vbQuestion, "Fechar sem salvar")
''
''    ' Se o usuário escolher 'Sim', permitir o fechamento sem salvar
''    If resposta = vbYes Then
''        Me.Saved = True
''    Else
''        ' Caso contrário, cancelar o fechamento
''        Cancel = True
''    End If
'
''    Me.Saved = True
'End Sub

'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    Dim resposta As Integer
'
'    ' Perguntar ao usuário se deseja fechar sem salvar
'    resposta = MsgBox("Você não tem permissão para salvar esse arquivo! Entre contato com o suporte. Clique em 'SIM' para continuar.", vbYesNo + vbQuestion, "Fechar sem salvar")
'
'    ' Se o usuário escolher 'Sim', permitir o fechamento sem salvar
'    If resposta = vbYes Then
'        Me.Saved = True
'    Else
'        ' Caso contrário, cancelar o fechamento
'        Cancel = True
'    End If
'End Sub

Private Sub Workbook_Open()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Call telaMenuPrincipal
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub