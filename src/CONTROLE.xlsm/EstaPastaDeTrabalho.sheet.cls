Option Explicit


Private Sub Workbook_Open()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Call telaMenuPrincipal
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub