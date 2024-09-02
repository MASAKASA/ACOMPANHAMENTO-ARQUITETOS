Attribute VB_Name = "Botoes"
Option Explicit

Public Sub carregarListaArquitetos()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Call DaoArquiteto.limpaPlanilhaMenuPrincipal
    
    Call DaoArquiteto.listarArquitetosMenuPrincipal
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub telaMenuPrincipal()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Call DaoArquiteto.limparCamposArquiteto
    
    Call DaoArquiteto.limpaPlanilhaMenuPrincipal
    
    Call DaoArquiteto.listarArquitetosMenuPrincipal
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub telaArquiteto()
    
    Dim i As Long
    Dim arquiteto As objArquiteto
    
    If LINHA_SELECIONADA = 0 Then
        MsgBox "Selecione um arquiteto da lista!", vbCritical, "Nada selecioando"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    For i = 1 To LISTA_ARQUITETOS.Count
    
        Set arquiteto = LISTA_ARQUITETOS.Item(i)
        
        If COD_SELECIONADO = arquiteto.codigo Then
            Call DaoArquiteto.pesquisarAquitetoPorIdNaLista(arquiteto)
            Exit For
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub salvarArquiteto()
    
    Dim ws As Worksheet
    Dim arquiteto As objArquiteto
    Dim arquitetoJaCadastrado As objArquiteto
    Dim arquitetoBanco As objArquiteto
    Dim contato As objContato
    Dim romaneio As objRomaneio
    Dim listaContatos As Collection
    Dim listaRomaneios As Collection
    Dim caminhoArquivo As String
    Dim linha As Long
    Dim ultimaLinha As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    Set arquiteto = New objArquiteto
    Set listaContatos = New Collection
    Set listaRomaneios = New Collection
    linha = 17
    
    With PlanArquiteto
         
        arquiteto.nome = UCase(.Cells(1, 2).Value)
'        For i = 1 To LISTA_ARQUITETOS.Count
'            Set arquitetoJaCadastrado = LISTA_ARQUITETOS.Item(i)
'
'            If arquiteto.nome = arquitetoJaCadastrado.nome Then
'                MsgBox "Esse nome do arquiteto já cadastrado no sistema!", vbCritical, "Cadastrado arquiteto"
'                Exit Sub
'            End If
'        Next i
        
        If arquiteto.nome = "NOME DO ARQUITETO" Or arquiteto.nome = "" _
                Or arquiteto.nome = " " Then
                        
           MsgBox "Adicione o nome do arquiteto!", vbCritical, "Arquiteto sem nome"
           Exit Sub
        End If
         
        If IsDate(.Cells(5, 5).Value) = False Then
            If .Cells(5, 5).Value = "" Or .Cells(5, 5).Value = " " Then
                arquiteto.aniversario = UCase(.Cells(5, 5).Value)
            Else
                MsgBox "Adicione uma data válida para aniversário!", vbCritical, "Data inválida"
                Exit Sub
            End If
        Else
            arquiteto.aniversario = UCase(.Cells(5, 5).Value)
        End If
        
        If IsDate(.Cells(5, 7).Value) = False Then
            If .Cells(5, 7).Value = "" Or .Cells(5, 5).Value = " " Then
                arquiteto.ultimoContato = UCase(.Cells(5, 7).Value)
            Else
                MsgBox "Adicione uma data válida para último contato!", vbCritical, "Data invlida"
                Exit Sub
            End If
        Else
            arquiteto.ultimoContato = UCase(.Cells(5, 7).Value)
        End If
         
        If .Cells(11, 8).Value = "" Or .Cells(11, 8).Value = " " Then
           MsgBox "Adicione um status para 'retorno' ao arquiteto!", vbCritical, "Sem status para retorno"
           Exit Sub
        End If
        
        If .Cells(11, 9).Value = "" Or .Cells(11, 9).Value = " " Then
           MsgBox "Adicione um status para 'pendência' ao arquiteto!", vbCritical, "Sem status para pencêcia"
           Exit Sub
        End If
        
        If ValidarEmail(.Cells(14, 3).Value) = False And .Cells(14, 3).Value <> "" Then
            MsgBox "Adicione um e-mail válido!", vbCritical, "E-mail inválido"
           Exit Sub
        End If
        
        arquiteto.codigo = UCase(.Cells(5, 10).Value)
        arquiteto.escritorio = UCase(.Cells(5, 3).Value)
        arquiteto.marmoarias = UCase(.Cells(11, 3).Value)
        arquiteto.observacaoGeral = UCase(.Cells(11, 5).Value)
        arquiteto.retorno = UCase(.Cells(11, 8).Value)
        arquiteto.pendencia = UCase(.Cells(11, 9).Value)
        arquiteto.email = UCase(.Cells(14, 3).Value)
        arquiteto.cnpj = UCase(.Cells(14, 5).Value)
        arquiteto.cpf = UCase(.Cells(14, 7).Value)
        arquiteto.telefone = UCase(.Cells(14, 8).Value)
        arquiteto.logradouro = UCase(.Cells(8, 3).Value)
        arquiteto.bairro = UCase(.Cells(8, 5).Value)
        arquiteto.cidade = UCase(.Cells(8, 7).Value)
        arquiteto.uf = UCase(.Cells(8, 8).Value)
        arquiteto.cep = UCase(.Cells(8, 9).Value)
        
        ultimaLinha = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
        For i = linha To ultimaLinha
           Set contato = New objContato
           
            If IsDate(.Cells(linha, 2).Value) = False Then
                If .Cells(linha, 2).Value = "" Or .Cells(linha, 2).Value = " " Then
                    contato.dataContato = UCase(.Cells(linha, 2).Value)
                Else
                    MsgBox "Adicione uma data válida em 'data contato' na linha: " & linha & "!", vbCritical, "Data inválida"
                    Exit Sub
                End If
            Else
                contato.dataContato = UCase(.Cells(linha, 2).Value)
            End If
            
            If IsDate(.Cells(linha, 4).Value) = False Then
                If .Cells(linha, 4).Value = "" Or .Cells(linha, 4).Value = " " Then
                    contato.dataRetorno = UCase(.Cells(linha, 4).Value)
                Else
                    MsgBox "Adicione uma data válida em 'data retorno' na linha: " & linha & "!", vbCritical, "Data inválida"
                    Exit Sub
                End If
            Else
                contato.dataRetorno = UCase(.Cells(linha, 4).Value)
            End If
            
            contato.codigo = UCase(.Cells(linha, 6).Value)
            contato.relatoContato = UCase(.Cells(linha, 3).Value)
            contato.obsevacao = UCase(.Cells(linha, 5).Value)
            Set contato.arquiteto = arquiteto
            
            If contato.codigo <> "0" And contato.dataContato <> "" Then
                listaContatos.Add contato
            ElseIf contato.codigo = "0" And contato.dataContato <> "" Then
                listaContatos.Add contato
            End If
            
            Set contato = Nothing
            linha = linha + 1
        Next i
        
        linha = 17
        ultimaLinha = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
        For i = linha To ultimaLinha
            Set romaneio = New objRomaneio
            
            romaneio.codigo = UCase(.Cells(linha, 11).Value)
            romaneio.numeroRomaneio = UCase(.Cells(linha, 8).Value)
            romaneio.pontuacao = UCase(.Cells(linha, 9).Value)
           
            If romaneio.pontuacao <> "" Or romaneio.pontuacao <> " " Then
                If romaneio.numeroRomaneio = "" Or romaneio.numeroRomaneio = " " Then
                    MsgBox "Adicione um número de romaneio para pontuação na linha: " & linha & "!", vbCritical, "Número de romaneio inválido"
                    Exit Sub
                End If
            End If
            
            Set romaneio.arquiteto = arquiteto
            
            If romaneio.codigo <> "0" And romaneio.numeroRomaneio <> "" Then
                listaRomaneios.Add romaneio
            ElseIf romaneio.codigo = "0" And romaneio.numeroRomaneio <> "" Then
                listaRomaneios.Add romaneio
            End If
            
            Set romaneio = Nothing
            linha = linha + 1
        Next i
    End With
    
    Set arquiteto.listaContatos = listaContatos
    Set arquiteto.listaRomaneios = listaRomaneios
    
    Set arquitetoBanco = DaoArquiteto.cadastrarEditarArquiteto(arquiteto)
    
    ' Verifica se o usuario tem foto de perfil
    caminhoArquivo = ThisWorkbook.Path & "\FOTOS\" & arquitetoBanco.codigo & ".jpg"
    If Dir(caminhoArquivo) <> "" Then
        ' Já existe foto de perfil
    Else
    
        With PlanArquiteto
            SavePicture .fotoArquiteto.Picture, caminhoArquivo
        End With
    
    End If
    
    Call DaoArquiteto.pesquisarAquitetoPorIdNaLista(arquitetoBanco)
    
    For i = 1 To LISTA_ARQUITETOS.Count
        Set arquiteto = LISTA_ARQUITETOS.Item(i)
        
        If arquiteto.codigo = arquitetoBanco.codigo Then
            Set arquiteto = arquitetoBanco
            Exit For
        End If
    Next i
    
    MsgBox "Alteração feita com sucesso!", vbInformation, "Alteração realizada"
End Sub

Public Sub trocarFotoArquiteto()
    
On Error Resume Next
    
    Dim caminhoArquivo As String
    Dim extensaoArquivo As String
    
    MsgBox "Só fotos com extesão '.jpg' são aceitas para foto de perfil!", vbInformation, "Tipo de foto"
    
    caminhoArquivo = Application.GetOpenFilename( _
                    FileFilter:="Image Files(*.jpg),*jpg")
    
    If caminhoArquivo = "Falso" Then
        Exit Sub
    End If
    
    extensaoArquivo = Mid(caminhoArquivo, InStrRev(caminhoArquivo, ".") + 1)
    
    If extensaoArquivo <> "jpg" Then
        MsgBox "Só fotos com extesão '.jpg' são aceitas para foto de perfil!", vbInformation, "Tipo de foto"
        Exit Sub
    End If
    
    With PlanArquiteto
        .fotoArquiteto.Picture = LoadPicture("")
        .fotoArquiteto.Picture = LoadPicture(caminhoArquivo)
    End With
End Sub

Public Sub salvarFotoArquiteto()
  
    Dim caminhoArquivo As String
    
    If PlanArquiteto.Cells(5, 10).Value = "0" Then
        MsgBox "Primeiro salve cadastro de arquiteto pra depois inserir uma foto!", vbCritical, "Primeiro salvar arquiteto"
        Exit Sub
    End If
        
    caminhoArquivo = ThisWorkbook.Path & "\FOTOS\" & PlanArquiteto.Cells(5, 10).Value & ".jpg"
    
    With PlanArquiteto
        SavePicture .fotoArquiteto.Picture, caminhoArquivo
    End With
    
    MsgBox "Foto salva com sucesso!", vbInformation, "Alteração realizada"
End Sub

Function ValidarEmail(ByVal email As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Padrão de expressão regular para validar o e-mail
    regex.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    regex.IgnoreCase = True
    regex.Global = False
    
    ' Verifica se o e-mail é válido
    If regex.Test(email) Then
        ValidarEmail = True
    Else
        ValidarEmail = False
    End If
    
    Set regex = Nothing
End Function
