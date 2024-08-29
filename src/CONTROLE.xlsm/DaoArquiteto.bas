Attribute VB_Name = "DaoArquiteto"
Option Explicit

' Atualiza os dados na tela de menu principal
Function listarArquitetosMenuPrincipal()

    Dim rs As ADODB.Recordset
    Dim rsAuxiliar As ADODB.Recordset
    Dim arquiteto As objArquiteto
    Dim contato As objContato
    Dim romaneio As objRomaneio
    Dim listaContatos As Collection
    Dim listaRomaneios As Collection
    Dim strSql As String
    Dim totalVendas As Long
    Dim totalPontos As Long
    Dim linha As Long
    Dim i As Long
    Dim j As Long
    
    '-----------------------------------------------------------------------------------------
    ' Pesquisa os arquitetos -----------------------------------------------------------------
    
    ' String para consulta
    strSql = "SELECT * FROM Arquitetos ORDER BY nome;"
    
    ' Abrindo conexão com banco
    Call conctarBanco
    
    ' Criação e atribuição dos objeto
    Set LISTA_ARQUITETOS = New Collection
    Set rs = New ADODB.Recordset
    
    ' Consulta banco
    rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        Set arquiteto = New objArquiteto
        
        arquiteto.codigo = rs.Fields("pf_codigo").Value
        arquiteto.nome = rs.Fields("nome").Value
        
        If IsNull(rs.Fields("escritorio").Value) Then
            arquiteto.escritorio = ""
        Else
            arquiteto.escritorio = rs.Fields("escritorio").Value
        End If
        
        If IsNull(rs.Fields("aniversario").Value) Then
            arquiteto.aniversario = ""
        Else
            arquiteto.aniversario = rs.Fields("aniversario").Value
        End If
        
        If IsNull(rs.Fields("ultimo_contato").Value) Then
            arquiteto.ultimoContato = ""
        Else
            arquiteto.ultimoContato = rs.Fields("ultimo_contato").Value
        End If
        
        If IsNull(rs.Fields("marmoarias").Value) Then
            arquiteto.marmoarias = ""
        Else
            arquiteto.marmoarias = rs.Fields("marmoarias").Value
        End If
        
        If IsNull(rs.Fields("observacao_geral").Value) Then
            arquiteto.observacaoGeral = ""
        Else
            arquiteto.observacaoGeral = rs.Fields("observacao_geral").Value
        End If
        
        If IsNull(rs.Fields("retorno").Value) Then
            arquiteto.retorno = ""
        Else
            arquiteto.retorno = rs.Fields("retorno").Value
        End If
        
        If IsNull(rs.Fields("pendencia").Value) Then
            arquiteto.pendencia = ""
        Else
            arquiteto.pendencia = rs.Fields("pendencia").Value
        End If
        
        If IsNull(rs.Fields("email").Value) Then
            arquiteto.email = ""
        Else
            arquiteto.email = rs.Fields("email").Value
        End If
        
        If IsNull(rs.Fields("telefone").Value) Then
            arquiteto.telefone = ""
        Else
            arquiteto.telefone = rs.Fields("telefone").Value
        End If
        
        If IsNull(rs.Fields("cpf").Value) Then
            arquiteto.cpf = ""
        Else
            arquiteto.cpf = rs.Fields("cpf").Value
        End If
        
        If IsNull(rs.Fields("cnpj").Value) Then
            arquiteto.cnpj = ""
        Else
            arquiteto.cnpj = rs.Fields("cnpj").Value
        End If
        
        If IsNull(rs.Fields("logradouro").Value) Then
            arquiteto.logradouro = ""
        Else
            arquiteto.logradouro = rs.Fields("logradouro").Value
        End If
        
        If IsNull(rs.Fields("bairro").Value) Then
            arquiteto.bairro = ""
        Else
            arquiteto.bairro = rs.Fields("bairro").Value
        End If
        
        If IsNull(rs.Fields("cidade").Value) Then
            arquiteto.cidade = ""
        Else
            arquiteto.cidade = rs.Fields("cidade").Value
        End If
        
        If IsNull(rs.Fields("uf").Value) Then
            arquiteto.uf = ""
        Else
            arquiteto.uf = rs.Fields("uf").Value
        End If
        
        If IsNull(rs.Fields("cep").Value) Then
            arquiteto.cep = ""
        Else
            arquiteto.cep = rs.Fields("cep").Value
        End If
        
        '-----------------------------------------------------------------------------------------
        ' Pesquisa os contatos -------------------------------------------------------------------

        Set rsAuxiliar = New Recordset
        Set listaContatos = New Collection
        
        ' String para consulta
        strSql = "SELECT * FROM Contatos_Arquiteto" _
                    & " WHERE fk_arquiteto = " & arquiteto.codigo _
                    & " ORDER BY data_contato desc;"
        
        ' Consulta banco
        rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
                
        While Not rsAuxiliar.EOF
            Set contato = New objContato
            
            Set contato.arquiteto = arquiteto
            contato.codigo = rsAuxiliar.Fields("cod_contato")
            
            If IsNull(rsAuxiliar.Fields("data_contato").Value) Then
                contato.dataContato = ""
            Else
                contato.dataContato = rsAuxiliar.Fields("data_contato")
            End If
            
            If IsNull(rsAuxiliar.Fields("relato_contato").Value) Then
                contato.relatoContato = rsAuxiliar.Fields("relato_contato")
            Else
                contato.relatoContato = rsAuxiliar.Fields("relato_contato")
            End If
            
            If IsNull(rsAuxiliar.Fields("data_retorno").Value) Then
                contato.dataRetorno = ""
            Else
                contato.dataRetorno = rsAuxiliar.Fields("data_retorno")
            End If
            
            If IsNull(rsAuxiliar.Fields("observacao").Value) Then
                contato.obsevacao = ""
            Else
                contato.obsevacao = rsAuxiliar.Fields("observacao")
            End If
            
            listaContatos.Add contato
            
            Set contato = Nothing
            
            rsAuxiliar.MoveNext
        Wend
        
        rsAuxiliar.Close
        Set rsAuxiliar = Nothing
        
        '-----------------------------------------------------------------------------------------
        ' Pesquisa os romaneios ------------------------------------------------------------------
        
        Set rsAuxiliar = New Recordset
        Set listaRomaneios = New Collection
        
        ' String para consulta
        strSql = "SELECT * FROM Romaneios_Arquitetos" _
                        & " WHERE fk_arquiteto = " & arquiteto.codigo _
                        & " ORDER BY numero_romaneio;"
        
        ' Consulta banco
        rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
                
        While Not rsAuxiliar.EOF
            Set romaneio = New objRomaneio
            
            Set romaneio.arquiteto = arquiteto
            romaneio.codigo = rsAuxiliar.Fields("cod_romaneio")
            
            If IsNull(rsAuxiliar.Fields("numero_romaneio").Value) Then
                romaneio.numeroRomaneio = ""
            Else
                romaneio.numeroRomaneio = rsAuxiliar.Fields("numero_romaneio")
            End If
            
            If IsNull(rsAuxiliar.Fields("pontuacao").Value) Then
                romaneio.pontuacao = ""
            Else
                romaneio.pontuacao = rsAuxiliar.Fields("pontuacao")
            End If

            listaRomaneios.Add romaneio
            
            Set romaneio = Nothing
            
            rsAuxiliar.MoveNext
        Wend
        
        rsAuxiliar.Close
        Set rsAuxiliar = Nothing
        
        Set arquiteto.listaContatos = listaContatos
        Set arquiteto.listaRomaneios = listaRomaneios
        
        ' Adiciona na lista
        LISTA_ARQUITETOS.Add arquiteto
        
        ' Libera espaço para nova pesquisa se ouver
        Set arquiteto = Nothing
        Set listaContatos = Nothing
        Set listaRomaneios = Nothing
        
        rs.MoveNext
    Wend
    
    ' Libera recurso Recordset
    rs.Close
    
    Set rs = Nothing
    
    ' Fechar conexão com banco
    Call fecharConexaoBanco
    
    '-----------------------------------------------------------------------------------------
    ' Carrega planilha com a lista -----------------------------------------------------------
    
    ' Atribuições
    linha = 6 ' Linha da tabela onde vai começar ser setado os dados
    
    ' Seleciona a planilha
    PlanMenuPrincipal.Select
    
    With PlanMenuPrincipal
        ' Criação e atribuição do objeto
        Set arquiteto = New objArquiteto
        ' Apaga se tiver conteúdo na planilha
        .Range("A8:X1048564").ClearContents
        'Percorre a lista o cola os valores na planilha
        For i = 1 To LISTA_ARQUITETOS.Count
        
            Set arquiteto = LISTA_ARQUITETOS.Item(i)
            totalVendas = 0
            totalPontos = 0
            
            ' Cola os dados
            .Cells(linha, 2).Value = arquiteto.codigo
            .Cells(linha, 3).Value = arquiteto.nome
            .Cells(linha, 4).Value = arquiteto.aniversario
            .Cells(linha, 6).Value = arquiteto.retorno
            .Cells(linha, 7).Value = arquiteto.pendencia
            .Cells(linha, 8).Value = arquiteto.ultimoContato
            
            totalVendas = arquiteto.listaRomaneios.Count
            For j = 1 To arquiteto.listaRomaneios.Count
                Set romaneio = arquiteto.listaRomaneios.Item(j)
                totalPontos = totalPontos + CLng(romaneio.pontuacao)
            Next j
            
            .Cells(linha, 10).Value = totalVendas
            .Cells(linha, 11).Value = totalPontos
            
            linha = linha + 1
            ' Libera espaço
            Set arquiteto = Nothing
        Next i
    End With
End Function

' Pesquisa arquiteto direto na variavel de lista
Function pesquisarAquitetoPorIdNaLista(arquiteto As objArquiteto)
    
    Dim contato As objContato
    Dim romaneio As objRomaneio
    Dim caminhoArquivo As String
    Dim linha As Long
    Dim i As Long
    
    ' Seleciona a planilha
    PlanArquiteto.Select
    
    With PlanArquiteto
    
        .Cells(1, 2).Value = arquiteto.nome
        
        .Cells(5, 10).Value = arquiteto.codigo
        .Cells(5, 3).Value = arquiteto.escritorio
        .Cells(5, 5).Value = arquiteto.aniversario
        .Cells(5, 7).Value = arquiteto.ultimoContato
        
        .Cells(8, 3).Value = arquiteto.logradouro
        .Cells(8, 5).Value = arquiteto.bairro
        .Cells(8, 7).Value = arquiteto.cidade
        .Cells(8, 8).Value = arquiteto.uf
        .Cells(8, 9).Value = arquiteto.cep
        
        .Cells(11, 3).Value = arquiteto.marmoarias
        .Cells(11, 5).Value = arquiteto.observacaoGeral
        .Cells(11, 8).Value = arquiteto.retorno
        .Cells(11, 9).Value = arquiteto.pendencia
        
        .Cells(14, 3).Value = arquiteto.email
        .Cells(14, 5).Value = arquiteto.cnpj
        .Cells(14, 7).Value = arquiteto.cpf
        .Cells(14, 8).Value = arquiteto.telefone
        
        Range("C17").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Range("E17").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        linha = 17
        For i = 1 To arquiteto.listaContatos.Count
            Set contato = arquiteto.listaContatos.Item(i)
            
            .Cells(linha, 2).Value = contato.dataContato
            .Cells(linha, 3).Value = contato.relatoContato
            .Cells(linha, 4).Value = contato.dataRetorno
            .Cells(linha, 5).Value = contato.obsevacao
            .Cells(linha, 6).Value = contato.codigo
            
            linha = linha + 1
        Next i
    
        linha = 17
        For i = 1 To arquiteto.listaRomaneios.Count
            Set romaneio = arquiteto.listaRomaneios.Item(i)
            
            .Cells(linha, 8).Value = romaneio.numeroRomaneio
            .Cells(linha, 9).Value = romaneio.pontuacao
            .Cells(linha, 11).Value = romaneio.codigo
            
            linha = linha + 1
        Next i
        
        caminhoArquivo = ThisWorkbook.Path & "\FOTOS\" & .Cells(5, 10).Value & ".jpg"
        
        If Dir(caminhoArquivo) <> "" Then
            .fotoArquiteto.Picture = LoadPicture(caminhoArquivo)
        Else
            caminhoArquivo = ThisWorkbook.Path & "\FOTOS\" & "0.jpg"
            .fotoArquiteto.Picture = LoadPicture(caminhoArquivo)
        End If
        
        
        Range("A1").Select
    End With
End Function

' Cadastrar e editar Arquitetos
Function cadastrarEditarArquiteto(arquiteto As objArquiteto) As objArquiteto

    Dim rs As ADODB.Recordset
    Dim rsAuxiliar As ADODB.Recordset
    Dim listaContatosBanco As Collection
    Dim listaRomaneiosBanco As Collection
    Dim arquitetoBanco As objArquiteto
    Dim fkObject As Variant
    Dim strSql As String
    Dim campos() As String
    Dim valoresCampos As String
    Dim cadastro As Boolean
    Dim i As Long
    Dim j As Long
    
    ' Seta true em cadastro
    cadastro = True
    
    ' Faz a consulta para saber se o código do bloco já exite
    strSql = "SELECT * FROM Arquitetos WHERE pf_codigo = " & arquiteto.codigo & ";"
    
    ' Abrindo conexão com banco
    Call conctarBanco
    ' Criando e abrindo Recordset para consulta
    Set rs = New Recordset
    Set rsAuxiliar = New Recordset
    ' Abrindo Recordset para consulta
    rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
    ' Retorno da consulta
    While Not rsAuxiliar.EOF
        ' Seta false porquê vai ser uma edição
        cadastro = False
        
        rsAuxiliar.MoveNext
    Wend
    ' Fecha conexão do Recordset
    rsAuxiliar.Close
    
    ' Direciona para os comandos certos de cadastro ou edição
    If cadastro = True Then ' Se cadastro
        
        ' Realoca espaço da variavel
        ReDim campos(1 To 17)
        ' Colocando vingulas, Parenteses e  arpas simples os valores
        campos(1) = "('" & arquiteto.nome & "', "
        campos(2) = "'" & arquiteto.escritorio & "', "
        
        If arquiteto.aniversario = "" Or arquiteto.aniversario = " " Then
            campos(3) = "" & "NULL" & ", "
        Else
            campos(3) = "'" & arquiteto.aniversario & "', "
        End If
        
        If arquiteto.ultimoContato = "" Or arquiteto.ultimoContato = " " Then
            campos(4) = "" & "NULL" & ", "
        Else
            campos(4) = "'" & arquiteto.ultimoContato & "', "
        End If
        
        campos(5) = "'" & arquiteto.marmoarias & "', "
        campos(6) = "'" & arquiteto.observacaoGeral & "', "
        campos(7) = "'" & arquiteto.retorno & "', "
        campos(8) = "'" & arquiteto.pendencia & "', "
        campos(9) = "'" & arquiteto.email & "', "
        campos(10) = "'" & arquiteto.telefone & "', "
        campos(11) = "'" & arquiteto.cpf & "', "
        campos(12) = "'" & arquiteto.cnpj & "', "
        campos(13) = "'" & arquiteto.logradouro & "', "
        campos(14) = "'" & arquiteto.bairro & "', "
        campos(15) = "'" & arquiteto.cidade & "', "
        campos(16) = "'" & arquiteto.uf & "', "
        campos(17) = "'" & arquiteto.cep & "');"
        
        ' Concatenando os valores
        For i = 1 To 17
            valoresCampos = valoresCampos & campos(i)
        Next i
    
        ' Concatenando comando SQL e cadastrando bloco no banco de dados
        strSql = "INSERT INTO Arquitetos (nome, escritorio, aniversario, ultimo_contato, marmoarias, observacao_geral, " _
                    & "retorno, pendencia, email, telefone, cpf, cnpj, logradouro, bairro, cidade, uf, cep) " _
                    & "VALUES " & valoresCampos
        
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
        
        If arquiteto.listaContatos.Count > 0 Then
            ' Cadastra os contatos
            Set listaContatosBanco = DaoContato.cadastrarEEditar(arquiteto.listaContatos, False)
            
            Set arquiteto.listaContatos = listaContatosBanco
        End If
        
        If arquiteto.listaRomaneios.Count > 0 Then
            ' Cadastra os romaneios
            Set listaRomaneiosBanco = DaoRomaneio.cadastrarEEditar(arquiteto.listaRomaneios, False)
            
            Set arquiteto.listaRomaneios = listaRomaneiosBanco
        End If
        
        ' Captura o código que foi gerado no banco de dados
        Set rsAuxiliar = New Recordset
        
        strSql = "SELECT MAX(pf_codigo) AS id FROM Arquitetos"
        
        rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
        
        While Not rsAuxiliar.EOF
            
            arquiteto.codigo = rsAuxiliar.Fields("id").Value
            
            rsAuxiliar.MoveNext
        Wend
        
        ' Seta arquiteto cadastrado
        Set arquitetoBanco = arquiteto
    Else ' Se edição
        
        If arquiteto.aniversario = "" Or arquiteto.aniversario = " " Then
            arquiteto.aniversario = "NULL"
            
        Else
            arquiteto.aniversario = "'" & arquiteto.aniversario & "'"
        End If
        
        If arquiteto.ultimoContato = "" Or arquiteto.ultimoContato = " " Then
            arquiteto.ultimoContato = "NULL"
        Else
            arquiteto.ultimoContato = "'" & arquiteto.ultimoContato & "'"
        End If
        
        ' Edição do bloco com serraria e polideira
        strSql = "UPDATE Arquitetos SET nome = '" & arquiteto.nome & "', " _
                            & "escritorio = '" & arquiteto.escritorio & "', " _
                            & "aniversario = " & arquiteto.aniversario & ", " _
                            & "ultimo_contato = " & arquiteto.ultimoContato & ", " _
                            & "marmoarias = '" & arquiteto.marmoarias & "', " _
                            & "observacao_geral = '" & arquiteto.observacaoGeral & "', " _
                            & "retorno = '" & arquiteto.retorno & "', " _
                            & "pendencia = '" & arquiteto.pendencia & "', " _
                            & "email = '" & arquiteto.email & "', " _
                            & "telefone = '" & arquiteto.telefone & "', " _
                            & "cpf = '" & arquiteto.cpf & "', " _
                            & "cnpj = '" & arquiteto.cnpj & "', " _
                            & "logradouro = '" & arquiteto.logradouro & "', " _
                            & "bairro = '" & arquiteto.bairro & "', " _
                            & "cidade = '" & arquiteto.cidade & "',  " _
                            & "uf = '" & arquiteto.uf & "', " _
                            & "cep = '" & arquiteto.cep & "' " _
                            & "WHERE pf_codigo = " & arquiteto.codigo & ";"
            
        rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
        
        If arquiteto.listaContatos.Count > 0 Then
            ' Editar os contatos
            Set listaContatosBanco = DaoContato.cadastrarEEditar(arquiteto.listaContatos, False)
            
            Set arquiteto.listaContatos = listaContatosBanco
        End If
        
        If arquiteto.listaRomaneios.Count > 0 Then
            ' Editar os romaneios
            Set listaRomaneiosBanco = DaoRomaneio.cadastrarEEditar(arquiteto.listaRomaneios, False)
            
            Set arquiteto.listaRomaneios = listaRomaneiosBanco
        End If
        
        ' Seta arquiteto editado
        Set arquitetoBanco = arquiteto
    End If
    
    ' Libera espaço da memoria
    Set rs = Nothing
    Set rsAuxiliar = Nothing
    'Fechando conexão com banco
    Call fecharConexaoBanco
    ' Retorno arquiteto cadastrado
    Set cadastrarEditarArquiteto = arquitetoBanco
End Function

' Limpa a planilha principal
Function limpaPlanilhaMenuPrincipal()
        
    PlanMenuPrincipal.Select
    
    IF_LIMPAR_MENU_PRINCIPAL = True
    
    '-----------------------------------------------------------------------------------------
    ' Limpa a tabela -------------------------------------------------------------------------
    
    Rows("7:7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    '---------------------f--------------------------------------------------------------------
    ' Limpa primeira linha da tabela ---------------------------------------------------------
    Range("MENU_PRINCIPAL[COD]").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("MENU_PRINCIPAL[NOME]").Select
    Selection.ClearContents
    Range("MENU_PRINCIPAL[ANIVERSÁRIO]").Select
    Selection.ClearContents
    Range("MENU_PRINCIPAL[RETORNO]").Select
    Selection.ClearContents
    Range("MENU_PRINCIPAL[PENDÊNCIA]").Select
    Selection.ClearContents
    Range("MENU_PRINCIPAL[ÚLTIMO CONTATO]").Select
    Selection.ClearContents
    Range("MENU_PRINCIPAL[VENDAS]").Select
    Selection.ClearContents
    Range("MENU_PRINCIPAL[PONTUAÇÃO]").Select
    Selection.ClearContents
    
    Range("A1").Select
    
    ' Limpa memoria
    Set LISTA_ARQUITETOS = Nothing
    IF_LIMPAR_MENU_PRINCIPAL = False
End Function

' Limpa os campos da tela arquiteto
Function limparCamposArquiteto()

    Dim caminhoArquivo As String
    Dim totalLinhas As Long
    Dim i As Long
    
    PlanArquiteto.Select
        
    '-----------------------------------------------------------------------------------------
    ' Limpa as celulas -----------------------------------------------------------------------

    Cells(1, 2).Value = "NOME DO ARQUITETO"   ' NOME
    Cells(5, 3).Value = ""   ' ESCRITORIO
    Cells(5, 5).Value = ""   ' ANIVERSARIO
    Cells(5, 7).Value = ""   ' ULTIMO CONTATO
    Cells(5, 10).Value = "0" ' CODIGO
    
    Cells(8, 3).Value = ""   ' LOGRADOURO
    Cells(8, 5).Value = ""   ' BAIRRO
    Cells(8, 7).Value = ""   ' CIDADE
    Cells(8, 8).Value = ""   ' UF
    Cells(8, 9).Value = ""   ' CEP
    
    Cells(11, 3).Value = ""  ' MARMOARIAS
    Cells(11, 5).Value = ""  ' OBSERVACAO GERAL
    Cells(11, 8).Value = ""  ' RETORNO
    Cells(11, 9).Value = ""  ' PENDENCIA
    
    Cells(14, 3).Value = ""  ' EMAIL
    Cells(14, 5).Value = ""  ' CNPJ
    Cells(14, 7).Value = ""  ' CPF
    Cells(14, 8).Value = ""  ' TELEFONE
    
    '-----------------------------------------------------------------------------------------
    ' Limpa a foto ---------------------------------------------------------------------------
    
    caminhoArquivo = ThisWorkbook.Path & "\FOTOS\" & Cells(5, 10).Value & ".jpg"
    PlanArquiteto.fotoArquiteto.Picture = LoadPicture(caminhoArquivo)
    
    '-----------------------------------------------------------------------------------------
    ' Limpa a tabela de contatos -------------------------------------------------------------
    
    Range("B17").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    Range("CONTATOS[COD CONTATO]").Select
    
    totalLinhas = Selection.Rows.Count
    
    For i = 1 To totalLinhas
        Selection.ListObject.ListRows(1).Delete
    Next i
    
    Cells(17, 6) = "=COD[COD]"
    
    Range("B17").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    '-----------------------------------------------------------------------------------------
    ' Limpa a tabela de romaneios ------------------------------------------------------------
    
    Range("H17").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    Range("VENDA_ARQUITETOS[COD VENDA]").Select
    
    totalLinhas = Selection.Rows.Count
    
    For i = 1 To totalLinhas
        Selection.ListObject.ListRows(1).Delete
    Next i
    
    Range("VENDA_ARQUITETOS[VENDA]").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R17C13"
    
    Cells(17, 11) = "=COD[COD]"
    
    Range("H17").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    Range("B17,D17,F17").Select
    Range("CONTATOS[COD CONTATO]").Activate
    Range("H17,I17,J17,K17").Select
    Range("VENDA_ARQUITETOS[COD VENDA]").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("C17,E17").Select
    Range("CONTATOS[OBSERVAÇÃO]").Activate
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("A1").Select
End Function
