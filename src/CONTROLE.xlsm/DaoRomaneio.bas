Attribute VB_Name = "DaoRomaneio"
Option Explicit

' Cadastra e edita objeto
Function cadastrarEEditar(lista As Collection, fecharConexao As Boolean) As Collection
    
    Dim rs As ADODB.Recordset
    Dim rsAuxiliar As ADODB.Recordset
    Dim listaRomaneiosBanco As Collection
    Dim romaneio As objRomaneio
    Dim strSql As String
    Dim campos() As String
    Dim valoresCampos As String
    Dim cadastro As Boolean
    Dim i As Long
    Dim j As Long

    ' Seta true em cadastro
    cadastro = True
    
    Set listaRomaneiosBanco = New Collection
    
    ' Abrindo conexão com banco
    Call conctarBanco
    
    ' Loop através dos itens da coleção
    For i = 1 To lista.Count
    
            ' Criando e abrindo Recordset para consulta
        Set rs = New Recordset
        Set rsAuxiliar = New Recordset
        
        ' Seta o ojeto
        Set romaneio = lista(i)
        
        ' String para consulta
        strSql = "SELECT * FROM Romaneios_Arquitetos " & "WHERE cod_romaneio = " & romaneio.codigo & "" _
                & " AND fk_arquiteto = " & romaneio.arquiteto.codigo & ";"
        
        ' Consulta banco
        rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly
        
        ' Retorno da consulta
        While Not rsAuxiliar.EOF
            ' Irá ser um edição
            cadastro = False
            rsAuxiliar.MoveNext
        Wend
        
        Set rsAuxiliar = Nothing
        
        ' Direciona para os comandos certos de cadastro ou edição
        If cadastro = True Then ' Se cadastro
            ' Realoca espaço da variavel
            ReDim campos(1 To 3)
            ' Colocando vingulas, Parenteses e  arpas simples os valores
            campos(1) = "(" & romaneio.arquiteto.codigo & ", "
            campos(2) = "'" & romaneio.numeroRomaneio & "', "
            campos(3) = "'" & romaneio.pontuacao & "');"
            
            ' Concatenando os valores
            For j = 1 To 3
                valoresCampos = valoresCampos & campos(j)
            Next j
            
            ' Concatenando comando SQL e cadastrando bloco no banco de dados
            strSql = "INSERT INTO Romaneios_Arquitetos ( fk_arquiteto, numero_romaneio, pontuacao ) " _
                        & "VALUES " & valoresCampos
                    
            rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
            
            ' Limpa variavel para proximo cadastro
            valoresCampos = ""
            
            ' Captura o código que foi gerado no banco de dados
            Set rsAuxiliar = New Recordset

            strSql = "SELECT MAX(cod_romaneio) AS id FROM Romaneios_Arquitetos"

            rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly

            While Not rsAuxiliar.EOF

                romaneio.codigo = rsAuxiliar.Fields("id").Value

                rsAuxiliar.MoveNext
            Wend
            
            listaRomaneiosBanco.Add romaneio
        Else
            ' Se edição
            
            strSql = "UPDATE Romaneios_Arquitetos SET " _
                & "fk_arquiteto = " & romaneio.arquiteto.codigo & ", " _
                & "numero_romaneio = '" & romaneio.numeroRomaneio & "', " _
                & "pontuacao = '" & romaneio.pontuacao & "' " _
                & " WHERE cod_romaneio = " & romaneio.codigo & ";"
                        
            ' Abrindo Recordset para consulta
            rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
            ' Retorna o valor
            cadastro = True
            
            listaRomaneiosBanco.Add romaneio
        End If
        
            ' Libera recurso Recordset
            Set romaneio = Nothing
            Set rsAuxiliar = Nothing
            Set rs = Nothing
    Next i
    
    ' Fecha a conexão se não for pesquisa de chapa quem chamou esse metodo
    If fecharConexao = True Then
        ' Fechar conexão com banco
        Call fecharConexaoBanco
    End If
    
    Set cadastrarEEditar = listaRomaneiosBanco
End Function




