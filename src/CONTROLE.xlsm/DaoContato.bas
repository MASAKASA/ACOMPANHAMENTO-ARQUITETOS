Attribute VB_Name = "DaoContato"
Option Explicit

' Cadastra e edita objeto
Function cadastrarEEditar(lista As Collection, fecharConexao As Boolean) As Collection
    
    Dim rs As ADODB.Recordset
    Dim rsAuxiliar As ADODB.Recordset
    Dim listaContatosBanco As Collection
    Dim contato As objContato
    Dim strSql As String
    Dim campos() As String
    Dim valoresCampos As String
    Dim cadastro As Boolean
    Dim i As Long
    Dim j As Long

    ' Seta true em cadastro
    cadastro = True
    
    Set listaContatosBanco = New Collection
    
    ' Abrindo conexão com banco
    Call conctarBanco
    
    ' Loop através dos itens da coleção
    For i = 1 To lista.Count
    
            ' Criando e abrindo Recordset para consulta
        Set rs = New Recordset
        Set rsAuxiliar = New Recordset
        
        ' Seta o ojeto
        Set contato = lista(i)
        
        ' String para consulta
        strSql = "SELECT * FROM Contatos_Arquiteto " & "WHERE cod_contato = " & contato.codigo & "" _
                & " AND cod_contato = " & contato.arquiteto.codigo & ";"
        
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
            ReDim campos(1 To 5)
            
            If contato.dataContato = "" Or contato.dataContato = " " Then
                campos(2) = "" & "NULL" & ", "
            Else
                campos(2) = "'" & contato.dataContato & "', "
            End If
            
            If contato.dataRetorno = "" Or contato.dataRetorno = " " Then
                campos(4) = "" & "NULL" & ", "
            Else
                campos(4) = "'" & contato.dataRetorno & "', "
            End If
            
            ' Colocando vingulas, Parenteses e  arpas simples os valores
            campos(1) = "(" & contato.arquiteto.codigo & ", "
            campos(3) = "'" & contato.relatoContato & "', "
            campos(5) = "'" & contato.obsevacao & "');"
            
            ' Concatenando os valores
            For j = 1 To 5
                valoresCampos = valoresCampos & campos(j)
            Next j
            
            ' Concatenando comando SQL e cadastrando bloco no banco de dados
            strSql = "INSERT INTO Contatos_Arquiteto ( fk_arquiteto, data_contato, relato_contato, data_retorno, observacao ) " _
                        & "VALUES " & valoresCampos
                    
            rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
            
            ' Limpa variavel para proximo cadastro
            valoresCampos = ""
            
            
            ' Captura o código que foi gerado no banco de dados
            Set rsAuxiliar = New Recordset

            strSql = "SELECT MAX(cod_contato) AS id FROM Contatos_Arquiteto"

            rsAuxiliar.Open strSql, CONEXAO_BD, adOpenKeyset, adLockReadOnly

            While Not rsAuxiliar.EOF

                contato.codigo = rsAuxiliar.Fields("id").Value

                rsAuxiliar.MoveNext
            Wend
            
            listaContatosBanco.Add contato
        Else
            ' Se edição
            If contato.dataContato = "" Or contato.dataContato = " " Then
                contato.dataContato = "NULL"
            Else
                contato.dataContato = "'" & contato.dataContato & "'"
            End If
            
            If contato.dataRetorno = "" Or contato.dataRetorno = " " Then
                contato.dataRetorno = "NULL"
            Else
                contato.dataRetorno = "'" & contato.dataRetorno & "'"
            End If
            
            strSql = "UPDATE Contatos_Arquiteto SET " _
                & "fk_arquiteto = " & contato.arquiteto.codigo & ", " _
                & "data_contato = " & contato.dataContato & ", " _
                & "relato_contato = '" & contato.relatoContato & "', " _
                & "data_retorno = " & contato.dataRetorno & ", " _
                & "observacao = '" & contato.obsevacao & "', " _
                & " WHERE cod_contato = " & contato.codigo & ";"
                        
            ' Abrindo Recordset para consulta
            rs.Open strSql, CONEXAO_BD, adOpenKeyset, adLockPessimistic
            ' Retorna o valor
            cadastro = True
            
            listaContatosBanco.Add contato
        End If
        
            ' Libera recurso Recordset
            rsAuxiliar.Close
            Set contato = Nothing
            Set rsAuxiliar = Nothing
            Set rs = Nothing
    Next i
    
    ' Fecha a conexão se não for pesquisa de chapa quem chamou esse metodo
    If fecharConexao = True Then
        ' Fechar conexão com banco
        Call fecharConexaoBanco
    End If
    
    Set cadastrarEEditar = listaContatosBanco
End Function

