VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'Classe StringSql v1.0
'Responsavel pela formacao de string sql para operacoes basicas.
'Os métodos devem ser usados na ordem em que aparecem para que o resultado seja correto.
'---------------------------------------------------------------------------------------

Private St As String
Private Condicao As String
Private contCondicao As Integer
Private contJoins As Integer
Private contParent As Integer

'1. Inicio da formacao do string de busca, com a indicacao da tabela
Public Sub sqlSelect(tabela As String)
    St = "SELECT * FROM " + tabela
    contCondicao = 0
    contJoins = 0
    contParent = 0
End Sub

'1. Inicio da formacao do string de insercao, com a indicacao da tabela
Public Sub sqlInsert(tabela As String)
    St = "INSERT INTO " + tabela
End Sub

'1. Campos da tabela para Insert
'O Parâmetro deve ser um array com a quantidade desejada de campos
Public Sub insertCampos(arrCampos As Variant)
    Dim campos As String
    For i = 0 To UBound(arrCampos)
        If i > 0 And i < UBound(arrCampos) + 1 Then
            campos = campos + ","
            campos = campos + arrCampos(i)
        Else
            campos = campos + "(" + arrCampos(i)
        End If
    Next i
    campos = campos + ")"
    St = St + campos
End Sub

'1. Inicio da formacao do string de atualização, com a indicacao da tabela
Public Sub sqlUpdate(tabela As String)
    St = "UPDATE " + tabela + " SET "
End Sub

'1. Campos e valores para atualização
Public Sub updateCampos(camposValores As String)
    St = St + camposValores
End Sub

'1. Valores dos campos para Insert
'Parâmetro deve ser um array com a quantidade desejada de valores
Public Sub sqlValores(arrValores As Variant)
    Dim valores As String
    valores = " VALUES "
    For i = 0 To UBound(arrValores)
        If i > 0 And i < UBound(arrValores) + 1 Then
            valores = valores + ","
            valores = valores + arrValores(i)
        Else
            valores = valores + "(" + arrValores(i)
        End If
    Next i
    valores = valores + ")"
    St = St + valores
End Sub

'1. Inicio da formacao do string de exclusão, com a indicacao da tabela
Public Sub sqlDelete(tabela As String)
    St = "DELETE FROM " + tabela
End Sub

'2. INNER JOIN - com opcao de alias
Public Sub sqlJoin(tabela2 As String, id1 As String, id2 As String, Optional alias1 As String, Optional alias2 As String)
    Dim al1 As String, al2 As String
    If St = "" Then
        Exit Sub
    Else
        If alias1 = "" Then
            al1 = ""
        Else
            al1 = " AS " + alias1
        End If
        If alias2 = "" Then
            al2 = ""
        Else
            al2 = " AS " + alias2
        End If
        St = St + al1 + " INNER JOIN " + tabela2 + al2 + " ON " + id1 + "=" + id2
    End If
End Sub

'2. INNER JOIN - com opcao de alias, para sintaxe do Ms Access
Public Sub sqlJoinAccess(tabela2 As String, id1 As String, id2 As String, Optional alias1 As String, Optional alias2 As String)
    Dim al1 As String, al2 As String, stAux As String, stAux2 As String
    If St = "" Then
        Exit Sub
    Else
        If alias1 = "" Then
            al1 = ""
        Else
            al1 = " AS " + alias1
        End If
        If alias2 = "" Then
            al2 = ""
        Else
            al2 = " AS " + alias2
        End If
        St = St + al1 + " INNER JOIN " + tabela2 + al2 + " ON " + id1 + "=" + id2 + ")"
        contJoins = contJoins + 1
    End If
    For i = 1 To contJoins
        stAux = stAux + "("
    Next i
    If contParent > 0 Then
        For i = 1 To contParent
            stAux2 = stAux2 + "("
        Next i
    End If
    St = Replace(St, "FROM" + stAux2, "FROM" + stAux)
    contParent = contParent + 1
End Sub

'3. Clausula WHERE com suas condicoes
    'cond -> condicao a ser incluida
    'op   -> operador(AND, OR...)
Public Sub sqlWhere(cond As String, op As String)
    If St = "" Then
        Exit Sub
    Else
        If Condicao = "" Or contCondicao = 0 Then
            Condicao = cond
            St = St + " WHERE " + Condicao
            contCondicao = contCondicao + 1
        Else
            Condicao = " " + op + " " + cond
            St = St + Condicao
            contCondicao = contCondicao + 1
        End If
    End If
End Sub

'4. Retorna o string SQL formado
Public Function getString() As String
    If St <> "" Then
        getString = St
    Else
        getString = ""
    End If
End Function



