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
'Classe StringSql
'Responsavel pela formacao de string sql para operacoes basicas.
'Os métodos devem ser usados na ordem em que aparecem para que o resultado seja correto.
'---------------------------------------------------------------------------------------

Private St As String
Private Condicao As String

'1. Inicio da formacao do string de busca, com a indicacao da tabela
Sub sqlSelect(tabela As String)
    St = "SELECT * FROM " + tabela
End Sub

'1. Inicio da formacao do string de insercao, com a indicacao da tabela
Sub sqlInsert(tabela As String)
    St = "INSERT INTO " + tabela
End Sub

Sub sqlCampos(campo1 As String, Optional campo2 As String, Optional campo3 As String, Optional campo4 As String, Optional campo5 As String)
    Dim campos As String
    campos = "(" + campo1
    If campo2 <> "" Then
        campos = campos + "," + campo2
    End If
End Sub

'2. INNER JOIN - com opcao de alias
Public Function sqlJoin(tabela2 As String, id1 As String, id2 As String, Optional alias1 As String, Optional alias2 As String)
    Dim al1 As String, al2 As String
    If St = "" Then
        Exit Function
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
End Function

'3. Clausula WHERE com suas condicoes
    'cond -> condicao a ser incluida
    'op   -> operador(AND, OR...)
Public Function sqlWhere(cond As String, op As String)
    If St = "" Then
        Exit Function
    Else
        If Condicao = "" Then
            Condicao = cond
            St = St + " WHERE " + Condicao
        Else
            Condicao = " " + op + " " + cond
            St = St + Condicao
        End If
    End If
End Function

'4. Retorna o string SQL formado
Public Function getString() As String
    If St <> "" Then
        getString = St
    Else
        getString = ""
    End If
End Function



