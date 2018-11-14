VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailCDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'Classe EmailCDO versão 1.0.0
'Envio de email através da biblioteca Microsoft CDO
'Uso:
    '-> Dim email As New email
    '-> email.de = "dnl@mail.caixa"
    '-> email.para = "dnl@mail.caixa"
    '-> email.cc = "dnl@mail.com"
    '-> email.assunto = "Teste"
    '-> email.texto = "<h1>Teste de email CDO</h1>"
    '-> email.enviar
'---------------------------------------------------------------------------------------

Public de As String
Public para As String
Public cc As String
Public bcc As String
Public assunto As String
Public texto As String
Public anexo As String

Private iConf As New CDO.Configuration
Private iMsg As New CDO.Message
Private fld As Object

Public Function enviar() As Boolean

    On Error GoTo Err

    setConfig
    With iMsg
        Set .Configuration = iConf
        .From = de
        .To = para
        .Subject = assunto
        .HTMLBody = texto
        .cc = cc
        .bcc = bcc
        If (anexo <> "") Then
            iMsg.AddAttachment (anexo)
        End If
        .send
    End With
    enviar = True
    
    Set iConf = Nothing
    Set iMsg = Nothing
    Set fld = Nothing
    Exit Function
    
Err:
    MsgBox "Erro ao enviar mensagem. " & Err.Number & " - " & Err.Description, vbCritical
End Function

Private Function setConfig()
    
    Set fld = iConf.Fields
    fld.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    fld.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = ""
    fld.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    fld.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 2
    fld.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = de
    fld.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""
    fld.item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
    fld.update

End Function
