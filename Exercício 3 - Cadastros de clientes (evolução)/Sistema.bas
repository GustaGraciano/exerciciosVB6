Attribute VB_Name = "Sistema"
Option Explicit
Public cn As ADODB.Connection

Public Enum eQuery
    Consulta = 0
    Executar = 1
End Enum

Public Enum TipoComunicacao
    Email = 0
    Telefone = 1
    Whatsapp = 2
    Carta = 3
End Enum

Public Sub Conectar()
On Error GoTo Trata

Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=SQLOLEDB;Data Source=DEV-GRACIANO;Initial Catalog=exercicio3;Integrated Security=SSPI;"
cn.Open
Exit Sub

Trata:
MsgBox "Erro na conexão: " & Err.Description
End Sub

Public Function IniP(ByVal parTipoConsulta As eQuery, ByVal sSql As String, Optional ByRef rs As ADODB.Recordset) As Boolean
On Error GoTo Trata

If parTipoConsulta = Consulta Then
    If rs Is Nothing Then
        Set rs = New ADODB.Recordset
    ElseIf rs.State <> 0 Then
        rs.Close
    End If
    rs.CursorLocation = adUseClient
    rs.Open sSql, cn, adOpenStatic, adLockReadOnly
Else
    cn.Execute sSql
End If

IniP = True

Exit Function
Trata:
IniP = False

MsgBox "Erro ao executar consulta:" & vbCrLf & _
        "SQL: " & sSql & vbCrLf & _
        "Descrição: " & Err.Description, vbCritical, "Erro"
End Function

Public Sub ApenasNumeros(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
    If KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End If
End Sub

Public Sub ApenasCaracteres(KeyAscii As Integer)

' Letras maiúsculas
If KeyAscii >= 65 And KeyAscii <= 90 Then Exit Sub

' Letras minúsculas
If KeyAscii >= 97 And KeyAscii <= 122 Then Exit Sub

' Letras acentuadas e ç
If KeyAscii >= 192 And KeyAscii <= 255 Then Exit Sub

' Espaço
If KeyAscii = Asc(" ") Then Exit Sub

' Backspace
If KeyAscii = vbKeyBack Then Exit Sub

' Qualquer outra tecla
KeyAscii = 0
End Sub
