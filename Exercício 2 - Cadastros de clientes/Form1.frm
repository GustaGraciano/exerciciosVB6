VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Lista de registros"
   ClientHeight    =   6240
   ClientLeft      =   6225
   ClientTop       =   2790
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame frameDados 
      Caption         =   "Dados"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chbEspecial 
         Caption         =   "Especial"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtIdade 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Endereço completo"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Idade"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ListBox lstNomes 
      Height          =   5130
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0007
      TabIndex        =   0
      Top             =   600
      Width           =   6375
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim conn As String
Dim cli As Clientes

Private Sub Form_Load()
On Error GoTo Trata

cn.ConnectionString = "Provider=SQLOLEDB;Data Source=DEV-GRACIANO;Initial Catalog=exercicio2;Integrated Security=SSPI;"
cn.Open

lstNomes.Clear

Call CarregarClientes

Exit Sub

Trata:
MsgBox "Erro na conexão: " & Err.Description

End Sub

Private Sub txtIdade_KeyPress(KeyAscii As Integer)
    Call ApenasNumeros(KeyAscii)
End Sub

Private Function ChecaSeIdExiste(ByVal parId As Long) As Boolean
Dim i As Integer

For i = 0 To lstNomes.ListCount - 1
    If lstNomes.ItemData(i) = parId Then
        ChecaSeIdExiste = True
        Exit Function
    End If
Next i

ChecaSeIdExiste = False
End Function

Private Sub AdicionaAtualizaCliente(ByVal parId As Long, ByVal parNome As String)
    Dim sIndice As Integer
    Dim sPosicao As Integer
    Dim sAtualizado As Boolean
    Dim sIndiceAntigo As Integer
    
    For sIndice = 0 To lstNomes.ListCount - 1
        If lstNomes.ItemData(sIndice) = parId Then
            lstNomes.List(sIndice) = parNome
            sAtualizado = True
            Exit For
        End If
    Next sIndice
    
    If sAtualizado Then
        sIndiceAntigo = sIndice
        lstNomes.RemoveItem sIndiceAntigo
    End If
    
    sPosicao = lstNomes.ListCount
    For sIndice = 0 To lstNomes.ListCount - 1
        If StrComp(parNome, lstNomes.List(sIndice), vbTextCompare) < 0 Then
            sPosicao = sIndice
            Exit For
        End If
    Next sIndice
    
    lstNomes.AddItem parNome, sPosicao
    lstNomes.ItemData(sPosicao) = parId
End Sub

Public Sub CarregarClientes()
Dim sLinha As String
If rs.State = adStateOpen Then
    rs.Close
End If

lstNomes.Clear

rs.Open "SELECT * FROM Clientes", cn, adOpenStatic, adLockReadOnly

Do Until rs.EOF
    sLinha = "[" & rs!ID & "] - " & _
            "[" & rs!Nome & "] - " & _
            "[" & rs!Idade & "] - " & _
            "[" & rs!Endereco & "]"

    If rs!Especial = True Then
        sLinha = sLinha & " - [Especial]"
    End If
    Call AdicionaAtualizaCliente(rs!ID, sLinha)
    rs.MoveNext
Loop

rs.Close
End Sub

Private Sub cmdIncluir_Click()
frameDados.Visible = True
lstNomes.Visible = False
cmdIncluir.Visible = False
cmdEditar.Visible = False
cmdExcluir.Visible = False
cmdConfirmar.Visible = True
cmdCancelar.Visible = True
End Sub

Private Sub cmdCancelar_Click()
frameDados.Visible = False
lstNomes.Visible = True
cmdIncluir.Visible = True
cmdEditar.Visible = True
cmdExcluir.Visible = True
cmdConfirmar.Visible = False
cmdCancelar.Visible = False
End Sub

Private Sub cmdConfirmar_Click()
If txtNome.Text = "" Or txtIdade.Text = "" Or txtEndereco = "" Then
    MsgBox "Nenhum campo deve ficar vazio!", vbExclamation
    Exit Sub
End If

If txtIdade.Text > 125 Then
    MsgBox "Idade não pode ser superior a 125 anos!", vbExclamation
    Exit Sub
End If

If Len(txtNome.Text) > 50 Then
    MsgBox "Campo nome não pode ter mais de 50 caracteres!", vbExclamation
    Exit Sub
End If

If Len(txtEndereco.Text) > 250 Then
    MsgBox "Campo nome não pode ter mais de 50 caracteres!", vbExclamation
    Exit Sub
End If

sSql = "INSERT INTO Clientes (Nome, Idade, Endereco, Especial) VALUES (" & _
        "'" & txtNome.Text & "', " & _
        "'" & txtIdade.Text & "', " & _
        "'" & txtEndereco.Text & "', " & _
        IIf(chbEspecial.Value = 1, 1, 0) & ")"

cn.Execute sSql

txtNome.Text = ""
txtIdade.Text = ""
txtEndereco.Text = ""
chbEspecial = 0

Call CarregarClientes

frameDados.Visible = False
lstNomes.Visible = True
cmdIncluir.Visible = True
cmdEditar.Visible = True
cmdExcluir.Visible = True
cmdConfirmar.Visible = False
cmdCancelar.Visible = False
End Sub

Private Sub cmdEditar_Click()
If lstNomes.ListIndex = -1 Then
    MsgBox "Favor selecionar um cliente!", vbExclamation
    Exit Sub
End If

Dim sIdSelecionado As Long
sIdSelecionado = lstNomes.ItemData(lstNomes.ListIndex)


frmEditar.IdCliente = sIdSelecionado
frmEditar.Show
End Sub


Private Sub cmdExcluir_Click()
If lstNomes.ListIndex = -1 Then
    MsgBox "Favor selecionar um cliente!", vbExclamation
    Exit Sub
End If

Dim sIdSelecionado As Long
sIdSelecionado = lstNomes.ItemData(lstNomes.ListIndex)

Dim sOk As VbMsgBoxResult
sOk = MsgBox("Deseja mesmo deletar esse cliente?", vbQuestion + vbYesNo, "frmPrincipal.cmdExcluir_Click")
If sOk = VbMsgBoxResult.vbYes Then
    sSql = "DELETE FROM Clientes WHERE ID = " & sIdSelecionado

    cn.Execute sSql

    Call CarregarClientes
End If
End Sub

Public Sub ApenasNumeros(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
    If KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End If
End Sub
