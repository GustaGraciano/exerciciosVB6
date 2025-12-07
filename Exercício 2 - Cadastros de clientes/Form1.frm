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
   Begin VB.Frame Frame1 
      Caption         =   "Dados"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1200
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
      Height          =   3180
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0007
      TabIndex        =   0
      Top             =   2880
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
Dim i As Integer

For i = 0 To lstNomes.ListCount - 1

    If lstNomes.ItemData(i) = parId Then
        lstNomes.List(i) = parNome
        lstNomes.ItemData(i) = parId
        Exit Sub
    End If

Next i

lstNomes.AddItem parNome
lstNomes.ItemData(lstNomes.NewIndex) = parId
End Sub

Public Sub CarregarClientes()
If rs.State = adStateOpen Then
    rs.Close
End If

lstNomes.Clear

rs.Open "SELECT * FROM Clientes", cn, adOpenStatic, adLockReadOnly

Do Until rs.EOF
    Call AdicionaAtualizaCliente(rs!ID, rs!Nome)
    rs.MoveNext
Loop

rs.Close
End Sub

Private Sub cmdIncluir_Click()
If txtNome.Text = "" Or txtIdade.Text = "" Or txtEndereco = "" Then
    MsgBox "Nenhum campo deve ficar vazio!", vbExclamation
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
