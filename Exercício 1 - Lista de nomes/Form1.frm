VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   5880
   ClientTop       =   2535
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdicionaNome 
      Caption         =   "Adicionar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtListaNomes 
      BackColor       =   &H8000000F&
      Height          =   4815
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtNome 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Lista de nomes"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.KeyPreview = True
End Sub

Private Sub cmdAdicionaNome_Click()
    If txtNome.Text = "" Then
        MsgBox "Campo não pode ficar vazio! Um nome deve ser digitado."
    End If


    If txtListaNomes.Text = "" Then
        txtListaNomes.Text = txtNome.Text
    Else
        txtListaNomes.Text = txtListaNomes.Text & vbCrLf & txtNome.Text
    End If
    txtNome = ""
    txtNome.SetFocus
End Sub

Private Sub cmdLimpar_Click()
    txtListaNomes.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim sOk As VbMsgBoxResult
    
    sOk = MsgBox("Deseja continuar?", vbQuestion + vbYesNo, "Form1.Form_QueryUnload")
    
    If sOk = VbMsgBoxResult.vbNo Then
        Cancel = True
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdAdicionaNome_Click
    End If
End Sub
