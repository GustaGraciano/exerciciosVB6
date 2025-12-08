VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   6225
   ClientTop       =   2790
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6585
   Begin VB.CommandButton Command3 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
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
      Begin VB.CheckBox Check1 
         Caption         =   "Especial"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
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
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0007
      TabIndex        =   0
      Top             =   2880
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub List1_Click()

End Sub
