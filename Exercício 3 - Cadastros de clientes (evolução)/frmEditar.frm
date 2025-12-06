VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmEditar 
   Caption         =   "Editar cliente"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame1 
      Caption         =   "Dados"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   495
         Left            =   7080
         TabIndex        =   19
         Top             =   2520
         Width           =   915
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   8160
         TabIndex        =   18
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chbEspecial 
         Caption         =   "Especial"
         Height          =   195
         Left            =   8160
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de comunicação"
         Height          =   735
         Left            =   3480
         TabIndex        =   11
         Top             =   1320
         Width           =   4575
         Begin VB.OptionButton opbCom 
            Caption         =   "Email"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton opbCom 
            Caption         =   "Telefone"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton opbCom 
            Caption         =   "WhatsApp"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   13
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton opbCom 
            Caption         =   "Carta"
            Height          =   195
            Index           =   3
            Left            =   3360
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmbSexo 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtIdade 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtNome 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin EditLib.fpDateTime fpDataNasc 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   1455
         _Version        =   196608
         _ExtentX        =   2566
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   1
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "18/09/2020"
         DateCalcMethod  =   0
         DateTimeFormat  =   0
         UserDefinedFormat=   ""
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle fpCredito 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
         _Version        =   196608
         _ExtentX        =   1931
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         DecimalPlaces   =   -1
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label5 
         Caption         =   "Valor de crédito"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Sexo"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Endereço completo"
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Idade"
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim conn As String
Private fIdCliente As Long

Public Property Let IdCliente(ByVal parIdCliente As Long)
fIdCliente = parIdCliente
End Property

Private Sub txtIdade_KeyPress(KeyAscii As Integer)
    Call ApenasNumeros(KeyAscii)
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    Call ApenasCaracteres(KeyAscii)
End Sub

Private Sub Form_Load()
Conectar
CarregarCliente
End Sub

Private Sub CarregarCliente()
Dim sSql As String

sSql = "SELECT * FROM Clientes WHERE ID = " & fIdCliente

If IniP(Consulta, sSql, rs) Then

End If

If Not rs.EOF Then
    txtNome.Text = rs!Nome
    fpDataNasc = rs!dataNasc
    txtIdade.Text = rs!Idade
    txtEndereco.Text = rs!Endereco
    chbEspecial.Value = IIf(rs!Especial = True, 1, 0)
    cmbSexo.Text = rs!Sexo
    fpCredito.Text = rs!ValCredito
    sTipo = rs!TipoComunicacao
    opbCom(sTipo).Value = True
End If
End Sub

Private Sub cmdEditar_Click()
Dim sTipo As TipoComunicacao

sSql = "UPDATE Clientes SET " & _
        "nome = '" & txtNome.Text & "', " & _
        "dataNasc ='" & Format(fpDataNasc, "yyyy-mm-dd") & "', " & _
        "idade = '" & txtIdade.Text & "', " & _
        "endereco = '" & txtEndereco.Text & "', " & _
        "sexo = '" & cmbSexo.Text & "', " & _
        "valCredito = " & Replace(fpCredito.Text, ",", ".") & ", " & _
        "tipoComunicacao = " & sTipo & ", " & _
        "Especial = " & IIf(chbEspecial.Value = 1, 1, 0) & " " & _
        "WHERE ID = " & fIdCliente & ";"

If IniP(Executar, sSql) Then
    MsgBox "Cliente editado com sucesso!"
End If

frmPrincipal.CarregarClientes
rs.Close

Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
rs.Close
End Sub
