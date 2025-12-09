VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmPrincipal 
   Caption         =   "Cadastro de clientes"
   ClientHeight    =   6240
   ClientLeft      =   3930
   ClientTop       =   2730
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   9330
   Begin FPSpreadADO.fpSpread gridDados 
      Height          =   2775
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   9135
      _Version        =   458752
      _ExtentX        =   16113
      _ExtentY        =   4895
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   1
      SpreadDesigner  =   "Form1.frx":0000
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9135
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de comunicação"
         Height          =   735
         Left            =   3360
         TabIndex        =   17
         Top             =   1200
         Width           =   4575
         Begin VB.OptionButton opbCom 
            Caption         =   "Carta"
            Height          =   195
            Index           =   3
            Left            =   3360
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton opbCom 
            Caption         =   "WhatsApp"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton opbCom 
            Caption         =   "Telefone"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton opbCom 
            Caption         =   "Email"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
      End
      Begin EditLib.fpDateTime fpDataNasc 
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   600
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
         Left            =   2160
         TabIndex        =   14
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
      Begin VB.ComboBox cmbSexo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox chbEspecial 
         Caption         =   "Especial"
         Height          =   255
         Left            =   8040
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtIdade 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtNome 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Valor de crédito"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Sexo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Endereço completo"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Idade"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
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

Private Sub Form_Load()
Conectar
CarregarClientes

With cmbSexo
    .AddItem "Masculino"
    .AddItem "Feminino"
    .AddItem "Indefinido"
End With

opbCom(2).Value = True

txtIdade.Text = CalcularIdade(CDate(fpDataNasc.Text))
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    ApenasCaracteres (KeyAscii)
End Sub

Private Sub fpDataNasc_LostFocus()
    If IsDate(fpDataNasc.Text) Then
        txtIdade.Text = CalcularIdade(CDate(fpDataNasc.Text))
    End If
End Sub

Public Sub CarregarClientes()
If rs.State = adStateOpen Then rs.Close

sSql = "SELECT id, nome, dataNasc, endereco, especial FROM Clientes"

If IniP(eConsulta, sSql, rs) Then
    If Not rs.EOF Then
        rs.MoveLast
        gridDados.MaxCols = rs.Fields.Count
        gridDados.MaxRows = rs.RecordCount
        rs.MoveFirst

        Dim col As Long
        For col = 0 To rs.Fields.Count - 1
            gridDados.SetText col + 1, 0, rs.Fields(col).Name
        Next col

        Dim row As Long
        row = 1
        Do Until rs.EOF
            For col = 0 To rs.Fields.Count - 1
                gridDados.SetText col + 1, row, rs.Fields(col).Value & ""
            Next col
            row = row + 1
            rs.MoveNext
        Loop
    End If
End If
End Sub

Private Function ClienteSelecionado() As Variant
Dim sLinha As Long
Dim sIdSelecionado As Variant

sLinha = gridDados.ActiveRow

gridDados.GetText 1, sLinha, sIdSelecionado

ClienteSelecionado = sIdSelecionado
End Function

Private Sub cmdIncluir_Click()
Dim sTipo As eTipoComunicacao
Dim sMarcado As Boolean
Dim sSexo As Integer
sSexo = cmbSexo.ListIndex + 1
Dim i As Integer

sMarcado = False

For i = 0 To 3
    If opbCom(i).Value = True Then
        sTipo = i
        sMarcado = True
        Exit For
    End If
Next

If Trim(txtNome.Text) = "" _
   Or Trim(txtIdade.Text) = "" _
   Or Trim(txtEndereco.Text) = "" _
   Or cmbSexo.ListIndex = -1 _
   Or Not sMarcado Then

       MsgBox "Nenhum campo deve ficar vazio!", vbExclamation
       Exit Sub
End If

If Not ChecaFormatoData(fpDataNasc.Text) Then
    MsgBox "A data contém caracteres inválidos!", vbExclamation
    Exit Sub
End If

If fpCredito.Text > 5000 Then
    MsgBox "Valor inválido!", vbExclamation
    Exit Sub
End If

sSql = "INSERT INTO Clientes (nome, dataNasc, endereco, sexo, valCredito, tipoComunicacao, Especial) VALUES (" & _
        "'" & txtNome.Text & "', " & _
        "'" & Format(fpDataNasc, "yyyy-mm-dd") & "', " & _
        "'" & txtEndereco.Text & "', " & _
        "'" & sSexo & "', " & _
        Replace(fpCredito.Text, ",", ".") & ", " & _
        sTipo & ", " & _
        IIf(chbEspecial.Value = 1, 1, 0) & _
        ")"

If IniP(eExecutar, sSql) Then
    MsgBox "Cliente cadastrado com sucesso!"
End If

txtNome.Text = ""
fpDataNasc = "18/09/2020"
txtIdade.Text = ""
txtEndereco.Text = ""
fpCredito.Value = 0
chbEspecial = 0

CarregarClientes
End Sub

Private Sub cmdEditar_Click()
Dim sLinha As Long
Dim sIdSelecionado As Variant

sIdSelecionado = ClienteSelecionado()

frmEditar.IdCliente = (sIdSelecionado)
frmEditar.Show
End Sub

Private Sub cmdExcluir_Click()
Dim sOk As VbMsgBoxResult
Dim sLinha As Long
Dim sIdSelecionado As Variant

sIdSelecionado = ClienteSelecionado()

sSql = "DELETE FROM Clientes WHERE ID = " & sIdSelecionado

sOk = MsgBox("Tem certeza que deseja excluir esse cliente?", vbQuestion + vbYesNo, "frmPrincipal.cmdExcluir_Click")
If sOk = VbMsgBoxResult.vbYes Then
    If IniP(Executar, sSql) Then
        MsgBox "Cliente excluído com sucesso."
        CarregarClientes
    End If
End If
End Sub


