VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debitar Cheque"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   6630
   Begin MSACAL.Calendar Calendar1 
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   4471
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   1999
      Month           =   8
      Day             =   4
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form7.frx":0442
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "Form7.frx":0452
      TabIndex        =   6
      Top             =   0
      Width           =   6615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "devolvido"
      DataField       =   "Devolvido/semfundo/etc"
      DataSource      =   "Data1"
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Localizar"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Debitar"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   195
      Left            =   4800
      TabIndex        =   3
      Top             =   2880
      Width           =   390
   End
   Begin VB.Label Label2 
      DataField       =   "Cheque nº"
      DataSource      =   "Data1"
      Height          =   135
      Left            =   1680
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim retorno As Integer
Dim search As String

Private Sub Calendar1_Click()
Text1.Text = Calendar1.Value

End Sub

Private Sub Command1_Click()
On Error GoTo lbt
If Label2.Caption = "" Then
Else
    If Check1.Value = 1 Then
    retorno = MsgBox("O cheque n.º " & Label2.Caption & ", já apresentou problemas bancários, Você tem certeza que deseja debita-lo ?", 52)
    Else
    retorno = MsgBox("Você deseja debitar o cheque n.º " & Label2.Caption & "  ?", 4)
    End If
        If retorno = 6 Then
        Data1.Recordset.Delete
        End If
End If
lbt:
Select Case Error
Case Error(3021)
MsgBox "Não existe mais nenhum cheque ou nenhum cheque foi selecionado"
End Select


End Sub

Private Sub Command2_Click()
search = "select * from cheque where creditar like '" & Text1.Text & "*'"
Data1.RecordSource = search
Data1.Refresh
End Sub

Private Sub Command3_Click()
Form7.ScaleHeight = 3925

End Sub

Private Sub Form_Load()

On Error GoTo lberro
Calendar1.Value = Date

Data1.DatabaseName = Form1.Text1.Text
Data1.RecordSource = "cheque"
Data1.Refresh
Text1.Text = Date
search = "select * from cheque where creditar like '" & Text1.Text & "*'"
Data1.RecordSource = search
Data1.Refresh

lberro:
Select Case Error
Case Error(3197)
MsgBox "Registro alterado por outro usuario.", vbOKOnly = 36
Case Error(3024)
MsgBox "O local do banco de dados foi movido ou não está acessivel !", 16
Case Error(3078)
MsgBox "O Banco de dados não é um banco de dados SmartV2 valido !", 16
Case Error(3022)
MsgBox "Este cliente já esta cadastrado", 16, "erro"
End Select


End Sub

