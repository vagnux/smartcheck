VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheques a Debitar"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   6030
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   4215
      _Version        =   524288
      _ExtentX        =   7435
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   1999
      Month           =   8
      Day             =   2
      DayLength       =   1
      MonthLength     =   0
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
   Begin VB.CommandButton Command1 
      Caption         =   "Verificar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3360
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
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form6.frx":0442
      Height          =   1935
      Left            =   0
      OleObjectBlob   =   "Form6.frx":0452
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      Height          =   195
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Width           =   390
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim search As String

Private Sub Calendar1_Click()
Text1.Text = Calendar1.Value

End Sub

Private Sub Command1_Click()
On Error GoTo lberro
If Text1.Text = "todos" Then
        Text1.Text = ""
        search = "select * from cliente_valor where creditar like '" & Text1.Text & "*'"
    Else
        search = "select * from cliente_valor where creditar like '" & Text1.Text & "*'"
    End If
    Data1.RecordSource = search
    Data1.Refresh
lberro:

End Sub

Private Sub Form_Load()
On Error GoTo lberro

Data1.DatabaseName = Form1.Text1.Text
Data1.RecordSource = "cliente_valor"
Data1.Refresh
Text1.Text = Date
search = "select * from cliente_valor where creditar like '" & Text1.Text & "*'"
    Data1.RecordSource = search
    Data1.Refresh
Calendar1.Value = Date

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

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo lberro
If KeyAscii = 13 Then
    If Text1.Text = "todos" Then
        Text1.Text = ""
        search = "select * from cliente_valor where creditar like '" & Text1.Text & "*'"
    Else
        search = "select * from cliente_valor where creditar like '" & Text1.Text & "*'"
    End If
    Data1.RecordSource = search
    Data1.Refresh

lberro:

End If


End Sub

