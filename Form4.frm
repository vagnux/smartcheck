VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Novo Cheque"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   8400
   Begin VB.CommandButton Command4 
      Caption         =   "Procurar"
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      DataField       =   "cliente"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7680
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Corrigir Tudo"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adicionar"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Cheque"
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   8175
      Begin VB.CheckBox Check1 
         Caption         =   "Cheque devolvido "
         Height          =   255
         Left            =   4200
         TabIndex        =   20
         Top             =   1680
         Width           =   2535
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2415
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   3975
         _Version        =   524288
         _ExtentX        =   7011
         _ExtentY        =   4260
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   1999
         Month           =   8
         Day             =   4
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
      Begin VB.TextBox Text8 
         Height          =   1215
         Left            =   4200
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4200
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observação sobre o cheque :"
         Height          =   195
         Left            =   4200
         TabIndex        =   15
         Top             =   2160
         Width           =   2115
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data a ser creditado :"
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor : "
         Height          =   195
         Left            =   4200
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia :"
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco :"
         Height          =   195
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Numero do cheque"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   1140
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
      RecordSource    =   "cheque"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Cliente : "
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim smartfind As String

Private Sub Calendar1_Click()
Text7.Text = Calendar1.Value

End Sub

Private Sub Command1_Click()
On Error GoTo lberro
If Text8.Text = "" Then
Text8.Text = "Dado inexistente"
End If
Data1.Recordset.AddNew
Data1.Recordset.Fields("cliente") = Text1.Text
Data1.Recordset.Fields("cheque nº") = Text3.Text
Data1.Recordset.Fields("banco") = Text4.Text
Data1.Recordset.Fields("agencia") = Text5.Text
Data1.Recordset.Fields("valor") = Text6.Text
Data1.Recordset.Fields("creditar") = Text7.Text
Data1.Recordset.Fields("obs:") = Text8.Text
Data1.Recordset.Fields("Devolvido/semfundo/etc") = Check1.Enabled
Data1.Recordset.Update
Data1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text1.SetFocus


lberro:
Select Case Error
Case Error(3421)
MsgBox "O formato da data de credito não está correto", vbOKOnly = 36
Text7.SetFocus
End Select
End Sub

Private Sub Command2_Click()
retorno = MsgBox("Você Tem certeza que deseja apagar tudo ?", 36)
If retorno = 6 Then
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text1.SetFocus
End If



End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Form_Load()
On Error GoTo lberro

Data1.DatabaseName = Form1.Text1.Text
Data2.DatabaseName = Form1.Text1.Text
Data1.RecordSource = "cheque"
Data2.RecordSource = "cliente"
Data1.Refresh
Data2.Refresh
Calendar1.Value = Date

lberro:
Select Case Error
Case Error(3197)
MsgBox "Registro alterado por outro usuario.", vbOKOnly = 36
Case Error(3024)
MsgBox "O local do banco de dados foi movido ou não está acessivel !", 16
Timer1.Enabled = False
Case Error(3078)
MsgBox "O Banco de dados não é um banco de dados SmartV2 valido !", 16
Timer1.Enabled = False

Case Error(3022)
MsgBox "Este cliente já esta cadastrado", 16, "erro"
End Select

End Sub

Private Sub Text1_GotFocus()
Timer1.Enabled = True

End Sub

Private Sub Text2_GotFocus()
Text3.SetFocus

End Sub

Private Sub Text3_GotFocus()
If Text2.Text = "Cliente não cadastrado" Then
MsgBox "Este Cliente não se encontra na ficha, Por faver cadastre-o !"
Form3.Show
Text2.Visible = False
Timer1.Enabled = False
Text2.Text = ""
Form3.Text1.Text = Text1.Text
Form3.Text1.SetFocus

Else
Text1.Text = Text2.Text
Text2.Visible = False
Timer1.Enabled = False

End If

End Sub

Private Sub Text4_GotFocus()
If Text2.Text = "Cliente não cadastrado" Then
MsgBox "Este Cliente não se encontra na ficha, Por faver cadastre-o !"
Form3.Show
Text2.Visible = False
Timer1.Enabled = False
Else
Text1.Text = Text2.Text
Text2.Visible = False
Timer1.Enabled = False

End If

End Sub

Private Sub Timer1_Timer()
If Text1.Text <> "" Then
    Text2.Visible = True
    smartfind = "select * from cliente where cliente like '" & Text1.Text & "*'"
    Data2.RecordSource = smartfind
    Data2.Refresh
    If Text2.Text = "" Then
        Text2.Text = "Cliente não cadastrado"
        smartfind = "select * from cliente where cliente like '*'"
    Data2.RecordSource = smartfind
    End If
Else
    Text2.Visible = False
End If



End Sub
