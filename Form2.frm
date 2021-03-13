VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Sistema"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5820
   Begin MSComDlg.CommonDialog dalog1 
      Left            =   5160
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      DataField       =   "fundo"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   5415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Selecionar"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "color"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Procucar"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\smarts\config.dat"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "config"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "database"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pagina inicial do navegador"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor de Fundo:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Caminho do Banco de Dados Principal:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.Edit
Data1.Recordset.Update
MsgBox "O Sistema será reiniciado para que as alterações sejam efetuadas !"
Unload MDIForm1
Form1.reset.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
dalog1.filename = ""
dalog1.Filter = "Smarcheck x06 data base |*.mdb|"
dalog1.ShowOpen
Text1.Text = dalog1.filename


End Sub

Private Sub Command4_Click()
dalog1.filename = ""
dalog1.ShowColor
Text2.Text = dalog1.Color
Text2.BackColor = Text2.Text


End Sub

Private Sub Command5_Click()
dalog1.filename = ""
dalog1.Filter = "Papel de Parede |*.bmp|"
dalog1.ShowOpen
Text3.Text = dalog1.filename

End Sub

Private Sub Form_Load()
Text2.BackColor = MDIForm1.BackColor

End Sub
