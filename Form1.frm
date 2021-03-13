VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "config"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.TextBox Text3 
      DataField       =   "fundo"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Timer reset 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.TextBox Text2 
      DataField       =   "color"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "000"
      Top             =   1440
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      DataField       =   "database"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "c:\smart\smartv2.mdb"
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\smarts\config.dat"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "config"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "SPCS : Smart Personal Config Service - 20/05/1999   "
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dat As String




Private Sub Form_Load()
Dim SaveTitle As String
 If App.PrevInstance Then
   SaveTitle = App.Title
   App.Title = "... segunda chamada ao mesmo programa."
   Me.Caption = "... segunda chamada ao mesmo programa, serei fechado"
   'se for a Sub Main, a linha acima, obviamente, não existe
   'as linhas abaixo fecham a segunda chamada e alternam para
   'a primeira
   AppActivate SaveTitle
   SendKeys "% R", True
   End
 End If

frmSplash.Show
reset.Enabled = True
Form1.Visible = False



End Sub

Private Sub reset_Timer()
MDIForm1.Show
reset.Enabled = False
End Sub

Private Sub Timer1_Timer()

End Sub
