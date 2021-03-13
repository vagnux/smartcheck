VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procurar"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4230
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   4230
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Data Data1 
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
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form10.frx":0442
      Height          =   5775
      Left            =   0
      OleObjectBlob   =   "Form10.frx":0452
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "cliente"
      DataSource      =   "Data1"
      Height          =   135
      Left            =   480
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim search As String
Private Sub Command1_Click()
search = "select * from cliente where cliente like '" & Text1.Text & "*'"
    Data1.RecordSource = search
    Data1.Refresh

End Sub

Private Sub DBGrid1_DblClick()
Form9.Label8.Caption = Form10.Label1.Caption
Form9.Timer1.Enabled = True

End Sub

Private Sub Form_Load()
Data1.DatabaseName = Form1.Text1.Text
Data1.RecordSource = "cliente"
Data1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
search = "select * from cliente where cliente like '" & Text1.Text & "*'"
    Data1.RecordSource = search
    Data1.Refresh
End If

End Sub
