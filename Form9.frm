VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar informações do cliente"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7050
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   7050
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   5880
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "observ"
      DataSource      =   "Data1"
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   3840
      Width           =   6615
   End
   Begin VB.TextBox Text5 
      DataField       =   "tel"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3840
      TabIndex        =   7
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "cep"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "endereco"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   6615
   End
   Begin VB.TextBox Text2 
      DataField       =   "cgc"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      DataField       =   "cpf"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procurar"
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nome do Cliente"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.Label Label1 
         DataField       =   "cliente"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observações do Cliente"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefones:"
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   2880
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cep :"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço :"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CGC:"
      Height          =   195
      Left            =   4200
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   345
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim smartfind As String

Private Sub Command1_Click()
Form10.Show

End Sub

Private Sub Command2_Click()
If Label1.Caption <> "" Then
    Data1.Recordset.Edit
    Data1.Recordset.Update
    Data1.RecordSource = "select * from cliente where cliente like ''"
    Data1.Refresh
    
Else
    MsgBox "Não Foi especificado Nenhum Cliente !!!"
End If

End Sub

Private Sub Command3_Click()
Unload Form9

End Sub

Private Sub Form_Load()
On Error GoTo trat


Data1.DatabaseName = Form1.Text1.Text
Data1.RecordSource = "cliente"
Data1.Refresh
Data1.RecordSource = "select * from cliente where cliente like ''"
Data1.Refresh

trat:
Select Case Error
Case Error(3024)
MsgBox "O arquivo de banco de dados não é valido ou não foi encontrado !"
Form2.Show
End Select
End Sub

Private Sub Timer1_Timer()
smartfind = "select * from cliente where cliente like '" & Label8.Caption & "*'"
Data1.RecordSource = smartfind
Data1.Refresh
Timer1.Enabled = False
End Sub
