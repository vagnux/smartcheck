VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adicionar um novo Cliente"
   ClientHeight    =   5235
   ClientLeft      =   255
   ClientTop       =   2895
   ClientWidth     =   6975
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   6975
   Begin VB.CommandButton Command3 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Corrigir Tudo"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adicionar cliente"
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   3240
      Width           =   6495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   6495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6495
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
      RecordSource    =   "cliente"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observações do Cliente :"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone :"
      Height          =   195
      Left            =   3720
      TabIndex        =   11
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cep :"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Endereço :"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CGC :"
      Height          =   195
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   390
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim falta As String

Private Sub Command1_Click()
On Error GoTo lberro
If Text2.Text = "" Then
Text2.Text = "Dado inexistente"
End If
If Text3.Text = "" Then
Text3.Text = "Dado inexistente"
End If
If Text5.Text = "" Then
Text5.Text = "0000"
End If
If Text7.Text = "" Then
Text7.Text = "Dado inexistente"
End If

Data1.Recordset.AddNew
Data1.Recordset.Fields("cliente") = Text1.Text
Data1.Recordset.Fields("cpf") = Text2.Text
Data1.Recordset.Fields("cgc") = Text3.Text
Data1.Recordset.Fields("endereco") = Text4.Text
Data1.Recordset.Fields("cep") = Text5.Text
Data1.Recordset.Fields("tel") = Text6.Text
Data1.Recordset.Fields("observ") = Text7.Text
Data1.Recordset.Update
Data1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

lberro:
Select Case Error
Case Error(3197)
MsgBox "Registro alterado por outro usuario.", vbOKOnly = 36
Case Error(3024)
MsgBox "O local do banco de dados foi movido ou não está acessivel !", 16
Form2.Show
Case Error(3078)
MsgBox "O Banco de dados não é um banco de dados SmartV2 valido !", 16
Form2.Show
Case Error(3022)
MsgBox "Este cliente já esta cadastrado", 16, "erro"
End Select

End Sub

Private Sub Command2_Click()
retorno = MsgBox("Você Tem certeza que deseja apagar tudo ?", 36)
If retorno = 6 Then
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
End If

End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Form_Load()
On Error GoTo lberro

Data1.DatabaseName = Form1.Text1.Text
Data1.RecordSource = "cliente"
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

Private Sub Text2_Change()
If Text2.Text <> "" Then
Text3.Enabled = False
Text3.BackColor = 12632256
Else
Text3.Enabled = True
Text3.BackColor = 16777215

End If

End Sub

Private Sub Text3_Change()
If Text3.Text <> "" Then
Text2.Enabled = False
Text2.BackColor = 12632256
Else
Text2.Enabled = True
Text2.BackColor = 16777215

End If

End Sub
