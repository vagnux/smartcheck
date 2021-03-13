VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   3435
   ClientLeft      =   5850
   ClientTop       =   2700
   ClientWidth     =   6270
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form8.frx":0442
      Top             =   480
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form8

End Sub

Private Sub Form_Load()
Form8.BackColor = MDIForm1.BackColor

End Sub

