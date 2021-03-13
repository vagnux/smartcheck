VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sobre"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4875
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "SmartCheck 3"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Software de uso gratuito , não pode ser vendido !"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Versão: "
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "07 de Setembro de 1999"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data de lançamento : "
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label5.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub
