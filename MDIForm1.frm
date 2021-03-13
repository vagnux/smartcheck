VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00800000&
   Caption         =   "SMARTCHECK 3"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10110
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Adicionar Cliente"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Adicionar Cheque"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Consultar Cliente"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Internet"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Consultar cheques a serem debitados"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Debitar Cheques"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Alterar Dados de um cliente"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Configuração do Sistema"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8475
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   450
      SimpleText      =   "SmartCheck x06 Small Aplication pack 1"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:37"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "07/09/99"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":16DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":1D12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu opt 
      Caption         =   "&Opções"
      Begin VB.Menu add 
         Caption         =   "&Adicionar"
         Begin VB.Menu newclient 
            Caption         =   "Novo Cliente"
         End
         Begin VB.Menu newcheck 
            Caption         =   "Novo Cheque "
         End
      End
      Begin VB.Menu find 
         Caption         =   "&Consultas"
         Begin VB.Menu findclient 
            Caption         =   "Consultar Cliente"
         End
         Begin VB.Menu net 
            Caption         =   "Internet"
         End
         Begin VB.Menu checknow 
            Caption         =   "Cheques a Serem Debitados"
         End
      End
      Begin VB.Menu debitar 
         Caption         =   "Debitar"
         Begin VB.Menu cheq 
            Caption         =   "Cheque"
         End
      End
      Begin VB.Menu modify 
         Caption         =   "Al&terar"
         Begin VB.Menu modclient 
            Caption         =   "Dados de um Cliente"
         End
      End
      Begin VB.Menu config 
         Caption         =   "Con&figurar"
         Begin VB.Menu system 
            Caption         =   "Sistema"
         End
      End
      Begin VB.Menu tabs 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&?"
      Begin VB.Menu about 
         Caption         =   "Sobre ..."
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function falha()
Select Case Num_erro
Case Else
MsgBox "Erro numero : " & Err & " =>  " & Error$, 16
End Select

End Function

Private Sub about_Click()
Form8.Show

End Sub

Private Sub checknow_Click()
Form6.Show

End Sub

Private Sub cheq_Click()
Form7.Show

End Sub

Private Sub exit_Click()
retorno = MsgBox("Você tem Certeza que deseja Sai ?", 36)
If retorno = 6 Then
Unload MDIForm1
End If

End Sub

Private Sub findclient_Click()
Form5.Show

End Sub

Private Sub MDIForm_Load()
MDIForm1.BackColor = Form1.Text2.Text


End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu opt
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload Form1

End Sub

Private Sub modclient_Click()
Form9.Show

End Sub

Private Sub net_Click()
frmBrowser.Show

End Sub

Private Sub newcheck_Click()
Form4.Show

End Sub

Private Sub newclient_Click()
Form3.Show

End Sub

Private Sub system_Click()
Form2.Show

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
    Case 1
    Form3.Show
    Case 2
    Form4.Show
    Case 4
    Form5.Show
    Case 5
    frmBrowser.Show
    Case 6
    Form6.Show
    Case 8
    Form7.Show
    Case 10
    Form9.Show
    Case 12
    Form2.Show
    
    
    End Select
    
End Sub
