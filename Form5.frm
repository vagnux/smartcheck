VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Cliente"
   ClientHeight    =   5970
   ClientLeft      =   3255
   ClientTop       =   2265
   ClientWidth     =   7050
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   7050
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ficha cliente"
      TabPicture(0)   =   "Form5.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label30"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Data3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Data4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Timer2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Cliente e Cheques"
      TabPicture(1)   =   "Form5.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label29"
      Tab(1).Control(1)=   "DBGrid1"
      Tab(1).Control(2)=   "Text2"
      Tab(1).Control(3)=   "Text1"
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(6)=   "Data2"
      Tab(1).Control(7)=   "Data1"
      Tab(1).ControlCount=   8
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   5400
         Top             =   4680
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "cliente"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Consultar"
         Height          =   255
         Left            =   5280
         TabIndex        =   30
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   4800
         Width           =   5055
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H0000FFFF&
         DataField       =   "cliente"
         DataSource      =   "Data4"
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
         Left            =   840
         TabIndex        =   28
         Top             =   4560
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ficha do Cliente"
         Height          =   3015
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   6735
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Label22"
            DataField       =   "observ"
            DataSource      =   "Data3"
            Height          =   795
            Left            =   840
            TabIndex        =   37
            Top             =   2040
            Width           =   5730
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label22"
            DataField       =   "Tel"
            DataSource      =   "Data3"
            Height          =   195
            Left            =   3600
            TabIndex        =   36
            Top             =   1440
            Width           =   570
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label22"
            DataField       =   "cep"
            DataSource      =   "Data3"
            Height          =   195
            Left            =   840
            TabIndex        =   35
            Top             =   1560
            Width           =   570
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label22"
            DataField       =   "endereco"
            DataSource      =   "Data3"
            Height          =   195
            Left            =   1200
            TabIndex        =   34
            Top             =   1200
            Width           =   570
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label22"
            DataField       =   "CGC"
            DataSource      =   "Data3"
            Height          =   195
            Left            =   3240
            TabIndex        =   33
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label22"
            DataField       =   "CPF"
            DataSource      =   "Data3"
            Height          =   195
            Left            =   840
            TabIndex        =   32
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label22"
            DataField       =   "cliente"
            DataSource      =   "Data3"
            Height          =   195
            Left            =   960
            TabIndex        =   31
            Top             =   480
            Width           =   570
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OBS:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   360
            TabIndex        =   27
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2880
            TabIndex        =   26
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CEP:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   360
            TabIndex        =   25
            Top             =   1560
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   360
            TabIndex        =   24
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CGC:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2760
            TabIndex        =   23
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPF:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   360
            TabIndex        =   22
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   480
            Width           =   465
         End
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -71640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "cliente_cheque"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -69720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "cliente"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados do Clinte"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   6735
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   6120
            Top             =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPF:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CGC:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2880
            TabIndex        =   17
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CEP:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2880
            TabIndex        =   14
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OBS:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label8"
            DataField       =   "cliente"
            DataSource      =   "Data1"
            Height          =   195
            Left            =   720
            TabIndex        =   12
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label9"
            DataField       =   "CPF"
            DataSource      =   "Data1"
            Height          =   195
            Left            =   600
            TabIndex        =   11
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label10"
            DataField       =   "CGC"
            DataSource      =   "Data1"
            Height          =   195
            Left            =   3360
            TabIndex        =   10
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label11"
            DataField       =   "endereco"
            DataSource      =   "Data1"
            Height          =   195
            Left            =   1080
            TabIndex        =   9
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            DataField       =   "cep"
            DataSource      =   "Data1"
            Height          =   195
            Left            =   600
            TabIndex        =   8
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label13"
            DataField       =   "Tel"
            DataSource      =   "Data1"
            Height          =   195
            Left            =   3600
            TabIndex        =   7
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Label14"
            DataField       =   "observ"
            DataSource      =   "Data1"
            Height          =   615
            Left            =   600
            TabIndex        =   6
            Top             =   1320
            Width           =   6015
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar"
         Height          =   255
         Left            =   -69720
         TabIndex        =   4
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox Text1 
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
         Left            =   -74160
         TabIndex        =   2
         Top             =   5040
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74880
         TabIndex        =   1
         Top             =   5280
         Width           =   5055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Form5.frx":047A
         Height          =   2175
         Left            =   -74880
         OleObjectBlob   =   "Form5.frx":048A
         TabIndex        =   3
         Top             =   2760
         Width           =   6735
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   4560
         Width           =   525
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   38
         Top             =   5040
         Width           =   525
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim search As String


Private Sub Command1_Click()
Text2.Text = Text1.Text
    search = "select * from cliente_cheque where cliente like '" & Text2.Text & "*'"
    Data1.RecordSource = search
    Data1.Refresh
    Text2.Text = ""
End Sub

Private Sub Form_Activate()
Text2.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo lberro

Data1.DatabaseName = Form1.Text1.Text
Data2.DatabaseName = Form1.Text1.Text
Data3.DatabaseName = Form1.Text1.Text
Data4.DatabaseName = Form1.Text1.Text
Data3.RecordSource = "cliente"
Data4.RecordSource = "cliente"
Data1.RecordSource = "cliente_cheque"
Data2.RecordSource = "cliente"
Data1.Refresh
Data2.Refresh
Data1.RecordSource = "select * from cliente where cliente like ''"
Data3.RecordSource = "select * from cliente where cliente like ''"
Data1.Refresh
Data3.Refresh

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

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.Text = Text1.Text
    search = "select * from cliente_cheque where cliente like '" & Text2.Text & "*'"
    Data1.RecordSource = search
    Data1.Refresh
    Text2.Text = ""
    If Label8.Caption = "" Then
    MsgBox "Este Cliente não tem Nenhum Cheque retornado ou a ser Debitado"
    End If
    
    
    
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.Text = Text3.Text
    search = "select * from cliente where cliente like '" & Text4.Text & "*'"
    Data3.RecordSource = search
    Data3.Refresh
    Text4.Text = ""
End If

End Sub

Private Sub Timer1_Timer()

If Text2.Text <> "" Then
    Text1.Visible = True
    search = "select * from cliente where cliente like '" & Text2.Text & "*'"
    Data2.RecordSource = search
    Data2.Refresh
    If Text1.Text = "" Then
       Text1.Text = "Cliente não cadastrado"
    End If
Else
    Text1.Visible = False
End If
End Sub

Private Sub Timer2_Timer()
If Text4.Text <> "" Then
    Text3.Visible = True
    search = "select * from cliente where cliente like '" & Text4.Text & "*'"
    Data4.RecordSource = search
    Data4.Refresh
    If Text3.Text = "" Then
       Text3.Text = "Cliente não cadastrado"
    End If
Else
    Text3.Visible = False
End If
End Sub
