VERSION 5.00
Begin VB.Form frmEnvEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de envio de e-mail"
   ClientHeight    =   4080
   ClientLeft      =   4590
   ClientTop       =   4350
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8460
   Begin VB.CommandButton cmdenvmsg 
      Caption         =   "Enviar"
      Height          =   765
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Autorizar Orde de Serviço"
      Top             =   3240
      Width           =   855
   End
   Begin VB.ComboBox cbMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   8
      Top             =   240
      Width           =   3705
   End
   Begin VB.TextBox ctTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   90
      MaxLength       =   39
      TabIndex        =   5
      Top             =   240
      Width           =   4575
   End
   Begin VB.TextBox ctEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3810
      MaxLength       =   39
      TabIndex        =   3
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox ctNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   90
      MaxLength       =   39
      TabIndex        =   1
      Top             =   840
      Width           =   3705
   End
   Begin VB.TextBox ctMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1500
      Width           =   8295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4710
      TabIndex        =   9
      Top             =   0
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Texto"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   1260
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3810
      TabIndex        =   4
      Top             =   630
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   630
      Width           =   465
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuTbMsg 
         Caption         =   "Tabela de mensagens"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmEnvEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbMsg_Click()
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from tbmsg where titulo='" & cbMsg & "'", conn
ctMsg = rsTabelas!msg
ctTitulo = rsTabelas!titulo
rsTabelas.Close
Set rsTabelas = Nothing

End Sub


Private Sub cmdenvmsg_Click()
    Call fcOutlook(ctTitulo, ctNome, ctEmail, ctMsg)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from tbmsg order by titulo ASC", conn
Do While Not rsTabelas.EOF
    cbMsg.AddItem rsTabelas!titulo
    rsTabelas.MoveNext
Loop
rsTabelas.Close
Set rsTabelas = Nothing

End Sub

Private Sub mnuTbMsg_Click()
frmcadMsg.Show
End Sub
