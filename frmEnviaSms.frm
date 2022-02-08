VERSION 5.00
Begin VB.Form frmEnviaSms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de envio de Sms"
   ClientHeight    =   1770
   ClientLeft      =   4305
   ClientTop       =   3735
   ClientWidth     =   7965
   Icon            =   "frmEnviaSms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnviar 
      Height          =   405
      Left            =   5790
      Picture         =   "frmEnviaSms.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Enviar Sms"
      Top             =   180
      Width           =   495
   End
   Begin VB.CommandButton cmdConsultar 
      Height          =   405
      Left            =   5280
      Picture         =   "frmEnviaSms.frx":08D3
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Consultar cadastro de contatos"
      Top             =   180
      Width           =   495
   End
   Begin VB.TextBox ctNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      MaxLength       =   25
      TabIndex        =   2
      Top             =   210
      Width           =   3045
   End
   Begin VB.TextBox ctFone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3150
      MaxLength       =   25
      TabIndex        =   1
      Top             =   210
      Width           =   2055
   End
   Begin VB.TextBox ctsms 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   7875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mensagen"
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
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label8 
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
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fone"
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
      Left            =   3150
      TabIndex        =   3
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "frmEnviaSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1


Private Sub cmdConsultar_Click()
nIDsms = 0
frmCadsms.Show (1)

If nIDsms <> 0 Then
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from contatossms where id=" & nIDsms & "", conn
    ctNome = rsTabelas!nome
    ctFone = rsTabelas!fone
    rsTabelas.Close
    Set rsTabelas = Nothing
End If
End Sub

Private Sub cmdEnviar_Click()
'ShellExecute hwnd, "open", "http://www.vbmania.com.br", vbNullString, vbNullString, conSwNo
'ShellExecute hwnd, "open", "http://www.smsemmassa.com.br/gateway.php?usuario=ernande&senha=sms3302&fone=55" & ctFone & "&msg=" & ctsms & "", vbNullString, vbNullString, conSwNo
'ShellExecute hwnd, "open", "http://www.smsemmassa.com.br/gateway.php?usuario=brsofthouse&senha=ernandesms&fone=55" & ctFone & "&msg=" & ctsms & "", vbNullString, vbNullString, conSwNo
ShellExecute hwnd, "open", "http://www.smsemmassa.com.br/gateway.php?usuario=ernande&senha=sms3302&fone=55" & ctFone & "&msg=" & ctsms & "", vbNullString, vbNullString, conSwNo

'$sms ="http://www.smsemmassa.com.br/gateway.php?usuario=brsofthouse&senha=ernandesms&fone=55".$Fone."&msg=".$Mensage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True

End Sub
