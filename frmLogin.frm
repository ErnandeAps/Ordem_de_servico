VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de Login"
   ClientHeight    =   3375
   ClientLeft      =   6885
   ClientTop       =   5190
   ClientWidth     =   5295
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleMode       =   0  'User
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbConexão 
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
      ItemData        =   "frmLogin.frx":1C7A
      Left            =   90
      List            =   "frmLogin.frx":1C84
      TabIndex        =   13
      Text            =   "Remota"
      Top             =   2940
      Width           =   2835
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   405
      Left            =   4290
      TabIndex        =   10
      Top             =   2910
      Width           =   945
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Entrar"
      Height          =   405
      Left            =   3240
      TabIndex        =   9
      Top             =   2910
      Width           =   1035
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000009&
      Height          =   945
      Left            =   60
      TabIndex        =   4
      Top             =   1920
      Width           =   2835
      Begin VB.Frame Frame5 
         BackColor       =   &H80000009&
         Caption         =   "Data"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Width           =   1425
         Begin VB.TextBox Ct_Data 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   405
            Left            =   90
            TabIndex        =   8
            Top             =   240
            Width           =   1245
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000009&
         Caption         =   "Hora"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   1530
         TabIndex        =   5
         Top             =   150
         Width           =   1245
         Begin VB.TextBox CT_Hora 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   405
            Left            =   90
            TabIndex        =   6
            Top             =   240
            Width           =   1065
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1290
      Top             =   2190
   End
   Begin VB.TextBox ctLogin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   3750
      TabIndex        =   0
      Top             =   1950
      Width           =   1485
   End
   Begin VB.TextBox ctSenha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   3750
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2430
      Width           =   1485
   End
   Begin VB.Label lblRegistro 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Licenciado para :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label lblLicenca 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Licenciado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Login:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3060
      TabIndex        =   3
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2970
      TabIndex        =   2
      Top             =   2490
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   30
      Picture         =   "frmLogin.frx":1C97
      Stretch         =   -1  'True
      Top             =   30
      Width           =   5250
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdLogin_Click()
Select Case cbConexão
Case "Local"
    strDns = "DSN=suportekDb;Uid=root;pwd=(#suporte#)"
Case "Remota"
    strDns = "DSN=suportek-srv;Uid=root;pwd=(#suporte#)"
End Select

Call Conecta_BD

Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select * from clientes where login='" & ctlogin & "' and senha='" & ctsenha & "'", conn


'Set Reg_Dados = BANCO.OpenRecordset("Select * from contseg where usuario='" & ctLogin & "' and senha='" & ctSenha & "'")
If rsTabelas.EOF Then
    MsgBox "Usuário ou senha inválido.", vbCritical, "HairSoft"
    Exit Sub
End If

MOP = ctlogin
'MNIVEL = Reg_Dados!Nivel

rsTabelas.Close
Set rsTabelas = Nothing
Unload Me

frmtelaPrin.Show

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub ctLogin_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    If ctlogin.Text = "" Then
        MsgBox "Didite o seu login para continuar.", vbCritical
        Exit Sub
    End If
    ctsenha.SetFocus
End Select
End Sub

Private Sub ctSenha_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    Select Case cbConexão
    Case "Local"
        'strDns = "DSN=suportekDb;Uid=root;pwd=(#suporte#)"
        strDns = "DSN=suportekDb;Uid=root;pwd=(#suporte#)"
    Case "Remota"
        strDns = "DSN=suportek-srv;Uid=root;pwd=(#suporte#)"
    End Select
    
    strDnsRemoto = "DSN=Suportekreomoto;Uid=suportek_supbr;pwd=(#suporte#)"
    
    Call Conecta_BD
    
    Set rsTabelas = New ADODB.Recordset
    
    rsTabelas.Open "select * from clientes where login='" & ctlogin & "' and senha='" & ctsenha & "'", conn
    
    
    'Set Reg_Dados = BANCO.OpenRecordset("Select * from contseg where usuario='" & ctLogin & "' and senha='" & ctSenha & "'")
    If rsTabelas.EOF Then
        MsgBox "Usuário ou senha inválido.", vbCritical, "HairSoft"
        Exit Sub
    End If
    
    MOP = ctlogin
    'MNIVEL = Reg_Dados!Nivel
    
    rsTabelas.Close
    Set rsTabelas = Nothing
    Unload Me
    'frmAgenda.Show
    'frmtelaPrin.Show
    frmFlexGrid.Show
End Select
End Sub

Private Sub Form_Load()
Dim Ret As Integer
Dim nDias As Integer

Ct_Data.Text = Date
'Call VerifRegistro
'If VerifRegistro = "Registrado" Then
'    Set Reg_Dados = BANCO.OpenRecordset("select * from cad_empresa")
'    Set Reg_Dados3 = BANCOLIB.OpenRecordset("select * from trava")
'    Set Reg_Dados2 = BANCOLIB.OpenRecordset("select * from cad_empresa")
'    If Reg_Dados!serial <> Reg_Dados2!serial Then
'        Ret = MsgBox("Este programa não esta registrado para este computador. Consulte o suporte técnico da brsofthouse - www.brsofthouse.com.br - Email suporte@brsofthouse.com.br", vbYesNo, "Assistente de Registro.")
'        If Ret = 6 Then
'            GoTo E:
'        Else
'
'            nDias = DateDiff("d", Date, Reg_Dados3!data_fin)
'            lblRegistro.Caption = "Registro temporário:"
'            If nDias > 1 Then
'                lblLicenca.Caption = "Faltam " & nDias & " dias."
'            Else
'                lblLicenca.Caption = "Falta " & nDias & " dias."
'            End If
'        End If
'
'    Else
'        lblLicenca.Caption = Reg_Dados!empresa
'    End If
'
'Else
'    Ret = MsgBox("Seu sistema ainda não foi ativado deseja faze-lo agora ???", vbYesNo, "Brsoftware")
'    If Ret = 6 Then
'        frmRegistro.Show
'    Else
'        Exit Sub
'    End If
'End If
'
'Reg_Dados.Close
'Set Reg_Dados = Nothing
'Reg_Dados2.Close
'Set Reg_Dados2 = Nothing
'Reg_Dados3.Close
'Set Reg_Dados3 = Nothing
'
'E:

'Exit Sub
End Sub

Private Sub Timer2_Timer()
CT_Hora.Text = Format(Time, "HH:MM:SS")
End Sub
