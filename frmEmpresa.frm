VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmEmpresa 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro da Empresa"
   ClientHeight    =   4575
   ClientLeft      =   3120
   ClientTop       =   645
   ClientWidth     =   8805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3645
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   8715
      Begin VB.ComboBox cbTabRegiao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         ItemData        =   "frmEmpresa.frx":26E2
         Left            =   5640
         List            =   "frmEmpresa.frx":26EC
         TabIndex        =   27
         Top             =   450
         Width           =   2925
      End
      Begin VB.TextBox ctNomeFranquia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         MaxLength       =   15
         TabIndex        =   25
         Top             =   1110
         Width           =   1755
      End
      Begin VB.TextBox ctrazao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   90
         TabIndex        =   11
         Top             =   450
         Width           =   5535
      End
      Begin VB.TextBox ctfantasia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   1110
         Width           =   5535
      End
      Begin VB.TextBox ctcgc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   1770
         Width           =   1875
      End
      Begin VB.TextBox ctinsc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   1980
         TabIndex        =   8
         Top             =   1770
         Width           =   1875
      End
      Begin VB.TextBox ctend 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   3870
         TabIndex        =   7
         Top             =   1770
         Width           =   4695
      End
      Begin VB.TextBox ctbairro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   2430
         Width           =   4425
      End
      Begin VB.TextBox ctcidade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   4530
         TabIndex        =   5
         Top             =   2430
         Width           =   3615
      End
      Begin VB.TextBox ctuf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   8160
         TabIndex        =   4
         Top             =   2430
         Width           =   405
      End
      Begin VB.TextBox ctcontato 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   2700
         TabIndex        =   3
         Top             =   3090
         Width           =   5865
      End
      Begin MSMask.MaskEdBox MASKFONE 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3090
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483624
         MaxLength       =   13
         Mask            =   "(##)####-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MASKFAX 
         Height          =   375
         Left            =   1410
         TabIndex        =   13
         Top             =   3090
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483624
         MaxLength       =   13
         Mask            =   "(##)####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tab-Região"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5640
         TabIndex        =   28
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Franquia"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5670
         TabIndex        =   26
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "C.G.C."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1500
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "INSC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2010
         TabIndex        =   21
         Top             =   1500
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3900
         TabIndex        =   20
         Top             =   1500
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4560
         TabIndex        =   18
         Top             =   2160
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Uf"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8190
         TabIndex        =   17
         Top             =   2160
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Contato"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2670
         TabIndex        =   16
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fone Fax :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1380
         TabIndex        =   15
         Top             =   2820
         Width           =   1050
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Fone :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   2820
         Width           =   630
      End
   End
   Begin VB.CommandButton BtFechar 
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4170
      Picture         =   "frmEmpresa.frx":270D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3750
      Width           =   1155
   End
   Begin VB.CommandButton BtAlterar 
      Caption         =   "Al&terar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3030
      Picture         =   "frmEmpresa.frx":374F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3750
      Width           =   1155
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Reg_Dados As Recordset
Private Sub BtAdicionar_Click()

End Sub

Private Sub BtAlterar_Click()
Set Reg_Dados = BANCO.OpenRecordset("select * from cad_empresa")
Reg_Dados.Edit
Reg_Dados!franquiaID = ctNomeFranquia
Reg_Dados!empresa = CTRAZAO
Reg_Dados!n_fantasia = ctfantasia
Reg_Dados!End = CTEND
Reg_Dados!bairro = CTBAIRRO
Reg_Dados!Cidade = CTCIDADE
Reg_Dados!Uf = CTUF
Reg_Dados!cgc = CTCGC
Reg_Dados!insc = CTINSC
Reg_Dados!Fone = MASKFONE
Reg_Dados!fax = MASKFAX
Reg_Dados!contato = CTCONTATO
If cbTabRegiao.Text = "Norte/Nordeste" Then
    Reg_Dados!segat = 0
Else
  Reg_Dados!segat = 1
End If
Reg_Dados!segex = ctsegEx
Reg_Dados.Update
Reg_Dados.Close
Set Reg_Dados = Nothing
'*********************************atualiza tabela de endereço ******************************
Set Reg_Dados = BANCO.OpenRecordset("select * from tbperiodorel")
    If Reg_Dados.EOF Then
        Reg_Dados.AddNew
    Else
        Reg_Dados.Edit
    End If
    Reg_Dados!empresa = CTRAZAO
    Reg_Dados!n_fantasia = ctfantasia
    Reg_Dados!End = CTEND
    Reg_Dados!bairro = CTBAIRRO
    Reg_Dados!Cidade = CTCIDADE
    Reg_Dados!Uf = CTUF
    Reg_Dados!cgc = CTCGC
    Reg_Dados!insc = CTINSC
    Reg_Dados!Fone = MASKFONE
    Reg_Dados!fax = MASKFAX
    Reg_Dados!contato = CTCONTATO
    Reg_Dados.Update
    Reg_Dados.Close
    Set Reg_Dados = Nothing

MsgBox "Informações do Cadastro atualizadas com sucesso.", vbInformation, "Assistente de verificação."
CTRAZAO.SetFocus

End Sub

Private Sub btFechar_Click()
Unload Me
End Sub

Private Sub CTBAIRRO_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    CTCIDADE.SetFocus
End Select
End Sub

Private Sub ctcgc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    CTINSC.SetFocus
End Select
End Sub

Private Sub CTCIDADE_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    CTUF.SetFocus
End Select
End Sub

Private Sub ctcontato_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    BtAlterar.SetFocus
End Select
End Sub

Private Sub ctend_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    CTBAIRRO.SetFocus
End Select
End Sub

Private Sub ctfantasia_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    CTCGC.SetFocus
End Select
End Sub

Private Sub ctinsc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    CTEND.SetFocus
End Select
End Sub

Private Sub ctrazao_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    ctfantasia.SetFocus
End Select
End Sub

Private Sub CTUF_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    MASKFONE.SetFocus
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape:
    Unload Me
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
On Error Resume Next

Set Reg_Dados = BANCO.OpenRecordset("select * from cad_Empresa")
strSegAt = Reg_Dados!segat
Reg_Dados.Close
Set Reg_Dados = Nothing

If strSegAt = 0 Then
    cbTabRegiao.Text = "Norte/Nordeste"
Else
    cbTabRegiao.Text = "Sul/Suldeste"
    
End If

On Error Resume Next
Set Reg_Dados = BANCO.OpenRecordset("select * from cad_empresa")
ctNomeFranquia = Reg_Dados!franquiaID
CTRAZAO = Reg_Dados!empresa
ctfantasia = Reg_Dados!n_fantasia
CTEND = Reg_Dados!End
CTBAIRRO = Reg_Dados!bairro
CTCIDADE = Reg_Dados!Cidade
CTUF = Reg_Dados!Uf
CTCGC = Reg_Dados!cgc
CTINSC = Reg_Dados!insc
MASKFONE = Reg_Dados!Fone
MASKFAX = Reg_Dados!fax
CTCONTATO = Reg_Dados!contato
ctsegAt = Reg_Dados!segat
ctsegEx = Format(Reg_Dados!segex, "##,##0.00")
Reg_Dados.Close
Set Reg_Dados = Nothing

End Sub

Private Sub MASKFAX_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    CTCONTATO.SetFocus
End Select
End Sub

Private Sub MASKFONE_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    MASKFAX.SetFocus
End Select
End Sub

