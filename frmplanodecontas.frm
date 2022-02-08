VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmplanodecontas 
   AutoRedraw      =   -1  'True
   Caption         =   "Plano de Contas Caixa Geral"
   ClientHeight    =   5310
   ClientLeft      =   2925
   ClientTop       =   3615
   ClientWidth     =   8550
   ClipControls    =   0   'False
   Icon            =   "frmplanodecontas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8550
   Begin Crystal.CrystalReport CRP 
      Left            =   330
      Top             =   4620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton btFechar 
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6060
      Picture         =   "frmplanodecontas.frx":17B2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1185
   End
   Begin VB.CommandButton BTEXCLUIR 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4890
      Picture         =   "frmplanodecontas.frx":27F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1185
   End
   Begin VB.CommandButton BTATUALIZAR 
      Caption         =   "A&tualizar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3720
      Picture         =   "frmplanodecontas.frx":4F96
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1185
   End
   Begin VB.CommandButton BTCONSULTAR 
      Caption         =   "&Consultar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2550
      Picture         =   "frmplanodecontas.frx":6A08
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1185
   End
   Begin VB.CommandButton BTADICIONAR 
      Caption         =   "&Adicionar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1380
      Picture         =   "frmplanodecontas.frx":7A4A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1185
   End
   Begin VB.TextBox CTDESCRICAO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Left            =   750
      TabIndex        =   2
      Top             =   270
      Width           =   7755
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmplanodecontas.frx":8A8C
      Height          =   3705
      Left            =   0
      OleObjectBlob   =   "frmplanodecontas.frx":8AA0
      TabIndex        =   1
      Top             =   660
      Width           =   8535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSMask.MaskEdBox Maskcod 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   -2147483628
      MaxLength       =   6
      Mask            =   "#.#.##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DESCRICAO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   780
      TabIndex        =   4
      Top             =   30
      Width           =   1170
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "DOC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   480
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuRel 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuPlanoContas 
         Caption         =   "Plano de Contas"
      End
   End
End
Attribute VB_Name = "frmplanodecontas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtAdicionar_Click()
If Maskcod.Text = "_._.__" Then
    MsgBox "Digite o código da conta para continuar.", vbCritical, "Assistente de conta"
    Exit Sub
End If

Set Reg_Dados = BANCO.OpenRecordset("select * from plano_contas where cod='" & Maskcod & "'")
If Not Reg_Dados.EOF Then
    MsgBox "Conta ja cadastrada.", vbCritical, "Assistente de conta"
    Exit Sub
End If

Reg_Dados.AddNew
Reg_Dados!franquiaID = strEmpresa
Reg_Dados!cod = Maskcod
Reg_Dados!descricao = CTDESCRICAO
Reg_Dados.Update
Reg_Dados.Close
Set Reg_Dados = Nothing
Call sbLimpa_Campos(Me)
Maskcod.SetFocus

Data1.Refresh

End Sub

Private Sub BTATUALIZAR_Click()
Set Reg_Dados = BANCO.OpenRecordset("select * from plano_contas where cod='" & Maskcod & "'")
If Reg_Dados.EOF Then
    MsgBox "Conta não cadastrada.", vbCritical, "Assistente de conta"
    Exit Sub
End If

Reg_Dados.Edit
'Reg_Dados!franquiaID = strEmpresa
'Reg_Dados!cod = Maskcod
Reg_Dados!descricao = CTDESCRICAO
Reg_Dados.Update
Reg_Dados.Close
Set Reg_Dados = Nothing
Call sbLimpa_Campos(Me)
BTATUALIZAR.Enabled = False
BtExcluir.Enabled = False
BtAdicionar.Enabled = True


Maskcod.SetFocus

Data1.Refresh
End Sub

Private Sub BtConsultar_Click()
If Maskcod.Text = "_._.__" Then
    MsgBox "Digite o código da conta para continuar.", vbCritical, "Assistente de conta"
    Exit Sub
End If
Set Reg_Dados = BANCO.OpenRecordset("select * from plano_contas where cod='" & Maskcod & "'")
If Reg_Dados.EOF Then
    MsgBox "Conta não cadastrada.", vbCritical, "Assistente de conta"
    Exit Sub
End If
Do While Not Data1.Recordset.EOF
    If Data1.Recordset!cod = Maskcod.Text Then
    
        Exit Sub
    End If
    Data1.Recordset.MoveNext
Loop
Reg_Dados.Close
Set Reg_Dados = Nothing

Data1.Refresh
End Sub

Private Sub BtExcluir_Click()
Set Reg_Dados = BANCO.OpenRecordset("select * from plano_contas where cod='" & Maskcod & "'")
If Reg_Dados.EOF Then
    MsgBox "Conta não cadastrada.", vbCritical, "Assistente de conta"
    Exit Sub
End If

Reg_Dados.Delete
Reg_Dados.Close
Set Reg_Dados = Nothing
Call sbLimpa_Campos(Me)
BTATUALIZAR.Enabled = False
BtExcluir.Enabled = False
BtAdicionar.Enabled = True


Maskcod.SetFocus

Data1.Refresh
End Sub

Private Sub btFechar_Click()
Unload Me
End Sub

Private Sub CTDESCRICAO_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        If CTDESCRICAO.Text = "" Then
            MsgBox "Digite a descrição da conta para continuar.", vbCritical, "Assistente de conta"
            Exit Sub
        End If
        
        BtAdicionar.SetFocus
End Select
End Sub

Private Sub DBGrid1_Click()
Maskcod.Text = DBGrid1.Columns(0)
CTDESCRICAO.Text = DBGrid1.Columns(1)
BtAdicionar.Enabled = False
BTATUALIZAR.Enabled = True
BtExcluir.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape:
        If Maskcod.Text = "_._.__" Then
            Unload Me
        Else
            Call sbLimpa_Campos(Me)
        End If
        
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Data1.DatabaseName = BANCO.Name
Data1.RecordSource = "SELECT * FROM PLANO_CONTAS order by cod"
Data1.Refresh
End Sub

Private Sub Maskcod_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        If Maskcod.Text = "_._.__" Then
            MsgBox "Digite o código da conta para continuar.", vbCritical, "Assistente de conta"
            Exit Sub
        End If
        
        CTDESCRICAO.SetFocus
End Select
End Sub

Private Sub mnuPlanoContas_Click()
On Error GoTo E:
'mped = InputBox("Digite o nùmero do pedido para continuar", "Assistente de impressão")

CRP.Password = Mpassword
CRP.DataFiles(0) = BANCO.Name
CRP.SelectionFormula = ""
CRP.ReportFileName = App.Path & "\..\relatorios\PlanoContas.rpt"
CRP.Destination = crptToWindow
CRP.WindowState = crptMaximized
CRP.Action = 1
E:
If Err.Number = 20515 Then
    MsgBox "Data inválida.", vbCritical, "Assistente de verificação."
    Exit Sub
End If
End Sub

Private Sub MnuSair_Click()
Unload Me
End Sub
