VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmentprod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulário de Entrada de Produtos"
   ClientHeight    =   8385
   ClientLeft      =   2130
   ClientTop       =   1785
   ClientWidth     =   11130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   2  'Dot
   Icon            =   "frmentprod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11130
   Begin VB.Frame framRel 
      Caption         =   "Período"
      ForeColor       =   &H00800000&
      Height          =   7095
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   11085
      Begin VB.Frame Frame3 
         Height          =   915
         Left            =   150
         TabIndex        =   30
         Top             =   210
         Width           =   2535
         Begin MSMask.MaskEdBox maskDataIni 
            Height          =   375
            Left            =   90
            TabIndex        =   31
            Top             =   390
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483628
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskDataFin 
            Height          =   375
            Left            =   1290
            TabIndex        =   33
            Top             =   390
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483628
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Data final"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1380
            TabIndex        =   34
            Top             =   150
            Width           =   795
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Data inicial"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   180
            TabIndex        =   32
            Top             =   150
            Width           =   945
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11085
      Begin VB.TextBox ctTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   35
         Top             =   6630
         Width           =   1665
      End
      Begin VB.TextBox ctsub_total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   7380
         TabIndex        =   25
         Top             =   1050
         Width           =   1275
      End
      Begin VB.ListBox ltcliente 
         Height          =   1620
         ItemData        =   "frmentprod.frx":17B2
         Left            =   90
         List            =   "frmentprod.frx":17B4
         TabIndex        =   20
         Top             =   780
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.ListBox LTPROD 
         Height          =   1620
         ItemData        =   "frmentprod.frx":17B6
         Left            =   90
         List            =   "frmentprod.frx":17B8
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   7305
      End
      Begin VB.TextBox ctnfiscal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   4650
         TabIndex        =   16
         Top             =   390
         Width           =   1485
      End
      Begin VB.TextBox ctfornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   90
         TabIndex        =   0
         Top             =   390
         Width           =   4545
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmentprod.frx":17BA
         Height          =   5115
         Left            =   90
         OleObjectBlob   =   "frmentprod.frx":17CE
         TabIndex        =   8
         Top             =   1440
         Width           =   10905
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2730
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1590
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox ctcodentrada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   90
         TabIndex        =   4
         Top             =   1050
         Width           =   795
      End
      Begin VB.TextBox ctproduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   1620
         TabIndex        =   3
         Top             =   1050
         Width           =   4935
      End
      Begin VB.TextBox ctquant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   900
         TabIndex        =   2
         Top             =   1050
         Width           =   705
      End
      Begin MSMask.MaskEdBox Maskdata 
         Height          =   375
         Left            =   7290
         TabIndex        =   18
         Top             =   390
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483628
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox ctfrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   6150
         TabIndex        =   21
         Top             =   390
         Width           =   1125
      End
      Begin VB.TextBox ctval_unit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   6570
         TabIndex        =   23
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Total ="
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
         Left            =   8340
         TabIndex        =   36
         Top             =   6690
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Sub. Tot."
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
         Left            =   7410
         TabIndex        =   26
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "V. Unit."
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
         Left            =   6600
         TabIndex        =   24
         Top             =   780
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Frete/imp."
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
         Left            =   6180
         TabIndex        =   22
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nota Fiscal Nº"
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
         Left            =   4710
         TabIndex        =   17
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
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
         TabIndex        =   15
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data"
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
         Left            =   7590
         TabIndex        =   9
         Top             =   150
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cód. P."
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
         TabIndex        =   7
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
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
         Left            =   1650
         TabIndex        =   6
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quant."
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
         Left            =   930
         TabIndex        =   5
         Top             =   780
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1185
      Left            =   0
      TabIndex        =   10
      Top             =   7110
      Width           =   11085
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
         Height          =   795
         Left            =   7410
         Picture         =   "frmentprod.frx":2A25
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton btexc_item 
         Caption         =   "Exclui Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6270
         Picture         =   "frmentprod.frx":3A67
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton BtExcluir 
         Caption         =   "&Excluir"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5130
         Picture         =   "frmentprod.frx":6209
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1155
      End
      Begin Crystal.CrystalReport crp 
         Left            =   210
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton BtConsultar 
         Caption         =   "&Consultar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3990
         Picture         =   "frmentprod.frx":89AB
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton btabrir 
         Caption         =   "&Abrir"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2850
         Picture         =   "frmentprod.frx":99ED
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton BtAdicionar 
         Caption         =   "&Salvar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2850
         Picture         =   "frmentprod.frx":AA2F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Menu marq 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir pedido de entrada"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuConsultar 
         Caption         =   "Consultar pedido de entrada"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuExcluir 
         Caption         =   "Excluir pedido de entrada"
         Shortcut        =   ^X
      End
      Begin VB.Menu msair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mrel 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuRelacaoPedido 
         Caption         =   "Relação de Pedidos"
      End
      Begin VB.Menu mrel1 
         Caption         =   "Relatório de Entrada de mercadoria"
      End
   End
End
Attribute VB_Name = "frmentprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Reg_Dados As Recordset
Dim Reg_Dados2 As Recordset
Dim Reg_Dados3 As Recordset
Dim Mnum As Integer
Dim Mcont As Integer
Dim Vmodo As String
Dim McodF As String
Dim Mfrete As Currency

Private Sub btabrir_Click()
btabrir.Visible = False
BtAdicionar.Enabled = True
BtConsultar.Enabled = False
btexc_item.Enabled = False
BTEXCLUIR.Enabled = False
Call sbDesTravar_Campos(Me)
ctfornecedor.Enabled = False
ctnfiscal.SetFocus
End Sub

Private Sub BtAdicionar_Click()
Dim Ret As Integer
Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & "  and modo='" & Vmodo & "'")
If Reg_Dados.EOF Then
    MsgBox "Impossível atualizar o estoque, você deve lançar um produto para continuar.", vbCritical, "Assistente de verificação."
    ctcodentrada.SetFocus
    Exit Sub
End If
Reg_Dados.Close
Set Reg_Dados = Nothing


Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Mnum & " and modo='" & Vmodo & "'")
If Reg_Dados.EOF Then
    Reg_Dados.AddNew
    Reg_Dados!franquiaID = strEmpresa
    Reg_Dados!op = MOP
    Reg_Dados!modo = Vmodo
    Reg_Dados!Data = MaskDATA
    Reg_Dados!num = Mnum
    Reg_Dados!numf = ctnfiscal
    Reg_Dados!COD_F = McodF
    Reg_Dados!fornecedor = ctfornecedor
    Reg_Dados.Update
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    
    '***********************************************************************************
    Dim McodEntrada As String
    Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & "  and modo='" & Vmodo & "' and num=" & Mnum & "")
    Do While Not Reg_Dados.EOF
        McodEntrada = Reg_Dados!CODENTRADA
        Set Reg_Dados2 = BANCO.OpenRecordset("select * from cad_produto where codigo=" & McodEntrada & "")
        Reg_Dados2.Edit
        Reg_Dados2!quant = Reg_Dados2!quant + Reg_Dados!quant
        Reg_Dados2.Update
        Reg_Dados2.Close
        
        Reg_Dados.MoveNext
    Loop
    Reg_Dados.Close
    Set Reg_Dados = Nothing
        
    MsgBox "Seu estoque foi atualizado com sucesso", vbInformation, "assistente de inclusão"
    Call sbLimpa_Campos(Me)
    MaskDATA = Date
    Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq WHERE modo='" & Vmodo & "' order by num")
    Reg_Dados.MoveLast
    Mnum = Format(Reg_Dados!num + 1, "0000")
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    
    frmentprod.Caption = "Formulário de Entrada de Produtos Pedido Nº " & Format(Mnum, "0000")
    Data1.RecordSource = "select * from desc_movest  where modo='" & Vmodo & "' and num=" & Mnum & ""
    Data1.Refresh
    Call sbTrav_button(Me)
    
    ctfornecedor.Text = "KETURA COSMÉTICOS"
    ctnfiscal.Text = 0
    ctfrete.Text = 0
    btabrir.Visible = True
    BtAdicionar.Enabled = False
    BtConsultar.Enabled = True
    btexc_item.Enabled = False
    BTEXCLUIR.Enabled = False
    btabrir.Enabled = True
    
    btabrir.SetFocus
    
End If
End Sub




Private Sub BtConsultar_Click()
Dim Hnum As String
On Error Resume Next
Hnum = InputBox("Digite o número do pedido", "Assistente de consulta")
Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Hnum & "")
If Reg_Dados.EOF Then
    MsgBox "Pedido de entrada de mercadoria não encontrado", vbCritical, "Assistente de consulta"
    Exit Sub
Else
MaskDATA = Reg_Dados!Data
Mnum = Reg_Dados!num
ctnfiscal = Reg_Dados!numf
ctfornecedor = Reg_Dados!fornecedor
frmentprod.Caption = "Formulário de Entrada de Produtos Pedido Nº " & Format(Mnum, "0000")

Set Reg_Dados = BANCO.OpenRecordset("select sum(sub_total) as vnumSubTotal from desc_movest where num=" & Mnum & "")
ctTotal = Format(Reg_Dados!vnumSubTotal, "##,##0.00")
Reg_Dados.Close
Set Reg_Dados = Nothing

Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Hnum & "")
Data1.RecordSource = "select * from desc_movest where num=" & Mnum & ""
Data1.Refresh
btabrir.Enabled = False
BtConsultar.Enabled = False
If Reg_Dados!modo = "Estornado" Then
    MsgBox "Este pedido se encontra estornado.", vbCritical, "Assistente de verificação"
    Call sbTravar_Campos(Me)
Else
    BtAlterar.Enabled = True
    BTEXCLUIR.Enabled = True
End If
Reg_Dados.Close
Set Reg_Dados = Nothing
End If
End Sub

Private Sub btexc_item_Click()
Dim Mitem As Integer
On Error GoTo E:
Mitem = InputBox("Digite o número do controle do lançamento p/ continuar.", "Assistente de exclusão")
Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & " and cont=" & Mitem & "")
If Reg_Dados.EOF Then
    MsgBox "Ítem não encontrado.", vbCritical, "Assistente de verificação"
    Exit Sub
End If
Mitem = MsgBox("Tem certeza que deseja excluir este ítem do lançamento ???", vbYesNo, "Assistente de verificação")
If Mitem = 6 Then
    Reg_Dados.Delete
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    Data1.Refresh
    ctcodentrada.SetFocus
End If

Set Reg_Dados = BANCO.OpenRecordset("select sum(sub_total) as vnumTotal  from desc_movest where num=" & Mnum & "")
ctTotal = Format(Reg_Dados!vnumTotal, "##,##0.00")
Reg_Dados.Close
Set Reg_Dados = Nothing



Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & "")
If Reg_Dados.EOF Then
    btexc_item.Enabled = False
End If
Reg_Dados.Close
Set Reg_Dados = Nothing

E:
If Err.Number = 13 Then
    Exit Sub
End If
End Sub

Private Sub BtExcluir_Click()
Dim Ret As Integer

Ret = MsgBox("Tem certeza que deseja excluir este pedido de entrada de mercadoria ???", vbYesNo, "Assistente de verificação")
If Ret <> 6 Then
    Exit Sub
End If

Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Mnum & "")
If Not Reg_Dados.EOF Then
    Reg_Dados.Edit
    Reg_Dados!modo = "Estornado"
    Reg_Dados.Update
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    
    Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & "")
    Do While Not Reg_Dados.EOF
    McodEntrada = Reg_Dados!CODENTRADA
    Reg_Dados.Edit
    Reg_Dados!modo = "Estornado"
    Reg_Dados.Update
    
    '***************************atualiza estoque*******************************************
    Set Reg_Dados2 = BANCO.OpenRecordset("select * from cad_produto where codigo=" & McodEntrada & "")
    Reg_Dados2.Edit
    Reg_Dados2!quant = Reg_Dados2!quant - Reg_Dados!quant
    Reg_Dados2.Update
    Reg_Dados2.Close
    
    Reg_Dados.MoveNext
    Loop
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    MsgBox "Pedido de entrada de produtos estornado com sucesso.", vbInformation, "assistente de Exclusão"
    Unload Me
    frmentprod.Show
Else
    MsgBox "Selecione um pedido para continuar", vbCritical, "Assistente de Exclusão"
    Exit Sub
End If
End Sub

Private Sub btFechar_Click()
If btabrir.Visible = True Then
            Unload Me
        Else
            
            Dim Ret As Integer
            Ret = MsgBox("Ao sair do pedido de entrada de mercadoria sem salvar voce perderá todos os lançamentos !!!, deseja sair mesmo assim ???", vbYesNo, "Assistente de verificação")
            If Ret <> 6 Then
                Exit Sub
            End If
            
            Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Mnum & "")
            If Reg_Dados.EOF Then
                Reg_Dados.Close
                Set Reg_Dados = Nothing
                Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & "")
                Do While Not Reg_Dados.EOF
                Reg_Dados.Delete
                Reg_Dados.MoveNext
                Loop
                Reg_Dados.Close
                Set Reg_Dados = Nothing
                Call sbLimpa_Campos(Me)
                MaskDATA = Date
                Data1.Refresh
                btabrir.Visible = True
                btabrir.Enabled = True
                btabrir.SetFocus
            End If
            Unload Me
            'frmentprod.Show
        End If
End Sub

Private Sub Btsalvar_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub ctcodentrada_Click()
ctcodentrada.Text = ""
ctproduto.Text = ""
ctquant.Text = ""
ctval_unit.Text = ""
ctsub_total.Text = ""

End Sub

Private Sub ctcodentrada_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
        Dim McodEntrada As String
        If ctcodentrada.Text = "" Then
            MsgBox "Digite o codigo do produto para continuar", vbCritical, "Assistente de inclusão"
            ctcodentrada.SetFocus
            Exit Sub
        Else
        McodEntrada = Format(ctcodentrada, "0000")
        Set Reg_Dados = BANCO.OpenRecordset("select * from cad_produto where codigo=" & McodEntrada & "")
            If Reg_Dados.EOF Then
                MsgBox "Produto não cadastrado", vbCritical, "Assistente de inclusão"
                ctcodentrada.Text = ""
                ctcodentrada.SetFocus
                Exit Sub
            Else
                ctcodentrada = Format(McodEntrada, "0000")
                ctproduto = Reg_Dados!desci
                ctval_unit = Reg_Dados!VENDAD
                ctquant.SetFocus
            End If
        Reg_Dados.Close
        Set Reg_Dados = Nothing
        End If
End Select
End Sub

Private Sub ctcodentrada_KeyPress(KeyAscii As Integer)
'If ctcodentrada.Text = "" Then
'    Exit Sub
'End If
'LTPROD.Clear
'Mnome = ctcodentrada.Text
'Set Reg_Dados = BANCO.OpenRecordset("SELECT * FROM cad_produto where codigo like '" & Mnome & "*' order by descricao")
'Do While Not Reg_Dados.EOF
'LTPROD.AddItem Reg_Dados!CODigo & "   " & Reg_Dados!descricao
'Reg_Dados.MoveNext
'Loop
'Reg_Dados.Close
'Set Reg_Dados = Nothing
'LTPROD.Visible = True
End Sub

Private Sub ctfornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
    If ctfornecedor.Text = "" Then
        MsgBox "Digite o Fornecedor para continuar.", vbCritical, "Assistente de consulta"
        ctfornecedor.SetFocus
        Exit Sub
    End If
    Dim Ret
'    Set Reg_Dados = BANCO.OpenRecordset("select * from cad_fornecedor where n_fantasia='" & ctfornecedor.Text & "'")
'    If Reg_Dados.EOF Then
'        ret = MsgBox("Este Fornecedor não está cadastrado, deseja seleciona-lo de uma lista ?", vbYesNo, "Assistente de consulta")
'        If ret = 6 Then
'            Reg_Dados.Close
'            Set Reg_Dados = BANCO.OpenRecordset("select * from cad_fornecedor where n_fantasia like '" & ctfornecedor.Text & "*'")
'            ltcliente.Visible = True
'            ltcliente.Clear
'            Do While Not Reg_Dados.EOF
'                ltcliente.AddItem Reg_Dados!n_fantasia
'                Reg_Dados.MoveNext
'            Loop
'            ltcliente.SetFocus
'            Exit Sub
'        End If
'    End If
        ctnfiscal.SetFocus
End Select
End Sub

Private Sub ctfrete_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
        If ctfrete.Text = "" Then
            MsgBox "Digite o valor do frete p/ continuar.", vbCritical, "Assistente de verificação"
            Exit Sub
        End If
        ctfrete = Format(ctfrete, "##,##0.00")
        MaskDATA.SetFocus
End Select

End Sub

Private Sub ctnfiscal_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
        If ctnfiscal.Text = "" Then
            MsgBox "Digite o número da Nota fiscal p/ continuar.", vbCritical, "Assistente de verificação"
            Exit Sub
        End If
        ctfrete.SetFocus
End Select

End Sub

Private Sub ctquant_GotFocus()
LTPROD.Visible = False
End Sub

Private Sub ctQuant_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
        If ctquant.Text = "" Then
            MsgBox "Digite a quantidade p/ continuar.", vbCritical, "Assistente de verificação."
            ctquant.SetFocus
            Exit Sub
        End If
        ctsub_total = Format(ctquant * ctval_unit, "##,##0.00")
        If ctquant.Text = "" Then
            MsgBox "Digite a quantidade para continuar", vbCritical, "assistente de inclusão"
            ctquant.SetFocus
            Exit Sub
        End If
        If ctsub_total.Text = "" Then
            MsgBox "Digite o valor da nota para continuar", vbCritical, "assistente de inclusão"
            ctsub_total.SetFocus
            Exit Sub
        End If
        ctval_unit = Format(ctsub_total / ctquant, "##,##0.00")
        
        Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & " order by cont")
        If Reg_Dados.EOF Then
            Mcont = 1
        Else
            Reg_Dados.MoveLast
            Mcont = Reg_Dados!cont + 1
        End If
        Reg_Dados.AddNew
        Reg_Dados!franquiaID = strEmpresa
        Reg_Dados!Data = MaskDATA
        Reg_Dados!modo = Vmodo
        Reg_Dados!num = Mnum
        Reg_Dados!cont = Mcont
        Reg_Dados!CODENTRADA = ctcodentrada
        Reg_Dados!PRODUTO = ctproduto
        Reg_Dados!historico = ctfornecedor
        Reg_Dados!quant = ctquant
        Reg_Dados!val_unit = ctval_unit
        Reg_Dados!sub_total = Format(ctsub_total, "##,##0.00") 'Format(ctQuant * ctval_unit, "##,##0.00")
        Reg_Dados.Update
        Reg_Dados.Close
        Set Reg_Dados = Nothing
        
        Set Reg_Dados = BANCO.OpenRecordset("select sum(sub_total) as vnumTotal  from desc_movest where num=" & Mnum & "")
        ctTotal = Format(Reg_Dados!vnumTotal, "##,##0.00")
        Reg_Dados.Close
        Set Reg_Dados = Nothing
        
        ctcodentrada.Text = ""
        ctproduto.Text = ""
        ctquant.Text = ""
        ctval_unit.Text = ""
        ctsub_total.Text = ""
        ctcodentrada.SetFocus
        Data1.Refresh
        
        If btabrir.Visible = False Then
            btexc_item.Enabled = True
        End If
        
    Case vbKeyLeft
        ctcodentrada.Text = ""
        ctproduto.Text = ""
        ctquant.Text = ""
        ctval_unit.Text = ""
        ctsub_total.Text = ""
        ctcodentrada.SetFocus
        
        
End Select
End Sub

Private Sub ctsub_total_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
        
'        If ctval_unit.Text = "" Then
'            MsgBox "Digite o valor unitário para continuar", vbCritical, "assistente de inclusão"
'            ctval_unit.SetFocus
'            Exit Sub
'        End If
'        If ctquant.Text = "" Then
'            MsgBox "Digite a quantidade para continuar", vbCritical, "assistente de inclusão"
'            ctquant.SetFocus
'            Exit Sub
'        End If
'        If ctsub_total.Text = "" Then
'            MsgBox "Digite o valor da nota para continuar", vbCritical, "assistente de inclusão"
'            ctsub_total.SetFocus
'            Exit Sub
'        End If
'        ctval_unit = Format(ctsub_total / ctquant, "##,##0.00")
'
'        Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & " order by cont")
'        If Reg_Dados.EOF Then
'            Mcont = 1
'        Else
'            Reg_Dados.MoveLast
'            Mcont = Reg_Dados!cont + 1
'        End If
'        Reg_Dados.AddNew
'        Reg_Dados!Data = Maskdata
'        Reg_Dados!MODO = Vmodo
'        Reg_Dados.num = Mnum
'        Reg_Dados!cont = Mcont
'        Reg_Dados!codentrada = ctcodentrada
'        Reg_Dados!produto = ctproduto
'        Reg_Dados!historico = ctfornecedor
'        Reg_Dados!QUANT = ctquant
'        Reg_Dados!val_unit = ctval_unit
'        Reg_Dados!sub_total = Format(ctsub_total, "##,##0.00") 'Format(ctQuant * ctval_unit, "##,##0.00")
'        Reg_Dados.Update
        '*********************************ATUALIZA VALOR***********************************
'        Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & " order by cont")
'        If Reg_Dados.EOF Then
'            Mcont = 1
'        Else
'            Reg_Dados.MoveLast
'            Mcont = Reg_Dados!cont
'            Reg_Dados.MoveFirst
'        End If
'        Mfrete = Format(ctfrete / Mcont, "##,##0.00")
'        Do While Not Reg_Dados.EOF
'        Reg_Dados.Edit
'        Reg_Dados!frete = Mfrete
'        'Reg_Dados!sub_total = Format((Reg_Dados!val_unit * Reg_Dados!QUANT) + Mfrete, "##,##0.00")
'        Reg_Dados!val_unit_frete = Format(Reg_Dados!val_unit + Mfrete, "##,##0.00")
'        'Reg_Dados!sub_total = Format(((Reg_Dados!val_unit * Reg_Dados!QUANT) + Mfrete), "##,##0.00")
'        Reg_Dados.Update
'        Reg_Dados.MoveNext
'        Loop
'        Reg_Dados.Close
'        Set Reg_Dados = Nothing
        ctcodentrada.Text = ""
        ctproduto.Text = ""
        ctquant.Text = ""
        ctval_unit.Text = ""
        ctsub_total.Text = ""
        ctcodentrada.SetFocus
        Data1.Refresh
        
End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape:
        If btabrir.Visible = True Then
            Unload Me
        Else
            
            Dim Ret As Integer
            Ret = MsgBox("Ao sair do pedido de entrada de mercadoria sem salvar voce perderá todos os lançamentos !!!, deseja sair mesmo assim ???", vbYesNo, "Assistente de verificação")
            If Ret <> 6 Then
                Exit Sub
            End If
            
            Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Mnum & "")
            If Reg_Dados.EOF Then
                Reg_Dados.Close
                Set Reg_Dados = Nothing
                Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & "")
                Do While Not Reg_Dados.EOF
                Reg_Dados.Delete
                Reg_Dados.MoveNext
                Loop
                Reg_Dados.Close
                Set Reg_Dados = Nothing
                Call sbLimpa_Campos(Me)
                MaskDATA = Date
                Data1.Refresh
                btabrir.Visible = True
                btabrir.Enabled = True
                btabrir.SetFocus
            End If
            Unload Me
            'frmentprod.Show
        End If
End Select

End Sub

Private Sub Form_Load()
KeyPreview = True
ctfornecedor.Text = "KETURA COSMÉTICOS"
ctnfiscal.Text = 0
ctfrete.Text = 0
Vmodo = "Entrada"
Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where modo='" & Vmodo & "' order by num")
If Reg_Dados.EOF Then
    Mnum = Format(1, "0000")
Else
    Reg_Dados.MoveLast
    Mnum = Format(Reg_Dados!num + 1, "0000")
End If
frmentprod.Caption = "Formulário de Entrada de Produtos Pedido Nº " & Format(Mnum, "0000")
Reg_Dados.Close
Set Reg_Dados = Nothing
Data1.DatabaseName = BANCO.Name
Data1.RecordSource = "select * from desc_movest  where modo='" & Vmodo & "' and num=" & Mnum & ""
Data1.Refresh
MaskDATA = Date
Call sbTravar_Campos(Me)
btabrir.Visible = True
BtAdicionar.Enabled = False
BtConsultar.Enabled = True

btexc_item.Enabled = False
BTEXCLUIR.Enabled = False
btabrir.Visible = True
btabrir.Enabled = True
End Sub

Private Sub ltcliente_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
    On Error Resume Next
    Set Reg_Dados = BANCO.OpenRecordset("select * from cad_fornecedor where n_fantasia='" & ltcliente.List(ltcliente.ListIndex) & "'")
    ctfornecedor = Reg_Dados!n_fantasia
    McodF = Reg_Dados!cod
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    ltcliente.Visible = False
    ctcodentrada.SetFocus
End Select
End Sub

Private Sub LTPROD_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
    Set Reg_Dados = BANCO.OpenRecordset("select * from cad_fornecedor where n_fantasia='" & ltcliente.List(ltcliente.ListIndex) & "'")
    ctfornecedor = Reg_Dados!n_fantasia
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    ltcliente.Visible = False
    ctcodentrada.SetFocus
End Select
End Sub

Private Sub Maskdata_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn:
        ctcodentrada.SetFocus
End Select

End Sub

Private Sub maskdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
Dim Mdatini As String
Dim Mdatfin As String
On Error GoTo E:
If MaskDataFin.Text = "__/__/____" Then
    MsgBox "Digite a data final da consulta para continuar.", vbCritical, "Assistente de verificação"
    Exit Sub
End If
Mdatini = MaskDataIni
Mdatfin = MaskDataFin
'**********************atualiza data rel*******************************************
Set Reg_Dados = BANCO.OpenRecordset("select * from tbperiodorel")
Reg_Dados.Edit
Reg_Dados!tbdatData_ini = Mdatini
Reg_Dados!tbdatData_fin = Mdatfin
Reg_Dados.Update
Reg_Dados.Close
Set Reg_Dados = Nothing


framRel.Visible = False
crp.Password = Mpassword
crp.DataFiles(0) = BANCO.Name
crp.SelectionFormula = ""
crp.ReportFileName = App.Path & "\..\relatorios\relentmerc.rpt"
crp.SelectionFormula = " {desc_movest.data} >= Date (" & Format(Mdatini, "yyyy,mm,dd") & ") and  {desc_movest.data} <= Date (" & Format(Mdatfin, "yyyy,mm,dd") & ")and {desc_movest.modo}= '" & Vmodo & "'"
crp.Destination = crptToWindow
crp.WindowState = crptMaximized
crp.Action = 1
Call sbLimpa_Campos(Me)
E:
If Err.Number = 20515 Then
    MsgBox "Data inválida, digite novamente.", vbCritical, "Assistente de verificação"
    Exit Sub
End If

End Select
End Sub

Private Sub maskdataini_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
If MaskDataIni.Text = "__/__/____" Then
    MsgBox "Digite a data inicial da consulta para continuar.", vbCritical, "Assistente de verificação"
    Exit Sub
End If
MaskDataFin.SetFocus
End Select
End Sub

Private Sub mnuAbrir_Click()
btabrir.Visible = False
BtAdicionar.Enabled = True
BtConsultar.Enabled = False
btexc_item.Enabled = False
BTEXCLUIR.Enabled = False
Call sbDesTravar_Campos(Me)
ctfornecedor.Enabled = False
ctnfiscal.SetFocus
End Sub

Private Sub mnuConsultar_Click()
Dim Hnum As String
On Error Resume Next
Hnum = InputBox("Digite o número do pedido", "Assistente de consulta")
Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Hnum & "")
If Reg_Dados.EOF Then
    MsgBox "Pedido de entrada de mercadoria não encontrado", vbCritical, "Assistente de consulta"
    Exit Sub
Else
MaskDATA = Reg_Dados!Data
Mnum = Reg_Dados!num
ctnfiscal = Reg_Dados!numf
ctfornecedor = Reg_Dados!fornecedor
frmentprod.Caption = "Formulário de Entrada de Produtos Pedido Nº " & Format(Mnum, "0000")

Set Reg_Dados = BANCO.OpenRecordset("select sum(sub_total) as vnumSubTotal from desc_movest where num=" & Mnum & "")
ctTotal = Format(Reg_Dados!vnumSubTotal, "##,##0.00")
Reg_Dados.Close
Set Reg_Dados = Nothing


Data1.RecordSource = "select * from desc_movest where num=" & Mnum & ""
Data1.Refresh
btabrir.Enabled = False
BtConsultar.Enabled = False
If Reg_Dados!modo = "Estornado" Then
    MsgBox "Este pedido se encontra estornado.", vbCritical, "Assistente de verificação"
    Call sbTravar_Campos(Me)
Else
    BtAlterar.Enabled = True
    BTEXCLUIR.Enabled = True
End If
Reg_Dados.Close
Set Reg_Dados = Nothing
End If
BTEXCLUIR.Enabled = True
End Sub

Private Sub mnuExcluir_Click()
Dim Ret As Integer

Ret = MsgBox("Tem certeza que deseja excluir este pedido de entrada de mercadoria ???", vbYesNo, "Assistente de verificação")
If Ret <> 6 Then
    Exit Sub
End If

Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Mnum & "")
If Not Reg_Dados.EOF Then
    Reg_Dados.Edit
    Reg_Dados!modo = "Estornado"
    Reg_Dados.Update
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    
    Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & "")
    Do While Not Reg_Dados.EOF
    McodEntrada = Reg_Dados!CODENTRADA
    Reg_Dados.Edit
    Reg_Dados!modo = "Estornado"
    Reg_Dados.Update
    
    '***************************atualiza estoque*******************************************
    Set Reg_Dados2 = BANCO.OpenRecordset("select * from cad_produto where codigo=" & McodEntrada & "")
    Reg_Dados2.Edit
    Reg_Dados2!quant = Reg_Dados2!quant - Reg_Dados!quant
    Reg_Dados2.Update
    Reg_Dados2.Close
    
    Reg_Dados.MoveNext
    Loop
    Reg_Dados.Close
    Set Reg_Dados = Nothing
    MsgBox "Pedido de entrada de produtos estornado com sucesso.", vbInformation, "assistente de Exclusão"
    Unload Me
    frmentprod.Show
Else
    MsgBox "Selecione um pedido para continuar", vbCritical, "Assistente de Exclusão"
    Exit Sub
End If
End Sub

Private Sub mnuRelacaoPedido_Click()
crp.Password = Mpassword
crp.DataFiles(0) = BANCO.Name
crp.SelectionFormula = ""
crp.ReportFileName = App.Path & "\..\relatorios\RelacaoPedEnt.rpt"
'crp.SelectionFormula = " {desc_movest.data} >= Date (" & Format(Mdatini, "yyyy,mm,dd") & ") and  {desc_movest.data} <= Date (" & Format(Mdatfin, "yyyy,mm,dd") & ")and {desc_movest.modo}= '" & Vmodo & "'"
crp.Destination = crptToWindow
crp.WindowState = crptMaximized
crp.Action = 1
End Sub

Private Sub mrel1_Click()
'Mdatini = Format(InputBox("Digite a Data Inicial da consulta segundo Exemplo : " & Date & "", "Assistente de Consulta"), "yyyy,mm,dd")
'Mdatfin = Format(InputBox("Digite a Data Final da consulta segundo Exemplo : " & Date & "", "Assistente de Consulta"), "yyyy,mm,dd")
framRel.Visible = True
MaskDataIni.Enabled = True
MaskDataFin.Enabled = True
MaskDataIni.SetFocus
End Sub

Private Sub msair_Click()
If btabrir.Visible = True Then
            Unload Me
        Else
            
            Dim Ret As Integer
            Ret = MsgBox("Ao sair do pedido de entrada de mercadoria sem salvar voce perderá todos os lançamentos !!!, deseja sair mesmo assim ???", vbYesNo, "Assistente de verificação")
            If Ret <> 6 Then
                Exit Sub
            End If
            
            Set Reg_Dados = BANCO.OpenRecordset("select * from mov_estoq where num=" & Mnum & "")
            If Reg_Dados.EOF Then
                Reg_Dados.Close
                Set Reg_Dados = Nothing
                Set Reg_Dados = BANCO.OpenRecordset("select * from desc_movest where num=" & Mnum & "")
                Do While Not Reg_Dados.EOF
                Reg_Dados.Delete
                Reg_Dados.MoveNext
                Loop
                Reg_Dados.Close
                Set Reg_Dados = Nothing
                Call sbLimpa_Campos(Me)
                MaskDATA = Date
                Data1.Refresh
                btabrir.Visible = True
                btabrir.Enabled = True
                btabrir.SetFocus
            End If
            Unload Me
            'frmentprod.Show
        End If
End Sub
