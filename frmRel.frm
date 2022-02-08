VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de Relatórios"
   ClientHeight    =   4155
   ClientLeft      =   5445
   ClientTop       =   4920
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7065
   Begin Crystal.CrystalReport CRP 
      Left            =   6180
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton BtAdicionar 
      Caption         =   "Visualizar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   300
      Picture         =   "frmRel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3150
      Width           =   1155
   End
   Begin VB.ComboBox cbRelatorios 
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
      Left            =   120
      TabIndex        =   2
      Top             =   330
      Width           =   4515
   End
   Begin VB.ComboBox cbcolaborador 
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
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4515
   End
   Begin MSMask.MaskEdBox maskDataIni 
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   1860
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox maskDataFin 
      Height          =   375
      Left            =   1590
      TabIndex        =   6
      Top             =   1860
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
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
      Left            =   180
      TabIndex        =   8
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final"
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
      Left            =   1620
      TabIndex        =   7
      Top             =   1620
      Width           =   810
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório"
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
      TabIndex        =   3
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label48 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Parceiros"
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
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtAdicionar_Click()
Dim strRel As String
Dim strCriterio As String

Select Case cbRelatorios
Case "Relação de OS por Parceiro"
    strCriterio = "{osdb.colaborador}='" & cbcolaborador & "'"

Case "Relação de OS por Cliente"

End Select

Set rsTabelas = New ADODB.Recordset
rsTabelas.Open "select * from tbRelatorios where nome='" & cbRelatorios & "'", conn
strRel = rsTabelas!rel
rsTabelas.Close
Set rsTabelas = Nothing

Sql = "DELETE FROM tbperiodorel"
conn.Execute Sql

Sql = "INSERT INTO tbperiodorel (IND, tbdatData_ini, tbdatData_fin) VALUES (" & "0" & ", '" & Format(maskDataIni, "yyyy-mm-dd") & "', '" & Format(maskDataFin, "yyyy-mm-dd") & "')"
conn.Execute Sql

crp.SelectionFormula = ""

If maskDataIni.Text = "__/__/____" Then
    crp.SelectionFormula = strCriterio
Else
    strCriterio = "{osdb.data}>= DateTime (" & Format(maskDataIni, "yyyy,mm,dd") & " , 00, 00, 00) and {osdb.data}<= DateTime (" & Format(maskDataFin, "yyyy,mm,dd") & " , 00, 00, 00) and {osdb.colaborador}='" & cbcolaborador & "' and {osdb.servico}<>'" & "AUTORIZAÇÃO DE USO" & "'"

End If
'DateTime (" & Mdata_ini & " , 00, 00, 00)

'crp.SelectionFormula = "{osdb.id}=" & ctNos & ""
crp.Connect = strDns
crp.ReportFileName = App.Path & "\..\relatorios\" & strRel & ".rpt"
crp.SelectionFormula = strCriterio
crp.Destination = crptToWindow
crp.WindowState = crptMaximized
crp.Action = 1

End Sub

Private Sub cbcolaborador_GotFocus()
   On Error Resume Next
   cbcolaborador.SelStart = 0
   cbcolaborador.SelLength = Len(cbcolaborador.Text)
End Sub

Private Sub Combo1_GotFocus()
   On Error Resume Next
   Combo1.SelStart = 0
   Combo1.SelLength = Len(Combo1.Text)
End Sub

Private Sub cbRelatorios_Click()
Select Case cbRelatorios
Case "Relação de OS por Parceiro"
    Call sbFiltro(1)
Case "Relação de OS por Cliente"
    Call sbFiltro(2)
End Select

End Sub

Private Sub Form_Load()
KeyPreview = True

Call sbAddRelatorios


End Sub
Private Sub sbAddRelatorios()
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "select * from tbrelatorios order by nome", conn
    cbRelatorios.Clear
    Do While Not rsTabelas.EOF
        cbRelatorios.AddItem rsTabelas!nome
        rsTabelas.MoveNext
    Loop
    rsTabelas.Close
    Set rsTabelas = Nothing
End Sub


Private Sub sbFiltro(nFiltroCombo As Integer)
10000   Select Case nFiltroCombo
            Case 1

10005           Set rsTabelas = New ADODB.Recordset
10010           rsTabelas.Open "select * from clientes where titulo='" & "PARCEIRO" & "' order by nome", conn
10015           cbcolaborador.Clear
10020           Do While Not rsTabelas.EOF
10025               cbcolaborador.AddItem rsTabelas!nome
10030               rsTabelas.MoveNext
10035           Loop
10040           rsTabelas.Close
10045           Set rsTabelas = Nothing
            Case 2
    
    
    
10050           Set rsTabelas = New ADODB.Recordset
10055           rsTabelas.Open "select * from clientes where titulo='" & "ADM" & "' order by nome", conn
                'cbcolaborador.Clear
                'Do While Not rsTabelas.EOF
10060           cbcolaborador.AddItem rsTabelas!nome
                '    rsTabelas.MoveNext
                'Loop
10065           rsTabelas.Close
10070           Set rsTabelas = Nothing
    
10075   End Select
End Sub

Private Sub ctCompetencia_GotFocus()
   On Error Resume Next
   ctCompetencia.SelStart = 0
   ctCompetencia.SelLength = Len(ctCompetencia.Text)
End Sub

Private Sub ctCompetencia_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub maskDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyReturn
    maskDataFin.SetFocus
End Select
End Sub
