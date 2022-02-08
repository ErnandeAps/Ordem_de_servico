VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmFlexGrid 
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   555
   ClientTop       =   1365
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   16575
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   6975
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   12303
      _Version        =   393216
      BackColor       =   16744576
      Appearance      =   0
   End
End
Attribute VB_Name = "frmFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlexGrid_DblClick()
Retorno = FlexGrid.TextMatrix(FlexGrid.Row, 0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
KeyPreview = True
Call preencheFlexGrid
End Sub
Public Sub preencheFlexGrid()

'On Error GoTo Macoratti

'verifica os tipos dos objetos

If Not TypeOf FlexGrid Is MSFlexGrid Then Exit Sub

If Not TypeOf rs Is ADODB.Recordset Then Exit Sub

Dim I As Integer
Dim j As Integer

   'define linha e coluna do msflexgrid
    FlexGrid.FixedRows = 1
    FlexGrid.FixedCols = 0
    
    FlexGrid.Cols = 8
    FlexGrid.ColWidth(0) = 800
    FlexGrid.ColWidth(1) = 1600
    FlexGrid.ColWidth(2) = 2500
    FlexGrid.ColWidth(3) = 1000
    FlexGrid.ColWidth(4) = 3500
    FlexGrid.ColWidth(5) = 5000
    FlexGrid.ColWidth(6) = 4500
    FlexGrid.ColWidth(7) = 2500
    FlexGrid.TextMatrix(0, 0) = "N° OS"
    FlexGrid.TextMatrix(0, 1) = "Equipamento"
    FlexGrid.TextMatrix(0, 2) = "N° de Série"
    FlexGrid.TextMatrix(0, 3) = "Data"
    FlexGrid.TextMatrix(0, 4) = "Serviço"
    FlexGrid.TextMatrix(0, 5) = "Cliente"
    FlexGrid.TextMatrix(0, 6) = "Endereço"
    FlexGrid.TextMatrix(0, 7) = "Status"
    FlexGrid.ColAlignment(0) = 3
    FlexGrid.ColAlignment(2) = 1
    FlexGrid.ColAlignment(3) = 3
   'se o recordset tiver dados então...
    
    Conecta_BD
    Set rsTabelas = New ADODB.Recordset
    rsTabelas.Open "Select * From osdb order by id DESC", conn
    
    Dim nCont As Variant
    'Contador
    Do While Not rsTabelas.EOF
        nCont = nCont + 1
        rsTabelas.MoveNext
    Loop
    
    rsTabelas.MoveFirst
   If Not rsTabelas.EOF Then
    
       FlexGrid.Rows = nCont + 1

       Do While Not rsTabelas.EOF
            For I = 1 To nCont
                FlexGrid.TextMatrix(I, 0) = rsTabelas.Fields(0).Value
                FlexGrid.TextMatrix(I, 1) = rsTabelas.Fields(26).Value
                FlexGrid.TextMatrix(I, 2) = rsTabelas.Fields(24).Value
                FlexGrid.TextMatrix(I, 3) = rsTabelas.Fields(19).Value
                FlexGrid.TextMatrix(I, 4) = rsTabelas.Fields(11).Value
                FlexGrid.TextMatrix(I, 5) = rsTabelas.Fields(1).Value
                FlexGrid.TextMatrix(I, 6) = rsTabelas.Fields(4).Value
                FlexGrid.TextMatrix(I, 7) = rsTabelas.Fields(12).Value
                If fColorGrdi(rsTabelas.Fields(12).Value) <> "" Then
                FlexGrid.Row = I
                For j = 0 To 7
                    FlexGrid.Col = j
                    FlexGrid.CellBackColor = fColorGrdi(rsTabelas.Fields(12).Value)
                Next
                    
                End If
           rsTabelas.MoveNext
           Next
       Loop
   End If


Macoratti:

 

   Exit Sub

End Sub
Private Function fColorGrdi(strSituacao As String) As String
Select Case strSituacao
Case "Coleta de equip. em andamento"
        fColorGrdi = &HC0FFFF
        
Case "Aguardando aprovacao cliente"
        fColorGrdi = &HC0FFFF

Case "Ordem de servico fechada"
        fColorGrdi = &HE0E0E0

Case "Aguardando peças"
        fColorGrdi = &H80C0FF

Case "Equip. entregue"
        fColorGrdi = &HFFFFC0
        
Case "Aguardando Cotação Peça"
        fColorGrdi = &H80C0FF

End Select
End Function
