Attribute VB_Name = "mdlFuncoesDiversas"
Private Const MAX_FILENAME_LEN = 256
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Public M_seg_frm As Variant
Public MrDias As Variant
Public MrMult As Variant
Public MrReg As Variant
Public VsaldoEstoque As Variant
Option Explicit
Public Sub sbLimpa_Campos(ByVal frmForm As Form)
   Dim sMask As String
   Dim j As Integer
    
    For j = 0 To frmForm.Controls.Count - 1
        If TypeOf frmForm.Controls(j) Is TextBox Then
            frmForm.Controls(j).Text = ""
            frmForm.Controls(j).BackColor = &HFFFFFF
        ElseIf TypeOf frmForm.Controls(j) Is MaskEdBox Then
            sMask = frmForm.Controls(j).Mask
            frmForm.Controls(j).Mask = "": frmForm.Controls(j).Text = ""
            frmForm.Controls(j).Mask = sMask
            frmForm.Controls(j).BackColor = &HFFFFFF
        ElseIf TypeOf frmForm.Controls(j) Is ComboBox Then
            frmForm.Controls(j).Text = ""
            frmForm.Controls(j).BackColor = &HFFFFFF
        End If
    Next j
End Sub
Public Sub sbTravar_Campos(ByVal frmForm As Form)
       Dim I As Integer
       
       For I = 0 To frmForm.Controls.Count - 1
           
           If TypeOf frmForm.Controls(I) Is TextBox Then
              frmForm.Controls(I).Enabled = False
           End If
           
           If TypeOf frmForm.Controls(I) Is MaskEdBox Then
              frmForm.Controls(I).Enabled = False
           End If
           
           If TypeOf frmForm.Controls(I) Is ComboBox Then
              frmForm.Controls(I).Enabled = False
           End If
                      
       Next I
End Sub
Public Sub sbDesTravar_Campos(ByVal frmForm As Form)
        
       Dim I As Integer
       
       For I = 0 To frmForm.Controls.Count - 1
           
           If TypeOf frmForm.Controls(I) Is TextBox Then
              frmForm.Controls(I).Enabled = True
           End If
           
           If TypeOf frmForm.Controls(I) Is MaskEdBox Then
              frmForm.Controls(I).Enabled = True
           End If
                      
            If TypeOf frmForm.Controls(I) Is ComboBox Then
              frmForm.Controls(I).Enabled = True
           End If
       Next I
End Sub

Public Function fcWhereOuAnd(ByRef sString As String) As String
   
    On Error GoTo erro
    
    If UCase(sString) Like "*WHERE*" Then
        fcWhereOuAnd = " AND "
    Else
        fcWhereOuAnd = " Where "
    End If
     
'    If UCase(sString) Like "*AND*" Then
'       fcWhereOuAnd = " AND "
'    End If
    
    Exit Function
erro:
    MsgBox Err.Number & ", " & Err.Description, vbCritical, "Erro"
End Function

Public Function desturn()
Select Case Time
Case Is <= "12:00"
    mturno = "Dia"
Case Is < "18:00"
    mturno = "Tarde"
Case Is >= "18:00"
    mturno = "Noite"
End Select

End Function

Public Function SbDiasemana()
Select Case Weekday(Date)
    Case 1
        MDIASEMANA = "Domingo"
    Case 2
        MDIASEMANA = "Segunda"
    Case 3
        MDIASEMANA = "Terça"
    Case 4
        MDIASEMANA = "Quarta"
    Case 5
        MDIASEMANA = "Quinta"
    Case 6
        MDIASEMANA = "Sexta"
    Case 7
        MDIASEMANA = "Sábado"
End Select
End Function

Public Function SBMEN()
MsgBox "Não habilitado Versão para demonstração, consulte o Suporte técnico pelo fone: 30913302/34326216/91295280", vbCritical, "Assistente MedeSistem"
End Function

Public Function trat_turn()
Select Case Time
Case Is <= "12:00"
    Mtrat_turn = "Bom Dia"
Case Is < "18:00"
    Mtrat_turn = "Boa Tarde"
Case Is >= "18:00"
    Mtrat_turn = "Boa Noite"
End Select
End Function

Public Function SbSaldo_conta(ByVal mcod_conta As String) As Currency
Dim Reg_soma As Recordset
Dim Mdebito As Currency
Dim Mcredito As Currency
Dim Msaldo_conta As Variant
Msaldo_conta = 0
'Set TBSOMADEBITO = BANCO.OpenRecordset("select SUM(DEBITO) AS VDEBITO from CAIXA WHERE MONTH(DATA)=" & MMES & " AND Year(DATA)=" & ctcomp & "")
On Error Resume Next
Set Reg_soma = BANCO.OpenRecordset("Select SUM(DEBITO) as Vdebito from caixa where doc='" & mcod_conta & "'")
If Not Reg_soma.EOF Then
    Mdebito = Reg_soma!VDEBITO
    Reg_soma.Close
    Set Reg_soma = Nothing
End If
Set Reg_soma = BANCO.OpenRecordset("Select SUM(credito) as Vdebito from caixa where doc='" & mcod_conta & "'")
Mcredito = Reg_soma!VDEBITO
Reg_soma.Close
Set Reg_soma = Nothing

Msaldo_conta = Mdebito - Mcredito

End Function
Public Function GetSerialNumber(sDrive As String) As Long
 
Dim ser As Long
Dim s As String * MAX_FILENAME_LEN
Dim s2 As String * MAX_FILENAME_LEN
Dim I As Long
Dim j As Long

Call GetVolumeInformation(sDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, I, j, s2, MAX_FILENAME_LEN)
'GetSerialNumber = ser
'Mserial = ser
End Function

Public Function SbDecod_serial(strText As Variant, strText2 As Variant, strText3 As Variant, strText4 As Variant, strserial01 As Variant, strserial02 As Variant, strserial03 As Variant, strserial04 As Variant, ByVal strPwd As String)
Dim M1 As String
Dim M2 As String
Dim M3 As String
Dim M4 As String
Dim Vm As Variant
Dim Vm2 As Variant
Dim Vm3 As Variant
Dim Vm4 As Variant
Dim Mct1 As Variant
Dim Mct2 As Variant
Dim Mct3 As Variant
Dim Mct4 As Variant

M4 = Right(strText4, 4)
'Vm = Right(cte02, 2) & cte03 & Left(cte04, 1)
Vm = Right(strText2, 2) & strText3 & Left(strText4, 1)
Vm2 = strText & Left(strText2, 3)
M1 = Left(Vm2, 4)
M2 = Right(M1, 2)
M1 = Left(M1, 2)

M3 = Right(Vm2, 4)
M4 = Right(M3, 2)
M3 = Left(M3, 2)

Mct1 = M1
Mct2 = M2
Mct3 = M3
Mct4 = M4

Vm3 = Right(strText2, 2) & Left(strText3, 4)
Vm4 = Right(strText3, 1) & strText4
Mct1 = Mct1 & Left(Vm3, 3)
Mct2 = Mct2 & Right(Vm3, 3)
Mct3 = Mct3 & Left(Vm4, 3)
Mct4 = Mct4 & Right(Vm4, 3)

If strserial01 <> Mct1 Or strserial03 <> Mct3 Or strserial04 <> Mct4 Or Right(strserial02, 2) = Right(Mct2, 2) Then
    MsgBox "Serial inválido, digite novamente", vbCritical, "Assistente de registro do sistema"
    MrDias = 0
    MrMult = 0
    MrReg = 0
    Exit Function
End If

MrDias = Left(Mct2, 1) * 10
MrMult = Left(Mct2, 3)
MrMult = Right(MrMult, 2)
MrReg = Mct2

End Function

Public Function SbErr(ByVal frmForm As Form)
Dim Mcont  As Integer
erro:
If Err.Number <> 0 Then
Set Reg_Erro = BANCO.OpenRecordset("select * from Tb_erro order by cont")
If Reg_Erro.EOF Then
    Mcont = 1
Else
    Reg_Erro.MoveLast
    Mcont = Reg_Erro!cont + 1
End If
Reg_Erro.AddNew
Reg_Erro!cont = Mcont
Reg_Erro!d_data = Date
Reg_Erro!T_form = frmForm.Name
Reg_Erro!T_controle = frmForm.ActiveControl
Reg_Erro!T_erro = Err.Number
Reg_Erro!T_desc = Err.Description
Reg_Erro.Update
Reg_Erro.Close
Set Reg_Erro = Nothing

    'MsgBox Err.Number & frmForm.Name & ", " & Err.Description, vbCritical, "Erro"
'Open App.Path & "\..\log\Logerro.txt" For Output As #1
   'Open "TESTFILE" For Output As #1   ' Open file for output.
  
   
   'Write #1,
'   Write #1, Date & " " & frmForm.Name & " " & frmForm.ActiveControl & " " & Err.Number
'   Close #1
    
End If

End Function
Public Sub sbTrav_button(ByVal frmForm As Form)
On Error Resume Next
Set Reg_Seg = BANCO.OpenRecordset("select * from assist_seg where n_cod_usu=" & Mcod_usu & " and frm='" & frmForm.Name & "'")
If Reg_Seg.EOF Then
    MsgBox "Permissão negada, consulte o administrador do sistema.", vbCritical, "Assistente de verificação"
    M_seg_frm = 2
    Exit Sub
End If
M_seg_frm = 1
M_inc = Reg_Seg!n_inc
M_con = Reg_Seg!n_con
M_at = Reg_Seg!n_at
M_exc = Reg_Seg!n_exc
Reg_Seg.Close
Set Reg_Seg = Nothing
'************************************************incluir**************************
If M_inc = 1 Then
    frmForm.Controls("Btsalvar").Enabled = True
Else
    frmForm.Controls("Btsalvar").Enabled = False
End If
'************************************************consultar**************************
If M_con = 1 Then
    frmForm.Controls("Btconsultar").Enabled = True
Else
    frmForm.Controls("Btconsultar").Enabled = False
End If
'************************************************atualizar**************************
If M_at = 1 Then
    frmForm.Controls("Btalterar").Enabled = True
Else
    frmForm.Controls("Btalterar").Enabled = False
End If
'************************************************Excluir**************************
If M_exc = 1 Then
    frmForm.Controls("Btexcluir").Enabled = True
Else
    frmForm.Controls("Btexcluir").Enabled = False
End If
End Sub
Public Sub sbDesTrav_button_Campos(ByVal frmForm As Form)

End Sub

Public Function AtivaForm(ByVal frmForm As Form) As Integer
'Public Function AtivaForm(ByVal frmForm As Form, vForm As String) As Integer
Dim iTop As Integer
Dim iLeft As Integer

iTop = ((11670 - frmForm.Height) \ 2)
iLeft = ((19320 - frmForm.Width) \ 2)
frmForm.Move iLeft, iTop
   
frmForm.Show (1)
'Set Reg_Dados = BANCO.OpenRecordset("select * from assist_seg where t_login='" & MOP & "' and frm='" & vForm & "'")
'If Reg_Dados.EOF Then
'    AtivaForm = 0
'    MsgBox "Acesso negado, consulte o administrador do sistema.", vbCritical, "Assistente de verificação"
'    Reg_Dados.Close
'    Set Reg_Dados = Nothing
'    Exit Function
'Else
'    Reg_Dados.Close
'    Set Reg_Dados = Nothing
'    frmForm.Show (0)
    '*************************************trava butões***************************************
'    M_seg_frm = 1
'    M_inc = Reg_Dados!n_inc
'    M_con = Reg_Dados!n_con
'    M_at = Reg_Dados!n_at
'    M_exc = Reg_Dados!n_exc
    
    '************************************************incluir**************************
'    If M_inc = 1 Then
'        frmForm.Controls("Btsalvar").Enabled = True
'    Else
'        frmForm.Controls("Btsalvar").Enabled = False
'    End If
'    '************************************************consultar**************************
'    If M_con = 1 Then
'        frmForm.Controls("Btconsultar").Enabled = True
'    Else
'        frmForm.Controls("Btconsultar").Enabled = False
'    End If
'    '************************************************atualizar**************************
'    If M_at = 1 Then
'        frmForm.Controls("Btalterar").Enabled = True
'    Else
'        frmForm.Controls("Btalterar").Enabled = False
'    End If
'    '************************************************Excluir**************************
'    If M_exc = 1 Then
'        frmForm.Controls("Btexcluir").Enabled = True
'    Else
'        frmForm.Controls("Btexcluir").Enabled = False
'    End If
'    End If
End Function

Public Function FatEstoque(numCodProd As String) As Variant
Dim Vqtdent As Variant
Dim Vqtdsaid As Variant
Dim VqtdsaidEstoque As Variant
Dim Vmodo As String

'**************************************Verifica a quantidade de entrada de produtos*************************************
Vmodo = "Entrada"
Set Reg_Dados3 = BANCO.OpenRecordset("select * from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "'")
If Reg_Dados3.EOF Then
    Vqtdent = 0
Else

    Set Reg_Dados3 = BANCO.OpenRecordset("select sum(quant) as Vqtd from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "'")
    Vqtdent = Reg_Dados3!vQtd
End If

Reg_Dados3.Close
Set Reg_Dados3 = Nothing

'**************************************Verifica a quantidade de saida de produtos*************************************
Vmodo = "Saida"
Set Reg_Dados3 = BANCO.OpenRecordset("select * from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "'")
If Reg_Dados3.EOF Then
    Vqtdsaid = 0
Else
    Set Reg_Dados3 = BANCO.OpenRecordset("select sum(quant) as Vqtd from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "'")
    
    Vqtdsaid = Reg_Dados3!vQtd
End If
Reg_Dados3.Close
Set Reg_Dados3 = Nothing


'**************************************Verifica a quantidade de saida avulsa de produtos*************************************
Vmodo = "SaidaEstoque"
Set Reg_Dados3 = BANCO.OpenRecordset("select * from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "'")
If Reg_Dados3.EOF Then
    VqtdsaidEstoque = 0
Else
    Set Reg_Dados3 = BANCO.OpenRecordset("select sum(quant) as Vqtd from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "'")
    
    VqtdsaidEstoque = Reg_Dados3!vQtd
End If
Reg_Dados3.Close
Set Reg_Dados3 = Nothing

VsaldoEstoque = Vqtdent - Vqtdsaid - VqtdsaidEstoque


End Function

Public Function VerificaPedAberto(vnumPed As Integer) As String

Set Reg_Estoq = BANCO.OpenRecordset("select sum(qtdtemp) as vnumSit from desc_ped where num=" & vnumPed & " and modo<>'" & "Estornado" & "'")
If Reg_Estoq!vnumSit > 0 Then

    VerificaPedAberto = "Aberto"
Else

    VerificaPedAberto = "Fechado"
End If

Reg_Estoq.Close
Set Reg_Estoq = Nothing
End Function

Public Sub geraArquivotxt(strTabela As String, strFranquia As String)
Call Conecta_BD

rsTabelas.Open "Select * from tbcadtabelas where tbstrexportar='" & "SIM" & "'", conn

Do While Not rsTabelas.EOF

    'rs.Open "Select * from " & rsTabelas!tbstrnome & " , Conn, adOpenForwardOnly, adLockReadOnly, adCmdText"
    'rs.Open "select * from " & rsTabelas!tbstrnome, Conn
    rs.Open "select * from " & strTabela, conn
    
    'Open App.Path & "\..\ftp\outbox\" & strEmpresa & "_" & rsTabelas!tbstrnome & ".txt" For Output As #1
    'Print #1, strEmpresa & ";" & rsTabelas!tbstrnome & ";" & Date & ";" & Time
    Open App.Path & "\..\ftp\outbox\" & strEmpresa & "_" & strTabela & ".txt" For Output As #1
    Print #1, strEmpresa & ";" & strTabela & ";" & Date & ";" & Time
    
    
    
    Do Until rs.EOF
        'Print #1, rs.GetString(, 100, vbTab, vbCrLf, "");
        Print #1, rs.GetString(, 100, ";", vbCrLf, "");
    Loop
    
    Close #1
    
    rs.Close
    Set rs = Nothing
    
    rsTabelas.MoveNext
Loop
    
rsTabelas.Close
Set rsTabelas = Nothing

Call Desconecta_BD
End Sub

Public Function Conecta_BD()
    Set conn = New ADODB.Connection
    'Set rsTabelas = New ADODB.Recordset
    conn.Open strDns
End Function

Public Function Conecta_BDRemoto()
    Set connRemoto = New ADODB.Connection
    'Set rsTabelas = New ADODB.Recordset
    connRemoto.Open strDnsRemoto
End Function

Public Function Desconecta_BD()
  conn.Close
  Set conn = Nothing
End Function

Public Function Desconecta_BDRemoto()
  connRemoto.Close
  Set connRemoto = Nothing
End Function
Public Function DescMes(vnumMes As Integer) As String
Select Case vnumMes
Case 1
    DescMes = "JANEIRO"
Case 2
    DescMes = "FEVEREIRO"
Case 3
    DescMes = "MARÇO"
Case 4
    DescMes = "ABRIL"
Case 5
    DescMes = "MAIO"
Case 6
    DescMes = "JUNHO"
Case 7
    DescMes = "JULHO"
Case 8
    DescMes = "AGOSTO"
Case 9
    DescMes = "SETEMBRO"
Case 10
    DescMes = "OUTUBRO"
Case 11
    DescMes = "NOVEMBRO"
Case 12
    DescMes = "DEZEMBRO"
   
End Select
End Function


'Public Sub sbVerificaDebidoFranquia()
'Set REG_STATUS = BANCO.OpenRecordset("select * from cad_distribuidor order by cod")
'Do While Not REG_STATUS.EOF
'    Set Reg_Dados = BANCO.OpenRecordset("Select * from tb_lancheq where t_distribuidor='" & REG_STATUS!N_FANTASIA & "' and t_sit='" & "Devolvido" & "'")
'    If Not Reg_Dados.EOF Then
'        Reg_Dados.MoveLast
'        Reg_Dados.MoveFirst
'        If Reg_Dados.RecordCount > 1 Then
'            REG_STATUS.Edit
'            REG_STATUS!statusCredito = "Impedimento de compra"
'            REG_STATUS.Update
'        Else
'            REG_STATUS.Edit
'            REG_STATUS!statusCredito = "Com restrição de Crédito"
'            REG_STATUS.Update
'        End If
'    Else
'            REG_STATUS.Edit
'            REG_STATUS!statusCredito = "Sem restrição de Crédito"
'            REG_STATUS.Update
'    End If
'    Reg_Dados.Close
'    Set Reg_Dados = Nothing
'    REG_STATUS.MoveNext
'Loop
'
'REG_STATUS.Close
'Set REG_STATUS = Nothing
'End Sub

Public Function incPontuacao(vnumCodpontuacao As Integer, vstrFranquiado As String)
Set Reg_Dados3 = BANCO.OpenRecordset("select * from cadPontuacao where codigo=" & vnumCodpontuacao & "")
Set Reg_Dados2 = BANCO.OpenRecordset("select * from tbPontuacao")
Reg_Dados2.AddNew
Reg_Dados2!op = MOP
Reg_Dados2!Data = Date
Reg_Dados2!Franquiado = vstrFranquiado
Reg_Dados2!Historico = Reg_Dados3!Descricao
If Reg_Dados3!Valor > 0 Then
    Reg_Dados2!Credito = Reg_Dados3!Valor
    Reg_Dados2!Debito = 0
Else
    Reg_Dados2!Credito = 0
    Reg_Dados2!Debito = Reg_Dados3!Valor

End If
Reg_Dados2.Update
Reg_Dados2.Close
Reg_Dados3.Close
Set Reg_Dados2 = Nothing
Set Reg_Dados3 = Nothing
End Function

Public Function VerifPromocao(numCodigo As Integer, numPedido As Variant) As Variant
Set Reg_Function = BANCO.OpenRecordset("Select * from desc_ped where num=" & numPedido & " and codEntrada=" & numCodigo & "")
If Reg_Function.EOF Then
    VerifPromocao = 0
Else
    VerifPromocao = Reg_Function!Valor
End If
Reg_Function.Close
Set Reg_Function = Nothing
End Function

Public Sub validaPromocao(vnumPedido As Variant, strNomeFranquiado As String)
Dim nValor As Variant
Dim somaPromocao As Variant
Dim vQtdPontuacao As Integer
Dim I As Integer
nValor = 0

Set Reg_Dados3 = BANCO.OpenRecordset("Select * from tbPromocao order by tbpronumCod")
Do While Not Reg_Dados3.EOF
    If VerifPromocao(Reg_Dados3!tbPronumCod, vnumPedido) >= Reg_Dados3!tbPronumValor Then
        nValor = nValor + VerifPromocao(Reg_Dados3!tbPronumCod, vnumPedido)
    End If
    Reg_Dados3.MoveNext
Loop
Reg_Dados3.Close
Set Reg_Dados3 = Nothing

Set Reg_Dados3 = BANCO.OpenRecordset("Select sum(tbpronumValor) as vsomaPromocao from tbpromocao")
somaPromocao = Reg_Dados3!vsomaPromocao
Reg_Dados3.Close
Set Reg_Dados3 = Nothing

If nValor - somaPromocao > 1 Then

    vQtdPontuacao = nValor / somaPromocao
Else
    vQtdPontuacao = 0
End If

If vQtdPontuacao > 0 Then
    
    For I = 1 To vQtdPontuacao
        Call incPontuacao(4, strNomeFranquiado)
    Next
End If

End Sub

Public Function FatEstoqueP(numCodProd As String, vDate As Variant) As Variant
Dim Vqtdent As Variant
Dim Vqtdsaid As Variant
Dim VqtdsaidEstoque As Variant
Dim Vmodo As String

'**************************************Verifica a quantidade de entrada de produtos*************************************
Vmodo = "Entrada"
Set Reg_Dados3 = BANCO.OpenRecordset("select * from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "' and data <=#" & Format(vDate, "mm/dd/yyyy") & "#")
If Reg_Dados3.EOF Then
    Vqtdent = 0
Else

    Set Reg_Dados3 = BANCO.OpenRecordset("select sum(quant) as Vqtd from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "' and data <=#" & Format(vDate, "mm/dd/yyyy") & "#")
    Vqtdent = Reg_Dados3!vQtd
End If

Reg_Dados3.Close
Set Reg_Dados3 = Nothing

'**************************************Verifica a quantidade de saida de produtos*************************************
Vmodo = "Saida"
Set Reg_Dados3 = BANCO.OpenRecordset("select * from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "' and data <=#" & Format(vDate, "mm/dd/yyyy") & "#")
If Reg_Dados3.EOF Then
    Vqtdsaid = 0
Else
    Set Reg_Dados3 = BANCO.OpenRecordset("select sum(quant) as Vqtd from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "' and data <=#" & Format(vDate, "mm/dd/yyyy") & "#")
    
    Vqtdsaid = Reg_Dados3!vQtd
End If
Reg_Dados3.Close
Set Reg_Dados3 = Nothing


'**************************************Verifica a quantidade de saida avulsa de produtos*************************************
Vmodo = "SaidaEstoque"
Set Reg_Dados3 = BANCO.OpenRecordset("select * from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "' and data <=#" & Format(vDate, "mm/dd/yyyy") & "#")
If Reg_Dados3.EOF Then
    VqtdsaidEstoque = 0
Else
    Set Reg_Dados3 = BANCO.OpenRecordset("select sum(quant) as Vqtd from desc_movest where codentrada=" & numCodProd & " and modo='" & Vmodo & "'and data <=#" & Format(vDate, "mm/dd/yyyy") & "#")
    
    VqtdsaidEstoque = Reg_Dados3!vQtd
End If
Reg_Dados3.Close
Set Reg_Dados3 = Nothing

VsaldoEstoque = Vqtdent - Vqtdsaid - VqtdsaidEstoque
End Function



Public Sub sbBaixacheq()
Set Reg_Dados = BANCO.OpenRecordset("select * from tb_lancheq where d_data_pg<=#" & Format(Date, "mm/dd/yyyy") & "# order by d_data_pg")
Do While Not Reg_Dados.EOF
    Reg_Dados.Edit
    Reg_Dados!t_cond = "Pagos"
    Reg_Dados.Update
    Reg_Dados.MoveNext
Loop
Reg_Dados.Close
Set Reg_Dados = Nothing

End Sub


Public Function fcCalcLimite(nCodigo As Integer) As Variant
Set reg_Limite = BANCO.OpenRecordset("select sum (n_valor) as nLimite  from tb_lancheq where n_codDist=" & nCodigo & " and t_cond='" & "Nao Pagos" & "' and  Month(d_data_pg)=" & Month(Date) & "")
If IsNull(reg_Limite!nLimite) = True Then
    
    fcCalcLimite = 0
Else
    fcCalcLimite = reg_Limite!nLimite
End If
reg_Limite.Close
Set reg_Limite = Nothing

End Function

Public Function fcVerfQuant(nCodProd As Integer) As Integer
Set Reg_Estoq = BANCO.OpenRecordset("select * from cad_produto where codigo=" & nCodProd & "")
Select Case Left(Reg_Estoq!tipoemb, 2)
Case "CX"
    fcVerfQuant = Mid(Reg_Estoq!tipoemb, 4)
Case "FD"
    fcVerfQuant = Mid(Reg_Estoq!tipoemb, 4)
Case "UN"
    fcVerfQuant = 1
End Select
End Function


Public Function fcVerMult(nQtdmin, nQtdProd As Integer) As Integer
Dim nValorQtd As Currency

nValorQtd = nQtdProd / nQtdmin

If nValorQtd - CInt(nValorQtd) <> 0 Then
    fcVerMult = 0
Else
    fcVerMult = 1
End If

End Function

Public Function fcVerPrePedido(nfcNum As Variant, srtfcFranquia As String) As Boolean
Dim vSubTPed As Currency
Dim vSubTDescPed As Currency


Set Reg_Dados3 = BANCO.OpenRecordset("Select * from FranquiaMov_Pre_Venda where num=" & nfcNum & " and franquiaID='" & srtfcFranquia & "' and sit='" & "Novo" & "'")
vSubTPed = Format(Reg_Dados3!Valor, "##,##0.00")
Reg_Dados3.Close
Set Reg_Dados3 = Nothing

Set Reg_Dados3 = BANCO.OpenRecordset("Select SUM(valor) as nfcValor from franquiadecs_pre_vendas where num=" & nfcNum & " and franquiaid='" & srtfcFranquia & "'")
vSubTDescPed = Format(Reg_Dados3!nfcValor, "##,##0.00")
Reg_Dados3.Close
Set Reg_Dados3 = Nothing

If vSubTPed <> vSubTDescPed Then
    fcVerPrePedido = False
Else
    fcVerPrePedido = True
End If
End Function

Public Sub sbDesableOpt(ByVal frmForm As Form)
Dim I As Integer
       
       For I = 0 To frmForm.Controls.Count - 1
           
           If TypeOf frmForm.Controls(I) Is OptionButton Then
              frmForm.Controls(I).Enabled = False
           End If
                                
       Next I
End Sub

Public Function fStatusOs(nOs As Variant) As String
Dim nStatus As String

Call sbDesableOpt(frmOs)

Set rsTabelas = New ADODB.Recordset

rsTabelas.Open "select * from osdb where id=" & nOs & "", conn
nStatus = rsTabelas!Status

Select Case nStatus
Case 1
    frmOs.OptSolicitado.Value = True
    frmOs.OptSolicitado.Enabled = True
Case 2
    frmOs.OptAutorizado.Value = True
    frmOs.OptAutorizado.Enabled = True

Case 3
    frmOs.OptConcluido.Value = True
    frmOs.OptConcluido.Enabled = True
Case 4
    frmOs.OptFaturado.Value = True
    frmOs.OptFaturado.Enabled = True
Case 5
    frmOs.OptCancelado.Value = True
    frmOs.OptCancelado.Enabled = True
    
End Select

rsTabelas.Close
Set rsTabelas = Nothing

End Function

Public Function converte(Valor As Double)
Dim NovoValor As String
NovoValor = Format(Valor)

If InStr(NovoValor, ",") <> 0 Then
  Mid(NovoValor, InStr(NovoValor, ","), 1) = "."
  converte = NovoValor
End If

converte = NovoValor

End Function

'***********************************************************************************************************************
Public Function fcOutlook(strTitulo, strNome, strEmail, strMenssagem As String)
' Start Outlook.
 ' If it is already running, you'll use the same instance...
   Dim olApp As Outlook.Application
   Set olApp = CreateObject("Outlook.Application")
    
 ' Logon. Doesn't hurt if you are already running and logged on...
   Dim olNs As Outlook.NameSpace
   Set olNs = olApp.GetNamespace("MAPI")
   olNs.Logon

 ' Create and Open a new contact.
   Dim olItem As Outlook.ContactItem
   Set olItem = olApp.CreateItem(olContactItem)

 ' Setup Contact information...
   With olItem
      .FullName = strNome
      .Birthday = Date
      .CompanyName = "Suportek consultoria e serviços"
      .HomeTelephoneNumber = "84 3302-6370"
      .Email1Address = strEmail
      .JobTitle = strTitulo
      .HomeAddress = "www.suportek.net"
   End With
   
 ' Save Contact...
   olItem.Save
    
 ' Create a new appointment.
   Dim olAppt As Outlook.AppointmentItem
   Set olAppt = olApp.CreateItem(olAppointmentItem)
    
 ' Set start time for 2-minutes from now...
   olAppt.Start = Now() + (2# / 24# / 60#)
    
 ' Setup other appointment information...
'   With olAppt
'      .Duration = 60
'      .Subject = "Meeting to discuss plans..."
'      .Body = "Meeting with " & olItem.FullName & " to discuss plans."
'      .Location = "Home Office"
'      .ReminderMinutesBeforeStart = 1
'      .ReminderSet = True
'   End With
    
 ' Save Appointment...
   'olAppt.Save
    
 ' Send a message to your new contact.
   Dim olMail As Outlook.MailItem
   Set olMail = olApp.CreateItem(olMailItem)
 ' Fill out & send message...
   olMail.To = olItem.Email1Address
   olMail.Subject = strTitulo
   olMail.Body = _
        "ATT:. " & strNome & ", " & vbCr & vbCr & vbTab & strMenssagem & vbCr & vbCr & _
        "Obrigado" & vbCr & "Fone: 84 3302-6370"
   olMail.Send
    
 ' Clean up...
   MsgBox "Menssagem enviada...", vbMsgBoxSetForeground
   olNs.Logoff
   Set olNs = Nothing
   Set olMail = Nothing
   Set olAppt = Nothing
   Set olItem = Nothing
   Set olApp = Nothing
End Function
