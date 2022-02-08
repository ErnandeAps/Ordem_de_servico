Attribute VB_Name = "mdlMedmain"
Public conn As ADODB.Connection
Public connRemoto As ADODB.Connection
Public rs As New ADODB.Recordset
Public rsTabelas As New ADODB.Recordset
Public rsTabelas2 As New ADODB.Recordset
Public BANCO As Database
Public BANCOLIB As Database
Public BANCOLoja As Database
Public BANCOFtp As Database
Public Reg_DadosFtp As Recordset
Public REG_STATUS As Recordset
Public Reg_Dados As Recordset
Public Reg_Dados2 As Recordset
Public Reg_Dados3 As Recordset
Public reg_Limite As Recordset
Public Reg_Estoq As Recordset
Public Reg_Erro As Recordset
Public Reg_Seg As Recordset
Public Reg_Function As Recordset
Public MOP As String
Public Mcod_usu As Integer
Public MNIVEL As String
Public mturno As String
Public mData As Variant
Public MDIASEMANA As String
Public MusuSenha As String
Public Mtrat_turn As String
Public MUSUARIO As String
Public VSENHA As String
Public VNIVEL As String
Public M_inc As Integer
Public M_con As Integer
Public M_at As Integer
Public M_exc As Integer
Public M_seg_frm As Integer
Public Mpassword As String
Public MpendAssoc As String
Public MultiPag As Integer
Public Vmes As Integer
Public MsitDeb As Integer
Public Msitaut As Integer
Public glb_vnumPedido As Integer
Public strFranquiado As String
Public mNumF As Variant
Public strEmpresa As String
Public numUniqueID As Variant
Public strPasswordLoja As String
Public str As String
Public strDns As String
Public strDnsRemoto As String
Public strData As String
Public Sql As String
Public Ret As Integer
Public nIDsms As Integer
Public strArqPdf As String


Public Sub Main()
    mData = Format(Date, "mm/dd/yyyy")
    'Set BANCO = OpenDatabase(App.Path & "\..\Bank\basedados.mdb", False, False, ";PWD=EAPSSINDPRO")
'    Set BANCOLoja = OpenDatabase(App.Path & "\..\lojaketura\Bank\basedist.mdb", False, False, ";PWD=BRSOFTHOUSE1501")
    strDns = "DSN=Suportek-srv;Uid=root;pwd=(#suporte#)"
    
    'strDnsRemoto = "DSN=Suportekreomoto;Uid=suportek_supbr;pwd=(#suporte#)"
    
    mData = Format(Date, "MM/DD/YYYY")
    'Mpassword = "(#suporte#)"
    'strPasswordLoja = "BRSOFTHOUSE1501"
    
    'Call Conecta_BDRemoto
    
    MOP = "Admin"
    'Call AtivaForm(frmLogin)
    Call AtivaForm(frmFlexGrid)
End Sub
