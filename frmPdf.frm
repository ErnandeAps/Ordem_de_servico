VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmPdf 
   Caption         =   "Visualizador de PDF"
   ClientHeight    =   9510
   ClientLeft      =   2730
   ClientTop       =   1065
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   11985
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   9285
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11835
      _cx             =   20876
      _cy             =   16378
   End
End
Attribute VB_Name = "frmPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   
'   'define o tamanho do formulário e do controle WebBrowser Maximizados
  ' frmPdf.WindowState = 2 'Maximizado
  'AcroPDF1.Width = 11520
  ' AcroPDF1.Height = 11520
'
'   'exibe o arquivo pdf selecionado no formulário anterior
'   'strArquivo = "C:\Projeto Bematech\pdfVB\VB Net.pdf"
'   'wbrs1.Navigate frmBusca.cmdlg1.FileName
'    wbrs1.Navigate strArqPdf

AcroPDF1.src = strArqPdf
End Sub
