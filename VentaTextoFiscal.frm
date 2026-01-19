VERSION 5.00
Begin VB.Form VentaTextoFiscal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Texto del Comprobante"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   525
      Left            =   1065
      TabIndex        =   1
      Top             =   4755
      Width           =   2085
   End
   Begin VB.TextBox TxtTexto 
      Height          =   4530
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "VentaTextoFiscal.frx":0000
      Top             =   120
      Width           =   3900
   End
End
Attribute VB_Name = "VentaTextoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset, pMovi As Long

Private Sub CmdAceptar_Click()
    
  sTexFiscal = ImprimirXLinea(TxtTexto)
  Grabar
  Unload Me
End Sub

Private Sub Form_Load()
   Me.Width = Me.Width + sAnchoTexto
   TxtTexto.Width = TxtTexto.Width + sAnchoTexto
   TxtTexto.text = ""
   pMovi = VentaComprobantes.nMovi
   LlenarRs
End Sub

Private Sub LlenarRs()
  Dim cRsl As ClsLectura, sCli As String, nTra As Integer
  
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRsCondi("CabTexto", "Id", "Movimiento=" & pMovi)
    
  TxtTexto.text = ""
  
  If Rs.RecordCount = 0 Then
     Rs.AddNew
     Rs!Movimiento = pMovi
     Rs!Texto1 = ""
     Rs!Texto2 = ""
     Rs!Texto3 = ""
     Rs!Texto4 = ""
     Rs!TextoFiscal = TxtTexto.text
     Rs.Update
  End If

  TxtTexto.text = "" & Rs!TextoFiscal
 
  
  Set cRsl = Nothing

End Sub

Private Sub Grabar()
  Dim cRle As ClsComprobantesE
  
  Set cRle = New ClsComprobantesE
  Rs!Movimiento = pMovi
  Rs!Texto1 = ""
  Rs!Texto2 = ""
  Rs!Texto3 = ""
  Rs!Texto4 = ""
  Rs!TextoFiscal = TxtTexto.text
  Rs.Update
  
  cRle.GrabarTexto Rs, pMovi
  LlenarRs
  Set cRle = Nothing
End Sub

