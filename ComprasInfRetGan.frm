VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasInfRetGan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retención de Ganancia"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   3780
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Aplicar"
      Height          =   420
      Left            =   2700
      TabIndex        =   5
      Top             =   2355
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   3510
      Begin VB.CheckBox ChkExportar 
         Caption         =   "Exportar al Sicore"
         Height          =   225
         Left            =   135
         TabIndex        =   6
         Top             =   1890
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   345
         Left            =   1140
         TabIndex        =   1
         Top             =   450
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   52690945
         CurrentDate     =   36983
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   1185
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   52690945
         CurrentDate     =   36983
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   420
         TabIndex        =   4
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   435
         TabIndex        =   3
         Top             =   450
         Width           =   645
      End
   End
End
Attribute VB_Name = "ComprasInfRetGan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdImprimir_Click()
  If ChkExportar.Value = 1 Then
     ExportarSicore
  End If
End Sub

Private Sub Form_Load()
  DTPDesde.Value = Date
  DTPHasta.Value = Date
End Sub

Private Sub ExportarSicore()
  Dim cRx As ClsComprobantesL, cRsl As ClsLectura
  
  Set cRx = New ClsComprobantesL
  Set cRsl = New ClsLectura
  
  CmdImprimir.Enabled = False
  
  Dim sCodComprobante As String, sFechaEmi As String, sNumeroCom As String, sImpoCompr As String, sCodImpuesto As String, sRegimen As String
  Dim sCodOperacion As String, sBase As String, sFechaRet As String, sCondicion As String, sRetePrac As String, sImporRet As String, sPorExlucion As String
  Dim sFechaEmision As String, sTipoDoc As String, sDocumento As String, sNumCerti As String
  
  Dim Rx As Recordset, sVariable As String
    
  sCodComprobante = "06"
  sCodImpuesto = "217"
  sCodOperacion = "1"
  sCondicion = "01"
  sRetePrac = "0"
  sPorExlucion = "  0.00"
  sTipoDoc = "80"
  sVariable = ""
  
  Set Rx = cRx.TraerGananciaResumen(DTPDesde.Value, DTPHasta.Value)
  
  If Rx.RecordCount <> 0 Then
     Dim sNom As String
     Dim sNomDet As String
        
     sNom = CurDir() & "\RetGanancia\RetGan" & Replace(sCuitAfip, "-", "") & Month(DTPDesde.Value) & Year(DTPDesde.Value) & ".txt"
   '  sNom = "C:\RetGanancia\RetGan" & Replace(sCuitAfip, "-", "") & Month(DTPDesde.Value) & Year(DTPDesde.Value) & ".txt"
     
     Open sNom For Output As #1
        
      Do While Not Rx.EOF
         sVariable = sCodComprobante
         sVariable = sVariable & Rx!fecha
         sVariable = sVariable & "0001" & Format(Rx!nCompro, "000000000000")
         
         Dim nNeto As String, nCat As Byte, sReten As String
         
         nNeto = Format(Rx!Neto, "0.00")
         sImpoCompr = Space(16 - Len(nNeto)) & nNeto
         sVariable = sVariable & sImpoCompr & sCodImpuesto
                  
         nCat = cRsl.TraerValorDeUnCampo("Proveedores", "CatGanancia", "Cuit='" & Rx!Cuit & "'")
         sRegimen = cRsl.TraerValorDeUnCampo("TblGananciaItems", "Codigo", "Id_Ganancia=" & nCat)
         sVariable = sVariable & sRegimen & sCodOperacion
         sBase = Space(14 - Len(nNeto)) & nNeto
         sVariable = sVariable & sBase
         sVariable = sVariable & Rx!fecha
         sVariable = sVariable & sCondicion
         sVariable = sVariable & sRetePrac
         
         sReten = Format(Rx!importe, "0.00")
         sImporRet = Space(14 - Len(sReten)) & sReten
         sVariable = sVariable & sImporRet & sPorExlucion & Rx!fecha & sTipoDoc
         
         sDocumento = Space(20 - Len(Rx!Cuit)) & Rx!Cuit
         sVariable = sVariable & sDocumento
         
         sNumCerti = "0001" & Format(Rx!Numero, "0000000000")
         sVariable = sVariable & sNumCerti
         
         Print #1, sVariable
         Rx.MoveNext
      Loop
      Close #1
        
      MsgBox "Archivos de Exportación Para el Sicore Creado Con Exito", vbExclamation, "Atención"

  End If
  
  CmdImprimir.Enabled = True
End Sub
