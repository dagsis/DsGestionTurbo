VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CitiCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Citi Compras"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   5595
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
      Height          =   1905
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   5415
      Begin VB.ComboBox CmbPeriodo 
         Height          =   315
         ItemData        =   "CitiCompra.frx":0000
         Left            =   2010
         List            =   "CitiCompra.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   1980
         TabIndex        =   3
         Top             =   870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60424193
         CurrentDate     =   42270
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   1980
         TabIndex        =   4
         Top             =   1350
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60424193
         CurrentDate     =   42270
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   270
         Left            =   375
         TabIndex        =   7
         Top             =   825
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   270
         Left            =   375
         TabIndex        =   6
         Top             =   1305
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Peróodo :"
         Height          =   225
         Left            =   375
         TabIndex        =   5
         Top             =   300
         Width           =   1770
      End
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   435
      Left            =   3615
      TabIndex        =   0
      Top             =   2010
      Width           =   1830
   End
End
Attribute VB_Name = "CitiCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdGenerar_Click()

  CmdGenerar.Enabled = False
  
  Dim cRsl As ClsLectura, Rx As Recordset, cRx As ClsComprobantesL, Rd As Recordset
  
  Set cRsl = New ClsLectura
  Set cRx = New ClsComprobantesL
  
  ' Cabecera
  Dim sFechaCompro, sTipo, sPunto, sNumero, sNumeroDespacho, sCodDoc, sNumeroDoc, sApellido, sImporteTotal As String
  Dim sNeto, sPerNoCat, sExentas, sPerNacio, sPercIB, sPerMuni, sInternos, sCodMoneda, sTipoCambio, sCantAlicuota As String
  Dim sCodOperacion, sCreditoFiscal, sOtrosTributos, sCuitEmisor, sDenoEmisor, sIvaEmisor As String
  Dim sAlicuota, sIva As String, nNeto As Double
  
  
  If CmbPeriodo.ListIndex <> -1 Then
     Set Rx = cRx.TraerCabezaCtiCompra(DtpDesde.Value, DtpHasta.Value)
     If Rx.RecordCount <> 0 Then
        Dim sNom As String, sVariable As String
        Dim sNomDet As String, sVariableDet As String
        
        sNom = CurDir() & "\CitiCompras\CompraComprobante" & Replace(sCuitAfip, "-", "") & CmbPeriodo.text & Year(DtpDesde.Value) & ".txt"
        sNomDet = CurDir() & "\CitiCompras\CompraAlicuotas" & Replace(sCuitAfip, "-", "") & CmbPeriodo.text & Year(DtpDesde.Value) & ".txt"
          
        Open sNom For Output As #1
        Open sNomDet For Output As #2
        Dim J As Integer
        Do While Not Rx.EOF
seguir:
           If Rx!Numero <> 0 Then
              sVariable = ""
           
              sFechaCompro = Year(Rx!FechaIva) & Format(Month(Rx!FechaIva), "00") & Format(Day(Rx!FechaIva), "00")
              sVariable = sVariable & sFechaCompro
              sTipo = Format(cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & Rx!Comprobante), "000")
              sVariable = sVariable & sTipo
              If sTipo = "033" Then
                  sPunto = "00000"
                  sVariable = sVariable & sPunto
                  sNumero = Format(Trim(Str(Rx!Sucursal)) & Trim(Str(Rx!Numero)), "00000000000000000000")
                  sVariable = sVariable & sNumero
                  sNumeroDespacho = Space(16)
                  sVariable = sVariable & sNumeroDespacho
              Else
                  sPunto = Format(IIf(Rx!Sucursal = 0, 1, Rx!Sucursal), "00000")
                  sVariable = sVariable & sPunto
                  sNumero = Format(Rx!Numero, "00000000000000000000")
                  sVariable = sVariable & sNumero
                  sNumeroDespacho = Space(16)
                  sVariable = sVariable & sNumeroDespacho
              End If
              
              sCodDoc = "80"
              If cRsl.TraerValorDeUnCampo("Proveedores", "Cuit", "Proveedor='" & Rx!Cliente & "'") <> "" Then
                 sNumeroDoc = cRsl.TraerValorDeUnCampo("Proveedores", "Cuit", "Proveedor='" & Rx!Cliente & "'")
              Else
                  sCodDoc = "99"
                  sNumeroDoc = "00000000000000000000"
              End If
              sVariable = sVariable & sCodDoc
              sVariable = sVariable & Format(sNumeroDoc, "00000000000000000000")
              sApellido = cRsl.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Proveedor='" & Rx!Cliente & "'")
              sApellido = Left(sApellido, 30)
              sVariable = sVariable & sApellido & Space(30 - Len(sApellido))
              Dim nTotal As Double
              nTotal = (Rx!Neto) + Rx!Iva1 + Abs(Rx!NoGravados) + Abs(Rx!PIvaCompras) + Abs(Rx!IBrutosCompras)
              
              sImporteTotal = Format(Str(nTotal) * 100, "000000000000000")
              
            '  sImporteTotal = Format(Str(Abs(Rx!Haber - Rx!Debe + IIf(Rx!NoGravados < 0, Abs(Rx!NoGravados), 0)) * 100), "000000000000000")
              sVariable = sVariable & sImporteTotal
              sNeto = "000000000000000"
              sVariable = sVariable & sNeto
              sPerNoCat = "000000000000000"
              sExentas = "000000000000000"
              sPerNacio = Format(Str(Abs(Rx!PIvaCompras) * 100), "000000000000000")
              sPercIB = Format(Str(Abs(Rx!IBrutosCompras) * 100), "000000000000000")
              sPerMuni = "000000000000000"
              sInternos = Format(Str(Abs(Rx!NoGravados) * 100), "000000000000000")
              sVariable = sVariable & sPerNoCat & sExentas & sPerNacio & sPercIB & sPerMuni & sInternos
              sCodMoneda = "PES"
              sTipoCambio = "0001000000"
              sVariable = sVariable & sCodMoneda & sTipoCambio
              Set Rd = cRx.TraerDetalleCitiVenta(Rx!Movimiento)
              sCantAlicuota = IIf(sTipo = "033", 1, Rd.RecordCount)
              sCantAlicuota = IIf(sTipo = "011" Or sTipo = "012" Or sTipo = "013", 0, Rd.RecordCount)

              sCodOperacion = "0"
              If cRsl.TraerCantidad("DetallesComprobantes", "Tasa", "Movimiento=" & Rx!Movimiento & " and Tasa=10") > 0 Then
                  sCodOperacion = "E"
              End If
              If cRsl.TraerCantidad("DetallesComprobantes", "Tasa", "Movimiento=" & Rx!Movimiento & " and Tasa=21") > 0 Then
                 If cRsl.TraerCantidad("DetallesComprobantes", "Tasa", "Movimiento=" & Rx!Movimiento & " and Tasa>=9 and Tasa<=16") > 0 Then
                     sCodOperacion = "0"
                 Else
                     sCodOperacion = "N"
                 End If
              End If
              sVariable = sVariable & sCantAlicuota & sCodOperacion
              
              sCreditoFiscal = Format(Str(Abs(Rx!Iva1) * 100), "000000000000000")
              sVariable = sVariable & sCreditoFiscal
              sOtrosTributos = "000000000000000"
              sVariable = sVariable & sOtrosTributos
              sCuitEmisor = IIf(sTipo = "033", Format(sNumeroDoc, "00000000000"), "00000000000")
              sDenoEmisor = IIf(sTipo = "033", sApellido & Space(30 - Len(sApellido)), Space(30))
              sIvaEmisor = "000000000000000"
              sVariable = sVariable & sCuitEmisor & sDenoEmisor & sIvaEmisor
              Print #1, sVariable
                         
              Do While Not Rd.EOF
                 sVariableDet = ""
                 sVariableDet = sVariableDet & sTipo
                 sVariableDet = sVariableDet & sPunto
                 sVariableDet = sVariableDet & sNumero & sCodDoc & Format(sNumeroDoc, "00000000000000000000")
                 Dim cNeto As Double
                 
                 cNeto = Rx!Neto
                 If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Afip=" & Val(sTipo)) = 0 Then
                    cNeto = Rd!SumaDePrecioTotal '- Rd!SumaDeImpuesto
                 Else
                    cNeto = Rd!SumaDePrecioTotal
                 End If
                 
                 sNeto = Format(Str(Abs(cNeto) * 100), "000000000000000")
                  If sCodOperacion = "N" Then
                    sNeto = "000000000000000"
                 End If
                 sVariableDet = sVariableDet & sNeto
                 sAlicuota = Format(Rd!CodAfip, "0000")
                 sVariableDet = sVariableDet & sAlicuota
                 Dim nIva As Double
                 nIva = Rd!SumaDeImpuesto - IIf(sTipo = "033", Abs(Rx!NoGravados), 0)
                 If sCodOperacion = "N" Then
                    nIva = 0
                 End If
                 sIva = Format(Round(nIva, 2) * 100, "000000000000000")
                 sVariableDet = sVariableDet & sIva
                 If sAlicuota <> 0 Then
                    If nIva >= 0 Then
                       Print #2, sVariableDet
                    End If
                 End If
                 Rd.MoveNext
              Loop
           End If
           Rx.MoveNext
        Loop
        Close #1
        Close #2
        
        MsgBox "Archivos de Exportación Para el Citi Venta Creados Con Exito", vbExclamation, "Atención"
    Else
        MsgBox "No Hay Registros para Realizar esta Operación", vbExclamation, "Atención"
    End If
  Else
     MsgBox "Debe Seleccionar un Período", vbExclamation, "Atención"
  End If
  
  CmdGenerar.Enabled = True

End Sub

Private Sub Form_Load()
  Limpiar
End Sub

Private Sub Limpiar()
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  
  CmbPeriodo.AddItem "Enero"
  CmbPeriodo.ItemData(0) = 1
  CmbPeriodo.AddItem "Febrero"
  CmbPeriodo.ItemData(0) = 2
  CmbPeriodo.AddItem "Marzo"
  CmbPeriodo.ItemData(0) = 3
  CmbPeriodo.AddItem "Abril"
  CmbPeriodo.ItemData(0) = 4
  CmbPeriodo.AddItem "Mayo"
  CmbPeriodo.ItemData(0) = 5
  CmbPeriodo.AddItem "Junio"
  CmbPeriodo.ItemData(0) = 6
  CmbPeriodo.AddItem "Julio"
  CmbPeriodo.ItemData(0) = 7
  CmbPeriodo.AddItem "Agosto"
  CmbPeriodo.ItemData(0) = 8
  CmbPeriodo.AddItem "Septiembre"
  CmbPeriodo.ItemData(0) = 9
  CmbPeriodo.AddItem "Octubre"
  CmbPeriodo.ItemData(0) = 10
  CmbPeriodo.AddItem "Noviembre"
  CmbPeriodo.ItemData(0) = 11
  CmbPeriodo.AddItem "Diciembre"
  CmbPeriodo.ItemData(0) = 12
           
End Sub

