VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CitiVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CITI Ventas"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   5655
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   435
      Left            =   3645
      TabIndex        =   1
      Top             =   2010
      Width           =   1830
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
      Height          =   1905
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   5415
      Begin VB.ComboBox CmbPeriodo 
         Height          =   315
         ItemData        =   "CitiVenta.frx":0000
         Left            =   2010
         List            =   "CitiVenta.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   315
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   1980
         TabIndex        =   2
         Top             =   870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51052545
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
         Format          =   51052545
         CurrentDate     =   42270
      End
      Begin VB.Label Label3 
         Caption         =   "Peróodo :"
         Height          =   225
         Left            =   375
         TabIndex        =   6
         Top             =   300
         Width           =   1770
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   270
         Left            =   375
         TabIndex        =   5
         Top             =   1305
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   270
         Left            =   375
         TabIndex        =   3
         Top             =   825
         Width           =   660
      End
   End
End
Attribute VB_Name = "CitiVenta"
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
  Dim sFechaCompro, sTipo, sPunto, sNumero, sNumeroHasta, sCodDoc, sNumeroDoc, sApellido, sImporteTotal As String
  Dim sNeto, sPerNoCat, sExentas, sPerNacio, sPercIB, sPerMuni, sInternos, sCodMoneda, sTipoCambio, sCantAlicuota As String
  Dim sCodOperacion, sOtrosTributos, sFechaVenc As String
  Dim sAlicuota, sIva As String, nNeto As Double
  
  
  If CmbPeriodo.ListIndex <> -1 Then
     Set Rx = cRx.TraerCabezaCtiVenta(DtpDesde.Value, DtpHasta.Value)
     If Rx.RecordCount <> 0 Then
        Dim sNom As String, sVariable As String
        Dim sNomDet As String, sVariableDet As String
        
        sNom = CurDir() & "\CitiVentas\VentaComprobante" & Replace(sCuitAfip, "-", "") & CmbPeriodo.text & Year(DtpDesde.Value) & ".txt"
        sNomDet = CurDir() & "\CitiVentas\VentaAlicuotas" & Replace(sCuitAfip, "-", "") & CmbPeriodo.text & Year(DtpDesde.Value) & ".txt"
          
        Open sNom For Output As #1
        Open sNomDet For Output As #2
        Dim j As Integer
        Do While Not Rx.EOF
seguir:
           If Rx!Numero <> 0 Then
              sVariable = ""
              ' Verificador de Facturas
           '   If cRsl.TraerCantidad("CabComprobantes", "Comprobante", "Comprobante=" & Rx!Comprobante & " AND Numero=" & Rx!Numero & " AND Sucursal=" & Rx!Sucursal) = 2 Then
           '      MsgBox "Factura Repetida o Sin Autorización de AFIP :" & cRsl.TraerValorDeUnCampo("Comprobantes", "Descripcion", "Id=" & Rx!Comprobante) & " " & Format(Rx!Sucursal, "0000") & "-" & Format(Rx!Numero, "00000000"), vbCritical, "Atención"
'                 Close #1
'                 Close #2
'                 Exit Sub
            '  End If
              sFechaCompro = Year(Rx!fecha) & Format(Month(Rx!fecha), "00") & Format(Day(Rx!fecha), "00")
              sVariable = sVariable & sFechaCompro
              sTipo = Format(cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & Rx!Comprobante), "000")
              sVariable = sVariable & sTipo
              If sTipo = "033" Then
                  sPunto = "00000"
                  sVariable = sVariable & sPunto
                  sNumero = Format(Trim(Str(Rx!Sucursal)) & Trim(Str(Rx!Numero)), "00000000000000000000")
                  sVariable = sVariable & sNumero
                  sNumeroHasta = sNumero
                  sVariable = sVariable & sNumeroHasta
              Else
                  sPunto = Format(Rx!Sucursal, "00000")
                  sVariable = sVariable & sPunto
                  sNumero = Format(Rx!Numero, "00000000000000000000")
                  sVariable = sVariable & sNumero
                  sNumeroHasta = Format(Rx!Numero, "00000000000000000000")
                  sVariable = sVariable & sNumeroHasta
              End If
              sCodDoc = "80"
              If Rx!Cliente = "." Or cRsl.TraerValorDeUnCampo("Clientes", "TipoIva", "Cliente='" & Rx!Cliente & "'") = 1 Then
                sCodDoc = "96"
              End If
              sNumeroDoc = cRsl.TraerValorDeUnCampo("AfipCaes", "Documento", "Movimiento=" & Rx!Movimiento)
              If Val(sNumeroDoc) = 0 Then
                 If cRsl.TraerValorDeUnCampo("Clientes", "Cuit", "Cliente='" & Rx!Cliente & "'") <> "" Then
                    sNumeroDoc = cRsl.TraerValorDeUnCampo("Clientes", "Cuit", "Cliente='" & Rx!Cliente & "'")
                    If sNumeroDoc = "00000000000" Then
                       sNumeroDoc = "99999999999"
                    End If
                 Else
                    sCodDoc = "96"
                    sNumeroDoc = "99999999999"
'                    Rx.MoveNext
'                    GoTo seguir
                 End If
              End If
              sVariable = sVariable & sCodDoc
              sVariable = sVariable & Format(sNumeroDoc, "00000000000000000000")
              sApellido = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & Rx!Cliente & "'")
              sApellido = Left(sApellido, 30)
              sVariable = sVariable & sApellido & Space(30 - Len(sApellido))

              Dim nTotal As Double
              nTotal = (Rx!Neto - Rx!Descuentos) + Rx!Iva1 + Abs(Rx!NoGravados) + Abs(Rx!PIvaCompras) + Abs(Rx!IBrutosCompras)
              
              sImporteTotal = Format(Str(nTotal) * 100, "000000000000000")
                            
              sVariable = sVariable & sImporteTotal
              sNeto = "000000000000000"
              sVariable = sVariable & sNeto
              sPerNoCat = "000000000000000"
              sExentas = "000000000000000"
              sPerNacio = Format(Str(Rx!PIvaCompras * 100), "000000000000000")
              sPercIB = Format(Str(Rx!IBrutosCompras * 100), "000000000000000")
              sPerMuni = "000000000000000"
              sInternos = Format(Str(Abs(Rx!NoGravados) * 100), "000000000000000")
              sVariable = sVariable & sPerNoCat & sExentas & sPerNacio & sPercIB & sPerMuni & sInternos
              sCodMoneda = "PES"
              sTipoCambio = "0001000000"
              sVariable = sVariable & sCodMoneda & sTipoCambio
              Set Rd = cRx.TraerDetalleCitiVenta(Rx!Movimiento)
              sCantAlicuota = IIf(sTipo = "033", 1, Rd.RecordCount)
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
              sOtrosTributos = "000000000000000"
              sVariable = sVariable & sCantAlicuota & sCodOperacion & sOtrosTributos
              If Rx!Cliente <> "." Then
                 sFechaVenc = Year(Rx!fecha) & Format(Month(Rx!fecha), "00") & Format(Day(Rx!fecha), "00")
              Else
                 sFechaVenc = "00000000"
              End If
              If sTipo = "051" Or sTipo = "033" Or Val(sTipo) >= 81 Then
                 sFechaVenc = "00000000"
              End If
              sVariable = sVariable & sFechaVenc
              If Rd.RecordCount <> 0 Then
                 Print #1, sVariable
              End If
              Do While Not Rd.EOF
                 sVariableDet = ""
                 sVariableDet = sVariableDet & sTipo
                 sVariableDet = sVariableDet & sPunto
                 sVariableDet = sVariableDet & sNumero
                 Dim cNeto As Double
                 
                 cNeto = (Rx!Neto - Rx!Descuentos)

                 If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Afip=" & Val(sTipo)) = 0 Then
                    cNeto = Rd!SumaDePrecioTotal - Rd!SumaDeImpuesto
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
