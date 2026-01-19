VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaLibroIva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Iva Ventas"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   5445
   Begin VB.CheckBox ChkResumen 
      Caption         =   "Resumen"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3045
      Width           =   1095
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   4095
      TabIndex        =   7
      Top             =   2925
      Width           =   1245
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
      Height          =   2745
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   5235
      Begin VB.CheckBox ChkAgrupado 
         Caption         =   "Agrupado por Comprobantes"
         Height          =   195
         Left            =   315
         TabIndex        =   11
         Top             =   1995
         Width           =   2970
      End
      Begin VB.TextBox TxtPagina 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4035
         TabIndex        =   9
         Text            =   "0"
         Top             =   2340
         Width           =   930
      End
      Begin VB.TextBox TxtPeriodo 
         Height          =   285
         Left            =   1095
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "TxtPeriodo"
         Top             =   1470
         Width           =   3885
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   1575
         TabIndex        =   1
         Top             =   405
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   38638
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   915
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   38638
      End
      Begin VB.Label Label4 
         Caption         =   "Ultimo Número de Página :"
         Height          =   180
         Left            =   315
         TabIndex        =   10
         Top             =   2385
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha :"
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   390
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta  Fecha :"
         Height          =   180
         Left            =   330
         TabIndex        =   4
         Top             =   915
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Período :"
         Height          =   210
         Left            =   330
         TabIndex        =   3
         Top             =   1470
         Width           =   780
      End
   End
End
Attribute VB_Name = "VentaLibroIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As ADODB.Recordset

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
     End If
End Sub

Private Sub CmdAplicar_Click()
  Dim Rx As ADODB.Recordset, Rd As ADODB.Recordset, cRsl As ClsLectura, cRx As ClsComprobantesL
  Dim i As Byte
  
  Set cRsl = New ClsLectura
  Set cRx = New ClsComprobantesL
  If ChkResumen.Value = 0 Then
     CrearRs
     Set Rx = cRx.TraerIvaVenta(DtpDesde.Value, DtpHasta.Value, ChkAgrupado.Value)
     Set db1 = New Connection
     db1.Open sDb

     Do While Not Rx.EOF
        Set Rd = cRx.TraerDetIvaVenta(Rx!Movimiento)
        i = 1
        Do While Not Rd.EOF
           Rs.AddNew
           If i = 1 Then
              Rs!fecha = Rx!fecha
              Rs!Comprobante = Rx!Descripcion
'              If Rx!Numero = 12434 Then
'                 Dim a As Integer
'                 a = 1
'              End If
              Rs!Numero = Format(Rx!sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
              Rs!Cliente = Left(Rx!RazonSocial, 20)
              Rs!Cuit = Rx!Cuit
              Rs!RetenIva = 0
              Rs!NetoNoGravado = 0 ' Rx!NoGravados
              ' Para Abel diaz siguiente campos =0
              Rs!RetenPercep = IIf(cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & Rx!Compr) = 5, -Rx!PIvaCompras, Rx!PIvaCompras) ' 0
              Rs!Exento = 0
              Rs!Tasa = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & Rd!Tasa)
              Rs!Debito = Rd!Impuesto
              If Rd!Tasa <> 21 Then
                 If Rx!Debe <> 0 Then
                    If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Compr) = 0 Then
                       If Rd!Tasa = 10 Then
                          Rs!NetoNoGravado = Rd!PrecioTotal - Rd!Impuesto
                          Rs!NetoNoGravado = Rd!PrecioTotal - Rd!Descuento '- Rd!Impuesto
                          Rs!NetoGravado = 0
                       Else
                          Rs!NetoNoGravado = Rx!NoGravados
                          Rs!NetoGravado = Rd!PrecioTotal - Rd!Impuesto - Rd!Descuento - Rx!NoGravados
                       End If
                     Else
                        If Rd!Tasa = 10 Then
                           Rs!NetoGravado = 0
                           Rs!NetoNoGravado = Rd!PrecioTotal - Rd!Descuento
                        Else
                            Rs!NetoNoGravado = Rx!NoGravados
                            Rs!NetoGravado = Rd!PrecioTotal - Rd!Descuento '- Rd!Impuesto
                        End If
                      End If
                      Rs!Debito = Rd!Impuesto
                      If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Compr) = 0 Then
                          Rs!Total = Rd!PrecioTotal
                          Rs!Total = Rd!PrecioTotal - Rd!Descuento
                      Else
                         Rs!Total = Rs!NetoGravado + Rd!Impuesto - Rd!Descuento
                      End If
                 Else
                    If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Compr) = 0 Then
                       If Rd!Tasa <> 10 Then
                          Rs!NetoGravado = -Rd!PrecioTotal + Rd!Impuesto + Rd!Descuento
                       Else
                           Rs!NetoGravado = 0
                           Rs!NetoNoGravado = -Rd!PrecioTotal + Rd!Descuento
                       End If
                    Else
                       Rs!NetoGravado = -Rd!PrecioTotal + Rd!Descuento
                       Rs!NetoNoGravado = -Rx!NoGravados
                    End If
                    Rs!Debito = -Rd!Impuesto
                    If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Compr) = 0 Then
                       Rs!Total = -(Rd!PrecioTotal + Rx!NoGravados)
                    Else
                       Rs!Total = -(Rd!PrecioTotal + Rx!NoGravados + Rd!Impuesto)
                     End If
                 End If
              Else
                 Rs!NetoNoGravado = Rx!NoGravados
                 Rs!NetoGravado = 0
                 Rs!Total = Rx!NoGravados
             End If
              If Rx!NoGravados = 0 Then
                i = 2
              Else
                i = 0
              End If
           Else
              If ChkAgrupado.Value = 1 Then
                 Rs!fecha = Rx!fecha
                 Rs!Comprobante = Rx!Descripcion
                 Rs!Numero = Format(Rx!sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
                 Rs!Cliente = Left(Rx!RazonSocial, 20)
                 Rs!Cuit = Rx!Cuit
              End If
             Rs!NetoNoGravado = 0 ' Rx!NoGravados
             Rs!RetenPercep = 0
             Rs!Exento = 0
             Rs!Tasa = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & Rd!Tasa)
             Rs!Debito = Rd!Impuesto
             If Rd!Tasa <> 21 Then
                If Rx!Debe <> 0 Then
                   If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Compr) = 0 Then
                      Rs!NetoGravado = Rd!PrecioTotal - Rd!Descuento - Rd!Impuesto
                   Else
                      Rs!NetoGravado = Rd!PrecioTotal - Rd!Descuento
                   End If
                   Rs!Debito = Rd!Impuesto
                   If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Compr) = 0 Then
                      Rs!Total = Rd!PrecioTotal + Rx!NoGravados
                   Else
                      Rs!Total = Rd!PrecioTotal + Rx!NoGravados + Rd!Impuesto
                   End If
                Else
                   If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Compr) = 0 Then
                      Rs!NetoGravado = -Rd!PrecioTotal + Rd!Impuesto
                   Else
                      Rs!NetoGravado = -Rd!PrecioTotal
                   End If
                   Rs!Debito = -Rd!Impuesto
                   If cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Compr) = 0 Then
                      Rs!Total = -(Rd!PrecioTotal)
                   Else
                      Rs!Total = -(Rd!PrecioTotal + Rd!Impuesto)
                   End If
                End If
            Else
               Rs!NetoNoGravado = Rx!NoGravados
               Rs!NetoGravado = 0
               Rs!Total = Rx!NoGravados
            End If
          End If
       '   If Rd!Impuesto <> 0 Or Rs!Total = 0 Then
             Rs.Update
'          Else
'             Rs.CancelUpdate
'          End If
          If i = 0 Then
             Exit Do
          End If
          Rd.MoveNext
       Loop
       Rx.MoveNext
     Loop
     Set Rx = cRx.TraerRetenciones(DtpDesde.Value, DtpHasta.Value)
     Dim sCuit1 As String, sNom1 As String
  
     Do While Not Rx.EOF
        Rs.AddNew
        Rs!fecha = Rx!fecha
        Rs!Comprobante = Rx!Descripcion
        Rs!Numero = Rx!Numero
        sNom1 = Rx!RazonSocial
        sCuit1 = cRsl.TraerValorDeUnCampo("Valores", "NumeroTarjeta", "Venta=" & Rx!Movimiento)
        If sCuit1 <> "" Then
           sNom1 = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cuit='" & sCuit1 & "'")
        Else
           sCuit1 = Rx!Cuit
        End If
        If sNom1 = "0" Or sNom1 = "" Then
           Rs!Cliente = Left(Rx!RazonSocial, 20)
           Rs!Cuit = Rx!Cuit
        Else
           Rs!Cliente = Left(sNom1, 20)
           Rs!Cuit = sCuit1
        End If
        Rs!NetoGravado = 0
        Rs!NetoNoGravado = 0
        Select Case Rx!tipo
               Case 9
                    Rs!RetenIva = 0
                    Rs!RetenPercep = 0
                    Rs!Exento = -Rx!importe
                    Rs!Debito = 0
               Case 10
                    Rs!RetenIva = -Rx!importe
                    Rs!Debito = 0
                    Rs!RetenPercep = 0
                    Rs!Exento = 0
               Case 11
                    Rs!RetenIva = 0
                    Rs!RetenPercep = -Rx!importe
                    Rs!Exento = 0
                    Rs!Debito = 0
               Case 11
                    Rs!RetenIva = 0
                    Rs!RetenPercep = -Rx!importe
                    Rs!Exento = 0
                    Rs!Debito = 0
               Case 12
                    Rs!RetenIva = 0
                    Rs!RetenPercep = -Rx!importe
                    Rs!Exento = 0
                    Rs!Debito = 0
               Case 13
                    Rs!NetoNoGravado = -Rx!importe
                    Rs!RetenIva = 0
                    Rs!RetenPercep = 0
                 Rs!Exento = 0
                 Rs!Debito = 0
     End Select
     Rs!Tasa = 0
     Rs!Total = -Rx!importe
     Rs.Update
     Rx.MoveNext
  Loop
  db1.Close
'  If ChkAgrupado.Value = 0 Then
'     rs.Sort = "Fecha,Numero"
'  End If
  nImpr = 38
  FrmImpresor.Show
  Else
     CrearRsResumen
'     Set Rx = cRx.TraerNetoGrabado(DTPDesde.Value, DTPHasta.Value, 1)
'     If Rx.RecordCount <> 0 Then
'        rs.AddNew
'        rs!Descripcion = "Neto Grabado"
'        rs!Total = Rx!NetoTotal
'        rs.Update
'     End If
     Set Rx = cRx.TraerResumemIva(DtpDesde.Value, DtpHasta.Value, 1)
     Dim nTasa As String, nNeto As Currency, nTotal As Currency, nPorc As Single
     If Rx.RecordCount <> 0 Then
        nTasa = Rx!Descripcion
     End If
     Do While Not Rx.EOF
        nTotal = 0
        nNeto = 0
        Rs.AddNew
        Rs!Descripcion = Rx!Descripcion
        Do While nTasa = Rx!Descripcion
            nTotal = nTotal + Rx!Total
            nNeto = nNeto + Rx!NetoTotal
            Rx.MoveNext
            If Rx.EOF Then
               Exit Do
            End If
        Loop
        Rs!Total = nTotal
        Rs!Neto = nNeto
        Rs.Update
        If Not Rx.EOF Then
           nTasa = Rx!Descripcion
        End If
     Loop
     nImpr = 56
     FrmImpresor.Show
  End If
End Sub

Private Sub CrearRsResumen()
   Set Rs = New ADODB.Recordset

  Rs.Fields.Append "Descripcion", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Neto", adCurrency, , adFldIsNullable
  Rs.Fields.Append "Total", adCurrency, , adFldIsNullable
  Rs.Open
End Sub

Private Sub Form_Load()
  Limpiar
End Sub

Private Sub Limpiar()
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  TxtPeriodo.text = ""
End Sub

Private Sub CrearRs()
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs.Fields.Append "Comprobante", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Numero", adVarChar, 18, adFldIsNullable
  Rs.Fields.Append "Cliente", adVarChar, 20, adFldIsNullable
  Rs.Fields.Append "Cuit", adVarChar, 20, adFldIsNullable
  Rs.Fields.Append "NetoGravado", adCurrency, , adFldIsNullable
  Rs.Fields.Append "NetoNoGravado", adCurrency, , adFldIsNullable
  Rs.Fields.Append "Exento", adCurrency, , adFldIsNullable
  Rs.Fields.Append "RetenIva", adCurrency, , adFldIsNullable
  Rs.Fields.Append "RetenPercep", adCurrency, , adFldIsNullable
  Rs.Fields.Append "Tasa", adCurrency, , adFldIsNullable
  Rs.Fields.Append "Debito", adCurrency, , adFldIsNullable
  Rs.Fields.Append "Total", adCurrency, , adFldIsNullable
  Rs.Open
  
End Sub

Private Sub TxtPagina_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPagina_LostFocus()
  If TxtPagina.text = "" Then
     TxtPagina.text = 0
  End If
  TxtPagina.text = Format(TxtPagina.text, "0")
End Sub

Private Sub TxtPagina__GotFocus()
  TxtPagina.SelStart = 0
  TxtPagina.SelLength = Len(TxtPagina.text)
End Sub

