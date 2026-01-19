VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasLibroIva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Iva Compras"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   5475
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   4125
      TabIndex        =   7
      Top             =   2985
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
      Height          =   2805
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   5235
      Begin VB.CheckBox ChkAgrupado 
         Caption         =   "Agrupado por Comprobantes"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1920
         Width           =   2970
      End
      Begin VB.TextBox TxtPagina 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3990
         TabIndex        =   9
         Text            =   "0"
         Top             =   2190
         Width           =   930
      End
      Begin VB.CheckBox ChkResumen 
         Caption         =   "Resumen"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   2490
         Width           =   1095
      End
      Begin VB.TextBox TxtPeriodo 
         Height          =   285
         Left            =   1095
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "TxtPeriodo"
         Top             =   1470
         Width           =   3885
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Top             =   405
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   38638
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   915
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   38638
      End
      Begin VB.Label Label4 
         Caption         =   "Ultimo Número de Página :"
         Height          =   180
         Left            =   270
         TabIndex        =   10
         Top             =   2220
         Width           =   1950
      End
      Begin VB.Label Label3 
         Caption         =   "Período :"
         Height          =   210
         Left            =   330
         TabIndex        =   6
         Top             =   1470
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta  Fecha :"
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   915
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha :"
         Height          =   180
         Left            =   330
         TabIndex        =   4
         Top             =   390
         Width           =   1110
      End
   End
End
Attribute VB_Name = "ComprasLibroIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Private Sub CmdAplicar_Click()
  Dim Rx As ADODB.Recordset, Rd As ADODB.Recordset, cRsl As ClsLectura, cRx As ClsComprobantesL
  Dim i As Byte
  
  Set cRsl = New ClsLectura
  Set cRx = New ClsComprobantesL
  If ChkResumen.Value = 0 Then
  CrearRs
  Set Rx = cRx.TraerIvaCompra(DtpDesde.Value, DtpHasta.Value, ChkAgrupado.Value)
  Set db1 = New Connection
  db1.Open sDb
  
  Do While Not Rx.EOF
     Set Rd = cRx.TraerDetIvaCompra(Rx!Movimiento)
     i = 1
     Do While Not Rd.EOF
        rs.AddNew
        If i = 1 Then
           rs!fecha = Rx!FechaIva
           rs!Comprobante = Rx!Descripcion
           rs!Numero = Format(Rx!Sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
           rs!Cliente = Left(Rx!RazonSocial, 20)
           rs!Cuit = Rx!Cuit
           rs!RetenIva = 0
           If Rx!Haber <> 0 Then
              rs!NetoNoGravado = Rx!NoGravados
              rs!RetenPercep = Rx!IBrutosCompras
           Else
              rs!NetoNoGravado = -Rx!NoGravados
              rs!RetenPercep = -Rx!IBrutosCompras
           End If
           rs!Exento = IIf(cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & Rx!Compr) = 12, -Rx!PIvaCompras, Rx!PIvaCompras)     'Rx!PIvaCompras
           rs!Tasa = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & Rd!Tasa)
           rs!Debito = Rd!Impuesto
           If rs!Tasa <> 0 Then
              If Rx!Haber <> 0 And cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & Rx!Compr) <> 12 Then
                 rs!NetoGravado = IIf(Rd!Tasa = 21, 0, Rd!PrecioTotal)
                 rs!Debito = Rd!Impuesto
                 rs!Total = IIf(Rd!Tasa = 21, 0, Rd!PrecioTotal) + Rx!NoGravados + Rd!Impuesto + rs!RetenPercep + rs!Exento
              Else
                 rs!NetoGravado = -Rd!PrecioTotal
                 rs!RetenPercep = -Rx!IBrutosCompras
                 rs!Debito = -Rd!Impuesto
                 rs!Total = -(Rd!PrecioTotal + Rx!NoGravados + Rd!Impuesto + Abs(rs!RetenPercep) + rs!Exento)
              End If
           Else
               rs!NetoGravado = 0
               rs!NetoNoGravado = IIf(cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & Rx!Compr) = 12, -(Rd!PrecioTotal + Rx!NoGravados), Rd!PrecioTotal + Rx!NoGravados)
               rs!Debito = Rd!Impuesto
               rs!Total = Rd!PrecioTotal
           End If
           i = 2
        Else
           If ChkAgrupado.Value = 1 Then
              rs!fecha = FechaAmericanaIva(Rx!FechaIva)
              rs!Comprobante = Rx!Descripcion
              rs!Numero = Format(Rx!Sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
              rs!Cliente = Left(Rx!RazonSocial, 20)
              rs!Cuit = Rx!Cuit
           End If
           rs!RetenIva = 0
           rs!NetoNoGravado = Rx!NoGravados
           rs!RetenPercep = Rx!IBrutosCompras
           rs!Exento = Rx!PIvaCompras
           rs!Tasa = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & Rd!Tasa)
           rs!Debito = Rd!Impuesto
           rs!NetoGravado = Rd!PrecioTotal
           rs!Tasa = cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & Rd!Tasa)
           rs!Debito = Rd!Impuesto
           rs!NetoNoGravado = 0
           rs!RetenPercep = 0
           rs!Exento = 0
           If Rx!Haber <> 0 Then
              rs!NetoGravado = Rd!PrecioTotal
              rs!Debito = Rd!Impuesto
              rs!Total = Rd!PrecioTotal + Rd!Impuesto + rs!RetenPercep
           Else
              rs!NetoGravado = -Rd!PrecioTotal
              rs!Debito = -Rd!Impuesto
              rs!Total = -(Rd!PrecioTotal + Rd!Impuesto + rs!RetenPercep)
           End If
        End If
        rs.Update
        Rd.MoveNext
     Loop
     Rx.MoveNext
  Loop
  db1.Close
  nImpr = 39

  FrmImpresor.Show
  Else
     CrearRsResumen
'     Set Rx = cRx.TraerNetoGrabado(DtpDesde.Value, DtpHasta.Value, 8)
'     If Rx.RecordCount <> 0 Then
'        rs.AddNew
'        rs!Descripcion = "Neto Grabado"
'        rs!Total = Rx!NetoTotal
'        rs.Update
'     End If
     Set Rx = cRx.TraerResumenIvaC(DtpDesde.Value, DtpHasta.Value, 8)
     Dim nTasa As String, nNeto As Currency, nTotal As Currency, nPorc As Single
     If Rx.RecordCount <> 0 Then
        nTasa = Rx!Descripcion
     End If
     Do While Not Rx.EOF
        nTotal = 0
        nNeto = 0
        rs.AddNew
        rs!Descripcion = Rx!Descripcion
        Do While nTasa = Rx!Descripcion
            nTotal = nTotal + Rx!Total
            nNeto = nNeto + Rx!NetoTotal
            Rx.MoveNext
            If Rx.EOF Then
               Exit Do
            End If
        Loop
        rs!Total = nTotal
        rs!Neto = nNeto
        rs.Update
        If Not Rx.EOF Then
           nTasa = Rx!Descripcion
        End If
     Loop
     nImpr = 57
     FrmImpresor.Show
  End If
End Sub

Private Sub CrearRsResumen()
   Set rs = New ADODB.Recordset

  rs.Fields.Append "Descripcion", adVarChar, 30, adFldIsNullable
  rs.Fields.Append "Neto", adCurrency, , adFldIsNullable
  rs.Fields.Append "Total", adCurrency, , adFldIsNullable
  rs.Open
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
  Set rs = New ADODB.Recordset
  
  rs.Fields.Append "Fecha", adDate, , adFldIsNullable
  rs.Fields.Append "Comprobante", adVarChar, 30, adFldIsNullable
  rs.Fields.Append "Numero", adVarChar, 13, adFldIsNullable
  rs.Fields.Append "Cliente", adVarChar, 20, adFldIsNullable
  rs.Fields.Append "Cuit", adVarChar, 20, adFldIsNullable
  rs.Fields.Append "NetoGravado", adCurrency, , adFldIsNullable
  rs.Fields.Append "NetoNoGravado", adCurrency, , adFldIsNullable
  rs.Fields.Append "Exento", adCurrency, , adFldIsNullable
  rs.Fields.Append "RetenIva", adCurrency, , adFldIsNullable
  rs.Fields.Append "RetenPercep", adCurrency, , adFldIsNullable
  rs.Fields.Append "Tasa", adCurrency, , adFldIsNullable
  rs.Fields.Append "Debito", adCurrency, , adFldIsNullable
  rs.Fields.Append "Total", adCurrency, , adFldIsNullable
  rs.Open
  
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

