VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaInfResumenDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Venta con Detalles"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   6390
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   225
      TabIndex        =   16
      Top             =   3360
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox ChkCero 
      Caption         =   "Sin Saldo a Cero"
      Height          =   285
      Left            =   420
      TabIndex        =   15
      Top             =   2910
      Width           =   1650
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5295
      TabIndex        =   6
      Top             =   3330
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   225
      TabIndex        =   1
      Top             =   60
      Width           =   5970
      Begin VB.ComboBox CmbZona 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   2205
      End
      Begin VB.CheckBox ChkActivos 
         Caption         =   "Solo Activos"
         Height          =   210
         Left            =   4470
         TabIndex        =   18
         Top             =   2865
         Width           =   1200
      End
      Begin VB.CheckBox ChkIva 
         Caption         =   "Iva Incluido en el Unitario"
         Height          =   195
         Left            =   1965
         TabIndex        =   17
         Top             =   2865
         Width           =   2280
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4905
         TabIndex        =   8
         Top             =   1695
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   180
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "TxtCliente"
         Top             =   1695
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4920
         TabIndex        =   7
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   180
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "TxtCliente"
         Top             =   1035
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   0
         Left            =   825
         TabIndex        =   4
         Top             =   2235
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   708771841
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   1
         Left            =   3045
         TabIndex        =   5
         Top             =   2235
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   708837377
         CurrentDate     =   37021
      End
      Begin VB.Label Label5 
         Caption         =   "Zona :"
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente :"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   765
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   165
         Left            =   2490
         TabIndex        =   12
         Top             =   2235
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   210
         Left            =   195
         TabIndex        =   11
         Top             =   2235
         Width           =   615
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1095
         TabIndex        =   10
         Top             =   1695
         Width           =   3705
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   1035
         Width           =   3705
      End
   End
End
Attribute VB_Name = "VentaInfResumenDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsClienteL
Public Rs As ADODB.Recordset

Private Sub CmdAplicar_Click()
  CrearRs
  TraerAnterior
  TraerSiguiente
  nImpr = 34
  FrmImpresor.Show
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
  Dim crs As ClsLectura
  Set crs = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs = crs.TraerTodos("Clientes", "Id")
     Rs.MoveFirst
     Rs.Find "Id=" & nDat
     Set Rs = crs.TraerRS("ClienteTraerUno", nDat, True)
     LinkearTexto Index
     If Index = 0 Then
        TxtCliente(1).text = Rs!Cliente
        LblCliente(1).Caption = Rs!RazonSocial
     End If
  End If
  Set cRsl = Nothing

End Sub
Private Sub LinkearTexto(Index As Integer)
  TxtCliente(Index).text = Rs!Cliente
  LblCliente(Index).Caption = Rs!RazonSocial
End Sub

Private Sub DtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
     KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  ChkIva.Value = 1
  DtpFecha(0).Value = Date - 30
  DtpFecha(1).Value = Date
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
End Sub

Private Sub Form_Load()
  Set cRsl = New ClsClienteL
  Set Rs = New ADODB.Recordset

  Limpiar
  CargarEtiquetas
  
End Sub

Private Sub CargarEtiquetas()
  Dim cRcli As ClsLectura
  
  Dim Rz As ADODB.Recordset

  Set Rs = cRsl.ClientesAlfabetico(0, False)
  
  Set cRcli = New ClsLectura
  
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Cliente
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Cliente
  LblCliente(1).Caption = Rs!RazonSocial
  
  
  CmbZona.AddItem "TODAS"
  CmbZona.ItemData(CmbZona.NewIndex) = 0
  
  Set Rz = cRcli.TraerRS("ZonaTraerTodos", , True)
  
  Do While Not Rz.EOF
     CmbZona.AddItem Rz!Descripcion
     CmbZona.ItemData(CmbZona.NewIndex) = Rz!Zona
     Rz.MoveNext
  Loop
  CmbZona.ListIndex = 0
  
End Sub
Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub
Private Sub TxtCliente_LostFocus(Index As Integer)
  CmdAplicar.Enabled = False
  LblCliente(Index).Caption = ""
  Set cRsl = New ClsClienteL
  If TxtCliente(Index).text <> "" Then
     If cRsl.BuscarNombreCliente(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cRsl.BuscarNombreCliente(TxtCliente(Index).text)
        TxtCliente(1).text = TxtCliente(Index).text
        LblCliente(1).Caption = LblCliente(Index).Caption
        CmdAplicar.Enabled = True
     End If
  End If
End Sub

Private Sub CrearRs()
  
  Set Rs = New ADODB.Recordset
  
  Rs.Fields.Append "Cliente", adVarChar, 15
  Rs.Fields.Append "Razon", adVarChar, 50
  Rs.Fields.Append "Domicilio", adVarChar, 50, adFldIsNullable
  Rs.Fields.Append "Postal", adVarChar, 8, adFldIsNullable
  Rs.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Iva", adVarChar, 25, adFldIsNullable
  Rs.Fields.Append "Cuit", adVarChar, 13, adFldIsNullable
  Rs.Fields.Append "Telefono", adVarChar, 20, adFldIsNullable
  Rs.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs.Fields.Append "Comprobante", adVarChar, 40, adFldIsNullable
  Rs.Fields.Append "Numero", adDouble, , adFldIsNullable
  Rs.Fields.Append "Sucursal", adInteger, , adFldIsNullable
  Rs.Fields.Append "Cuota", adInteger, , adFldIsNullable
  Rs.Fields.Append "Debe", adDouble
  Rs.Fields.Append "Haber", adDouble
  Rs.Fields.Append "Detalle", adVarChar, 100, adFldIsNullable
  Rs.Fields.Append "Cantidad", adDouble, , adFldIsNullable
  Rs.Fields.Append "Unitario", adDouble, , adFldIsNullable
  Rs.Fields.Append "Total", adDouble, , adFldIsNullable
  Rs.Open
End Sub

Private Sub TraerAnterior()
  Dim RsAux As ADODB.Recordset
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsClienteL
  
  Set RsAux = cRsl.ResumenSaldoAnterior(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, CmbZona.ItemData(CmbZona.ListIndex))
  
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Cliente = RsAux!Cliente
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = IIf(Trim(RsAux!CodigoPostal) = "", 0, RsAux!CodigoPostal)
     Rs!Localidad = RsAux!Ciudad
     Rs!Provincia = RsAux!Provincia
     Rs!Iva = RsAux!TipoIva
     Rs!Cuit = RsAux!Cuit
     Rs!Telefono = Left(RsAux!Telefono, 20)
     Rs!fecha = DtpFecha(0).Value
     Rs!Comprobante = Null
     Rs!Numero = Null
     Rs!Sucursal = Null
     Rs!Cuota = Null
     Rs!Debe = IIf(IsNull(RsAux!SumaDeDebe), 0, RsAux!SumaDeDebe)
     Rs!Haber = IIf(IsNull(RsAux!SumaDeHaber), 0, RsAux!SumaDeHaber)
     Rs!Detalle = "SALDO ANTERIOR"
     Rs!Total = RsAux!SumaDeDebe - RsAux!SumaDeHaber
     If Rs!Total = 0 Then
        Rs.CancelUpdate
     Else
        Rs.Update
     End If
     RsAux.MoveNext
  Loop
  
End Sub

Private Sub TraerSiguiente()
  Dim RsAux As ADODB.Recordset, cComp As ClsComprobantesL, nVenta As Long, j As Long
  Dim cRx As ClsLectura, bPercep As Long
  
  Set RsAux = New ADODB.Recordset
  Set cComp = New ClsComprobantesL
  Set cRx = New ClsLectura
  
  j = 1
  ProgressBar1.Value = 1
  Set RsAux = cRsl.ResumenCuentaDetalles(LblCliente(0).Caption, LblCliente(1).Caption, DtpFecha(0).Value, DtpFecha(1).Value, CmbZona.ItemData(CmbZona.ListIndex))
   
  nVenta = 0
  If RsAux.RecordCount <> 0 Then
     ProgressBar1.Max = RsAux.RecordCount
  End If
  bPercep = 0
  Do While Not RsAux.EOF
     Rs.AddNew
     Rs!Cliente = RsAux!Cliente
     Rs!Razon = RsAux!RazonSocial
     Rs!Domicilio = RsAux!Domicilio
     Rs!Postal = IIf(Trim(RsAux!CodigoPostal) = "", 0, RsAux!CodigoPostal)
     Rs!Localidad = RsAux!Ciudad
     Rs!Provincia = "" ' RsAux![Provincias.Descripcion]
     Rs!Iva = "" 'RsAux![TipoIva.Descripcion]
     Rs!Cuit = "" ' RsAux!Cuit
     Rs!Telefono = "" ' RsAux!Telefono
     Rs!fecha = RsAux!fecha
     Rs!Comprobante = RsAux!Comprobante & " " & RsAux!Motivo
     Rs!Numero = RsAux!Numero
     Rs!Sucursal = 0 ' RsAux!Sucursal
     Rs!Cuota = 0 ' RsAux!Cuota
     Rs!Debe = 0 'RsAux!Debe
     Rs!Haber = 0 ' RsAux!Haber
     Rs!Detalle = RsAux!Detalle
     Rs!Cantidad = RsAux!Cantidad
     If cRx.TraerValorDeUnCampo("Comprobantes", "Costo", "Comprobante='" & RsAux!Comprobante & "'") = 1 Then
         Rs!Unitario = RsAux!PrecioUnitario + IIf(ChkIva.Value = 1, (RsAux!Impuesto / IIf(RsAux!Cantidad = 0, 1, RsAux!Cantidad)), 0) - (RsAux!Descuento / RsAux!Cantidad)
         Rs!Total = (Rs!Unitario * Rs!Cantidad) + RsAux!Impuesto - RsAux!Descuento
         If cRx.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & RsAux!Comprobante & "'") = 8 Then
            Rs!Total = -Rs!Unitario
         End If
        If cRx.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & RsAux!Comprobante & "'") = 5 Then
           Rs!Unitario = -RsAux!PrecioTotal
           Rs!Total = -(RsAux!PrecioTotal + RsAux!Impuesto)
        End If
     Else
        Rs!Unitario = RsAux!PrecioUnitario
        If cRx.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & RsAux!Comprobante & "'") = 4 Then
           Rs!Total = -cRx.TraerValorDeUnCampo("CabComprobantes", "Haber", "Movimiento=" & RsAux!Movimiento) '-RsAux!PrecioTotal
           Rs!Cantidad = 1
           Rs!Unitario = 0
           If RsAux!Lista = 0 Then
              Rs!Detalle = "COBRANZA"
           Else
              Rs!Detalle = Null
           End If
        End If
        If cRx.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & RsAux!Comprobante & "'") = 5 Then
           Rs!Total = -RsAux!PrecioTotal
        End If
        If cRx.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & RsAux!Comprobante & "'") = 8 Then
           Rs!Total = -RsAux!PrecioTotal
        End If
        If cRx.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & RsAux!Comprobante & "'") = 9 Then
           Rs!Total = RsAux!PrecioTotal
        End If
     End If
     If IsNull(Rs!Detalle) Then
        Rs.Update
        Rs.Delete
     Else
        If cComp.SumoOresto(RsAux!Comprobante) = False Then
           Rs!Total = RsAux!PrecioTotal
           If cRx.TraerValorDeUnCampo("ListadePrecio", "PrecioIva", "ListaPrecio=" & RsAux!Lista) = 0 Then
              Rs!Total = Rs!Total + RsAux!Impuesto - -RsAux!Descuento
           End If
           If cRx.TraerValorDeUnCampo("Comprobantes", "Costo", "Comprobante='" & RsAux!Comprobante & "'") = 1 Then
              Rs!Total = (RsAux!PrecioUnitario * Rs!Cantidad) + RsAux!Impuesto - RsAux!Descuento 'rs!Unitario * rs!Cantidad '
              Rs!Unitario = Rs!Total / Rs!Cantidad
           Else
               Rs!Total = (RsAux!PrecioUnitario * Rs!Cantidad) - RsAux!Descuento
               If Rs!Total > 0 Then
                  Rs!Unitario = Rs!Total / Rs!Cantidad
               Else
                  Rs!Unitario = 0
               End If
           End If
           If cRx.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & RsAux!Comprobante & "'") = 9 Then
              Rs!Total = RsAux!PrecioTotal
           End If
           Rs.Update
           ' Percepciones
           If cRx.TraerValorDeUnCampo("CabComprobantes", "PIvaCompras", "Movimiento=" & RsAux!Movimiento) <> 0 Then
               If bPercep <> RsAux!Movimiento Then
                    Rs.AddNew
                    Rs!Cliente = RsAux!Cliente
                    Rs!Razon = RsAux!RazonSocial
                    Rs!Domicilio = RsAux!Domicilio
                    Rs!Postal = IIf(Trim(RsAux!CodigoPostal) = "", 0, RsAux!CodigoPostal)
                    Rs!Localidad = RsAux!Ciudad
                    Rs!Provincia = ""
                    Rs!Iva = ""
                    Rs!Cuit = ""
                    Rs!Telefono = ""
                    Rs!fecha = RsAux!fecha
                    Rs!Comprobante = RsAux!Comprobante & " " & RsAux!Motivo
                    Rs!Numero = RsAux!Numero
                    Rs!Sucursal = 0 ' RsAux!Sucursal
                    Rs!Cuota = 0 ' RsAux!Cuota
                    Rs!Debe = 0 'RsAux!Debe
                    Rs!Haber = 0 ' RsAux!Haber
                    Rs!Detalle = "PERCEPCIONES"
                    Rs!Total = cRx.TraerValorDeUnCampo("CabComprobantes", "PIvaCompras", "Movimiento=" & RsAux!Movimiento)
                    Rs.Update
                    bPercep = RsAux!Movimiento
                End If
           Else
               bPercep = 0
           End If
        Else
           Rs.Update
        End If
     End If
   ProgressBar1.Value = j
   j = j + 1
   RsAux.MoveNext
  Loop

  Set cRx = Nothing

End Sub

