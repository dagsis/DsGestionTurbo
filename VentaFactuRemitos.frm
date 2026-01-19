VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VentaFactuRemitos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturar Remitos"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   6285
   Begin VB.CheckBox ChkDatos 
      Caption         =   "Datos del Remito"
      Height          =   240
      Left            =   150
      TabIndex        =   12
      Top             =   2880
      Width           =   1875
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5115
      TabIndex        =   0
      Top             =   2835
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
      Height          =   2640
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   5970
      Begin VB.Frame Frame2 
         Height          =   930
         Left            =   195
         TabIndex        =   10
         Top             =   1620
         Width           =   5550
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   660
            Left            =   90
            TabIndex        =   11
            Top             =   165
            Width           =   5340
            _ExtentX        =   9419
            _ExtentY        =   1164
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4980
         TabIndex        =   5
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   255
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "TxtCliente"
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4980
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   255
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "TxtCliente"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente :"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
         Height          =   195
         Left            =   255
         TabIndex        =   8
         Top             =   330
         Width           =   1350
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
         Left            =   1155
         TabIndex        =   7
         Top             =   1260
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
         Left            =   1155
         TabIndex        =   6
         Top             =   600
         Width           =   3705
      End
   End
End
Attribute VB_Name = "VentaFactuRemitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsClienteL
Dim crs As ClsLectura
Public Rs As ADODB.Recordset, rRem As ADODB.Recordset
Dim RsCv As ADODB.Recordset, Rsd As ADODB.Recordset
Dim rRemitos As ADODB.Recordset
Private Sub CmdAplicar_Click()
  ProgressBar1.Min = 0
  TraerRemitos
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
  Dim crs As ClsLectura
  Set crs = New ClsLectura
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs = crs.TraerTodos("Clientes", "Id")
     Rs.MoveFirst
     Rs.Find "Id=" & nDat
     Set Rs = crs.TraerRS("ClienteTraerUno", nDat, True)
     LinkearTexto Index
  End If
  Set cRsl = Nothing

End Sub
Private Sub LinkearTexto(Index As Integer)
  TxtCliente(Index).text = Rs!Cliente
  LblCliente(Index).Caption = Rs!RazonSocial
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
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
  Set Rs = cRsl.ClientesAlfabetico(0, False)
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Cliente
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Cliente
  LblCliente(1).Caption = Rs!RazonSocial
End Sub
Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub
Private Sub TxtCliente_LostFocus(Index As Integer)
  CmdAplicar.Enabled = False
  LblCliente(Index).Caption = ""
  If TxtCliente(Index).text <> "" Then
     If cRsl.BuscarNombreCliente(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cRsl.BuscarNombreCliente(TxtCliente(Index).text)
        CmdAplicar.Enabled = True
     End If
  End If
End Sub
Private Sub TraerRemitos()
  Dim cRem As ClsComprobantesL, i As Long
  
  Set cRem = New ClsComprobantesL
  
  Set rRem = cRem.TraerTodosRemitosPend(LblCliente(0).Caption, LblCliente(1).Caption)
  If rRem.RecordCount <> 0 Then
     rRem.MoveLast
     ProgressBar1.Max = rRem.RecordCount
     rRem.MoveFirst
     i = 1
     Do While Not rRem.EOF
        CrearRsVenta
        CrearRsDetalles
        GrabarTodo
        ProgressBar1.Value = i
        rRem.MoveNext
        i = i + 1
     Loop
  Else
     MsgBox "No Existen Remitos Pendientes", vbCritical, "Atención"
  End If
End Sub
Private Sub CrearRsVenta()
 On Error GoTo errHandler

 Set RsCv = New ADODB.Recordset
 Set crs = New ClsLectura
 
 Set RsCv = crs.RsVacio("CabComprobantes", "Id", "N")
 
Exit Sub

errHandler:
   ManejaErrores
 
End Sub

Private Sub CrearRsDetalles()
 On Error GoTo errHandler

 Set Rsd = New ADODB.Recordset
 Set crs = New ClsLectura

 Set Rsd = crs.RsVacio("DetallesComprobantes", "Id", "N")

Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub GrabarTodo()
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cTipo As ClsComprobantesL
  Dim rCuen As ADODB.Recordset, rValo As ADODB.Recordset, rApli As ADODB.Recordset
  Dim rReci As ADODB.Recordset, rCant As ADODB.Recordset, rCaja As ADODB.Recordset
  Dim cVeri As ClsComprobantesL, cCuota As ClsComprobantesL, cDias As ClsComprobantesL, cStock As ClsComprobantesL
  Dim i As Byte, nCuotas As Byte, nDias As Integer, rValoAux As ADODB.Recordset
  Dim cPre As ClsProductoL, cTotal As Currency
  Dim nCanti As Single
  
  ' declaracion para clientes que emiten pagare
  
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cTipo = New ClsComprobantesL
  Set rCuen = New ADODB.Recordset
  Set cVeri = New ClsComprobantesL
  Set cCuota = New ClsComprobantesL
  Set cDias = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  Set rReci = New ADODB.Recordset
  Set rCant = New ADODB.Recordset
  Set rCaja = New ADODB.Recordset
  
  Set rReci = crs.RsVacio("CuentaCorriente", "Id", "N")
  Set rValo = crs.RsVacio("Valores", "Id", "N")
  Set rApli = crs.RsVacio("Aplicaciones", "Id", "N")
  Set rCant = crs.RsVacio("Cantidades", "Id", "N")
  Set rCaja = crs.RsVacio("Caja", "Id", "N")
  
  ' Grabar Cabecera Venta
  cGrab.ActualizarMovimiento 5
    
  RsCv.AddNew
 ' RsCv!Id = 0
  RsCv!tipo = 1
  RsCv!Movimiento = cNum.TraerUltimoMovimiento
  RsCv!Comprobante = 5
  RsCv!fecha = Date
  RsCv!FechaIva = Date
  RsCv!Numero = Format(cNum.TraerUltimoNumero(RsCv!Comprobante), "00000000")
  RsCv!Sucursal = Format(cNum.TraerSucursalNumero(RsCv!Comprobante), "0000")
  RsCv!Cliente = rRem!Cliente
  RsCv!Lista = 1
  RsCv!Vendedor = 1
  RsCv!Caja = 1
  RsCv!CondVenta = 2
  RsCv!Neto = 0
  RsCv!Iva1 = 0
  RsCv!Iva2 = 0
  RsCv!Descuentos = 0
  RsCv!Financiacion = 0
  Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
         Case 1
              RsCv!Debe = 0
              RsCv!Haber = 0
         Case 2
              RsCv!Debe = 0
              RsCv!Haber = 0
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = ""
  RsCv!Anulado = 0
  RsCv.Update
  
  Dim cRem As ClsComprobantesL, rRem1 As ADODB.Recordset
  Dim rAux As ADODB.Recordset
  
  
  Set rAux = New ADODB.Recordset
  Set cRem = New ClsComprobantesL
  Set rRem1 = cRem.TraerClientesRemitos(rRem!Cliente)
  
  rAux.Fields.Append "Descripcion", adVarChar, 80
  rAux.Fields.Append "Movimiento", adDouble
  rAux.Open
  Do While Not rRem1.EOF
     rAux.AddNew
     If ChkDatos.Value = 1 Then
        rAux!Descripcion = rRem1!fecha & " " & rRem1!Descripcion & " " & Format(rRem1!Sucursal, "0000") & "-" & Format(rRem1!Numero, "00000000")
     Else
        rAux!Descripcion = rRem1!fecha
     End If
     rAux!Movimiento = rRem1!Movimiento
     rRem1.MoveNext
     rAux.Update
  Loop
  
  Set rRemitos = New ADODB.Recordset
  rRemitos.Fields.Append "Movimiento", adDouble
  rRemitos.Fields.Append "Descripcion", adVarChar, 50
  rRemitos.Open
  
  If rAux.RecordCount <> 0 Then
     rAux.MoveFirst
  Else
     MsgBox "Problemas con el Cliente :" & rRem!Cliente & ". Verifique en Stock Informes, Remitos Pendientes.", vbCritical, "Atención"
     Exit Sub
  End If
  Do While Not rAux.EOF
     rRemitos.AddNew
     rRemitos!Movimiento = rAux!Movimiento
     rRemitos!Descripcion = rAux!Descripcion
     rRemitos.Update
     rAux.MoveNext
  Loop
  ' Devoluciones
  If rRemitos.RecordCount <> 0 Then
     Set rRem1 = cRem.TraerClientesRemitosDev(rRem!Cliente)
     Do While Not rRem1.EOF
        rRemitos.AddNew
        rRemitos!Movimiento = rRem1!Movimiento
        If ChkDatos.Value = 1 Then
           rRemitos!Descripcion = rRem1!fecha & " " & rRem1!Descripcion & " " & Format(rRem1!Sucursal, "0000") & "-" & Format(rRem1!Numero, "00000000")
        Else
            rRemitos!Descripcion = rRem1!fecha
        End If
        rRemitos.Update
        rRem1.MoveNext
     Loop
  End If

  Set rAux = New ADODB.Recordset
  Set cPre = New ClsProductoL
  Set rRem1 = rRemitos
  If Not rRem1 Is Nothing Then
     If rRem1.RecordCount <> 0 Then
        rRem1.MoveFirst
        Set cRem = New ClsComprobantesL
        Do While Not rRem1.EOF
           Set rAux = cRem.TraerDetallesMovimiento(rRem1!Movimiento)
           Do While Not rAux.EOF
              Rsd.AddNew
             ' Rsd!Id = 0
              Rsd!Producto = rAux!Producto
              Rsd!Descripcion = rRem1!Descripcion + " " & rAux!Descripcion
              nCanti = rAux!Cantidad
              If cRem.SumaCantidad(rAux!Comprobante) = False Then
                 nCanti = -nCanti
              End If
              Rsd!Cantidad = nCanti
              Rsd!PrecioUnitario = rAux!PrecioUnitario
              Rsd!PrecioTotal = nCanti * rAux!PrecioUnitario
              cTotal = cTotal + Rsd!PrecioTotal
              Rsd.Update
              rAux.MoveNext
           Loop
           rRem1.MoveNext
        Loop
        RsCv!Neto = cTotal
        Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
               Case 1
                    RsCv!Debe = cTotal
                    RsCv!Haber = 0
               Case 2
                    RsCv!Debe = 0
                    RsCv!Haber = cTotal
        End Select
        RsCv.Update
    End If
  End If

  ' Detalle Venta ya esta Armado en el recordset Rsd
  ' Forma de Pagos,Valores,Aplicaciones,Cantidades,Caja
  rReci.AddNew
'  rReci!Id = 0
  rReci!fecha = Date
  rReci!Cliente = rRem!Cliente
  rReci!Venta = RsCv!Movimiento
  rReci!Comprobante = 5
  rReci!Numero = cTipo.TraerUltimoNumero(rReci!Comprobante)
  rReci!Sucursal = cTipo.TraerSucursalNumero(rReci!Comprobante)
  rReci!Cuota = 1
  rReci!Debe = Format(cTotal, nCantDecimales)
  rReci!Haber = 0
  rReci!Anulado = 0
  rReci.Update
'  ' Actualizo Stock
  If cStock.ActualizaStock(RsCv!Comprobante) = True Then
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        rCant.AddNew
        rCant!Producto = Rsd!Producto
        rCant!Deposito = cStock.DepositoVenta
        Select Case cTipo.QueMoviStock(RsCv!Comprobante)
               Case "Suma"
                     rCant!CantidadDebe = Rsd!Cantidad
                     rCant!CantidadHaber = 0
               Case "Resta"
                     rCant!CantidadDebe = 0
                     rCant!CantidadHaber = Rsd!Cantidad
        End Select
        rCant!Anulado = 0
        rCant.Update
        Rsd.MoveNext
     Loop
     Rsd.MoveFirst
  End If

'
' If RsCv!Movimiento >= 90 Then
'    MsgBox "Periodo de Demostración Finalizado", vbCritical, "Atención"
'    Exit Sub
' Else
    cGrab.Agregar RsCv, Rsd, rReci, rValo, rApli, rCant, rCaja, True, rRem1
' End If
'
'
'  ' Fin del proceso de facturacion fiscal
'
  Set cGrab = Nothing
  Set cNum = Nothing
  Set cTipo = Nothing
  Set cVeri = Nothing
  Set cCuota = Nothing
  Set cDias = Nothing
  Set cStock = Nothing
Exit Sub
 
errores:
  MsgBox Err.Description, vbCritical, "Atención"
End Sub

