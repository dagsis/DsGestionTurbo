VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasNotaCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas de Creditos/Debitos a Proveedores"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   9360
   Begin VB.TextBox TxtCliente 
      Height          =   285
      Left            =   1065
      MaxLength       =   15
      TabIndex        =   29
      Text            =   "TxtCliente"
      Top             =   570
      Width           =   795
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   750
      Index           =   5
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5235
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Anular"
      Height          =   750
      Index           =   4
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   960
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   3
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3570
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   2
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2700
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   750
      Index           =   1
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1830
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   750
      Index           =   0
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   90
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4530
      Left            =   105
      TabIndex        =   4
      Top             =   1470
      Width           =   8010
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   2835
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2895
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2895
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   5265
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2895
         Width           =   870
      End
      Begin VB.Frame Frame2 
         Height          =   960
         Left            =   15
         TabIndex        =   11
         Top             =   3570
         Width           =   6210
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   1890
         MaxLength       =   40
         TabIndex        =   10
         Text            =   "TxtDetalle"
         Top             =   2550
         Width           =   3510
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   1905
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2895
         Width           =   870
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   690
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "TxtPrecio"
         Top             =   3195
         Width           =   1095
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   855
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "TxtCa"
         Top             =   2880
         Width           =   930
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   5445
         TabIndex        =   6
         Top             =   2565
         Width           =   705
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   705
         MaxLength       =   25
         TabIndex        =   5
         Text            =   "TxtProducto"
         Top             =   2550
         Width           =   1080
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2130
         Left            =   165
         TabIndex        =   15
         Top             =   255
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   3757
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTotal"
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
         Left            =   6840
         TabIndex        =   22
         Top             =   4215
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Total :"
         Height          =   225
         Left            =   6270
         TabIndex        =   21
         Top             =   4215
         Width           =   495
      End
      Begin VB.Label LblNeto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNeto"
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
         Left            =   6840
         TabIndex        =   20
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Neto :"
         Height          =   255
         Left            =   6270
         TabIndex        =   19
         Top             =   2550
         Width           =   525
      End
      Begin VB.Label Label14 
         Caption         =   "Precio :"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   3195
         Width           =   750
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   90
         TabIndex        =   17
         Top             =   2895
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   90
         TabIndex        =   16
         Top             =   2565
         Width           =   735
      End
   End
   Begin VB.ComboBox CmbCaja 
      Height          =   315
      Left            =   6090
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1050
      Width           =   2010
   End
   Begin VB.ComboBox CmBVendedor 
      Height          =   315
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1050
      Width           =   2025
   End
   Begin VB.CommandButton CmdClientes 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   5055
      TabIndex        =   1
      Top             =   555
      Width           =   735
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   105
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   315
      Left            =   4500
      TabIndex        =   30
      Top             =   90
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   63897601
      CurrentDate     =   36783
   End
   Begin VB.Label LblNumero 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblNumero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6615
      TabIndex        =   40
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label LbLIva 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LbLIva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6285
      TabIndex        =   39
      Top             =   555
      Width           =   1815
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
      Height          =   285
      Left            =   1875
      TabIndex        =   38
      Top             =   570
      Width           =   3105
   End
   Begin VB.Label Label9 
      Caption         =   "Caja :"
      Height          =   240
      Left            =   5130
      TabIndex        =   37
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label Label8 
      Caption         =   "Operador :"
      Height          =   225
      Left            =   135
      TabIndex        =   36
      Top             =   1050
      Width           =   870
   End
   Begin VB.Label Label5 
      Caption         =   "Iva :"
      Height          =   225
      Left            =   5865
      TabIndex        =   35
      Top             =   570
      Width           =   405
   End
   Begin VB.Label Label4 
      Caption         =   "Proveedor :"
      Height          =   255
      Left            =   135
      TabIndex        =   34
      Top             =   570
      Width           =   840
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   3915
      TabIndex        =   33
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   240
      Left            =   5865
      TabIndex        =   32
      Top             =   90
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   135
      TabIndex        =   31
      Top             =   90
      Width           =   1155
   End
End
Attribute VB_Name = "ComprasNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean                     ' Suich Cabecera Venta
Dim sw1 As Boolean                    ' Suich Detalles Venta
Dim sw2 As Boolean                    ' Suich Valores
Dim bCliente As Boolean
Dim cImpor As Currency                ' Importe Total de Venta
Dim Rs_Ven As ADODB.Recordset         ' Cabecera Venta Tabla Uno
Dim RsCv As ADODB.Recordset           ' Cabecera Virtual Venta
Dim Rsd As ADODB.Recordset            ' Cabecera Vitual VentaDetalle
Dim Rs As ADODB.Recordset             ' Cabecera Venta Tabla Todos
Dim Rs_Sis As ADODB.Recordset         ' Tabla Sistema, Ultima Venta
Public nMovi As Long
Public nCas As Integer
Dim cRsl As ClsLectura
Private Sub CmbCaja_Click()
  VerEstado
  VerCabecera
End Sub
Private Sub CmbComprobante_Click()
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler

  If sw = False Then
     Set cNum = New ClsComprobantesL
     nNum = Format(cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
     nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
     LblNumero.Caption = nSuc & "-" & nNum
  End If
  Set cNum = Nothing
  VerEstado
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmbFormaPago_Click()
  VerEstado
  VerCabecera
End Sub

Private Sub CmbVendedor_Click()
  VerEstado
  VerCabecera
End Sub
Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Buscar
         Case 2
              Anterior
         Case 3
              Siguiente
         Case 4
              Anular
         Case 5
              Salir
  End Select
End Sub
Private Sub CmdClientes_Click()
  Dim RsC As ADODB.Recordset, cRiL As ClsClienteL, cRcL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRiL = New ClsClienteL
  Set cRcL = New ClsClienteL
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 2
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ProveedorTraerUno", nDat, True)
     TxtCliente.Text = RsC!Proveedor
     LblCliente.Caption = RsC!RazonSocial
     LbLIva.Caption = cRiL.BuscarTipoIvaProv(RsC!Proveedor)
     CmBVendedor.ListIndex = 0
     CmBVendedor.SetFocus
     CalcularTotales
     bCliente = True
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub CmdClientes_LostFocus()
  On Error GoTo errHandler
  If bCliente = False Then
     If TxtCliente.Enabled = True Then
        TxtCliente.SetFocus
        MsgBox "El Proveedor no es el Correcto", vbCritical, "Atención"
     End If
  End If
Exit Sub
errHandler:
 Exit Sub
End Sub
Private Sub CmdProductos_Click()
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  nDat = 0
  nLista = 1
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.Text = RsP!Producto
     TxtDetalle.Text = RsP!Descripcion
'     TxtPrecio.Text = Format(cRpl.TraerPrecio(CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.Text), nCantDecimales)
     TxtCantidad.SetFocus
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub DtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
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
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.Text = ""
      End If
  Next ctl
  DTPFecha.Value = Date
  LblCliente.Caption = ""
  LblNumero.ForeColor = &H80000012
  LblNumero.Caption = "0000-00000000"
  LbLIva.Caption = ""
  LblNeto.Caption = "0.00"
  LblTotal.Caption = "0.00"
  CmbComprobante.ListIndex = -1
  CmBVendedor.ListIndex = -1
  CmbCaja.ListIndex = -1
End Sub
Private Sub HabilitarTodo(bEstado As Boolean, bGrid As Boolean)
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
         Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
         Or TypeOf ctl Is ComboBox Then
         ctl.Enabled = bEstado
      End If
  Next ctl
  Grid1.Enabled = bGrid
End Sub
Private Sub LinkearTexto()
  Dim cRcL As ClsClienteL, cRiL As ClsClienteL, nSuc As String * 4, nNum As String * 8
  Dim cComp As ClsComprobantesL
  Dim cCompr As ClsLectura, cLista As ClsLectura, cVend As ClsLectura, cCaja As ClsLectura
  On Error GoTo errHandler
  
  Set cCompr = New ClsLectura
  Set cLista = New ClsLectura
  Set cVend = New ClsLectura
  Set cCaja = New ClsLectura
  Set cRcL = New ClsClienteL
  Set cRiL = New ClsClienteL
  Set cComp = New ClsComprobantesL
  
  CmbComprobante.Text = cCompr.DatoCombo("Comprobantes", "Id", "Descripcion", Rs!Comprobante)
  DTPFecha.Value = Rs!fecha
  nNum = Format(Rs!Numero, "00000000")
  nSuc = Format(Rs!sucursal, "0000")
  LblNumero.ForeColor = &H80000012
  If cComp.ComprobanteAnulado(Rs!Movimiento) = True Then
     LblNumero.ForeColor = &H40C0&
  End If
  LblNumero.Caption = nSuc & "-" & nNum
  nMovi = Rs!Movimiento
  TxtCliente.Text = Rs!Cliente
  LblCliente.Caption = cRcL.BuscarNombreProveedor(Rs!Cliente)
  LbLIva.Caption = cRiL.BuscarTipoIva(Rs!Cliente)
  CmBVendedor.Text = cVend.DatoCombo("Vendedores", "Vendedor", "Descripcion", Rs!Vendedor)
  CmbCaja.Text = cCaja.DatoCombo("Cajas", "Caja", "Descripcion", Rs!Caja)
  LblNeto.Caption = Format(Rs!Neto, nCantDecimales)
  LblTotal.Caption = Format(Rs!Debe + Rs!Haber, nCantDecimales)
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub CmdDetalle_Click(Index As Integer)
  Select Case Index
         Case 0
              NuevoRsD
         Case 1
              BorrarRsD
         Case 2
              ModificarRsD
         Case 3
              CancelarRsd
  End Select
End Sub
Private Sub Form_Load()
 On Error GoTo errHandler

 Set Rs = New ADODB.Recordset
 Set Rs_Sis = New ADODB.Recordset
 
 sw = True
 sw1 = True
 sw2 = True
 nCas = 0
 nLlama = 5
 
 Botones True, False, False, False, False, True
 BotonDetalles False, False, False, False
 BotonBuscar False, False

 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Buscar", 0)
' CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
' CmdBotones(3).Picture = LoadResPicture("Buscar", 0)
 CmdBotones(4).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(5).Picture = LoadResPicture("Salir", 0)
 
 CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
 CmdDetalle(1).Picture = LoadResPicture("Borrar", 0)
 CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
 CmdDetalle(3).Picture = LoadResPicture("Cancelar", 0)
 
 CargarCombos
 CrearRsDetalles
 Limpiar
 HabilitarTodo False, False
 HabilitarDetalles False
 
 ArmaRsParaMovimiento
 Me.Top = 0
 Me.Left = 0
 
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CargarCombos()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=5 and Suc=" & nSucursal
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmBVendedor, "Vendedores", "Vendedor", "Descripcion", ""
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bMod As Boolean, bBus As Boolean, _
  bVal As Boolean, bSal As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bBorr
  CmdBotones(2).Enabled = bMod
  CmdBotones(3).Enabled = bBus
  CmdBotones(4).Enabled = bVal
  If LblNumero.ForeColor = &H40C0& Then
     CmdBotones(4).Enabled = False
  End If
  CmdBotones(5).Enabled = bSal
End Sub
Private Sub BotonDetalles(bAgr As Boolean, bBorr As Boolean, bMod As Boolean, bCan As Boolean)
  CmdDetalle(0).Enabled = bAgr
  CmdDetalle(1).Enabled = bBorr
  CmdDetalle(2).Enabled = bMod
  CmdDetalle(3).Enabled = bCan
End Sub
Private Sub BotonBuscar(bCli As Boolean, bPro As Boolean)
  CmdClientes.Enabled = bCli
  CmdProductos.Enabled = bPro
End Sub
Private Sub Form_Resize()
  CmdBotones(0).SetFocus
End Sub
Private Sub CrearRsVenta()
 On Error GoTo errHandler

 Set RsCv = New ADODB.Recordset
 Set cRsl = New ClsLectura
 
 Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
 
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub CrearRsDetalles()
 On Error GoTo errHandler

 Set Rsd = New ADODB.Recordset
 Set cRsl = New ClsLectura
 
 Set Rsd = cRsl.RsVacio("DetallesComprobantes", "Id", "N")
 CabGrid
 
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub CabGrid()
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.Bold = True
  With Grid1
      .Columns(0).Visible = False
      .Columns(1).Visible = False
      .Columns(2).Caption = "Producto"
      .Columns(2).Width = 1000
      .Columns(3).Caption = "Descripción"
      .Columns(3).Width = 5150
      .Columns(4).Visible = False
      .Columns(5).Caption = "Cant."
      .Columns(5).Width = 650
      .Columns(5).Alignment = dbgRight
      .Columns(5).NumberFormat = nCantDecimales
      .Columns(6).Visible = False
      .Columns(7).Visible = False
      .Columns(8).Visible = False
      .Columns(9).Caption = "P.Uni."
      .Columns(9).Width = 800
      .Columns(9).NumberFormat = nCantDecimales
      .Columns(9).Alignment = dbgRight
      .Columns(10).Caption = "Total"
      .Columns(10).Width = 1000
      .Columns(10).NumberFormat = nCantDecimales
      .Columns(10).Alignment = dbgRight
 End With
End Sub
Private Sub NuevoRsD()
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     CmdBotones(5).Enabled = False
     BotonBuscar False, True
     HabilitarTodo False, False
     HabilitarDetalles True
     LimpiarDetalles
     BotonDetalles False, False, False, True
     CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(0).Caption = "Grabar"
     TxtProducto.SetFocus
  Else
     CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
     CmdDetalle(0).Caption = "Agregar"
     Rsd.AddNew
     GrabarRsd
     BotonDetalles True, True, True, False
     BotonBuscar True, False
     CalcularTotales
     sw1 = True
     HabilitarTodo True, True
     HabilitarDetalles False
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub BorrarRsD()
  On Error GoTo errHandler
  If MsgBox("Borra el Producto Seleccionado ?", 20, "Atención") = 6 Then
     LimpiarDetalles
     BotonDetalles True, False, False, False
     Rsd.Delete
     Rsd.MovePrevious
     If Rsd.BOF = False And Rsd.EOF = False Then
        If Rsd.BOF Then
           Rsd.MoveFirst
        End If
        LinkearDetalles
        Grid1.Bookmark = Rsd.Bookmark
        BotonDetalles True, True, True, False
'        CmdBotones(4).Enabled = HabiValores()
        CmdBotones(0).Enabled = False
     Else
        Botones False, False, False, False, False, True
     End If
  End If
  CalcularTotales
Exit Sub
  
errHandler:
   ManejaErrores

End Sub
Private Sub ModificarRsD()
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     HabilitarTodo True, False
     HabilitarDetalles True
     BotonBuscar True, True
     BotonDetalles False, False, True, True
     CmdDetalle(2).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(2).Caption = "Grabar"
  Else
     CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
     CmdDetalle(2).Caption = "Actualizar"
     GrabarRsd
     BotonDetalles True, True, True, False
     BotonBuscar True, False
     CalcularTotales
     sw1 = True
     HabilitarTodo True, True
     HabilitarDetalles False
     CalcularTotales
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CancelarRsd()
  On Error GoTo errHandler
  
  CmdBotones(5).Enabled = True
  If sw1 = False Then
     If CmdDetalle(0).Caption = "Grabar" Then
        If CmdDetalle(0).Enabled = True Then
           If MsgBox("El Detalle no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        End If
        CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
        CmdDetalle(0).Caption = "Nuevo"
        BotonDetalles True, False, False, False
        HabilitarTodo True, True
        HabilitarDetalles False
        BotonBuscar True, False
        LimpiarDetalles
        If Rsd.EOF = False Or Rsd.BOF = False Then
           BotonDetalles True, True, True, False
           LinkearDetalles
        End If
        sw1 = True
     Else
        CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
        CmdDetalle(2).Caption = "Actualizar"
        BotonDetalles True, True, True, False
        HabilitarTodo True, True
        HabilitarDetalles False
        BotonBuscar True, False
        LinkearDetalles
        sw1 = True
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub GrabarRsd()
  On Error GoTo errHandler

'  Rsd!Id = 0
  Rsd!Producto = TxtProducto.Text
  Rsd!Descripcion = TxtDetalle.Text
  Rsd!Cantidad = TxtCantidad.Text
  Rsd!PrecioUnitario = TxtPrecio.Text
  Rsd!PrecioTotal = TxtCantidad.Text * TxtPrecio.Text
  Rsd.Update
  CabGrid
  VerEstado
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub LimpiarDetalles()
 TxtProducto.Text = ""
 TxtCantidad.Text = "1.00"
 TxtDetalle.Text = ""
 TxtPrecio.Text = "0.00"
End Sub
Private Sub LinkearDetalles()
  Dim cRsPl As ClsProductoL
  On Error GoTo errHandler
  
  Set cRsPl = New ClsProductoL
  
  TxtProducto.Text = Rsd!Producto
  TxtCantidad.Text = Format(Rsd!Cantidad, nCantDecimales)
  TxtDetalle.Text = cRsPl.BuscarNombreProducto(TxtProducto.Text)
  TxtPrecio.Text = Format(Rsd!PrecioUnitario, nCantDecimales)
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Public Sub CalcularTotales()
  Dim rReg As Variant, cDescTotal As Currency, cDescParcial As Currency, sClien As String
  Dim cImpIva1 As Currency, cImpIva2 As Currency, i As Byte, cTotPar As Currency
  Dim rAux As ADODB.Recordset
  
  On Error GoTo errHandler
  
  LblNeto.Caption = "0.00"
  ' Calcular Precio Neto
  If Rsd.BOF = False And Rsd.EOF = False Then
     rReg = Rsd.Bookmark
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        LblNeto.Caption = CCur(LblNeto.Caption) + Rsd!PrecioTotal
        LblNeto.Caption = Format(LblNeto.Caption, nCantDecimales)
        Rsd.MoveNext
     Loop
     Rsd.Bookmark = rReg
     ' Calcular Descuento
     Set cRsl = New ClsLectura
     Set rAux = New ADODB.Recordset
     sClien = TxtCliente.Text
     If sClien <> "" Then
        Set rAux = cRsl.TraerRS("ClienteDescuentoTraer", sClien, True)
        cDescTotal = 0
        cDescParcial = 0
        Do While Not rAux.EOF
           cDescParcial = CCur(LblNeto.Caption) * rAux!Porcentaje / 100
           cDescTotal = cDescTotal + cDescParcial
           rAux.MoveNext
        Loop
     Else
        MsgBox "No se Puede Calcular Los Descuentos el Campo Cliente está Vacio", vbCritical, "Atención"
     End If
     cDescParcial = 0
'     cDescParcial = CCur(LblNeto.Caption) * CDbl(TxtDescuento.Text) / 100
     cDescTotal = cDescTotal + cDescParcial
'     LblBonificacion.Caption = Format(cDescTotal, nCantDecimales)
     cTotPar = CCur(LblNeto) - cDescTotal
     ' Calcular Financiación
     ' CalcularImpuestos
     Set cRsl = New ClsLectura
     Set rAux = New ADODB.Recordset
     sClien = TxtCliente.Text
     Set rAux = cRsl.TraerRS("ClienteImpuestoTraer", sClien, True)
     i = 1
     cImpIva1 = 0
     cImpIva2 = 0
     Dim cLista As ClsComprobantesL
     Set cLista = New ClsComprobantesL
'     If cLista.ListaConIva(CmbLista.ItemData(CmbLista.ListIndex)) = False Then
'        Do While Not rAux.EOF
'           If i = 1 Then
'              cImpIva1 = cTotPar
'              LblIva1.Caption = Format(cImpIva1 * rAux!Porcentaje / 100, nCantDecimales)
'              i = 2
'           Else
'              cImpIva2 = cTotPar
'              LblIva2.Caption = Format(cImpIva2 * rAux!Porcentaje / 100, nCantDecimales)
'           End If
'           rAux.MoveNext
'        Loop
'    End If
    ' Calcular Total General
    LblTotal.Caption = CCur(LblNeto.Caption) - cDescTotal
    LblTotal.Caption = Format(LblTotal.Caption, nCantDecimales)
 End If
Exit Sub

errHandler:
   ManejaErrores
 
End Sub

Private Sub TxtCantidad_GotFocus()
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCantidad_LostFocus()
'  Dim cStock As ClsComprobantesL, cDepo As ClsProductoL
'  Dim nDeposito As Integer
'
'  Set cStock = New ClsComprobantesL
'  Set cDepo = New ClsProductoL
'
'  ' Si comprobante controla Stock, Verifica las Cantidades Disponibles
'  If TxtProducto.Text <> "" Then
'     If cStock.ActualizaStock(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
'        If cStock.DepositoVenta <> 0 Then
'           nDeposito = cStock.DepositoVenta
'           If TxtCantidad.Text > cDepo.DameCantidad(TxtProducto.Text, nDeposito) Then
'              MsgBox "Cantidad Superior a la del Deposito de Venta", vbCritical, "Atención"
'              TxtCantidad.Text = cDepo.DameCantidad(TxtProducto.Text, nDeposito)
'           End If
'           TxtCantidad.SetFocus
'        Else
'           MsgBox "No hay Deposito de Venta Definido", vbCritical, "Atención"
'           TxtCantidad.SetFocus
'        End If
'     End If
'  End If
  If TxtCantidad.Text = "" Then TxtCantidad.Text = 1
  TxtCantidad.Text = Format(TxtCantidad.Text, nCantDecimales)
End Sub

Private Sub TxtCliente_Change()
  VerEstado
  VerCabecera
End Sub

Private Sub TxtCliente_GotFocus()
  TxtCliente.SelStart = 0
  TxtCliente.SelLength = Len(TxtCliente.Text)
End Sub

Private Sub TxtCliente_LostFocus()
  Dim cRcL As ClsClienteL, cRiL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set cRcL = New ClsClienteL
  Set cRiL = New ClsClienteL
  LblCliente.Caption = ""
  LbLIva.Caption = ""
  If TxtCliente.Text <> "" Then
     If cRcL.BuscarNombreProveedor(TxtCliente.Text) <> "" Then
        LblCliente.Caption = cRcL.BuscarNombreProveedor(TxtCliente.Text)
        LbLIva.Caption = cRiL.BuscarTipoIvaProv(TxtCliente.Text)
        CmBVendedor.Text = cRsl.DatoCombo("Vendedores", "Vendedor", "Descripcion", cRcL.BuscarVendedor(TxtCliente.Text))
        CmBVendedor.SetFocus
        CalcularTotales
        bCliente = True
     Else
        MsgBox "El Proveedor no es el Correto", vbCritical, "Atención"
        If CmdClientes.Enabled = True Then
           CmdClientes.SetFocus
        End If
        bCliente = False
     End If
  End If
  Set cRcL = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub TxtDetalle_Change()
  If CmdDetalle(0).Caption = "Grabar" Then
     CmdDetalle(0).Enabled = False
     If Len(TxtDetalle.Text) <> 0 Then
        CmdDetalle(0).Enabled = True
     End If
  Else
     CmdDetalle(2).Enabled = False
     If Len(TxtDetalle.Text) <> 0 Then
        CmdDetalle(2).Enabled = True
     End If
  End If
End Sub
Private Sub TxtDetalle_LostFocus()
  If TxtDetalle.Text <> "" Then
     TxtCantidad.SetFocus
  End If
End Sub
Private Sub TxtPrecio_GotFocus()
  TxtPrecio.SelStart = 0
  TxtPrecio.SelLength = Len(TxtPrecio.Text)
End Sub
Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtPrecio) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtPrecio_LostFocus()
  If TxtPrecio.Text = "" Then TxtPrecio.Text = 0
  TxtPrecio.Text = Format(TxtPrecio.Text, nCantDecimales)
  If CmdDetalle(0).Enabled = True Then
     CmdDetalle(0).SetFocus
  End If
End Sub
Private Sub TxtProducto_LostFocus()
  Dim cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set cRpl = New ClsProductoL
  If TxtProducto.Text <> "" Then
     If cRpl.BuscarNombreProducto(TxtProducto.Text) <> "" Then
        TxtDetalle.Text = cRpl.BuscarNombreProducto(TxtProducto.Text)
'        TxtPrecio.Text = Format(cRpl.TraerPrecio(CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.Text), nCantDecimales)
        TxtCantidad.SetFocus
     End If
  End If
  Set cRpl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtProducto.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  TxtPrecio.Enabled = bHabi
  CmdProductos.Enabled = bHabi
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler

  If sw = False And sw1 = True Then
     If Grid1.Row <> -1 Then
        Rsd.Bookmark = Grid1.Bookmark
        LinkearDetalles
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub ArmaRsParaMovimiento()
  Dim cComp As ClsComprobantesL
  On Error GoTo errHandler

  Set cComp = New ClsComprobantesL
  Set Rs_Ven = cComp.TraerTodos("CabComprobantes", "Id", 5)
  Botones True, False, False, False, False, True
  If Rs_Ven.BOF = False And Rs_Ven.EOF = False Then
     Rs_Ven.MoveLast
     Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Ven!ID, True)
     LinkearTexto
     LLenarDetalles
     nCas = 0
     Botones True, True, True, True, True, True
  End If
  Set cRsl = Nothing
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub Siguiente()
  On Error GoTo errHandler
  Rs_Ven.MoveNext
  If Rs_Ven.EOF Then
     Rs_Ven.MoveLast
     Beep
  End If
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Ven!ID, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Anterior()
  On Error GoTo errHandler
  Rs_Ven.MovePrevious
  If Rs_Ven.BOF Then
     Rs_Ven.MoveFirst
     Beep
  End If
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Ven!ID, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub LLenarDetalles()
  Dim rAuxD As ADODB.Recordset
  On Error GoTo errHandler
 
  Set rAuxD = New ADODB.Recordset
  Set cRsl = New ClsLectura
  Set rAuxD = cRsl.TraerRS("DetallesTraerUnos", Rs!Movimiento, True)
  rAuxD.MoveFirst
  Do While Not rAuxD.EOF
     Rsd.AddNew
     Rsd!Movimiento = rAuxD!Movimiento
     Rsd!Producto = rAuxD!Producto
     Rsd!Descripcion = rAuxD!Descripcion
     Rsd!Cantidad = rAuxD!Cantidad
     Rsd!PrecioUnitario = rAuxD!PrecioUnitario
     Rsd!PrecioTotal = rAuxD!PrecioTotal
     Rsd.Update
     rAuxD.MoveNext
  Loop
  CabGrid
  Set cRsl = Nothing
  rAuxD.Close
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub Nuevo()
  On Error GoTo errHandler

  If sw = True Then
     Limpiar
     HabilitarTodo True, False
     HabilitarDetalles False
     Botones False, False, False, False, False, True
     BotonDetalles False, False, False, False
     BotonBuscar True, False
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(5).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(5).Caption = "Cancelar"
     sw = False
     CrearRsVenta
     CrearRsDetalles
     CmbComprobante.SetFocus
     nMovi = 0
     nCas = 1
     CmbComprobante.ListIndex = 0
     Set cRsl = New ClsLectura
     CmBVendedor.Text = cRsl.DatoCombo("Vendedores", "Vendedor", "Descripcion", cRsl.RegPorDefecto("Vendedor"))
     CmbCaja.Text = cRsl.DatoCombo("Cajas", "Caja", "Descripcion", cRsl.RegPorDefecto("Caja"))
  Else
     sw = True
     Botones False, False, False, False, False, False
     GrabarTodo
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
     CmdBotones(5).Caption = "Salir"
     HabilitarTodo False, True
     Botones True, True, True, True, True, True
     BotonDetalles False, False, False, False
     BotonBuscar False, False
     nMovi = RsCv!Movimiento
     nCas = 0
  End If
Exit Sub

errHandler:
     ManejaErrores
     Recargar
End Sub
Private Sub Salir()
  If sw = False Then
     If CmdBotones(0).Enabled = True Then
        If MsgBox("El Comprobante no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
           Exit Sub
        End If
     Else
        If MsgBox("Cancela la Crearción del Comprobante ?", 20, "Atención") = vbNo Then
           Exit Sub
        End If
     End If
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
     CmdBotones(5).Caption = "Salir"
     HabilitarTodo False, False
     sw = True
     Limpiar
     LimpiarDetalles
     CrearRsDetalles
     Botones True, False, False, False, False, True
     BotonDetalles False, False, False, False
     BotonBuscar False, False
     If Rs_Ven.BOF = False Or Rs_Ven.EOF = False Then
        Rs_Ven.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Ven!ID, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True
     End If
  Else
     Unload Me
  End If
End Sub
Private Sub VerEstado()
  CmdBotones(0).Enabled = False
  CmdBotones(4).Enabled = False
  If Len(CmbComprobante.Text) <> 0 And Len(TxtCliente.Text) <> 0 _
     And Len(CmBVendedor.Text) <> 0 And Len(CmbCaja.Text) <> 0 And Rsd.RecordCount <> 0 Then
     CmdBotones(0).Enabled = True
     CmdBotones(5).Enabled = True
  Else
     If Len(CmbComprobante.Text) <> 0 And Len(TxtCliente.Text) <> 0 _
        And Len(CmBVendedor.Text) <> 0 And Len(CmbCaja.Text) <> 0 And Rsd.RecordCount <> 0 Then
        CmdBotones(4).Enabled = True
     End If
  End If
End Sub
Private Sub VerCabecera()
  CmdDetalle(0).Enabled = False
  If Len(CmbComprobante.Text) <> 0 And Len(TxtCliente.Text) <> 0 _
     And Len(CmBVendedor.Text) <> 0 And Len(CmbCaja.Text) <> 0 Then
     CmdDetalle(0).Enabled = True
  End If
End Sub
Private Sub GrabarTodo()
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cTipo As ClsComprobantesL
  Dim rReci As ADODB.Recordset, rCant As ADODB.Recordset, cStock As ClsComprobantesL
  ' declaracion para clientes que emiten pagare
  
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cTipo = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  Set rReci = New ADODB.Recordset
  Set rCant = New ADODB.Recordset
  
  Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
  Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
    
  ' Grabar Cabecera Venta
  
  cGrab.ActualizarMovimiento CmbComprobante.ItemData(CmbComprobante.ListIndex)
  
  RsCv.AddNew
 ' RsCv!Id = 0
  RsCv!Suc = nSucursal
  RsCv!tipo = 5
  RsCv!Movimiento = cNum.TraerUltimoMovimiento
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DTPFecha.Value
  RsCv!FechaIva = DTPFecha.Value
  RsCv!Numero = Right(LblNumero.Caption, 8)
  RsCv!sucursal = Left(LblNumero.Caption, 4)
  RsCv!Cliente = TxtCliente.Text
  RsCv!Lista = 1
  RsCv!Vendedor = CmBVendedor.ItemData(CmBVendedor.ListIndex)
  RsCv!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
  RsCv!CondVenta = 1
  RsCv!Neto = CCur(LblNeto.Caption)
  RsCv!Iva1 = 0
  RsCv!Iva2 = 0
  RsCv!Descuentos = 0
  RsCv!Financiacion = 0
  Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
         Case 11
              RsCv!Debe = CCur(LblTotal.Caption)
              RsCv!Haber = 0
         Case 12
              RsCv!Debe = 0
              RsCv!Haber = CCur(LblTotal.Caption)
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = ""
  RsCv!Anulado = 0
  RsCv.Update
  ' Detalle Venta ya esta Armado en el recordset Rsd
  ' Forma de Pagos,Valores,Aplicaciones,Cantidades,Caja
  If cStock.ActualizaCtaCte(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
     rReci.AddNew
    ' rReci!Id = 0
     rReci!fecha = DTPFecha.Value
     rReci!Cliente = TxtCliente.Text
     rReci!Venta = RsCv!Movimiento
     rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
     rReci!Numero = cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex))
     rReci!sucursal = cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex))
     Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
             Case 11
                  rReci!Cuota = 1000
                  rReci!Debe = CCur(LblTotal.Caption)
                  rReci!Haber = 0
             Case 12
                  rReci!Cuota = 1001
                  rReci!Debe = 0
                  rReci!Haber = CCur(LblTotal.Caption)
      End Select
      rReci!Anulado = 0
      rReci.Update
  End If
  ' Actualizo Stock
  If cStock.ActualizaStock(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        rCant.AddNew
        rCant!Suc = nSucursal
        rCant!Producto = Rsd!Producto
        rCant!Deposito = cStock.DepositoVenta
        Select Case cTipo.QueMoviStock(CmbComprobante.ItemData(CmbComprobante.ListIndex))
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
  
  
 cGrab.AgregarNotaCD RsCv, Rsd, rReci, rCant
  
 CrearRsDetalles
 Limpiar
 
 ArmaRsParaMovimiento
  
  ' Fin del proceso de facturacion fiscal
  
  Set cGrab = Nothing
  Set cNum = Nothing
  Set cTipo = Nothing
  Set cStock = Nothing
Exit Sub
 
Errores:
    ManejaErrores
End Sub
Private Sub Recargar()
  CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
  CmdBotones(0).Caption = "Nuevo"
  CmdBotones(5).Picture = LoadResPicture("Salir", 0)
  CmdBotones(5).Caption = "Salir"
  HabilitarTodo False, False
  sw = True
  Limpiar
  LimpiarDetalles
  CrearRsDetalles
  Botones True, False, False, False, False, True
  BotonDetalles False, False, False, False
  BotonBuscar False, False
  If Rs_Ven.BOF = False Or Rs_Ven.EOF Then
     Rs_Ven.MoveLast
     Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Ven!ID, True)
     LinkearTexto
     LLenarDetalles
     nCas = 0
     BotonDetalles False, False, False, False
     Botones True, True, True, True, True, True
  End If
End Sub

Private Sub Buscar()
  ComproBuscar.Show 1
  If ComproBuscar.nId <> 0 Then
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
     CmdBotones(5).Caption = "Salir"
     HabilitarTodo False, False
     sw = True
     Limpiar
     LimpiarDetalles
     CrearRsDetalles
     Botones True, False, False, False, False, True
     BotonDetalles False, False, False, False
     BotonBuscar False, False
     If Rs_Ven.BOF = False Or Rs_Ven.EOF Then
        Rs_Ven.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", ComproBuscar.nId, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True
     End If
  End If
End Sub
Private Sub Anular()
  Dim cBor As ClsComprobantesE
  
  Set cBor = New ClsComprobantesE
  If MsgBox("Anula el Comprobante Seleccionado ?", 20, "Atención") = 6 Then
     cBor.AnularComprobanteCantidades (Rs!Movimiento)
     LinkearTexto
     Botones True, True, True, True, True, True
  End If
End Sub


