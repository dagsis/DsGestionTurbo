VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ProduccionOrden 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Fabricación"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   10950
   Begin VB.ComboBox CmbHasta 
      Height          =   315
      Left            =   4905
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   525
      Width           =   3045
   End
   Begin VB.TextBox TxtObservacion 
      Height          =   285
      Left            =   1185
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "TxtObservacion"
      Top             =   960
      Width           =   7350
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   690
      Index           =   5
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4800
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   3
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1620
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   2
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   855
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Anular"
      Height          =   690
      Index           =   1
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   855
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   690
      Index           =   0
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   105
      Width           =   1005
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
      Height          =   3915
      Left            =   75
      TabIndex        =   16
      Top             =   1575
      Width           =   8580
      Begin VB.TextBox TxtCosto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "TxtCa"
         Top             =   3060
         Width           =   930
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3075
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   5865
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3075
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   7455
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3075
         Width           =   870
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   2010
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "TxtDetalle"
         Top             =   2700
         Width           =   5490
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   3990
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3075
         Width           =   870
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   885
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "TxtCa"
         Top             =   3060
         Width           =   930
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   7620
         TabIndex        =   17
         Top             =   2700
         Width           =   705
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   870
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "TxtProducto"
         Top             =   2700
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2280
         Left            =   165
         TabIndex        =   20
         Top             =   255
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   4022
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
      Begin VB.Label Label7 
         Caption         =   "P.Costo :"
         Height          =   255
         Left            =   2220
         TabIndex        =   35
         Top             =   3060
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   105
         TabIndex        =   22
         Top             =   3045
         Width           =   780
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   90
         TabIndex        =   21
         Top             =   2700
         Width           =   735
      End
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2820
   End
   Begin VB.ComboBox CmbDesde 
      Height          =   315
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   525
      Width           =   3045
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   690
      Index           =   4
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2370
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   690
      Index           =   6
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   7
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   105
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   8
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2370
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Modificar"
      Height          =   690
      Index           =   9
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1620
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   315
      Left            =   4905
      TabIndex        =   1
      Top             =   90
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   53346305
      CurrentDate     =   36783
   End
   Begin VB.Label Label5 
      Caption         =   "Hasta :"
      Height          =   225
      Left            =   4275
      TabIndex        =   34
      Top             =   540
      Width           =   660
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
      Left            =   7065
      TabIndex        =   33
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label4 
      Caption         =   "Observación :"
      Height          =   255
      Left            =   75
      TabIndex        =   32
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   4140
      TabIndex        =   31
      Top             =   90
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   240
      Left            =   6315
      TabIndex        =   30
      Top             =   90
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   75
      TabIndex        =   29
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label Label6 
      Caption         =   "Desde :"
      Height          =   225
      Left            =   75
      TabIndex        =   28
      Top             =   540
      Width           =   660
   End
End
Attribute VB_Name = "ProduccionOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean                     ' Suich Cabecera Stock
Dim sw1 As Boolean                    ' Suich Detalles Stock
Dim sw2 As Boolean                    ' Suich Valores
Dim bProveedor As Boolean
Dim cImpor As Currency                ' Importe Total de Stock
Dim Rs_Sto As ADODB.Recordset         ' Cabecera Stock Tabla Uno
Dim RsCv As ADODB.Recordset           ' Cabecera Virtual Stock
Dim Rsd As ADODB.Recordset            ' Cabecera Vitual StockDetalle
Dim Rs As ADODB.Recordset             ' Cabecera Stock Tabla Todos
Dim Rs_Sis As ADODB.Recordset         ' Tabla Sistema, Ultima Stock
Dim nTip As Integer                  ' Valor de tipo de Movimiento
Public nMovi As Long
Public nCas As Integer
Public rImp As ADODB.Recordset
Dim cRsl As ClsLectura
Dim bRemi As Boolean
Dim rRem As ADODB.Recordset
Dim nReg As Variant

Private Sub CmbDesde_Click()
  Dim cRsl As ClsLectura
  VerEstado
'  VerCabecera
  Set cRsl = New ClsLectura
  CmbHasta.ListIndex = -1
  If CmbDesde.ListIndex <> -1 Then
     cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", "Deposito<>" & CmbDesde.ItemData(CmbDesde.ListIndex)
  End If
  Set cRsl = Nothing
  If CmbHasta.ListCount <> 0 Then
     CmbHasta.ListIndex = 0
  End If
End Sub
Private Sub CmbComprobante_Click()
  Dim cRsl As ClsLectura
volver:
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler
  Set cRsl = New ClsLectura
  If sw = False Then
     Set cNum = New ClsComprobantesL
         If CmbComprobante.ListIndex <> -1 Then
            nNum = Format(cNum.TraerUltimoNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
            nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
            LblNumero.Caption = nSuc & "-" & nNum
            HabiCombo True, True
            cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", ""
            CmbHasta.Clear
            Set cRsl = Nothing
            CrearRsDetalles
            LimpiarDetalles
            If CmbDesde.ListCount <> 0 Then
               CmbDesde.ListIndex = 0
            End If
         End If
  End If
  Set cNum = Nothing
  Set cRsl = Nothing
  VerCabecera
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub HabiCombo(bDes As Boolean, bHas As Boolean)
  CmbDesde.Enabled = bDes
  CmbHasta.Enabled = bHas
End Sub
Private Sub CmbHasta_Click()
  VerEstado
'  VerCabecera
End Sub

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Anular
         Case 2
              Anterior
         Case 3
              Siguiente
         Case 4
              Buscar
         Case 5
              Salir
         Case 6
              Imprimir
         Case 7
              Primero
         Case 8
              Ultimo
         Case 9
              Modificar
  End Select
End Sub
Private Sub CmdProveedor_Click()
'  Dim RsC As ADODB.Recordset, cRiL As ClsClienteL, cRcl As ClsClienteL
'  On Error GoTo errHandler
'
'  Set cRiL = New ClsClienteL
'  Set cRcl = New ClsClienteL
'  Set RsC = New ADODB.Recordset
'
'  Set cRsl = New ClsLectura
'  nBuscar = 1
'  FrmBuscar.Show 1
'  If nDat <> 0 Then
'     Set RsC = cRsl.TraerRS("ClienteTraerUno", nDat, True)
'     TxtProveedor.Text = RsC!Cliente
'     LblProveedor.Caption = RsC!RazonSocial
'     bProveedor = True
'     VerPedidos
'     CmdDetalle(0).SetFocus
'     RsC.Close
'  End If
'  Set cRsl = Nothing
'Exit Sub
'
'errHandler:
'   ManejaErrores
  
End Sub
Private Sub VerPedidos()
'  On Error GoTo errHandler
'
'  Dim cRem As ClsComprobantesL, rRem As ADODB.Recordset
'
'  Set cRem = New ClsComprobantesL
'  Set rRem = New ADODB.Recordset
'  bRemi = False
'  Set rRem = cRem.TraerClientesPedidos(TxtProveedor.Text)
'  If rRem.RecordCount <> 0 Then
'     PedidosPendientes.Show 1
'     CrearRsDetalles
'     LlenarDetallesPedidos
'     Botones True, False, False, False, False, True, False
'  End If
'  VerCabecera
'Exit Sub
'
'errHandler:
'   MsgBox Err.Description, "VerPedidos"
  
End Sub

Private Sub LlenarDetallesPedidos()
'  On Error GoTo errHandler
'
'  Dim cRem As ClsComprobantesL, cPre As ClsProductoL
'  Dim rAux As ADODB.Recordset, nCanti As Single, cimp As ClsProductoL
'  Dim cLista As ClsComprobantesL, bDiscr As Boolean, cCli As ClsClienteL, cRpl As ClsProductoL
'
'  Set rAux = New ADODB.Recordset
'  Set cPre = New ClsProductoL
'  Set cimp = New ClsProductoL
'  Set cLista = New ClsComprobantesL
'  Set cCli = New ClsClienteL
'  Set cRpl = New ClsProductoL
'
'  Set rRem = PedidosPendientes.rPedidos
'
'  bDiscr = cCli.ClienteDiscrimina(TxtProveedor)
'
'  If Not rRem Is Nothing Then
'     If rRem.RecordCount <> 0 Then
'        rRem.MoveFirst
'        Set cRem = New ClsComprobantesL
'        sw1 = False
'        Do While Not rRem.EOF
'           Set rAux = cRem.TraerDetallesMovimiento(rRem!Movimiento)
'           Do While Not rAux.EOF
'              Rsd.AddNew
'              Rsd!Id = 0
'              Rsd!Producto = rAux!Producto
'              Rsd!Descripcion = rRem!Descripcion + " " & rAux!Descripcion
'              Rsd!Cantidad = rAux!Cantidad
'              Rsd!PrecioUnitario = rAux!PrecioUnitario
'              Rsd!PrecioTotal = nCanti * rAux!PrecioUnitario
'              Rsd.Update
'              rAux.MoveNext
'           Loop
'           rRem.MoveNext
'           bRemi = True
'           CmdBotones(4).Enabled = True
'           Grid1.Enabled = True
'        Loop
'        CabGrid
'        sw1 = True
'    End If
'  End If
'Exit Sub
'
'errHandler:
'   MsgBox Err.Description, "LlenarDetallesPedidos"
  
End Sub

Private Sub CmdProductos_Click()
  Dim RsP As ADODB.Recordset
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nBuscar = 3
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.Text = RsP!Producto
     TxtDetalle.Text = RsP!Descripcion
     TxtCosto.Text = Format(cRsl.TraerValorDeUnCampo("Productos", "PrecioCompra", "Producto='" & TxtProducto.Text & "'"), nCantDecimales)
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub DtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
        Exit Sub
     End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
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
  LblNumero.ForeColor = &H80000012
  LblNumero.Caption = "0000-00000000"
  TxtCantidad.Text = "0.00"
  TxtCosto.Text = "0.00"
  CmbComprobante.ListIndex = -1
  CmbDesde.ListIndex = -1
  CmbHasta.ListIndex = -1
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
  Dim cRcl As ClsClienteL, cRiL As ClsClienteL, nSuc As String * 4, nNum As String * 8
  Dim cComp As ClsComprobantesL
  Dim cCompr As ClsLectura, cLista As ClsLectura, cVend As ClsLectura, cCaja As ClsLectura
  On Error GoTo errHandler
  
  Set cCompr = New ClsLectura
  Set cLista = New ClsLectura
  Set cVend = New ClsLectura
  Set cCaja = New ClsLectura
  Set cRcl = New ClsClienteL
  Set cRiL = New ClsClienteL
  Set cComp = New ClsComprobantesL
  
  CmbComprobante.Text = cCompr.DatoCombo("Comprobantes", "Id", "Descripcion", Rs!Comprobante)
  If Rs!Desde <> 0 Then
     CmbDesde.Text = cCompr.DatoCombo("Depositos", "Deposito", "Descripcion", Rs!Desde)
  Else
     CmbDesde.ListIndex = -1
  End If
  If Rs!Hasta <> 0 Then
     CmbHasta.Text = cCompr.DatoCombo("Depositos", "Deposito", "Descripcion", Rs!Hasta)
  Else
     CmbHasta.ListIndex = -1
  End If
  DTPFecha.Value = Rs!Fecha
  nNum = Format(Rs!Numero, "00000000")
  nSuc = Format(Rs!Sucursal, "0000")
  LblNumero.ForeColor = &H80000012
  If cComp.ComprobanteAnulado(Rs!Movimiento) = True Then
     LblNumero.ForeColor = &H40C0&
  End If
  LblNumero.Caption = nSuc & "-" & nNum
  nMovi = Rs!Movimiento
  TxtObservacion.Text = Rs!Motivo
'  TxtProveedor.Text = Rs!Cliente
'  LblProveedor.Caption = cRcl.BuscarNombreCliente(Rs!Cliente)
 ' TxtNumero.Text = Format(Rs!factura, "000000")
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
 nLlama = 2
 nTip = 5
 Botones True, False, False, False, False, True, False
 BotonDetalles False, False, False, False
 BotonBuscar False, False

 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(4).Picture = LoadResPicture("Buscar", 0)
 CmdBotones(5).Picture = LoadResPicture("Salir", 0)
 CmdBotones(6).Picture = LoadResPicture("Imprimir", 0)
 CmdBotones(9).Picture = LoadResPicture("Modificar", 0)
 
 
 CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
 CmdDetalle(1).Picture = LoadResPicture("Borrar", 0)
 CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
 CmdDetalle(3).Picture = LoadResPicture("Cancelar", 0)
 
 CargarCombos
 CrearRsDetalles
 Limpiar
 HabilitarTodo False, False
 HabilitarDetalles False
 ArmaRSparaMovimiento
 Me.Top = 0
 Me.Left = 0
 If CmbComprobante.ListCount = 0 Or CmbDesde.ListCount = 0 Or CmbHasta.ListCount = 0 Then
    MsgBox "Verifique que haya comprobantes de Stock o Depositos", vbCritical, "Atención"
    Botones False, False, False, False, False, True, False
 End If
 
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=11"
  cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", ""
  cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", ""
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bMod As Boolean, bBus As Boolean, _
  bVal As Boolean, bSal As Boolean, bModi As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bBorr
  If LblNumero.ForeColor = &H40C0& Then
     CmdBotones(1).Enabled = False
  End If
  CmdBotones(2).Enabled = bBus
  CmdBotones(3).Enabled = bBus
  CmdBotones(4).Enabled = bVal
  CmdBotones(5).Enabled = bSal
  CmdBotones(6).Enabled = bBus
  CmdBotones(7).Enabled = bBus
  CmdBotones(8).Enabled = bBus
  CmdBotones(9).Enabled = bModi
End Sub
Private Sub BotonDetalles(bAgr As Boolean, bBorr As Boolean, bMod As Boolean, bCan As Boolean)
  CmdDetalle(0).Enabled = bAgr
  CmdDetalle(1).Enabled = bBorr
  CmdDetalle(2).Enabled = bMod
  CmdDetalle(3).Enabled = bCan
End Sub
Private Sub BotonBuscar(bProv As Boolean, bPro As Boolean)
  CmdProductos.Enabled = bPro
End Sub
Private Sub Form_Resize()
  On Error Resume Next
  CmdBotones(0).SetFocus
End Sub
Private Sub CrearRsStock()
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
      .Columns(3).Width = 5000
      .Columns(4).Caption = "Cant."
      .Columns(4).Width = 800
      .Columns(4).Alignment = dbgRight
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(5).Visible = False
      .Columns(6).Visible = False
      .Columns(7).Visible = False
      .Columns(8).Visible = False
      .Columns(9).Visible = False
      .Columns(10).Visible = False
      .Columns(11).Visible = False
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
     sw1 = True
     HabilitarDetalles False
     HabilitarTodo True, True
     If nTip <> 5 Then
        CmbDesde.Enabled = False
     End If
     If CmdBotones(0).Caption = "Grabar" Then
        Botones True, False, False, False, False, True, False
     Else
        Botones False, False, False, False, False, True, True
     End If
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
        Botones False, False, False, False, False, True, False
     End If
  End If
Exit Sub
  
errHandler:
   ManejaErrores

End Sub
Private Sub ModificarRsD()
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     nReg = Rsd.Bookmark
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
     HabilitarTodo True, True
     HabilitarDetalles False
     Rsd.Bookmark = nReg
     Botones False, False, False, False, False, True, False
     If Rsd.RecordCount <> 0 Then
        Botones False, False, False, False, False, True, True
     End If
     sw1 = True
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

  Rsd!Id = 0
  Rsd!Producto = TxtProducto.Text
  Rsd!Descripcion = TxtDetalle.Text
  Rsd!Cantidad = TxtCantidad.Text
  Rsd!Tasa = 0
  Rsd!Deposito = 0
  Rsd!Medida = 0
  Rsd!Impuesto = 0
  Rsd!PrecioUnitario = TxtCosto.Text
  Rsd!PrecioTotal = TxtCosto.Text
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
End Sub
Private Sub LinkearDetalles()
  Dim cRsPl As ClsProductoL, cRsl As ClsLectura
  On Error GoTo errHandler
  
  Set cRsPl = New ClsProductoL
  Set cRsl = New ClsLectura
  
  TxtProducto.Text = Rsd!Producto
  TxtCantidad.Text = Format(Rsd!Cantidad, nCantDecimales)
  TxtDetalle.Text = Rsd!Descripcion
  TxtCosto.Text = Format(Rsd!PrecioTotal, nCantDecimales)
Exit Sub

errHandler:
   ManejaErrores
  
End Sub


Private Sub Text1_Change()

End Sub

Private Sub TxtCantidad_GotFocus()
  Dim cCan As ClsComprobantesL
  
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.Text)
  
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimal(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCantidad_LostFocus()
  If TxtCantidad.Text = "" Then TxtCantidad.Text = 1
  TxtCantidad.Text = Format(TxtCantidad.Text, "0.00")
End Sub

Private Sub TxtProveedor_Change()
  VerEstado
  VerCabecera
End Sub

Private Sub TxtCosto_GotFocus()
  TxtCosto.SelStart = 0
  TxtCosto.SelLength = Len(TxtCosto.Text)
End Sub

Private Sub TxtCosto_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimal(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCosto_LostFocus()
  If TxtCosto.Text = "" Then TxtCosto.Text = 0
  TxtCosto.Text = Format(TxtCosto.Text, nCantDecimales)
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
     If TxtCantidad.Enabled = True Then
'        CmbTalle.SetFocus
     End If
  Else
     If CmdProductos.Enabled = True Then
        CmdProductos.SetFocus
     End If
  End If
End Sub
Private Sub TxtProducto_LostFocus()
  Dim cRpl As ClsProductoL, cRx As ClsLectura
  On Error GoTo errHandler
  
  Set cRpl = New ClsProductoL
  Set cRx = New ClsLectura
  
  If TxtProducto.Text <> "" Then
     If cRpl.BuscarNombreProducto(TxtProducto.Text) <> "" Then
        TxtDetalle.Text = cRpl.BuscarNombreProducto(TxtProducto.Text)
        TxtCosto.Text = Format(cRx.TraerValorDeUnCampo("Productos", "PrecioCompra", "Producto='" & TxtProducto.Text & "'"), nCantDecimales)
     End If
  End If
  Set cRpl = Nothing
  Set cRx = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtProducto.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  TxtCosto.Enabled = bHabi
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
Private Sub ArmaRSparaMovimiento()
  Dim cComp As ClsComprobantesL
  On Error GoTo errHandler

  Set cComp = New ClsComprobantesL
  Set Rs_Sto = cComp.TraerTodos("CabComprobantes", "Id", 11)
  Botones True, False, False, False, False, True, False
  If Rs_Sto.BOF = False And Rs_Sto.EOF = False Then
     Rs_Sto.MoveLast
     Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
     LinkearTexto
     LLenarDetalles
     nCas = 0
     Botones True, True, True, True, True, True, True
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
  Rs_Sto.MoveNext
  If Rs_Sto.EOF Then
     Rs_Sto.MoveLast
     Beep
  End If
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Anterior()
  On Error GoTo errHandler
  Rs_Sto.MovePrevious
  If Rs_Sto.BOF Then
     Rs_Sto.MoveFirst
     Beep
  End If
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Primero()
  On Error GoTo errHandler
  Rs_Sto.MoveFirst
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Ultimo()
  On Error GoTo errHandler
  Rs_Sto.MoveLast
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True
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
     Botones False, False, False, False, False, True, False
     BotonDetalles False, False, False, False
     BotonBuscar True, False
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(5).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(5).Caption = "Cancelar"
     sw = False
     CrearRsStock
     CrearRsDetalles
     CmbComprobante.SetFocus
     nMovi = 0
     nCas = 1
     Set cRsl = New ClsLectura
     CmbComprobante.ListIndex = 0
  Else
     sw = True
     Botones False, False, False, False, False, False, False
     GrabarTodo
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
     CmdBotones(5).Caption = "Salir"
     HabilitarTodo False, True
     Botones True, True, True, True, True, True, True
     BotonDetalles False, False, False, False
     BotonBuscar False, False
     nMovi = RsCv!Movimiento
     nCas = 0
  End If
Exit Sub

errHandler:
     ManejaErrores
End Sub
Private Sub Salir()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
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
     Else
        If CmdBotones(9).Enabled = True Then
           If MsgBox("El Comprobante no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        Else
           If MsgBox("Cancela la Modificación del Comprobante ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        End If
        CmdBotones(9).Picture = LoadResPicture("Modificar", 0)
        CmdBotones(9).Caption = "Modificar"
        CmdBotones(5).Picture = LoadResPicture("Salir", 0)
        CmdBotones(5).Caption = "Salir"
     End If
     HabilitarTodo False, False
     sw = True
     nTip = 0
     Limpiar
     LimpiarDetalles
     CrearRsDetalles
     Botones True, False, False, False, False, True, False
     BotonDetalles False, False, False, False
     BotonBuscar False, False
     If Rs_Sto.BOF = False Or Rs_Sto.EOF = False Then
        Rs_Sto.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True, True
     End If
  Else
     Unload Me
  End If
End Sub
Private Sub VerEstado()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        CmdBotones(4).Enabled = False
        If Len(CmbComprobante.Text) <> 0 And Len(CmbHasta.Text) <> 0 _
           And Rsd.RecordCount <> 0 And nCas = 2 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(9).Enabled = False
        CmdBotones(4).Enabled = False
        If Len(CmbComprobante.Text) <> 0 And Len(CmbHasta.Text) <> 0 _
           And Rsd.RecordCount <> 0 And nCas = 2 Then
           CmdBotones(9).Enabled = True
        End If
     End If
  End If
End Sub
Private Sub VerCabecera()
  CmdDetalle(0).Enabled = False
  If Len(CmbComprobante.Text) <> 0 Or CmbHasta.Enabled = True Then
     CmdDetalle(0).Enabled = True
  End If
End Sub
Private Sub GrabarTodo()
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cTipo As ClsComprobantesL
  Dim cStock As ClsComprobantesL, rCant As ADODB.Recordset, cMat As ADODB.Recordset
  Dim I As Byte
  
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cTipo = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  Set cRsl = New ClsLectura
  Set rCant = New ADODB.Recordset
  Set cMat = New ADODB.Recordset
  
  Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
  
  ' Grabar Cabecera Stock
  
  RsCv.AddNew
  RsCv!Id = 0
  RsCv!Tipo = 11
  If CmdBotones(0).Caption = "Grabar" Then
     RsCv!Movimiento = cNum.TraerUltimoMovimiento
  Else
     RsCv!Movimiento = Rs!Movimiento
  End If
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!Fecha = DTPFecha.Value
  RsCv!Numero = Right(LblNumero.Caption, 8)
  RsCv!Sucursal = Left(LblNumero.Caption, 4)
  RsCv!Cliente = 0
  RsCv!Neto = 0
  RsCv!Debe = 0
  RsCv!Haber = 0
  If CmbDesde.ListIndex <> -1 Then
     RsCv!Desde = CmbDesde.ItemData(CmbDesde.ListIndex)
  Else
     RsCv!Desde = 0
  End If
  If CmbHasta.ListIndex <> -1 Then
     RsCv!Hasta = CmbHasta.ItemData(CmbHasta.ListIndex)
  Else
    RsCv!Hasta = 0
  End If
  RsCv!factura = 0
  RsCv!Motivo = TxtObservacion.Text
  RsCv.Update
  ' Actualizo Stock
  
  Rsd.MoveFirst
  Do While Not Rsd.EOF
     Set cMat = cNum.TraerMateriaPrima(Rsd!Producto)
     Do While Not cMat.EOF
        rCant.AddNew
        rCant!Movimiento = RsCv!Movimiento
        rCant!Producto = cMat!Producto
        If cRsl.TraerValorDeUnCampo("Productos", "Tipo", "Producto='" & cMat!Producto & "'") <> 1 Then
           rCant!Deposito = CmbDesde.ItemData(CmbDesde.ListIndex)
        Else
           rCant!Deposito = cRsl.TraerValorDeUnCampo("Productos", "Deposito", "Producto='" & cMat!Producto & "'")
        End If
        rCant!CantidadDebe = 0
        rCant!CantidadHaber = cMat!Cantidad * Rsd!Cantidad
        rCant.Update
        cMat.MoveNext
     Loop
     Rsd.MoveNext
  Loop
  
  
  Rsd.MoveFirst
  Do While Not Rsd.EOF
     rCant.AddNew
     rCant!Movimiento = RsCv!Movimiento
     rCant!Producto = Rsd!Producto
     rCant!Deposito = CmbHasta.ItemData(CmbHasta.ListIndex)
     rCant!CantidadDebe = Rsd!Cantidad
     rCant!CantidadHaber = 0
     rCant.Update
     Rsd.MoveNext
  Loop
  Rsd.MoveFirst
  
  ' Grabo Todo en los Store Procedure
'  If RsCv!Movimiento >= 80 Then
'          MsgBox "Periodo de Demostración Finalizado", vbCritical, "Atención"
'          Exit Sub
'  Else
    If CmdBotones(0).Caption = "Grabar" Then
       cGrab.GrabarStock RsCv, Rsd, rCant, 6, False, rRem
       CrearRsDetalles
       Limpiar
       ArmaRSparaMovimiento
    Else
       cGrab.ModificarStock RsCv, Rsd, rCant, 6, False
    End If
 ' End If
  nTip = 0

  Set cGrab = Nothing
  Set cNum = Nothing
  Set cTipo = Nothing
  Set cStock = Nothing

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
     Botones True, False, False, False, False, True, False
     BotonDetalles False, False, False, False
     BotonBuscar False, False
     If Rs_Sto.BOF = False Or Rs_Sto.EOF Then
        Rs_Sto.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", ComproBuscar.nId, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True, True
     End If
  End If
End Sub

Private Sub Anular()
  Dim cBor As ClsComprobantesE
  
  Set cBor = New ClsComprobantesE
  If MsgBox("Anula el Comprobante Seleccionado ?", 20, "Atención") = 6 Then
     cBor.AnularComprobanteCantidades (Rs!Movimiento)
     LinkearTexto
     Botones True, True, True, True, True, True, True
  End If
End Sub

Private Sub Imprimir()
'  Dim cClie As ClsClienteL
'  Dim rCli As ADODB.Recordset
'
'  Set rImp = New ADODB.Recordset
'  Set cClie = New ClsClienteL
'
'  rImp.Fields.Append "Fecha", adDate
'  rImp.Fields.Append "Numero", adVarChar, 15, adFldIsNullable
'  rImp.Fields.Append "Cliente", adVarChar, 15, adFldIsNullable
'  rImp.Fields.Append "RazonSocial", adVarChar, 50, adFldIsNullable
'  rImp.Fields.Append "Direccion", adVarChar, 50, adFldIsNullable
'  rImp.Fields.Append "Iva", adVarChar, 30, adFldIsNullable
'  rImp.Fields.Append "CPostal", adVarChar, 8, adFldIsNullable
'  rImp.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
'  rImp.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
'  rImp.Fields.Append "Cuit", adVarChar, 13, adFldIsNullable
'  rImp.Fields.Append "Codigo", adVarChar, 15, adFldIsNullable
'  rImp.Fields.Append "Cantidad", adDouble
'  rImp.Fields.Append "Descripcion", adVarChar, 80, adFldIsNullable
'  rImp.Fields.Append "Corte", adInteger, , adFldIsNullable
'  rImp.Fields.Append "PUnitario", adCurrency
'  rImp.Open
'  Rsd.MoveFirst
'  Set rCli = cClie.BuscarDatosCli(TxtProveedor.Text)
'
'  Do While Not Rsd.EOF
'     rImp.AddNew
'     rImp!Fecha = DTPFecha.Value
'     rImp!Numero = LblNumero.Caption
'     rImp!Cliente = TxtProveedor.Text
'     rImp!RazonSocial = LblProveedor.Caption
'     If rCli.RecordCount <> 0 Then
'        rImp!Direccion = rCli!Domicilio
'        rImp!Iva = rCli![TipoIva.Descripcion]
'        rImp!CPostal = rCli!CodigoPostal
'        rImp!Localidad = rCli![Localidades.Descripcion]
'        rImp!Provincia = rCli![Provincias.Descripcion]
'        rImp!Cuit = rCli!Cuit
'     End If
'     rImp!Codigo = Rsd!Producto
'     rImp!Cantidad = Rsd!Cantidad
'     rImp!Descripcion = Rsd!Descripcion
'     rImp!Corte = Rsd!Corte
'     rImp!PUnitario = 0
'     rImp.Update
'     Rsd.MoveNext
'  Loop
'  nImpr = 26
'  FrmImpresor.Show
End Sub
Private Sub Modificar()
  On Error GoTo errHandler

  If sw = True Then
     HabilitarTodo True, True
     HabilitarDetalles False
     BotonBuscar True, False
     CmbComprobante.Enabled = False
     CmbDesde.Enabled = False
     CmbHasta.Enabled = False
     DTPFecha.Enabled = False
     CrearRsStock
     Rsd.MoveLast
     Grid1.Bookmark = Rsd.Bookmark
     LinkearDetalles
     Botones False, False, True, False, False, True, True
     BotonDetalles True, True, True, False
     CmdBotones(9).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(9).Caption = "Grabar"
     CmdBotones(5).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(5).Caption = "Cancelar"
     sw = False
     sw1 = True
  Else
     Botones False, False, False, False, False, False, False
     GrabarTodo
     sw = True
     CmdBotones(9).Picture = LoadResPicture("Modificar", 0)
     CmdBotones(9).Caption = "Modificar"
     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
     CmdBotones(5).Caption = "Salir"
     HabilitarTodo False, True
     BotonBuscar False, False
     Botones True, True, True, True, True, True, True
     BotonDetalles False, False, False, False
     CmdBotones(0).SetFocus
  End If
Exit Sub

errHandler:
  MsgBox Err.Description, vbCritical, "Modificar"

End Sub



