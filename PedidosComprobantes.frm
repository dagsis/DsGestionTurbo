VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form PedidosComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   10425
   Begin VB.ComboBox CmbLista 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   990
      Width           =   3375
   End
   Begin VB.CommandButton CmdProveedor 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   5400
      TabIndex        =   31
      Top             =   570
      Width           =   735
   End
   Begin VB.TextBox TxtProveedor 
      Height          =   285
      Left            =   1245
      MaxLength       =   15
      TabIndex        =   29
      Text            =   "TxtProveedor"
      Top             =   570
      Width           =   795
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   690
      Index           =   5
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4965
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
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1635
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
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   885
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Anular"
      Height          =   690
      Index           =   1
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   885
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   690
      Index           =   0
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
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
      Height          =   4200
      Left            =   105
      TabIndex        =   6
      Top             =   1470
      Width           =   8010
      Begin VB.ComboBox CmbUnidad 
         Height          =   315
         Left            =   2670
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3075
         Width           =   2880
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3465
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3465
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3465
         Width           =   870
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   8
         Text            =   "TxtDetalle"
         Top             =   2700
         Width           =   3510
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3465
         Width           =   870
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   735
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "TxtPrecio"
         Top             =   3435
         Width           =   1050
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   825
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "TxtCa"
         Top             =   3075
         Width           =   930
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   5430
         TabIndex        =   13
         Top             =   2700
         Width           =   705
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   675
         MaxLength       =   15
         TabIndex        =   7
         Text            =   "TxtProducto"
         Top             =   2715
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2400
         Left            =   165
         TabIndex        =   16
         Top             =   240
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   4233
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
      Begin VB.Label Label6 
         Caption         =   "U.Medida :"
         Height          =   195
         Left            =   1845
         TabIndex        =   41
         Top             =   3075
         Width           =   915
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
         TabIndex        =   23
         Top             =   3855
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Total :"
         Height          =   225
         Left            =   6255
         TabIndex        =   22
         Top             =   3855
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
         Left            =   6825
         TabIndex        =   21
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Neto :"
         Height          =   255
         Left            =   6255
         TabIndex        =   20
         Top             =   2700
         Width           =   525
      End
      Begin VB.Label Label14 
         Caption         =   "Precio :"
         Height          =   195
         Left            =   75
         TabIndex        =   19
         Top             =   3480
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   3075
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   75
         TabIndex        =   17
         Top             =   2745
         Width           =   735
      End
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   690
      Index           =   4
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2385
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   690
      Index           =   6
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3135
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
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
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
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2385
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Modificar"
      Height          =   690
      Index           =   9
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1635
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   315
      Left            =   4485
      TabIndex        =   30
      Top             =   120
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52559873
      CurrentDate     =   36783
   End
   Begin VB.Label Label5 
      Caption         =   "L.de Precio :"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   990
      Width           =   1095
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
      TabIndex        =   37
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label LblProveedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblProveedor"
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
      Left            =   2160
      TabIndex        =   36
      Top             =   570
      Width           =   3105
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente :"
      Height          =   255
      Left            =   135
      TabIndex        =   35
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   3900
      TabIndex        =   34
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   240
      Left            =   5850
      TabIndex        =   33
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "PedidosComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean                     ' Suich Cabecera Pedido
Dim sw1 As Boolean                    ' Suich Detalles Pedidos
Dim sw2 As Boolean                    ' Suich Valores
Dim bProveedor As Boolean
Dim cImpor As Currency                ' Importe Total de Pedidos
Dim Rs_Sto As ADODB.Recordset         ' Cabecera Pedido Tabla Uno
Dim RsCv As ADODB.Recordset           ' Cabecera Virtual Pedido
Dim Rsd As ADODB.Recordset            ' Cabecera Vitual PedidoDetalle
Dim rs As ADODB.Recordset             ' Cabecera Pedido Tabla Todos
Dim Rs_Sis As ADODB.Recordset         ' Tabla Sistema, Ultima Pedido
Dim nTip As Integer                   ' Valor de tipo de Movimiento
Public nMovi As Long
Public nCas As Integer
Public rImp As ADODB.Recordset
Dim cRsl As ClsLectura
Dim nReg As Variant

Private Sub CmbComprobante_Click()
volver:
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler
  Set cRsl = New ClsLectura
  If sw = False Then
     Set cNum = New ClsComprobantesL
        If CmbComprobante.ListIndex <> -1 Then
        nNum = Format(cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
        nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
        LblNumero.Caption = nSuc & "-" & nNum
        End If
  End If
  Set cNum = Nothing
  VerCabecera
Exit Sub

errHandler:
   ManejaErrores
  
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
              ultimo
         Case 9
              Modificar
  End Select
End Sub
Private Sub CmdProveedor_Click()
  Dim RsC As ADODB.Recordset, cRiL As ClsClienteL, cRcL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRiL = New ClsClienteL
  Set cRcL = New ClsClienteL
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ClienteTraerUno", nDat, True)
     TxtProveedor.text = RsC!Cliente
     LblProveedor.Caption = RsC!RazonSocial
     CmbLista.text = cRsl.DatoCombo("ListadePrecio", "ListaPrecio", "Descripcion", cRcL.BuscarLista(RsC!Cliente))
     bProveedor = True
     CmdDetalle(0).SetFocus
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmdProductos_Click()
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL, nUni As Integer
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  
  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.text = RsP!Producto
     TxtDetalle.text = RsP!Descripcion
     TxtPrecio.text = Format(cRpl.TraerPrecio(CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.text), nCantDecimales)
     nUni = cRsl.TraerValorDeUnCampo("Productos", "UMCompra", "Producto='" & TxtProducto.text & "'")
     CmbUnidad.text = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & nUni)
     TxtCantidad.SetFocus
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
         ctl.text = ""
      End If
  Next ctl
  DTPFecha.Value = Date
  LblProveedor.Caption = ""
  LblNumero.ForeColor = &H80000012
  LblNumero.Caption = "0000-00000000"
  LblNeto.Caption = "0.00"
  LblTotal.Caption = "0.00"
  CmbComprobante.ListIndex = -1
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
  
  CmbComprobante.text = cCompr.DatoCombo("Comprobantes", "Id", "Descripcion", rs!Comprobante)
  DTPFecha.Value = rs!fecha
  nNum = Format(rs!Numero, "00000000")
  nSuc = Format(rs!Sucursal, "0000")
  LblNumero.ForeColor = &H80000012
  If cComp.ComprobanteAnulado(rs!Movimiento) = True Then
     LblNumero.ForeColor = &H40C0&
  End If
  LblNumero.Caption = nSuc & "-" & nNum
  nMovi = rs!Movimiento
  TxtProveedor.text = rs!Cliente
  LblProveedor.Caption = cRcL.BuscarNombreCliente(rs!Cliente)
  CmbLista.text = cCompr.DatoCombo("ListadePrecio", "ListaPrecio", "Descripcion", rs!Lista)
  LblNeto.Caption = Format(rs!Debe + rs!Haber, nCantDecimales)
  LblTotal.Caption = Format(rs!Debe + rs!Haber, nCantDecimales)
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

 Set rs = New ADODB.Recordset
 Set Rs_Sis = New ADODB.Recordset
 
 sw = True
 sw1 = True
 sw2 = True
 nCas = 0
 nLlama = 2
 
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
 ArmaRsParaMovimiento
 Me.Top = 0
 Me.Left = 0
 If CmbComprobante.ListCount = 0 Then
    MsgBox "Verifique que haya comprobantes de Stock o Depositos", vbCritical, "Atención"
    Botones False, False, False, False, False, True, False
 End If
 
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CargarCombos()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=10"
  cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmbUnidad, "UnidadMedida", "Unidad", "Descripcion", ""

  Set cRsl = New ClsLectura
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
  CmdProveedor.Enabled = bProv
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
  Grid1.HeadFont.size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Visible = False
      .Columns(1).Visible = False
      .Columns(2).Caption = "Producto"
      .Columns(2).Width = 1000
      .Columns(3).Caption = "Descripción"
      .Columns(3).Width = 5000
      .Columns(4).Caption = "Cant."
      .Columns(4).Width = 1000
      .Columns(4).Alignment = dbgRight
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(5).Visible = False
      .Columns(6).Visible = False
      .Columns(7).Visible = False
      .Columns(8).Visible = False
      .Columns(9).Visible = False
      .Columns(10).Caption = "P.Uni."
      .Columns(10).Width = 800
      .Columns(10).NumberFormat = nCantDecimales
      .Columns(10).Alignment = dbgRight
      .Columns(11).Caption = "Total"
      .Columns(11).Width = 1000
      .Columns(11).NumberFormat = nCantDecimales
      .Columns(11).Alignment = dbgRight
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
     CalcularTotales
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

  Dim cRx As ClsLectura
  
  Set cRx = New ClsLectura

  On Error GoTo errHandler

 ' Rsd!Id = 0
  Rsd!Producto = TxtProducto.text
  Rsd!Descripcion = TxtDetalle.text
  Rsd!Cantidad = TxtCantidad.text
  Rsd!Tasa = cRx.TraerValorDeUnCampo("ProductoImpuesto", "Impuesto", "Producto='" & TxtProducto.text & "'")
  Rsd!Deposito = 0
  Rsd!Medida = CmbUnidad.ItemData(CmbUnidad.ListIndex)
  Rsd!Impuesto = 0
  Rsd!PrecioUnitario = TxtPrecio.text
  Rsd!PrecioTotal = TxtCantidad.text * TxtPrecio.text
  Rsd.Update
  CabGrid
  CalcularTotales
  VerEstado
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub LimpiarDetalles()
 TxtProducto.text = ""
 TxtCantidad.text = "1.00"
 TxtDetalle.text = ""
 TxtPrecio.text = "0.00"
End Sub
Private Sub LinkearDetalles()
  Dim cRsPl As ClsProductoL, cRa As ClsLectura
  On Error GoTo errHandler
  
  Set cRsPl = New ClsProductoL
  Set cRa = New ClsLectura
  
  TxtProducto.text = Rsd!Producto
  TxtCantidad.text = Format(Rsd!Cantidad, nCantDecimales)
  TxtDetalle.text = Rsd!Descripcion
  CmbUnidad.text = cRa.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & Rsd!Medida)
  TxtPrecio.text = Format(Rsd!PrecioUnitario, nCantDecimales)
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub TxtCantidad_GotFocus()
  Dim cCan As ClsComprobantesL
  Set cCan = New ClsComprobantesL
  
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.text)
  
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCantidad_LostFocus()
  If TxtCantidad.text = "" Then TxtCantidad.text = 1
  TxtCantidad.text = Format(TxtCantidad.text, "0.00")
End Sub

Private Sub TxtProveedor_Change()
  VerEstado
  VerCabecera
End Sub
Private Sub TxtProveedor_GotFocus()
  TxtProveedor.SelStart = 0
  TxtProveedor.SelLength = Len(TxtProveedor.text)
End Sub

Private Sub TxtProveedor_LostFocus()
  Dim cRcL As ClsClienteL, cRiL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set cRcL = New ClsClienteL
  Set cRiL = New ClsClienteL
  LblProveedor.Caption = ""
  If TxtProveedor.text <> "" Then
     If cRcL.BuscarNombreCliente(TxtProveedor.text) <> "" Then
        LblProveedor.Caption = cRcL.BuscarNombreCliente(TxtProveedor.text)
        CmbLista.text = cRsl.DatoCombo("ListadePrecio", "ListaPrecio", "Descripcion", cRcL.BuscarLista(TxtProveedor.text))
        bProveedor = True
        VerCabecera
        CmdDetalle(0).SetFocus
     Else
        CmdProveedor.SetFocus
        bProveedor = False
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
     If Len(TxtDetalle.text) <> 0 Then
        CmdDetalle(0).Enabled = True
     End If
  Else
     CmdDetalle(2).Enabled = False
     If Len(TxtDetalle.text) <> 0 Then
        CmdDetalle(2).Enabled = True
     End If
  End If
End Sub
Private Sub TxtDetalle_LostFocus()
  If TxtDetalle.text <> "" Then
     If TxtCantidad.Enabled = True Then
        TxtCantidad.SetFocus
     End If
  Else
     If CmdProductos.Enabled = True Then
        CmdProductos.SetFocus
     End If
  End If
End Sub
Private Sub TxtPrecio_GotFocus()
  TxtPrecio.SelStart = 0
  TxtPrecio.SelLength = Len(TxtPrecio.text)
End Sub
Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimal(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtPrecio_LostFocus()
  If TxtPrecio.text = "" Then TxtPrecio.text = 0
  TxtPrecio.text = Format(TxtPrecio.text, nCantDecimales)
End Sub
Private Sub TxtProducto_LostFocus()
  Dim cRpl As ClsProductoL, sProdu As String, rPrd As ADODB.Recordset
  Dim cRx As ClsLectura, nUni As Integer
  
  On Error GoTo errHandler
  
  Set cRpl = New ClsProductoL
  Set cRx = New ClsLectura
  If TxtProducto.text <> "" Then
     If cRpl.EstaElCodigoCliente(TxtProveedor.text, TxtProducto.text) = False Then
        If cRpl.BuscarNombreProducto(TxtProducto.text) <> "" Then
           TxtDetalle.text = cRpl.BuscarNombreProducto(TxtProducto.text)
        End If
     Else
        Set rPrd = cRpl.TraerUnProductoCliente(TxtProveedor, TxtProducto.text)
        sProdu = rPrd!Producto
        TxtDetalle.text = cRpl.BuscarNombreProducto(sProdu)
     End If
     nUni = cRx.TraerValorDeUnCampo("Productos", "UMCompra", "Producto='" & TxtProducto.text & "'")
     CmbUnidad.text = cRx.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & nUni)
     TxtCantidad.SetFocus
  End If
  Set cRpl = Nothing
  Set cRx = Nothing
Exit Sub

errHandler:
   ManejaErrores

  Set cRpl = Nothing
End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtProducto.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  TxtPrecio.Enabled = bHabi
  CmbUnidad.Enabled = bHabi
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
  Set Rs_Sto = cComp.TraerTodos("CabComprobantes", "Id", 10)
  Botones True, False, False, False, False, True, False
  If Rs_Sto.BOF = False And Rs_Sto.EOF = False Then
     Rs_Sto.MoveLast
     Set rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
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
  Set rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
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
  Set rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
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
  Set rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
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
Private Sub ultimo()
  On Error GoTo errHandler
  Rs_Sto.MoveLast
  Set cRsl = New ClsLectura
  Set rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
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
  Set rAuxD = cRsl.TraerRS("DetallesTraerUnos", rs!Movimiento, True)
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
  CalcularTotales
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
     CmbLista.ListIndex = 0
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
        Set rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!Id, True)
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
        If Len(CmbComprobante.text) <> 0 And Len(TxtProveedor.text) <> 0 _
           And Rsd.RecordCount <> 0 And nCas = 2 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(9).Enabled = False
        CmdBotones(4).Enabled = False
        If Len(CmbComprobante.text) <> 0 And Len(TxtProveedor.text) <> 0 _
           And Rsd.RecordCount <> 0 And nCas = 2 Then
           CmdBotones(9).Enabled = True
        End If
     End If
  End If
End Sub
Private Sub VerCabecera()
  CmdDetalle(0).Enabled = False
  If Len(CmbComprobante.text) <> 0 Then
     CmdDetalle(0).Enabled = True
  End If
End Sub
Private Sub GrabarTodo()
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cTipo As ClsComprobantesL
  Dim cStock As ClsComprobantesL, rCant As ADODB.Recordset
  Dim i As Byte
  
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cTipo = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  Set cRsl = New ClsLectura
  Set rCant = New ADODB.Recordset
  
  Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
  
  ' Grabar Cabecera Stock
    
  RsCv.AddNew
 ' RsCv!Id = 0
  RsCv!tipo = cTipo.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  If CmdBotones(0).Caption = "Grabar" Then
     cGrab.ActualizarMovimiento CmbComprobante.ItemData(CmbComprobante.ListIndex)
     RsCv!Movimiento = cNum.TraerUltimoMovimiento
  Else
     RsCv!Movimiento = rs!Movimiento
  End If
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DTPFecha.Value
  RsCv!FechaIva = DTPFecha.Value
  RsCv!Numero = Right(LblNumero.Caption, 8)
  RsCv!Sucursal = Left(LblNumero.Caption, 4)
  RsCv!Cliente = TxtProveedor.text
  RsCv!Lista = CmbLista.ItemData(CmbLista.ListIndex)
  RsCv!Neto = CCur(LblNeto.Caption)
  RsCv!Debe = 0
  RsCv!Haber = 0
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = ""
  RsCv!Anulado = 0
  RsCv.Update
  ' Actualizo Stock
  
  Rsd.MoveFirst
  Do While Not Rsd.EOF
     rCant.AddNew
     rCant!Movimiento = RsCv!Movimiento
     rCant!Producto = Rsd!Producto
     rCant!Deposito = 0
     rCant!CantidadHaber = Rsd!Cantidad
     rCant!CantidadDebe = 0
     rCant!Anulado = 0
     rCant.Update
     Rsd.MoveNext
  Loop
  Rsd.MoveFirst
  
    If CmdBotones(0).Caption = "Grabar" Then
       cGrab.GrabarStock RsCv, Rsd, rCant, nTip
       CrearRsDetalles
       Limpiar
       ArmaRsParaMovimiento
    Else
       cGrab.ModificarStock RsCv, Rsd, rCant, nTip
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
        Set rs = cRsl.TraerRS("CabezaTraerUno", ComproBuscar.nId, True)
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
     cBor.AnularComprobanteCantidades (rs!Movimiento)
     LinkearTexto
     Botones True, True, True, True, True, True, True
  End If
End Sub

Private Sub Imprimir()
  Dim cClie As ClsClienteL
  Dim rCli As ADODB.Recordset
  
  Set rImp = New ADODB.Recordset
  Set cClie = New ClsClienteL
  
  rImp.Fields.Append "Fecha", adDate
  rImp.Fields.Append "Numero", adVarChar, 15, adFldIsNullable
  rImp.Fields.Append "Cliente", adVarChar, 15, adFldIsNullable
  rImp.Fields.Append "RazonSocial", adVarChar, 50, adFldIsNullable
  rImp.Fields.Append "Direccion", adVarChar, 50, adFldIsNullable
  rImp.Fields.Append "Iva", adVarChar, 30, adFldIsNullable
  rImp.Fields.Append "CPostal", adVarChar, 8, adFldIsNullable
  rImp.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
  rImp.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
  rImp.Fields.Append "Cuit", adVarChar, 13, adFldIsNullable
  rImp.Fields.Append "Codigo", adVarChar, 15, adFldIsNullable
  rImp.Fields.Append "Cantidad", adDouble
  rImp.Fields.Append "Descripcion", adVarChar, 80, adFldIsNullable
  rImp.Fields.Append "PUnitario", adCurrency
  rImp.Open
  Rsd.MoveFirst
  Set rCli = cClie.BuscarDatosCli(TxtProveedor.text)
  
  Do While Not Rsd.EOF
     rImp.AddNew
     rImp!fecha = DTPFecha.Value
     rImp!Numero = LblNumero.Caption
     rImp!Cliente = TxtProveedor.text
     rImp!RazonSocial = LblProveedor.Caption
     rImp!Direccion = rCli!Domicilio
     rImp!Iva = rCli![TipoIva.Descripcion]
     rImp!CPostal = rCli!CodigoPostal
     rImp!Localidad = rCli!Ciudad
     rImp!Provincia = rCli!Provincia
     rImp!Cuit = rCli!Cuit
     rImp!codigo = Rsd!Producto
     rImp!Cantidad = Rsd!Cantidad
     rImp!Descripcion = Rsd!Descripcion
     rImp!PUnitario = 0
     rImp.Update
     Rsd.MoveNext
  Loop
  nImpr = 26
  FrmImpresor.Show
End Sub

Private Sub CalcularTotales()
  Dim cTotal As Currency
  
  If Rsd.RecordCount <> 0 Then
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        cTotal = cTotal + Rsd!PrecioTotal
        Rsd.MoveNext
     Loop
  End If
  LblNeto.Caption = Format(cTotal, nCantDecimales)
  LblTotal.Caption = Format(cTotal, nCantDecimales)
End Sub

Private Sub Modificar()
  On Error GoTo errHandler

  If sw = True Then
     HabilitarTodo True, True
     HabilitarDetalles False
     BotonBuscar True, False
     CmbComprobante.Enabled = False
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


