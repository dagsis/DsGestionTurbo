VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasRemitos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos de Compras"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   10995
   Begin VB.TextBox TxtMotivo 
      Height          =   285
      Left            =   1185
      MaxLength       =   80
      TabIndex        =   34
      Text            =   "TxtMotivo"
      Top             =   1380
      Width           =   7485
   End
   Begin VB.CommandButton CmdOrden 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7770
      TabIndex        =   33
      Top             =   990
      Width           =   540
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Modificar"
      Height          =   690
      Index           =   9
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1620
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
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2370
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
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   120
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   690
      Index           =   4
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2370
      Width           =   1005
   End
   Begin VB.ComboBox CmbDesde 
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   540
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   90
      Width           =   2535
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
      Height          =   4410
      Left            =   45
      TabIndex        =   10
      Top             =   1710
      Width           =   8640
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   930
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "TxtImporte"
         Top             =   3930
         Width           =   930
      End
      Begin VB.ComboBox CmbImpuesto 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3255
         Width           =   2880
      End
      Begin VB.ComboBox CmbUnidad 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3255
         Width           =   2880
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   960
         MaxLength       =   25
         TabIndex        =   18
         Text            =   "TxtProducto"
         Top             =   2940
         Width           =   1095
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   6990
         TabIndex        =   17
         Top             =   2940
         Width           =   705
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   945
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "TxtCa"
         Top             =   3600
         Width           =   930
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3600
         Width           =   870
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   2115
         MaxLength       =   100
         TabIndex        =   14
         Text            =   "TxtDetalle"
         Top             =   2940
         Width           =   4800
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3600
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3600
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3600
         Width           =   870
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2640
         Left            =   165
         TabIndex        =   22
         Top             =   210
         Width           =   8370
         _ExtentX        =   14764
         _ExtentY        =   4657
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
      Begin VB.Label Label9 
         Caption         =   "Importe :"
         Height          =   255
         Left            =   165
         TabIndex        =   26
         Top             =   3945
         Width           =   780
      End
      Begin VB.Label Label14 
         Caption         =   "U.Medida :"
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   165
         TabIndex        =   24
         Top             =   2940
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   165
         TabIndex        =   23
         Top             =   3600
         Width           =   780
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   690
      Index           =   0
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Anular"
      Height          =   690
      Index           =   1
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   870
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
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   870
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
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1620
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   690
      Index           =   5
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5415
      Width           =   1005
   End
   Begin VB.TextBox TxtProveedor 
      Height          =   285
      Left            =   1185
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "TxtProveedor"
      Top             =   990
      Width           =   795
   End
   Begin VB.CommandButton CmdProveedor 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   5985
      TabIndex        =   3
      Top             =   990
      Width           =   735
   End
   Begin VB.ComboBox CmbHasta 
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   555
      Width           =   3045
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   7740
      TabIndex        =   1
      Text            =   "00000000"
      Top             =   90
      Width           =   855
   End
   Begin VB.TextBox TxtSucursal 
      Height          =   315
      Left            =   6945
      TabIndex        =   0
      Text            =   "0000"
      Top             =   90
      Width           =   495
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   315
      Left            =   4710
      TabIndex        =   35
      Top             =   90
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   708182017
      CurrentDate     =   36783
   End
   Begin VB.Label Label8 
      Caption         =   "Observación :"
      Height          =   165
      Left            =   90
      TabIndex        =   45
      Top             =   1395
      Width           =   990
   End
   Begin VB.Label Label7 
      Caption         =   "Factura :"
      Height          =   180
      Left            =   6960
      TabIndex        =   44
      Top             =   990
      Width           =   705
   End
   Begin VB.Label Label5 
      Caption         =   "Desde :"
      Height          =   195
      Left            =   4425
      TabIndex        =   43
      Top             =   570
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   90
      TabIndex        =   42
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   4125
      TabIndex        =   41
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Proveedor :"
      Height          =   255
      Left            =   90
      TabIndex        =   40
      Top             =   990
      Width           =   900
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
      Left            =   2025
      TabIndex        =   39
      Top             =   990
      Width           =   3885
   End
   Begin VB.Label Label6 
      Caption         =   "Depósito :"
      Height          =   225
      Left            =   90
      TabIndex        =   38
      Top             =   540
      Width           =   885
   End
   Begin VB.Label Label17 
      Caption         =   "--"
      Height          =   195
      Left            =   7530
      TabIndex        =   37
      Top             =   135
      Width           =   150
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   240
      Left            =   6060
      TabIndex        =   36
      Top             =   90
      Width           =   765
   End
End
Attribute VB_Name = "ComprasRemitos"
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
Dim nTip As Integer                   ' Valor de tipo de Movimiento
Public nMovi As Long, nClient As String
Public nCas As Integer
Public rImp As ADODB.Recordset
Dim cRsl As ClsLectura
Dim bRemi As Boolean
Dim rRem As ADODB.Recordset
Dim nReg As Variant
Dim bStock As Boolean, nFactura As Long

Private Sub CmbDesde_Click()
  VerEstado
'  VerCabecera
  Set cRsl = New ClsLectura
  CmbHasta.ListIndex = -1
  If nTip = 5 Then
     If CmbDesde.ListIndex <> -1 Then
        cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", "Deposito<>" & CmbDesde.ItemData(CmbDesde.ListIndex)
     End If
     Set cRsl = Nothing
     If CmbHasta.ListCount <> 0 Then
        CmbHasta.ListIndex = 0
     End If
  End If
End Sub
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
          TxtSucursal.text = nSuc
          TxtNumero.text = nNum
          HabiCombo False, True
          cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", ""
          CmbHasta.ListIndex = 0
          TxtProveedor.Enabled = True
          LblProveedor.Enabled = True
          CmdProveedor.Enabled = True
          CmdOrden.Enabled = False
        End If
  End If
  Set cNum = Nothing
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


Private Sub CmbUnidad_Click()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  If CmbUnidad.ListIndex <> -1 And TxtProducto.text <> "" Then
     If cRsl.TraerValorDeUnCampo("ProductoUnidades", "UnidadA", "Producto='" & TxtProducto.text & "' and UnidadA=" & CmbUnidad.ItemData(CmbUnidad.ListIndex)) = 0 Then
        MsgBox "Esta Unidad de Medida no esta definida en este articulo", vbCritical, "Atención"
        CmbUnidad.ListIndex = -1
     End If
  End If
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
         Case 7
              Primero
         Case 8
              ultimo
         Case 9
              Modificar
  End Select
End Sub

Private Sub CmdOrden_Click()
  Dim rAuxD As ADODB.Recordset, nMovi As Long
  On Error GoTo errHandler
 
  Set rAuxD = New ADODB.Recordset
  Set cRsl = New ClsLectura
  
  nLlama = 3
  ComproBuscar.Show 1
  If ComproBuscar.nId <> 0 Then
      nMovi = cRsl.TraerValorDeUnCampo("CabComprobantes", "Movimiento", "Id=" & ComproBuscar.nId)
      TxtProveedor.text = cRsl.TraerValorDeUnCampo("CabComprobantes", "Cliente", "Id=" & ComproBuscar.nId)
      LblProveedor.Caption = cRsl.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Proveedor='" & TxtProveedor.text & "'")
      TxtMotivo.text = "Factura Nº:" & Format(cRsl.TraerValorDeUnCampo("CabComprobantes", "Sucursal", "Id=" & ComproBuscar.nId), "0000") & "-" & Format(cRsl.TraerValorDeUnCampo("CabComprobantes", "Numero", "Id=" & ComproBuscar.nId), "00000000")
  End If
    
  If nMovi <> 0 Then
     CrearRsDetalles
     Set rAuxD = cRsl.TraerRS("DetallesTraerUnos", nMovi, True)
     sw1 = False
     If rAuxD.RecordCount <> 0 Then
        rAuxD.MoveFirst
        Do While Not rAuxD.EOF
           Rsd.AddNew
         '  Rsd!Id = 0
           Rsd!Producto = rAuxD!Producto
           Rsd!Descripcion = rAuxD!Descripcion
           Rsd!Cantidad = rAuxD!Cantidad
           Rsd!Tasa = cRsl.TraerValorDeUnCampo("ProductoImpuesto", "Impuesto", "Producto='" & TxtProducto.text & "'")
           If CmbDesde.ListIndex <> -1 Then
              Rsd!Deposito = CmbDesde.ItemData(CmbDesde.ListIndex)
           Else
             Rsd!Deposito = 0
           End If
           Rsd!Medida = cRsl.TraerValorDeUnCampo("Productos", "UMVenta", "Producto='" & Rsd!Producto & "'")
           Rsd!Impuesto = 0
           Rsd.Update
           rAuxD.MoveNext
        Loop
        CmdBotones(0).Enabled = True
     End If
     nFactura = 1
     bStock = True
     BotonDetalles False, False, False, False
     HabilitarTodo True, False
     HabilitarDetalles False
     rAuxD.Close
  Else
     MsgBox "No existe La Factura Ingresada", vbCritical, "Atención"
  End If

  CabGrid
  sw1 = True
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
   
End Sub

Private Sub CmdProveedor_Click()
  Dim RsC As ADODB.Recordset, cRiL As ClsClienteL, cRcL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRiL = New ClsClienteL
  Set cRcL = New ClsClienteL
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 2
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ProveedorTraerUno", nDat, True)
     TxtProveedor.text = RsC!Proveedor
     LblProveedor.Caption = RsC!RazonSocial
     bProveedor = True
     BotonDetalles True, False, False, False
     CmdDetalle(0).Enabled = True
     CmdDetalle(0).SetFocus
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub


Private Sub CmdProductos_Click()
  Dim RsP As ADODB.Recordset, nUni As Integer
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nLista = 1
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.text = RsP!Producto
     TxtDetalle.text = RsP!Descripcion
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
         ctl.text = ""
      End If
  Next ctl
  DtpFecha.Value = Date
  LblProveedor.Caption = ""
  TxtImporte.text = "0.00"
  TxtSucursal.text = "0000"
  TxtNumero.text = "00000000"
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
  CmdOrden.Enabled = bEstado
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
  
  CmbComprobante.text = cCompr.DatoCombo("Comprobantes", "Id", "Descripcion", Rs!Comprobante)
  If Rs!Desde <> 0 Then
     CmbDesde.text = cCompr.DatoCombo("Depositos", "Deposito", "Descripcion", Rs!Desde)
  Else
     CmbDesde.ListIndex = -1
  End If
  If Rs!Hasta <> 0 Then
     CmbHasta.text = cCompr.DatoCombo("Depositos", "Deposito", "Descripcion", Rs!Hasta)
  Else
     CmbHasta.ListIndex = -1
  End If
  DtpFecha.Value = Rs!fecha
  TxtSucursal.text = Format(Rs!Sucursal, "0000")
  TxtNumero.text = Format(Rs!Numero, "00000000")
  nMovi = Rs!Movimiento
  TxtProveedor.text = Rs!Cliente
  LblProveedor.Caption = cRcL.BuscarNombreProveedor(Rs!Cliente)
  nClient = Rs!Cliente
  TxtMotivo.text = "" & Rs!Motivo
  nFactura = "" & Rs!Lista
  
  Botones True, True, True, True, True, True, True
  If Rs!factura <> 0 Then
     Botones True, True, False, True, True, True, True
  End If

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
 
 Botones True, False, False, False, False, True, False
 BotonDetalles False, False, False, False
 BotonBuscar False, False

 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(4).Picture = LoadResPicture("Buscar", 0)
 CmdBotones(5).Picture = LoadResPicture("Salir", 0)
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
 If CmbComprobante.ListCount = 0 Or CmbDesde.ListCount = 0 Or CmbHasta.ListCount = 0 Then
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
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=7 and Suc=" & nSucursal
  cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  cRsl.CargaCombo CmbUnidad, "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbImpuesto, "Impuestos", "Impuesto", "Descripcion", "Producto='1' OR Retencion='1' OR Percepcion='1' OR Interno='1'"
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bMod As Boolean, bBus As Boolean, _
  bVal As Boolean, bSal As Boolean, bModi As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bBorr
  CmdBotones(2).Enabled = bBus
  CmdBotones(3).Enabled = bBus
  CmdBotones(4).Enabled = bVal
  CmdBotones(5).Enabled = bSal
  CmdBotones(7).Enabled = bBus
  CmdBotones(8).Enabled = bBus
  CmdBotones(9).Enabled = bMod
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
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Visible = False
      .Columns(1).Visible = False
      .Columns(2).Caption = "Producto"
      .Columns(2).Width = 1000
      .Columns(3).Caption = "Descripción"
      .Columns(3).Width = 5000
      .Columns(4).Visible = False
      .Columns(5).Caption = "Cant."
      .Columns(5).Width = 800
      .Columns(5).Alignment = dbgRight
      .Columns(5).NumberFormat = "#0.00"
      .Columns(6).Visible = False
      .Columns(7).Visible = False
      .Columns(8).Visible = False
      .Columns(9).Visible = False
      .Columns(10).Visible = False
      .Columns(11).Visible = False
      .Columns(12).Visible = False
      .Columns(13).Visible = False
      .Columns(14).Visible = False
      .Columns(15).Visible = False
      .Columns(16).Visible = False
      .Columns(17).Visible = False
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
     BotonDetalles True, False, False, True
     CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(0).Caption = "Grabar"
     CmbUnidad.ListIndex = 0
     CmbImpuesto.ListIndex = 1
     TxtCantidad.text = "0.00"
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
        Botones False, False, True, False, False, True, True
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
     If Rsd.RecordCount <> 0 Then
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
     Botones False, False, False, False, False, False, False
     If CmdBotones(0).Caption = "Grabar" Then
        Botones True, False, False, False, False, True, False
     Else
        Botones False, False, True, False, False, True, True
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
  Dim crs As ClsLectura
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
 ' Rsd!Id = 0
  Rsd!Producto = TxtProducto.text
  Rsd!Descripcion = TxtDetalle.text
  Rsd!Cantidad = TxtCantidad.text
  Rsd!Tasa = CmbImpuesto.ItemData(CmbImpuesto.ListIndex)
  Rsd!Cuenta = 0
  Rsd!Descuento = 0
  If CmbDesde.ListIndex <> -1 Then
     Rsd!Deposito = CmbDesde.ItemData(CmbDesde.ListIndex)
  Else
    Rsd!Deposito = 0
  End If
  If CmbUnidad.ListIndex <> -1 Then
     Rsd!Medida = CmbUnidad.ItemData(CmbUnidad.ListIndex)
  Else
     Rsd!Medida = 0
  End If
  Rsd!PrecioUnitario = TxtImporte.text
  Rsd.Update
  VerEstado
  Set crs = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub LimpiarDetalles()
 TxtProducto.text = ""
 TxtCantidad.text = "1.00"
 TxtDetalle.text = ""
 CmbUnidad.ListIndex = -1
 TxtImporte.text = "0.00"
 CmbImpuesto.ListIndex = -1
End Sub
Private Sub LinkearDetalles()
  Dim cRsPl As ClsProductoL
  On Error GoTo errHandler
  
  Set cRsPl = New ClsProductoL
  Set cRsl = New ClsLectura
  
  TxtProducto.text = Rsd!Producto
  TxtCantidad.text = Format(Rsd!Cantidad, "#0.00")
  TxtDetalle.text = Rsd!Descripcion
  CmbImpuesto.text = cRsl.TraerValorDeUnCampo("Impuestos", "Descripcion", "Impuesto=" & Rsd!Tasa)
  If Not IsNull(Rsd!Medida) Then
     CmbUnidad.text = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & Rsd!Medida)
  End If
  TxtImporte.text = Format(Rsd!PrecioUnitario, "#0.000")
  
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub TxtCantidad_GotFocus()
  Dim cCan As ClsComprobantesL
  Set cCan = New ClsComprobantesL
  
  If nTip = 5 Or nTip = 7 Then
     TxtCantidad.text = cCan.TraerCantidad(TxtProducto.text, CmbDesde.ItemData(CmbDesde.ListIndex))
  End If
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.text)
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtCantidad) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCantidad_LostFocus()
  If TxtCantidad.text = "" Then TxtCantidad.text = 1
  TxtCantidad.text = Format(TxtCantidad.text, "0.00")
End Sub

Private Sub TxtImporte_GotFocus()
  TxtImporte.SelStart = 0
  TxtImporte.SelLength = Len(TxtImporte.text)
End Sub

Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtImporte) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtImporte_LostFocus()
  If TxtImporte.text = "" Then TxtImporte.text = 1
  TxtImporte.text = Format(TxtImporte.text, "0.00")
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
     If cRcL.BuscarNombreProveedor(TxtProveedor.text) <> "" Then
        LblProveedor.Caption = cRcL.BuscarNombreProveedor(TxtProveedor.text)
        bProveedor = True
        VerCabecera
        CmdDetalle(0).Enabled = True
        CmdDetalle(0).SetFocus
     Else
        CmdDetalle(0).Enabled = False
        CmdProveedor.SetFocus
        bProveedor = False
     End If
  End If
  Set cRcL = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
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
Exit Sub

End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtProducto.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  CmbUnidad.Enabled = bHabi
  CmbImpuesto.Enabled = bHabi
  TxtImporte.Enabled = bHabi
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
  Dim cComp As ClsComprobantesL, cRsx As ClsLectura
  On Error GoTo errHandler

  Set cComp = New ClsComprobantesL
  Set cRsx = New ClsLectura
  Set Rs_Sto = cComp.TraerComprobantesRemitoCompras
  Botones True, False, False, False, False, True, False
  If Rs_Sto.BOF = False And Rs_Sto.EOF = False Then
     Rs_Sto.MoveFirst
     Set Rs = cRsx.TraerRS("CabezaTraerUno", Rs_Sto!ID, True)
     LinkearTexto
     LLenarDetalles
     nCas = 0
  End If
  Set cRsx = Nothing
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub Siguiente()
   On Error GoTo errHandler
  If Rs_Sto.RecordCount <> 0 Then
     Rs_Sto.MovePrevious
     If Rs_Sto.BOF Then
        Rs_Sto.MoveFirst
        Beep
     End If
     Set cRsl = New ClsLectura
     Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!ID, True)
     Set cRsl = Nothing
     CrearRsDetalles
     Limpiar
     LinkearTexto
     LLenarDetalles
     HabilitarTodo False, True
     BotonDetalles False, False, False, False
   End If
Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Anterior()
  On Error GoTo errHandler
  Rs_Sto.MoveNext
  If Rs_Sto.EOF Then
     Rs_Sto.MoveLast
     Beep
  End If
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!ID, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Primero()
  On Error GoTo errHandler
  Rs_Sto.MoveLast
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!ID, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub ultimo()
  On Error GoTo errHandler
  Rs_Sto.MoveFirst
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!ID, True)
  Set cRsl = Nothing
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
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
     Rsd!Tasa = rAuxD!Tasa
     Rsd!Deposito = rAuxD!Deposito
     Rsd!Medida = rAuxD!Medida
     Rsd!Cuenta = rAuxD!Cuenta
     Rsd!Cantidad = rAuxD!Cantidad
     Rsd!Impuesto = rAuxD!Impuesto
     Rsd!Descuento = rAuxD!Descuento
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
     Botones False, False, False, False, False, True, False
     BotonDetalles False, False, False, False
     BotonBuscar True, False
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(5).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(5).Caption = "Cancelar"
     sw = False
     bStock = False
     CrearRsStock
     CrearRsDetalles
     CmbComprobante.SetFocus
     nMovi = 0
     nCas = 1
     nFactura = 0
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
       ' Rs_Sto.MoveLast
        Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Sto!ID, True)
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
        If Len(CmbComprobante.text) <> 0 And Len(TxtProveedor.text) <> 0 And Len(CmbHasta.text) <> 0 _
           And Rsd.RecordCount <> 0 And nCas = 2 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(9).Enabled = False
        CmdBotones(4).Enabled = False
        If Len(CmbComprobante.text) <> 0 And Len(TxtProveedor.text) <> 0 And Len(CmbHasta.text) <> 0 _
           And Rsd.RecordCount <> 0 And nCas = 2 Then
           CmdBotones(9).Enabled = True
        End If
     End If
  End If
End Sub
Private Sub VerCabecera()
  CmdDetalle(0).Enabled = False
  If Len(CmbComprobante.text) <> 0 And Len(LblProveedor.Caption) <> 0 Or CmbHasta.Enabled = True Then
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
  RsCv!Suc = nSucursal
  RsCv!tipo = cTipo.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  If CmdBotones(0).Caption = "Grabar" Then
     RsCv!Movimiento = cNum.TraerUltimoMovimiento
  Else
     RsCv!Movimiento = Rs!Movimiento
  End If
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DtpFecha.Value
  RsCv!FechaIva = DtpFecha.Value
  RsCv!Numero = TxtNumero.text
  RsCv!Sucursal = TxtSucursal.text
  RsCv!Cliente = TxtProveedor.text
  RsCv!Lista = nFactura
  RsCv!Neto = 0
  RsCv!Debe = 0
  RsCv!Haber = 0
  RsCv!Descuentos = 0
  RsCv!Iva1 = 0
  RsCv!Iva2 = 0
  RsCv!COTIZACION = 1
  RsCv!Descuentos = 0
  RsCv!Financiacion = 0
  RsCv!NoGravados = 0
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
  RsCv!factura = IIf(nFactura = 1, 1, 0)
  RsCv!Motivo = TxtMotivo.text
  RsCv!Anulado = 0
  RsCv.Update
  ' Actualizo Stock
  
  Dim nCantidadA As Single, nCantidadB As Single, nTotConver As Single
  
  If bStock = False Then
     Rsd.MoveFirst
     nTip = cNum.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))
     If nTip = 30 Then
        Do While Not Rsd.EOF
           rCant.AddNew
           rCant!Suc = nSucursal
           rCant!Movimiento = RsCv!Movimiento
           rCant!Estado = 1
           rCant!Producto = Rsd!Producto
           nCantidadA = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadA", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!Medida)
           nCantidadB = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadB", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!Medida)
           nTotConver = nCantidadA * nCantidadB
           rCant!Deposito = CmbHasta.ItemData(CmbHasta.ListIndex)
           rCant!CantidadDebe = Rsd!Cantidad * nTotConver
           rCant!CantidadHaber = 0
           rCant!Anulado = 0
           rCant.Update
           Rsd.MoveNext
        Loop
     End If
     If nTip = 34 Then
        Do While Not Rsd.EOF
           rCant.AddNew
           rCant!Suc = nSucursal
           rCant!Movimiento = RsCv!Movimiento
           rCant!Estado = 1
           rCant!Producto = Rsd!Producto
           nCantidadA = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadA", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!Medida)
           nCantidadB = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadB", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!Medida)
           nTotConver = nCantidadA * nCantidadB
           rCant!Deposito = CmbHasta.ItemData(CmbHasta.ListIndex)
           rCant!CantidadDebe = 0
           rCant!CantidadHaber = Rsd!Cantidad * nTotConver
           rCant!Anulado = 0
           rCant.Update
           Rsd.MoveNext
        Loop
     End If
  End If
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
  nLlama = 5
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
  
  Dim nOpe As Integer, nTipo As Integer, cLer As ClsComprobantesL
   
  Set cLer = New ClsComprobantesL
  
  nOpe = cLer.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  nTipo = cLer.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))

  If MsgBox("Anula el Comprobante Seleccionado ?", 20, "Atención") = 6 Then
     cBor.AnularComprobantes Rs!Movimiento, nOpe, nTipo
     Rs_Sto.Delete
     Botones True, False, False, False, False, False, True
     Anterior
  End If
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
     DtpFecha.Enabled = False
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
  MsgBox err.Description, vbCritical, "Modificar"

End Sub


Private Sub TxtSucursal_GotFocus()
  TxtSucursal.SelStart = 0
  TxtSucursal.SelLength = Len(TxtSucursal.text)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSucursal_LostFocus()
  If TxtSucursal.text = "" Then TxtSucursal.text = 1
  TxtSucursal.text = Format(TxtSucursal.text, "0000")
End Sub

Private Sub TxtNumero_GotFocus()
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtNumero_LostFocus()
  If TxtNumero.text = "" Then TxtNumero.text = 0
  TxtNumero.text = Format(TxtNumero.text, "00000000")
End Sub


