VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaCompRemitos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos de Ventas"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   10965
   Begin VB.TextBox TxtSucursal 
      Height          =   315
      Left            =   6945
      TabIndex        =   2
      Text            =   "0000"
      Top             =   90
      Width           =   495
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   7740
      TabIndex        =   3
      Text            =   "00000000"
      Top             =   90
      Width           =   855
   End
   Begin VB.ComboBox CmbHasta 
      Height          =   315
      Left            =   75
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   285
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton CmdProveedor 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   5985
      TabIndex        =   7
      Top             =   990
      Width           =   735
   End
   Begin VB.TextBox TxtProveedor 
      Height          =   285
      Left            =   1185
      MaxLength       =   15
      TabIndex        =   6
      Text            =   "TxtProveedor"
      Top             =   990
      Width           =   795
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   690
      Index           =   5
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5415
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
      TabIndex        =   36
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
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   870
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Anular"
      Height          =   690
      Index           =   1
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   870
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   690
      Index           =   0
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   33
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
      Height          =   4410
      Left            =   45
      TabIndex        =   24
      Top             =   1710
      Width           =   8640
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   660
         Index           =   1
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3615
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   660
         Index           =   2
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3615
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   660
         Index           =   3
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3615
         Width           =   870
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   2115
         MaxLength       =   100
         TabIndex        =   10
         Text            =   "TxtDetalle"
         Top             =   2925
         Width           =   4800
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   660
         Index           =   0
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3615
         Width           =   870
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   945
         MaxLength       =   10
         TabIndex        =   14
         Text            =   "TxtCa"
         Top             =   3600
         Width           =   930
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   7005
         TabIndex        =   11
         Top             =   2940
         Width           =   705
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   960
         MaxLength       =   25
         TabIndex        =   9
         Text            =   "TxtProducto"
         Top             =   2940
         Width           =   1095
      End
      Begin VB.ComboBox CmbUnidad 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3255
         Width           =   2880
      End
      Begin VB.ComboBox CmbImpuesto 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3255
         Width           =   2880
      End
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   930
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "TxtImporte"
         Top             =   3930
         Width           =   930
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2640
         Left            =   165
         TabIndex        =   28
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
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   165
         TabIndex        =   32
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   165
         TabIndex        =   31
         Top             =   2940
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "U.Medida :"
         Height          =   195
         Left            =   165
         TabIndex        =   30
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "Importe :"
         Height          =   255
         Left            =   165
         TabIndex        =   29
         Top             =   3945
         Width           =   780
      End
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2535
   End
   Begin VB.ComboBox CmbDesde 
      Height          =   315
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   585
      Width           =   3015
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   690
      Index           =   4
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2370
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   690
      Index           =   6
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   22
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
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   21
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
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2370
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Modificar"
      Height          =   690
      Index           =   9
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1620
      Width           =   1005
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
      TabIndex        =   18
      Top             =   990
      Width           =   540
   End
   Begin VB.TextBox TxtMotivo 
      Height          =   285
      Left            =   1185
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   8
      Text            =   "TxtMotivo"
      Top             =   1380
      Width           =   7485
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Transporte"
      Height          =   690
      Index           =   10
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3870
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   315
      Left            =   4710
      TabIndex        =   1
      Top             =   90
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   312016897
      CurrentDate     =   36783
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   240
      Left            =   6060
      TabIndex        =   46
      Top             =   90
      Width           =   765
   End
   Begin VB.Label Label17 
      Caption         =   "--"
      Height          =   195
      Left            =   7530
      TabIndex        =   45
      Top             =   135
      Width           =   150
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
      TabIndex        =   44
      Top             =   990
      Width           =   3885
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente :"
      Height          =   255
      Left            =   90
      TabIndex        =   43
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   4125
      TabIndex        =   42
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   90
      TabIndex        =   41
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "Depósito :"
      Height          =   195
      Left            =   90
      TabIndex        =   40
      Top             =   585
      Width           =   930
   End
   Begin VB.Label Label7 
      Caption         =   "Factura :"
      Height          =   180
      Left            =   6960
      TabIndex        =   39
      Top             =   990
      Width           =   705
   End
   Begin VB.Label Label8 
      Caption         =   "Observación :"
      Height          =   165
      Left            =   90
      TabIndex        =   38
      Top             =   1395
      Width           =   990
   End
End
Attribute VB_Name = "VentaCompRemitos"
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
Dim rs As ADODB.Recordset             ' Cabecera Stock Tabla Todos
Dim Rs_Sis As ADODB.Recordset         ' Tabla Sistema, Ultima Stock
Dim nTip As Integer                   ' Valor de tipo de Movimiento
Public nMovi As Long, nClient As String, nMoviFac As Long, nMoviTrans As Long
Public nCas As Integer
Public rImp As ADODB.Recordset
Dim cRsl As ClsLectura
Dim cComp As ClsComprobantesL
Dim bRemi As Boolean, nLin As Byte
Dim rRem As ADODB.Recordset
Dim nReg As Variant, nLista As Integer
Dim bStock As Boolean, nFactura As Long

Private Sub CmbDesde_Click()
  VerEstado
'  VerCabecera
  
  CmbHasta.ListIndex = -1
  If nTip = 5 Then
     If CmbDesde.ListIndex <> -1 Then
        cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", "Deposito<>" & CmbDesde.ItemData(CmbDesde.ListIndex & " And Suc=" & nSucursal)
     End If
     
     If CmbHasta.ListCount <> 0 Then
        CmbHasta.ListIndex = 0
     End If
  End If
End Sub
Private Sub CmbComprobante_Click()
volver:
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler
 
  If sw = False Then
     Set cNum = New ClsComprobantesL
        If CmbComprobante.ListIndex <> -1 Then
           If cRsl.TraerValorDeUnCampo("Impresion", "Numero", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) = False Then
              nNum = Format(cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
              nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
           Else
              nNum = Format(cNum.TraerUltimoNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
              nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
           End If
           TxtSucursal.text = nSuc
           TxtNumero.text = nNum
           HabiCombo True, False
           cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
           CmbDesde.ListIndex = 0
           CmbHasta.ListIndex = -1
           TxtProveedor.Enabled = True
           LblProveedor.Enabled = True
           CmdProveedor.Enabled = True
          ' CmdOrden.Enabled = False
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

Private Sub CmbImpuesto_Change()
   TodoBien
End Sub

Private Sub CmbUnidad_Change()
   TodoBien
End Sub

Private Sub CmbUnidad_Click()
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
         Case 6
              Imprimir
         Case 7
              Primero
         Case 8
              ultimo
         Case 9
              Modificar
         Case 10
              Transporte
  End Select
End Sub

Private Sub Transporte()
  StockTransporte.Show 1
End Sub

Private Sub CmdOrden_Click()
  Dim rAuxD As ADODB.Recordset, nMovi As Long, nCantRem As Single, nCantRemAux As Single, cRem As ClsComprobantesL, bSaldo As Byte
  On Error GoTo errHandler
 
  Set rAuxD = New ADODB.Recordset
  Set cRem = New ClsComprobantesL
  
  nLlama = 1
  ComproBuscar.Show 1
  If ComproBuscar.nId <> 0 Then
      nMovi = cRsl.TraerValorDeUnCampo("CabComprobantes", "Movimiento", "Id=" & ComproBuscar.nId)
      If cRsl.TraerValorDeUnCampo("CabComprobantes", "Descuentos", "Descuentos=" & nMovi) = 0 Then
         TxtProveedor.text = cRsl.TraerValorDeUnCampo("CabComprobantes", "Cliente", "Id=" & ComproBuscar.nId)
         LblProveedor.Caption = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & TxtProveedor.text & "'")
         TxtMotivo.text = "Factura Nº:" & Format(cRsl.TraerValorDeUnCampo("CabComprobantes", "Sucursal", "Id=" & ComproBuscar.nId), "0000") & "-" & Format(cRsl.TraerValorDeUnCampo("CabComprobantes", "Numero", "Id=" & ComproBuscar.nId), "00000000")
      Else
         MsgBox "Factura ya Aplicada en un Remito", vbCritical, "Atención"
         Exit Sub
      End If
  End If
  nMoviFac = nMovi
  CrearRsDetalles
  BotonDetalles True, False, False, False
  
   Set rAuxD = cRsl.TraerRS("DetallesTraerUnos", nMovi, True)
   sw1 = False
   bSaldo = 0
   If rAuxD.RecordCount <> 0 Then
      rAuxD.MoveFirst
      Do While Not rAuxD.EOF
         Rsd.AddNew
         Rsd!Producto = rAuxD!Producto
         Rsd!Descripcion = rAuxD!Descripcion
         Rsd!Cantidad = rAuxD!Cantidad
         Rsd!Tasa = cRsl.TraerValorDeUnCampo("ProductoImpuesto", "Impuesto", "Producto='" & rAuxD!Producto & "'")
         If CmbDesde.ListIndex <> -1 Then
            Rsd!deposito = CmbDesde.ItemData(CmbDesde.ListIndex)
         Else
            Rsd!deposito = 0
         End If
         Rsd!medida = cRsl.TraerValorDeUnCampo("Productos", "UMVenta", "Producto='" & rAuxD!Producto & "'")
         Rsd!Impuesto = 0
         Rsd!PrecioUnitario = rAuxD!PrecioUnitario
         Rsd!PrecioTotal = rAuxD!PrecioTotal
         Rsd.Update
         rAuxD.MoveNext
      Loop
      If Rsd.RecordCount = 0 Then
         BotonDetalles False, False, False, False
         Botones False, False, False, False, False, True, False
         nMoviFac = 0
      Else
         nFactura = 1
         bStock = True
         BotonDetalles False, True, False, False
         Botones True, False, False, False, False, True, False
         HabilitarTodo True, True
         HabilitarDetalles False
         rAuxD.Close
      End If
   Else
       MsgBox "No existe La Factura Ingresada", vbCritical, "Atención"
   End If
   sw1 = True
  
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
  
 
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ClienteTraerUno", nDat, True)
     TxtProveedor.text = RsC!Cliente
     LblProveedor.Caption = RsC!RazonSocial
     nLista = RsC!precio
     bProveedor = True
     BotonDetalles True, False, False, False
     CmdOrden.Enabled = False
     CmdDetalle(0).Enabled = True
     CmdDetalle(0).SetFocus
     RsC.Close
  End If
 
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmdProductos_Click()
  Dim RsP As ADODB.Recordset, nUni As Integer, sImp As String, cRpl As ClsProductoL, cRem As ClsComprobantesL
  
  On Error GoTo errHandler
  Dim nCantPen As Single
  
  Set RsP = New ADODB.Recordset
    
  Set cRpl = New ClsProductoL
  Set cRem = New ClsComprobantesL
  
  nDat = 0
  nBuscar = 3

  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.text = RsP!Producto
     TxtDetalle.text = RsP!Descripcion
     sImp = cRpl.ProductoImpuestoDesc(TxtProducto.text)
     If sImp <> "" Then
        CmbImpuesto.text = sImp
     End If
     nUni = cRsl.TraerValorDeUnCampo("Productos", "UMCompra", "Producto='" & TxtProducto.text & "'")
     CmbUnidad.text = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & nUni)
      TxtImporte.text = Format(cRpl.TraerPrecio(nLista, TxtProducto.text), nCantDecimales)
      If nMoviFac <> 0 Then
         nCantPen = cRem.TraerCantidadFactura(TxtProducto.text, nMoviFac, True)
         If nCantPen <> 0 Then
           TxtCantidad.text = Format(nCantPen, "#0.00")
         Else
            MsgBox "Limite de Producto Consumido", vbCritical, "Atención"
            LimpiarDetalles
            TxtProducto.SetFocus
            Exit Sub
        End If
     End If
     TxtCantidad.SetFocus
     RsP.Close
  End If
  
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
      KeyAscii = 0
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
  TxtImporte.text = nCantDecimales
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
         Or TypeOf ctl Is ComboBox _
         Or TypeOf ctl Is OptionButton Then
         ctl.Enabled = bEstado
      End If
  Next ctl
  CmdOrden.Enabled = False
  Grid1.Enabled = bGrid
End Sub
Private Sub LinkearTexto()
  On Error GoTo errHandler
  Dim nSuc As String * 4, nNum As String * 8, p As ClsPrograma
          
  Set p = New ClsPrograma
   
  If Not IsNull(rs!Comprobante) Then p.SetComboByItemData CmbComprobante, CLng(rs!Comprobante)
  
  If rs!desde <> 0 Then
     p.SetComboByItemData CmbDesde, CLng(rs!desde)
  Else
     CmbDesde.ListIndex = -1
  End If
  If rs!hasta <> 0 Then
     p.SetComboByItemData CmbHasta, CLng(rs!hasta)
  Else
     CmbHasta.ListIndex = -1
  End If
  
  DTPFecha.Value = rs!fecha
  TxtSucursal.text = Format(rs!sucursal, "0000")
  TxtNumero.text = Format(rs!Numero, "00000000")
  nMovi = rs!Movimiento
  nMoviTrans = rs!Movimiento
  TxtProveedor.text = rs!Cliente
  LblProveedor.Caption = rs!RazonSocial
  nClient = rs!Cliente
  TxtMotivo.text = "" & rs!Motivo
  nFactura = "" & rs!Lista
  nLista = rs!precio

  Botones True, True, True, True, True, True, True
  If rs!factura <> 0 Then
     Botones True, True, False, True, True, True, True
  End If
  
  If cRsl.TraerValorDeUnCampo("DetallesComprobantes", "MoviRemi", "MoviRemi=" & rs!Movimiento) <> 0 Then
      Botones True, False, False, True, True, True, True
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

 Set rs = New ADODB.Recordset
 Set Rs_Sis = New ADODB.Recordset
 Set cRsl = New ClsLectura
 Set cComp = New ClsComprobantesL
  
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
 CmdBotones(10).Picture = LoadResPicture("Valores", 0)
 
 
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
 
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=6 And Suc=" & nSucursal
  cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  cRsl.CargaCombo CmbUnidad, "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbImpuesto, "Impuestos", "Impuesto", "Descripcion", "Producto='1' OR Retencion='1' OR Percepcion='1' OR Interno='1'"
  
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
  CmdBotones(6).Enabled = bBus
  CmdBotones(7).Enabled = bBus
  CmdBotones(8).Enabled = bBus
  CmdBotones(9).Enabled = bMod
  CmdBotones(10).Enabled = bBus
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
 
 Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
 
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub CrearRsDetalles()
 On Error GoTo errHandler

 Set Rsd = New ADODB.Recordset
  
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

Private Sub HabilitarTipo(pHabi As Boolean)
  CmdOrden.Enabled = pHabi
End Sub

Private Sub NuevoRsD()
  Dim cRp As ClsLectura
  
  Set cRp = New ClsLectura
  
  On Error GoTo errHandler
  If nMoviFac = 0 Then pClave = False
  If pClave = True Then
      FrmClave.Show 1
      If bClave = True Then
Entrar:
          If sw1 = True Then
             sw1 = False
             CmdBotones(5).Enabled = False
             BotonBuscar False, True
             HabilitarTodo False, False
             HabilitarDetalles True
             LimpiarDetalles
             BotonDetalles False, False, False, True
             Botones False, False, False, False, False, True, False
             HabilitarTipo False
             CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
             CmdDetalle(0).Caption = "Grabar"
             CmbUnidad.text = cRp.DatoCombo("UnidadMedida", "Unidad", "Descripcion", cRp.RegPorDefecto("Medida"))
             If cRp.TraerValorDeUnCampo("Comprobantes", "Iva", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 0 Then
                CmbImpuesto.text = "IVA EXENTO"
             Else
                CmbImpuesto.text = cRp.DatoCombo("Impuestos", "Impuesto", "Descripcion", cRp.RegPorDefecto("Pais"))
             End If
        
             TxtCantidad.text = "0.00"
             TxtProducto.SetFocus
          Else
             If TxtCantidad.text <= 0 Then
                MsgBox "El campo Cantidad no puede ser igual a cero", vbCritical, "Atención"
                TxtCantidad.SetFocus
                Exit Sub
             End If
             CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
             CmdDetalle(0).Caption = "Agregar"
             Rsd.AddNew
             GrabarRsd
             BotonDetalles True, True, True, False
             BotonBuscar True, False
             sw1 = True
             If nTip <> 5 Then
                CmbDesde.Enabled = False
             End If
             If CmdBotones(0).Caption = "Grabar" Then
                Botones True, False, False, False, False, True, False
             Else
                Botones False, False, True, False, False, True, True
             End If
             HabilitarTodo True, True
             HabilitarDetalles False
          End If
     Else
         MsgBox "No Tiene Permisos para Realizar Esta Operación", vbExclamation, "Atención"
      End If
 Else
    GoTo Entrar:
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
        If CmdBotones(0).Caption = "Grabar" Then
           Botones True, False, False, False, False, True, False
        Else
           Botones False, False, True, False, False, True, True
        End If
        If nMoviFac <> 0 Then
           BotonDetalles False, True, False, False
        End If
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
     Botones False, False, False, False, False, True, False
     CmdDetalle(2).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(2).Caption = "Grabar"
  Else
     If TxtCantidad.text <= 0 Then
          MsgBox "El campo Cantidad no puede ser igual a cero", vbCritical, "Atención"
          TxtCantidad.SetFocus
          Exit Sub
     End If
     CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
     CmdDetalle(2).Caption = "Actualizar"
     GrabarRsd
     BotonDetalles True, True, True, False
     If nMoviFac <> 0 Then
        BotonDetalles False, True, True, False
     End If
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
           If CmdBotones(0).Caption = "Grabar" Then
              Botones True, False, False, False, False, True, False
           Else
              Botones False, False, True, False, False, True, False
           End If
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
        If CmdBotones(0).Caption = "Grabar" Then
           Botones True, False, False, False, False, True, False
        Else
           Botones False, False, True, False, False, True, False
        End If
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

  
'  Rsd!Id = 0
  Rsd!Producto = TxtProducto.text
  Rsd!Descripcion = TxtDetalle.text
  Rsd!Cantidad = TxtCantidad.text
  Rsd!Tasa = CmbImpuesto.ItemData(CmbImpuesto.ListIndex)
  Rsd!Cuenta = 0
  Rsd!PDesc = 0
  Rsd!Descuento = 0
  If CmbDesde.ListIndex <> -1 Then
     Rsd!deposito = CmbDesde.ItemData(CmbDesde.ListIndex)
  Else
    Rsd!deposito = 0
  End If
  If CmbUnidad.ListIndex <> -1 Then
     Rsd!medida = CmbUnidad.ItemData(CmbUnidad.ListIndex)
  Else
     Rsd!medida = 0
  End If
  Rsd!precio = TxtImporte.text
  Rsd!PrecioUnitario = TxtImporte.text
  Rsd!PrecioTotal = CDec(TxtImporte.text) * CDec(TxtCantidad.text)
  Rsd!MoviRemi = nMoviFac
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
 TxtImporte.text = nCantDecimales
 CmbImpuesto.ListIndex = -1
End Sub
Private Sub LinkearDetalles()
  On Error GoTo errHandler
  
  Dim p As ClsPrograma
  Set p = New ClsPrograma
      
  TxtProducto.text = Rsd!Producto
  TxtCantidad.text = Format(Rsd!Cantidad, "#0.00")
  TxtDetalle.text = Rsd!Descripcion
  If Rsd!Tasa <> 0 Then
      p.SetComboByItemData CmbImpuesto, CLng(Rsd!Tasa)
  End If
  If Not IsNull(Rsd!medida) And Rsd!medida <> 0 Then
     p.SetComboByItemData CmbUnidad, CLng(Rsd!medida)
  End If
  TxtImporte.text = Format(Rsd!PrecioUnitario, nCantDecimales)
  

Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub TxtCantidad_Change()
   TodoBien
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
  Dim nCantPen As Single
  
  If TxtCantidad.text = "" Then TxtCantidad.text = 1
  TxtCantidad.text = Format(TxtCantidad.text, "0.00")
  If nMoviFac <> 0 Then
     Dim cRem As ClsComprobantesL

     Set cRem = New ClsComprobantesL
     
     nCantPen = cRem.TraerCantidadFactura(TxtProducto.text, nMoviFac, True)
       
     If nCantPen >= TxtCantidad.text Then
        TxtCantidad.text = Format(TxtCantidad.text, "#0.00")
     Else
        MsgBox "El Limite de Producto no puede ser Superior a:" & Format(nCantPen, "#0.00"), vbCritical, "Atención"
         TxtCantidad.text = Format(nCantPen, "#0.00")
         TxtCantidad.SetFocus
        Exit Sub
    End If
  End If
End Sub

Private Sub TxtDetalle_Change()
  TodoBien
End Sub

Private Sub TodoBien()
  If CmdDetalle(0).Caption = "Grabar" Then
     CmdDetalle(0).Enabled = False
     If Len(CmbImpuesto.text) <> 0 And Len(TxtDetalle.text) <> 0 Then
        CmdDetalle(0).Enabled = True
     End If
  Else
     CmdDetalle(2).Enabled = False
     If Len(CmbImpuesto.text) <> 0 And Len(TxtDetalle.text) <> 0 Then
        CmdDetalle(2).Enabled = True
        If bRemi = True Then
           BotonDetalles True, False, True, False
        End If
     End If
  End If
End Sub

Private Sub TxtDetalle_LostFocus()
  If TxtDetalle.text <> "" Then
     TxtCantidad.SetFocus
  End If
End Sub

Private Sub TxtImporte_Change()
 ' TodoBien
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
  TxtImporte.text = Format(TxtImporte.text, nCantDecimales)
End Sub

Private Sub TxtProducto_Change()
  TodoBien
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
  
  
  Set cRcL = New ClsClienteL
  Set cRiL = New ClsClienteL
  LblProveedor.Caption = ""
  If TxtProveedor.text <> "" Then
     If cRcL.BuscarNombreCliente(TxtProveedor.text) <> "" Then
        LblProveedor.Caption = cRcL.BuscarNombreCliente(TxtProveedor.text)
        nLista = cRsl.TraerValorDeUnCampo("Clientes", "Precio", "Cliente='" & TxtProveedor.text & "'")
        bProveedor = True
        VerCabecera
        CmdOrden.Enabled = False
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
  Dim cRx As ClsLectura, nUni As Integer, sImp As String
  Dim cRem As ClsComprobantesL, nCantPen As Single
  
  On Error GoTo errHandler
  
  Set cRpl = New ClsProductoL
  Set cRx = New ClsLectura
  Set cRem = New ClsComprobantesL
  If TxtProducto.text <> "" Then
  
     TxtProducto.text = cRx.TraerCodigoProducto(TxtProducto.text)

     If cRpl.EstaElCodigoCliente(TxtProveedor.text, TxtProducto.text) = False Then
        If cRpl.BuscarNombreProducto(TxtProducto.text) <> "" Then
           TxtDetalle.text = cRpl.BuscarNombreProducto(TxtProducto.text)
        Else
           LimpiarDetalles
           TxtProducto.SetFocus
           Exit Sub
        End If
     Else
        Set rPrd = cRpl.TraerUnProductoCliente(TxtProveedor, TxtProducto.text)
        sProdu = rPrd!Producto
        TxtDetalle.text = cRpl.BuscarNombreProducto(sProdu)
     End If
     sImp = cRpl.ProductoImpuestoDesc(TxtProducto.text)
     If sImp <> "" Then
        CmbImpuesto.text = sImp
     End If
     TxtImporte.text = Format(cRpl.TraerPrecio(nLista, TxtProducto.text), nCantDecimales)
     nUni = cRx.TraerValorDeUnCampo("Productos", "UMVenta", "Producto='" & TxtProducto.text & "'")
     CmbUnidad.text = cRx.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & nUni)
     If nMoviFac <> 0 Then
        nCantPen = cRem.TraerCantidadFactura(TxtProducto.text, nMoviFac, True)
        If nCantPen <> 0 Then
           TxtCantidad.text = Format(nCantPen, nDecimalCant)
        Else
           MsgBox "Limite de Producto Consumido o no Esta en el Pedido", vbCritical, "Atención"
           LimpiarDetalles
           TxtProducto.SetFocus
           Exit Sub
        End If
     End If
     TodoBien
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
 
  On Error GoTo errHandler
  
  Set Rs_Sto = cComp.TraerComprobantesRemitoVentas
  
  Botones True, False, False, False, False, True, False
  If Rs_Sto.BOF = False And Rs_Sto.EOF = False Then
     Rs_Sto.MoveFirst
     Set rs = cComp.TraerRsComprobante(Rs_Sto!ID)
     LinkearTexto
     LLenarDetalles
     nCas = 0
  End If
  
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
     
    Set rs = cComp.TraerRsComprobante(Rs_Sto!ID)
     
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
  
  Set rs = cComp.TraerRsComprobante(Rs_Sto!ID)
  
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
  
  Set rs = cComp.TraerRsComprobante(Rs_Sto!ID)
  
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
  
  Set rs = cComp.TraerRsComprobante(Rs_Sto!ID)
  
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
  
  Set rAuxD = cRsl.TraerRS("DetallesTraerUnos", rs!Movimiento, True)
  rAuxD.MoveFirst
  Do While Not rAuxD.EOF
     Rsd.AddNew
     Rsd!Movimiento = rAuxD!Movimiento
     Rsd!Producto = rAuxD!Producto
     Rsd!Descripcion = rAuxD!Descripcion
     Rsd!Tasa = rAuxD!Tasa
     Rsd!deposito = rAuxD!deposito
     Rsd!medida = rAuxD!medida
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
       nMoviFac = 0
       
       CmbComprobante.ListIndex = 0
       HabilitarTipo True
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
     If Rs_Sto.RecordCount <> 0 Then
        If Rs_Sto.BOF = False Or Rs_Sto.EOF = False Then
         '  Rs_Sto.MoveLast
           Set rs = cComp.TraerRsComprobante(Rs_Sto!ID)
           LinkearTexto
           LLenarDetalles
           nCas = 0
           BotonDetalles False, False, False, False
           Botones True, True, True, True, True, True, True
           
           If cRsl.TraerValorDeUnCampo("DetallesComprobantes", "MoviRemi", "MoviRemi=" & rs!Movimiento) <> 0 Then
              Botones True, False, False, True, True, True, True
           End If
        End If
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
  
  Set rCant = New ADODB.Recordset
  
  Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
  
  ' Grabar Cabecera Stock
  
  RsCv.AddNew
  RsCv!Suc = nSucursal
'  RsCv!Id = 0
  RsCv!tipo = cTipo.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  If CmdBotones(0).Caption = "Grabar" Then
     RsCv!Movimiento = cNum.TraerUltimoMovimiento
  Else
     RsCv!Movimiento = rs!Movimiento
  End If
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DTPFecha.Value
  RsCv!FechaIva = DTPFecha.Value
  RsCv!Numero = TxtNumero.text
  RsCv!sucursal = TxtSucursal.text
  RsCv!Cliente = TxtProveedor.text
  RsCv!Lista = nFactura
  RsCv!Vendedor = cRsl.TraerValorDeUnCampo("CabComprobantes", "Vendedor", "Movimiento=" & nMoviFac)
  RsCv!Caja = 0
  RsCv!CondVenta = 0
  RsCv!PIvaCompras = 0
  RsCv!IBrutosCompras = 0
  RsCv!Neto = 0
  RsCv!Debe = 0
  RsCv!Haber = 0
  RsCv!Iva1 = 0
  RsCv!Iva2 = 0
  RsCv!COTIZACION = 1
  RsCv!Descuentos = nMoviFac  ' Lo utilizo para saber con que factura esta relacionada
  RsCv!Financiacion = 0
  RsCv!NoGravados = 0
  If CmbDesde.ListIndex <> -1 Then
     RsCv!desde = CmbDesde.ItemData(CmbDesde.ListIndex)
  Else
     RsCv!desde = 0
  End If
  If CmbHasta.ListIndex <> -1 Then
     RsCv!hasta = CmbHasta.ItemData(CmbHasta.ListIndex)
  Else
    RsCv!hasta = 0
  End If
  RsCv!factura = IIf(nFactura = 1, 1, 0)
  RsCv!Motivo = TxtMotivo.text
  RsCv!Anulado = 0
  RsCv.Update
  ' Actualizo Stock
  
  Dim nCantidadA As Single, nCantidadB As Single, nTotConver As Single
  
  If bStock = False Then
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        nTip = cNum.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))
        If nTip = 17 Then
           rCant.AddNew
           rCant!Suc = nSucursal
           rCant!Movimiento = RsCv!Movimiento
           rCant!Estado = 1
           rCant!Producto = Rsd!Producto
           nCantidadA = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadA", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!medida)
           nCantidadB = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadB", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!medida)
           nTotConver = nCantidadA * nCantidadB
           rCant!deposito = CmbDesde.ItemData(CmbDesde.ListIndex)
           rCant!CantidadHaber = Rsd!Cantidad * nTotConver
           rCant!CantidadDebe = 0
           rCant!Anulado = 0
           rCant.Update
           Rsd.MoveNext
        End If
     Loop
  End If
  Rsd.MoveFirst
  
  If CmdBotones(0).Caption = "Grabar" Then
     cGrab.GrabarRemito RsCv, Rsd, rCant, nTip
     CrearRsDetalles
     Limpiar
     ArmaRsParaMovimiento
  Else
     cGrab.ModificarRemito RsCv, Rsd, rCant, nTip
  End If
 ' End If
  nTip = 0

  Set cGrab = Nothing
  Set cNum = Nothing
  Set cTipo = Nothing
  Set cStock = Nothing

End Sub
Private Sub Buscar()
  nLlama = 4
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
  
  Dim nOpe As Integer, nTipo As Integer, cLer As ClsComprobantesL
   
  Set cLer = New ClsComprobantesL
  
  
  nOpe = cLer.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  nTipo = cLer.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))

  If cRsl.TraerValorDeUnCampo("RegTransporte", "Estado", "Movimiento=" & Rs_Sto!Movimiento) <= 1 Then
    If MsgBox("Anula el Comprobante Seleccionado ?", 20, "Atención") = 6 Then
       cBor.AnularComprobantes rs!Movimiento, nOpe, nTipo
       Rs_Sto.Delete
       Botones True, False, False, False, False, True, False
       If Rs_Sto.RecordCount = 0 Then
          Limpiar
          LimpiarRecorset Rsd
       Else
          Anterior
       End If
    End If
  End If
End Sub
Private Sub Imprimir()
  Dim Hoja As Byte, cHoja As Byte, Pagina As Byte, i As Single, i1 As Single, h As Byte, Linea As Byte, nLinea As Byte
  Dim nUlti As Double, nPorIva As Single, nTotal As Currency

  Dim Rx As ADODB.Recordset, nReg As Variant
  
  Dim SubTotal As Currency, Impuesto As Currency, Subtotal1 As Currency, Iva1 As Currency, Iva2 As Currency, Total As Currency
    
  Dim nSubTotal As Double, nCantipage As Single
    
  Set Rx = cRsl.TraerRsCondi("Impresion", "Id", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
  
  
  Dim p As Printer, sImpre As String
        
  sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
  If sImpre = "0" Then
     MsgBox "Debe Definir Un Modelo de Impresora al Comprobante...", vbCritical, "Atención"
     Exit Sub
  End If
  
  For Each p In Printers
      If p.DeviceName = sImpre Then
         Set Printer = p
        ' Printer.FontName = "Courier New"
         Printer.ScaleMode = 7
         Printer.FontBold = cRsl.TraerValorDeUnCampo("Impresion", "Negrita", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Printer.FontSize = cRsl.TraerValorDeUnCampo("Impresion", "Tamaño", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Exit For
      End If
  Next
  Rsd.MoveFirst
  
  nSubTotal = 0
  nCantipage = Rsd.RecordCount / Rx!CantidadDetalle

  i1 = 0
  Pagina = 0
  Hoja = 1
  cHoja = Rx!CantidadHoja
  Linea = Rx!CantidadDetalle
  For h = 1 To cHoja
      If Hoja = 2 Then
         i1 = Rx!Salto
      Else
         i1 = 0
      End If
      i = Rx!PosicionDetalle + i1
      CabezaFactura i, i1, Rx
      nLinea = 1
      Do While Not Rsd.EOF
         ' Codigo Articulo
         If Rx!codigo <> "0" Then
            Printer.CurrentY = i
            Printer.CurrentX = Rx!codigo
            Printer.Print Rsd!Producto
         End If

         If Rx!unidad <> "0" Then
            Printer.CurrentY = i
            Printer.CurrentX = Rx!unidad
            Printer.Print cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & Rsd!medida)
         End If

         Printer.CurrentY = i
         Printer.CurrentX = Rx!Detalle
         Printer.Print Rsd!Descripcion
         If Rsd!Cantidad <> 0 Then
            Printer.CurrentY = i
            Printer.CurrentX = Rx!Cantidad - Printer.TextWidth(Format(Rsd!Cantidad, nCantDecimales))
            Printer.Print Format(Rsd!Cantidad, nCantDecimales)
         End If
         
         If Rx!PUnitario <> "0" Then
            Printer.CurrentY = i
            Printer.CurrentX = Rx!PUnitario - Printer.TextWidth(Format(Rsd!PrecioUnitario, nCantDecimales))
            Printer.Print Format(Rsd!PrecioUnitario, nCantDecimales)
         End If
         
         If Rx!importe <> 0 Then
            Printer.CurrentY = i
            Printer.RightToLeft = True
            Printer.CurrentX = Rx!importe - Printer.TextWidth(Format(Rsd!PrecioTotal, nCantDecimales))
            Printer.Print Format(Rsd!PrecioTotal, nCantDecimales)
            nTotal = nTotal + Rsd!PrecioTotal
            
            nSubTotal = nSubTotal + nTotal
         End If
         If nLinea = Linea Then
            nLinea = 1
            If Hoja = 1 Then
               nReg = Rsd.Bookmark
            End If
            If nCantipage <> Printer.Page Then
               If Rx!ResInscrip <> "0" Then      ' Saldo de Cuenta Corriente
                  Printer.CurrentY = TraerXY(Rx!ResInscrip, "Y") + i1
                  Printer.CurrentX = TraerXY(Rx!ResInscrip, "X")
                  Printer.Print "Sub Total Página: " & CInt(Printer.Page)
                End If
                If Rx!ResInscrip <> "0" Then      ' Saldo de Cuenta Corriente
                   Printer.CurrentY = TraerXY(Rx!Total, "Y") + i1
                   Printer.CurrentX = TraerXY(Rx!Total, "X")
                   Printer.CurrentX = TraerXY(Rx!Total, "X")
                   Printer.Print Format(nSubTotal, "#0.00")
                End If
                Printer.NewPage
                CabezaFactura i, i1, Rx
                i = Rx!PosicionDetalle + i1
            End If
         Else
            nLinea = nLinea + 1
            i = i + Rx!Interlineado
         End If
         Rsd.MoveNext
      Loop
      If Hoja = 1 Then
         i = Rx!Salto
         Hoja = 2
         If Pagina = 0 Then
            Rsd.MoveFirst
         End If
      End If
      If Rx!Total <> "0" Then
         Printer.CurrentY = TraerXY(Rx!Total, "Y") + i1
         Printer.CurrentX = TraerXY(Rx!Total, "X") - Printer.TextWidth(Format(nTotal, "#0.00"))
         Printer.Print Format(nTotal, "#0.00")
      End If

  Next h
  Printer.EndDoc
  BotonDetalles False, False, False, False
End Sub

Private Sub CabezaFactura(pI As Single, pId As Single, Rx As ADODB.Recordset)
  Dim cClie As ClsClienteL
  
  Dim rCli As ADODB.Recordset
  
  Set cClie = New ClsClienteL

  Set rCli = cClie.BuscarDatosCli(TxtProveedor.text)
  
  Dim sFech As Date, sNum As String, sCuen As String, sRazon As String, sDire As String, sLocal As String, sTel As String
  Dim sCuit As String, sIva As String, sCondi As String, sProvi As String
  
  Dim sRazonT As String, sDireT As String, sLocalT As String, sTelT As String
  Dim sCuitT As String, sIvaT As String, sCondiT As String, sProviT As String
    
  If rCli.RecordCount <> 0 Then
     sFech = DTPFecha.Value
     sNum = TxtSucursal.text & "-" & TxtNumero.text
     sCuen = TxtProveedor.text
     sRazon = LblProveedor.Caption
     sDire = "" & rCli!Domicilio
     sLocal = rCli!CodigoPostal & " - " & rCli!Ciudad
     sProvi = rCli!Provincia
     sTel = cRsl.TraerValorDeUnCampo("Clientes", "Telefono", "Cliente='" & TxtProveedor.text & "'")
     sCuit = cRsl.TraerValorDeUnCampo("Clientes", "CUIT", "Cliente='" & TxtProveedor.text & "'")
     sIva = rCli!TipoIva
  End If
  
  sRazonT = cRsl.TraerValorDeUnCampo("RegTransporte", "Destinatario", "Movimiento=" & nMovi)
  sDireT = cRsl.TraerValorDeUnCampo("RegTransporte", "Direccion", "Movimiento=" & nMovi)
  sLocalT = cRsl.TraerValorDeUnCampo("RegTransporte", "Localidad", "Movimiento=" & nMovi)
  sCuitT = cRsl.TraerValorDeUnCampo("RegTransporte", "Cuit", "Movimiento=" & nMovi)
  sCondi = cRsl.TraerValorDeUnCampo("Sistema", "FormaPago", "")

 ' Empresa
  If Rx!NEmpresa <> "0" Then
     Printer.CurrentY = TraerXY(Rx!NEmpresa, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!NEmpresa, "X")
     Printer.Print sNomFan
  End If

  If Rx!NDireccion <> "0" Then
     Printer.CurrentY = TraerXY(Rx!NDireccion, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!NDireccion, "X")
     Printer.Print sDirFan
  End If

  If Rx!NTelefono <> "0" Then
     Printer.CurrentY = TraerXY(Rx!NTelefono, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!NTelefono, "X")
     Printer.Print sTelFan
  End If

  If Rx!NCiudad <> "0" Then
     Printer.CurrentY = TraerXY(Rx!NCiudad, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!NCiudad, "X")
     Printer.Print sCiuFan
  End If
  ' Fecha
  Printer.ScaleMode = 7
  Printer.CurrentY = TraerXY(Rx!fecha, "Y") + pId
  Printer.CurrentX = TraerXY(Rx!fecha, "X")
  Printer.Print Left(sFech, 2)
  
  Printer.CurrentY = TraerXY(Rx!Fecha1, "Y") + pId
  Printer.CurrentX = TraerXY(Rx!Fecha1, "X")
  Printer.Print Mid(sFech, 4, 2)
  Printer.CurrentY = TraerXY(Rx!Fecha2, "Y") + pId
  Printer.CurrentX = TraerXY(Rx!Fecha2, "X")
  Printer.Print Right(sFech, 2)
  

  If Rx!CordTexto <> "0" Then
     Printer.CurrentY = TraerXY(Rx!CordTexto, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!CordTexto, "X")
     Dim sTit As String

     Printer.Print cRsl.TraerValorDeUnCampo("Impresion", "Texto", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) & "  " & TxtSucursal.text & "-" & TxtNumero.text & " " & sTit
  End If
  ' Cod Cliente
  If Rx!Cliente <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Cliente, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Cliente, "X")
     Printer.Print sCuen
  End If
  ' Razon Social
  If Rx!Señor <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Señor, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Señor, "X")
     If sRazonT <> "0" Then
        Printer.Print sRazonT
     Else
        Printer.Print sRazon
     End If
  End If
  ' Direccion
  If Rx!Domicilio <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Domicilio, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Domicilio, "X")
     If sDireT <> "0" Then
        Printer.Print sDireT
     Else
        Printer.Print sDire
     End If
  End If
  
  ' Localidad
  If Rx!Localidad <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Localidad, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Localidad, "X")
     If sLocalT <> "0" Then
        Printer.Print sLocalT
     Else
        Printer.Print sLocal
     End If
  End If
  ' Provincia
  If Rx!Provincia <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Provincia, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Provincia, "X")
     Printer.Print sProvi
  End If
  ' Telefono
  If Rx!Telefono <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Telefono, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Telefono, "X")
     Printer.Print sTel
  End If
  
  If Rx!CondIva <> "0" Then
     Printer.CurrentY = TraerXY(Rx!CondIva, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!CondIva, "X")
     Printer.Print sIva
     GoTo formapago
  End If
  Select Case sIva
         Case "Consumidor Final"
               If Rx!ConsFinal <> "0" Then
                  Printer.CurrentY = TraerXY(Rx!ConsFinal, "Y") + pId
                  Printer.CurrentX = TraerXY(Rx!ConsFinal, "X")
                  Printer.Print "X"
               End If
         Case "Responsable Inscripto"
              If Rx!ResInscrip <> "0" Then
                 Printer.CurrentY = TraerXY(Rx!ResInscrip, "Y") + pId
                 Printer.CurrentX = TraerXY(Rx!ResInscrip, "X")
                 Printer.Print "X"
              End If
         Case "Responsable no Inscripto"
           ' Responsable No Inscrp.
              If Rx!ResNInscrip <> "0" Then
                 Printer.CurrentY = TraerXY(Rx!ResNInscrip, "Y") + pId
                 Printer.CurrentX = TraerXY(Rx!ResNInscrip, "X")
                 Printer.Print "X"
              End If
         Case "Monotributo"
              If Rx!ResNInscrip <> "0" Then
                 Printer.CurrentY = TraerXY(Rx!ResNInscrip, "Y") + pId
                 Printer.CurrentX = TraerXY(Rx!ResNInscrip, "X")
                 Printer.Print "X"
              End If
         Case "Exento"
               If Rx!Exento <> "0" Then
                  Printer.CurrentY = TraerXY(Rx!Exento, "Y") + pId
                  Printer.CurrentX = TraerXY(Rx!Exento, "X")
                  Printer.Print "X"
               End If
  End Select
  ' Cuit
formapago:
  
  If Rx!Cuit <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Cuit, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Cuit, "X")
     Printer.Print sCuit
  End If
  ' Cond Pago
  If Rx!CondPago <> "0" Then
     Printer.CurrentY = TraerXY(Rx!CondPago, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!CondPago, "X")
     Printer.Print sCondi
     GoTo Transporte
  End If
  Select Case sCondi
         Case "CONTADO"
              If Rx!Contado <> "0" Then
                 Printer.CurrentY = TraerXY(Rx!Contado, "Y") + pId
                 Printer.CurrentX = TraerXY(Rx!Contado, "X")
                 Printer.Print "X"
               End If
         Case "CUENTA CORRIENTE"
              If Rx!ctacte <> "0" Then
                 Printer.CurrentY = TraerXY(Rx!ctacte, "Y") + pId
                 Printer.CurrentX = TraerXY(Rx!ctacte, "X")
                 Printer.Print "X"
              End If
  End Select
  
Transporte:
  ' Transporte
  If Rx!TTransporte <> "0" Then
     Printer.CurrentY = TraerXY(Rx!TTransporte, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!TTransporte, "X")
     Printer.Print cRsl.TraerValorDeUnCampo("RegTransporte", "Transporte", "Movimiento=" & nMovi)
  End If
  ' Direccion Transporte
  If Rx!TDomicilio <> "0" Then
     Printer.CurrentY = TraerXY(Rx!TDomicilio, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!TDomicilio, "X")
     Printer.Print cRsl.TraerValorDeUnCampo("RegTransporte", "Domicilio", "Movimiento=" & nMovi)
  End If
  ' Cuit Transporte
  If Rx!TCuitTrans <> "0" Then
     Printer.CurrentY = TraerXY(Rx!TCuitTrans, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!TCuitTrans, "X")
     Printer.Print cRsl.TraerValorDeUnCampo("RegTransporte", "CuitTrans", "Movimiento=" & nMovi)
  End If
  ' Patente Camion
  If Rx!TPatenteCamion <> "0" Then
     Printer.CurrentY = TraerXY(Rx!TPatenteCamion, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!TPatenteCamion, "X")
     Printer.Print cRsl.TraerValorDeUnCampo("RegTransporte", "PatenteCamion", "Movimiento=" & nMovi)
  End If
  ' Patente Acoplado
  If Rx!TPatenteAcoplado <> "0" Then
     Printer.CurrentY = TraerXY(Rx!TPatenteAcoplado, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!TPatenteAcoplado, "X")
     Printer.Print cRsl.TraerValorDeUnCampo("RegTransporte", "PatenteAcoplado", "Movimiento=" & nMovi)
  End If

End Sub

Private Function PonerNumeros(pValor As Single) As Single
   Dim n As Byte
   n = Len(pValor)
   PonerNumeros = pValor + 10 - n
End Function

Private Function TraerXY(pValor As String, pCor As String) As Single
   Dim n As Byte
   n = InStr(1, pValor, "-")
   If pCor = "Y" Then
      TraerXY = CSng(Mid(pValor, 1, n - 1))
      Debug.Print TraerXY
   Else
      TraerXY = CSng(Mid(pValor, n + 1, 4))
   End If
End Function

Private Sub Modificar()
  On Error GoTo errHandler

  If sw = True Then
     HabilitarTodo True, True
     CmdOrden.Enabled = False
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
     sw1 = False
     GrabarTodo
     sw = True
     sw1 = True
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

