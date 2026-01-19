VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form StockComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Stock"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   10995
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   690
      Index           =   6
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3120
      Width           =   1005
   End
   Begin VB.TextBox TxtMotivo 
      Height          =   285
      Left            =   1185
      MaxLength       =   80
      TabIndex        =   4
      Text            =   "TxtMotivo"
      Top             =   975
      Width           =   7485
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Modificar"
      Height          =   690
      Index           =   9
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1605
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
      TabIndex        =   35
      Top             =   2355
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
      TabIndex        =   34
      Top             =   105
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   690
      Index           =   4
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2355
      Width           =   1005
   End
   Begin VB.ComboBox CmbHasta 
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Width           =   3045
   End
   Begin VB.ComboBox CmbDesde 
      Height          =   315
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   3015
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   2535
   End
   Begin VB.Frame Chk1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   45
      TabIndex        =   21
      Top             =   1380
      Width           =   8640
      Begin VB.TextBox TxtPorc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2835
         TabIndex        =   11
         Text            =   "21.00"
         Top             =   3930
         Width           =   570
      End
      Begin VB.TextBox TxtIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2145
         TabIndex        =   10
         Text            =   "21.00"
         Top             =   3930
         Width           =   570
      End
      Begin VB.CheckBox ChkRep 
         Caption         =   "A. Representantes"
         Height          =   195
         Left            =   1980
         TabIndex        =   42
         Top             =   4305
         Width           =   2115
      End
      Begin VB.TextBox TxtVenta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "TxtImporte"
         Top             =   3930
         Width           =   1170
      End
      Begin VB.CheckBox ChkActualiza 
         Caption         =   "Actualizo P.Ventas"
         Height          =   225
         Left            =   165
         TabIndex        =   41
         Top             =   4305
         Width           =   1740
      End
      Begin VB.TextBox TxtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   930
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "TxtImporte"
         Top             =   3930
         Width           =   1155
      End
      Begin VB.ComboBox CmbUnidad 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3255
         Width           =   2880
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   975
         MaxLength       =   25
         TabIndex        =   5
         Text            =   "TxtProducto"
         Top             =   2940
         Width           =   1095
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   6990
         TabIndex        =   15
         Top             =   2940
         Width           =   705
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   945
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "TxtCa"
         Top             =   3600
         Width           =   930
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   4665
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3600
         Width           =   870
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   2115
         MaxLength       =   100
         TabIndex        =   6
         Text            =   "TxtDetalle"
         Top             =   2940
         Width           =   4800
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3600
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3600
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   5595
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3600
         Width           =   870
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2640
         Left            =   165
         TabIndex        =   24
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
      Begin VB.Label Label10 
         Caption         =   "Venta"
         Height          =   180
         Left            =   3660
         TabIndex        =   45
         Top             =   3630
         Width           =   570
      End
      Begin VB.Label Label7 
         Caption         =   "Porc."
         Height          =   180
         Left            =   2865
         TabIndex        =   44
         Top             =   3630
         Width           =   390
      End
      Begin VB.Label Label4 
         Caption         =   "I."
         Height          =   180
         Left            =   2280
         TabIndex        =   43
         Top             =   3630
         Width           =   180
      End
      Begin VB.Label Label9 
         Caption         =   "Importe :"
         Height          =   255
         Left            =   165
         TabIndex        =   39
         Top             =   3945
         Width           =   780
      End
      Begin VB.Label Label14 
         Caption         =   "U.Medida :"
         Height          =   195
         Left            =   165
         TabIndex        =   37
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   165
         TabIndex        =   26
         Top             =   2940
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   165
         TabIndex        =   25
         Top             =   3600
         Width           =   780
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   690
      Index           =   0
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   105
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Anular"
      Height          =   690
      Index           =   1
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   855
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
      TabIndex        =   18
      Top             =   855
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
      TabIndex        =   17
      Top             =   1605
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   690
      Index           =   5
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5310
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   315
      Left            =   4710
      TabIndex        =   1
      Top             =   75
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   119209985
      CurrentDate     =   36783
   End
   Begin VB.Label Label8 
      Caption         =   "Observación :"
      Height          =   165
      Left            =   90
      TabIndex        =   38
      Top             =   990
      Width           =   990
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta :"
      Height          =   225
      Left            =   4620
      TabIndex        =   32
      Top             =   540
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "Desde :"
      Height          =   195
      Left            =   90
      TabIndex        =   31
      Top             =   570
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   90
      TabIndex        =   30
      Top             =   75
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   240
      Left            =   6075
      TabIndex        =   29
      Top             =   75
      Width           =   765
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   4125
      TabIndex        =   28
      Top             =   75
      Width           =   975
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
      Left            =   6825
      TabIndex        =   27
      Top             =   90
      Width           =   1485
   End
End
Attribute VB_Name = "StockComprobantes"
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
Dim nReg As Variant, nLinea As Byte, nLin As Byte
Dim bStock As Boolean, nFactura As Long, bActualizo As Byte, sImpuesto As String

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
        If CmbComprobante.ListIndex <> -1 Then
           Set cNum = New ClsComprobantesL

          nNum = Format(cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
          nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
          

          nLinea = cRsl.TraerValorDeUnCampo("Impresion", "CantidadDetalle", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
          LblNumero.Caption = nSuc & "-" & nNum
          nTip = cNum.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))
          Select Case nTip
               Case 5
                    HabiCombo True, True
                    cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
                    CmbHasta.Clear
                    Set cRsl = Nothing
                    CrearRsDetalles
                    LimpiarDetalles
                    If CmbDesde.ListCount <> 0 Then
                       CmbDesde.ListIndex = 0
                       CmdDetalle(0).Enabled = True
                    End If
               Case 6
                    HabiCombo False, True
                    cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", ""
                  '  CmbHasta.Clear
                    CmbDesde.ListIndex = -1
                    CmbHasta.ListIndex = 0 ' cRsl.RegPorDefecto("Deposito")
                    CmdDetalle(0).Enabled = True
               Case 7
                    HabiCombo True, False
                    cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
                    CmbHasta.Clear
                    CmbHasta.ListIndex = -1
                    CmbDesde.ListIndex = 0
                    CrearRsDetalles
                    LimpiarDetalles
               Case 17
                    HabiCombo False, True
                    cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
                    CmbHasta.Clear
                    CmbDesde.ListIndex = -1
                    CmbHasta.ListIndex = 0
               Case 18
                    HabiCombo True, False
                    cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
                    CmbHasta.Clear
                    CmbHasta.ListIndex = -1
                    CmbDesde.ListIndex = 0 ' cRsl.RegPorDefecto("Deposito")
                    CrearRsDetalles
                    LimpiarDetalles
                    CmdDetalle(0).Enabled = False
        End Select
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
  Dim cLc As ClsLectura, cNum As ClsComprobantesL
  
  Set cNum = New ClsComprobantesL
  Set cLc = New ClsLectura
  
  VerEstado
  
  Set cNum = New ClsComprobantesL
  If CmbComprobante.ListIndex <> -1 Then
     BotonDetalles True, False, False, False
     nTip = cNum.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))
     If nTip = 5 And sw = False Then
        Dim nMov As Long
        If CmbHasta.ListIndex <> -1 Then
           nMov = cNum.TraerMoviDesp(CmbHasta.ItemData(CmbHasta.ListIndex), CmbDesde.ItemData(CmbDesde.ListIndex))
           If nMov <> 0 Then
              MsgBox "No Puede Transferir a este deposito mientras haya Productos sin Conciliar", vbCritical, "Atención"
              BotonDetalles False, False, False, False
           End If
        End If
     End If
  End If
   
'  VerCabecera
End Sub


Private Sub CmbUnidad_Click()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  If CmbUnidad.ListIndex <> -1 And TxtProducto.Text <> "" Then
     If cRsl.TraerValorDeUnCampo("ProductoUnidades", "UnidadA", "Producto='" & TxtProducto.Text & "' and UnidadA=" & CmbUnidad.ItemData(CmbUnidad.ListIndex)) = 0 Then
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
  End Select
End Sub

Private Sub Imprimir()
  Dim Hoja As Byte, cHoja As Byte, Pagina As Byte, i As Single, i1 As Single, h As Byte, Linea As Byte, nLinea As Byte
  Dim nUlti As Double, nPorIva As Single, nTotal As Currency

  Dim cRsl As ClsLectura, Rx As ADODB.Recordset, nReg As Variant
  
  Dim SubTotal As Currency, Impuesto As Currency, Subtotal1 As Currency, Iva1 As Currency, Iva2 As Currency, Total As Currency
  
  Set cRsl = New ClsLectura
  
 ' Set Rx = cRsl.TraerRsCondi("Impresion", "Id", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
  
   Set Rx = cRsl.TraerRsCondi("Impresion", "Id", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)

   Dim P As Printer, sImpre As String, nCantipage As Single
        
   sImpre = cRsl.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
   If sImpre = "0" Then
      MsgBox "Debe Definir Un Modelo de Impresora al Comprobante...", vbCritical, "Atención"
      Exit Sub
  End If
    
  For Each P In Printers
      If P.DeviceName = sImpre Then
         Set Printer = P
      '   Printer.FontName = "Courier New"
         Printer.ScaleMode = 7
         Printer.FontBold = cRsl.TraerValorDeUnCampo("Impresion", "Negrita", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Printer.FontSize = cRsl.TraerValorDeUnCampo("Impresion", "Tamaño", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Exit For
      End If
  Next
  Rsd.MoveFirst
  i1 = 0
  Pagina = 0
  Hoja = 1
  
  nCantipage = Rsd.RecordCount / Rx!CantidadDetalle
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
            Printer.Print cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & Rsd!Medida)
         End If

         Printer.CurrentY = i
         Printer.CurrentX = Rx!Detalle
         Printer.Print Rsd!Descripcion & " - " & cRsl.TraerValorDeUnCampo("Productos", "Descripcion1", "Producto='" & Rsd!Producto & "'")
         If Rsd!Cantidad <> 0 Then
            Printer.CurrentY = i
            Printer.CurrentX = Rx!Cantidad - Printer.TextWidth(Format(Rsd!Cantidad, nCantDecimales))
            Printer.Print Format(Rsd!Cantidad, nCantDecimales)
            nTotal = nTotal + Rsd!Cantidad
         End If
         If nLinea = Linea Then
            nLinea = 1
            If Hoja = 1 Then
               nReg = Rsd.Bookmark
            End If
               
            Printer.NewPage
            CabezaFactura i, i1, Rx
            i = Rx!PosicionDetalle + i1
          '  Exit Do
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
End Sub

Private Sub CabezaFactura(pI As Single, pId As Single, Rx As ADODB.Recordset)
  Dim cRsl As ClsLectura, cClie As ClsClienteL
  
  Dim rCli As ADODB.Recordset
  Set cRsl = New ClsLectura
  Set cClie = New ClsClienteL
  
  Dim sFech As Date, sNum As String, sCuen As String, sRazon As String, sDire As String, sLocal As String, sTel As String
  Dim sCuit As String, sIva As String, sCondi As String, sProvi As String
  
  Dim sRazonT As String, sDireT As String, sLocalT As String, sTelT As String
  Dim sCuitT As String, sIvaT As String, sCondiT As String, sProviT As String
    
  sFech = DtpFecha.Value

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
     Printer.Print cRsl.TraerValorDeUnCampo("Impresion", "Texto", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
  End If
  ' Cod Cliente
  If Rx!Cliente <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Cliente, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Cliente, "X")
     Printer.Print "Desde Deposito :" & CmbDesde.Text
  End If
  ' Razon Social
  If Rx!Señor <> "0" Then
     Printer.CurrentY = TraerXY(Rx!Señor, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!Señor, "X")
     Printer.Print "Hasta Deposito :" & CmbHasta.Text
  End If
  If Rx!TCuitTrans <> "0" Then
     Printer.CurrentY = TraerXY(Rx!TCuitTrans, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!TCuitTrans, "X")
     Printer.Print "========================================================================================="
     Printer.CurrentY = TraerXY(Rx!TPatenteCamion, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!TPatenteCamion, "X")
     Printer.Print "Cantidad         Codigo                          Descripción/Adicional"
     Printer.CurrentY = TraerXY(Rx!TPatenteAcoplado, "Y") + pId
     Printer.CurrentX = TraerXY(Rx!TPatenteAcoplado, "X")
     Printer.Print "========================================================================================="
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

Private Sub CmdProductos_Click()
  Dim RsP As ADODB.Recordset, nUni As Integer, sImp As String, cRpl As ClsProductoL
  
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  Set cRpl = New ClsProductoL
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.Text = RsP!Producto
     TxtDetalle.Text = RsP!Descripcion
     nUni = cRsl.TraerValorDeUnCampo("Productos", "UMCompra", "Producto='" & TxtProducto.Text & "'")
     CmbUnidad.Text = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & nUni)
     sImp = "IVA EXENTO"
     TxtVenta.Text = Format(cRpl.TraerPrecioImp(TxtProducto.Text, str(nListaStock)), "#0.00")
     TxtIva.Text = Format(cRpl.ProductoCompuesto(TxtProducto.Text), "#0.00")
     TxtPorc.Text = Format(cRsl.TraerValorDeUnCampo("Precios", "Porcentaje", "Producto='" & TxtProducto.Text & "'"), "0.00")

     TraerCantidadPrecio RsP!Producto
     TxtCantidad.SetFocus
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub TraerCantidadPrecio(pProducto As String)
  Dim cRi As ClsProductoL
  
  Set cRi = New ClsProductoL
  
  TxtImporte.Text = Format(cRi.TraerPrecioCompra(pProducto), nCantDecimales)
  If CmbDesde.ListIndex <> -1 Then
     TxtCantidad.Text = Format(cRi.DameCantidad(pProducto, CmbDesde.ItemData(CmbDesde.ListIndex)), nCantDecimales)
  End If
  
  Set cRi = Nothing
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
         ctl.Text = ""
      End If
  Next ctl
  DtpFecha.Value = Date
  TxtIva.Text = "0.00"
  TxtPorc.Text = "0.00"
  TxtVenta.Text = "0.00"
  TxtImporte.Text = "0.00"
  LblNumero.ForeColor = &H80000012
  LblNumero.Caption = "0000-00000000"
  CmbComprobante.ListIndex = -1
  CmbDesde.ListIndex = -1
  CmbHasta.ListIndex = -1
  ChkActualiza.Value = bActualizo
  ChkRep.Value = False
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
  
  cCompr.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  cCompr.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", ""

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
  DtpFecha.Value = Rs!fecha
  nNum = Format(Rs!Numero, "00000000")
  nSuc = Format(Rs!sucursal, "0000")
  LblNumero.ForeColor = &H80000012
  If cComp.ComprobanteAnulado(Rs!Movimiento) = True Then
     LblNumero.ForeColor = &H40C0&
  End If
  LblNumero.Caption = nSuc & "-" & nNum
  nMovi = Rs!Movimiento
  TxtMotivo.Text = "" & Rs!Motivo
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
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=2 and Suc=" & nSucursal
  cRsl.CargaCombo CmbDesde, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  cRsl.CargaCombo CmbHasta, "Depositos", "Deposito", "Descripcion", ""
  cRsl.CargaCombo CmbUnidad, "UnidadMedida", "Unidad", "Descripcion", ""
  bActualizo = IIf(cRsl.RegPorDefecto("Precio") = True, 1, 0)
  sImpuesto = cRsl.DatoCombo("Impuestos", "Impuesto", "Descripcion", cRsl.RegPorDefecto("Pais"))
  nListaStock = cRsl.RegPorDefecto("ListaPrecio")
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
  CmdBotones(6).Enabled = bVal
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
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
 
  If sw1 = True Then
     sw1 = False
     CmdBotones(5).Enabled = False
     CmdBotones(9).Enabled = False
     CmdBotones(0).Enabled = False
     BotonBuscar False, True
     HabilitarTodo False, False
     HabilitarDetalles True
     LimpiarDetalles
     BotonDetalles True, False, False, True
     CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(0).Caption = "Grabar"
     CmbUnidad.Text = cRsl.DatoCombo("UnidadMedida", "Unidad", "Descripcion", cRsl.RegPorDefecto("Medida"))
     TxtCantidad.Text = "0.00"
     TxtProducto.SetFocus
  Else
     If Len(TxtDetalle.Text) <> 0 Then
        CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
        CmdDetalle(0).Caption = "Agregar"
        Rsd.AddNew
        GrabarRsd
        BotonDetalles True, True, True, False
        BotonBuscar True, False
        sw1 = True
        HabilitarTodo True, True
        HabilitarDetalles False
        If nTip <> 5 Then
           CmbDesde.Enabled = False
        End If
        If Rsd.RecordCount < nLinea Then
           BotonDetalles True, True, True, False
        Else
           BotonDetalles False, True, True, False
        End If
        CmbComprobante.Enabled = False
        If CmdBotones(0).Caption = "Grabar" Then
           Botones True, False, False, False, False, True, False
        Else
           Botones False, False, True, False, False, True, True
        End If
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
     If Rsd.RecordCount <> 0 Then
        Rsd.MovePrevious
        If Rsd.BOF Then
           Rsd.MoveFirst
        End If
        LinkearDetalles
        Grid1.Bookmark = Rsd.Bookmark
        BotonDetalles True, True, True, False
'        CmdBotones(4).Enabled = HabiValores()
        If CmdBotones(0).Caption = "Grabar" Then
           CmdBotones(0).Enabled = True
        Else
           CmdBotones(9).Enabled = True
        End If
     Else
        If CmdBotones(9).Caption = "Grabar" Then
             Botones False, False, False, False, False, True, False
        Else
           Botones False, False, False, False, False, True, False
        End If
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
     CmdBotones(5).Enabled = False
     CmdBotones(9).Enabled = False
     CmdBotones(0).Enabled = False
     BotonDetalles False, False, True, True
     CmdDetalle(2).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(2).Caption = "Grabar"
     CmdBotones(9).Enabled = False
     CmdBotones(5).Enabled = False
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
     If Rsd.RecordCount = nLinea Then
        BotonDetalles False, True, True, False
     End If
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
        If Rsd.RecordCount = nLinea Then
           BotonDetalles False, True, True, False
        End If
        HabilitarTodo True, True
        HabilitarDetalles False
        BotonBuscar True, False
        LimpiarDetalles
        If Rsd.EOF = False Or Rsd.BOF = False Then
           BotonDetalles True, True, True, False
           If Rsd.RecordCount <> 0 Then
              LinkearDetalles
           End If
        End If
        If Rsd.RecordCount <> 0 Then
            Botones True, False, False, False, False, True, False
        End If
        If CmdBotones(9).Caption = "Grabar" Then
           Botones False, False, True, False, False, True, True
           BotonDetalles True, True, True, False
           CmdBotones(9).Picture = LoadResPicture("Grabar", 0)
           CmdBotones(9).Caption = "Grabar"
'           CmdBotones(5).Picture = LoadResPicture("Cancelar", 0)
'           CmdBotones(5).Caption = "Cancelar"
        End If
        sw1 = True
     Else
        CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
        CmdDetalle(2).Caption = "Actualizar"
        BotonDetalles True, True, True, False
        If Rsd.RecordCount = nLinea Then
           BotonDetalles False, True, True, False
        End If
        HabilitarTodo True, True
        HabilitarDetalles False
        BotonBuscar True, False
        LinkearDetalles
        If CmdBotones(9).Caption = "Grabar" Then
           CmdBotones(9).Enabled = True
        Else
           CmdBotones(0).Enabled = True
        End If
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
'  Rsd!Id = 0
  Rsd!Producto = TxtProducto.Text
  Rsd!Descripcion = TxtDetalle.Text
  Rsd!Cantidad = TxtCantidad.Text
  Rsd!Tasa = 0
  Rsd!Cuenta = 0
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
  If ChkRep.Value = 1 Then
     Rsd!Cuenta = cRsl.TraerValorDeUnCampo("Productos", "Representante", "Producto='" & TxtProducto.Text & "'")
  End If
  Rsd!Impuesto = TxtImporte.Text * (cRsl.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & Rsd!Tasa)) / 100
  Rsd!precio = TxtVenta.Text
  If cRsl.TraerValorDeUnCampo("Precios", "Precio", "Producto='" & TxtProducto.Text & "'") <> TxtVenta.Text Then
     Rsd!Tasa = 1
  End If
  Rsd!PrecioUnitario = TxtImporte.Text
  Rsd!MoviRemi = ChkActualiza.Value
  Rsd.Update
  VerEstado
  Set crs = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub LimpiarDetalles()
 
 TxtProducto.Text = ""
 TxtCantidad.Text = "1.00"
 TxtDetalle.Text = ""
 CmbUnidad.ListIndex = -1
 TxtVenta.Text = "0.00"
 TxtIva.Text = "0.00"
 TxtPorc.Text = "0.00"
 TxtImporte.Text = "0.00"
End Sub
Private Sub LinkearDetalles()
  Dim cRsPl As ClsProductoL
  On Error GoTo errHandler
  
  Set cRsPl = New ClsProductoL
  Set cRsl = New ClsLectura
  
  TxtProducto.Text = Rsd!Producto
  TxtCantidad.Text = Format(Rsd!Cantidad, "#0.00")
  TxtDetalle.Text = Rsd!Descripcion
  If Not IsNull(Rsd!Medida) Then
     CmbUnidad.Text = cRsl.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & Rsd!Medida)
  End If
  TxtImporte.Text = Format(Rsd!PrecioUnitario, "#0.000")
  TxtVenta.Text = Format(Rsd!precio, "#0.000")
  ChkActualiza.Value = Rsd!MoviRemi
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub TxtCantidad_GotFocus()
  Dim cCan As ClsComprobantesL
  Set cCan = New ClsComprobantesL
  
  If nTip = 5 Or nTip = 7 Then
     TxtCantidad.Text = cCan.TraerCantidad(TxtProducto.Text, CmbDesde.ItemData(CmbDesde.ListIndex))
  End If
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtCantidad) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCantidad_LostFocus()
  If TxtCantidad.Text = "" Then TxtCantidad.Text = 1
  TxtCantidad.Text = Format(TxtCantidad.Text, "0.00")
End Sub

Private Sub TxtImporte_GotFocus()
  TxtImporte.SelStart = 0
  TxtImporte.SelLength = Len(TxtImporte.Text)
End Sub

Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
    If SoloNumeroDecimalFinal(KeyAscii, TxtImporte) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtImporte_LostFocus()
  If TxtImporte.Text = "" Then TxtImporte.Text = 0
  TxtImporte.Text = Format(TxtImporte.Text, "0.00")
 ' CalcularVenta
End Sub

Private Sub TxtProducto_LostFocus()
  Dim cRpl As ClsProductoL, sProdu As String, rPrd As ADODB.Recordset
  Dim cRx As ClsLectura, nUni As Integer, sImp As String
  
  
  Set cRpl = New ClsProductoL
  Set cRx = New ClsLectura
  If TxtProducto.Text <> "" Then
     If cRpl.BuscarNombreProducto(TxtProducto.Text) <> "" Then
        TxtDetalle.Text = cRpl.BuscarNombreProducto(TxtProducto.Text)
        nUni = cRx.TraerValorDeUnCampo("Productos", "UMVenta", "Producto='" & TxtProducto.Text & "'")
        If nUni <> 0 Then
           CmbUnidad.Text = cRx.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & nUni)
           TxtVenta.Text = Format(cRpl.TraerPrecioImp(TxtProducto.Text, str(nLista)), "#0.00")
           TxtIva.Text = Format(cRpl.ProductoCompuesto(TxtProducto.Text), "#0.00")
           TxtPorc.Text = Format(cRx.TraerValorDeUnCampo("Precios", "Porcentaje", "Producto='" & TxtProducto.Text & "'"), "0.00")
           TraerCantidadPrecio TxtProducto.Text
           TxtCantidad.SetFocus
        Else
           TxtProducto.Text = ""
           TxtProducto.SetFocus
        End If
     Else
         TxtProducto.Text = ""
         TxtProducto.SetFocus
     End If
  End If
  Set cRpl = Nothing
  Set cRx = Nothing

End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtProducto.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  CmbUnidad.Enabled = bHabi
  TxtImporte.Enabled = bHabi
  TxtIva.Enabled = bHabi
  TxtPorc.Enabled = bHabi
  TxtVenta.Enabled = bHabi
  ChkActualiza.Enabled = bHabi
  ChkRep.Enabled = bHabi
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
  Dim cComp As ClsComprobantesL, cRsx As ClsLectura
  On Error GoTo errHandler

  Set cComp = New ClsComprobantesL
  Set cRsx = New ClsLectura
  Set Rs_Sto = cComp.TraerComprobantesStock
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
     Rsd!PDesc = rAuxD!PDesc
     Rsd!Descuento = rAuxD!Descuento
     Rsd!precio = rAuxD!precio
     Rsd!PrecioUnitario = rAuxD!PrecioUnitario
     Rsd!PrecioTotal = rAuxD!PrecioTotal
     Rsd!MoviRemi = rAuxD!MoviRemi
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
     nLin = 1
     nFactura = 0
     Set cRsl = New ClsLectura
     CmbComprobante.ListIndex = 0
  Else
     sw = True
     Botones False, False, False, False, False, False, False
     If pClave = True Then
        FrmClave.Show 1
        If bClave = True Then
           GrabarTodo
        Else
           MsgBox "No Tiene Permisos para Realizar Esta Operación", vbExclamation, "Atención"
        End If
     Else
        GrabarTodo
     End If
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
        sw = False
        LinkearTexto
        sw = True
        HabilitarTodo False, False
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
     If CmbComprobante.Text <> "TRANFERENCIA" Then
        CmdDetalle(0).Enabled = True
     End If
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
  RsCv!Numero = Right(LblNumero.Caption, 8)
  RsCv!sucursal = Left(LblNumero.Caption, 4)
  RsCv!Cliente = Null
  RsCv!Iva2 = 0
  RsCv!Lista = nFactura
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
  RsCv!factura = IIf(nFactura = 1, 1, 0)
  RsCv!Motivo = TxtMotivo.Text
  RsCv!Anulado = 0
  RsCv.Update
  ' Actualizo Stock
  
  Dim nCantidadA As Single, nCantidadB As Single, nTotConver As Single
  sw1 = False
  If bStock = False Then
     Rsd.MoveFirst
     nTip = cNum.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))
     If nTip <> 22 Then
        Do While Not Rsd.EOF
           rCant.AddNew
           rCant!Suc = nSucursal
           rCant!Movimiento = RsCv!Movimiento
           rCant!Producto = Rsd!Producto
           nCantidadA = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadA", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!Medida)
           nCantidadB = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadB", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!Medida)
           nTotConver = nCantidadA * nCantidadB
           Select Case nTip
                  Case 5
                       rCant!Estado = 0
                       rCant!Deposito = 0
                       rCant!CantidadDebe = Rsd!Cantidad * nTotConver
                       rCant!CantidadHaber = Rsd!Cantidad * nTotConver
                  Case 6, 17
                       rCant!Estado = 1
                       rCant!Deposito = CmbHasta.ItemData(CmbHasta.ListIndex)
                       rCant!CantidadDebe = Rsd!Cantidad * nTotConver
                       rCant!CantidadHaber = 0
                  Case 7, 18
                       rCant!Estado = 1
                       rCant!Deposito = CmbDesde.ItemData(CmbDesde.ListIndex)
                       rCant!CantidadHaber = Rsd!Cantidad * nTotConver
                       rCant!CantidadDebe = 0
           End Select
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
  sw1 = True

  Set cGrab = Nothing
  Set cNum = Nothing
  Set cTipo = Nothing
  Set cStock = Nothing

End Sub
Private Sub Buscar()
  nLlama = 2
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
     Botones True, False, False, False, False, True, True
     If Rs_Sto.RecordCount <> 0 Then
        Anterior
     Else
        LimpiarRecorset Rsd
     End If
  End If
End Sub

Private Sub Modificar()
  Set cRsl = New ClsLectura
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
     nLinea = cRsl.TraerValorDeUnCampo("Impresion", "CantidadDetalle", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
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

Private Sub TxtVenta_GotFocus()
  TxtVenta.SelStart = 0
  TxtVenta.SelLength = Len(TxtVenta.Text)
End Sub

Private Sub TxtVenta_KeyPress(KeyAscii As Integer)
    If SoloNumeroDecimalFinal(KeyAscii, TxtVenta) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtVenta_LostFocus()
  If TxtVenta.Text = "" Then TxtVenta.Text = 0
  TxtVenta.Text = Format(TxtVenta.Text, "0.00")
End Sub

Private Sub TxtIva_GotFocus()
  TxtIva.SelStart = 0
  TxtIva.SelLength = Len(TxtIva.Text)
End Sub

Private Sub TxtIva_KeyPress(KeyAscii As Integer)
    If SoloNumeroDecimalFinal(KeyAscii, TxtIva) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtIva_LostFocus()
  If TxtIva.Text = "" Then TxtIva.Text = 0
  TxtIva.Text = Format(TxtIva.Text, "0.00")
  If TxtIva.Text <> 0 Then
     CalcularVenta
  End If
End Sub

Private Sub TxtPorc_GotFocus()
  TxtPorc.SelStart = 0
  TxtPorc.SelLength = Len(TxtPorc.Text)
End Sub

Private Sub TxtPorc_KeyPress(KeyAscii As Integer)
   If SoloNumeroDecimalFinal(KeyAscii, TxtPorc) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtPorc_LostFocus()
  If TxtPorc.Text = "" Then TxtPorc.Text = 0
  TxtPorc.Text = Format(TxtPorc.Text, "0.00")
    db.Open sDb
    db.Execute "UPDATE PRECIOS SET Porcentaje=" & CDbl(TxtPorc.Text) & "  WHERE Producto='" & TxtProducto.Text & "'"
    db.Close
  If TxtPorc.Text <> 0 Then
     CalcularVenta
  End If
End Sub

Private Sub CalcularVenta()
  Dim nVenta As Double
  
  nVenta = TxtImporte.Text + (TxtImporte.Text * TxtIva.Text / 100)
  nVenta = nVenta + (nVenta * TxtPorc.Text / 100)
  nVenta = RedondearPrecios(nVenta)
  TxtVenta.Text = Format(nVenta, "#0.00")
  
  
End Sub
