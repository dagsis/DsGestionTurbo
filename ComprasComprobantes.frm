VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprabantes de Compras"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   12795
   Begin VB.CommandButton CmdBotones 
      Caption         =   "M.Fechas"
      Height          =   750
      Index           =   8
      Left            =   11715
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Etiquetas"
      Height          =   750
      Index           =   7
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   4599
      Width           =   1005
   End
   Begin VB.TextBox TxtCotizacion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2790
      MaxLength       =   7
      TabIndex        =   8
      Text            =   "TxtCotizac"
      Top             =   960
      Width           =   795
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   1245
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   2805
   End
   Begin VB.TextBox TxtSucursal 
      Height          =   315
      Left            =   8850
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "0000"
      Top             =   75
      Width           =   495
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   9570
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "00000000"
      Top             =   75
      Width           =   825
   End
   Begin VB.CommandButton CmdClientes 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   6525
      TabIndex        =   6
      Top             =   555
      Width           =   735
   End
   Begin VB.TextBox TxtDescuento 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1065
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "TxtDe"
      Top             =   960
      Width           =   810
   End
   Begin VB.ComboBox CmbCaja 
      Height          =   315
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   2010
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
      Height          =   4965
      Left            =   120
      TabIndex        =   18
      Top             =   1335
      Width           =   10395
      Begin VB.CheckBox ChkActualiza 
         Caption         =   "Actualizo P.Ventas"
         Height          =   225
         Left            =   105
         TabIndex        =   70
         Top             =   4695
         Width           =   1740
      End
      Begin VB.TextBox TxtDesc 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   930
         MaxLength       =   10
         TabIndex        =   31
         Text            =   "TxtDesc"
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox TxtLeyenda 
         Height          =   285
         Left            =   1920
         MaxLength       =   13
         TabIndex        =   24
         Text            =   "TxtDetalle"
         Top             =   2745
         Width           =   2475
      End
      Begin VB.ComboBox CmbCuenta 
         Height          =   315
         Left            =   4875
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3555
         Width           =   2700
      End
      Begin VB.TextBox TxtIBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   64
         Text            =   "TxtIBru"
         Top             =   3735
         Width           =   1050
      End
      Begin VB.TextBox TxtPIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   63
         Text            =   "TxtPIva"
         Top             =   3420
         Width           =   1050
      End
      Begin VB.TextBox TxtNoGrav 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9390
         MaxLength       =   10
         TabIndex        =   62
         Text            =   "TxtNoGrav"
         Top             =   3075
         Width           =   915
      End
      Begin VB.ComboBox CmbDeposito 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3150
         Width           =   2070
      End
      Begin VB.ComboBox CmbUnidad 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3150
         Width           =   1740
      End
      Begin VB.ComboBox CmbImpuesto 
         Height          =   315
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3555
         Width           =   2805
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   735
         MaxLength       =   25
         TabIndex        =   22
         Text            =   "TxtProducto"
         Top             =   2370
         Width           =   1080
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   6825
         TabIndex        =   35
         Top             =   2385
         Width           =   705
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   930
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "TxtCa"
         Top             =   3555
         Width           =   930
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   930
         MaxLength       =   10
         TabIndex        =   30
         Text            =   "TxtPrecio"
         Top             =   3930
         Width           =   1095
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3915
         Width           =   870
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   23
         Text            =   "TxtDetalle"
         Top             =   2385
         Width           =   4890
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3915
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   3975
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3915
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3915
         Width           =   870
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2130
         Left            =   165
         TabIndex        =   21
         Top             =   195
         Width           =   10125
         _ExtentX        =   17859
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
      Begin VB.Label Label15 
         Caption         =   "Desc. :"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   4320
         Width           =   750
      End
      Begin VB.Label Label8 
         Caption         =   "Leyenda Etiqueta :"
         Height          =   180
         Left            =   165
         TabIndex        =   66
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "I.Bru. Compras  :"
         Height          =   195
         Left            =   8070
         TabIndex        =   61
         Top             =   3735
         Width           =   1395
      End
      Begin VB.Label Label23 
         Caption         =   "P.Iva Compras :"
         Height          =   210
         Left            =   8070
         TabIndex        =   60
         Top             =   3420
         Width           =   1230
      End
      Begin VB.Label Label22 
         Caption         =   "Gastos No Grav. :"
         Height          =   195
         Left            =   8070
         TabIndex        =   59
         Top             =   3075
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   "Depósito :"
         Height          =   195
         Left            =   135
         TabIndex        =   58
         Top             =   3150
         Width           =   870
      End
      Begin VB.Label Label21 
         Caption         =   "U.De Medida :"
         Height          =   225
         Left            =   3015
         TabIndex        =   57
         Top             =   3150
         Width           =   1140
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   120
         TabIndex        =   45
         Top             =   2385
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   150
         TabIndex        =   44
         Top             =   3540
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Precio :"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   3900
         Width           =   750
      End
      Begin VB.Label Label10 
         Caption         =   "Neto :"
         Height          =   255
         Left            =   8070
         TabIndex        =   42
         Top             =   2385
         Width           =   525
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
         Left            =   8580
         TabIndex        =   41
         Top             =   2385
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "Desc.:"
         Height          =   225
         Left            =   8070
         TabIndex        =   40
         Top             =   2715
         Width           =   495
      End
      Begin VB.Label LblBonificacion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblBonificacion"
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
         Left            =   9180
         TabIndex        =   39
         Top             =   2715
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "I.V.A. :"
         Height          =   180
         Left            =   8070
         TabIndex        =   38
         Top             =   4095
         Width           =   660
      End
      Begin VB.Label Label16 
         Caption         =   "Total :"
         Height          =   225
         Left            =   8070
         TabIndex        =   37
         Top             =   4470
         Width           =   495
      End
      Begin VB.Label LblIva1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblIva1"
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
         Left            =   8835
         TabIndex        =   36
         Top             =   4095
         Width           =   1455
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
         Left            =   8730
         TabIndex        =   32
         Top             =   4470
         Width           =   1545
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   750
      Index           =   0
      Left            =   10590
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   750
      Index           =   1
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1673
      Width           =   1005
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
      Height          =   585
      Index           =   2
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2487
      Width           =   1005
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
      Height          =   585
      Index           =   3
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3136
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Aplicacion"
      Height          =   750
      Index           =   4
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3785
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   750
      Index           =   5
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1005
   End
   Begin VB.TextBox TxtCliente 
      Height          =   285
      Left            =   1050
      MaxLength       =   15
      TabIndex        =   5
      Text            =   "TxtCliente"
      Top             =   555
      Width           =   795
   End
   Begin VB.ComboBox CmbFormaPago 
      Height          =   315
      Left            =   7185
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   960
      Width           =   3210
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Anular"
      Height          =   750
      Index           =   6
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   859
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
      Format          =   135069697
      CurrentDate     =   36783
   End
   Begin MSComCtl2.DTPicker DtpLIva 
      Height          =   315
      Left            =   6735
      TabIndex        =   2
      Top             =   75
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   135069699
      CurrentDate     =   36783
   End
   Begin VB.Label Label20 
      Caption         =   "F.LIva:"
      Height          =   210
      Left            =   6120
      TabIndex        =   69
      Top             =   75
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Cotización :"
      Height          =   195
      Left            =   1935
      TabIndex        =   65
      Top             =   975
      Width           =   960
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   4140
      TabIndex        =   56
      Top             =   75
      Width           =   540
   End
   Begin VB.Label Label18 
      Caption         =   "Suc/Nº :"
      Height          =   165
      Left            =   8115
      TabIndex        =   55
      Top             =   105
      Width           =   660
   End
   Begin VB.Label Label17 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9390
      TabIndex        =   54
      Top             =   -30
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   120
      TabIndex        =   53
      Top             =   75
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "Proveedor :"
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   555
      Width           =   840
   End
   Begin VB.Label Label5 
      Caption         =   "Iva :"
      Height          =   225
      Left            =   7350
      TabIndex        =   51
      Top             =   555
      Width           =   360
   End
   Begin VB.Label Label7 
      Caption         =   "Descuento :"
      Height          =   240
      Left            =   105
      TabIndex        =   50
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Caja :"
      Height          =   240
      Left            =   3675
      TabIndex        =   49
      Top             =   975
      Width           =   480
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
      Left            =   1860
      TabIndex        =   48
      Top             =   555
      Width           =   4620
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
      Left            =   7770
      TabIndex        =   47
      Top             =   555
      Width           =   2610
   End
   Begin VB.Label Label30 
      Caption         =   "C.de Pago :"
      Height          =   210
      Left            =   6255
      TabIndex        =   46
      Top             =   975
      Width           =   885
   End
End
Attribute VB_Name = "ComprasComprobantes"
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
Public RsC As ADODB.Recordset
Dim cRsl As ClsLectura, Rsd1 As ADODB.Recordset
Dim bRemi As Boolean, bActualizo As Byte
Dim rRem As ADODB.Recordset

Private Sub CmbCaja_Click()
  VerEstado
  VerCabecera
End Sub
Private Sub CmbComprobante_Click()
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler

  If sw = False Then
     Set cNum = New ClsComprobantesL
     Set cRsl = New ClsLectura
     nNum = Format(cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
     nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
     TxtSucursal.text = nSuc
     TxtNumero.text = nNum
  End If
  Set cNum = Nothing
  VerEstado
  VerCabecera
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmbFormaPago_Click()
  VerEstado
  VerCabecera
End Sub

Private Sub CmbLista_Click()
  VerEstado
  VerCabecera
End Sub
Private Sub CmbVendedor_Click()
  VerEstado
  VerCabecera
End Sub

Private Sub CmbFormaPago_LostFocus()
  If CmdDetalle(0).Enabled = True Then
     CmdDetalle(0).SetFocus
  End If
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
              Buscar
         Case 2
              Anterior
         Case 3
              Siguiente
         Case 4
              Aplicacion
         Case 5
              Salir
         Case 6
              Anular
         Case 7
              Etiquetas
         Case 8
              MFechas
  End Select
End Sub

Private Sub MFechas()
   Dim cRc As ClsComprobantesE
   
   Set cRc = New ClsComprobantesE
   
   If CmdBotones(8).Caption = "M.Fechas" Then
      Botones False, False, False, False, False, False
      CmdBotones(8).Enabled = True
      CmdBotones(8).Caption = "Actualizar"
      DtpFecha.Enabled = True
      DtpLIva.Enabled = True
   Else
      Botones True, True, True, True, True, True
      CmdBotones(8).Caption = "M.Fechas"
      DtpFecha.Enabled = False
      DtpLIva.Enabled = False
      cRc.ActualizarFechasCompras DtpFecha.Value, DtpLIva.Value, nMovi
   End If
End Sub

Public Sub Etiquetas()
  Dim cRsl As ClsLectura, i As Integer
  
  Set cRsl = New ClsLectura
  Set RsC = New ADODB.Recordset
  
  RsC.Fields.Append "Descripcion", adVarChar, 50
  RsC.Fields.Append "DescripAdi", adVarChar, 50
  RsC.Fields.Append "Leyenda", adVarChar, 13
  RsC.Fields.Append "Fecha", adDate
  RsC.Fields.Append "Porcentaje", adCurrency
  RsC.Fields.Append "Codigo", adVarChar, 15
  RsC.Fields.Append "Barra", adVarChar, 15
  RsC.CursorType = adOpenKeyset
  RsC.LockType = adLockOptimistic
  RsC.Open

  Me.MousePointer = 11

  Rsd.MoveFirst
  Do While Not Rsd.EOF
     For i = 1 To Rsd!Cantidad
         RsC.AddNew
         RsC!Descripcion = Rsd!Descripcion
         RsC!DescripAdi = cRsl.TraerValorDeUnCampo("Productos", "Descripcion1", "Producto='" & Rsd!Producto & "'")
         RsC!Leyenda = "" & Rsd!Leyenda
         RsC!fecha = Date
         RsC!Porcentaje = cRsl.TraerValorDeUnCampo("Proveedores", "Porcentaje", "Proveedor='" & TxtCliente.text & "'")
         RsC!codigo = Rsd!Producto
         RsC!Barra = Rsd!Producto
         RsC.Update
     Next i
     Rsd.MoveNext
  Loop
  
  nImpr = 49
  FrmImpresor.Show
  Me.MousePointer = 0
  Set cRsl = Nothing
End Sub

Private Sub Aplicacion()
 Dim cTi As ClsComprobantesL, nCom As Integer
 
 Set cTi = New ClsComprobantesL
 
 nCom = CmbComprobante.ItemData(CmbComprobante.ListIndex)
 If CmdBotones(4).Caption = "Aplicación" Then
    If cTi.TraerTipoMovimiento(nCom) = 1 Or cTi.TraerTipoMovimiento(nCom) = 24 Or cTi.TraerTipoMovimiento(nCom) = 25 Then
       Dim cRp As ClsLectura
       Set cRp = New ClsLectura

       ComprasAplicacion.Show 1
       If cRp.TraerCantidad("Aplicaciones", "Id", "MoviCta=" & Rs!Movimiento) <> 0 Then
              CmdBotones(4).Caption = "DesAplicar"
       End If
    End If
 Else
  Dim cBor As ClsComprobantesE
      Set cBor = New ClsComprobantesE

      If MsgBox("Anula la Aplicación ?", 20, "Atención") = 6 Then
         cBor.AnulaAplicacion Rs!Movimiento
         CmdBotones(4).Caption = "Aplicación"
      End If
 End If
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
     TxtCliente.text = RsC!Proveedor
     LblCliente.Caption = RsC!RazonSocial
     LbLIva.Caption = cRiL.BuscarTipoIvaProv(RsC!Proveedor)
     VerRemitos
     CmbFormaPago.SetFocus
     CalcularTotales
     bCliente = True
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub VerRemitos()
  On Error GoTo errHandler

  Dim cRem As ClsComprobantesL, rRem As ADODB.Recordset
  
  Set cRem = New ClsComprobantesL
  Set rRem = New ADODB.Recordset
  bRemi = False
  Set rRem = cRem.TraerProveedorRemitos(TxtCliente.text)
  If rRem.RecordCount <> 0 Then
     nLlama = 2
     VentaRemitosPendientes.Show 1
     CrearRsDetalles
     LlenarDetallesRemito
     BotonDetalles True, True, True, False
  End If
Exit Sub

errHandler:
   On Error Resume Next
  
End Sub

Private Sub LlenarDetallesRemito()
  On Error GoTo errHandler

  Dim cRem As ClsComprobantesL, cPre As ClsProductoL, nCantRem As Double, nCantRemAux As Double
  Dim rAux As ADODB.Recordset, nCanti As Double, cImp As ClsProductoL
  Dim cLista As ClsComprobantesL, cCli As ClsClienteL, cRpl As ClsProductoL
  
  Set rAux = New ADODB.Recordset
  Set cPre = New ClsProductoL
  Set cImp = New ClsProductoL
  Set cLista = New ClsComprobantesL
  Set cCli = New ClsClienteL
  Set cRpl = New ClsProductoL
  
  Set rRem = VentaRemitosPendientes.rRemitos
  
  If Not rRem Is Nothing Then
     If rRem.RecordCount <> 0 Then
        rRem.MoveFirst
        Set cRem = New ClsComprobantesL
        sw1 = False
        Do While Not rRem.EOF
           Set rAux = cRem.TraerDetallesMovimiento(rRem!Movimiento)
           Do While Not rAux.EOF
              nCantRem = Round(cRem.TraerCantidadRemito(rAux!Producto, rRem!Movimiento, rAux!ID), 2)
              nCantRemAux = Round(rAux!Cantidad - nCantRem, 2)
              If nCantRemAux > 0 Or nCantRem = -1 Then
                 Rsd.AddNew
               '  Rsd!Id = 0
                 Rsd!Producto = rAux!Producto
                 Rsd!Descripcion = rRem!Descripcion + " " & rAux!Descripcion
                 nCanti = IIf(nCantRem = -1, rAux!Cantidad, rAux!Cantidad - nCantRem)
                 bRemi = False
                 If cRem.SumaCantidad(rAux!Comprobante) = True Then
                    bRemi = True
                 End If
              If cRsl.TraerValorDeUnCampo("Comprobantes", "TipoMovimiento", "Id=" & cRsl.TraerValorDeUnCampo("CabComprobantes", "Comprobante", "Movimiento=" & rRem!Movimiento)) <> 34 Then
                 Rsd!Cantidad = nCanti
              Else
                 Rsd!Cantidad = -(nCanti)
              End If
                 Rsd!Cuenta = cRsl.TraerValorDeUnCampo("OperacionesContables", "Cuenta", "Id=2")
                 Rsd!Deposito = rAux!Deposito ' rAuxD!Deposito
                 Rsd!Medida = rAux!Medida  'rAuxD!Medida
                 Rsd!Tasa = rAux!Tasa
                 Rsd!precio = rAux!precio
                 Rsd!PrecioUnitario = rAux!PrecioUnitario
                 Rsd!PrecioTotal = IIf(IsNull(rAux!PrecioTotal), 0, rAux!PrecioTotal)
                 Rsd!MoviRemi = rRem!Movimiento
                 Rsd.Update
              End If
             rAux.MoveNext
           Loop
           rRem.MoveNext
           Grid1.Enabled = True
        Loop
        CmbComprobante.Enabled = False
        CmdBotones(0).Enabled = True
        sw1 = True
    End If
  End If
Exit Sub

errHandler:
   MsgBox err.Description, "LlenarDetallesRemitos"
  
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
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL, sImp As String, cRx As ClsLectura
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  Set cRx = New ClsLectura
  nDat = 0
  
  nLista = 1

  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.text = RsP!Producto
     TxtDetalle.text = RsP!Descripcion
     TxtPrecio.text = Format(0, nCantDecimales)
     CmbUnidad.text = cRx.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & cRx.TraerValorDeUnCampo("Productos", "UmVenta", "Producto='" & TxtProducto.text & "'"))
     CmbDeposito.text = cRx.TraerValorDeUnCampo("Depositos", "Descripcion", "Deposito=" & cRx.TraerValorDeUnCampo("Productos", "Deposito", "Producto='" & TxtProducto.text & "'"))
     If cRx.TraerValorDeUnCampo("Comprobantes", "Iva", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 0 Then
        CmbImpuesto.text = "IVA EXENTO"
     Else
         sImp = cRpl.ProductoImpuestoDesc(TxtProducto.text)
         If sImp <> "" Then
            CmbImpuesto.text = sImp
         End If
     End If
     TxtLeyenda.SetFocus
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

Private Sub DTPFecha_LostFocus()
   DtpLIva.Value = DtpFecha.Value
End Sub

Private Sub DtpLIva_KeyDown(KeyCode As Integer, Shift As Integer)
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
  DtpFecha.Value = Date
  DtpLIva.Value = Date
  TxtDescuento.text = "0.00"
  LblCliente.Caption = ""
  TxtSucursal.text = &H80000012
  TxtNumero.ForeColor = &H80000012
  TxtSucursal.text = "0001"
  TxtNumero.text = "00000000"
  LbLIva.Caption = ""
  TxtCotizacion.text = "1.000"
  LblNeto.Caption = "0.00"
  LblBonificacion.Caption = "0.00"
  LblIva1.Caption = "0.00"
  lblTotal.Caption = "0.00"
  TxtNoGrav.text = "0.00"
  TxtPIva.text = "0.00"
  TxtIBruto.text = "0.00"
  CmbComprobante.ListIndex = -1
  CmbFormaPago.ListIndex = 0
  CmbCaja.ListIndex = -1
  CmbCuenta.ListIndex = -1
  ChkActualiza.Value = bActualizo

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
  Dim cRcL As ClsClienteL, cRiL As ClsClienteL, nSuc As String * 6, nNum As String * 6
  
  Dim cCompr As ClsLectura, cLista As ClsLectura, cVend As ClsLectura, cCaja As ClsLectura
  Dim cComp As ClsComprobantesL
  On Error GoTo errHandler
  
  Set cCompr = New ClsLectura
  Set cComp = New ClsComprobantesL
  Set cLista = New ClsLectura
  Set cVend = New ClsLectura
  Set cCaja = New ClsLectura
  Set cRcL = New ClsClienteL
  Set cRiL = New ClsClienteL
  
  CmbComprobante.text = cCompr.DatoCombo("Comprobantes", "Id", "Descripcion", Rs!Comprobante)
  DtpFecha.Value = Rs!fecha
  DtpLIva.Value = Rs!FechaIva
  TxtSucursal.ForeColor = &H80000012
  TxtNumero.ForeColor = &H80000012
  If cComp.ComprobanteAnulado(Rs!Movimiento) = True Then
     TxtSucursal.ForeColor = &H40C0&
     TxtNumero.ForeColor = &H40C0&
  End If
  TxtSucursal.text = Format(Rs!Sucursal, "0000")
  TxtNumero.text = Format(Rs!Numero, "00000000")
  nMovi = Rs!Movimiento
  TxtCliente.text = Rs!Cliente
  LblCliente.Caption = cRcL.BuscarNombreProveedor(Rs!Cliente)
  TxtCotizacion.text = Format(Rs!COTIZACION, "0.000")
  LbLIva.Caption = cRiL.BuscarTipoIvaProv(Rs!Cliente)
  TxtNoGrav.text = Format(Rs!NoGravados, "0.00")
  TxtPIva.text = Format(Rs!PIvaCompras, "0.00")
  TxtIBruto.text = Format(Rs!IBrutosCompras, "0.00")
  TxtDescuento = "0.00"
  If Rs!Neto <> 0 Then
     TxtDescuento.text = Format(Rs!Descuentos * 100 / Rs!Neto, "0.00")
  End If
  CmbCaja.text = cCaja.DatoCombo("Cajas", "Caja", "Descripcion", Rs!Caja)
  LblNeto.Caption = Format(Rs!Neto, "0.00")
  LblIva1.Caption = Format(Rs!Iva1, "0.00")
  LblBonificacion.Caption = Format(Rs!Descuentos, "0.00")
  lblTotal.Caption = Format(Rs!Debe + Rs!Haber, "0.00")
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
 nLlama = 1
 
 Botones True, False, False, False, False, True
 BotonDetalles False, False, False, False
 BotonBuscar False, False

 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Buscar", 0)
' CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
 CmdBotones(6).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(4).Picture = LoadResPicture("Valores", 0)
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
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "(TipoOperacion=8 or TipoOperacion=12) and Suc=" & nSucursal, "Descripcion"
  cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
  cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
  cRsl.CargaCombo CmbImpuesto, "Impuestos", "Impuesto", "Descripcion", "(Producto='1' OR Retencion='1' OR Percepcion='1' OR Interno='1') and Impuesto<>21"
  cRsl.CargaCombo CmbFormaPago, "CondVenta", "CondVta", "Descripcion", "tipo=6 and Dias=0"
  cRsl.CargaCombo CmbUnidad, "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbDeposito, "Depositos", "Deposito", "Descripcion", ""
  cRsl.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", "Grupo=5", "Descripcion"
  bActualizo = IIf(cRsl.RegPorDefecto("Precio") = True, 1, 0)

  Set cRsl = Nothing

Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bMod As Boolean, bBus As Boolean, _
  bVal As Boolean, bSal As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bVal
  CmdBotones(6).Enabled = bBorr
  If TxtNumero.ForeColor = &H40C0& Then
     CmdBotones(6).Enabled = False
  End If
  CmdBotones(2).Enabled = bMod
  CmdBotones(3).Enabled = bBus
  CmdBotones(4).Enabled = bVal
  CmdBotones(5).Enabled = bSal
  CmdBotones(8).Enabled = bBus
  CmdBotones(7).Enabled = bBorr
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
  If CmdBotones(0).Enabled = True Then
     CmdBotones(0).SetFocus
  End If
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
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Visible = False
      .Columns(1).Visible = False
      .Columns(2).Caption = "Producto"
      .Columns(2).Width = 1000
      .Columns(3).Caption = "Descripción"
      .Columns(3).Width = 4800
      .Columns(4).Visible = False
      .Columns(5).Caption = "Cant."
      .Columns(5).Width = 1000
      .Columns(5).Alignment = dbgRight
      .Columns(5).NumberFormat = nDecimalCant
      .Columns(6).Visible = False
      .Columns(7).Visible = False
      .Columns(8).Visible = False
      .Columns(9).Visible = False
      .Columns(10).Visible = False
      .Columns(11).Visible = False
      .Columns(12).Visible = False
      .Columns(13).Visible = False
      .Columns(14).Caption = "P.Uni."
      .Columns(14).Width = 800
      .Columns(14).NumberFormat = "0.00"
      .Columns(14).Alignment = dbgRight
      .Columns(15).Caption = "Total"
      .Columns(15).Width = 1000
      .Columns(15).NumberFormat = "0.00"
      .Columns(15).Alignment = dbgRight
      .Columns(16).Visible = False
      .Columns(17).Visible = False
 End With
End Sub
Private Sub NuevoRsD()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     CmdBotones(5).Enabled = False
     BotonBuscar False, True
     HabilitarTodo False, False
     HabilitarDetalles True
     LimpiarDetalles
     CmbDeposito.text = cRsl.TraerValorDeUnCampo("Depositos", "Descripcion", "Deposito=" & cRsl.RegPorDefecto("Deposito"))
     CmbUnidad.text = cRsl.DatoCombo("UnidadMedida", "Unidad", "Descripcion", cRsl.RegPorDefecto("Medida"))
     If cRsl.TraerValorDeUnCampo("Comprobantes", "Iva", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 0 Then
        CmbImpuesto.text = "IVA EXENTO"
     Else
        CmbImpuesto.text = cRsl.DatoCombo("Impuestos", "Impuesto", "Descripcion", cRsl.RegPorDefecto("Pais"))
     End If
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
     CmdBotones(0).Enabled = True
     CmdBotones(5).Enabled = True
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
        Botones True, False, False, False, False, True
        BotonDetalles True, True, True, False
'        CmdBotones(4).Enabled = HabiValores()
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
     Botones True, False, False, False, False, True
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
           Rsd.MoveLast
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

  Set crs = New ClsLectura
'  Rsd!Id = 0
  Rsd!Producto = TxtProducto.text
  Rsd!Descripcion = TxtDetalle.text
  Rsd!Leyenda = TxtLeyenda.text
  Rsd!Cantidad = TxtCantidad.text
  Rsd!Tasa = CmbImpuesto.ItemData(CmbImpuesto.ListIndex)
  Rsd!Cuenta = CmbCuenta.ItemData(CmbCuenta.ListIndex)
  Rsd!PDesc = 0
  Rsd!Descuento = 0
  If CmbDeposito.ListIndex <> -1 Then
     Rsd!Deposito = CmbDeposito.ItemData(CmbDeposito.ListIndex)
  Else
     Rsd!Deposito = 0
  End If
  If CmbUnidad.ListIndex <> -1 Then
     Rsd!Medida = CmbUnidad.ItemData(CmbUnidad.ListIndex)
  Else
     Rsd!Medida = 0
  End If
  TxtDesc.text = IIf(TxtDesc.text = "", "0.00", TxtDesc.text)
  Rsd!Descuento = TxtDesc.text
  Rsd!PrecioUnitario = TxtPrecio.text - TxtDesc.text
  Rsd!PrecioTotal = TxtCantidad.text * Rsd!PrecioUnitario
  Rsd!Impuesto = 0
  Rsd!MoviRemi = ChkActualiza.Value
  Rsd.Update
  VerEstado
  Set crs = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub LimpiarDetalles()
 Dim cRx As ClsLectura
 
 Set cRx = New ClsLectura
 

 CmbCuenta.text = cRx.TraerValorDeUnCampo("PlanDeCuenta", "Descripcion", "Id=" & cRx.TraerValorDeUnCampo("OperacionesContables", "Cuenta", "Id=2"))
 TxtProducto.text = ""
 TxtCantidad.text = Format(1, "0.00")
 TxtDetalle.text = ""
 TxtLeyenda.text = ""
 TxtPrecio.text = nCantDecimales
 TxtDesc.text = "0.00"
 If CmbComprobante.ListIndex <> -1 Then
    If cRx.TraerValorDeUnCampo("Comprobantes", "Iva", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 0 Then
       CmbImpuesto.text = "IVA EXENTO"
    Else
       CmbImpuesto.ListIndex = 1
    End If
 End If

 Set cRx = Nothing
End Sub
Private Sub LinkearDetalles()
  Dim cRsPl As ClsProductoL, cRa As ClsLectura
  On Error GoTo errHandler
  
  Set cRsPl = New ClsProductoL
  Set cRa = New ClsLectura
  
  TxtProducto.text = Rsd!Producto
  TxtCantidad.text = Format(Rsd!Cantidad, "0.00")
  TxtDetalle.text = Rsd!Descripcion
  CmbImpuesto.text = cRa.TraerValorDeUnCampo("Impuestos", "Descripcion", "Impuesto=" & Rsd!Tasa)
  CmbDeposito.ListIndex = -1
  CmbUnidad.ListIndex = -1
  If Rsd!Deposito <> 0 Then
     CmbDeposito.text = cRa.TraerValorDeUnCampo("Depositos", "Descripcion", "Deposito=" & Rsd!Deposito)
  End If
  If Rsd!Medida <> 0 Then
     CmbUnidad.text = cRa.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & Rsd!Medida)
  End If
  If Rsd!Cuenta <> 0 Then
     CmbCuenta.text = cRa.TraerValorDeUnCampo("PlanDeCuenta", "Descripcion", "Id=" & Rsd!Cuenta)
  End If
  TxtPrecio.text = Format(Rsd!PrecioUnitario, nCantDecimales)
  TxtDesc.text = Format(Rsd!Descuento, "#0.00")
  ChkActualiza.Value = Rsd!MoviRemi

Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Public Sub CalcularTotales()
  Dim cRx As ClsLectura, cDes As Currency, nTasa As Double, nNeto As Double, nDes As Double
  Dim nIva As Double, nReg As Variant, nOtros As Double, nNoGra As Double
  
  Set cRx = New ClsLectura
  On Error GoTo errHandler
  
  Set Rsd1 = cRx.RsVacio("DetallesComprobantes", "Id", "N")
  LblNeto.Caption = "0.00"
  LblIva1.Caption = "0.00"
  
  cDes = TxtDescuento.text
  nNeto = 0
  nDes = 0
  nIva = 0
  nNoGra = CCur(TxtNoGrav.text)
  
  If Rsd.RecordCount <> 0 Then
     nReg = Rsd.Bookmark
     Rsd.MoveFirst
     If cRx.TraerValorDeUnCampo("Proveedores", "TipoIva", "Proveedor='" & TxtCliente.text & "'") = 2 Then
        Do While Not Rsd.EOF
           Rsd1.AddNew
        '   Rsd1!Id = 0
           nTasa = cRx.TraerValorDeUnCampo("Impuestos", "Porcentaje", "Impuesto=" & Rsd!Tasa)
           Rsd1!Producto = Rsd!Producto
           Rsd1!Descripcion = Rsd!Descripcion
           Rsd1!Leyenda = Rsd!Leyenda
           Rsd1!Cantidad = Rsd!Cantidad
           If cDes <> 0 Then
              Rsd1!Descuento = (Rsd!PrecioUnitario * cDes / 100) * Rsd!Cantidad
              nDes = nDes + Rsd1!Descuento
           Else
              Rsd1!Descuento = IIf(IsNull(Rsd!Descuento), 0, Rsd!Descuento)
           End If
           Rsd1!Cuenta = Rsd!Cuenta
           Rsd1!Tasa = Rsd!Tasa
           Rsd1!Deposito = Rsd!Deposito
           Rsd1!Medida = Rsd!Medida
           Rsd1!PrecioUnitario = Rsd!PrecioUnitario
           Rsd1!Impuesto = ((Rsd1!PrecioUnitario - (Rsd1!Descuento / Rsd1!Cantidad)) * nTasa / 100) * Rsd!Cantidad
           nIva = nIva + Rsd1!Impuesto
           Rsd1!PrecioTotal = Rsd1!PrecioUnitario * Rsd!Cantidad ' + Rsd1!Descuento
           If cRx.TraerValorDeUnCampo("Impuestos", "Interno", "Impuesto=" & Rsd!Tasa) = 0 Then
              If cRx.TraerValorDeUnCampo("Impuestos", "Retencion", "Impuesto=" & Rsd!Tasa) = 0 Then
                 If cRx.TraerValorDeUnCampo("Impuestos", "Percepcion", "Impuesto=" & Rsd!Tasa) = 0 Then
                    nNeto = nNeto + Rsd1!PrecioTotal ' - (Rsd1!Descuento / Rsd1!Cantidad)
                 Else
                    nNoGra = nNoGra + Rsd1!PrecioTotal
                 End If
              Else
                 nNoGra = nNoGra + Rsd1!PrecioTotal
              End If
           Else
              nNoGra = nNoGra + Rsd1!PrecioTotal
           End If
           Rsd1!MoviRemi = Rsd!MoviRemi
           Rsd1.Update
           Rsd.MoveNext
        Loop
    Else
     Do While Not Rsd.EOF
        Rsd1.AddNew
      '  Rsd1!Id = 0
        Rsd1!Producto = Rsd!Producto
        Rsd1!Descripcion = Rsd!Descripcion
        Rsd1!Leyenda = Rsd!Leyenda
        Rsd1!Cantidad = Rsd!Cantidad
        If cDes <> 0 Then
           Rsd1!Descuento = (Rsd!PrecioUnitario * cDes / 100) * Rsd!Cantidad
           nDes = nDes + Rsd1!Descuento
        Else
           Rsd1!Descuento = Rsd!Descuento
        End If
        Rsd1!Tasa = Rsd!Tasa
        Rsd1!Cuenta = Rsd!Cuenta
        Rsd1!Deposito = Rsd!Deposito
        Rsd1!Medida = Rsd!Medida
        Rsd1!PrecioUnitario = Round(Rsd!PrecioUnitario, 2)
        Rsd1!Impuesto = 0
        nIva = nIva + Rsd1!Impuesto
        Rsd1!PrecioTotal = Rsd1!PrecioUnitario * Rsd!Cantidad ' + Rsd1!Descuento
        If cRx.TraerValorDeUnCampo("Impuestos", "Interno", "Impuesto=" & Rsd!Tasa) = 0 Then
           If cRx.TraerValorDeUnCampo("Impuestos", "Retencion", "Impuesto=" & Rsd!Tasa) = 0 Then
              If cRx.TraerValorDeUnCampo("Impuestos", "Percepcion", "Impuesto=" & Rsd!Tasa) = 0 Then
                 nNeto = nNeto + Rsd1!PrecioTotal ' - (Rsd1!Descuento / Rsd1!Cantidad)
              Else
                 nNoGra = nNoGra + Rsd1!PrecioTotal
              End If
           Else
              nNoGra = nNoGra + Rsd1!PrecioTotal
           End If
        Else
           nNoGra = nNoGra + Rsd1!PrecioTotal
        End If
        Rsd1!MoviRemi = Rsd!MoviRemi
        Rsd1.Update
        Rsd.MoveNext
    Loop
  End If
  sw1 = False
  Rsd.Bookmark = nReg
  sw1 = True
  TxtNoGrav.text = Format(nNoGra, "0.00")
  LblNeto.Caption = Format(nNeto, "0.00")
  LblBonificacion.Caption = Format(nDes, "0.00")
  LblIva1.Caption = Format(nIva, "0.00")
  nOtros = CCur(TxtNoGrav.text) + CCur(TxtPIva.text) + CCur(TxtIBruto.text)
  lblTotal.Caption = Format(nNeto - nDes + nIva + nOtros, "0.00")
 End If
Exit Sub

errHandler:
   ManejaErrores
 
End Sub

Private Sub TxtCantidad_GotFocus()
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.text)
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtCantidad) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCantidad_LostFocus()
  Dim cStock As ClsComprobantesL, cDepo As ClsProductoL
  Dim nDeposito As Integer
  
  Set cStock = New ClsComprobantesL
  Set cDepo = New ClsProductoL
  
  ' Si comprobante controla Stock, Verifica las Cantidades Disponibles
  If TxtProducto.text <> "" Then
     If cStock.ActualizaStock(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
        If cStock.DepositoVenta <> 0 Then
           nDeposito = cStock.DepositoVenta
        Else
           MsgBox "No hay Deposito de Venta Definido", vbCritical, "Atención"
           TxtCantidad.SetFocus
        End If
     End If
  End If
  If TxtCantidad.text = "" Then TxtCantidad.text = 1
  TxtCantidad.text = Format(TxtCantidad.text, "0.00")
End Sub

Private Sub TxtCliente_Change()
  VerEstado
  VerCabecera
End Sub

Private Sub TxtCliente_GotFocus()
  TxtCliente.SelStart = 0
  TxtCliente.SelLength = Len(TxtCliente.text)
End Sub

Private Sub TxtCliente_LostFocus()
  Dim cRcL As ClsClienteL, cRiL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set cRcL = New ClsClienteL
  Set cRiL = New ClsClienteL
  LblCliente.Caption = ""
  LbLIva.Caption = ""
  If TxtCliente.text <> "" Then
     If cRcL.BuscarNombreProveedor(TxtCliente.text) <> "" Then
        LblCliente.Caption = cRcL.BuscarNombreProveedor(TxtCliente.text)
        LbLIva.Caption = cRiL.BuscarTipoIvaProv(TxtCliente.text)
        CmbFormaPago.SetFocus
        CalcularTotales
        bCliente = True
     Else
        MsgBox "El Proveedor  no es el Correto", vbCritical, "Atención"
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
Private Sub TxtDescuento_GotFocus()
  TxtDescuento.SelStart = 0
  TxtDescuento.SelLength = Len(TxtDescuento.text)
End Sub
Private Sub TxtDescuento_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtDescuento) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtDescuento_LostFocus()
  If TxtDescuento.text = "" Then TxtDescuento.text = 0
  TxtDescuento.text = Format(TxtDescuento.text, "0.00")
  CalcularTotales
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
Private Sub TxtPrecio_GotFocus()
  TxtPrecio.SelStart = 0
  TxtPrecio.SelLength = Len(TxtPrecio.text)
End Sub
Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtPrecio) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtPrecio_LostFocus()
  If TxtPrecio.text = "" Then TxtPrecio.text = 0
  TxtPrecio.text = Format(TxtPrecio.text, nCantDecimales)
End Sub
Private Sub TxtProducto_LostFocus()
  Dim cRpl As ClsProductoL, sImp As String, cRx As ClsLectura
  
  On Error GoTo errHandler
  
  Set cRx = New ClsLectura
  Set cRpl = New ClsProductoL
  If TxtProducto.text <> "" Then
     If cRpl.BuscarNombreProducto(TxtProducto.text) <> "" Then
        TxtDetalle.text = cRpl.BuscarNombreProducto(TxtProducto.text)
        sImp = cRpl.ProductoImpuestoDesc(TxtProducto.text)
        CmbUnidad.text = cRx.TraerValorDeUnCampo("UnidadMedida", "Descripcion", "Unidad=" & cRx.TraerValorDeUnCampo("Productos", "UmVenta", "Producto='" & TxtProducto.text & "'"))
        CmbDeposito.text = cRx.TraerValorDeUnCampo("Depositos", "Descripcion", "Deposito=" & cRx.TraerValorDeUnCampo("Productos", "Deposito", "Producto='" & TxtProducto.text & "'"))
        If sImp <> "" Then
           CmbImpuesto.text = sImp
        End If
        If CmbComprobante.ListIndex <> -1 Then
           If cRx.TraerValorDeUnCampo("Comprobantes", "Iva", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)) = 0 Then
              CmbImpuesto.text = "IVA EXENTO"
           End If
        End If
        If TxtCantidad.Enabled = True Then
            TxtLeyenda.SetFocus
        End If
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
  TxtLeyenda.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  TxtDesc.Enabled = bHabi
  TxtPrecio.Enabled = bHabi
  CmbImpuesto.Enabled = bHabi
  CmbDeposito.Enabled = bHabi
  CmbUnidad.Enabled = bHabi
  CmbCuenta.Enabled = bHabi
  ChkActualiza.Enabled = bHabi
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler

  If sw = False And sw1 = True Then
     If Grid1.Row <> -1 Then
        Rsd.Bookmark = Grid1.Bookmark
        LinkearDetalles
     End If
  Else
      If CmdBotones(0).Caption <> "Grabar" Then
         Dim cRx As New ClsLectura
         Set cRx = New ClsLectura
         If Not Rsd.EOF Then
            Rsd.Bookmark = Grid1.Bookmark
            If Not IsNull(Rsd!Cuenta) Then
               LinkearDetalles
          ' CmbCuenta.Text = cRx.TraerValorDeUnCampo("PlanDeCuenta", "Descripcion", "Id=" & Rsd!Cuenta)
               BotonDetalles False, False, False, False
            End If
         End If
         Set cRx = Nothing
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
  Set Rs_Ven = cComp.TraerComprobantesCompras
  Botones True, False, False, False, False, True
  If Rs_Ven.BOF = False And Rs_Ven.EOF = False Then
     Rs_Ven.MoveFirst
     Set Rs = cRsl.TraerRS("CabezaTraerUno", Rs_Ven!ID, True)
     LinkearTexto
     LLenarDetalles
     nCas = 0
     Botones True, True, True, True, True, True
     Set cRsl = New ClsLectura
     If cRsl.TraerValorDeUnCampo("Aplicaciones", "Id", "Venta=" & Rs!Movimiento) <> 0 Then
        Botones True, False, True, True, True, True
     End If
     CmdBotones(4).Caption = "Aplicación"
     If cRsl.TraerCantidad("Aplicaciones", "Id", "MoviCta=" & Rs!Movimiento) <> 0 Then
        If CmbFormaPago.text <> "CONTADO" Then
           CmdBotones(4).Caption = "DesAplicar"
        End If
     End If
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
  If Rs_Ven.RecordCount <> 0 Then
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
     Set cRsl = New ClsLectura
     If cRsl.TraerValorDeUnCampo("Aplicaciones", "Id", "Venta=" & Rs!Movimiento) <> 0 Then
        Botones True, False, True, True, True, True
     End If
  End If
Exit Sub
  Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub Anterior()
 
  On Error GoTo errHandler
  If Rs_Ven.RecordCount <> 0 Then
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
        Set cRsl = New ClsLectura
        If cRsl.TraerValorDeUnCampo("Aplicaciones", "Id", "Venta=" & Rs!Movimiento) <> 0 Then
           Botones True, False, True, True, True, True
        End If
  End If
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
     Botones False, False, False, False, False, True
     BotonDetalles False, False, False, False
     BotonBuscar True, False
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(5).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(5).Caption = "Cancelar"
     CmdBotones(4).Caption = "Aplicacion"
     sw = False
     CrearRsVenta
     CrearRsDetalles
     CmbComprobante.SetFocus
     nMovi = 0
     nCas = 1
     CmbCaja.ListIndex = 0
     Set cRsl = New ClsLectura
  Else
     sw = True
     Botones False, False, False, False, False, False
     sw1 = True
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
     sw1 = True
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
  Dim cRx As ClsLectura
  
  Set cRx = New ClsLectura
  
  CmdBotones(0).Enabled = False
  CmdBotones(4).Enabled = False
  
  If Len(TxtCliente.text) <> 0 And Len(TxtNumero.text) <> 0 And Len(TxtSucursal.text) <> 0 And sw = False And CmbComprobante.ListIndex <> -1 Then
     If cRx.TraerCantidad("CabComprobantes", "Cliente", "Cliente='" & TxtCliente.text & "' AND Sucursal=" & TxtSucursal.text & " AND Numero=" & TxtNumero.text & " AND Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)) <> 0 Then
         MsgBox "Este Comprobante ya ha sido cargado", vbCritical, "Atención"
         TxtCliente.text = ""
         LblCliente.Caption = ""
         TxtCliente.SetFocus
         Exit Sub
     End If
 End If
 If Len(CmbComprobante.text) <> 0 And Len(TxtCliente.text) <> 0 _
     And Len(CmbCaja.text) <> 0 And Rsd.RecordCount <> 0 Then
     CmdBotones(0).Enabled = True
     CmdBotones(4).Enabled = True
  Else
     If Len(CmbComprobante.text) <> 0 And Len(TxtCliente.text) <> 0 _
         And Len(CmbCaja.text) <> 0 And Rsd.RecordCount <> 0 Then
        CmdBotones(4).Enabled = True
     End If
  End If
  Set cRx = Nothing
End Sub
Private Sub VerCabecera()
  CmdDetalle(0).Enabled = False
  If Len(CmbComprobante.text) <> 0 And Len(TxtCliente.text) <> 0 _
     And Len(CmbCaja.text) <> 0 Then
     CmdDetalle(0).Enabled = True
  End If
End Sub
Private Sub GrabarTodo()
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cTipo As ClsComprobantesL
  Dim rCuen As ADODB.Recordset, rValo As ADODB.Recordset, rApli As ADODB.Recordset
  Dim rReci As ADODB.Recordset, rCant As ADODB.Recordset, rCaja As ADODB.Recordset
  Dim cVeri As ClsComprobantesL, cCuota As ClsComprobantesL, cDias As ClsComprobantesL, cStock As ClsComprobantesL
  Dim i As Byte, nCuotas As Byte, nDias As Integer, rValoAux As ADODB.Recordset, cRsl As ClsLectura
  ' declaracion para clientes que emiten pagare

  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cTipo = New ClsComprobantesL
  Set rCuen = New ADODB.Recordset
  Set rValo = New ADODB.Recordset
  Set rValoAux = New ADODB.Recordset
  Set rApli = New ADODB.Recordset
  Set cVeri = New ClsComprobantesL
  Set cCuota = New ClsComprobantesL
  Set cDias = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  Set rReci = New ADODB.Recordset
  Set rCant = New ADODB.Recordset
  Set rCaja = New ADODB.Recordset
  Set cRsl = New ClsLectura
  
  Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
  Set rValo = cRsl.RsVacio("Valores", "Id", "N")
  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
  Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
  Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
    
  ' Grabar Cabecera Venta
    
  RsCv.AddNew
'  RsCv!Id = 0
  RsCv!Suc = nSucursal
  RsCv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
  RsCv!Movimiento = cNum.TraerUltimoMovimiento
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DtpFecha.Value
  RsCv!FechaIva = DtpLIva.Value
  RsCv!Numero = TxtNumero.text
  RsCv!Sucursal = TxtSucursal.text
  RsCv!Cliente = TxtCliente.text
  RsCv!Lista = 0
  RsCv!Vendedor = 0
  RsCv!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
  RsCv!CondVenta = CmbFormaPago.ItemData(CmbFormaPago.ListIndex)
  RsCv!Neto = CCur(LblNeto.Caption)
  RsCv!Iva1 = CCur(LblIva1.Caption)
  RsCv!Iva2 = CCur(TxtDescuento.text)
  RsCv!COTIZACION = CCur(TxtCotizacion.text)
  RsCv!Descuentos = CCur(LblBonificacion.Caption)
  RsCv!Financiacion = 0
  RsCv!NoGravados = CCur(TxtNoGrav.text)
  RsCv!PIvaCompras = CCur(TxtPIva.text)
  RsCv!IBrutosCompras = CCur(TxtIBruto.text)
  Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
         Case 22, 11, 24
              RsCv!Debe = CCur(lblTotal.Caption)
              RsCv!Haber = 0
         Case 14, 12
              RsCv!Debe = 0
              RsCv!Haber = CCur(lblTotal.Caption)
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = ""
  RsCv!Anulado = 0
  RsCv.Update
  ' Detalle Venta ya esta Armado en el recordset Rsd
  ' Forma de Pagos,Valores,Aplicaciones,Cantidades,Caja
  Set rValoAux = ComprasValores.RsV
  If cStock.ActualizaCtaCte(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
     Select Case cRsl.TraerValorDeUnCampo("CondVenta", "TipoCuota", "CondVta=" & CmbFormaPago.ItemData(CmbFormaPago.ListIndex))
            Case 1
                 nDias = -(cTipo.CantDias(CmbFormaPago.ItemData(CmbFormaPago.ListIndex)))
            Case 2
                 nDias = cTipo.CantDias(CmbFormaPago.ItemData(CmbFormaPago.ListIndex)) + 29
            Case 3
                 nDias = 0
            Case 4
                 nDias = 0
     End Select
     Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
            Case 1, 11, 24
                 If cTipo.TraerTipoCondicion(CmbFormaPago.ItemData(CmbFormaPago.ListIndex)) = 6 Then
                     For i = 1 To cTipo.CantCuotas(CmbFormaPago.ItemData(CmbFormaPago.ListIndex))
                         rReci.AddNew
                      '   rReci!Id = 0
                         rReci!fecha = DtpFecha.Value + cTipo.CantDias(CmbFormaPago.ItemData(CmbFormaPago.ListIndex)) + nDias
                         nDias = nDias + cTipo.CantDias(CmbFormaPago.ItemData(CmbFormaPago.ListIndex))
                         rReci!Cliente = TxtCliente.text
                         rReci!Venta = RsCv!Movimiento
                         rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
                         rReci!Numero = TxtNumero.text
                         rReci!Sucursal = TxtSucursal.text
                         If cTipo.TraerTipoMovimiento(RsCv!Comprobante) = 8 Then
                            rReci!Cuota = i
                         Else
                            rReci!Cuota = 1000
                         End If
                         rReci!Debe = Format(lblTotal.Caption, nCantDecimales) / cTipo.CantCuotas(CmbFormaPago.ItemData(CmbFormaPago.ListIndex))
                         rReci!Haber = 0
                         rReci!Anulado = 0
                         rReci.Update
                     Next i
                 End If
            Case 14, 12
                 If cTipo.TraerTipoCondicion(CmbFormaPago.ItemData(CmbFormaPago.ListIndex)) = 6 Then
                     For i = 1 To cTipo.CantCuotas(CmbFormaPago.ItemData(CmbFormaPago.ListIndex))
                         rReci.AddNew
                       '  rReci!Id = 0
                         rReci!fecha = DtpFecha.Value + cTipo.CantDias(CmbFormaPago.ItemData(CmbFormaPago.ListIndex)) + nDias
                         nDias = nDias + cTipo.CantDias(CmbFormaPago.ItemData(CmbFormaPago.ListIndex))
                         rReci!Cliente = TxtCliente.text
                         rReci!Venta = RsCv!Movimiento
                         rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
                         rReci!Numero = TxtNumero.text
                         rReci!Sucursal = TxtSucursal.text
                         If cTipo.TraerTipoMovimiento(RsCv!Comprobante) = 1 Then
                            rReci!Cuota = i
                         Else
                            rReci!Cuota = 1001
                         End If
                         rReci!Debe = 0
                         rReci!Haber = Format(lblTotal.Caption, nCantDecimales) / cTipo.CantCuotas(CmbFormaPago.ItemData(CmbFormaPago.ListIndex))
                         rReci!Anulado = 0
                         rReci.Update
                     Next i
                 End If
    End Select
  End If
  
  ' Actualizo Stock
  If cStock.ActualizaStock(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
     Dim nCantidadA As Double, nCantidadB As Double, nTotConver As Double
     If bRemi = False Then
        Rsd.MoveFirst
        Do While Not Rsd.EOF
           rCant.AddNew
           rCant!Suc = nSucursal
           rCant!Producto = Rsd!Producto
           rCant!Deposito = Rsd!Deposito
           rCant!Estado = 1
           nCantidadA = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadA", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!Medida)
           nCantidadB = cRsl.TraerValorDeUnCampo("ProductoUnidades", "CantidadB", "Producto='" & Rsd!Producto & "' AND UnidadA=" & Rsd!Medida)
           nTotConver = nCantidadA * nCantidadB
           Select Case cTipo.QueMoviStock(CmbComprobante.ItemData(CmbComprobante.ListIndex))
                  Case "Suma"
                       rCant!CantidadDebe = Rsd!Cantidad * nTotConver
                       rCant!CantidadHaber = 0
                  Case "Resta"
                       rCant!CantidadDebe = 0
                       rCant!CantidadHaber = Rsd!Cantidad * nTotConver
           End Select
           rCant!Anulado = 0
           rCant.Update
           Rsd.MoveNext
        Loop
        Rsd.MoveFirst
     End If
  End If
  ' Actualizo Caja
  If cStock.ActualizaCaja(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
     If rValo.EOF = False Or rValo.BOF = False Then
        rValo.MoveFirst
        Do While Not rValo.EOF
           If cTipo.TraerTipoCondicion(rValo!CodPago) < 6 Then
              rCaja.AddNew
              rCaja!Suc = nSucursal
              rCaja!Movimiento = RsCv!Movimiento
              rCaja!Comprobante = cTipo.TraerTipoRecibo(RsCv!Comprobante)
              rCaja!Numero = cTipo.TraerUltimoNumeroVenta(RsCv!Comprobante)
              rCaja!Sucursal = cTipo.TraerSucursalNumero(rCaja!Comprobante)
              rCaja!fecha = Date
              rCaja!Hora = Time
              rCaja!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
              rCaja!formapago = rValo!formapago
              rCaja!Nombre = rValo!Nombre
              rCaja!Banco = rValo!Banco
              rCaja!NumeroCheque = rValo!NumeroCheque
              rCaja!FechaAcreditacion = rValo!FechaAcreditacion
              rCaja!NumeroTarjeta = rValo!NumeroTarjeta
              rCaja!FechaVencimiento = rValo!FechaVencimiento
              rCaja!Autorizacion = rValo!Autorizacion
              rCaja!Cupon = rValo!Cupon
              If cStock.SumaCaja(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
                 rCaja!Debe = rValo!importe
                 rCaja!Haber = 0
              Else
                 rCaja!Debe = 0
                 rCaja!Haber = rValo!importe
              End If
              rCaja!Abierta = "S"
              rCaja!Anulado = 0
              rCaja.Update
           End If
           rValo.MoveNext
        Loop
     End If
  End If
   
  cGrab.Agregar RsCv, Rsd1, rReci, rValo, rApli, rCant, rCaja, bRemi, rRem
  
 CrearRsDetalles
 Limpiar
 
 ArmaRsParaMovimiento
  
  ' Fin del proceso de facturacion fiscal
  
  Set cGrab = Nothing
  Set cNum = Nothing
  Set cTipo = Nothing
  Set cVeri = Nothing
  Set cCuota = Nothing
  Set cDias = Nothing
  Set cStock = Nothing
Exit Sub
 
Errores:
  MsgBox err.Description, vbCritical, "GrabarTodos"
End Sub
Private Sub Valores()
  nLlama = 1
  ComprasValores.Show 1
  Select Case nCas
         Case 0
              BotonBuscar False, False
              HabilitarTodo False, False
              BotonDetalles False, False, False, False
              Botones True, True, True, True, True, True
         Case 2
              HabilitarTodo False, False
              BotonDetalles False, False, False, False
              Botones True, False, False, False, True, True
         Case 3
              HabilitarTodo True, True
              BotonDetalles True, True, True, False
              Botones False, False, False, False, True, True
   End Select
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
  If Rs_Ven.RecordCount <> 0 Then
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
  nLlama = 3
  ComproBuscar.Show 1
  If ComproBuscar.nId <> 0 Then
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
     CmdBotones(5).Caption = "Salir"
     HabilitarTodo False, True
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
  Dim cBor As ClsComprobantesE, cRw As ClsLectura
  
  Set cBor = New ClsComprobantesE
  
  Dim nOpe As Integer, nTipo As Integer, cLer As ClsComprobantesL
   
  Set cLer = New ClsComprobantesL
  Set cRw = New ClsLectura

  nOpe = cLer.TraerTipoOperacion(CmbComprobante.ItemData(CmbComprobante.ListIndex))
  nTipo = cLer.TraerTipoMovimiento(CmbComprobante.ItemData(CmbComprobante.ListIndex))

  If cRw.TraerValorDeUnCampo("Aplicaciones", "Id", "Venta=" & Rs!Movimiento & " AND Anulado=0") <> 0 Or CmdBotones(4).Caption = "DesAplicar" Then
     Exit Sub
  End If

  If MsgBox("Anula el Comprobante Seleccionado ?", 20, "Atención") = 6 Then
     cBor.AnularComprobantes Rs!Movimiento, nOpe, nTipo
     Rs_Ven.Delete
     CrearRsDetalles
     Limpiar
     Botones True, False, False, False, False, True
     Anterior
  End If
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

Private Sub TxtNoGrav_GotFocus()
  TxtNoGrav.SelStart = 0
  TxtNoGrav.SelLength = Len(TxtNoGrav.text)
End Sub
Private Sub TxtNoGrav_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtNoGrav) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtNoGrav_LostFocus()
  If TxtNoGrav.text = "" Then TxtNoGrav.text = 0
  TxtNoGrav.text = Format(TxtNoGrav.text, "#0.000")
  CalcularTotales
End Sub

Private Sub TxtPIva_GotFocus()
  TxtPIva.SelStart = 0
  TxtPIva.SelLength = Len(TxtPIva.text)
End Sub
Private Sub TxtPIva_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtPIva) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtPIva_LostFocus()
  If TxtPIva.text = "" Then TxtPIva.text = 0
  TxtPIva.text = Format(TxtPIva.text, "#0.00")
  CalcularTotales
End Sub
Private Sub TxtIBruto_GotFocus()
  TxtIBruto.SelStart = 0
  TxtIBruto.SelLength = Len(TxtIBruto.text)
End Sub
Private Sub TxtIBruto_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtIBruto) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtIBruto_LostFocus()
  If TxtIBruto.text = "" Then TxtIBruto.text = 0
  TxtIBruto.text = Format(TxtIBruto.text, "#0.000")
  CalcularTotales
End Sub

Private Sub TxtCotizacion_GotFocus()
  TxtCotizacion.SelStart = 0
  TxtCotizacion.SelLength = Len(TxtCotizacion.text)
End Sub
Private Sub TxtCotizacion_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtCotizacion) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCotizacion_LostFocus()
  If TxtCotizacion.text = "" Then TxtPrecio.text = 1
  TxtCotizacion.text = Format(TxtCotizacion.text, "#0.000")
End Sub

Private Sub TxtDesc_GotFocus()
  TxtDesc.SelStart = 0
  TxtDesc.SelLength = Len(TxtDesc.text)
End Sub
Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtDesc) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtDesc_LostFocus()
  If TxtDesc.text = "" Then TxtDesc.text = 0
  TxtDesc.text = Format(TxtDesc.text, "#0.000")
End Sub

