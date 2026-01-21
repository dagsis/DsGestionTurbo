VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Venta"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "VentaComprobantes.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   11790
   Begin VB.CheckBox ChkA4 
      Caption         =   "Fact.A4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10890
      TabIndex        =   85
      Top             =   3885
      Width           =   780
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "E-Mail"
      Height          =   675
      Index           =   8
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   879
      Width           =   750
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Recibo"
      Height          =   675
      Index           =   7
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   75
      Width           =   750
   End
   Begin VB.ComboBox CmbCorredor 
      Height          =   315
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1395
      Width           =   2040
   End
   Begin VB.TextBox TxtVendedor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6825
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "TxtVend"
      Top             =   75
      Width           =   675
   End
   Begin VB.TextBox TxtReparto 
      Height          =   315
      Left            =   8400
      MaxLength       =   13
      TabIndex        =   2
      Text            =   "0000000000000"
      Top             =   75
      Width           =   1335
   End
   Begin VB.TextBox LblIva1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8400
      MaxLength       =   10
      TabIndex        =   74
      Text            =   "LblIva1"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ComboBox CmbVend 
      Height          =   315
      Left            =   4650
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   2085
   End
   Begin VB.ComboBox CmbEmpresa 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   75
      Width           =   2790
   End
   Begin VB.TextBox TxtCotizacion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3960
      MaxLength       =   7
      TabIndex        =   10
      Text            =   "TxtCotizac"
      Top             =   1395
      Width           =   840
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   675
      Index           =   3
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2487
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   675
      Index           =   0
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   75
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   675
      Index           =   2
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   1680
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   675
      Index           =   9
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6270
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Anular"
      Height          =   675
      Index           =   1
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   879
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   675
      Index           =   5
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3885
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Aplicación"
      Height          =   675
      Index           =   6
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   4674
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Texto"
      Height          =   450
      Index           =   4
      Left            =   9840
      TabIndex        =   62
      Top             =   3291
      Width           =   1005
   End
   Begin VB.ComboBox CmbCondPago 
      Height          =   315
      Left            =   3975
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1830
      Width           =   3285
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   8835
      MaxLength       =   8
      TabIndex        =   6
      Text            =   "00000000"
      Top             =   510
      Width           =   855
   End
   Begin VB.TextBox TxtSucursal 
      Height          =   315
      Left            =   8040
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "0000"
      Top             =   510
      Width           =   495
   End
   Begin VB.ComboBox CmbFormaPago 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1830
      Width           =   2070
   End
   Begin VB.TextBox TxtCliente 
      Height          =   285
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   7
      Text            =   "TxtCliente"
      Top             =   945
      Width           =   795
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
      Height          =   4830
      Left            =   75
      TabIndex        =   38
      Top             =   2145
      Width           =   9540
      Begin VB.TextBox TxtPDesc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2775
         MaxLength       =   5
         TabIndex        =   21
         Text            =   "TxtPD"
         Top             =   3180
         Width           =   870
      End
      Begin VB.ComboBox CmbCuenta 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3525
         Width           =   2700
      End
      Begin VB.TextBox TxtNoGrav 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8475
         MaxLength       =   10
         TabIndex        =   58
         Text            =   "TxtNoGrav"
         Top             =   3052
         Width           =   915
      End
      Begin VB.ComboBox CmbUnidad 
         Height          =   315
         Left            =   4785
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2775
         Width           =   1740
      End
      Begin VB.ComboBox CmbDeposito 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2780
         Width           =   2820
      End
      Begin VB.ComboBox CmbImpuesto 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3180
         Width           =   2235
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3900
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   3870
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3900
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   6465
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3900
         Width           =   870
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2055
         Left            =   165
         TabIndex        =   51
         Top             =   255
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3625
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
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   1920
         MaxLength       =   80
         TabIndex        =   16
         Text            =   "TxtDetalle"
         Top             =   2400
         Width           =   4680
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3900
         Width           =   870
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   690
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "TxtPrecio"
         Top             =   3525
         Width           =   1095
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "TxtCa"
         Top             =   3180
         Width           =   930
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   6630
         TabIndex        =   17
         Top             =   2400
         Width           =   705
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   720
         MaxLength       =   25
         TabIndex        =   15
         Text            =   "TxtProducto"
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label LblPercepcion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblPercepcion"
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
         Left            =   8295
         TabIndex        =   81
         Top             =   4095
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "Perc.IB:"
         Height          =   195
         Left            =   7500
         TabIndex        =   80
         Top             =   4125
         Width           =   630
      End
      Begin VB.Label Label31 
         Caption         =   "C.Contable :"
         Height          =   240
         Left            =   2805
         TabIndex        =   79
         Top             =   3525
         Width           =   915
      End
      Begin VB.Label Label29 
         Caption         =   "Iva :"
         Height          =   210
         Left            =   3795
         TabIndex        =   78
         Top             =   3180
         Width           =   390
      End
      Begin VB.Label Label28 
         Caption         =   "P.Desc.:"
         Height          =   240
         Left            =   1905
         TabIndex        =   77
         Top             =   3180
         Width           =   675
      End
      Begin VB.Label Label22 
         Caption         =   "No Grav. :"
         Height          =   195
         Left            =   7500
         TabIndex        =   59
         Top             =   3045
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "U.De Medida :"
         Height          =   225
         Left            =   3690
         TabIndex        =   57
         Top             =   2780
         Width           =   1140
      End
      Begin VB.Label Label18 
         Caption         =   "Depósito :"
         Height          =   195
         Left            =   90
         TabIndex        =   56
         Top             =   2780
         Width           =   870
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
         Left            =   8010
         TabIndex        =   50
         Top             =   4440
         Width           =   1380
      End
      Begin VB.Label Label16 
         Caption         =   "Total :"
         Height          =   225
         Left            =   7500
         TabIndex        =   49
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Iva 1:"
         Height          =   180
         Left            =   7500
         TabIndex        =   48
         Top             =   3750
         Width           =   420
      End
      Begin VB.Label LblFinanciacion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblFinanciacion"
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
         Left            =   8295
         TabIndex        =   47
         Top             =   3408
         Width           =   1095
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
         Left            =   8295
         TabIndex        =   46
         Top             =   2711
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Finac.:"
         Height          =   195
         Left            =   7500
         TabIndex        =   45
         Top             =   3405
         Width           =   510
      End
      Begin VB.Label Label19 
         Caption         =   "Bonif.:"
         Height          =   225
         Left            =   7500
         TabIndex        =   44
         Top             =   2715
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
         Left            =   7965
         TabIndex        =   43
         Top             =   2370
         Width           =   1425
      End
      Begin VB.Label Label10 
         Caption         =   "Neto :"
         Height          =   255
         Left            =   7500
         TabIndex        =   42
         Top             =   2370
         Width           =   525
      End
      Begin VB.Label Label14 
         Caption         =   "Precio :"
         Height          =   195
         Left            =   75
         TabIndex        =   41
         Top             =   3525
         Width           =   750
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   60
         TabIndex        =   40
         Top             =   3180
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   2400
         Width           =   735
      End
   End
   Begin VB.ComboBox CmbLista 
      Height          =   315
      Left            =   1065
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1395
      Width           =   2040
   End
   Begin VB.CommandButton CmdClientes 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   4980
      TabIndex        =   8
      Top             =   915
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   315
      Left            =   5745
      TabIndex        =   4
      Top             =   510
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   159580161
      CurrentDate     =   36783
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   465
      Width           =   3825
   End
   Begin MSComCtl2.DTPicker DtpVenc 
      Height          =   315
      Left            =   8280
      TabIndex        =   14
      Top             =   1830
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   159580161
      CurrentDate     =   36783
   End
   Begin VB.PictureBox Picture1 
      Height          =   570
      Left            =   6915
      ScaleHeight     =   510
      ScaleWidth      =   210
      TabIndex        =   84
      Top             =   5190
      Width           =   270
   End
   Begin VB.Label Label27 
      Caption         =   "Cantidad :"
      Height          =   255
      Left            =   2025
      TabIndex        =   76
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label26 
      Caption         =   "Reparto :"
      Height          =   240
      Left            =   7650
      TabIndex        =   75
      Top             =   75
      Width           =   675
   End
   Begin VB.Label Label25 
      Caption         =   "Vendedor:"
      Height          =   240
      Left            =   3900
      TabIndex        =   73
      Top             =   75
      Width           =   825
   End
   Begin VB.Label Label24 
      Caption         =   "Empresa :"
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
      Left            =   135
      TabIndex        =   72
      Top             =   75
      Width           =   885
   End
   Begin VB.Label Label23 
      Caption         =   "Cotización :"
      Height          =   195
      Left            =   3105
      TabIndex        =   70
      Top             =   1395
      Width           =   960
   End
   Begin VB.Label LblDescuento 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblDescuento"
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
      Left            =   5775
      TabIndex        =   61
      Top             =   1410
      Width           =   765
   End
   Begin VB.Label Label15 
      Caption         =   "F.de Pago :"
      Height          =   210
      Left            =   3150
      TabIndex        =   60
      Top             =   1845
      Width           =   885
   End
   Begin VB.Label Label17 
      Caption         =   "--"
      Height          =   195
      Left            =   8610
      TabIndex        =   55
      Top             =   495
      Width           =   150
   End
   Begin VB.Label Label30 
      Caption         =   "C.de Pago :"
      Height          =   210
      Left            =   135
      TabIndex        =   54
      Top             =   1845
      Width           =   885
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
      Left            =   6270
      TabIndex        =   37
      Top             =   930
      Width           =   3390
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
      TabIndex        =   36
      Top             =   945
      Width           =   3105
   End
   Begin VB.Label Label9 
      Caption         =   "F.Venc. :"
      Height          =   240
      Left            =   7500
      TabIndex        =   35
      Top             =   1845
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Corredor :"
      Height          =   225
      Left            =   6660
      TabIndex        =   34
      Top             =   1410
      Width           =   780
   End
   Begin VB.Label Label7 
      Caption         =   "Descuento :"
      Height          =   240
      Left            =   4860
      TabIndex        =   33
      Top             =   1395
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Lista :"
      Height          =   330
      Left            =   135
      TabIndex        =   32
      Top             =   1395
      Width           =   555
   End
   Begin VB.Label Label5 
      Caption         =   "Iva :"
      Height          =   225
      Left            =   5850
      TabIndex        =   31
      Top             =   945
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente :"
      Height          =   255
      Left            =   135
      TabIndex        =   30
      Top             =   945
      Width           =   660
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   5160
      TabIndex        =   29
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      Height          =   240
      Left            =   7155
      TabIndex        =   28
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   135
      TabIndex        =   27
      Top             =   510
      Width           =   1155
   End
End
Attribute VB_Name = "VentaComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_NEXTDLGCTL As Long = &H28

Private mShell As Object

Private RsCv As ADODB.Recordset  'Cabecera
Private Rsd  As ADODB.Recordset  'Detalle
Private RsCli As ADODB.Recordset 'Cliente Elegido
Private RsComp As ADODB.Recordset 'Comprobante elegido
Private RsValor As ADODB.Recordset

Private cRsl As ClsLectura
Private P    As ClsPrograma

'======== ESTADOS ========
Private Enum eEstado
    stIdle = 0
    stNuevo = 1
    stViendo = 2
End Enum

Private Enum eEstadoDetalle
    detIdle = 0
    detnuevo = 1
    detEditando = 2
End Enum

Private mDetEstado As eEstadoDetalle

Private mEstado As eEstado
Private mSaltoADetalle As Boolean

'--- Producto
Private mProdId As Long
Private mProdCodigo As String
Private mProdImpuesto As Long
Private mProdUM As Long
Private mProdDeposito As Long
Private mProdTasa As Single
Private mProdCuenta As Long
Private mProdPrecio As Double
Private mProdDescripcion As String
Private mProdIntTipo As String
Private mProdIntValor As Double


'----Varios
Private mListaPrecio As Integer
Private mCodVentaTipo As Integer
Private nPorcValor As Double
Private mCliePercepcion As Double
Private mLoadingDetalle As Boolean
Private mRefreshingUI As Boolean


'========================
'  RS
'========================
Private Sub CrearRsVenta()
    On Error GoTo errHandler
    Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
    Exit Sub
errHandler:
    ManejaErrores
End Sub

Private Sub CrearRsDetalles()
    On Error GoTo errHandler
    Set Rsd = cRsl.RsVacio("DetallesComprobantes", "Id", "N")
    Exit Sub
errHandler:
    ManejaErrores
End Sub

Private Sub CrearRsdMemoria()
    On Error GoTo errHandler

    Set Rsd = New ADODB.Recordset
    Rsd.CursorLocation = adUseClient

    With Rsd.Fields
        .Append "Producto", adVarChar, 30
        .Append "Descripcion", adVarChar, 255
        .Append "Cuenta", adInteger
        .Append "Deposito", adInteger
        .Append "Medida", adInteger
        .Append "Cantidad", adDouble
        .Append "PDesc", adDouble
        .Append "Desc", adDouble
        .Append "ImpId", adInteger
        .Append "ImpPorc", adDouble             'Id impuesto o
        .Append "ImpIncluido", adInteger           '<<< NO existe en tabla real, pero te sirve
        .Append "PrecioUnitNeto", adDouble
        .Append "PrecioUnitFinal", adDouble
        .Append "Neto", adDouble
        .Append "Impuesto", adDouble
        .Append "Total", adDouble
        .Append "IntTipo", adVarChar, 20
        .Append "IntValor", adDouble
        .Append "IntUnit", adDouble
        .Append "IntTotal", adDouble
        .Append "PrecioBase", adDouble
        .Append "ImpProduc", adDouble
    End With

    Rsd.Open  'abre el rs en memoria
    
    

    Exit Sub
errHandler:
    ManejaErrores
End Sub

'========================
'  VALIDACIONES
'========================
Private Function CabeceraOK() As Boolean
    Dim ok As Boolean
    ok = True

    ok = ok And (CmbComprobante.ListIndex <> -1)

    ' Cliente: no solo TxtCliente, sino que esté linkeado (LblCliente)
    ok = ok And (Len(Trim$(TxtCliente.Text)) > 0)
    ok = ok And (Len(Trim$(LblCliente.Caption)) > 0)

    ok = ok And (CmbLista.ListIndex <> -1)

    ' Corredor: si aceptás "Ninguno", no lo bloquees
    ok = ok And (Len(Trim$(CmbCorredor.Text)) > 0)

    ' Forma de pago
    ok = ok And (Len(Trim$(CmbFormaPago.Text)) > 0)

    ' CondPago solo si CONTADO (según tu lógica)
    If UCase$(Trim$(CmbFormaPago.Text)) = "CONTADO" Then
        ok = ok And (Len(Trim$(CmbCondPago.Text)) > 0) _
                 And (UCase$(Trim$(CmbCondPago.Text)) <> "NINGUNO")
    End If

    CabeceraOK = ok
End Function

Private Function DetalleEditOK() As Boolean
    Dim ok As Boolean
    ok = True

    ok = ok And (Len(Trim$(TxtDetalle.Text)) > 0)
    ok = ok And (CmbDeposito.ListIndex <> -1)
    ok = ok And (CmbUnidad.ListIndex <> -1)
    ok = ok And (CmbImpuesto.ListIndex <> -1)
    ok = ok And (CmbCuenta.ListIndex <> -1)
    ok = ok And (CDbl(TxtCantidad.Text) > 0)
    ok = ok And (CDbl(TxtPrecio.Text) > 0)
    ok = ok And (Len(Trim$(TxtPDesc.Text)) > 0)

    DetalleEditOK = ok
End Function


'========================
'  UI REFRESH (LA CLAVE)
'========================
Private Sub RefrescarUI()
   If mRefreshingUI Then Exit Sub
    mRefreshingUI = True
    On Error GoTo Salir
  
    Select Case mEstado

        Case stIdle
            HabilitarTodo False
            HabilitarDetalles False
            BotonDetalles False, False, False, False
            BotonBuscar False, False

            'Nuevo, Buscar, Salir
            Botones True, False, False, False, False, True

            CmdBotones(0).Caption = "Nuevo"
            CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
            CmdBotones(9).Caption = "Salir"
            CmdBotones(9).Picture = LoadResPicture("Salir", 0)

      Case stNuevo
           Dim okCab As Boolean: okCab = CabeceraOK()
           Dim okDet As Boolean: okDet = DetalleOK()
           Dim okDetEdit As Boolean: okDetEdit = DetalleEditOK()
                      
            ' Cabecera: Grabar/Cancelar
           CmdBotones(0).Caption = "Grabar"
           CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
           CmdBotones(9).Caption = "Cancelar"
           CmdBotones(9).Picture = LoadResPicture("Cancelar", 0)
        
            ' Si estás editando un detalle => NO se puede grabar cabecera todavía
           If mDetEstado <> detIdle Then
              CmdBotones(0).Enabled = False
           Else
                ' Grabar cabecera solo si cabecera OK + hay detalle cargado
              CmdBotones(0).Enabled = (okCab And okDet)
            End If
        
            ' Deshabilitá buscar/acciones mientras estás cargando
'            CmdTexto.Enabled = False
'            CmdBotones(1).Enabled = False
'            CmdBotones(7).Enabled = False
'            CmdBotones(4).Enabled = False
'            CmdBotones(9).Enabled = False
'            CmdBotones(10).Enabled = False
        
            ' ===== DETALLE =====
            If mDetEstado = detIdle Then
                ' Todavía no estás editando un renglón
                HabilitarTodo True
                HabilitarDetalles False
                BotonBuscar True, False
        
                ' Botón "Nuevo detalle" solo si cabecera ok
                CmdDetalle(0).Caption = "Nuevo"
                CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
                CmdDetalle(0).Enabled = okCab
        
                CmdDetalle(1).Enabled = okDet          'Borrar (si hay registros)
                CmdDetalle(2).Enabled = okDet          'Modificar (si hay registros)
                CmdDetalle(3).Enabled = False          'Cancelar detalle no
        
            Else
                ' Estás en modo carga/edición de renglón
                HabilitarTodo False
                HabilitarDetalles True
                BotonBuscar False, True
        
                CmdDetalle(0).Caption = "Grabar"
                CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
                CmdDetalle(0).Enabled = okDetEdit
        
                CmdDetalle(1).Enabled = False
                CmdDetalle(2).Enabled = False
                CmdDetalle(3).Enabled = True           'Cancelar renglón
            End If
                   
        Case stViendo
            HabilitarTodo False
            HabilitarDetalles False

            'Nuevo, Buscar, Salir + acciones de ver
            Botones True, True, True, True, True, True

            CmdBotones(0).Caption = "Nuevo"
            CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
            CmdBotones(5).Caption = "Salir"
            CmdBotones(5).Picture = LoadResPicture("Salir", 0)

            BotonDetalles False, False, False, False
            BotonBuscar False, False
    End Select
Salir:
    mRefreshingUI = False
    
End Sub

Private Sub CmbCondPago_Click()
  If mLoadingDetalle Then Exit Sub
  Dim RsFPago As ADODB.Recordset
  
  nPorcValor = 0
  
  Set RsFPago = cRsl.TraerRsCondi("CondVenta", "CondVta", "CondVta=" & CmbCondPago.ItemData(CmbCondPago.ListIndex))
  
  If RsFPago.RecordCount <> 0 Then
      mCodVentaTipo = RsFPago!tipo
          
     If RsFPago!tipo = 7 Then
        ResetValores              'deja todo limpio
        If RsValor Is Nothing Then CrearRsValorMemoria Else VaciarRsValor
    
        nPorcValor = 0
        RecalcularRsd
    
        'Abrir FrmValores (después lo hacemos)
    Else
        'Si venís de pago mixto, borrar valores
        ResetValores
    
        nPorcValor = RsFPago!RgTarjeta
        RecalcularRsd
    End If

  End If
   
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub CrearRsValorMemoria()
    On Error GoTo errHandler

    If Not RsValor Is Nothing Then
        If RsValor.State = adStateOpen Then RsValor.Close
    End If

    Set RsValor = New ADODB.Recordset
    RsValor.CursorLocation = adUseClient

    With RsValor.Fields
        .Append "FormaPagoId", adInteger
        .Append "FormaPagoDesc", adVarChar, 50
        .Append "Importe", adDouble

        .Append "Nombre", adVarChar, 100
        .Append "BancoId", adInteger
        .Append "BancoDesc", adVarChar, 60
        .Append "NroCheque", adVarChar, 30
        .Append "FechaAcreditacion", adDate

        .Append "NroTarjeta", adVarChar, 30
        .Append "Vencimiento", adDate
        .Append "NroCupon", adVarChar, 30
        .Append "NroAutorizacion", adVarChar, 30

        .Append "Observaciones", adVarChar, 255
    End With

    RsValor.Open
    Exit Sub

errHandler:
    ManejaErrores
End Sub

Private Sub VaciarRsValor()
    On Error GoTo errHandler

    If RsValor Is Nothing Then Exit Sub
    If RsValor.State <> adStateOpen Then Exit Sub
    If RsValor.RecordCount <= 0 Then Exit Sub

    RsValor.MoveFirst
    Do While Not RsValor.EOF
        RsValor.Delete
        RsValor.MoveNext
    Loop

    Exit Sub
errHandler:
    ManejaErrores
End Sub

Private Sub ResetValores()
    On Error GoTo errHandler

    '1) No permitir que quede recargo “pegado”
    nPorcValor = 0

    '2) Vaciar valores (si existe)
    If RsValor Is Nothing Then
        'nada
    Else
        If RsValor.State <> adStateOpen Then
            'si por alguna razón está cerrado, lo recreamos
            CrearRsValorMemoria
        Else
            VaciarRsValor
        End If
    End If

    '3) Si tenés labels extra de financiación/valores, resetealos acá
    'LblFinanciacion.Caption = "0.00"

    Exit Sub
errHandler:
    ManejaErrores
End Sub

Private Function SumaValores(ByRef rs As ADODB.Recordset) As Double
    On Error GoTo errHandler

    Dim s As Double, bk As Variant
    s = 0

    If rs Is Nothing Then Exit Function
    If rs.State <> adStateOpen Then Exit Function
    If rs.RecordCount <= 0 Then Exit Function
    If (rs.BOF And rs.EOF) Then Exit Function

    bk = rs.Bookmark
    rs.MoveFirst
    Do While Not rs.EOF
        s = s + CDbl(Val(rs!importe & ""))
        rs.MoveNext
    Loop

    On Error Resume Next
    rs.Bookmark = bk
    On Error GoTo errHandler

    SumaValores = Round(s, 2)
    Exit Function

errHandler:
    ManejaErrores
End Function


Private Sub LinkearDetalleDesdeRsd()
    On Error GoTo Salir
    If mLoadingDetalle Then Exit Sub
    If Rsd Is Nothing Then Exit Sub
    If (Rsd.BOF Or Rsd.EOF) Then Exit Sub

    mLoadingDetalle = True

    TxtProducto.Text = Rsd!Producto & ""
    TxtDetalle.Text = Rsd!Descripcion & ""
    TxtCantidad.Text = Format(CDbl(Val(Rsd!Cantidad & "")), nDecimalCant)
    TxtPDesc.Text = Format(CDbl(Val(Rsd!PDesc & "")), "#0.00")
    TxtPrecio.Text = Format(CDbl(Val(Rsd!PrecioUnitFinal & "")), "#0.00")

    P.SetComboByItemData CmbCuenta, CLng(Val(Rsd!Cuenta & ""))
    P.SetComboByItemData CmbDeposito, CLng(Val(Rsd!Deposito & ""))
    P.SetComboByItemData CmbUnidad, CLng(Val(Rsd!Medida & ""))
    P.SetComboByItemData CmbImpuesto, CLng(Val(Rsd!ImpId & ""))

Salir:
    mLoadingDetalle = False
End Sub

Private Sub RecalcularRsd()
  On Error GoTo errHandler
  If mLoadingDetalle Then Exit Sub

  mLoadingDetalle = True

  Dim bk As Variant
  Dim tieneBk As Boolean
  Dim negro As Boolean

  If Rsd Is Nothing Then GoTo Salir
  If Rsd.State <> adStateOpen Then GoTo Salir
  If Rsd.RecordCount <= 0 Then GoTo Salir
  If (Rsd.BOF And Rsd.EOF) Then GoTo Salir

  negro = EsNegro()

  '--- Guardar bookmark solo si el actual es válido
  tieneBk = Not (Rsd.BOF Or Rsd.EOF)
  If tieneBk Then bk = Rsd.Bookmark

  Rsd.MoveFirst
  Do While Not Rsd.EOF

      Rsd!ImpIncluido = IIf(mListaPrecio = 1, 1, 0)

      If negro Then
          'NEGRO: NO IVA, NO INTERNO
          Rsd!ImpPorc = 0
          Rsd!IntTipo = ""
          Rsd!IntValor = 0
      Else
          'BLANCO: recuperar tasa real del producto
          If Not IsNull(Rsd!ImpProduc) Then
              Rsd!ImpPorc = CDbl(Val(Rsd!ImpProduc & ""))
          Else
              'por las dudas, si venís de registros viejos sin el campo seteado
              Rsd!ImpPorc = 0
          End If
      End If

      CalcularImportesRenglon Rsd
      Rsd.Update
      Rsd.MoveNext
  Loop

  '--- Restaurar bookmark si se puede
  On Error Resume Next
  If tieneBk Then Rsd.Bookmark = bk
  err.Clear
  On Error GoTo errHandler

Salir:
  mLoadingDetalle = False
  CalcularTotales
  Exit Sub

errHandler:
  mLoadingDetalle = False
  ManejaErrores
End Sub


Private Sub CmbCorredor_Click()
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub CmbLista_Click()
    If mLoadingDetalle Then Exit Sub
    If mEstado = stNuevo Then RefrescarUI
    If CmbLista.ListIndex <> -1 Then
        mListaPrecio = cRsl.TraerValorDeUnCampo("ListaDePrecio", "PrecioIva", _
                     "ListaPrecio=" & CmbLista.ItemData(CmbLista.ListIndex))
        RecalcularRsd
    End If
End Sub

Private Sub CmdClientes_Click()
    nDat = 0
    nBuscar = 1
    FrmBuscarEx.Show 1

    If nDat <> 0 Then
        LinkearCliente nDat

        If mDetEstado = detIdle Then
          If Len(Trim$(LblCliente.Caption)) > 0 Then CmbLista.SetFocus
        End If
    End If
End Sub

Private Sub CmdProductos_Click()
 nDat = 0
 nBuscar = 3
 FrmBuscarEx.Show 1
  If nDat <> 0 Then
     LinkearProducto nDat
  End If
End Sub

Private Sub DtpFecha_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub DtpVenc_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

'========================
'  FORM LOAD
'========================
Private Sub Form_Load()
    On Error GoTo errHandler

    Set mShell = CreateObject("WScript.Shell")
    Set cRsl = New ClsLectura
    Set P = New ClsPrograma

    CrearRsdMemoria
    CabGrid

    'imagenes
    CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
    CmdBotones(2).Picture = LoadResPicture("Buscar", 0)
    CmdBotones(9).Picture = LoadResPicture("Salir", 0)
  '  CmdBotones(4).Picture = LoadResPicture("Valores", 0)
    CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
    CmdBotones(5).Picture = LoadResPicture("Imprimir", 0)
    CmdBotones(3).Picture = LoadResPicture("Modificar", 0)
    CmdBotones(8).Picture = LoadResPicture("Email", 0)

    CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
    CmdDetalle(1).Picture = LoadResPicture("Borrar", 0)
    CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
    CmdDetalle(3).Picture = LoadResPicture("Cancelar", 0)
       
    CabGrid
    CargarCombos

    Limpiar

    mEstado = stIdle
    RefrescarUI

    Me.Top = 0
    Me.Left = 0
    Exit Sub

errHandler:
    On Error Resume Next
    ManejaErrores
End Sub


'========================
'  ENTER COMO TAB (TU MÉTODO)
'========================
Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim WshShell As Object
  If KeyAscii = vbKeyReturn Then
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub DtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
     Set WshShell = CreateObject("WScript.Shell")
     WshShell.SendKeys "{TAB}"
  End If
End Sub

Private Sub DtpVenc_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
     Set WshShell = CreateObject("WScript.Shell")
     WshShell.SendKeys "{TAB}"
  End If
End Sub


'========================
'  NUMÉRICOS / SELECCIONAR
'========================
Private Sub SoloNumEntero(ByRef KeyAscii As Integer)
    If SoloNumero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub SoloNumDecimal(ByRef KeyAscii As Integer, ByVal tb As TextBox)
    If SoloNumeroDecimalFinal(KeyAscii, tb) = False Then KeyAscii = 0
End Sub

Private Sub SeleccionarTodo(ByVal tb As TextBox)
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If mLoadingDetalle Then Exit Sub
    If Rsd Is Nothing Then Exit Sub
    If (Rsd.BOF Or Rsd.EOF) Then Exit Sub

    If mEstado = stNuevo And mDetEstado = detIdle Then
        LinkearDetalleActual
        RefrescarUI
    End If
End Sub

Private Sub LinkearDetalleActual()
    On Error GoTo Salir
    If mLoadingDetalle Then Exit Sub
    If Rsd Is Nothing Then Exit Sub
    If (Rsd.BOF Or Rsd.EOF) Then Exit Sub

    mLoadingDetalle = True

    TxtProducto.Text = Rsd!Producto & ""
    TxtDetalle.Text = Rsd!Descripcion & ""
    TxtCantidad.Text = Format(Val(Rsd!Cantidad & ""), nDecimalCant)
    TxtPDesc.Text = Format(Val(Rsd!PDesc & ""), "0.00")
    TxtPrecio.Text = Format(Val(Rsd!precioBase & ""), "0.00")

    P.SetComboByItemData CmbDeposito, CLng(Val(Rsd!Deposito & ""))
    P.SetComboByItemData CmbUnidad, CLng(Val(Rsd!Medida & ""))
    P.SetComboByItemData CmbCuenta, CLng(Val(Rsd!Cuenta & ""))
    P.SetComboByItemData CmbImpuesto, CLng(Val(Rsd!ImpId & ""))

Salir:
    mLoadingDetalle = False
End Sub

Private Sub TxtCantidad_LostFocus()
  If TxtCantidad.Text = "" Then TxtCantidad.Text = 1
  TxtCantidad.Text = Format(TxtCantidad.Text, nDecimalCant)
End Sub

Private Sub TxtCliente_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub TxtCliente_LostFocus()
    LinkearCliente 0

    If mDetEstado <> detIdle Then Exit Sub '<<< importante

    If Len(Trim$(LblCliente.Caption)) > 0 Then
        CmbLista.SetFocus
    End If
End Sub


Private Sub TxtCotizacion_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub TxtPrecio_LostFocus()
  If TxtPrecio.Text = "" Then TxtPrecio.Text = 0
  TxtPrecio.Text = Format(TxtPrecio.Text, nCantDecimales)
End Sub

Private Sub TxtPDesc_LostFocus()
  If TxtPDesc.Text = "" Then TxtPDesc.Text = 0
  TxtPDesc.Text = Format(TxtPDesc.Text, "#0.00")
End Sub

Private Sub TxtNumero_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub LinkearProducto(pProducto As Long)
  Dim cRpl As ClsProductoL
  
  Set cRpl = New ClsProductoL
  
  If TxtProducto.Text <> "" Or nDat <> 0 Then
     Dim RsProduc As ADODB.Recordset
     Set RsProduc = cRpl.TraerProducto(TxtProducto.Text, pProducto)
     
     If Not RsProduc Is Nothing And RsProduc.RecordCount > 0 Then
     
        mProdId = RsProduc!ID
        mProdCodigo = RsProduc!Producto
        mProdDescripcion = RsProduc!Descripcion
        mProdUM = RsProduc!UMVenta
        mProdTasa = RsProduc!Porcentaje
        mProdDeposito = RsProduc!Deposito
        mProdImpuesto = RsProduc!Impuesto
        mProdPrecio = CDbl(RsProduc!precio)
        mProdIntTipo = RsProduc!TipoInterno
        mProdIntValor = RsProduc!ValorInterno
        
        TxtProducto.Text = mProdCodigo
        TxtDetalle.Text = mProdDescripcion
        
        P.SetComboByItemData CmbUnidad, mProdUM
        P.SetComboByItemData CmbDeposito, mProdDeposito
         CmbImpuesto.Text = "IVA EXENTO"
         If Not RsComp Is Nothing Then
            If RsComp!Iva = 1 Then
              P.SetComboByItemData CmbImpuesto, mProdImpuesto
            End If
        End If
        If RsProduc!precio <> 0 Then
           TxtPrecio.Text = Format(mProdPrecio, "#0.00")
           TxtCantidad.SetFocus
        Else
           TxtPrecio.Text = "0.00"
           TxtPrecio.SetFocus
        End If
     Else
        LimpiarDetalles
        TxtProducto.SetFocus
     End If
  End If
End Sub


Private Sub TxtProducto_LostFocus()
   LinkearProducto 0
End Sub

Private Sub TxtReparto_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub TxtSucursal_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub TxtVendedor_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

'-- ejemplo (dejá el resto igual)
Private Sub TxtVendedor_GotFocus(): SeleccionarTodo TxtVendedor: End Sub
Private Sub TxtVendedor_KeyPress(KeyAscii As Integer): SoloNumDecimal KeyAscii, TxtVendedor: End Sub


'========================
'  LIMPIAR
'========================
Private Sub Limpiar()
    Dim ctl As Control

    For Each ctl In Controls
        If TypeOf ctl Is TextBox Then ctl.Text = ""
    Next ctl

    DTPFecha.Value = Date
    DtpVenc.Value = Date

    LblDescuento.Caption = "0.00"
    TxtReparto.Text = "0000000000000"
    LblCliente.Caption = ""
    TxtSucursal.Text = "0000"
    TxtNumero.Text = "00000000"
    TxtVendedor.Text = "0.00"
    LbLIva.Caption = ""
    TxtCotizacion.Text = "1.000"
    LblNeto.Caption = "0.00"
    LblBonificacion.Caption = "0.00"
    LblFinanciacion.Caption = "0.00"
    TxtNoGrav.Text = "0.00"
    LblIva1.Text = "0.00"
    LblPercepcion.Caption = "0.00"
    LblTotal.Caption = "0.00"

    CmbComprobante.ListIndex = -1
    CmbVend.Text = "NINGUNO"
    CmbCorredor.Text = "Ninguno"
    CmbLista.ListIndex = -1
    CmbCuenta.ListIndex = -1

    ' Forma pago / cond pago
    CmbFormaPago.Text = IIf(nCondPago = 0, "CUENTA CORRIENTE", "CONTADO")
    CmbCondPago.Text = "EFECTIVO"
      
End Sub

Private Sub LimpiarDetalles()
    TxtProducto.Text = ""
    P.SetComboByItemData CmbDeposito, nDepDetalle
    P.SetComboByItemData CmbUnidad, nMedDetalle
    P.SetComboByItemData CmbCuenta, nCueDetalle
    
    CmbImpuesto.Text = "IVA EXENTO"
    If Not RsComp Is Nothing Then
        If RsComp!Iva = 1 Then
           P.SetComboByItemData CmbImpuesto, nImpDetalle
        End If
    End If
    TxtCantidad.Text = Format(1, nDecimalCant)
    TxtDetalle.Text = ""
    TxtPDesc.Text = "0.00"
    TxtPrecio.Text = Format(0, nCantDecimales) '<<< corregido
End Sub


'========================
'  BOTONES
'========================
Private Sub CmdBotones_Click(Index As Integer)
    Select Case Index
        Case 0
            Nuevo
        Case 9
            Salir
    End Select
End Sub

Private Sub Nuevo()
    On Error GoTo errHandler

    If mEstado <> stNuevo Then
        ' pasar a nuevo
       mEstado = stNuevo
       ResetDetalleMemoria
       ResetValores          '<<< ACÁ
       LimpiarDetalles
       CalcularTotales
       RefrescarUI
        
       CmbCorredor.Text = "Ninguno"
       P.SetComboByItemData CmbComprobante, nVentaFactura
       
       CmbComprobante.SetFocus
    Else
        mEstado = stIdle
        Limpiar
        RefrescarUI
    End If
    Exit Sub

errHandler:
    MsgBox err.Description, vbCritical, "Nuevo"
End Sub

Private Sub Salir()
   If mEstado = stNuevo Then
    If MsgBox("Cancela la creación del comprobante?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
    mEstado = stIdle
    Limpiar
    LimpiarDetalles
    ResetDetalleMemoria
    ResetValores          '<<< ACÁ
    RefrescarUI
    Else
        Unload Me
    End If
End Sub


Private Sub CmbFormaPago_Click()
   If CmbFormaPago.Text <> "CONTADO" Then
      CmbCondPago.Enabled = False
      CmbCondPago.Text = "Ninguno"
      nPorcValor = 0
      If Rsd.RecordCount <> 0 Then RecalcularRsd
   Else
      CmbCondPago.Enabled = True
      CmbCondPago.Text = "EFECTIVO"
   End If
   If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub TxtReparto_GotFocus()
    SeleccionarTodo TxtReparto
End Sub

Private Sub TxtReparto_KeyPress(KeyAscii As Integer)
    SoloNumEntero KeyAscii
End Sub

Private Sub TxtSucursal_GotFocus()
    SeleccionarTodo TxtSucursal
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
    SoloNumEntero KeyAscii
End Sub

Private Sub TxtNumero_GotFocus()
    SeleccionarTodo TxtNumero
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
    SoloNumEntero KeyAscii
End Sub

Private Sub TxtCliente_GotFocus()
    SeleccionarTodo TxtCliente
End Sub

Private Sub TxtCotizacion_GotFocus()
    SeleccionarTodo TxtCotizacion
End Sub

Private Sub TxtCantidad_GotFocus()
    SeleccionarTodo TxtCantidad
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    SoloNumDecimal KeyAscii, TxtCantidad
End Sub

Private Sub TxtPDesc_GotFocus()
    SeleccionarTodo TxtPDesc
End Sub

Private Sub TxtPrecio_GotFocus()
    SeleccionarTodo TxtPrecio
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
    SoloNumDecimal KeyAscii, TxtPrecio
End Sub

Private Sub Botones(bNue As Boolean, bComun As Boolean, bBuscar As Boolean, bModi As Boolean, bAplica As Boolean, bSalir As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bComun
  CmdBotones(2).Enabled = bBuscar
  CmdBotones(3).Enabled = bModi
  CmdBotones(4).Enabled = bComun
  CmdBotones(5).Enabled = bComun
  CmdBotones(6).Enabled = bAplica
  CmdBotones(7).Enabled = bComun
  CmdBotones(8).Enabled = bComun
  CmdBotones(9).Enabled = bSalir
   
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

Private Sub HabilitarTodo(bEstado As Boolean)
  Dim ctl As Control

  For Each ctl In Controls

      '<<< NO TOCAR EL FRAME DEL DETALLE NI SUS CONTROLES
      If ctl.Container Is Frame1 Then
          'saltarlo
      Else
          If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
             Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
             Or TypeOf ctl Is ComboBox Then
             ctl.Enabled = bEstado
          End If
      End If

  Next ctl

  Grid1.Enabled = bEstado
End Sub


Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtProducto.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  CmbImpuesto.Enabled = bHabi
  CmbDeposito.Enabled = bHabi
  CmbUnidad.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  TxtPDesc.Enabled = bHabi
  TxtPrecio.Enabled = bHabi
  CmbCuenta.Enabled = bHabi
End Sub

Private Sub CabGrid()
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.Bold = True
  With Grid1
      .Columns(0).Caption = "Producto"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Descripción"
      .Columns(1).Width = 4400
      .Columns(2).Visible = False
      .Columns(3).Visible = False
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
      .Columns(11).Caption = "P.Uni."
      .Columns(11).Width = 1000
      .Columns(11).NumberFormat = "0.00"
      .Columns(11).Alignment = dbgRight
      .Columns(12).Visible = False
      .Columns(13).Visible = False
      .Columns(14).Visible = False
      .Columns(15).Caption = "Total"
      .Columns(15).Width = 1200
      .Columns(15).NumberFormat = "#0.00" ' nCantDecimales
      .Columns(15).Alignment = dbgRight
      .Columns(16).Visible = False
      .Columns(17).Visible = False
      .Columns(18).Visible = False
      .Columns(19).Visible = False
      .Columns(20).Visible = False
      .Columns(21).Visible = False
 End With
End Sub

Private Sub CargarCombos()
  On Error GoTo errHandler
  
 
  cRsl.CargaCombo CmbComprobante, "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
  cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmbCorredor, "Clientes", "Id", "RazonSocial", "Tipo=3"
     
  CmbCorredor.AddItem "Ninguno"
    
  cRsl.CargaCombo CmbImpuesto, "Impuestos", "Impuesto", "Descripcion", "Producto='1' OR Retencion='1' OR Percepcion='1' OR Interno='1'"
  cRsl.CargaCombo CmbCondPago, "CondVenta", "CondVta", "Descripcion", "Tipo<9 and tipo<>3"
  CmbCondPago.AddItem "Ninguno"
  
  CmbFormaPago.AddItem "CUENTA CORRIENTE"
  CmbFormaPago.ItemData(CmbFormaPago.NewIndex) = 0
  CmbFormaPago.AddItem "CONTADO"
  CmbFormaPago.ItemData(CmbFormaPago.NewIndex) = 1
  CmbFormaPago.AddItem "COTIZAR"
  CmbFormaPago.ItemData(CmbFormaPago.NewIndex) = 2
  
  cRsl.CargaCombo CmbUnidad, "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbDeposito, "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  cRsl.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", "Grupo=4"
  cRsl.CargaCombo CmbVend, "Vendedores", "Vendedor", "Descripcion", ""
   
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmbComprobante_Click()
  If mLoadingDetalle Then Exit Sub
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler
  
  If mEstado = stNuevo Then
     Set cNum = New ClsComprobantesL
     
     Set RsComp = cRsl.TraerRsCondi("Comprobantes", "Id", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
     
     ' Si en impresion es false Numera cuando Termina de grabar el comprobante, si es verdadero actualiza el numero aca
     
     If cRsl.TraerValorDeUnCampo("UsuariosComprobantesVentas", "Comanda", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) = False Then
        ChkA4.Enabled = False
     Else
        ChkA4.Enabled = True
     End If
     
     If cRsl.TraerValorDeUnCampo("Impresion", "Numero", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) = False Then
        nNum = Format(cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
        nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
     Else
        nNum = Format(cNum.TraerUltimoNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
        nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
     End If
     TxtSucursal.Text = nSuc
     TxtNumero.Text = nNum
     If Rsd.RecordCount <> 0 Then
        RecalcularRsd ' cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "CondVta=" & CmbCondPago.ItemData(CmbCondPago.ListIndex))
     End If
        
  End If
  If mEstado = stNuevo Then RefrescarUI
  
  Set cNum = Nothing
   
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Public Sub CalcularTotales()
    On Error GoTo errHandler

    Dim nNeto As Double, nIva As Double, nDes As Double
    Dim nNoGrabado As Double, nPercepcion As Double, nTot As Double
    Dim nNetoGravado As Double
    Dim bk As Variant
    Dim tieneBk As Boolean

    nNeto = 0: nIva = 0: nDes = 0: nNoGrabado = 0: nPercepcion = 0: nTot = 0
    nNetoGravado = 0

    If Rsd Is Nothing Then GoTo pintar
    If Rsd.RecordCount <= 0 Then GoTo pintar
    If (Rsd.BOF Or Rsd.EOF) Then GoTo pintar   '<<< CLAVE

    '--- guardar bookmark solo si hay registro actual válido
    tieneBk = True
    bk = Rsd.Bookmark

    Rsd.MoveFirst
    Do While Not Rsd.EOF
        nNeto = nNeto + CDbl(Val(Rsd!neto & ""))
        nIva = nIva + CDbl(Val(Rsd!Impuesto & ""))
        nDes = nDes + CDbl(Val(Rsd!desc & ""))
        nNoGrabado = nNoGrabado + CDbl(Val(Rsd!IntTotal & ""))
        nTot = nTot + CDbl(Val(Rsd!Total & ""))
        Rsd.MoveNext
    Loop

    '--- volver si el bookmark sigue siendo válido
    On Error Resume Next
    If tieneBk Then Rsd.Bookmark = bk
    On Error GoTo errHandler

pintar:
    nNetoGravado = Round(nNeto - nNoGrabado, 2)

    nPercepcion = 0
    If PuedePercepcion() Then
        nPercepcion = Round(nNetoGravado * mCliePercepcion / 100, 2)
    End If

    LblFinanciacion.Caption = "0.00"
    TxtNoGrav.Text = Format(Round(nNoGrabado, 2), "#0.00")
    LblNeto.Caption = Format(Round(nNetoGravado, 2), "#0.00")
    LblBonificacion.Caption = Format(Round(nDes, 2), "#0.00")
    LblIva1.Text = Format(Round(nIva, 2), "#0.00")
    LblPercepcion.Caption = Format(Round(nPercepcion, 2), "#0.00")
    LblTotal.Caption = Format(Round(nTot + nPercepcion, 2), "#0.00")
    Exit Sub

errHandler:
    ManejaErrores
End Sub



Private Function PuedePercepcion() As Boolean
    PuedePercepcion = False

    '1) Debe ser blanco (lleva IVA)
    If RsComp Is Nothing Then Exit Function
    If Val(RsComp!Iva & "") <> 1 Then Exit Function

    '2) Debe discriminar (Factura A típicamente)
    If Val(RsComp!Costo & "") <> 1 Then Exit Function

    '3) Cliente con percepción configurada
    If mCliePercepcion = 0 Then Exit Function

    PuedePercepcion = True
End Function


Private Sub LinkearCliente(nDat As Long)
  Dim cRiL As ClsClienteL
  On Error GoTo errHandler
    
  Set cRiL = New ClsClienteL
  LblCliente.Caption = ""
  LbLIva.Caption = ""
  If TxtCliente.Text <> "" Or nDat <> 0 Then
       Set RsCli = cRiL.TraerDatosCliente(TxtCliente.Text, nDat)
       TxtCliente.Text = RsCli!Cliente
       mCliePercepcion = RsCli!Percepcion
       
       If RsCli.RecordCount <> 0 Then
         If IsNull(RsCli!FechaBaja) Or RsCli!FechaBaja = "" Then
            LblCliente.Caption = RsCli!RazonSocial
            LbLIva.Caption = RsCli!TivaDesc
            P.SetComboByItemData CmbLista, RsCli!ListaPrecio
            LblDescuento.Caption = Format(RsCli!DescPorc, "0.00")
            If CmbVend.Text = "NINGUNO" Then
               CmbVend.Text = RsCli!VendDesc
               TxtVendedor.Text = Format(RsCli!VendComision, "0.00")
            End If
         End If
      End If
  End If
  
  Set cRiL = Nothing
  If mEstado = stNuevo Then RefrescarUI

Exit Sub

errHandler:
   ManejaErrores
End Sub

Private Sub CmdDetalle_Click(Index As Integer)
    Select Case Index
        Case 0 'Nuevo/Grabar
            If mDetEstado = detIdle Then
                'Nuevo renglón
                mDetEstado = detnuevo
                LimpiarDetalles
                RefrescarUI
                TxtProducto.SetFocus
            Else
                'Grabar (nuevo o edición)
                If mDetEstado = detnuevo Then
                    GrabarDetalleActual
                ElseIf mDetEstado = detEditando Then
                    ActualizarDetalleActual
                End If

                mDetEstado = detIdle
                RefrescarUI
                CmdDetalle(0).SetFocus
            End If

        Case 1 'Borrar
            BorrarDetalleActual

        Case 2 'Modificar
            If Rsd Is Nothing Then Exit Sub
            If (Rsd.BOF And Rsd.EOF) Then Exit Sub
            mDetEstado = detEditando
            RefrescarUI
            TxtDetalle.SetFocus

        Case 3 'Cancelar
            mDetEstado = detIdle
            LinkearDetalleDesdeRsd 'volver a lo que estaba seleccionado
            RefrescarUI
    End Select
End Sub



Private Sub BorrarDetalleActual()
    On Error GoTo errHandler

    If Rsd Is Nothing Then Exit Sub
    If Rsd.RecordCount <= 0 Then Exit Sub
    If (Rsd.BOF Or Rsd.EOF) Then Exit Sub

    If MsgBox("¿Eliminar el renglón seleccionado?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub

    Dim bkActual As Variant
    Dim bkIr As Variant
    Dim RsC As ADODB.Recordset

    bkActual = Rsd.Bookmark
    bkIr = Null

    '--- calcular bkIr sin tocar el Rsd real
    Set RsC = Rsd.Clone
    RsC.Bookmark = bkActual

    RsC.MoveNext
    If Not RsC.EOF Then
        bkIr = RsC.Bookmark
    Else
        RsC.MovePrevious
        If Not RsC.BOF Then bkIr = RsC.Bookmark
    End If

    RsC.Close
    Set RsC = Nothing

    '--- volver al actual y borrar ESE
    Rsd.Bookmark = bkActual
    Rsd.Delete   '<<< NO Update

    '--- reposicionarse
    If Rsd.RecordCount > 0 Then
        On Error Resume Next
        If Not IsNull(bkIr) Then
            Rsd.Bookmark = bkIr
        Else
            Rsd.MoveFirst
        End If
        If err.Number <> 0 Then
            err.Clear
            Rsd.MoveFirst
        End If
        On Error GoTo errHandler

        LinkearDetalleActual
    Else
        LimpiarDetalles
    End If

    CalcularTotales
    RefrescarUI
    Exit Sub

errHandler:
    ManejaErrores
End Sub


Private Sub IniciarEdicionDetalleActual()
    On Error GoTo errHandler

    If Rsd Is Nothing Then Exit Sub
    If (Rsd.BOF And Rsd.EOF) Then Exit Sub

    'paso a modo edición
    mDetEstado = detEditando
    RefrescarUI

    'cargo controles desde el rs
    TxtProducto.Text = Rsd!Producto & ""
    TxtDetalle.Text = Rsd!Descripcion & ""

    P.SetComboByItemData CmbCuenta, CLng(Val(Rsd!Cuenta & ""))
    P.SetComboByItemData CmbDeposito, CLng(Val(Rsd!Deposito & ""))
    P.SetComboByItemData CmbUnidad, CLng(Val(Rsd!Medida & ""))
    P.SetComboByItemData CmbImpuesto, CLng(Val(Rsd!ImpId & ""))

    TxtCantidad.Text = Format(CDbl(Val(Rsd!Cantidad & "")), nDecimalCant)
    TxtPDesc.Text = Format(CDbl(Val(Rsd!PDesc & "")), "0.00")

    'ojo: el precio base NO lo guardás en el rs (todavía).
    'lo más práctico: guardarlo también (te lo propongo abajo)
    'por ahora, lo dejamos como está (o lo re-traés del producto)

    TxtPrecio.SetFocus
    Exit Sub

errHandler:
    ManejaErrores
End Sub


Private Sub ActualizarDetalleActual()
    On Error GoTo errHandler

    Dim negro As Boolean
    negro = EsNegro()

    If Rsd Is Nothing Then Exit Sub
    If (Rsd.BOF Or Rsd.EOF) Then Exit Sub

    Rsd!Producto = TxtProducto.Text
    Rsd!Descripcion = TxtDetalle.Text
    Rsd!Cuenta = CmbCuenta.ItemData(CmbCuenta.ListIndex)
    Rsd!Deposito = CmbDeposito.ItemData(CmbDeposito.ListIndex)
    Rsd!Medida = CmbUnidad.ItemData(CmbUnidad.ListIndex)
    Rsd!Cantidad = CDbl(Val(TxtCantidad.Text))
    Rsd!PDesc = CDbl(Val(TxtPDesc.Text))
    Rsd!ImpProduc = mProdTasa
    Rsd!precioBase = CDbl(Val(TxtPrecio.Text))

    Rsd!ImpId = CmbImpuesto.ItemData(CmbImpuesto.ListIndex)
    Rsd!ImpIncluido = IIf(mListaPrecio = 1, 1, 0)

    If negro Then
        Rsd!ImpPorc = 0
        Rsd!IntTipo = ""
        Rsd!IntValor = 0
    Else
        'Si querés, acá podés respetar el impuesto elegido en combo
        Rsd!ImpPorc = mProdTasa
        Rsd!IntTipo = mProdIntTipo
        Rsd!IntValor = mProdIntValor
    End If

    CalcularImportesRenglon Rsd
    Rsd.Update

    CalcularTotales
    Exit Sub

errHandler:
    ManejaErrores
End Sub


Private Sub GrabarDetalleActual()
    On Error GoTo errHandler

    Dim negro As Boolean
    negro = EsNegro()

    Rsd.AddNew
    Rsd!Producto = TxtProducto.Text
    Rsd!Descripcion = TxtDetalle.Text
    Rsd!Cuenta = CmbCuenta.ItemData(CmbCuenta.ListIndex)
    Rsd!Deposito = CmbDeposito.ItemData(CmbDeposito.ListIndex)
    Rsd!Medida = CmbUnidad.ItemData(CmbUnidad.ListIndex)
    Rsd!Cantidad = CDbl(Val(TxtCantidad.Text))
    Rsd!PDesc = CDbl(Val(TxtPDesc.Text))
    Rsd!ImpProduc = mProdTasa
    Rsd!precioBase = CDbl(Val(TxtPrecio.Text))

    Rsd!ImpId = CmbImpuesto.ItemData(CmbImpuesto.ListIndex)
    Rsd!ImpIncluido = IIf(mListaPrecio = 1, 1, 0)

    If negro Then
        Rsd!ImpPorc = 0
        Rsd!IntTipo = ""
        Rsd!IntValor = 0
    Else
        Rsd!ImpPorc = mProdTasa
        Rsd!IntTipo = mProdIntTipo
        Rsd!IntValor = mProdIntValor
    End If

    CalcularImportesRenglon Rsd
    Rsd.Update

    CalcularTotales
    Exit Sub

errHandler:
    ManejaErrores
End Sub


Private Sub CalcularImportesRenglon(ByRef R As ADODB.Recordset)
    On Error GoTo errHandler
    
    Dim tipo As String
    tipo = UCase$(Trim$(R!IntTipo & ""))

    Dim cant As Double
    Dim precioBase As Double
    Dim tasa As Double

    Dim porcDescCli As Double, porcDescLin As Double
    Dim porcRecCli As Double, porcRecVal As Double

    Dim precioRecargado As Double
    Dim precioConDesc As Double

    Dim intUnit As Double, intTot As Double
    Dim finalUnit As Double, netoUnit As Double

    Dim neto As Double, impu As Double, tot As Double
    Dim descMonto As Double

    '--- básicos
    cant = CDbl(Val(R!Cantidad))
    If cant <= 0 Then cant = 0

    precioBase = CDbl(Val(R!precioBase & ""))
    tasa = CDbl(Val(R!ImpPorc)) '0 si negro

    '--- % desc/rec
    porcDescCli = 0
    If Not (RsCli Is Nothing) Then porcDescCli = CDbl(Val(RsCli!DescPorc & ""))

    porcDescLin = CDbl(Val(R!PDesc & ""))

    porcRecCli = 0
    If Not (RsCli Is Nothing) Then porcRecCli = CDbl(Val(RsCli!PorcRecargo & ""))

    porcRecVal = CDbl(Val(nPorcValor))

    '--- 1) recargos
    precioRecargado = precioBase * (1# + (porcRecCli + porcRecVal) / 100#)

    '--- 2) descuentos (cliente + línea)
    precioConDesc = precioRecargado * (1# - (porcDescCli + porcDescLin) / 100#)
    If precioConDesc < 0 Then precioConDesc = 0

    '--- descuento $ (sobre lo recargado)
    descMonto = (precioRecargado - precioConDesc) * cant
    If descMonto < 0 Then descMonto = 0

    '--- 3) interno (se calcula sobre el precio ya con rec/desc)
    intUnit = 0
    If CDbl(Val(R!IntValor & "")) <> 0 Then
      If tipo = "FIJO" Then
        intUnit = CDbl(Val(R!IntValor & ""))
      ElseIf tipo = "PORCENTAJE" Then
         intUnit = precioConDesc * (CDbl(Val(R!IntValor & "")) / 100#)
      Else
        intUnit = 0
      End If
    End If
    intTot = intUnit * cant

    '--- 4) neto/final/iva
    If Val(R!ImpIncluido & "") = 1 Then
        'precioConDesc es FINAL (con IVA) y puede incluir interno
        finalUnit = precioConDesc

        'IVA sale del monto SIN interno: (final - interno)
        Dim baseIvaFinal As Double
        baseIvaFinal = finalUnit - intUnit
        If baseIvaFinal < 0 Then baseIvaFinal = 0

        If tasa > 0 Then
            netoUnit = (baseIvaFinal / (1# + tasa / 100#)) + intUnit
        Else
            netoUnit = finalUnit
        End If

    Else
        'precioConDesc es NETO (sin IVA), interno se suma al neto
        netoUnit = precioConDesc + intUnit
        finalUnit = netoUnit * (1# + tasa / 100#)
    End If

    '--- 5) totales
    neto = netoUnit * cant
    tot = finalUnit * cant
    impu = tot - neto

    '--- 6) redondeos
    netoUnit = Round(netoUnit, nCantDecimales)
    finalUnit = Round(finalUnit, nCantDecimales)

    neto = Round(neto, 2)
    tot = Round(tot, 2)
    impu = Round(impu, 2)
    descMonto = Round(descMonto, 2)
    intUnit = Round(intUnit, nCantDecimales)
    intTot = Round(intTot, 2)

    '--- 7) asignar
    R!PrecioUnitNeto = netoUnit
    R!PrecioUnitFinal = finalUnit
    R!neto = neto
    R!Impuesto = impu
    R!desc = descMonto

    R!intUnit = intUnit
    R!IntTotal = intTot

    R!Total = tot  'tot ya incluye interno porque finalUnit lo incluye

    Exit Sub
errHandler:
    ManejaErrores
End Sub

Private Function DetalleOK() As Boolean
    On Error GoTo Salir

    DetalleOK = False

    If Rsd Is Nothing Then Exit Function
    If Rsd.State <> adStateOpen Then Exit Function
    If Rsd.RecordCount <= 0 Then Exit Function
    If (Rsd.BOF And Rsd.EOF) Then Exit Function

    Dim RsC As ADODB.Recordset
    Set RsC = Rsd.Clone   '<<< no toca el recordset del grid

    RsC.MoveFirst
    Do While Not RsC.EOF
        If Len(Trim$(RsC!Descripcion & "")) > 0 Then
            DetalleOK = True
            Exit Do
        End If
        RsC.MoveNext
    Loop

    RsC.Close
    Set RsC = Nothing

Salir:
End Function

Private Sub TxtDetalle_Change()
    If mEstado = stNuevo And mDetEstado <> detIdle Then
        RefrescarUI
    End If
End Sub

Private Sub TxtCantidad_Change()
  If mEstado = stNuevo And mDetEstado <> detIdle Then RefrescarUI
End Sub

Private Sub TxtPrecio_Change()
  If mEstado = stNuevo And mDetEstado <> detIdle Then RefrescarUI
End Sub

Private Sub CmbDeposito_Click()
  If mEstado = stNuevo And mDetEstado <> detIdle Then RefrescarUI
End Sub

Private Sub TxtPDesc_Change()
    If mEstado = stNuevo And mDetEstado <> detIdle Then RefrescarUI
End Sub

Private Function EsNegro() As Boolean
    EsNegro = False
    If Not (RsComp Is Nothing) Then
        EsNegro = (Val(RsComp!Iva & "") = 0)
    End If
End Function

Private Sub ResetDetalleMemoria()
    If Not Rsd Is Nothing Then
        If Rsd.State = adStateOpen Then Rsd.Close
    End If
    CrearRsdMemoria
    Set Grid1.DataSource = Rsd
    CabGrid
End Sub

