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
      Format          =   165543937
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
      Format          =   165543937
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
Private p    As ClsPrograma

'======== ESTADOS ========
Private Enum eEstado
    stIdle = 0
    stNuevo = 1
    stViendo = 2
    stEditando = 3
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
Private mprodActuaStock As String


'----Varios
Private mListaPrecio As Integer
Private mCodVentaTipo As Integer
Private nPorcValor As Double
Private mCliePercepcion As Double
Private mLoadingDetalle As Boolean
Private mRefreshingUI As Boolean
Private mMoviAfip As Long

Private mTipoOperacion As Byte
Private mTipoMovimiento As Byte
Dim imp_iva_fact As Double, imp_tributo_fact As Double, bEMail As Boolean

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
        .Append "ActuaStock", adVarChar, 2
    End With

    Rsd.Open  'abre el rs en memoria
    BindGridRsd
    

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
    ok = ok And (Len(Trim$(TxtCliente.text)) > 0)
    ok = ok And (Len(Trim$(LblCliente.Caption)) > 0)

    ok = ok And (CmbLista.ListIndex <> -1)

    ' Corredor: si aceptás "Ninguno", no lo bloquees
    ok = ok And (Len(Trim$(CmbCorredor.text)) > 0)

    ' Forma de pago
    ok = ok And (Len(Trim$(CmbFormaPago.text)) > 0)

    ' CondPago solo si CONTADO (según tu lógica)
    If UCase$(Trim$(CmbFormaPago.text)) = "CONTADO" Then
        ok = ok And (Len(Trim$(CmbCondPago.text)) > 0) _
                 And (UCase$(Trim$(CmbCondPago.text)) <> "NINGUNO")
    End If

    CabeceraOK = ok
End Function

Private Function DetalleEditOK() As Boolean
    Dim ok As Boolean
    ok = True

    ok = ok And (Len(Trim$(TxtDetalle.text)) > 0)
    ok = ok And (CmbDeposito.ListIndex <> -1)
    ok = ok And (CmbUnidad.ListIndex <> -1)
    ok = ok And (CmbImpuesto.ListIndex <> -1)
    ok = ok And (CmbCuenta.ListIndex <> -1)
    ok = ok And (CDbl(Val(TxtCantidad.text)) > 0)
    ok = ok And (CDbl(Val(TxtPrecio.text)) > 0)
    ok = ok And (Len(Trim$(Val(TxtPDesc.text))) > 0)

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
            Botones True, False, True, False, False, True

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
            CmdBotones(2).Enabled = False

        
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
             Dim puedeEditar As Boolean, puedeAnular As Boolean
             Dim okDetV As Boolean
             okDetV = DetalleOK() 'o (Not Rsd Is Nothing And Rsd.RecordCount > 0)
              
              BotonDetalles False, okDetV, okDetV, False
             
           '  puedeEditar = PuedeModificarComprobante(mMoviAfip)
           '  puedeAnular = PuedeAnularDirecto(mMoviAfip)
             
             HabilitarTodo False
             HabilitarDetalles False

            'Nuevo, Buscar, Salir + acciones de ver
           '  Botones True, True, True, puedeEditar, True, True
              
           '  CmdBotones(1).Enabled = puedeAnular

             CmdBotones(0).Caption = "Nuevo"
             CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
             CmdBotones(9).Caption = "Salir"
             CmdBotones(9).Picture = LoadResPicture("Salir", 0)
             CmdBotones(3).Caption = "Modificar"
             CmdBotones(3).Picture = LoadResPicture("Modificar", 0)

             Botones True, True, True, True, False, True
             BotonDetalles False, False, False, False
             BotonBuscar False, False
             
       Case stEditando
            Dim okCabE As Boolean: okCabE = CabeceraOK()
            Dim okDetE As Boolean: okDetE = DetalleOK()
            Dim okDetEditE As Boolean: okDetEditE = DetalleEditOK()
        
            ' ---- Cabecera: Modificar pasa a "Grabar", y Salir pasa a "Cancelar" ----
            CmdBotones(3).Caption = "Grabar"
            CmdBotones(3).Picture = LoadResPicture("Grabar", 0)
        
            CmdBotones(9).Caption = "Cancelar"
            CmdBotones(9).Picture = LoadResPicture("Cancelar", 0)
        
            ' ---- Habilitaciones generales ----
            ' Nuevo (0) lo podés deshabilitar en edición para no mezclar estados
            CmdBotones(0).Enabled = False
        
            ' Anular (1) NO mientras editás
            CmdBotones(1).Enabled = False
        
            ' Buscar (2) NO mientras editás
            CmdBotones(2).Enabled = False
        
            ' Modificar/Grabar (3) solo si cabecera + detalle OK y no estás editando un renglón
            If mDetEstado <> detIdle Then
                CmdBotones(3).Enabled = False
            Else
                CmdBotones(3).Enabled = (okCabE And okDetE)
            End If
        
            ' Imprimir (5) y demás acciones: deshabilitadas mientras editás
            CmdBotones(5).Enabled = False
            CmdBotones(4).Enabled = False
            CmdBotones(6).Enabled = False
            CmdBotones(7).Enabled = False
            CmdBotones(8).Enabled = False
        
            ' ---- Edición de controles ----
            HabilitarTodo True
        
            ' ---- Detalle (MISMA lógica que stNuevo) ----
            If mDetEstado = detIdle Then
                HabilitarDetalles False
                BotonBuscar True, False
        
               CmdDetalle(0).Caption = "Nuevo"
                CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
                CmdDetalle(0).Enabled = okCabE

                CmdDetalle(1).Enabled = okDetE
                CmdDetalle(2).Enabled = okDetE
                CmdDetalle(3).Enabled = False
            Else
                HabilitarTodo False
                HabilitarDetalles True
                BotonBuscar False, True
        
                CmdDetalle(0).Caption = "Grabar"
                CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
                CmdDetalle(0).Enabled = okDetEditE
        
                CmdDetalle(1).Enabled = False
                CmdDetalle(2).Enabled = False
                CmdDetalle(3).Enabled = True
            End If

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
    
      If mCodVentaTipo = 7 Then
          'PAGO MIXTO (VALORES)
          'No abrimos nada acá. Solo aseguramos que NO quede recargo pegado.
          nPorcValor = 0
          If Not (Rsd Is Nothing) Then
              If Rsd.RecordCount <> 0 Then RecalcularRsd
          End If
    
          'Opcional: si venías con RsValor cargado de antes, lo vaciamos para evitar basura
          ResetValores
    
      Else
          nPorcValor = RsFPago!RgTarjeta
          RecalcularRsd
      End If
  Else
       nPorcValor = 0
       RecalcularRsd
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
        .Append "FechaAcreditacion", adDate, , adFldIsNullable
        .Append "NroTarjeta", adVarChar, 30
        .Append "Vencimiento", adDate, , adFldIsNullable
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

    TxtProducto.text = Rsd!Producto & ""
    TxtDetalle.text = Rsd!Descripcion & ""
    TxtCantidad.text = Format(CDbl(Val(Rsd!Cantidad & "")), nDecimalCant)
    TxtPDesc.text = Format(CDbl(Val(Rsd!PDesc & "")), "#0.00")
    TxtPrecio.text = Format(CDbl(Val(Rsd!PrecioUnitFinal & "")), "#0.00")

    p.SetComboByItemData CmbCuenta, CLng(Val(Rsd!Cuenta & ""))
    p.SetComboByItemData CmbDeposito, CLng(Val(Rsd!Deposito & ""))
    p.SetComboByItemData CmbUnidad, CLng(Val(Rsd!Medida & ""))
    p.SetComboByItemData CmbImpuesto, CLng(Val(Rsd!ImpId & ""))

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
    Set p = New ClsPrograma

    CrearRsdMemoria

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
    tb.SelLength = Len(tb.text)
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If mLoadingDetalle Then Exit Sub
    If Rsd Is Nothing Then Exit Sub
    If (Rsd.BOF Or Rsd.EOF) Then Exit Sub

    If mEstado = stNuevo Or mEstado = stEditando And mDetEstado = detIdle Then
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

    TxtProducto.text = Rsd!Producto & ""
    TxtDetalle.text = Rsd!Descripcion & ""
    TxtCantidad.text = Format(Val(Rsd!Cantidad & ""), nDecimalCant)
    TxtPDesc.text = Format(Val(Rsd!PDesc & ""), "0.00")
    TxtPrecio.text = Format(Val(Rsd!precioBase & ""), "0.00")

    p.SetComboByItemData CmbDeposito, CLng(Val(Rsd!Deposito & ""))
    p.SetComboByItemData CmbUnidad, CLng(Val(Rsd!Medida & ""))
    p.SetComboByItemData CmbCuenta, CLng(Val(Rsd!Cuenta & ""))
    p.SetComboByItemData CmbImpuesto, CLng(Val(Rsd!ImpId & ""))
 

Salir:
    mLoadingDetalle = False
End Sub

Private Sub BindGridRsd()
    On Error Resume Next
    Set Grid1.DataSource = Nothing
    Set Grid1.DataSource = Rsd
    Grid1.ReBind
    Grid1.Refresh
    CabGrid
End Sub


Private Sub TxtCantidad_LostFocus()
  If TxtCantidad.text = "" Then TxtCantidad.text = 1
  TxtCantidad.text = Format(TxtCantidad.text, nDecimalCant)
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
  If TxtPrecio.text = "" Then TxtPrecio.text = 0
  TxtPrecio.text = Format(TxtPrecio.text, nCantDecimales)
End Sub

Private Sub TxtPDesc_LostFocus()
  If TxtPDesc.text = "" Then TxtPDesc.text = 0
  TxtPDesc.text = Format(TxtPDesc.text, "#0.00")
End Sub

Private Sub TxtNumero_Change()
  If mEstado = stNuevo Then RefrescarUI
End Sub

Private Sub LinkearProducto(pProducto As Long)
  Dim cRpl As ClsProductoL
  
  Set cRpl = New ClsProductoL
  
  If TxtProducto.text <> "" Or nDat <> 0 Then
     Dim RsProduc As ADODB.Recordset
     Set RsProduc = cRpl.TraerProducto(TxtProducto.text, pProducto)
     
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
        mprodActuaStock = RsProduc!ActuaStock
        
        TxtProducto.text = mProdCodigo
        TxtDetalle.text = mProdDescripcion
        
        p.SetComboByItemData CmbUnidad, mProdUM
        p.SetComboByItemData CmbDeposito, mProdDeposito
         CmbImpuesto.text = "IVA EXENTO"
         If Not RsComp Is Nothing Then
            If RsComp!Iva = 1 Then
              p.SetComboByItemData CmbImpuesto, mProdImpuesto
            End If
        End If
        If RsProduc!precio <> 0 Then
           TxtPrecio.text = Format(mProdPrecio, "#0.00")
           TxtCantidad.SetFocus
        Else
           TxtPrecio.text = "0.00"
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
        If TypeOf ctl Is TextBox Then ctl.text = ""
    Next ctl

    DtpFecha.Value = Date
    DtpVenc.Value = Date

    LblDescuento.Caption = "0.00"
    TxtReparto.text = "0000000000000"
    LblCliente.Caption = ""
    txtSucursal.text = "0000"
    TxtNumero.text = "00000000"
    TxtVendedor.text = "0.00"
    LbLIva.Caption = ""
    TxtCotizacion.text = "1.000"
    LblNeto.Caption = "0.00"
    LblBonificacion.Caption = "0.00"
    LblFinanciacion.Caption = "0.00"
    TxtNoGrav.text = "0.00"
    LblIva1.text = "0.00"
    LblPercepcion.Caption = "0.00"
    lblTotal.Caption = "0.00"

    CmbComprobante.ListIndex = -1
    CmbVend.text = "NINGUNO"
    CmbCorredor.text = "Ninguno"
    CmbLista.ListIndex = -1
    CmbCuenta.ListIndex = -1

    ' Forma pago / cond pago
    CmbFormaPago.text = IIf(nCondPago = 0, "CUENTA CORRIENTE", "CONTADO")
    CmbCondPago.text = "EFECTIVO"
      
End Sub

Private Sub LimpiarDetalles()
    TxtProducto.text = ""
    TxtDetalle.text = ""
    p.SetComboByItemData CmbUnidad, nMedDetalle
    p.SetComboByItemData CmbCuenta, nCueDetalle
    
    CmbImpuesto.text = "IVA EXENTO"
    If Not RsComp Is Nothing Then
        If RsComp!Iva = 1 Then
           p.SetComboByItemData CmbImpuesto, nImpDetalle
        End If
    End If
    TxtCantidad.text = Format(1, nDecimalCant)
    TxtDetalle.text = ""
    TxtPDesc.text = "0.00"
    TxtPrecio.text = Format(0, nCantDecimales) '<<< corregido
End Sub


'========================
'  BOTONES
'========================
Private Sub CmdBotones_Click(Index As Integer)
    Select Case Index
        Case 0
            If mEstado = stIdle Then
                Nuevo
            ElseIf mEstado = stNuevo Then
                GrabarComprobante
            End If
        Case 1 'ANULAR
            AnularComprobanteActual
        Case 2
            Buscar
        Case 3
             Modificar
        Case 5
             Imprimir
        Case 9
            Salir
    End Select
End Sub


Private Sub Nuevo()
    On Error GoTo errHandler
     
    If mEstado <> stNuevo Then
        mEstado = stNuevo

        mMoviAfip = 0   'opcional pero recomendable
         
        ResetDetalleMemoria
        ResetValores

        Limpiar
        LimpiarDetalles
        CalcularTotales
        RefrescarUI

        CmbCorredor.text = "Ninguno"
        p.SetComboByItemData CmbComprobante, nVentaFactura
        CmbComprobante.SetFocus


        CmbCorredor.text = "Ninguno"
        p.SetComboByItemData CmbComprobante, nVentaFactura
        CmbComprobante.SetFocus
    End If

    Exit Sub
errHandler:
    MsgBox err.Description, vbCritical, "Nuevo"
End Sub

Public Sub Buscar()
    Dim f As FrmBuscarComprobante

    Set f = New FrmBuscarComprobante
    f.TipoOperacionWhere = "1,5"
    Set f.OwnerForm = Me
    f.Show 1
             
    'NO lo descargues acá
End Sub


Private Sub Modificar()
    On Error GoTo errHandler

    Select Case mEstado
        Case stViendo
            mEstado = stEditando
            RefrescarUI

        Case stEditando
            'acá llamás a tu rutina de grabar modificación
            GuardarCambiosComprobante
            mEstado = stViendo
            RefrescarUI
    End Select
  Exit Sub
errHandler:
    ManejaErrores
End Sub

Private Sub Imprimir()
   Call ImprimirNuevo(Me, RsCli, RsComp, Rsd, nUsuario, sComputadora, mMoviAfip)
End Sub
Public Sub CargarComprobantePorMovimiento(ByVal pMov As Long)
    On Error GoTo errHandler

    MostrarComprobanteGrabado pMov
 
    mMoviAfip = pMov
    
    mEstado = stViendo
    mDetEstado = detIdle
 
    RefrescarUI
    
    Exit Sub

errHandler:
    MsgBox err.Description, vbCritical, "Cargar comprobante"
End Sub

Private Sub Salir()
   If mEstado = stNuevo Or mEstado = stEditando Then
      If MsgBox("¿Cancelar?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub

      If mEstado = stEditando Then
          'volver a ver el comprobante sin limpiar todo
          MostrarComprobanteGrabado mMoviAfip
          mEstado = stViendo
      Else
          mEstado = stIdle
          Limpiar
          LimpiarDetalles
          ResetDetalleMemoria
          ResetValores
      End If

      mEstado = stIdle
      
      mDetEstado = detIdle
      RefrescarUI
   Else
      Unload Me
   End If
End Sub


Private Sub GrabarComprobante()
    On Error GoTo errHandler

    '1) No permitir grabar si estás editando un renglón
    If mDetEstado <> detIdle Then
        MsgBox "Terminá de grabar/cancelar el renglón del detalle antes de grabar el comprobante.", vbExclamation, "Atención"
        Exit Sub
    End If

    '2) Validaciones
    If Not CabeceraOK() Then
        MsgBox "Faltan datos en la cabecera.", vbExclamation, "Atención"
        Exit Sub
    End If

    If Not DetalleOK() Then
        MsgBox "No hay renglones cargados en el detalle.", vbExclamation, "Atención"
        Exit Sub
    End If

    '3) Si es pago mixto (tipo 7), abrir Valores con el TOTAL FINAL
    
    If RsValor Is Nothing Then CrearRsValorMemoria
     
    If mCodVentaTipo = 7 Then
        Dim totalPagar As Double
        totalPagar = CDbl(Val(lblTotal.Caption)) 'esto ya incluye percepción en tu cálculo
       
        VentaValores.Inicializar RsValor, totalPagar
        VentaValores.Show vbModal

        'FrmValores debería setear una propiedad/flag si canceló
        If VentaValores.Cancelado Then
            'No grabar nada
            Exit Sub
        End If

        'Validar que sumen el total
        If Abs(SumaValores(RsValor) - Round(totalPagar, 2)) > 0.01 Then
            MsgBox "Los valores no suman el total a pagar. Revisá los importes.", vbExclamation, "Atención"
            Exit Sub
        End If
    Else
       VaciarRsValor
       
       RsValor.AddNew
       RsValor!FormaPagoId = CmbCondPago.ItemData(CmbCondPago.ListIndex)
       RsValor!FormaPagoDesc = CmbCondPago.text
       RsValor!importe = Round(CDbl(Val(lblTotal.Caption)), 2)

       RsValor!Nombre = LblCliente.Caption
       RsValor!BancoId = 0
    
       RsValor!BancoDesc = ""
       RsValor!NroCheque = ""
       RsValor!FechaAcreditacion = Null

       RsValor!NroTarjeta = ""
       RsValor!Vencimiento = Null
       RsValor!NroCupon = ""
       RsValor!NroAutorizacion = ""

       RsValor!Observaciones = ""
   
       RsValor.Update
    End If

    GrabarTodo
    
    mEstado = stIdle
    mDetEstado = detIdle
    
    RefrescarUI
    
    Botones True, True, True, True, False, True

    Exit Sub

errHandler:
    ManejaErrores
End Sub

Private Sub GuardarCambiosComprobante()
    On Error GoTo errHandler

    '1) Validaciones
    If mDetEstado <> detIdle Then
        MsgBox "Terminá de grabar/cancelar el renglón del detalle antes de grabar.", vbExclamation, "Atención"
        Exit Sub
    End If

    If Not CabeceraOK() Then
        MsgBox "Faltan datos en la cabecera.", vbExclamation, "Atención"
        Exit Sub
    End If

    If Not DetalleOK() Then
        MsgBox "No hay renglones cargados en el detalle.", vbExclamation, "Atención"
        Exit Sub
    End If

    '2) Armás TODOS los RsVacio igual que en GrabarTodo (pero con Movimiento existente)
    Dim RsCv As ADODB.Recordset, RsCd As ADODB.Recordset
    Dim rReci As ADODB.Recordset, rValo As ADODB.Recordset, rApli As ADODB.Recordset
    Dim rCant As ADODB.Recordset, rCaja As ADODB.Recordset

    Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
    Set RsCd = cRsl.RsVacio("DetallesComprobantes", "Id", "N")
    Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
    Set rValo = cRsl.RsVacio("Valores", "Id", "N")
    Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
    Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
    Set rCaja = cRsl.RsVacio("Caja", "Id", "N")

    '3) Cabecera (MISMO movimiento)
    RsCv.AddNew
    RsCv!Suc = nSucursal
    RsCv!tipo = RsComp!TipoOperacion
    RsCv!Movimiento = mMoviAfip   '<<< CLAVE: no se cambia
    RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
    RsCv!fecha = DtpFecha.Value
    RsCv!FechaIva = DtpVenc.Value
    RsCv!Numero = TxtNumero.text
    RsCv!sucursal = txtSucursal.text
    RsCv!Cliente = TxtCliente.text
    RsCv!lista = CmbLista.ItemData(CmbLista.ListIndex)
    RsCv!Vendedor = CmbVend.ItemData(CmbVend.ListIndex)
    RsCv!Caja = nCaja
    RsCv!CondVenta = CmbFormaPago.ItemData(CmbFormaPago.ListIndex)
    RsCv!neto = CCur(Val(LblNeto.Caption))
    RsCv!Iva1 = CCur(Val(LblIva1.text))
    RsCv!Iva2 = CCur(Val(LblDescuento.Caption))
    RsCv!COTIZACION = CCur(Val(TxtCotizacion.text))
    RsCv!Descuentos = CCur(Val(LblBonificacion.Caption))
    RsCv!Financiacion = CCur(Val(LblFinanciacion.Caption))
    RsCv!PIvaCompras = CCur(Val(LblPercepcion.Caption))
    RsCv!IBrutosCompras = CCur(Val(TxtVendedor.text))
    RsCv!NoGravados = CCur(Val(TxtNoGrav.text))

    Select Case RsComp!TipoMovimiento
        Case 1, 11
            RsCv!Debe = CCur(Val(lblTotal.Caption))
            RsCv!Haber = 0
        Case 2, 12, 13
            RsCv!Debe = 0
            RsCv!Haber = CCur(Val(lblTotal.Caption))
    End Select

    RsCv!Desde = CmbCondPago.ItemData(CmbCondPago.ListIndex)
    RsCv!Hasta = TxtReparto.text
    RsCv!factura = CmbCorredor.ItemData(CmbCorredor.ListIndex)
    RsCv!Motivo = ""
    RsCv!Anulado = 0
    RsCv.Update

    '4) Detalle desde tu Rsd (como ya lo tenés)
    Rsd.MoveFirst
    Do While Not Rsd.EOF
        RsCd.AddNew
        RsCd!Movimiento = mMoviAfip
        RsCd!Producto = Rsd!Producto
        RsCd!Descripcion = Rsd!Descripcion
        RsCd!Leyenda = ""
        RsCd!Cantidad = Rsd!Cantidad
        RsCd!Cuenta = Rsd!Cuenta
        RsCd!PDesc = Rsd!PDesc
        RsCd!Descuento = Rsd!desc
        RsCd!tasa = Rsd!ImpId
        RsCd!Deposito = Rsd!Deposito
        RsCd!Medida = Rsd!Medida
        RsCd!Impuesto = Rsd!ImpId
        RsCd!precio = Rsd!precioBase
        RsCd!PrecioUnitario = Rsd!PrecioUnitNeto + Rsd!desc
        RsCd!PrecioTotal = Rsd!Total + Rsd!desc
        RsCd!MoviRemi = 0
        RsCd!DetRemi = 0
        RsCd.Update
        Rsd.MoveNext
    Loop

    '5) Valores/Caja/Cantidades/CtaCte: armarlos igual que tu GrabarTodo actual
    '   (acá podés literalmente copiar y pegar tus bloques actuales)

    '6) Llamar a ModificarVenta
    Dim cGrab As ClsComprobantesE
    Set cGrab = New ClsComprobantesE
    cGrab.ModificarVenta RsCv, RsCd, rReci, rValo, rCant, rCaja

    '7) Volver a viendo
    mEstado = stViendo
    CmdBotones(3).Caption = "Modificar"
    CmdBotones(3).Picture = LoadResPicture("Modificar", 0)
    
    mEstado = stIdle
    mDetEstado = detIdle
    
    RefrescarUI

    MsgBox "Comprobante modificado.", vbInformation, "OK"
    Exit Sub

errHandler:
    ManejaErrores
End Sub


Private Sub MostrarComprobanteGrabado(ByVal movi As Long)
    On Error GoTo errHandler
    
    Dim sSql As String, cRiL As ClsClienteL
    Dim pvDefault As Integer
    pvDefault = Val(id_datos) ' o tu sucursal/punto de venta por defecto
    
    Set cRiL = New ClsClienteL
    
    sSql = ""
    sSql = sSql & "SELECT c.*, cli.RazonSocial, ti.Descripcion AS TipoIvaDesc " & vbCrLf
    sSql = sSql & "FROM CabComprobantes c " & vbCrLf
    sSql = sSql & "INNER JOIN Clientes cli ON c.Cliente = cli.Cliente " & vbCrLf
    sSql = sSql & "INNER JOIN TipoIva ti ON cli.TipoIva = ti.TipoIva " & vbCrLf
    sSql = sSql & "WHERE c.Movimiento = " & movi

    Set RsCv = cRsl.TraerRsSQL(sSql)
        
    Set RsCli = cRiL.TraerDatosCliente(RsCv!Cliente)
    
    sSql = "SELECT " & _
             "C.Id, UCV.Comprobante, UCV.Descripcion, UCV.Comanda, " & _
             "C.TipoMovimiento, C.TipoOperacion, C.Iva, C.Numero, C.CtaCte, C.Caja, C.Costo, C.MoviStock, " & _
             "C.Sucursal, C.Afip, C.Moneda, " & _
             "AD.Id_Datos, AD.Empresa, AD.Fantasia, AD.Cuit, AD.IIBB, AD.Iva AS IvaDesc, AD.Direccion, AD.Localidad, AD.Inicio, " & _
             "AD.Archivo, AD.Imagen, AD.Copias " & _
             "FROM UsuariosComprobantesVentas UCV " & _
             "LEFT JOIN Comprobantes C ON UCV.Comprobante = C.Id " & _
             "LEFT JOIN AfipDatos AD ON AD.Id_Datos = " & _
             "CASE WHEN ISNULL(C.Afip,0) > 0 THEN ISNULL(C.Sucursal," & pvDefault & ") ELSE " & pvDefault & " END " & _
             "WHERE UCV.Usuario = 1 AND UCV.Ver = 1 AND (C.TipoOperacion = 1 OR C.TipoOperacion = 5) " & _
             "AND C.Id=" & RsCv!Comprobante

      
    Set RsComp = cRsl.TraerRsSQL(sSql)
  
    If RsComp Is Nothing Or RsComp.RecordCount = 0 Then
        MsgBox "No se pudo cargar la definición del comprobante (Comprobantes.Id=" & RsCv!Comprobante & ").", vbCritical
        Exit Sub
    End If
      
    '2) Linkear cabecera a controles
    txtSucursal.text = Format$(RsCv!sucursal, "0000")
    TxtNumero.text = Format$(RsCv!Numero, "00000000")
    TxtCliente.text = RsCv!Cliente & ""
    
    LblCliente.Caption = RsCv!RazonSocial & ""
    LbLIva.Caption = RsCv!TipoIvaDesc & ""

    DtpFecha.Value = RsCv!fecha
    DtpVenc.Value = RsCv!FechaIva

    'si necesitás setear combos por ItemData:
    p.SetComboByItemData CmbComprobante, CLng(RsCv!Comprobante)
    p.SetComboByItemData CmbLista, CLng(RsCv!lista)
    p.SetComboByItemData CmbVend, CLng(RsCv!Vendedor)
    ' etc...

    '3) Traer detalle y cargarlo en tu Rsd (memoria) para mostrar en Grid
    ResetDetalleMemoria

    Dim rsDet As ADODB.Recordset
    Set rsDet = cRsl.TraerRsCondi("DetallesComprobantes", "Movimiento", "Movimiento=" & movi)

    If Not rsDet Is Nothing Then
        If rsDet.RecordCount > 0 Then
            rsDet.MoveFirst
            Do While Not rsDet.EOF
                Rsd.AddNew
                Rsd!Producto = rsDet!Producto & ""
                Rsd!Descripcion = rsDet!Descripcion & ""
                Rsd!Cantidad = CDbl(Val(rsDet!Cantidad & ""))
                Rsd!Cuenta = CLng(Val(rsDet!Cuenta & ""))
                Rsd!Medida = rsDet!Medida
                Rsd!Deposito = rsDet!Deposito
                Rsd!PDesc = CDbl(Val(rsDet!PDesc & ""))
                Rsd!desc = CDbl(Val(rsDet!Descuento & ""))
                Rsd!ImpId = rsDet!tasa
                Rsd!Impuesto = rsDet!Impuesto
                Rsd!neto = rsDet!PrecioUnitario
                Rsd!precioBase = CDbl(Val(rsDet!precio & ""))
                Rsd!PrecioUnitNeto = rsDet!PrecioUnitario
                Rsd!Total = rsDet!PrecioTotal
                'Acá mapear lo que tengas en DB a tu estructura en memoria
                'Rsd!PrecioUnitFinal / PrecioUnitNeto / Total, etc...
                Rsd.Update
                rsDet.MoveNext
            Loop
            
            If Rsd.RecordCount > 0 Then
                Rsd.MoveFirst
                BindGridRsd
                LinkearDetalleActual
            End If

        End If
    End If

    '4) Traer valores y dejarlos listos (si querés permitir reimprimir caja, etc.)
    If RsValor Is Nothing Then CrearRsValorMemoria Else VaciarRsValor

    Dim RSVal As ADODB.Recordset
    Set RSVal = cRsl.TraerRsCondi("Valores", "Venta", "Venta=" & movi)

    If Not RSVal Is Nothing Then
        If RSVal.RecordCount > 0 Then
            RSVal.MoveFirst
            Do While Not RSVal.EOF
                RsValor.AddNew
                RsValor!FormaPagoId = RSVal!CodPago
                RsValor!FormaPagoDesc = RSVal!formapago & ""
                RsValor!importe = CDbl(Val(RSVal!importe & ""))
                RsValor!Nombre = RSVal!Nombre & ""
                RsValor!BancoId = CLng(Val(RSVal!Banco & ""))
                RsValor!NroCheque = RSVal!NumeroCheque & ""
                'fechas/tarjeta según tu tabla
                RsValor.Update
                RSVal.MoveNext
            Loop
        End If
    End If

    '5) Recalcular totales y pintar labels (si tu CalcularTotales depende de Rsd)
    CalcularTotales
    
    mDetEstado = detIdle
  
    RefrescarUI
           
Exit Sub

errHandler:
    ManejaErrores
End Sub


Private Sub GrabarTodo()
  Dim RsCv As ADODB.Recordset
  Dim RsCd As ADODB.Recordset
  Dim rReci As ADODB.Recordset
  Dim rValo As ADODB.Recordset
  Dim rApli As ADODB.Recordset
  Dim rCant As ADODB.Recordset
  Dim rCaja As ADODB.Recordset
  Dim cComp As ClsComprobantesL
  Dim cGrab As ClsComprobantesE
  
  Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
  Set RsCd = cRsl.RsVacio("DetallesComprobantes", "Id", "N")
  Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
  Set rValo = cRsl.RsVacio("Valores", "Id", "N")
  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
  Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
  Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
  
  Set cComp = New ClsComprobantesL
  Set cGrab = New ClsComprobantesE
  
  ' Cabecera Comprobante
  RsCv.AddNew
  RsCv!Suc = nSucursal
'  RsCv!Id = 0
  RsCv!tipo = RsComp!TipoOperacion
  RsCv!Movimiento = cComp.TraerUltimoMovimiento
  
  mMoviAfip = RsCv!Movimiento
  
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DtpFecha.Value
  RsCv!FechaIva = DtpVenc.Value
  RsCv!Numero = TxtNumero.text
  RsCv!sucursal = txtSucursal.text
  RsCv!Cliente = TxtCliente.text
  RsCv!lista = CmbLista.ItemData(CmbLista.ListIndex)
  RsCv!Vendedor = CmbVend.ItemData(CmbVend.ListIndex)
  RsCv!Caja = nCaja
  RsCv!CondVenta = CmbFormaPago.ItemData(CmbFormaPago.ListIndex)
  RsCv!neto = CCur(Val(LblNeto.Caption))
  RsCv!Iva1 = CCur(Val(LblIva1.text))
  RsCv!Iva2 = CCur(Val(LblDescuento.Caption))
  RsCv!COTIZACION = CCur(Val(TxtCotizacion.text))
  RsCv!Descuentos = CCur(Val(LblBonificacion.Caption))
  RsCv!Financiacion = CCur(Val(LblFinanciacion.Caption))
  RsCv!PIvaCompras = CCur(Val(LblPercepcion.Caption))
  RsCv!IBrutosCompras = CCur(Val(TxtVendedor.text))
  RsCv!NoGravados = CCur(Val(TxtNoGrav.text))
  Select Case RsComp!TipoMovimiento
         Case 1, 11
              RsCv!Debe = CCur(Val(lblTotal.Caption))
              RsCv!Haber = 0
         Case 2, 12, 13
              RsCv!Debe = 0
              RsCv!Haber = CCur(Val(lblTotal.Caption))
  End Select
  RsCv!Desde = CmbCondPago.ItemData(CmbCondPago.ListIndex)
  RsCv!Hasta = TxtReparto.text
  RsCv!factura = CmbCorredor.ItemData(CmbCorredor.ListIndex)
  RsCv!Motivo = ""
  RsCv!Anulado = 0
  RsCv.Update
      
  'Detalle Comprobante
 
  If Not (Rsd Is Nothing) Then
      If Rsd.State = adStateOpen Then
          If Rsd.RecordCount > 0 Then
              Rsd.MoveFirst
              Do While Not Rsd.EOF
                  RsCd.AddNew
                  RsCd!Movimiento = mMoviAfip
                  RsCd!Producto = Rsd!Producto
                  RsCd!Descripcion = Rsd!Descripcion
                  RsCd!Leyenda = ""
                  RsCd!Cantidad = Rsd!Cantidad
                  RsCd!Cuenta = Rsd!Cuenta
                  RsCd!PDesc = Rsd!PDesc
                  RsCd!Descuento = Rsd!desc
                  RsCd!tasa = Rsd!ImpId        '<<< CONTRATO V0: tasa = ID impuesto
                  RsCd!Deposito = Rsd!Deposito
                  RsCd!Medida = Rsd!Medida
                  RsCd!Impuesto = Rsd!Impuesto
                  RsCd!precio = Rsd!precioBase
                  RsCd!PrecioUnitario = Rsd!PrecioUnitNeto + RsCd!Impuesto + Rsd!desc
                  RsCd!PrecioTotal = Rsd!Total + Rsd!desc
                  RsCd!DetRemi = 0
                  RsCd.Update
  
                  Rsd.MoveNext
              Loop
          End If
      End If
  End If

  
  ' Si es cuenta corriente
  If RsComp!ctacte <> "No" And CmbFormaPago.ItemData(CmbFormaPago.ListIndex) = 0 Then
      rReci.AddNew
      rReci!Venta = mMoviAfip
      rReci!fecha = DtpFecha.Value
      rReci!Cliente = TxtCliente.text
      rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
      rReci!Numero = TxtNumero.text
      rReci!sucursal = txtSucursal.text
      rReci!Cuota = 1
      Select Case RsComp!TipoMovimiento
            Case 1, 11
                rReci!Debe = CDbl(Val(lblTotal.Caption))
                rReci!Haber = 0
            Case 2, 12, 13
                 rReci!Debe = 0
                 rReci!Haber = CDbl(Val(lblTotal.Caption))
      End Select
      rReci!Anulado = 0
      rReci.Update
  End If
  
  ' Si Actualiza Stock
  If RsComp!MoviStock <> "No" And CmbFormaPago.text <> "COTIZAR" Then
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        If Rsd!ActuaStock = "Si" Then
            rCant.AddNew
            rCant!Suc = nSucursal
            rCant!Producto = Rsd!Producto
            rCant!Deposito = Rsd!Deposito
            rCant!Estado = 1
            Select Case RsComp!MoviStock
                   Case "Suma"
                        rCant!CantidadDebe = Rsd!Cantidad
                        rCant!CantidadHaber = 0
                   Case "Resta"
                        rCant!CantidadDebe = 0
                        rCant!CantidadHaber = Rsd!Cantidad
            End Select
            rCant!Anulado = 0
            rCant.Update
           End If
        Rsd.MoveNext
     Loop
  End If
  ' Valores
       
  If RsValor.RecordCount <> 0 Then
     RsValor.MoveFirst
     Do While Not RsValor.EOF
        rValo.AddNew
        rValo!Venta = mMoviAfip
        rValo!CodPago = RsValor!FormaPagoId
        rValo!formapago = RsValor!FormaPagoDesc
        rValo!Nombre = RsValor!Nombre
        rValo!Banco = RsValor!BancoId
        rValo!NumeroCheque = RsValor!NroCheque
        rValo!FechaAcreditacion = RsValor!FechaAcreditacion
        rValo!NumeroTarjeta = RsValor!NroTarjeta
        rValo!FechaVencimiento = RsValor!Vencimiento
        rValo!importe = RsValor!importe
        rValo.Update
        
        'Grabo la caja
        If RsComp!Caja <> "No" Then
           rCaja.AddNew
           rCaja!Suc = nSucursal
           rCaja!Movimiento = mMoviAfip
           rCaja!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
           rCaja!Numero = TxtNumero.text
           rCaja!sucursal = txtSucursal.text
           rCaja!fecha = DtpFecha.Value
           rCaja!Hora = Time
           rCaja!Caja = nCaja
           rCaja!formapago = RsValor!FormaPagoDesc
           rCaja!Nombre = RsValor!Nombre
           rCaja!Banco = RsValor!BancoId
           rCaja!NumeroCheque = RsValor!NroCheque
           rCaja!FechaAcreditacion = RsValor!FechaAcreditacion
           rCaja!NumeroTarjeta = RsValor!NroTarjeta
           rCaja!FechaVencimiento = RsValor!Vencimiento
           rCaja!Autorizacion = ""
           rCaja!Cupon = ""
           If RsComp!Caja = "Suma" Then
             rCaja!Debe = RsValor!importe
             rCaja!Haber = 0
           Else
              rCaja!Debe = 0
              rCaja!Haber = RsValor!importe
           End If
           rCaja!Abierta = "S"
           rCaja!Anulado = 0
           rCaja.Update
        End If
        
        RsValor.MoveNext
     Loop
  End If
  
  cGrab.Agregar RsCv, RsCd, rReci, rValo, rCant, rCaja
  
    
End Sub
Private Sub CmbFormaPago_Click()
   If CmbFormaPago.text <> "CONTADO" Then
      CmbCondPago.Enabled = False
      CmbCondPago.text = "Ninguno"
      nPorcValor = 0
      If Rsd.RecordCount <> 0 Then RecalcularRsd
   Else
      CmbCondPago.Enabled = True
      CmbCondPago.text = "EFECTIVO"
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
    SeleccionarTodo txtSucursal
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
  Grid1.HeadFont.bold = True
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
      .Columns(22).Visible = False
 End With
End Sub

Private Sub CargarCombos()
  On Error GoTo errHandler
  
 
  cRsl.CargaCombo CmbComprobante, "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
  cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmbCorredor, "Clientes", "Id", "RazonSocial", "Tipo=3"
     
  CmbCorredor.AddItem "Ninguno"
    
  cRsl.CargaCombo CmbImpuesto, "Impuestos", "Impuesto", "Descripcion", "Producto='1' OR Retencion='1' OR Percepcion='1' OR Interno='1'"
  cRsl.CargaCombo CmbCondPago, "CondVenta", "CondVta", "Descripcion", "Tipo<10 and tipo<>5"
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
  Dim nNum As Long, nSuc As Long, sSql As String
  On Error GoTo errHandler
  
  If CmbComprobante.ListIndex <> -1 Then
     If mEstado = stNuevo Then
     
       Dim pvDefault As Integer
       pvDefault = Val(id_datos) ' o tu sucursal/punto de venta por defecto

       sSql = "SELECT " & _
             "C.Id, UCV.Comprobante, UCV.Descripcion, UCV.Comanda, " & _
             "C.TipoMovimiento, C.TipoOperacion, C.Iva, C.Numero, C.CtaCte, C.Caja, C.Costo, C.MoviStock, " & _
             "C.Sucursal, C.Afip, C.Moneda, " & _
             "AD.Id_Datos, AD.Empresa, AD.Fantasia, AD.Cuit, AD.IIBB, AD.Iva AS IvaDesc, AD.Direccion, AD.Localidad, AD.Inicio, " & _
             "AD.Archivo, AD.Imagen, AD.Copias " & _
             "FROM UsuariosComprobantesVentas UCV " & _
             "LEFT JOIN Comprobantes C ON UCV.Comprobante = C.Id " & _
             "LEFT JOIN AfipDatos AD ON AD.Id_Datos = " & _
             "CASE WHEN ISNULL(C.Afip,0) > 0 THEN ISNULL(C.Sucursal," & pvDefault & ") ELSE " & pvDefault & " END " & _
             "WHERE UCV.Usuario = 1 AND UCV.Ver = 1 AND (C.TipoOperacion = 1 OR C.TipoOperacion = 5) " & _
             "AND C.Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)

      
       Set RsComp = cRsl.TraerRsSQL(sSql)
       
       If RsComp.RecordCount <> 0 Then
      
          mTipoOperacion = RsComp!TipoOperacion
          mTipoMovimiento = RsComp!TipoOperacion
          
          
          ' Si en impresion es false Numera cuando Termina de grabar el comprobante, si es verdadero actualiza el numero aca
          
          If cRsl.TraerValorDeUnCampo("UsuariosComprobantesVentas", "Comanda", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) = False Then
             ChkA4.Enabled = False
          Else
             ChkA4.Enabled = True
          End If
          
           nNum = RsComp!Numero + 1
           nSuc = RsComp!sucursal
               
           txtSucursal.text = Format(nSuc, "0000")
           TxtNumero.text = Format(nNum, "00000000")
          
          If Rsd.RecordCount <> 0 Then
             RecalcularRsd ' cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "CondVta=" & CmbCondPago.ItemData(CmbCondPago.ListIndex))
          End If
      End If
    End If
    If mEstado = stNuevo Then RefrescarUI
  End If
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
        nNoGrabado = nNoGrabado + CDbl(Val(Rsd!intTotal & ""))
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
    TxtNoGrav.text = Format(Round(nNoGrabado, 2), "#0.00")
    LblNeto.Caption = Format(Round(nNetoGravado, 2), "#0.00")
    LblBonificacion.Caption = Format(Round(nDes, 2), "#0.00")
    LblIva1.text = Format(Round(nIva, 2), "#0.00")
    LblPercepcion.Caption = Format(Round(nPercepcion, 2), "#0.00")
    lblTotal.Caption = Format(Round(nTot + nPercepcion, 2), "#0.00")
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
  If TxtCliente.text <> "" Or nDat <> 0 Then
       Set RsCli = cRiL.TraerDatosCliente(TxtCliente.text, nDat)
       If RsCli.RecordCount <> 0 Then
          TxtCliente.text = RsCli!Cliente
          mCliePercepcion = RsCli!Percepcion
        
         If IsNull(RsCli!FechaBaja) Or RsCli!FechaBaja = "" Then
            LblCliente.Caption = RsCli!RazonSocial
            LbLIva.Caption = RsCli!TivaDesc
            p.SetComboByItemData CmbLista, RsCli!ListaPrecio
            LblDescuento.Caption = Format(RsCli!DescPorc, "0.00")
            If CmbVend.text = "NINGUNO" Then
               CmbVend.text = RsCli!VendDesc
               TxtVendedor.text = Format(RsCli!VendComision, "0.00")
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
                mLoadingDetalle = True
                mDetEstado = detnuevo
                LimpiarDetalles
                mLoadingDetalle = False
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
          'Si estabas creando un renglón nuevo (todavía no existe en el RS),
            mLoadingDetalle = True
        
            mDetEstado = detIdle
            If Rsd Is Nothing Or Rsd.RecordCount <= 0 Or (Rsd.BOF And Rsd.EOF) Then
                LimpiarDetalles
            Else
                If (Rsd.BOF Or Rsd.EOF) Then Rsd.MoveFirst
                LinkearDetalleActual
            End If
        
            mLoadingDetalle = False
            RefrescarUI
        
            On Error Resume Next
            If CmdDetalle(0).Enabled Then CmdDetalle(0).SetFocus
            err.Clear

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
    TxtProducto.text = Rsd!Producto & ""
    TxtDetalle.text = Rsd!Descripcion & ""

    p.SetComboByItemData CmbCuenta, CLng(Val(Rsd!Cuenta & ""))
    p.SetComboByItemData CmbDeposito, CLng(Val(Rsd!Deposito & ""))
    p.SetComboByItemData CmbUnidad, CLng(Val(Rsd!Medida & ""))
    p.SetComboByItemData CmbImpuesto, CLng(Val(Rsd!ImpId & ""))

    TxtCantidad.text = Format(CDbl(Val(Rsd!Cantidad & "")), nDecimalCant)
    TxtPDesc.text = Format(CDbl(Val(Rsd!PDesc & "")), "0.00")

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

    Rsd!Producto = TxtProducto.text
    Rsd!Descripcion = TxtDetalle.text
    Rsd!Cuenta = CmbCuenta.ItemData(CmbCuenta.ListIndex)
    Rsd!Deposito = CmbDeposito.ItemData(CmbDeposito.ListIndex)
    Rsd!Medida = CmbUnidad.ItemData(CmbUnidad.ListIndex)
    Rsd!Cantidad = CDbl(Val(TxtCantidad.text))
    Rsd!PDesc = CDbl(Val(TxtPDesc.text))
    Rsd!ImpProduc = mProdTasa
    Rsd!precioBase = CDbl(Val(TxtPrecio.text))

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
    
    mLoadingDetalle = True

    Rsd.AddNew
    Rsd!Producto = TxtProducto.text
    Rsd!Descripcion = TxtDetalle.text
    Rsd!Cuenta = CmbCuenta.ItemData(CmbCuenta.ListIndex)
    Rsd!Deposito = CmbDeposito.ItemData(CmbDeposito.ListIndex)
    Rsd!Medida = CmbUnidad.ItemData(CmbUnidad.ListIndex)
    Rsd!Cantidad = CDbl(Val(TxtCantidad.text))
    Rsd!PDesc = CDbl(Val(TxtPDesc.text))
    Rsd!ImpProduc = mProdTasa
    Rsd!precioBase = CDbl(Val(TxtPrecio.text))

    Rsd!ImpId = CmbImpuesto.ItemData(CmbImpuesto.ListIndex)
    Rsd!ImpIncluido = IIf(mListaPrecio = 1, 1, 0)
    Rsd!ActuaStock = mprodActuaStock

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
    mLoadingDetalle = False
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
    netoUnit = Round(netoUnit, 4)
    finalUnit = Round(finalUnit, 4)

    neto = Round(neto, 4)
    tot = Round(tot, 4)
    impu = Round(impu, 4)
    descMonto = Round(descMonto, 4)
    intUnit = Round(intUnit, 4)
    intTot = Round(intTot, 4)

    '--- 7) asignar
    R!PrecioUnitNeto = netoUnit
    R!PrecioUnitFinal = finalUnit
    R!neto = neto
    R!Impuesto = impu
    R!desc = descMonto

    R!intUnit = intUnit
    R!intTotal = intTot

    R!Total = tot 'tot ya incluye interno porque finalUnit lo incluye

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
    EsNegro = (Val(RsComp!Iva & "") = 0)
End Function

Private Sub ResetDetalleMemoria()
    If Not Rsd Is Nothing Then
        If Rsd.State = adStateOpen Then Rsd.Close
    End If
    CrearRsdMemoria
    Set Grid1.DataSource = Rsd
    CabGrid
End Sub

Private Function TieneCae(ByVal movi As Long) As Boolean
    On Error GoTo errHandler
    TieneCae = (Val(cRsl.TraerValorDeUnCampo("AfipCaes", "Movimiento", "Movimiento=" & movi)) <> 0)
    Exit Function
errHandler:
    TieneCae = False
End Function

Private Function PuedeModificarComprobante(ByVal movi As Long) As Boolean
    'Negro => SI se puede (salvo permisos futuros)
    If EsNegro() Then
        PuedeModificarComprobante = True
        Exit Function
    End If

    'Blanco => si tiene CAE NO se puede
    If TieneCae(movi) Then
        PuedeModificarComprobante = False
    Else
        PuedeModificarComprobante = True
    End If

    'Futuro: permiso usuario
    'If Not UsuarioPuedeModificar Then PuedeModificarComprobante = False
End Function

Private Function PuedeAnularDirecto(ByVal movi As Long) As Boolean
    'Negro: siempre directo
    If EsNegro() Then
        PuedeAnularDirecto = True
        Exit Function
    End If

    'Blanco: si tiene CAE => NO directo
    PuedeAnularDirecto = (Not TieneCae(movi))
End Function

Private Sub AnularComprobanteActual()
    On Error GoTo errHandler

   ' If mEstado <> stViendo Then Exit Sub

    If Not PuedeAnularDirecto(mMoviAfip) Then
        MsgBox "Este comprobante es fiscal y ya tiene CAE." & vbCrLf & _
               "No se puede anular directo. Se anula con Nota de Crédito.", vbExclamation, "Atención"
        Exit Sub
    End If

    If MsgBox("¿Confirmás anular el comprobante " & mMoviAfip & "?", _
              vbYesNo + vbQuestion, "Anular") = vbNo Then Exit Sub

    '--- Borrado en DB
    Dim cGrab As ClsComprobantesE
    Set cGrab = New ClsComprobantesE

    cGrab.AnularComprobantes mMoviAfip, mTipoOperacion, mTipoMovimiento

    Set cGrab = Nothing

    MsgBox "Comprobante anulado.", vbInformation, "OK"

    '========================
    '  DEJAR FORM EN BLANCO
    '========================
    mMoviAfip = 0

    mEstado = stIdle
    mDetEstado = detIdle

    Limpiar
    LimpiarDetalles

    ResetDetalleMemoria   'recrea Rsd en memoria y lo linkea al grid
    ResetValores          'vacía RsValor / nPorcValor, etc.

    RefrescarUI
    Exit Sub

errHandler:
    On Error Resume Next
    Set cGrab = Nothing
    ManejaErrores
End Sub


Public Function Autorizar(pCompro As Byte, pConcepto As Byte, pTipoDoc As Byte, pNumDoc As String, pTotal As String, pNeto As String, _
   pIva As String, pPunto As Integer, pMovi As Long, pCosto As Byte) As Boolean
   
    Dim cRle As ClsEscritura
    
    Set cRle = New ClsEscritura
    Set cRsl = New ClsLectura

    Dim WSAA As Object, WSFEv1 As Object
    Dim ttl As Long, tra As String, Path As String, certificado As String, ClavePrivada As String
    Dim cms As String, cacert As String, wsdl As String
    Dim cache As String, wrapper As String, proxy As String, ok As Boolean, ta As String
    Dim tipo_cbte As Byte, punto_vta As Integer, cbte_nro As String, v As Variant, fecha As String
    Dim concepto As Byte, tipo_doc As Byte, nro_doc As String, cbt_desde As Long, cbt_hasta As Long
    Dim imp_total As String, imp_tot_conc As String, imp_neto As String, imp_iva As String, imp_trib As String
    Dim imp_op_ex As String, fecha_cbte As String, fecha_venc_pago As String
    Dim fecha_serv_desde As String, fecha_serv_hasta As String, moneda_id As String, moneda_ctz As String
    Dim evento As Variant, token As String, Sign As String
    Dim ID As Byte, base_imp As String, importe As String, CAE As String, cae2 As String
    Dim fd As String, Excepcion As String
    
    On Error GoTo ManejoError
    
    
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.version
    If WSAA.version < "2.02c" Then
       MsgBox "Debe instalar una versión más actualizada de PyAfipWs WSAA!"
       End
    End If
            
   ta = cRsl.TraerValorDeUnCampo("Sistema", "Ta", "")
   If ta <> "" Then
      ok = WSAA.AnalizarXml(ta)
       If Not WSAA.Expirado() Then
           token = WSAA.ObtenerTagXml("token")
           Sign = WSAA.ObtenerTagXml("sign")
       End If
   End If
            
   If token = "" Or Sign = "" Then
      
          ' deshabilito errores no manejados (version 2.04 o superior)
       WSAA.LanzarExcepciones = False
              
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
      ttl = 43200 ' tiempo de vida = 12hs hasta expiración
      tra = WSAA.CreateTRA("wsfe", ttl)
      ControlarExcepcion WSAA
      Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    
      Path = App.Path + "\"
 
      certificado = sCertificado '"dagsis.crt" ' certificado de prueba
      ClavePrivada = sClavePrivada '"dagsis.key" ' clave privada de prueba
        
    ' Generar el mensaje firmado (CMS)
      cms = WSAA.SignTRA(tra, Path + certificado, Path + ClavePrivada)
      ControlarExcepcion WSAA
      Debug.Print cms
    
    ' Conectarse con el webservice de autenticación:
      cache = ""
      proxy = "" '"usuario:clave@localhost:8000"
      wrapper = ""
      cacert = ""
          
      wsdl = sWebServiceLogin
   
      ok = WSAA.Conectar(cache, wsdl, proxy, wrapper, cacert) ' Homologación
      ControlarExcepcion WSAA
    
    ' Llamar al web service para autenticar:
      ta = WSAA.LoginCMS(cms)
      ControlarExcepcion WSAA

    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
      Debug.Print ta
      Debug.Print "Token:", WSAA.token
      Debug.Print "Sign:", WSAA.Sign
    
      If ta <> "" Then
           cRle.GrabarValordeUnCampo "Sistema", "Ta", "'" & ta & "'", ""
      End If
      token = WSAA.token
      Sign = WSAA.Sign

    End If
    
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.version
    If WSAA.version < "1.12" Then
        MsgBox "Debe instalar una versión mas actualizada de PyAfipWs WSFEv1!"
        End
    End If
    'Debug.Print WSFEv1.InstallDir

    ' Setear tocken y sing de autorización (pasos previos)
    WSFEv1.token = token ' WSAA.Token
    WSFEv1.Sign = Sign ' WSAA.Sign

    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = sCuitAfip
    ' deshabilito errores no manejados
    WSFEv1.LanzarExcepciones = False

    ' Conectar al Servicio Web de Facturación
    proxy = "" ' "usuario:clave@localhost:8000"
    proxy = ""

    wsdl = Trim(sWebServiceCae)

   ' wsdl = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL"

    cache = "" 'Path
    wrapper = "" ' libreria http (httplib2, urllib2, pycurl)
    cacert = "" ' WSAA.InstallDir & "\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)"

    ok = WSFEv1.Conectar(cache, wsdl, proxy, wrapper, cacert) ' homologación
    Debug.Print WSFEv1.version
    ControlarExcepcion WSFEv1

    ' mostrar bitácora de depuración:
    Debug.Print WSFEv1.DebugLog

    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    ControlarExcepcion WSFEv1
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus

 
   
   
    ' Establezco los valores de la factura a autorizar:
    tipo_cbte = pCompro '6
    punto_vta = pPunto
    
    cbte_nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta)
    ControlarExcepcion WSFEv1
    
    For Each v In WSFEv1.Errores
        Debug.Print v
    Next
    Debug.Print WSFEv1.errmsg
    Debug.Print WSFEv1.errcode
    
    If cbte_nro = "" Then
        cbte_nro = 0                ' no hay comprobantes emitidos
    Else
        cbte_nro = CLng(cbte_nro)   ' convertir a entero largo
    End If
    Dim sDate As Date
    sDate = DtpFecha.Value
    
    fecha = Format(sDate, "yyyymmdd")
    
    concepto = pConcepto
    tipo_doc = pTipoDoc: nro_doc = pNumDoc
    
    cbte_nro = cbte_nro + 1
    cbt_desde = cbte_nro: cbt_hasta = cbte_nro
 
   ' Factura C
   ' pIva = 3.4
    imp_total = pTotal: imp_tot_conc = "0.00": imp_neto = Format(pNeto, "0.00")  '16.22
    imp_iva = Format(pIva, "0.00"): imp_trib = LblPercepcion.Caption: imp_op_ex = "0.00"


    fecha_cbte = fecha: fecha_venc_pago = ""
    ' Fechas del período del servicio facturado (solo si concepto = 1?)
    If pConcepto = 1 Then
       fecha_serv_desde = "": fecha_serv_hasta = ""
    Else
       fecha_serv_desde = fecha: fecha_serv_hasta = fecha
       fecha_venc_pago = fecha
    End If
    moneda_id = "PES": moneda_ctz = "1.000"
    
   ' moneda_id = "DOL": moneda_ctz = "850.00"
    imp_tot_conc = TxtNoGrav.text
    
    
    Dim cRcli As ClsClienteL
    Set cRcli = New ClsClienteL
    
    Dim pTipoReceptor As Integer
    pTipoReceptor = cRcli.BuscarTipoIvaReceptor(TxtCliente.text)


    ok = WSFEv1.CrearFactura(concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta, _
        moneda_id, moneda_ctz, "N", 1)
        
    imp_iva_fact = imp_iva
    imp_tributo_fact = imp_trib


     If tipo_cbte = 201 Or tipo_cbte = 203 Then
        'if tipo_cbte in (203, 208, 213):
'
        ok = WSFEv1.AgregarOpcional(2101, cRsl.TraerValorDeUnCampo("Sistema", "Cbu", "")) ' CBU
        ok = WSFEv1.AgregarOpcional(2102, "pyafipws")                ' alias
        If tipo_cbte = 201 Then
           ok = WSFEv1.AgregarOpcional(27, "SCA")
        Else
           ok = WSFEv1.AgregarOpcional(22, "S")
        End If
     End If
        
    ' Agrego los comprobantes asociados:
    If tipo_cbte = 3 Or tipo_cbte = 2 Or tipo_cbte = 7 Or tipo_cbte = 8 Or tipo_cbte = 12 Or tipo_cbte = 13 Or tipo_cbte = 212 Or tipo_cbte = 203 Then   ' solo nc/nd
        Dim nNumeroAsoc As Integer, rAsoc As Recordset, tipoAsoc, puntoAsoc As Integer, numeroAsoc As Long
        nLlama = 9
        ComproBuscar.Show 1
        If ComproBuscar.nId <> 0 Then
           Set rAsoc = cRsl.TraerRS("CabezaTraerUno", ComproBuscar.nId, True)
           
           tipoAsoc = cRsl.TraerValorDeUnCampo("Comprobantes", "Afip", "Id=" & rAsoc!Comprobante)
           puntoAsoc = rAsoc!sucursal
           numeroAsoc = rAsoc!Numero
          ok = WSFEv1.AgregarCmpAsoc(tipoAsoc, puntoAsoc, numeroAsoc, sCuitAfip, Format(rAsoc!fecha, "yyyymmdd"))
       End If
    End If
'
'    ' Agrego impuestos varios
'    id = 99
'    Desc = "Impuesto Municipal Matanza'"
'    base_imp = "100.00"
'    alic = "0.10"
'    importe = "0.10"
'    ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)
'
'    ' Agrego impuestos varios

   Dim desc As String, nPerc As Currency, alic

   
'    Id = 4
'    desc = "Impuestos internos"
'    base_imp = imp_neto
'    alic = "1.00"
'    importe = "3168.90"
'    ok = WSFEv1.AgregarTributo(Id, desc, base_imp, alic, importe)
'
'    ' Agrego impuestos varios
'    id = 1
'    Desc = "Impuesto nacional"
'    base_imp = "50.00"
'    alic = "1.00"
'    importe = "0.50"
'    ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)
'
     Dim i As Byte
     
     Dim nTasa(3) As Byte
     Dim sBaseCam(3) As Currency, sImporte(3) As Currency
     
     
     Rsd.MoveFirst
    
    
      'Agrego impuestos varios
    If LblPercepcion.Caption <> 0 Then
       ID = 2
       nPerc = cRsl.TraerValorDeUnCampo("Clientes", "Percepcion", "Cliente='" & TxtCliente.text & "'")
       ok = WSFEv1.AgregarTributo(ID, "Perc.Ing.Bruto Bs.As.", LblNeto.Caption, str(nPerc), LblPercepcion.Caption)
    End If
    
    If bMono = False Then
        Do While Not Rsd.EOF
           Select Case Rsd!ImpId
               Case 9
                  nTasa(0) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId)
                  sBaseCam(0) = sBaseCam(0) + (Rsd!Total - Rsd!desc)
                  sImporte(0) = sImporte(0) + Rsd!Impuesto
               Case 10
                  nTasa(1) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId)
                  sBaseCam(1) = sBaseCam(1) + (Rsd!Total - Rsd!Des)
                  sImporte(1) = sImporte(1) + Rsd!Impuesto
               Case 15
                  nTasa(2) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId)
                  sBaseCam(2) = sBaseCam(2) + (Rsd!Total - Rsd!desc)
                  sImporte(2) = sImporte(2) + Rsd!Impuesto
               Case 16
                  nTasa(3) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId)
                  sBaseCam(3) = sBaseCam(3) + (Rsd!Total - Rsd!Descuento)
                  sImporte(3) = sImporte(3) + Rsd!Impuesto
           End Select
           Rsd.MoveNext
        Loop
        For i = 0 To 3
            If sImporte(i) <> 0 Or nTasa(i) = 3 Then
               ID = nTasa(i)
               base_imp = Round(sBaseCam(i), 2) - IIf(pCosto = 0, sImporte(i), 0)
               base_imp = Round(base_imp, 2)
               importe = Format(sImporte(i), "#0.00")
               ok = WSFEv1.AgregarIva(ID, base_imp, importe)
              ' ok = WSFEv1.AgregarIva(Id, 80, 18.9)
            End If
        Next
     End If

   
    ' Habilito reprocesamiento automático (predeterminado):
    WSFEv1.Reprocesar = True
    
    ok = WSFEv1.EstablecerCampoFactura("cancela_misma_moneda_ext", "N")
    ok = WSFEv1.EstablecerCampoFactura("condicion_iva_receptor_id", pTipoReceptor)

    ' Solicito CAE:
    CAE = WSFEv1.CAESolicitar()
    ControlarExcepcion WSFEv1

    MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & CAE & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly

    ' Muestro los errores
    If WSFEv1.errmsg <> "" Then
        MsgBox WSFEv1.errmsg, vbExclamation, "Error"
    End If


    If CAE = "" And tipo_cbte <> 203 Then
        ' hubo error, no comparo
    Else
        Set cRle = New ClsEscritura
        
        Dim rCae As ADODB.Recordset
        Set rCae = New ADODB.Recordset
        
        rCae.Fields.Append "Movimiento", adInteger
        rCae.Fields.Append "Fecha", adVarChar, 10
        rCae.Fields.Append "CAE", adVarChar, 50
        rCae.Fields.Append "Documento", adVarChar, 20
        rCae.Fields.Append "Comprobante", adInteger
        rCae.Fields.Append "Numero", adInteger
        rCae.Open
        
        rCae.AddNew
        rCae!Movimiento = mMoviAfip
        rCae!fecha = IIf(IsNull(WSFEv1.Vencimiento), Date, WSFEv1.Vencimiento)
        rCae!CAE = CAE
        rCae!Documento = pNumDoc
        rCae!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
        rCae!Numero = cbte_nro
        rCae.Update
        
        TxtNumero.text = Format(cbte_nro, "00000000")
        
        cRle.Actualizar "AfipCaes_A", rCae
        Botones True, True, True, True, False, True
        Autorizar = True
    End If
Exit Function
ManejoError:
    ' Si hubo error (tradicional, no controlado):
    
    ' Depuración (grabar a un archivo los detalles del error)
    fd = FreeFile
    Open "c:\error.txt" For Append As fd
    If Not WSAA Is Nothing Then
        If WSAA.version >= "1.02a" Then
            Print #fd, WSAA.Excepcion
            Print #fd, WSAA.Traceback
            Print #fd, WSAA.XmlRequest
            Print #fd, WSAA.XmlResponse
            ' guardo mensaje de error para mostrarlo:
            Excepcion = WSAA.Excepcion
        End If
    End If
    If Not WSFEv1 Is Nothing Then
        If WSFEv1.version >= "1.10a" Then
            Print #fd, WSFEv1.Excepcion
            Print #fd, WSFEv1.Traceback
            Print #fd, WSFEv1.XmlRequest
            Print #fd, WSFEv1.XmlResponse
            Print #fd, WSFEv1.DebugLog()
            ' guardo mensaje de error para mostrarlo:
            Excepcion = WSFEv1.Excepcion
        End If
    End If
    Close fd
    
    Debug.Print err.Description            ' descripción error afip
    Debug.Print err.Number - vbObjectError ' codigo error afip
    If Excepcion = "" Then                 ' si no tengo mensaje de excepcion
        Excepcion = err.Description        ' uso el error de VB
    End If
    
    ' Mostrar el mensaje de error
    Select Case MsgBox(Excepcion, vbCritical + vbRetryCancel, "Error:" & err.Number - vbObjectError & " en " & err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print err.Description
    End Select
End Function

Sub ControlarExcepcion(obj As Object)
    Dim fd As String

    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0
    If obj.Excepcion <> "" Then
        ' Depuración (grabar a un archivo los detalles del error)
        fd = FreeFile
        Open "c:\excepcion.txt" For Append As fd
        Print #fd, obj.Excepcion
        Print #fd, obj.Traceback
        Print #fd, obj.XmlRequest
        Print #fd, obj.XmlResponse
        Close fd
        MsgBox "Servidor de AFIP Caido. Intentelo Mas tarde....", vbCritical, "Atención"
        End
    End If
End Sub

Public Function ImprimirPdf(pCompro As Byte, pConcepto As Byte, pTipoDoc As Byte, pNumDoc As String, pTotal As String, pNeto As String, _
   pIva As String, pValIva() As String, pPorIva() As Byte, pPunto As Integer, pMovi As Long, pDire As String, pLocalidad As String, _
   pProvincia As String, pCosto As Byte) As Boolean


'      Dim PyEmail As Object
'
'      Set PyEmail = CreateObject("PyEmail")
'
'     ' Primer paso: conexión al servidor (por unica vez)
'     servidor = "mail.sistemasagiles.com.ar"
'     usuario = "no.responder@nsis.com.ar"
'     clave = "1238478"
'     ok = PyEmail.Conectar(servidor, usuario, clave)
'
'     ' Envio el o los correos (repetir por cada FE)
'     remitente = "no.responder@sistemasagiles.com.ar"
'     destinatario = "mariano@sistemasagiles.com.ar"
'     mensaje = "Se envia factura electronica adjunta"
'     archivo = "C:\FACTURA.PDF"
'
'     ok = PyEmail.Enviar(remitente, motivo, destinatario, mensaje, archivo)

    Dim PyFEPDF As Object
    
    Dim domicilio_cliente  As String, nombre_cliente As String, Path As String
    Dim cms As String, cacert As String, wsdl As String, pais_dst_cmp As String, Descuento As String, imp_subtotal As String
    Dim ok As Variant, id_impositivo As String
    Dim tipo_cbte As Byte, punto_vta As Integer, cbte_nro As String, fecha As String
    Dim concepto As Byte, tipo_doc As Byte, nro_doc As String
    Dim imp_total As String, imp_tot_conc As String, imp_neto As String, imp_iva As String, imp_trib As String
    Dim imp_op_ex As String, fecha_cbte As String, fecha_venc_pago As String
    Dim fecha_serv_desde As String, fecha_serv_hasta As String, moneda_id As String, moneda_ctz As String
    Dim ID As Byte, base_imp As String, importe As String, CAE As String
    Dim obs_generales As String, obs_comerciales As String, motivo_obs As String, forma_pago As String, fecha_vto_cae As String, incoterms As String
    Dim idioma_cbte As String
    Dim iva_id As Byte, u_mtx  As String, codigo As String, ds As String
    Dim qty As String, umed As Byte, precio As String, bonif As String, despacho As String
    Dim dato_a, dato_b, dato_c, dato_d, dato_e As String, cod_mtx As String
    Dim papel As String, orientacion As String, num_copias As Byte, lineas_max As Integer, qty_pos As String, salida As String, Imprimir As Boolean
    Dim cRsl As ClsLectura, i As Byte
    Dim cSal As ClsClienteL
    
    Set cRsl = New ClsLectura
    Set cSal = New ClsClienteL
    
    On Error GoTo ManejoError
    
    ' Crear objeto interface para generación de F.E. en PDF
    Set PyFEPDF = CreateObject("PyFEPDF")
    Debug.Print PyFEPDF.version
    Debug.Print PyFEPDF.InstallDir
        
        
    Path = App.Path
        
    ' CUIT del emisor
    PyFEPDF.Cuit = sCuitAfip
    
    tipo_cbte = pCompro         ' Factura A
    punto_vta = pPunto    ' prefijo
    cbte_nro = Val(TxtNumero.text) ' número de factura
    fecha = DtpFecha.Value
    concepto = pConcepto
    ' datos del cliente:
    
    
    tipo_doc = pTipoDoc: nro_doc = cRsl.TraerValorDeUnCampo("AfipCaes", "Documento", "Movimiento=" & pMovi)
    nombre_cliente = TxtCliente.text & " - " & LblCliente.Caption
    
    domicilio_cliente = pDire
    pais_dst_cmp = 16 ' código para exportación
    id_impositivo = LbLIva.Caption
    ' totales del comprobante:
    imp_total = pTotal: imp_tot_conc = "0.00"
    imp_neto = Format(pNeto, "0.00"): imp_iva = Format(pIva, "0.00")
    imp_trib = "0.00": imp_op_ex = "0.00": imp_subtotal = "0.00"
    Descuento = "0.00"
    fecha_cbte = fecha: fecha_venc_pago = fecha
    ' Fechas del período del servicio facturado
    fecha_serv_desde = fecha: fecha_serv_hasta = fecha
    moneda_id = "PES": moneda_ctz = "1.000"
    
   ' moneda_id = "DOL": moneda_ctz = "850.00"
    obs_generales = ""
    obs_comerciales = ""
   ' moneda_id = "012"
   ' moneda_ctz = 0.5
    forma_pago = IIf(CmbFormaPago.text = "COTIZAR", "CONTADO", CmbFormaPago.text)
    incoterms = "FOB" ' termino de comercio exterior para exportación
    idioma_cbte = 1   ' idioma para exportación (no usado por el momento)
    ' motivo de observación (F136 y otros - RG2485/08 Art. 30 inc. c):
    motivo_obs = ""

    ' Código de Autorización Electrónica y fecha de vencimiento:
    ' (para facturas tradicionales, no imprimir el CAE ni código de barras)
    CAE = cRsl.TraerValorDeUnCampo("AfipCaes", "Cae", "Movimiento=" & pMovi)
    fecha_vto_cae = cRsl.TraerValorDeUnCampo("AfipCaes", "Fecha", "Movimiento=" & pMovi)
    
    
    
    ' Creo la factura (internamente en la interfaz)
    ok = PyFEPDF.CrearFactura( _
        concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbte_nro, imp_total, imp_tot_conc, imp_neto, _
        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta, _
        moneda_id, moneda_ctz, CAE, fecha_vto_cae, id_impositivo, _
        nombre_cliente, domicilio_cliente, pais_dst_cmp, _
        obs_comerciales, obs_generales, forma_pago, incoterms, _
        idioma_cbte, motivo_obs, Descuento)
    
    
'    imp_iva_fact = imp_iva
'    imp_tributo_fact = imp_trib
    
    ' Agregar comprobantes asociados (si es una NC/ND):
    'tipo = 19
    'pto_vta = 2
    'nro = 1234
    'pyfepdf.AgregarCmpAsoc(tipo, pto_vta, nro)
    
     ' Path = CurDir() + "\"
     Path = App.Path + "\"
    ' Agrego subtotales de IVA (uno por alicuot,codigo )
    
     Dim nTasa(3) As Byte
     Dim sBaseCam(3) As Currency, sImporte(3) As Currency
    
     Rsd.MoveFirst
    
    If bMono = False Then
         Do While Not Rsd.EOF
           Select Case Rsd!ImpId
               Case 9
                  nTasa(0) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId)
                  sBaseCam(0) = sBaseCam(0) + (Rsd!Total - Rsd!desc)
                  sImporte(0) = sImporte(0) + Rsd!Impuesto
               Case 10
                  nTasa(1) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId)
                  sBaseCam(1) = sBaseCam(1) + (Rsd!Total - Rsd!Des)
                  sImporte(1) = sImporte(1) + Rsd!Impuesto
               Case 15
                  nTasa(2) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId)
                  sBaseCam(2) = sBaseCam(2) + (Rsd!Total - Rsd!desc)
                  sImporte(2) = sImporte(2) + Rsd!Impuesto
               Case 16
                  nTasa(3) = cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId)
                  sBaseCam(3) = sBaseCam(3) + (Rsd!Total - Rsd!Descuento)
                  sImporte(3) = sImporte(3) + Rsd!Impuesto
           End Select
           Rsd.MoveNext
        Loop
        For i = 0 To 3
            If sImporte(i) <> 0 Or nTasa(i) = 3 Then
               iva_id = nTasa(i)
               base_imp = sBaseCam(i) - IIf(pCosto = 0, sImporte(i), 0)
               importe = sImporte(i)
               ok = PyFEPDF.AgregarIva(iva_id, base_imp, importe)
            End If
        Next
     End If

    
    ' Agregar cada impuesto (por ej. IIBB, retenciones, percepciones, etc.):
    Dim alic As Currency
    
    If LblPercepcion.Caption <> 0 Then
       base_imp = LblNeto.Caption     ' importe sujeto a este tributo
       alic = cRsl.TraerValorDeUnCampo("Clientes", "Percepcion", "Cliente='" & TxtCliente.text & "'") ' alicuota (porcentaje) de este tributo
       importe = LblPercepcion.Caption        ' importe liquidado de este tributo
       ok = PyFEPDF.AgregarTributo(2, "Perc. Ing.Brutos", base_imp, Format(alic, "0.00"), importe)
    End If
    
'    Id = 4
'  '  desc = "Impuestos internos"
'    base_imp = LblNeto.Caption
'    alic = "1.00"
'    importe = "3168.90"
'    ok = PyFEPDF.AgregarTributo(4, "Percepcion Iva RG 2459", base_imp, alic, importe)
    ' Agrego detalles de cada item de la factura:
    
    Rsd.MoveFirst
    i = 0
    
    Dim nConta As Byte
    nConta = 1
    
    Dim bDetalle As Boolean
    
    Dim ret As Long
    ret = SafeUBound(sTexFiscal)  ' -1 si no está inicializado

   
    Do While Not Rsd.EOF
     '  If ret = 0 Then
          u_mtx = 123456              ' unidades
          cod_mtx = 1234567890123#    ' código de barras
          codigo = Rsd!Producto       ' codigo interno a imprimir (ej. "articulo")
          ds = Left(Rsd!Descripcion, 50)
          qty = Rsd!Cantidad           ' cantidad
          umed = Rsd!Medida ' código de unidad de medida (ej. 7 para "unidades")
          bonif = IIf(IsNull(Rsd!PDesc), 0, Rsd!PDesc) + LblDescuento.Caption       ' importe de descuentos
'       Else
'         If UBound(sTexFiscal) = 0 Then
'            ret = 0
'             u_mtx = 123456              ' unidades
'             cod_mtx = 1234567890123#    ' código de barras
'             codigo = Rsd!Producto       ' codigo interno a imprimir (ej. "articulo")
'             ds = Left(Rsd!Descripcion, 60)
'             qty = Rsd!Cantidad           ' cantidad
'             umed = 7                    ' código de unidad de medida (ej. 7 para "unidades")
'             bonif = Rsd!PDesc + LblDescuento.Caption   ' importe de descuentos
'          End If
'       End If
       precio = Rsd!PrecioUnitNeto ' precio neto (A) o iva incluido (B)
       If pCosto = 1 Then
          iva_id = Val(cRsl.TraerValorDeUnCampo("Impuestos", "CodAfip", "Impuesto=" & Rsd!ImpId))
          imp_iva = Rsd!Impuesto              ' importe liquidado de iva
          importe = Rsd!Total - Rsd!desc + IIf(pCompro = 11, 0, Rsd!Impuesto)
        Else
           iva_id = 0
           imp_iva = 0
           importe = Rsd!Total - Rsd!desc
        End If
       i = i + 1
       ' importe total del item
       despacho = ""               ' numero de despacho de importación
       dato_a = ""                 ' primer dato adicional del item
       dato_b = ""
       dato_c = ""
       dato_d = ""
       dato_e = ""           ' ultimo dato adicional del item
       
       ok = PyFEPDF.AgregarDetalleItem(u_mtx, cod_mtx, codigo, ds, qty, umed, _
            precio, bonif, iva_id, imp_iva, importe, despacho, _
            dato_a, dato_b, dato_c, dato_d, dato_e)
       Rsd.MoveNext
       nConta = nConta + 1
    Loop
    

    Dim rDatos As ADODB.Recordset
    
    Set rDatos = New ADODB.Recordset
    
    Set rDatos = cRsl.TraerRsCondi("AfipDatos", "Id_Datos", "Id_Datos=" & Val(txtSucursal.text))
    If rDatos.RecordCount = 0 Then
       MsgBox "Debe Definir las Sucursales en AFIP Datos", vbInformation, "Atención"
       Exit Function
    End If
    
    ' Agrego datos adicionales fijos:
     ok = PyFEPDF.AgregarDato("logo", Trim(Path & rDatos!Imagen))
     ok = PyFEPDF.AgregarDato("EMPRESA", Trim(rDatos!Fantasia))
     ok = PyFEPDF.AgregarDato("MEMBRETE1", Trim(rDatos!Empresa))
     ok = PyFEPDF.AgregarDato("MEMBRETE2", Trim(rDatos!Direccion))
     ok = PyFEPDF.AgregarDato("MEMBRETE3", Trim(rDatos!Localidad))
     ok = PyFEPDF.AgregarDato("CUIT", Trim(rDatos!Cuit))
     ok = PyFEPDF.AgregarDato("IIBB", Trim(rDatos!IIBB))
     ok = PyFEPDF.AgregarDato("IVA", Trim(rDatos!Iva))
     ok = PyFEPDF.AgregarDato("INICIO", "Inicio de Actividad: " & Trim(rDatos!Inicio))
     ok = PyFEPDF.AgregarDato("VENCI", Format(DtpVenc.Value, "dd/MM/yyyy"))

'     ok = PyFEPDF.AgregarDato("ObservacionesGenerales1", "Nota al pie1")
'     ok = PyFEPDF.AgregarDato("ObservacionesGenerales2", "")
'     ok = PyFEPDF.AgregarDato("ObservacionesGenerales3", "")

    ' Cargo el formato desde el archivo CSV (opcional)
    ' (carga todos los campos a utilizar desde la planilla)
    
    ok = PyFEPDF.CargarFormato(Path & rDatos!Archivo)
    
    ' Agrego campos manualmente (opcional):
    
    Dim Nombre As String, tipo As String, x1 As Single, x2 As Single, y1 As Single, y2 As Single
    Dim font As String, Size As Byte, bold As Boolean, italic As Boolean, underline As Boolean
    Dim foreground As Long, background As Long, Align As String, prioridad As Byte, text As String
    Dim priority As Byte
    
    
    Nombre = "Localidad": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
    x1 = 26.09: y1 = 57.5: x2 = 79.02: y2 = 59.7   ' coordenadas (en milimetros)
    font = "Arial": Size = 10: bold = 0: italic = 0: underline = 0 ' tipo de letra
    foreground = 0: background = 65535     ' colores de frente y fondo
    Align = "I" ' Alineación: Centrado, Izquierda, Derecha
    prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
    text = pLocalidad
    ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                        font, Size, bold, italic, underline, _
                        foreground, background, _
                        Align, text, priority)

    Nombre = "Provincia": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
    x1 = 99.29: y1 = 57.5: x2 = 140.3: y2 = 60.1     ' coordenadas (en milimetros)
    font = "Arial": Size = 10: bold = 0: italic = 0: underline = 0 ' tipo de letra
    foreground = 0: background = 65535     ' colores de frente y fondo
    Align = "I" ' Alineación: Centrado, Izquierda, Derecha
    prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
    text = pProvincia
    ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                        font, Size, bold, italic, underline, _
                        foreground, background, _
                        Align, text, priority)


    Nombre = "Texto": tipo = "T"           ' "T" texto, "L" lineas, "I" imagen, etc.
    x1 = 45.1: y1 = 79.01 + (4 * nConta): x2 = 105.4: y2 = 81.29 + (4 * nConta)      ' coordenadas (en milimetros)
    font = "Arial": Size = 8: bold = 0: italic = 0: underline = 0 ' tipo de letra
    foreground = 0: background = 65535   ' colores de frente y fondo
    Align = "I" ' Alineación: Centrado, Izquierda, Derecha
    prioridad = 2 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
   ' Text = sText
    
    Dim j As Byte
    
    Dim Ub As Long
      Ub = SafeUBound(sTexFiscal)
      If Ub >= 0 Then
          For j = 0 To Ub
             ok = PyFEPDF.AgregarCampo(Nombre & j, tipo, x1, y1, x2, y2, _
                   font, Size, bold, italic, underline, _
                   foreground, background, _
                   Align, sTexFiscal(j), priority)
                   y1 = y1 + 4
                   y2 = y2 + 4
          Next j
      End If
    
    If cRsl.TraerValorDeUnCampo("Impresion", "ResInscrip", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)) <> "0" Then
       Nombre = "Saldo": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
       x1 = 8.8: y1 = 239: x2 = 34.8: y2 = 243     ' coordenadas (en milimetros)
       font = "Arial": Size = 10: bold = 0: italic = 0: underline = 0 ' tipo de letra
       foreground = 0: background = 65535     ' colores de frente y fondo
       Align = "I" ' Alineación: Centrado, Izquierda, Derecha
       prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
    
       text = "Saldo Cta.Cte :" & Format(cSal.TraerSaldoCtaCte(TxtCliente.text), "#0.00")
       ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                        font, Size, bold, italic, underline, _
                        foreground, background, _
                        Align, text, priority)
    End If
    
    If pCompro = 6 Or pCompro = 11 Then
       Nombre = "Orientacion": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
       x1 = 13.3: y1 = 242.7: x2 = 108.3: y2 = 247.7      ' coordenadas (en milimetros)
       font = "Arial": Size = 7: bold = 0: italic = 0: underline = 0 ' tipo de letra
       foreground = 0: background = 65535     ' colores de frente y fondo
       Align = "I" ' Alineación: Centrado, Izquierda, Derecha
       prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
       text = "ORIENTACIÓN AL CONSUMIDOR PCIA BUENOS AIRES 0800-222-9042"
       
       ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                        font, Size, bold, italic, underline, _
                        foreground, background, _
                        Align, text, priority)
                        
       If pCompro = 6 Then
       
          Nombre = "regimen": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
          x1 = 13.3: y1 = 248.7: x2 = 109.3: y2 = 248.7      ' coordenadas (en milimetros)
          font = "Arial": Size = 7: bold = 0: italic = 0: underline = 0 ' tipo de letra
          foreground = 0: background = 65535     ' colores de frente y fondo
          Align = "I" ' Alineación: Centrado, Izquierda, Derecha
          prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
          text = "Régimen de Transparencia Fiscal al Consumidor. (Ley 27.743)"
          
          ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                           font, Size, bold, italic, underline, _
                           foreground, background, _
                           Align, text, priority)
                           
          If imp_iva_fact = 0 Then
              imp_iva_fact = Abs(((lblTotal.Caption / 1.21) - lblTotal.Caption))
          End If
                           
          Nombre = "regimen": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
          x1 = 13.3: y1 = 254.7: x2 = 109.3: y2 = 249.7      ' coordenadas (en milimetros)
          font = "Arial": Size = 7: bold = 0: italic = 0: underline = 0 ' tipo de letra
          foreground = 0: background = 65535     ' colores de frente y fondo
          Align = "I" ' Alineación: Centrado, Izquierda, Derecha
          prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
          text = "IVA Contenido :" & Format(imp_iva_fact, "#0.00")
          
          ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                           font, Size, bold, italic, underline, _
                           foreground, background, _
                           Align, text, priority)
                           
          Nombre = "regimen": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
          x1 = 13.3: y1 = 260.7: x2 = 109.3: y2 = 250.7      ' coordenadas (en milimetros)
          font = "Arial": Size = 7: bold = 0: italic = 0: underline = 0 ' tipo de letra
          foreground = 0: background = 65535     ' colores de frente y fondo
          Align = "I" ' Alineación: Centrado, Izquierda, Derecha
          prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
          text = "Otros Impuestos Nacionales Indirectos :" & Format(imp_tributo_fact, "#0.00")
          
          ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                           font, Size, bold, italic, underline, _
                           foreground, background, _
                           Align, text, priority)
                        
                        
       End If

    End If
    
    
    
    If (pCompro = 1 Or pCompro = 3) And cRsl.TraerValorDeUnCampo("Clientes", "TipoIva", "Cliente='" & TxtCliente.text & "'") = 4 Then
       Nombre = "Orientacion": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
       x1 = 13.3: y1 = 242.7: x2 = 108.3: y2 = 247.7      ' coordenadas (en milimetros)
       font = "Arial": Size = 7: bold = 0: italic = 0: underline = 0 ' tipo de letra
       foreground = 0: background = 65535     ' colores de frente y fondo
       Align = "I" ' Alineación: Centrado, Izquierda, Derecha
       prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
       text = "EL CREDITO FISCAL DISCRIMINADO EN EL PRESENTE COMPROBANTE SOLO PODRA SER COMPUTADO A EFECTOS"
            
       ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                        font, Size, bold, italic, underline, _
                        foreground, background, _
                        Align, text, priority)
                        
       Nombre = "Orientacion": tipo = "T" ' "T" texto, "L" lineas, "I" imagen, etc.
       x1 = 13.3: y1 = 245: x2 = 108.3: y2 = 250      ' coordenadas (en milimetros)
       font = "Arial": Size = 7: bold = 0: italic = 0: underline = 0 ' tipo de letra
       foreground = 0: background = 65535     ' colores de frente y fondo
       Align = "I" ' Alineación: Centrado, Izquierda, Derecha
       prioridad = 0 ' Orden Z, menor prioridad se dibuja primero (para superposiciones)
       text = "DEL REGIMEN DE SOSTENIMIENTO E INCLUSION FISCAL PARA PEQUENOS CONTRIBUYENTES DE LA LEY 27818."
                        
       ok = PyFEPDF.AgregarCampo(Nombre, tipo, x1, y1, x2, y2, _
                        font, Size, bold, italic, underline, _
                        foreground, background, _
                        Align, text, priority)

    End If
    
    
    
    ' Creo plantilla para esta factura (papel A4 vertical):
    papel = "A4" ' o "letter" para carta, "legal" para oficio
    orientacion = "portrait" ' o landscape (apaisado)
    ok = PyFEPDF.CrearPlantilla(papel, orientacion)
    num_copias = IIf(bEMail = False, rDatos!copias, 1) ' original, duplicado y triplicado
    lineas_max = IIf(nLineasElectronic = 0, 30, nLineasElectronic) ' cantidad de linas de items por página
    qty_pos = "izq" ' (cantidad a la izquierda de la descripción del artículo)
    ' Proceso la plantilla
    ok = PyFEPDF.ProcesarPlantilla(num_copias, lineas_max, qty_pos)
    ' Genero el PDF de salida según la plantilla procesada
    Dim sCob As String
    
    sCob = cRsl.TraerValorDeUnCampo("Comprobantes", "Comprobante", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
    
    salida = Path & "\Facturas\" & sCob & "-" & LblCliente.Caption & "-" & TxtNumero.text & ".pdf"
    ok = PyFEPDF.GenerarPDF(salida)
    
    ' Abro el visor de PDF y muestro lo generado
    ' (es necesario tener instalado Acrobat Reader o similar)
    Imprimir = False ' cambiar a True para que lo envie directo a la impresora
    If bEMail = False Then
       ok = PyFEPDF.MostrarPDF(salida, Imprimir)
    End If
    
    bEMail = False
 '   Erase sTexFiscal
    Exit Function
ManejoError:
    ' Si hubo error:
    Debug.Print err.Description            ' descripción error afip
    Debug.Print err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(err.Description, vbCritical + vbRetryCancel, "Error:" & err.Number - vbObjectError & " en " & err.Source)
        Case vbRetry
            Debug.Print PyFEPDF.Excepcion
            Debug.Print PyFEPDF.Traceback
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print err.Description
    End Select
    Debug.Assert False
End Function


Private Function SafeUBound(ByRef arr As Variant) As Long
    On Error GoTo EH
    SafeUBound = UBound(arr)
    Exit Function
EH:
    SafeUBound = -1
End Function

Private Function NzNum(ByVal v As Variant, Optional ByVal defValue As Double = 0) As Double
    On Error GoTo EH
    If IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then
        NzNum = defValue
    Else
        NzNum = CDbl(v)
    End If
    Exit Function
EH:
    NzNum = defValue
End Function

