VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MtrProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Productos"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   7665
   Begin VB.Frame FraProducto 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   300
      TabIndex        =   21
      Top             =   510
      Width           =   5895
      Begin VB.Frame FraImpuesto 
         Caption         =   "Ubicación e Impuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   5895
         Begin VB.Frame FraOtros 
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4650
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   5895
            Begin VB.Frame Frame4 
               Caption         =   "C.Producto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2025
               Left            =   3555
               TabIndex        =   77
               Top             =   210
               Width           =   2160
               Begin VB.ComboBox CmbCalcu 
                  Height          =   315
                  Left            =   1245
                  Style           =   2  'Dropdown List
                  TabIndex        =   80
                  Top             =   1470
                  Width           =   600
               End
               Begin VB.ComboBox CmbFijo 
                  Height          =   315
                  Left            =   1020
                  Style           =   2  'Dropdown List
                  TabIndex        =   79
                  Top             =   810
                  Width           =   825
               End
               Begin VB.ComboBox CmbCompuesto 
                  Height          =   315
                  Left            =   1245
                  Style           =   2  'Dropdown List
                  TabIndex        =   78
                  Top             =   315
                  Width           =   600
               End
               Begin VB.Label Label31 
                  Caption         =   "C.por Importe :"
                  Height          =   225
                  Left            =   135
                  TabIndex        =   83
                  Top             =   1470
                  Width           =   1140
               End
               Begin VB.Label Label30 
                  Caption         =   "Fijo :"
                  Height          =   210
                  Left            =   150
                  TabIndex        =   82
                  Top             =   795
                  Width           =   885
               End
               Begin VB.Label Label29 
                  Caption         =   "Compuesto :"
                  Height          =   195
                  Left            =   150
                  TabIndex        =   81
                  Top             =   315
                  Width           =   930
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Costo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2025
               Left            =   120
               TabIndex        =   70
               Top             =   210
               Width           =   3330
               Begin VB.ComboBox CmbMoneda 
                  Height          =   315
                  Left            =   1155
                  Style           =   2  'Dropdown List
                  TabIndex        =   72
                  Top             =   1500
                  Width           =   1905
               End
               Begin VB.TextBox TxtCompra 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   10
                  TabIndex        =   71
                  Text            =   "TxtCompra"
                  Top             =   840
                  Width           =   1140
               End
               Begin MSComCtl2.DTPicker DTFecha 
                  Height          =   315
                  Left            =   1755
                  TabIndex        =   73
                  Top             =   330
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   97779713
                  CurrentDate     =   36775
               End
               Begin VB.Label Label28 
                  Caption         =   "Moneda :"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   76
                  Top             =   1500
                  Width           =   930
               End
               Begin VB.Label Label27 
                  Caption         =   "Fecha Ult.Compra :"
                  Height          =   255
                  Left            =   135
                  TabIndex        =   75
                  Top             =   345
                  Width           =   1515
               End
               Begin VB.Label Label26 
                  Caption         =   "Precio Ult.Compra :"
                  Height          =   255
                  Left            =   135
                  TabIndex        =   74
                  Top             =   795
                  Width           =   1455
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Cantidades"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2310
               Left            =   120
               TabIndex        =   63
               Top             =   2250
               Width           =   5625
               Begin VB.ComboBox CmbCFamilia 
                  Height          =   315
                  ItemData        =   "MtrProductos.frx":0000
                  Left            =   4725
                  List            =   "MtrProductos.frx":0025
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   90
                  Top             =   1845
                  Width           =   750
               End
               Begin VB.ComboBox CmbCDepartamento 
                  Height          =   315
                  ItemData        =   "MtrProductos.frx":0055
                  Left            =   3345
                  List            =   "MtrProductos.frx":007A
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   88
                  Top             =   1830
                  Width           =   750
               End
               Begin VB.ComboBox CmbBalanza 
                  Height          =   315
                  ItemData        =   "MtrProductos.frx":00AA
                  Left            =   1935
                  List            =   "MtrProductos.frx":00C6
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   87
                  Top             =   1845
                  Width           =   750
               End
               Begin VB.TextBox TxtEmbase 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1905
                  MaxLength       =   10
                  TabIndex        =   84
                  Text            =   "TxtMaxima"
                  Top             =   1380
                  Width           =   1140
               End
               Begin VB.TextBox TxtMinima 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   10
                  TabIndex        =   65
                  Text            =   "TxtMinima"
                  Top             =   420
                  Width           =   1140
               End
               Begin VB.TextBox TxtMaxima 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1920
                  MaxLength       =   10
                  TabIndex        =   64
                  Text            =   "TxtMaxima"
                  Top             =   900
                  Width           =   1140
               End
               Begin VB.Label Label35 
                  Caption         =   "C.Fam :"
                  Height          =   225
                  Left            =   4125
                  TabIndex        =   91
                  Top             =   1845
                  Width           =   825
               End
               Begin VB.Label Label34 
                  Caption         =   "C.Dep :"
                  Height          =   225
                  Left            =   2760
                  TabIndex        =   89
                  Top             =   1845
                  Width           =   825
               End
               Begin VB.Label Label33 
                  Caption         =   "Balanza :"
                  Height          =   225
                  Left            =   390
                  TabIndex        =   86
                  Top             =   1845
                  Width           =   825
               End
               Begin VB.Label Label32 
                  Caption         =   "Valor Envase :"
                  Height          =   255
                  Left            =   390
                  TabIndex        =   85
                  Top             =   1365
                  Width           =   1275
               End
               Begin VB.Label Label25 
                  Caption         =   "Cant.Minima :"
                  Height          =   225
                  Left            =   390
                  TabIndex        =   69
                  Top             =   435
                  Width           =   1215
               End
               Begin VB.Label Label24 
                  Caption         =   "Cant.Maxima :"
                  Height          =   255
                  Left            =   390
                  TabIndex        =   68
                  Top             =   900
                  Width           =   1275
               End
               Begin VB.Label Label23 
                  Caption         =   "Cant.Actual :"
                  Height          =   240
                  Left            =   3210
                  TabIndex        =   67
                  Top             =   900
                  Width           =   1065
               End
               Begin VB.Label LblActual 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4305
                  TabIndex        =   66
                  Top             =   900
                  Width           =   1005
               End
            End
            Begin VB.TextBox TxtId 
               Height          =   300
               Left            =   4845
               TabIndex        =   62
               Text            =   "TxtId"
               Top             =   2595
               Visible         =   0   'False
               Width           =   765
            End
         End
         Begin VB.ComboBox CmbDeposito 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   2700
            Width           =   2520
         End
         Begin VB.ListBox LstImpuestos 
            Height          =   2085
            Left            =   3225
            Style           =   1  'Checkbox
            TabIndex        =   48
            Top             =   330
            Width           =   2475
         End
         Begin VB.ComboBox CmbCompra 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   555
            Width           =   2520
         End
         Begin VB.ComboBox CmbVenta 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1230
            Width           =   2520
         End
         Begin VB.ComboBox CmbOrigen 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   3360
            Width           =   2520
         End
         Begin VB.TextBox TxtCantDes 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   225
            MaxLength       =   7
            TabIndex        =   44
            Text            =   "TxtCant"
            Top             =   1965
            Width           =   1035
         End
         Begin VB.TextBox TxtCantVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1665
            MaxLength       =   7
            TabIndex        =   43
            Text            =   "TxtCant"
            Top             =   1965
            Width           =   1065
         End
         Begin VB.Frame Frame1 
            Caption         =   "Imp. Internos"
            Height          =   1215
            Left            =   3210
            TabIndex        =   39
            Top             =   2505
            Width           =   2550
            Begin VB.ComboBox CmbInterno 
               Height          =   315
               Left            =   150
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   300
               Width           =   2145
            End
            Begin VB.TextBox TxtvInterno 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1425
               MaxLength       =   7
               TabIndex        =   40
               Text            =   "TxtvInt"
               Top             =   735
               Width           =   840
            End
            Begin VB.Label Label17 
               Caption         =   "Valor :"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   735
               Width           =   495
            End
         End
         Begin MSComCtl2.DTPicker DTPBaja 
            Height          =   330
            Left            =   4110
            TabIndex        =   50
            Top             =   3945
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            Format          =   119209985
            CurrentDate     =   36768
         End
         Begin MSComCtl2.DTPicker DTPAlta 
            Height          =   330
            Left            =   1455
            TabIndex        =   51
            Top             =   3945
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            Format          =   119209985
            CurrentDate     =   36768
         End
         Begin VB.Label Label22 
            Caption         =   "Impuestos :"
            Height          =   210
            Left            =   3225
            TabIndex        =   60
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label21 
            Caption         =   "U.M.Principal :"
            Height          =   225
            Left            =   225
            TabIndex        =   59
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "U.M. Stock :"
            Height          =   240
            Left            =   225
            TabIndex        =   58
            Top             =   960
            Width           =   1470
         End
         Begin VB.Label Label19 
            Caption         =   "Depósito :"
            Height          =   225
            Left            =   210
            TabIndex        =   57
            Top             =   2445
            Width           =   1155
         End
         Begin VB.Label Label15 
            Caption         =   "Precios Claro :"
            Height          =   240
            Left            =   210
            TabIndex        =   56
            Top             =   3105
            Width           =   1185
         End
         Begin VB.Label Label16 
            Caption         =   "Cant. Princ.. :"
            Height          =   225
            Left            =   225
            TabIndex        =   55
            Top             =   1725
            Width           =   1050
         End
         Begin VB.Label Label18 
            Caption         =   "Cant. Stock :"
            Height          =   195
            Left            =   1665
            TabIndex        =   54
            Top             =   1725
            Width           =   1020
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha Alta :"
            Height          =   240
            Left            =   300
            TabIndex        =   53
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Baja :"
            Height          =   240
            Left            =   3090
            TabIndex        =   52
            Top             =   3960
            Width           =   930
         End
      End
      Begin VB.TextBox TxtCBarra 
         Height          =   315
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   11
         Text            =   "TxtCodigo"
         Top             =   4215
         Width           =   3945
      End
      Begin VB.CommandButton CmdRep 
         Caption         =   "..."
         Height          =   270
         Left            =   4905
         TabIndex        =   8
         Top             =   2985
         Width           =   435
      End
      Begin VB.TextBox TxtRep 
         Height          =   285
         Left            =   3320
         MaxLength       =   15
         TabIndex        =   7
         Text            =   "000000000000000"
         Top             =   2985
         Width           =   1485
      End
      Begin VB.TextBox TxtPeso 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   6
         Text            =   "TxtPe"
         Top             =   2985
         Width           =   990
      End
      Begin VB.ComboBox CmbFamiliaSub 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2550
         Width           =   3500
      End
      Begin VB.TextBox TxtDescripcion1 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "TxtDescripcion1"
         Top             =   1302
         Width           =   4260
      End
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1719
         Width           =   3500
      End
      Begin VB.ComboBox CmbStock 
         Height          =   315
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   450
         Width           =   630
      End
      Begin VB.ComboBox CmbMarca 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3390
         Width           =   3500
      End
      Begin VB.ComboBox CmbFamilia 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2136
         Width           =   3500
      End
      Begin VB.CheckBox ChkActivo 
         Caption         =   "Activo"
         Height          =   225
         Left            =   4695
         TabIndex        =   27
         Top             =   135
         Width           =   1080
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "TxtDescripcion"
         Top             =   885
         Width           =   4275
      End
      Begin VB.ComboBox CmbProveedor 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3810
         Width           =   3960
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   315
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "TxtCodigo"
         Top             =   450
         Width           =   1560
      End
      Begin VB.Label Label5 
         Caption         =   "C.Barra :"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   4215
         Width           =   885
      End
      Begin VB.Label Label13 
         Caption         =   "Repres.:"
         Height          =   225
         Left            =   2625
         TabIndex        =   36
         Top             =   2970
         Width           =   675
      End
      Begin VB.Label Label12 
         Caption         =   "Peso :"
         Height          =   255
         Left            =   165
         TabIndex        =   35
         Top             =   2970
         Width           =   900
      End
      Begin VB.Label Label11 
         Caption         =   "Sub Familia :"
         Height          =   240
         Left            =   165
         TabIndex        =   34
         Top             =   2550
         Width           =   1050
      End
      Begin VB.Label Label10 
         Caption         =   "Desc. Adicio. :"
         Height          =   255
         Left            =   165
         TabIndex        =   33
         Top             =   1290
         Width           =   1110
      End
      Begin VB.Label Label9 
         Caption         =   "A.Stock :"
         Height          =   210
         Left            =   4245
         TabIndex        =   31
         Top             =   450
         Width           =   690
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo :"
         Height          =   255
         Left            =   165
         TabIndex        =   30
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label Label8 
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   150
         TabIndex        =   28
         Top             =   3810
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   165
         TabIndex        =   25
         Top             =   450
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         Height          =   255
         Left            =   165
         TabIndex        =   24
         Top             =   900
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Marca :"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   3390
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Familia :"
         Height          =   300
         Left            =   165
         TabIndex        =   22
         Top             =   2085
         Width           =   705
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   780
      Index           =   0
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   780
      Index           =   1
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   780
      Index           =   2
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   780
      Index           =   3
      Left            =   3075
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Etiquetas"
      Height          =   780
      Index           =   4
      Left            =   4065
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5415
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   5
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   420
      Width           =   915
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
      Height          =   780
      Index           =   6
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1275
      Width           =   915
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
      Height          =   780
      Index           =   7
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2115
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   8
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2970
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   780
      Index           =   9
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5415
      Width           =   915
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5175
      Left            =   150
      TabIndex        =   26
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9128
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificación"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ubicación e Impuestos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MtrProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsC As Recordset
Dim sCod As Integer, sw As Boolean, nId As Integer
Dim sCod1 As String
Dim Rs As ADODB.Recordset, Rs_Pro As ADODB.Recordset
Dim cRse As ClsProductoE
Dim Cadena As String, ic As Integer, nRepAux As Long
Dim cRsl As ClsLectura

Private Sub CmbBalanza_Click()
  TodoBien
End Sub
Private Sub CmbCDepartamento_Click()
  TodoBien
End Sub
Private Sub CmbCFamilia_Click()

  If CmbFamilia.ListIndex <> -1 Then
     cRsl.CargaCombo CmbFamiliaSub, "FamiliasSub", "Id", "Descripcion", "Familia=" & CmbFamilia.ItemData(CmbFamilia.ListIndex)
  End If

  TodoBien
End Sub
Private Sub CmbCompra_Click()
  TodoBien
End Sub
Private Sub CmbDeposito_Click()
  TodoBien
End Sub

Private Sub CmbFamilia_Click()
   TodoBien
End Sub

Private Sub CmbInterno_Click()
  TodoBien
End Sub

Private Sub CmbMarca_Click()
  TodoBien
End Sub

Private Sub CmbMoneda_Change()
  TodoBien
End Sub

Private Sub CmbProveedor_Click()
  TodoBien
End Sub

Private Sub CmbVenta_Click()
  TodoBien
End Sub
Private Sub ChkActivo_Click()
  If sw = False Then
     DTPBaja.Enabled = False
     If ChkActivo.Value = 0 Then
        DTPBaja.Enabled = True
     End If
  End If
End Sub

Private Sub CmdRep_Click()
  Dim nRep As Long
    
  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     nRep = cRsl.TraerValorDeUnCampo("Productos", "Representante", "Id=" & nDat)
     If nRep <> 0 Then
        TxtRep.Text = nRep
        nRepAux = 0
     Else
        TxtRep.Text = nDat
        nRepAux = nDat
     End If
  End If
  
End Sub

Private Sub Form_Load()
  On Error GoTo errHandler

  Set Rs = New ADODB.Recordset
  Set Rs_Pro = New ADODB.Recordset
  
  Set cRsl = New ClsLectura

  VerFrame True, False, False
  sw = True
  Habilitar False
  Limpiar
  CargarCombos
  Botones pEdit2, False, False, False, False, False, False, False, False, True
  CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
  CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
  CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
  CmdBotones(3).Picture = LoadResPicture("Buscar", 0)
  CmdBotones(4).Picture = LoadResPicture("Imprimir", 0)
  CmdBotones(9).Picture = LoadResPicture("Salir", 0)
  
  ArmaRsParaMovimiento
   
  Cadena = ""
  ic = 1
  
  
  Me.Top = 0
  Me.Left = 0
Exit Sub

errHandler:
   ManejaErrores
 
End Sub

Private Sub Form_Resize()
  If CmdBotones(0).Enabled = True Then
     CmdBotones(0).SetFocus
  End If
End Sub

Private Sub LstImpuestos_Click()
  TodoBien
End Sub

Private Sub TabStrip1_Click()
 Select Case TabStrip1.SelectedItem.Index
        Case 1
             VerFrame True, False, False
        Case 2
             VerFrame True, True, False
        Case 3
             VerFrame True, True, True
 End Select
End Sub

Private Sub VerFrame(bIden As Boolean, bDomi As Boolean, bImp As Boolean)
  FraProducto.Visible = bIden
  FraImpuesto.Visible = bDomi
  FraOtros.Visible = bImp
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
      KeyAscii = 0
   End If
   KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.Text = ""
      End If
  Next ctl
  For Each ctl In Controls
      If TypeOf ctl Is ComboBox Then
         ctl.ListIndex = -1
      End If
  Next ctl
  TxtPeso.Text = "0.000"
  TxtRep.Text = "0"
  TxtCantDes.Text = "1.00"
  TxtCantVent.Text = "1.00"
  TxtvInterno.Text = "0.0000"
  TxtEmbase.Text = "0.00"
  
  TxtCompra.Text = Format(0, nCantDecimales)
  TxtMinima.Text = Format(0, nCantDecimales)
  TxtMaxima.Text = Format(0, nCantDecimales)
  DTPAlta.Value = Date
  DTPBaja.Value = Date
  DTFecha.Value = Date
  LblActual.Caption = Format(0, nCantDecimales)
  CmbBalanza.Text = "00"
  CmbCDepartamento.Text = "01"
  CmbCFamilia.Text = "01"
  
End Sub

Private Sub Habilitar(bEstado As Boolean)
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
         Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
         Or TypeOf ctl Is ComboBox Then
         ctl.Enabled = bEstado
      End If
  Next ctl
  CmdRep.Enabled = bEstado
End Sub

Private Sub Botones(bNue As Boolean, bBor As Boolean, bMod As Boolean, bBus As Boolean, _
 bImp As Boolean, bPri As Boolean, bSig As Boolean, bAnt As Boolean, bUlt As Boolean, _
 bSal As Boolean)
   CmdBotones(0).Enabled = bNue
   CmdBotones(1).Enabled = bBor
   CmdBotones(2).Enabled = bMod
   CmdBotones(3).Enabled = bBus
   CmdBotones(4).Enabled = bImp
   CmdBotones(5).Enabled = bPri
   CmdBotones(6).Enabled = bSig
   CmdBotones(7).Enabled = bAnt
   CmdBotones(8).Enabled = bUlt
   CmdBotones(9).Enabled = bSal
 End Sub

Private Sub CargarCombos()
    
  cRsl.CargaCombo CmbTipo, "[TipoProducto]", "Id", "Descripcion", ""
  cRsl.CargaCombo CmbFamilia, "Familias", "Familia", "Descripcion", ""
  cRsl.CargaCombo CmbMarca, "Marcas", "Marca", "Descripcion", ""
  cRsl.CargaCombo CmbProveedor, "Proveedores", "Id", "RazonSocial", ""
  cRsl.CargaCombo CmbCompra, "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbVenta, "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbDeposito, "Depositos", "Deposito", "Descripcion", ""
'  cRsl.CargaCombo CmbOrigen, "Paises", "Pais", "Descripcion", ""
  cRsl.CargaCombo CmbMoneda, "Monedas", "Id", "Descripcion", ""
 
  CmbOrigen.AddItem "Si"
  CmbOrigen.AddItem "No"

  cRsl.CargaCombo LstImpuestos, "Impuestos", "Impuesto", "Descripcion", "Producto='1'"
  CmbStock.AddItem "Si"
  CmbStock.AddItem "No"
  CmbStock.ListIndex = 0
  CmbCompuesto.AddItem "No"
  CmbCompuesto.AddItem "Si"
  CmbCompuesto.ListIndex = 0
  
  CmbFijo.AddItem "Um"
  CmbFijo.AddItem "Ux"
  CmbFijo.AddItem "Si"
  CmbFijo.AddItem "No"
  
  CmbFijo.ListIndex = 0
  
  CmbInterno.AddItem "No"
'  CmbInterno.AddItem "Fijo"
  CmbInterno.AddItem "Porcentaje"
    
  CmbInterno.Text = "No"
  
  CmbCalcu.AddItem "Si"
  CmbCalcu.AddItem "No"
 
  CmbCalcu.Text = "No"
  
  
End Sub

Private Sub MarcarImpuestosProducto(ByVal sProducto As String)
    Dim rSel As ADODB.Recordset
    Dim i As Long
    Dim impu As Long

    ' limpiar selección
    For i = 0 To LstImpuestos.ListCount - 1
        LstImpuestos.Selected(i) = False
    Next

    ' traer los impuestos del producto (1 consulta)
    Set rSel = cRsl.TraerRS("ImpuestoTraerUno", sProducto, True) ' SP nueva

    Do While Not rSel.EOF
        impu = CLng(rSel!Impuesto)

        ' como son 3, busqueda lineal y chau
        For i = 0 To LstImpuestos.ListCount - 1
            If CLng(LstImpuestos.ItemData(i)) = impu Then
                LstImpuestos.Selected(i) = True
                Exit For
            End If
        Next

        rSel.MoveNext
    Loop
End Sub

Private Sub LinkearTexto()
  Dim cRi As ClsProductoL, i As Integer, P As ClsPrograma
  
  Set cRi = New ClsProductoL
  
  
  Set P = New ClsPrograma
  
  Limpiar
  nRepAux = 0
  
  ChkActivo.Value = IIf(Rs!Activo = False, 0, 1)
  TxtId = Rs!ID
  TxtCodigo.Text = Rs!Producto
  TxtDescripcion.Text = "" & Rs!Descripcion
  TxtDescripcion1.Text = "" & Rs!Descripcion1
  CmbTipo.Text = cRsl.DatoCombo("TipoProducto", "Id", "Descripcion", Rs!tipo)
  P.SetComboByItemData CmbTipo, CLng(Rs!tipo)
  If Not IsNull(Rs!Familia) Then P.SetComboByItemData CmbFamilia, CLng(Rs!Familia)
  If Not IsNull(Rs!Marca) Then P.SetComboByItemData CmbMarca, CLng(Rs!Marca)
  If Not IsNull(Rs!Proveedor) Then P.SetComboByItemData CmbProveedor, CLng(Rs!Proveedor)
  If Not IsNull(Rs!UMCompra) Then P.SetComboByItemData CmbCompra, CLng(Rs!UMCompra)
  If Not IsNull(Rs!UMVenta) Then P.SetComboByItemData CmbVenta, CLng(Rs!UMVenta)
  If Not IsNull(Rs!Deposito) Then P.SetComboByItemData CmbDeposito, CLng(Rs!Deposito)
  If Not IsNull(Rs!Moneda) Then P.SetComboByItemData CmbMoneda, CLng(Rs!Moneda)

  TxtCantDes.Text = Format(Rs!CantDespacho, nCantDecimales)
  TxtCantVent.Text = Format(Rs!CantVenta, nCantDecimales)
  DTFecha.Value = Rs!FechaCompra
  TxtPeso.Text = Format(Rs!Peso, "0.000")
  TxtRep.Text = Rs!Representante
  TxtCompra.Text = Format(Rs!PrecioCompra, nCantDecimales)
  TxtMinima.Text = Format(Rs!CantMinima, nCantDecimales)
  TxtMaxima.Text = Format(Rs!CantMaxima, nCantDecimales)
  CmbStock.Text = "" & Rs!ActuaStock
  CmbCompuesto.Text = "" & Rs!CodCompu
  CmbFijo.Text = "" & Rs!fijo
  CmbInterno.Text = Rs!TipoInterno
  CmbCalcu.Text = Rs!Calcular
  TxtvInterno.Text = Format(Rs!ValorInterno, "#0.0000")
  TxtCBarra.Text = "" & Rs!CBarra
  TxtEmbase.Text = Rs!ValorEnvase
  TxtEmbase.Text = Format(Rs!ValorEnvase, nCantDecimales)
  CmbBalanza.Text = Format(Rs!Balanza, "00")
  CmbCDepartamento.Text = Format(Rs!CDepartamento, "00")
  CmbCFamilia.Text = Format(Rs!CFamilia, "00")
  
  Dim Rsi As ADODB.Recordset

  Set Rsi = New ADODB.Recordset

  Set Rsi = cRsl.TraerRS("ImpuestoTraerTodos", , True)
  If Rsi.EOF = False And Rsi.BOF = False Then
     LstImpuestos.Clear
     i = 0
     Do While Not Rsi.EOF
        If Rsi!Producto = "1" Then
           LstImpuestos.AddItem Rsi!Descripcion
           LstImpuestos.ItemData(LstImpuestos.NewIndex) = Rsi!Impuesto
           If cRi.TieneImpuesto(Trim(Rs!Producto), Rsi!Impuesto) = True Then
              LstImpuestos.Selected(i) = True
           End If
           i = i + 1
        End If
        Rsi.MoveNext
     Loop
  End If

 ' MarcarImpuestosProducto Trim$(rs!Producto)
  LblActual.Caption = Format(cRi.DameCantidad(Rs!Producto), nCantDecimales)
  
  Set cRi = Nothing
End Sub



Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Borrar
         Case 2
              Modificar
         Case 3
              Buscar
         Case 4
              Etiquetas
         Case 5
              Primero
         Case 6
              Anterior
         Case 7
              Siguiente
         Case 8
              ultimo
         Case 9
              Salir
  End Select
End Sub

Public Sub Etiquetas()
  Dim Rsd As Recordset
  Dim i As Integer
  
  
  Set RsC = New Recordset
  Set Rsd = New Recordset
  
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

  RsC.AddNew
  RsC!Descripcion = TxtDescripcion.Text
  RsC!DescripAdi = TxtDescripcion1.Text
  RsC!Leyenda = ""
  RsC!fecha = Date
  RsC!Porcentaje = cRsl.TraerValorDeUnCampo("Proveedores", "Porcentaje", "Id=" & CmbProveedor.ItemData(CmbProveedor.ListIndex))
  RsC!codigo = TxtCodigo.Text
  RsC!Barra = TxtCodigo.Text
  RsC.Update
  
  nImpr = 54
  FrmImpresor.Show
  Me.MousePointer = 0
  
End Sub

Private Sub Nuevo()
  On Error GoTo Errores
  
  If sw = True Then
     Cadena = ""
     ic = 1
     Habilitar True
     Botones False, False, False, False, False, False, False, False, False, True
     Limpiar
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(9).Caption = "Cancelar"
     CmdBotones(9).Picture = LoadResPicture("Cancelar", 0)
     TxtCodigo.SetFocus
     sw = False
     ChkActivo.Value = 1
     CmbTipo.ListIndex = 0
     CmbFijo.Text = "No"
     CmbCompuesto.Text = "No"
     CmbTipo.Text = "Producto Terminado"
     Dim RsReg As Recordset, P As ClsPrograma
     
     Set P = New ClsPrograma
     Set RsReg = cRsl.TraerRsCondi("Registros", "Usuario", "Usuario=" & nUsuario & " And Suc = " & nSucursal)
                
     P.SetComboByItemData CmbFamilia, RsReg!Familia ' CmbProvincia, RsReg!Provincia ' cRsl.DatoCombo("Familias", "Familia", "Descripcion", cRsl.RegPorDefecto("Familia"))
     P.SetComboByItemData CmbMarca, RsReg!Marca ' cRsl.DatoCombo("Marcas", "Marca", "Descripcion", cRsl.RegPorDefecto("Marca"))
     P.SetComboByItemData CmbProveedor, RsReg!Proveedor ' cRsl.DatoCombo("Proveedores", "Id", "RazonSocial", cRsl.RegPorDefecto("Proveedor"))
     P.SetComboByItemData CmbCompra, RsReg!Medida ' cRsl.DatoCombo("UnidadMedida", "Unidad", "Descripcion", cRsl.RegPorDefecto("Medida"))
     P.SetComboByItemData CmbVenta, RsReg!Medida ' cRsl.DatoCombo("UnidadMedida", "Unidad", "Descripcion", cRsl.RegPorDefecto("Medida"))
     P.SetComboByItemData CmbDeposito, RsReg!Deposito ' cRsl.DatoCombo("Depositos", "Deposito", "Descripcion", cRsl.RegPorDefecto("Deposito"))
     P.SetComboByItemData CmbMoneda, RsReg!Moneda ' cRsl.DatoCombo("Monedas", "Id", "Descripcion", cRsl.RegPorDefecto("Moneda"))
     CmbOrigen.Text = "No"  ' Precios Claros
     CmbStock.Text = IIf(cRsl.TraerValorDeUnCampo("Sistema", "MueveStock", "") = True, "Si", "No")
     CmbInterno.Text = "No"
     CmbCalcu.Text = "No"
     DTPBaja.Enabled = False
  Else
     Botones True, True, True, True, True, True, True, True, True, True
     Habilitar False
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(9).Caption = "Salir"
     CmdBotones(9).Picture = LoadResPicture("Salir", 0)
     sw = True
     Grabar
     ArmaRsParaMovimiento
  End If
  
  Exit Sub
Errores:
   ManejaErrores
End Sub

Private Sub Salir()
 On Error GoTo errHandler

  Botones True, False, False, False, False, False, False, False, False, True
  If sw = False Then
     sw = True
     Limpiar
     If CmdBotones(0).Caption = "Grabar" Then
        Habilitar False
        CmdBotones(0).Caption = "Nuevo"
        CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
        CmdBotones(9).Caption = "Salir"
        CmdBotones(9).Picture = LoadResPicture("Salir", 0)
     Else
        Habilitar False
        CmdBotones(2).Caption = "Modificar"
        CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
        CmdBotones(9).Caption = "Salir"
        CmdBotones(9).Picture = LoadResPicture("Salir", 0)
     End If
      If Rs_Pro.BOF = False And Rs_Pro.EOF = False Then
        Botones True, True, True, True, True, True, True, True, True, True
        LinkearTexto
     End If
  Else
     Unload Me
  End If
  Exit Sub
errHandler:
   ManejaErrores
End Sub
Private Sub Modificar()
  
  On Error GoTo Errores
  
  Static rList As Integer
  
  If sw = True Then
     Cadena = ""
     ic = 1
     sw = False
     sCod1 = Rs!Producto
     Botones False, False, True, False, False, False, False, False, False, True
     Habilitar True
     CmdBotones(2).Caption = "Grabar"
     CmdBotones(2).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(9).Caption = "Cancelar"
     CmdBotones(9).Picture = LoadResPicture("Cancelar", 0)
     DTPBaja.Enabled = False
  Else
     Botones True, True, True, True, True, True, True, True, True, True
     Habilitar False
     Grabar
     CmdBotones(2).Caption = "Modificar"
     CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
     CmdBotones(9).Caption = "Salir"
     CmdBotones(9).Picture = LoadResPicture("Salir", 0)
     sw = True
  End If
Exit Sub
Errores:
   ManejaErrores

End Sub
Private Sub Borrar()
  Dim cCant As ClsProductoL
  
  On Error GoTo errHandler
  
  Set cRse = New ClsProductoE
  Set cCant = New ClsProductoL
  
  
  If cRsl.TraerCantidad("DetallesComprobantes", "Producto", "Producto='" & TxtCodigo.Text & "'") = 0 Then
     If cCant.DameCantidad(Rs!Producto, Rs!Deposito) = 0 Then
        If MsgBox("Borra el Producto Seleccionado ?", 20, "Atención") = 6 Then
           cRse.Borrar "ProductoBorrar", Rs!Producto, TxtId.Text, Rs!Deposito
           Limpiar
           Set cRse = Nothing
           ArmaRsParaMovimiento
        End If
    Else
        MsgBox "El Producto no se Puede Borrar, Verifique las Cantidades", vbCritical, "Atención"
    End If
 Else
    MsgBox "Producto Asociado a una Factura de Compra o Venta", vbCritical, "Atención"
 End If
Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Grabar()
  On Error GoTo errHandler
  
  Dim Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, Rs3 As ADODB.Recordset
  Dim i As Integer
  
  Set Rs1 = New ADODB.Recordset
  Set Rs2 = New ADODB.Recordset
  Set Rs3 = New ADODB.Recordset
  
  
  Set cRse = New ClsProductoE
    
  Set Rs1 = cRsl.RsVacio("Productos", "Producto", "S")
  
  Rs1.AddNew
  Rs1!Activo = ChkActivo.Value
  Rs1!Producto = Replace(TxtCodigo.Text, "'", "")
  Rs1!Descripcion = Replace(TxtDescripcion.Text, "'", "")
  Rs1!Descripcion1 = Replace(TxtDescripcion1.Text, "'", "")
  Rs1!tipo = CmbTipo.ItemData(CmbTipo.ListIndex)
  If CmbFamilia.ListIndex <> -1 Then
     Rs1!Familia = CmbFamilia.ItemData(CmbFamilia.ListIndex)
  End If
  If CmbFamiliaSub.ListIndex <> -1 Then
     Rs1!FamiliaSub = CmbFamiliaSub.ItemData(CmbFamiliaSub.ListIndex)
  End If
  If CmbMarca.ListIndex <> -1 Then
     Rs1!Marca = CmbMarca.ItemData(CmbMarca.ListIndex)
  End If
  Rs1!FechaAlta = DTPAlta.Value
  Rs1!FechaBaja = DTPBaja.Value
  If CmbProveedor.ListIndex <> -1 Then
     Rs1!Proveedor = CmbProveedor.ItemData(CmbProveedor.ListIndex)
  End If
  If CmbCompra.ListIndex <> -1 Then
     Rs1!UMCompra = CmbCompra.ItemData(CmbCompra.ListIndex)
  End If
  If CmbVenta.ListIndex <> -1 Then
     Rs1!UMVenta = CmbVenta.ItemData(CmbVenta.ListIndex)
  End If
  If CmbDeposito.ListIndex <> -1 Then
     Rs1!Deposito = CmbDeposito.ItemData(CmbDeposito.ListIndex)
  Else
     Rs1!Deposito = Rs!Deposito
  End If
  If CmbOrigen.ListIndex <> -1 Then
     Rs1!Origen = IIf(CmbOrigen.Text = "Si", 1, 2)
  End If
  If CmbMoneda.ListIndex <> -1 Then
     Rs1!Moneda = CmbMoneda.ItemData(CmbMoneda.ListIndex)
  End If
  
  Rs1!Peso = TxtPeso.Text
  Rs1!Representante = TxtRep.Text
  Rs1!CantDespacho = TxtCantDes.Text
  Rs1!CantVenta = TxtCantVent.Text
  Rs1!FechaCompra = DTFecha.Value
  If TxtCompra.Text = "" Then TxtCompra.Text = 0
  Rs1!PrecioCompra = CCur(TxtCompra.Text)
  Rs1!Actualiza = False
  If TxtMinima.Text = "" Then TxtMinima.Text = 0
  Rs1!CantMinima = TxtMinima.Text
  If TxtMaxima.Text = "" Then TxtMaxima.Text = 0
  Rs1!CantMaxima = TxtMaxima.Text
  Rs1!ActuaStock = CmbStock.Text
  Rs1!CodCompu = CmbCompuesto.Text
  Rs1!fijo = CmbFijo.Text
  Rs1!TipoInterno = CmbInterno.Text
  Rs1!ValorInterno = TxtvInterno.Text
  Rs1!Calcular = CmbCalcu.Text
  Rs1!CBarra = "" & TxtCBarra.Text
  Rs1!ValorEnvase = TxtEmbase.Text
  Rs1!Balanza = CmbBalanza.Text
  Rs1!CDepartamento = CmbCDepartamento.Text
  Rs1!CFamilia = CmbCFamilia.Text
  Rs1.Update
  ' Carga Detalle de Impuestos
  
  Rs2.Fields.Append "Producto", adChar, 20
  Rs2.Fields.Append "Impuesto", adInteger
  Rs2.Open
  
  For i = 0 To LstImpuestos.ListCount - 1
    If LstImpuestos.Selected(i) = True Then
       Rs2.AddNew
       Rs2!Producto = TxtCodigo.Text
       Rs2!Impuesto = LstImpuestos.ItemData(i)
       Rs2.Update
    End If
  Next i
        
  If Rs2.RecordCount = 0 Then
     MsgBox "No se ha Grabado el Impuesto en el Producto", vbCritical, "Atención"
     Exit Sub
  End If
  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Modificar Rs1, Rs2, sCod1, nRepAux
  Else
     Set Rs3 = cRsl.TraerRS("ListaPrecioTraerTodos", , True)
     cRse.Agregar Rs1, Rs2, Rs3, nRepAux
  End If
  
  Rs1.Close
  Rs2.Close
  Set cRse = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub ArmaRsParaMovimiento()
  
  
  
  Set Rs_Pro = cRsl.TraerTodos("Productos", "Id")
  Botones pEdit2, False, False, False, False, False, False, False, False, True
  If Rs_Pro.BOF = False And Rs_Pro.EOF = False Then
     Rs_Pro.MoveLast
     Set Rs = cRsl.TraerRS("ProductoTraerUno", Rs_Pro!ID, True)
     LinkearTexto
     Botones pEdit2, pBorr2, pEdit2, True, True, True, True, True, True, True
  End If
  
End Sub

Private Sub Siguiente()
  
  On Error GoTo errHandler
  Rs_Pro.MoveNext
  If Rs_Pro.EOF Then
     Rs_Pro.MoveLast
     Beep
  End If
  
  Set Rs = cRsl.TraerRS("ProductoTraerUno", Rs_Pro!ID, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Anterior()
  
  On Error GoTo errHandler
  Rs_Pro.MovePrevious
  If Rs_Pro.BOF Then
     Rs_Pro.MoveFirst
     Beep
  End If
  
  Set Rs = cRsl.TraerRS("ProductoTraerUno", Rs_Pro!ID, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub Primero()
  
  On Error GoTo errHandler
  Rs_Pro.MoveFirst
  
  Set Rs = cRsl.TraerRS("ProductoTraerUno", Rs_Pro!ID, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub ultimo()
  
  On Error GoTo errHandler
  Rs_Pro.MoveLast
  
  Set Rs = cRsl.TraerRS("ProductoTraerUno", Rs_Pro!ID, True)
  
  LinkearTexto
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub TxtCBarra_Change()
  TodoBien
End Sub

Private Sub TxtCodigo_Change()
  TodoBien
End Sub

Private Sub TxtCodigo_LostFocus()
  Dim Rs1 As ADODB.Recordset
  
  Set Rs1 = New ADODB.Recordset

  If CmdBotones(2).Caption <> "Grabar" Then
     If TxtCodigo.Text <> "" Then
        Set Rs1 = cRsl.TraerRS("ProductoTraerUnoS", TxtCodigo.Text, True)
        If Rs1.EOF = False Or Rs1.BOF = False Then
           MsgBox "Codigo de Producto Repetido", vbCritical, "Atención"
           TxtCodigo.SetFocus
           TxtCodigo.Text = ""
        End If
     End If
  End If
  
End Sub

Private Sub Buscar()
  
  
  nDat = 0
  nBuscar = 3
  bBuscarActivo = False
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs_Pro = cRsl.TraerTodos("Productos", "Id")
     Rs_Pro.MoveFirst
     Rs_Pro.Find "Id=" & nDat
     Set Rs = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     LinkearTexto
  End If
  bBuscarActivo = True
  
  
End Sub

Private Sub TxtCompra_GotFocus()
  TxtCompra.SelStart = 0
  TxtCompra.SelLength = Len(TxtCompra.Text)
End Sub

Private Sub TxtCompra_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtCompra) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCompra_LostFocus()
  TxtCompra.Text = Format(TxtCompra.Text, nCantDecimales)
End Sub
Private Sub TxtDescripcion_Change()
  TodoBien
End Sub

Private Sub TxtDescripcion1_Change()
  TodoBien
End Sub

Private Sub TxtEmbase_Change()
  TodoBien
End Sub

Private Sub TxtMaxima_GotFocus()
  TxtMaxima.SelStart = 0
  TxtMaxima.SelLength = Len(TxtMaxima.Text)
End Sub
Private Sub TxtMaxima_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtMaxima) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtMaxima_LostFocus()
  TxtMaxima.Text = Format(TxtMaxima.Text, nCantDecimales)
End Sub
Private Sub TxtMinima_GotFocus()
  TxtMinima.SelStart = 0
  TxtMinima.SelLength = Len(TxtMinima.Text)
End Sub
Private Sub TxtMinima_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtMinima) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtMinima_LostFocus()
  TxtMinima.Text = Format(TxtMinima.Text, nCantDecimales)
End Sub
Private Sub TodoBien()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtCodigo.Text) <> 0 And Len(TxtDescripcion.Text) <> 0 And Len(CmbFamilia.Text) <> 0 And Len(CmbMarca.Text) <> 0 _
           And Len(CmbFamilia.Text) <> 0 And Len(CmbProveedor.Text) <> 0 And Len(CmbCompra.Text) <> 0 And Len(CmbVenta.Text) <> 0 _
           And Len(CmbDeposito.Text) <> 0 And LstImpuestos.SelCount = 1 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(2).Enabled = False
        If Len(TxtCodigo.Text) <> 0 And Len(TxtDescripcion.Text) <> 0 And Len(CmbFamilia.Text) <> 0 And Len(CmbMarca.Text) <> 0 _
           And Len(CmbFamilia.Text) <> 0 And Len(CmbProveedor.Text) <> 0 And Len(CmbCompra.Text) <> 0 And Len(CmbVenta.Text) <> 0 _
           And Len(CmbDeposito.Text) <> 0 And LstImpuestos.SelCount = 1 Then
           CmdBotones(2).Enabled = True
        End If
     End If
  End If
End Sub

Private Sub TxtPeso_Change()
  TodoBien
End Sub

Private Sub TxtRep_Change()
  TodoBien
End Sub

Private Sub TxtvInterno_Change()
  TodoBien
End Sub

Private Sub TxtvInterno_GotFocus()
  TxtvInterno.SelStart = 0
  TxtvInterno.SelLength = Len(TxtvInterno.Text)
End Sub

Private Sub TxtvInterno_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtvInterno) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtvInterno_LostFocus()
  If TxtvInterno.Text = "" Then TxtvInterno.Text = 0
   TxtvInterno.Text = Format(TxtvInterno.Text, "0.0000")
End Sub

Private Sub TxtPeso_GotFocus()
  TxtPeso.SelStart = 0
  TxtPeso.SelLength = Len(TxtPeso.Text)
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtPeso) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPeso_LostFocus()
  If TxtPeso.Text = "" Then TxtPeso.Text = 0
  TxtPeso.Text = Format(TxtPeso.Text, "0.000")
End Sub

Private Sub TxtRep_GotFocus()
  TxtRep.SelStart = 0
  TxtRep.SelLength = Len(TxtRep.Text)
End Sub

Private Sub TxtRep_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtRep_LostFocus()
  If TxtRep.Text = "" Then TxtRep.Text = 0
  If Val(TxtRep.Text) = 0 Then nRepAux = 0
End Sub

Private Sub TxtEmbase_GotFocus()
  TxtEmbase.SelStart = 0
  TxtEmbase.SelLength = Len(TxtEmbase.Text)
End Sub

Private Sub TxtEmbase_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtEmbase) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtEmbase_LostFocus()
  TxtEmbase.Text = Format(TxtEmbase.Text, nCantDecimales)
End Sub
