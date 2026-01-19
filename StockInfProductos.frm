VERSION 5.00
Begin VB.Form StockInfProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Productos"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   8970
   Begin VB.CommandButton Command1 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   7545
      TabIndex        =   6
      Top             =   3465
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   90
      TabIndex        =   7
      Top             =   120
      Width           =   8760
      Begin VB.Frame Frame5 
         Caption         =   "Marcas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   5580
         TabIndex        =   13
         Top             =   1755
         Width           =   3075
         Begin VB.ComboBox CmbVarios 
            Height          =   315
            Index           =   1
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   795
            Width           =   2235
         End
         Begin VB.ComboBox CmbVarios 
            Height          =   315
            Index           =   0
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   345
            Width           =   2235
         End
         Begin VB.Label Label6 
            Caption         =   "Hasta :"
            Height          =   195
            Left            =   105
            TabIndex        =   17
            Top             =   795
            Width           =   585
         End
         Begin VB.Label Label5 
            Caption         =   "Desde :"
            Height          =   210
            Left            =   90
            TabIndex        =   16
            Top             =   345
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   180
         TabIndex        =   12
         Top             =   1755
         Width           =   5310
         Begin VB.CommandButton CmdProveedor 
            Caption         =   "Buscar"
            Height          =   300
            Index           =   1
            Left            =   4560
            TabIndex        =   26
            Top             =   795
            Width           =   660
         End
         Begin VB.CommandButton CmdProveedor 
            Caption         =   "Buscar"
            Height          =   300
            Index           =   0
            Left            =   4560
            TabIndex        =   25
            Top             =   390
            Width           =   660
         End
         Begin VB.TextBox TxtProveedor 
            Height          =   300
            Index           =   1
            Left            =   690
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   22
            Text            =   "TxtProveedor"
            Top             =   810
            Width           =   735
         End
         Begin VB.TextBox TxtProveedor 
            Height          =   300
            Index           =   0
            Left            =   690
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   21
            Text            =   "TxtProveedor"
            Top             =   390
            Width           =   735
         End
         Begin VB.Label LblProveedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProveedor"
            Height          =   300
            Index           =   1
            Left            =   1470
            TabIndex        =   24
            Top             =   810
            Width           =   3045
         End
         Begin VB.Label LblProveedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProveedor"
            Height          =   300
            Index           =   0
            Left            =   1470
            TabIndex        =   23
            Top             =   390
            Width           =   3045
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta :"
            Height          =   210
            Left            =   75
            TabIndex        =   15
            Top             =   795
            Width           =   570
         End
         Begin VB.Label Label3 
            Caption         =   "Desde :"
            Height          =   165
            Left            =   45
            TabIndex        =   14
            Top             =   375
            Width           =   600
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   1875
         TabIndex        =   9
         Top             =   225
         Width           =   6780
         Begin VB.CheckBox ChkAlfabetico 
            Caption         =   "Alfabetico"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   1260
         End
         Begin VB.CommandButton CmdProducto 
            Caption         =   "Buscar"
            Height          =   300
            Index           =   1
            Left            =   5865
            TabIndex        =   29
            Top             =   675
            Width           =   720
         End
         Begin VB.TextBox TxtProducto 
            Height          =   300
            Index           =   1
            Left            =   720
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   27
            Text            =   "TxtProducto"
            Top             =   675
            Width           =   1470
         End
         Begin VB.CommandButton CmdProducto 
            Caption         =   "Buscar"
            Height          =   300
            Index           =   0
            Left            =   5865
            TabIndex        =   20
            Top             =   255
            Width           =   720
         End
         Begin VB.TextBox TxtProducto 
            Height          =   300
            Index           =   0
            Left            =   720
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   18
            Text            =   "TxtProducto"
            Top             =   255
            Width           =   1470
         End
         Begin VB.Label LblProducto 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProducto"
            Height          =   285
            Index           =   1
            Left            =   2235
            TabIndex        =   28
            Top             =   675
            Width           =   3570
         End
         Begin VB.Label LblProducto 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProducto"
            Height          =   285
            Index           =   0
            Left            =   2235
            TabIndex        =   19
            Top             =   255
            Width           =   3570
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta :"
            Height          =   195
            Left            =   75
            TabIndex        =   11
            Top             =   675
            Width           =   540
         End
         Begin VB.Label Label1 
            Caption         =   "Desde :"
            Height          =   180
            Left            =   75
            TabIndex        =   10
            Top             =   255
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   150
         TabIndex        =   8
         Top             =   225
         Width           =   1590
         Begin VB.OptionButton OptGrupo 
            Caption         =   "Ninguno"
            Height          =   210
            Index           =   3
            Left            =   240
            TabIndex        =   0
            Top             =   1095
            Width           =   1020
         End
         Begin VB.OptionButton OptGrupo 
            Caption         =   "Proveedor"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Top             =   210
            Width           =   1110
         End
         Begin VB.OptionButton OptGrupo 
            Caption         =   "Marca"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   2
            Top             =   530
            Width           =   960
         End
         Begin VB.OptionButton OptGrupo 
            Caption         =   "Familia"
            Height          =   210
            Index           =   2
            Left            =   240
            TabIndex        =   3
            Top             =   805
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "StockInfProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nOrd As Byte, sDesde As String, sHasta As String
Dim cRsl As ClsLectura
Private Sub CmbVarios_Click(Index As Integer)
  sDesde = CmbVarios(0).text
  sHasta = CmbVarios(1).text
End Sub
Private Sub CmdProducto_Click(Index As Integer)
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto(Index) = RsP!Producto
     LblProducto(Index).Caption = RsP!Descripcion
     If Index = 0 Then
        TxtProducto(1).text = RsP!Producto
        LblProducto(1).Caption = RsP!Descripcion
     End If
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CmdProveedor_Click(Index As Integer)
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  nBuscar = 2
  nDat = 0
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProveedorTraerUno", nDat, True)
     TxtProveedor(Index) = RsP!Proveedor
     LblProveedor(Index).Caption = RsP!RazonSocial
     sDesde = LblProveedor(0).Caption
     sHasta = LblProveedor(1).Caption
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Command1_Click()
  nImpr = 20
  FrmImpresor.Show
End Sub

Private Sub Form_Load()
  Limpiar
End Sub

Private Sub OptGrupo_Click(Index As Integer)
  Set cRsl = New ClsLectura
  Select Case Index
         Case 0
              nOrd = 8
              CmbVarios(0).Clear
              CmbVarios(1).Clear
              Habilitar True, False
         Case 1
              nOrd = 3
              Habilitar False, True
              Frame5.Caption = "Marcas"
              cRsl.CargaCombo CmbVarios(0), "Marcas", "Marca", "Descripcion", ""
              cRsl.CargaCombo CmbVarios(1), "Marcas", "Marca", "Descripcion", ""
              CmbVarios(0).ListIndex = 0
              CmbVarios(1).ListIndex = CmbVarios(1).ListCount - 1
         Case 2
              nOrd = 7
              Habilitar False, True
              Frame5.Caption = "Familias"
              cRsl.CargaCombo CmbVarios(0), "Familias", "Familia", "Descripcion", ""
              cRsl.CargaCombo CmbVarios(1), "Familias", "Familia", "Descripcion", ""
              CmbVarios(0).ListIndex = 0
              CmbVarios(1).ListIndex = CmbVarios(1).ListCount - 1
         Case 3
              nOrd = 0
              sDesde = ""
              sHasta = ""
              Habilitar False, False
              CmbVarios(0).Clear
              CmbVarios(1).Clear
  End Select
End Sub
Private Sub Habilitar(bProve As Boolean, bMarca As Boolean)
  Frame4.Enabled = bProve
  Frame5.Enabled = bMarca
End Sub

Private Sub Limpiar()
 Dim i As Byte
 For i = 0 To 1
     TxtProducto(i).text = ""
     TxtProveedor(i).text = ""
     LblProducto(i).Caption = ""
     LblProveedor(i).Caption = ""
 Next i
End Sub

