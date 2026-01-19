VERSION 5.00
Begin VB.Form ComprasAPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Precios de Venta Por Compras"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7470
   Begin VB.CommandButton CmbTodos 
      Caption         =   "Actualizar Todos Los Productos"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   5310
      Width           =   2610
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actualizar Por Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   120
      TabIndex        =   22
      Top             =   3885
      Width           =   7260
      Begin VB.TextBox TxtProveedor 
         Height          =   300
         Left            =   4500
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "TxtPr"
         Top             =   510
         Width           =   795
      End
      Begin VB.CommandButton CmdProvee 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   6120
         TabIndex        =   11
         Top             =   510
         Width           =   855
      End
      Begin VB.ComboBox CmbProveedor 
         Height          =   315
         Left            =   330
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   510
         Width           =   4110
      End
      Begin VB.Label Label5 
         Caption         =   "% "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   24
         Top             =   525
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Actualizar Por Familia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   90
      TabIndex        =   21
      Top             =   2460
      Width           =   7305
      Begin VB.TextBox TxtFamilia 
         Height          =   300
         Left            =   4500
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "TxtFa"
         Top             =   450
         Width           =   795
      End
      Begin VB.CommandButton CmdFamilia 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   6180
         TabIndex        =   8
         Top             =   435
         Width           =   855
      End
      Begin VB.ComboBox CmbFamilia 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   450
         Width           =   4110
      End
      Begin VB.Label Label1 
         Caption         =   "% "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   23
         Top             =   465
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actualizar Por Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   30
      TabIndex        =   12
      Top             =   0
      Width           =   7350
      Begin VB.TextBox TxtSubTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3435
         MaxLength       =   7
         TabIndex        =   4
         Text            =   "TxtVenta"
         Top             =   1635
         Width           =   2055
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         MaxLength       =   7
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1650
         Width           =   2055
      End
      Begin VB.CommandButton CmdProducto 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   6225
         TabIndex        =   5
         Top             =   1845
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "..."
         Height          =   270
         Left            =   6825
         TabIndex        =   1
         Top             =   1020
         Width           =   405
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   120
         MaxLength       =   25
         TabIndex        =   0
         Text            =   "TxtProducto"
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox TxtImpuestos 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2220
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1635
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "F.Ult.Compra :"
         Height          =   195
         Left            =   5820
         TabIndex        =   20
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label LblFecha 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblFecha"
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
         Left            =   5835
         TabIndex        =   19
         Top             =   525
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "Imp de Compra :"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label LblProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblProducto"
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
         Left            =   1230
         TabIndex        =   17
         Top             =   1020
         Width           =   5520
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
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   525
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Impuestos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2220
         TabIndex        =   14
         Top             =   1380
         Width           =   1185
      End
      Begin VB.Label Label9 
         Caption         =   "Sub Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3420
         TabIndex        =   13
         Top             =   1380
         Width           =   1185
      End
   End
End
Attribute VB_Name = "ComprasAPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cRsl As ClsLectura
Dim Cadena As String, ic As Byte, nProv As Long
Private Sub CmdSalir_Click()
  Unload Me
End Sub


Private Sub CmbTodos_Click()
  If MsgBox("Esta Seguro de Actualizar Todos Los Productos ?", 20, "Atención") = 6 Then
    Dim sCondi As String
  
  sCondi = "Todos"
  ActualizoPrecio sCondi, 0
  
  Limpiar
  TxtProducto.SetFocus

  End If
End Sub

Private Sub CmdFamilia_Click()
  Dim sCondi As String
  
  sCondi = "Familia=" & CmbFamilia.ItemData(CmbFamilia.ListIndex)
  ActualizoPrecio sCondi, CDbl(TxtFamilia.text)
  
  Limpiar
  TxtProducto.SetFocus
End Sub

Private Sub CmdProducto_Click()
  Dim sCondi As String
  sCondi = ""
  
  ActualizoPrecio sCondi, 0
  
  Limpiar
  TxtProducto.SetFocus
End Sub

Private Sub ActualizoPrecio(pCondi As String, pValor As Double)
  On Error GoTo errHandler
  Dim cPro As ClsPrecios, cRsl As ClsLectura, nCal As Double
  Set cPro = New ClsPrecios
  Set cRsl = New ClsLectura
  
  nCal = TxtSubTotal.text / ((TxtImpuestos.text / 100) + 1)
  
  cPro.ActulizarPreciosCostoLista TxtProducto.text, nCal, TxtSubTotal.text, pValor, pCondi
  
  MsgBox "Proceso Realizado Con Exito...", vbInformation, "Atención"
  Exit Sub

errHandler:
   ManejaErrores
End Sub


Private Sub CmdProvee_Click()
  Dim sCondi As String
  
  sCondi = "Proveedor=" & CmbProveedor.ItemData(CmbProveedor.ListIndex)
  ActualizoPrecio sCondi, CDbl(TxtProveedor.text)
  
  Limpiar
  TxtProducto.SetFocus

End Sub

Private Sub TxtImpuestos_LostFocus()
  Dim cRsl As New ClsProductoL
  
  If TxtPrecio.text <> "" Then
     TxtPrecio.text = Format(TxtPrecio.text, nCantDecimales)
     TxtSubTotal.text = TxtPrecio.text + (TxtPrecio.text * TxtImpuestos.text / 100)
     TxtSubTotal.text = Format(TxtSubTotal.text, nCantDecimales)
     TxtImpuestos.text = Format(TxtImpuestos.text, nCantDecimales)
  End If

End Sub

Private Sub TxtPrecio_LostFocus()
  Dim cRsl As New ClsProductoL
  Dim cSub As Double, cImp As Double, cPrecio As Double

  If TxtPrecio.text <> "" Then
     cPrecio = TxtPrecio.text
     cImp = TxtImpuestos.text
     TxtPrecio.text = Format(cPrecio, nCantDecimales)
     cSub = cPrecio + (cPrecio * cImp / 100)
     
     TxtSubTotal.text = Format(cSub, nCantDecimales)
  End If
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

Private Sub TxtImpuestos_GotFocus()
  TxtImpuestos.SelStart = 0
  TxtImpuestos.SelLength = Len(TxtImpuestos.text)
End Sub

Private Sub TxtImpuestos_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtImpuestos) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSubtotal_GotFocus()
  TxtSubTotal.SelStart = 0
  TxtSubTotal.SelLength = Len(TxtSubTotal.text)
End Sub

Private Sub TxtSubtotal_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtSubTotal) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtProducto_GotFocus()
  TxtProducto.SelStart = 0
  TxtProducto.SelLength = Len(TxtProducto.text)
End Sub

Private Sub TxtSubTotal_LostFocus()
  Dim cSub As Double, cImp As Double, cPrecio As Double
  
  If TxtPrecio.text <> "" Then
     cSub = TxtSubTotal.text
     cImp = TxtImpuestos.text
     
     cPrecio = cSub / ((cImp / 100) + 1)
       
     TxtSubTotal.text = Format(cSub, nCantDecimales)
     TxtPrecio.text = Format(cPrecio, nCantDecimales)
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
  '   KeyAscii = Upper(KeyAscii)
End Sub

Private Sub TxtProducto_LostFocus()
  Dim cDes As ClsProductoL, cRsl As ClsLectura
  Dim cSub As Double, cImp As Double, cPrecio As Double

  
  Set cDes = New ClsProductoL
  Set cRsl = New ClsLectura
  
  LblProducto.Caption = cDes.BuscarNombreProducto(TxtProducto.text)
  If LblProducto.Caption = "" Then
     TxtProducto.text = ""
  Else
     TxtPrecio.SetFocus
     cPrecio = cDes.TraerPrecioCompra(TxtProducto.text)
     TxtPrecio.text = Format(cPrecio, nCantDecimales)
     LblFecha.Caption = cRsl.TraerValorDeUnCampo("Productos", "FechaCompra", "Producto='" & TxtProducto.text & "'")
     nProv = cRsl.TraerValorDeUnCampo("Productos", "Proveedor", "Producto='" & TxtProducto.text & "'")
     LblProveedor.Caption = cRsl.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & nProv)
     
     cImp = cDes.ProductoImpuesto(TxtProducto.text)
     TxtImpuestos.text = Format(cImp, nCantDecimales)
     cSub = cPrecio + (cPrecio * cImp / 100)
     TxtSubTotal.text = Format(cSub, nCantDecimales)
     
     Set cRsl = Nothing
'     TxtProducto.SetFocus
  End If

End Sub

Private Sub Limpiar()
  LblFecha.Caption = ""
  LblProveedor.Caption = ""
  TxtImpuestos.text = "0.00"
  TxtSubTotal.text = "0.00"
  TxtProducto.text = ""
  LblProducto.Caption = ""
  TxtPrecio.text = nCantDecimales
  TxtFamilia.text = "0.00"
  TxtProveedor.text = "0.00"
End Sub

Private Sub CargarCombos()
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbFamilia, "Familias", "Familia", "Descripcion", ""
  cRsl.CargaCombo CmbProveedor, "Proveedores", "Id", "RazonSocial", ""
  
  Set cRsl = New ClsLectura
End Sub

Private Sub CmdBuscar_Click()
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL
  Dim cSub As Double, cImp As Double, cPrecio As Double

  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  
  nDat = 0
  nBuscar = 4
  FrmBuscarEx.Show 1
  Limpiar
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto = RsP!Producto
     LblProducto.Caption = RsP!Descripcion
     nProv = RsP!Proveedor
     cPrecio = cRpl.TraerPrecioCompra(TxtProducto.text)
     TxtPrecio.text = Format(cPrecio, nCantDecimales)
     
     LblFecha.Caption = cRsl.TraerValorDeUnCampo("Productos", "FechaCompra", "Producto='" & TxtProducto.text & "'")
     LblProveedor.Caption = cRsl.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & nProv)
     cImp = cRpl.ProductoImpuesto(TxtProducto.text)
     TxtImpuestos.text = Format(cImp, nCantDecimales)
     cSub = cPrecio + (cPrecio * cImp / 100)
     TxtSubTotal.text = Format(cSub, nCantDecimales)
     TxtPrecio.SetFocus
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Form_Load()
  Limpiar
  CargarCombos
End Sub
Private Sub TxtFamilia_GotFocus()
  TxtFamilia.SelStart = 0
  TxtFamilia.SelLength = Len(TxtFamilia.text)
End Sub

Private Sub TxtFamilia_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtFamilia) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtFamilia_LostFocus()
  If TxtFamilia.text = "" Then TxtFamilia.text = 0
  TxtFamilia.text = Format(TxtFamilia.text, "#0.00")
End Sub

Private Sub TxtProveedor_GotFocus()
  TxtProveedor.SelStart = 0
  TxtProveedor.SelLength = Len(TxtProveedor.text)
End Sub

Private Sub Txtproveedor_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtProveedor) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtProveedor_LostFocus()
  If TxtProveedor.text = "" Then TxtProveedor.text = 0
  TxtProveedor.text = Format(TxtProveedor, "#0.00")
End Sub
