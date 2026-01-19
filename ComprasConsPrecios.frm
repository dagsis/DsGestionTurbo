VERSION 5.00
Begin VB.Form ComprasConsPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas de Precios de Compras"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   10455
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   8985
      TabIndex        =   6
      Top             =   3795
      Width           =   1335
   End
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
      Left            =   3660
      MaxLength       =   12
      TabIndex        =   3
      Text            =   "TxtVenta"
      Top             =   2265
      Width           =   2460
   End
   Begin VB.Frame Frame1 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   30
      TabIndex        =   7
      Top             =   15
      Width           =   10320
      Begin VB.CheckBox ChkRep 
         Caption         =   "Actualiza Representantes"
         Height          =   195
         Left            =   3330
         TabIndex        =   23
         Top             =   2970
         Width           =   2430
      End
      Begin VB.CheckBox ChkActualiza 
         Caption         =   "Permitir Actualizar Precios de Venta"
         Height          =   240
         Left            =   270
         TabIndex        =   22
         Top             =   2955
         Width           =   3255
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
         Left            =   2385
         MaxLength       =   7
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2250
         Width           =   1155
      End
      Begin VB.TextBox TxtPorcentaje 
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
         Left            =   6165
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2250
         Width           =   1155
      End
      Begin VB.ComboBox CmbLista 
         Height          =   315
         Left            =   255
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   495
         Width           =   2895
      End
      Begin VB.TextBox TxtVenta 
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
         Left            =   7395
         TabIndex        =   5
         Text            =   "TxtVenta"
         Top             =   2250
         Width           =   2115
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   255
         MaxLength       =   25
         TabIndex        =   0
         Text            =   "TxtProducto"
         Top             =   1605
         Width           =   1050
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "..."
         Height          =   270
         Left            =   6960
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1605
         Width           =   405
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
         Left            =   240
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2250
         Width           =   2025
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
         Left            =   3645
         TabIndex        =   21
         Top             =   2025
         Width           =   1185
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
         Left            =   2355
         TabIndex        =   19
         Top             =   2025
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Lista :"
         Height          =   195
         Left            =   255
         TabIndex        =   18
         Top             =   255
         Width           =   540
      End
      Begin VB.Label Label6 
         Caption         =   "Precio Venta :"
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
         Left            =   7380
         TabIndex        =   16
         Top             =   2025
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Porc. Venta:"
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
         Left            =   6135
         TabIndex        =   15
         Top             =   2025
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   855
         Width           =   1335
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
         Left            =   255
         TabIndex        =   13
         Top             =   1110
         Width           =   5820
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
         Left            =   1365
         TabIndex        =   12
         Top             =   1605
         Width           =   5520
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
         Left            =   225
         TabIndex        =   11
         Top             =   2010
         Width           =   1515
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
         Left            =   6165
         TabIndex        =   10
         Top             =   1095
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "F.Ult.Compra :"
         Height          =   195
         Left            =   6150
         TabIndex        =   9
         Top             =   885
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Precio Venta :"
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
      Left            =   3135
      TabIndex        =   20
      Top             =   1980
      Width           =   1260
   End
End
Attribute VB_Name = "ComprasConsPrecios"
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

Private Sub CmbLista_Click()
  If TxtPrecio.Text <> "" Then
    TxtProducto_LostFocus
  End If
End Sub

Private Sub Command1_Click()
  ActualizoPrecio
  Limpiar
  TxtProducto.SetFocus
  
End Sub
Private Sub ActualizoPrecio()
  On Error GoTo errHandler
  Dim cProdu As ClsProductoL
  Dim cPro As ClsPrecios, nRep As Long, cRsl As ClsLectura
  Dim cLista As Recordset, sSql As String

  Set cProdu = New ClsProductoL
  Set cPro = New ClsPrecios
  Set cRsl = New ClsLectura

  
  If ChkActualiza.Value = 1 Then
     If cProdu.TraerPrecio(CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.Text) <> TxtVenta.Text Then
      If ChkRep.Value = 1 Then
        nRep = cRsl.TraerValorDeUnCampo("Productos", "Representante", "Producto='" & TxtProducto.Text & "'")
      End If
    End If
  End If
  
  cPro.ActulizarPreciosCosto CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.Text, TxtPrecio.Text, TxtPorcentaje.Text, TxtVenta.Text, nRep

  Set cLista = cRsl.TraerTodos("ListaDePrecio", "*")
      
  If CmbLista.ItemData(CmbLista.ListIndex) = cRsl.RegPorDefecto("ListaPrecio") Then
    If cLista.RecordCount > 1 Then
       db.Open sDb
       db.BeginTrans

       Do While Not cLista.EOF
          sSql = "Update Precios SET Precio=" & RedondearPrecios(TxtVenta.Text + (TxtVenta.Text * cLista!Porcentaje / 100)) & " WHERE Lista=" & cLista!ListaPrecio & " and Producto='" & TxtProducto.Text & "'"
          db.Execute sSql
          cLista.MoveNext
       Loop
       db.CommitTrans
       db.Close
    End If
  End If
  
  Exit Sub

errHandler:
   ManejaErrores

End Sub


Private Sub TxtImpuestos_LostFocus()
  Dim cRsl As New ClsProductoL
  If TxtImpuestos.Text = "" Then TxtImpuestos.Text = 0
  If ChkActualiza.Value = 1 Then
    If TxtPrecio.Text <> "" Then
       TxtPrecio.Text = Format(TxtPrecio.Text, nCantDecimales)
       If TxtPorcentaje.Text <> 0 Then
          TxtSubTotal.Text = TxtPrecio.Text + (TxtPrecio.Text * TxtImpuestos.Text / 100)
          TxtSubTotal.Text = Format(TxtSubTotal.Text, nCantDecimales)
          TxtVenta.Text = (TxtPorcentaje.Text * TxtSubTotal.Text / 100) + TxtSubTotal.Text
       Else
          TxtSubTotal.Text = TxtPrecio.Text + (TxtPrecio.Text * TxtImpuestos.Text / 100)
          TxtSubTotal.Text = Format(TxtSubTotal.Text, nCantDecimales)
        '  TxtVenta.Text = (TxtPorcentaje.Text * TxtSubTotal.Text / 100) + TxtSubTotal.Text
       End If
       TxtVenta.Text = RedondearPrecios(TxtVenta.Text)
       TxtVenta.Text = Format(TxtVenta.Text, nCantDecimales)
       TxtImpuestos.Text = Format(TxtImpuestos.Text, nCantDecimales)
    End If
  End If

End Sub

Private Sub TxtPrecio_LostFocus()
  Dim cRsl As New ClsProductoL
  
    If TxtPrecio.Text <> "" Then
          TxtPrecio.Text = Format(TxtPrecio.Text, nCantDecimales)
          If ChkActualiza.Value = 1 Then
             If TxtPorcentaje.Text <> 0 Then
                TxtSubTotal.Text = TxtPrecio.Text + (TxtPrecio.Text * TxtImpuestos.Text / 100)
                TxtSubTotal.Text = Format(TxtSubTotal.Text, nCantDecimales)
                TxtVenta.Text = (TxtPorcentaje.Text * TxtSubTotal.Text / 100) + TxtSubTotal.Text
             Else
                TxtSubTotal.Text = TxtPrecio.Text + (TxtPrecio.Text * TxtImpuestos.Text / 100)
                TxtSubTotal.Text = Format(TxtSubTotal.Text, nCantDecimales)
             '  TxtVenta.Text = (TxtPorcentaje.Text * TxtSubTotal.Text / 100) + TxtSubTotal.Text
            End If
      End If
      TxtVenta.Text = RedondearPrecios(TxtVenta.Text)
      TxtVenta.Text = Format(TxtVenta.Text, nCantDecimales)
    Else
       TxtPrecio.Text = Format(0, "0.00")
    End If
End Sub

Private Sub TxtPorcentaje_LostFocus()
 If ChkActualiza.Value = 1 Then
    If TxtPorcentaje.Text <> "" Then
       If TxtPrecio.Text <> "" Then
          TxtPrecio.Text = Format(TxtPrecio.Text, nCantDecimales)
          TxtSubTotal.Text = TxtPrecio.Text + (TxtPrecio.Text * TxtImpuestos.Text / 100)
          TxtSubTotal.Text = Format(TxtSubTotal.Text, nCantDecimales)
          TxtVenta.Text = (TxtPorcentaje.Text * TxtSubTotal.Text / 100) + TxtSubTotal.Text
          TxtVenta.Text = RedondearPrecios(TxtVenta.Text)
          TxtVenta.Text = Format(TxtVenta.Text, nCantDecimales)
          TxtPorcentaje.Text = Format(TxtPorcentaje.Text, nCantDecimales)
      End If
    Else
       TxtPorcentaje.Text = "0.00"
    End If
  End If
End Sub

Private Sub TxtPrecio_GotFocus()
  TxtPrecio.SelStart = 0
  TxtPrecio.SelLength = Len(TxtPrecio.Text)
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtPrecio) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtImpuestos_GotFocus()
  TxtImpuestos.SelStart = 0
  TxtImpuestos.SelLength = Len(TxtImpuestos.Text)
End Sub

Private Sub TxtImpuestos_KeyPress(KeyAscii As Integer)
If SoloNumeroDecimalFinal(KeyAscii, TxtImpuestos) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSubtotal_GotFocus()
  TxtSubTotal.SelStart = 0
  TxtSubTotal.SelLength = Len(TxtSubTotal.Text)
End Sub

Private Sub TxtSubtotal_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtSubTotal) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtProducto_GotFocus()
  TxtProducto.SelStart = 0
  TxtProducto.SelLength = Len(TxtProducto.Text)
End Sub

Private Sub TxtPorcentaje_GotFocus()
  TxtPorcentaje.SelStart = 0
  TxtPorcentaje.SelLength = Len(TxtPorcentaje.Text)
End Sub

Private Sub TxtPorcentaje_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtPorcentaje) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSubTotal_LostFocus()
 If ChkActualiza.Value = 1 Then
    If TxtPrecio.Text <> "" Then
         If TxtSubTotal.Text = "" Then TxtSubTotal.Text = 0
         If TxtImpuestos.Text = "" Then TxtImpuestos.Text = 0
         TxtPrecio.Text = CDec(TxtSubTotal.Text) / ((CDec(TxtImpuestos.Text) / 100) + 1)
         TxtVenta.Text = (TxtSubTotal.Text * TxtPorcentaje.Text / 100) + TxtSubTotal.Text
         TxtSubTotal.Text = Format(TxtSubTotal.Text, nCantDecimales)
         TxtPrecio.Text = Format(TxtPrecio.Text, nCantDecimales)
         TxtVenta.Text = RedondearPrecios(TxtVenta.Text)
         TxtVenta.Text = Format(TxtVenta.Text, nCantDecimales)
    End If
  End If
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


Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
  '   KeyAscii = Upper(KeyAscii)
End Sub

Private Sub TxtProducto_LostFocus()
  If TxtProducto.Text <> "" Then
  
     Dim cDes As ClsProductoL, cRsl As ClsLectura
  
  
     Set cDes = New ClsProductoL
     Set cRsl = New ClsLectura
      TxtProducto.Text = cRsl.TraerCodigoProducto(TxtProducto.Text)
    
      
      LblProducto.Caption = cDes.BuscarNombreProducto(TxtProducto.Text)
      If LblProducto.Caption = "" Then
         MsgBox "Producto no encontrado o inactivo", vbInformation, "Atención"
         TxtProducto.Text = ""
         TxtProducto.SetFocus
      Else
         TxtPrecio.SetFocus
         TxtPrecio.Text = Format(cDes.TraerPrecioCompra(TxtProducto.Text), nCantDecimales)
         LblFecha.Caption = cRsl.TraerValorDeUnCampo("Productos", "FechaCompra", "Producto='" & TxtProducto.Text & "'")
         nProv = cRsl.TraerValorDeUnCampo("Productos", "Proveedor", "Producto='" & TxtProducto.Text & "'")
         LblProveedor.Caption = cRsl.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & nProv)
         TxtImpuestos.Text = Format(cDes.ProductoImpuesto(TxtProducto.Text), nCantDecimales)
         TxtSubTotal.Text = TxtPrecio.Text + (TxtPrecio.Text * TxtImpuestos.Text / 100)
         TxtSubTotal.Text = Format(TxtSubTotal.Text, nCantDecimales)
         TxtPorcentaje.Text = Format(cRsl.TraerValorDeUnCampo("Precios", "Porcentaje", "Producto='" & TxtProducto.Text & "' and Lista=" & CmbLista.ItemData(CmbLista.ListIndex)), "0.00")
         If TxtPorcentaje.Text = 0 Then
            TxtPorcentaje.Text = Format(cRsl.TraerValorDeUnCampo("ListaDePrecio", "Porcentaje", "ListaPrecio=" & CmbLista.ItemData(CmbLista.ListIndex)), "0.00")
         End If
         TxtVenta.Text = Format(cDes.TraerPrecio(CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.Text), nCantDecimales)
    
         Set cRsl = Nothing
    '     TxtProducto.SetFocus
      End If
  End If

End Sub

Private Sub Limpiar()
  LblFecha.Caption = ""
  LblProveedor.Caption = ""
  TxtVenta.Text = "0.00"
  TxtImpuestos.Text = "0.00"
  TxtSubTotal.Text = "0.00"
  TxtPorcentaje.Text = "0.00"
  TxtProducto.Text = ""
  LblProducto.Caption = ""
  TxtPrecio.Text = nCantDecimales
End Sub

Private Sub CargarCombos()
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
  CmbLista.Text = cRsl.DatoCombo("ListaDePrecio", "ListaPrecio", "Descripcion", cRsl.RegPorDefecto("ListaPrecio"))
  Set cRsl = New ClsLectura
End Sub

Private Sub CmdBuscar_Click()
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  
  nLista = CmbLista.ItemData(CmbLista.ListIndex)
  nDat = 0
  
  nDat = 0
  nBuscar = 4
  FrmBuscarEx.Show 1
  Limpiar
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto = RsP!Producto
     LblProducto.Caption = RsP!Descripcion
     nProv = RsP!Proveedor
     TxtPrecio.Text = Format(cRpl.TraerPrecioCompra(TxtProducto.Text), nCantDecimales)
     LblFecha.Caption = cRsl.TraerValorDeUnCampo("Productos", "FechaCompra", "Producto='" & TxtProducto.Text & "'")
     LblProveedor.Caption = cRsl.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & nProv)
     TxtImpuestos.Text = Format(cRpl.ProductoImpuesto(TxtProducto.Text), "#0.00")
     TxtPorcentaje.Text = Format(cRsl.TraerValorDeUnCampo("Precios", "Porcentaje", "Producto='" & TxtProducto.Text & "' and Lista=" & CmbLista.ItemData(CmbLista.ListIndex)), "0.00")
     If TxtPorcentaje.Text = 0 Then
        TxtPorcentaje.Text = cRsl.TraerValorDeUnCampo("ListaDePrecio", "Porcentaje", "ListaPrecio=" & CmbLista.ItemData(CmbLista.ListIndex))
     End If
     TxtSubTotal.Text = TxtPrecio.Text + (TxtPrecio.Text * TxtImpuestos.Text / 100)
     TxtSubTotal.Text = Format(TxtSubTotal.Text, nCantDecimales)
     TxtVenta.Text = Format(cRpl.TraerPrecio(CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.Text), nCantDecimales)
     TxtPrecio.SetFocus
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Form_Load()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  
  Limpiar
  CargarCombos
 
  Cadena = ""
  ChkActualiza.Value = IIf(cRsl.TraerValorDeUnCampo("Registros", "Venta", "") = True, 1, 0)
  ic = 1
End Sub

Private Sub TxtVenta_LostFocus()
  If TxtVenta.Text <> "" Then
     TxtVenta.Text = Format(TxtVenta.Text, "0.00")
'     If ChkActualiza.Value = 1 Then
'        If TxtPorcentaje.text <> 0 Then
'           TxtVenta.text = (TxtPorcentaje.text * TxtSubTotal.text / 100) + TxtSubTotal.text
'        End If
'     End If
'     TxtVenta.text = RedondearPrecios(TxtVenta.text)
'     TxtVenta.text = Format(TxtVenta.text, "#0.00")
  End If
End Sub
