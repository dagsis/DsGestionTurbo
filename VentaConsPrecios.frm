VERSION 5.00
Begin VB.Form VentaConsPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas de Precios de Venta"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   7710
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
      Height          =   3510
      Left            =   90
      TabIndex        =   6
      Top             =   30
      Width           =   7440
      Begin VB.CheckBox ChkRep 
         Caption         =   "Actualiza Representantes"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   2580
         Width           =   2430
      End
      Begin VB.CheckBox ChkActualiza 
         Caption         =   "Permitir Actualizar Precio de Venta"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   2595
         Width           =   3195
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   225
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1770
         Width           =   3510
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   5595
         TabIndex        =   3
         Top             =   2535
         Width           =   1665
      End
      Begin VB.ComboBox CmbLista 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   510
         Width           =   2895
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "..."
         Height          =   270
         Left            =   5520
         TabIndex        =   1
         Top             =   1140
         Width           =   405
      End
      Begin VB.TextBox TxtProducto 
         Height          =   285
         Left            =   150
         MaxLength       =   25
         TabIndex        =   0
         Text            =   "TxtProducto"
         Top             =   1155
         Width           =   1050
      End
      Begin VB.Label LblMoneda 
         Caption         =   "Moneda"
         Height          =   210
         Left            =   1755
         TabIndex        =   13
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label LblPrecio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3810
         TabIndex        =   12
         Top             =   1770
         Width           =   3510
      End
      Begin VB.Label Label4 
         Caption         =   "Precio Venta $ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3780
         TabIndex        =   11
         Top             =   1545
         Width           =   1530
      End
      Begin VB.Label Label3 
         Caption         =   "F.Ult.Compra :"
         Height          =   195
         Left            =   4410
         TabIndex        =   10
         Top             =   255
         Width           =   1335
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
         Left            =   4545
         TabIndex        =   9
         Top             =   510
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Lista :"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Importe Moneda :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
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
         Left            =   1260
         TabIndex        =   4
         Top             =   1155
         Width           =   4200
      End
   End
End
Attribute VB_Name = "VentaConsPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura
Dim Cadena As String, ic As Byte, nCotizacion As Double
Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmbLista_Click()
  If TxtPrecio.Text <> "" Then
    TxtProducto_LostFocus
  End If
End Sub

Private Sub Command1_Click()
  If ChkActualiza.Value = 1 Then
     ActualizoPrecio
  End If
  Limpiar
  TxtProducto.SetFocus
End Sub
Private Sub ActualizoPrecio()
  On Error GoTo Errores
  Dim cPro As ClsPrecios, nRep As Long, cRsl As ClsLectura, cProdu As ClsProductoL
  Dim cLista As Recordset, sSql As String
  
  Set cPro = New ClsPrecios
  Set cRsl = New ClsLectura
  Set cProdu = New ClsProductoL
  Set cLista = New Recordset
  
  If cProdu.TraerPrecio(CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.Text) <> TxtPrecio.Text Then
  
      If ChkRep.Value = 1 Then
         nRep = cRsl.TraerValorDeUnCampo("Productos", "Representante", "Producto='" & TxtProducto.Text & "'")
      End If
    
      cPro.ActulizarPrecios CmbLista.ItemData(CmbLista.ListIndex), TxtProducto.Text, TxtPrecio.Text, nRep
      
      Set cLista = cRsl.TraerTodos("ListaDePrecio", "*")
      
      If CmbLista.ItemData(CmbLista.ListIndex) = cRsl.RegPorDefecto("ListaPrecio") Then
        If cLista.RecordCount > 1 Then
           db.Open sDb
           db.BeginTrans
  
           Do While Not cLista.EOF
              sSql = "Update Precios SET Precio=" & TxtPrecio.Text + (TxtPrecio.Text * cLista!Porcentaje / 100) & " WHERE Lista=" & cLista!ListaPrecio & " and Producto='" & TxtProducto.Text & "'"
              db.Execute sSql
              cLista.MoveNext
           Loop
           db.CommitTrans
           db.Close
        End If
      End If
      
 End If
  Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "Atención"
  db.RollbackTrans
  db.Close

Exit Sub

errHandler:
   ManejaErrores

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

Private Sub TxtPrecio_LostFocus()
  If TxtPrecio.Text = "" Then TxtPrecio.Text = 0
  LblPrecio.Caption = TxtPrecio.Text * nCotizacion
  LblPrecio.Caption = Format(LblPrecio.Caption, nCantDecimales)
  TxtPrecio.Text = Format(TxtPrecio.Text, nCantDecimales)
End Sub

Private Sub TxtProducto_GotFocus()
  TxtProducto.SelStart = 0
  TxtProducto.SelLength = Len(TxtProducto.Text)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub TxtProducto_LostFocus()
  Dim cDes As ClsProductoL, cRsl As New ClsLectura, nPrecio As Currency, nCoti As Double
  Dim Rs As Recordset
  
  Set cDes = New ClsProductoL
  
  If TxtProducto.Text <> "" Then
     TxtProducto.Text = cRsl.TraerCodigoProducto(TxtProducto.Text)
  
      Set Rs = cDes.TraerPrecioProducto(TxtProducto.Text, CmbLista.ItemData(CmbLista.ListIndex))
      
      TxtProducto.Text = ""
      If Rs.RecordCount <> 0 Then
         TxtProducto.Text = Rs!Producto
         LblProducto.Caption = Rs!Descripcion
         TxtPrecio.Text = Format(Rs!precio, nCantDecimales)
         LblFecha.Caption = Rs!fecha
         LblPrecio.Caption = Format(Rs!importe, nCantDecimales)
         LblMoneda.Caption = Rs!Moneda
         nCotizacion = Rs!COTIZACION
         TxtPrecio.SetFocus
      Else
         MsgBox "Producto no encontrado", vbInformation, "Atención"
         TxtProducto.SetFocus
      End If
  End If
  
End Sub

Private Sub Limpiar()
  LblFecha.Caption = ""
  TxtProducto.Text = ""
  LblProducto.Caption = ""
  LblPrecio.Caption = nCantDecimales
  TxtPrecio.Text = nCantDecimales
  LblMoneda.Caption = ""
End Sub

Private Sub CmdBuscar_Click()
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL, Rs As Recordset
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  
  nLista = CmbLista.ItemData(CmbLista.ListIndex)

  nDat = 0
  nBuscar = 3
  FrmBuscarEx.Show 1
  Limpiar
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto = RsP!Producto
     Set Rs = cRpl.TraerPrecioProducto(RsP!Producto, CmbLista.ItemData(CmbLista.ListIndex))
     If Rs.RecordCount <> 0 Then
        TxtProducto.Text = Rs!Producto
        LblProducto.Caption = Rs!Descripcion
        TxtPrecio.Text = Format(Rs!precio, nCantDecimales)
        LblFecha.Caption = Rs!fecha
        LblPrecio.Caption = Format(Rs!importe, nCantDecimales)
        LblMoneda.Caption = Rs!Moneda
        nCotizacion = Rs!COTIZACION

        TxtPrecio.SetFocus
        RsP.Close
     End If
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
  ChkActualiza.Value = IIf(cRsl.TraerValorDeUnCampo("Registros", "Venta", "Usuario=" & nUsuario) = True, 1, 0)
  ChkActualiza.Enabled = IIf(cRsl.TraerValorDeUnCampo("Registros", "Venta", "Usuario=" & nUsuario) = True, True, False)
  ChkRep.Enabled = IIf(cRsl.TraerValorDeUnCampo("Registros", "Venta", "Usuario=" & nUsuario) = True, True, False)
  ic = 1
End Sub

Private Sub CargarCombos()
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
  CmbLista.Text = cRsl.DatoCombo("ListaDePrecio", "ListaPrecio", "Descripcion", cRsl.RegPorDefecto("ListaPrecio"))
  Set cRsl = New ClsLectura
End Sub



