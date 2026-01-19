VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentaActuPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Precios"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   13530
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   75
      TabIndex        =   5
      Top             =   60
      Width           =   13395
      Begin VB.Frame Frame4 
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Left            =   180
         TabIndex        =   14
         Top             =   2032
         Width           =   4260
         Begin VB.ComboBox CmbMarcas 
            Height          =   315
            Left            =   195
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1800
            Width           =   2595
         End
         Begin VB.ComboBox CmbProveedor 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   480
            Width           =   3750
         End
         Begin VB.CheckBox Check1 
            Caption         =   "O.Código"
            Height          =   240
            Left            =   3060
            TabIndex        =   18
            Top             =   150
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.ComboBox CmbFamilias 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1110
            Width           =   2595
         End
         Begin VB.CommandButton CmdTodos 
            Caption         =   "Todos"
            Height          =   270
            Left            =   3240
            TabIndex        =   15
            Top             =   1125
            Width           =   810
         End
         Begin VB.Label Label5 
            Caption         =   "Marcas :"
            Height          =   225
            Left            =   180
            TabIndex        =   23
            Top             =   1545
            Width           =   810
         End
         Begin VB.Label Label3 
            Caption         =   "Proveedor :"
            Height          =   225
            Left            =   225
            TabIndex        =   20
            Top             =   225
            Width           =   945
         End
         Begin VB.Label Label4 
            Caption         =   "Familias :"
            Height          =   225
            Left            =   225
            TabIndex        =   17
            Top             =   840
            Width           =   810
         End
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   5550
         Left            =   4575
         TabIndex        =   13
         Top             =   315
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   9790
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
      Begin VB.ComboBox CmbLista 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Caption         =   "Copiar Precios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   180
         TabIndex        =   10
         Top             =   735
         Width           =   4260
         Begin VB.CommandButton CmdCajas 
            Caption         =   "Inicializar Cajas"
            Height          =   375
            Left            =   1320
            TabIndex        =   21
            Top             =   735
            Width           =   1800
         End
         Begin VB.ComboBox CmbCopiar 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   345
            Width           =   2745
         End
         Begin VB.CommandButton CmdCopiar 
            Caption         =   "&Copiar"
            Height          =   375
            Left            =   3180
            TabIndex        =   2
            Top             =   735
            Width           =   930
         End
         Begin VB.Label Label2 
            Caption         =   "Copiar  a  :"
            Height          =   255
            Left            =   195
            TabIndex        =   11
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   180
         TabIndex        =   6
         Top             =   4380
         Width           =   4260
         Begin VB.OptionButton OptImporte 
            Caption         =   "Por Importe Fijo"
            Height          =   240
            Left            =   240
            TabIndex        =   8
            Top             =   495
            Width           =   1425
         End
         Begin VB.OptionButton OptPorcentaje 
            Caption         =   "Por Porcentaje"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   870
            Width           =   1560
         End
         Begin VB.TextBox TxtImporte 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2100
            MaxLength       =   12
            TabIndex        =   3
            Text            =   "TxtImpo"
            Top             =   900
            Width           =   1020
         End
         Begin VB.CommandButton CmbActualizar 
            Caption         =   "&Actualizar"
            Height          =   375
            Left            =   3180
            TabIndex        =   4
            Top             =   900
            Width           =   915
         End
         Begin VB.Label LblTexto 
            Caption         =   "LblTexto"
            Height          =   225
            Left            =   2085
            TabIndex        =   9
            Top             =   645
            Width           =   1860
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Lista de Precio :"
         Height          =   255
         Left            =   300
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "VentaActuPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset, cRsl As ClsLectura, cPro As ClsPrecios
Dim Rs1 As ADODB.Recordset, nOp As Byte, sw As Boolean, nProv As Long

Private Sub CmbFamilias_Click()
  Dim Tip As String
  If sw = False Then
     Limpiar
     CrearRs
     If CmbFamilias.ListIndex <> -1 Then
        Set Rs1 = cPro.DameProductos(CmbLista.ItemData(CmbLista.ListIndex), "", "", CmbFamilias.ItemData(CmbFamilias.ListIndex), False, 0, CmbMarcas.ItemData(CmbMarcas.ListIndex))
     End If
     LLenarGrid
  End If
End Sub

Private Sub CmbMarcas_Click()
 Dim Tip As String
  If sw = False Then
     Limpiar
     CrearRs
     If CmbMarcas.ListIndex <> -1 Then
        Set Rs1 = cPro.DameProductos(CmbLista.ItemData(CmbLista.ListIndex), "", "", CmbFamilias.ItemData(CmbFamilias.ListIndex), False, 0, CmbMarcas.ItemData(CmbMarcas.ListIndex))
     End If
     LLenarGrid
  End If
End Sub

Private Sub CmbProveedor_Click()
  Dim Tip As String
  If sw = False Then
     Limpiar
     CrearRs
     If CmbProveedor.ListIndex <> -1 Then
        Set Rs1 = cPro.DameProductos(CmbLista.ItemData(CmbLista.ListIndex), "", "", , IIf(Check1.Value = 1, True, False), CmbProveedor.ItemData(CmbProveedor.ListIndex), 0)
     End If
     LLenarGrid
  End If
End Sub

Private Sub CmdCajas_Click()
   Dim cRsl As ClsLectura
   
   Set cRsl = New ClsLectura
   
   If cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=85 and Usuario=1") = True Then
      Dim cRx As ClsProductoL
      
      Set cRx = New ClsProductoL
  
      If MsgBox("Inicializa Los Precios Del Server a las Caja/s ?", 20, "Atención") = 6 Then
        cRx.AgegrarPreciosCajas
        MsgBox "Operación Realizada Con Exito", vbInformation, "Atención"
     End If
   End If
   
End Sub

Private Sub CmdTodos_Click()
  Dim Tip As String, cRsl As New ClsLectura

 Set cRsl = New ClsLectura
 Limpiar
 CrearRs
 Set Rs1 = cPro.DameProductos(CmbLista.ItemData(CmbLista.ListIndex), "", "", , IIf(Check1.Value = 1, True, False))
 LLenarGrid
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
     Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub CmbActualizar_Click()
  If Rs.RecordCount <> 0 Then
     Select Case nOp
            Case 1
                 ActualizoPrecio
            Case 2
                 ActualizoImporte
            Case 3
                 ActualizoPorcentaje
     End Select
     Grid1.SetFocus
  End If
End Sub
Private Sub ActualizoPrecio()
  Dim cRsl As ClsLectura, rLista As Recordset

  Set cRsl = New ClsLectura
  On Error GoTo errHandler

  Set rLista = cRsl.TraerTodos("ListaDePrecio", "*")

  Rs!precio = TxtImporte.text
'  If OptPrec(0).Value = True Then
'
'
'     cPro.ActulizarPrecios rLista, CmbLista.ItemData(CmbLista.ListIndex), Rs!codigo, TxtImporte.text
'  Else
'     nProv = cRsl.TraerValorDeUnCampo("Productos", "Proveedor", "Producto='" & Rs!codigo & "'")
'
'     cPro.ActulizarPreciosCosto rLista, Rs!codigo, TxtImporte, nProv, ChkActua.Value
'  End If
  Rs.Update
  Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub ActualizoImporte()
 On Error GoTo errHandler

 Rs.MoveFirst
 Do While Not Rs.EOF
    Rs!precio = Rs!precio + TxtImporte.text
    cPro.ActulizarImporte CmbLista.ItemData(CmbLista.ListIndex), Rs!codigo, TxtImporte.text
    Rs.Update
    Rs.MoveNext
 Loop
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub ActualizoPorcentaje()
   
 Dim pPrecio As Double
 
 On Error GoTo errHandler

 Rs.MoveFirst
 Do While Not Rs.EOF
    pPrecio = Rs!precio + (Rs!precio * TxtImporte.text / 100)
    Rs!precio = RedondearPrecios(pPrecio)
    cPro.ActulizarPorcentaje CmbLista.ItemData(CmbLista.ListIndex), Rs!codigo, Rs!precio
    Rs.Update
    Rs.MoveNext
 Loop
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub CmbLista_Click()
  Set cRsl = New ClsLectura
  CrearRs
  CmbCopiar.Enabled = True
  CmdCopiar.Enabled = True
  cRsl.CargaCombo CmbCopiar, "ListadePrecio", "ListaPrecio", "Descripcion", "ListaPrecio<>" & CmbLista.ItemData(CmbLista.ListIndex)
  If CmbCopiar.ListCount <> 0 Then
     CmbCopiar.ListIndex = 0
  Else
     CmdCopiar.Enabled = False
     CmbCopiar.Enabled = False
  End If
End Sub
Private Sub CmdCopiar_Click()
 On Error GoTo errHandler

  If MsgBox("Copia la Lista de Precio " & Trim(CmbLista.text) & " a " & Trim(CmbCopiar.text) & " ?", 20, "Atención") = 6 Then
     Me.MousePointer = 11
     cPro.BorrarPrecios (CmbCopiar.ItemData(CmbCopiar.ListIndex))
     cPro.CopiarPrecios Rs, CmbCopiar.ItemData(CmbCopiar.ListIndex)
     Me.MousePointer = 0
     MsgBox "Copia Realizada con Exito", vbCritical, "Atención"
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub Form_Load()
 On Error GoTo errHandler

 Set Rs1 = New ADODB.Recordset
 Set cRsl = New ClsLectura
 Set cPro = New ClsPrecios
 
 Limpiar
 CrearRs
 sw = True
 OptPorcentaje.Value = True
 CargarCombo
 sw = False
 
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub LlenarRs()
 On Error GoTo errHandler

 If Rs1.RecordCount <> 0 Then
    Do While Not Rs1.EOF
       If Not IsNull(Rs1!Descripcion) Then
          Rs.AddNew
          Rs!codigo = Rs1!Producto
          Rs!Producto = Rs1!Descripcion & " - " & Rs1!Descripcion1
          Rs!precio = Rs1!precio
          Rs!Porcentaje = Rs1!Porcentaje
          Rs!fecha = Rs1!fecha
       End If
       Rs1.MoveNext
    Loop
    TxtImporte.text = Format(0, nCantDecimales)
 End If
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub CargarCombo()
 On Error GoTo errHandler

 cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
 cRsl.CargaCombo CmbFamilias, "Familias", "Familia", "Descripcion", ""
 cRsl.CargaCombo CmbProveedor, "Proveedores", "Id", "RazonSocial", ""
 cRsl.CargaCombo CmbMarcas, "Marcas", "Marca", "Descripcion", ""


 CmbFamilias.AddItem "Ninguna", 0
 CmbMarcas.AddItem "Ninguna", 0
 
 CmbFamilias.text = "Ninguna"
 CmbMarcas.text = "Ninguna"
  
 If CmbLista.ListCount <> 0 Then
    CmbLista.ListIndex = 0
 Else
    MsgBox "Debe Tener al Menos una Lista de Precio", vbCritical, "Atención"
    Unload Me
 End If
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub CrearRs()
 On Error GoTo errHandler
 
 Set Rs = New ADODB.Recordset
 Rs.Fields.Append "Codigo", adVarChar, 25, adFldUpdatable
 Rs.Fields.Append "Producto", adVarChar, 200, adFldUpdatable
 Rs.Fields.Append "Precio", adCurrency, , adFldUpdatable
 Rs.Fields.Append "Fecha", adDate
 Rs.Fields.Append "Porcentaje", adCurrency, , adFldUpdatable
 Rs.Open
 
 Set Grid1.DataSource = Rs
 Grid1.HeadFont.Size = 10
 Grid1.HeadFont.bold = True
 With Grid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Producto"
      .Columns(1).Width = 5000
      .Columns(2).Caption = "Precio"
      .Columns(2).Width = 1000
      .Columns(2).Alignment = dbgRight
      .Columns(2).NumberFormat = nCantDecimales
      .Columns(3).Caption = "Fecha"
      .Columns(3).Width = 950
      .Columns(4).Visible = False
 End With
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
 Rs.Sort = Rs.Fields(ColIndex).Name & " Asc"
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler
  If OptImporte.Value = False And OptPorcentaje.Value = False Then
     If Grid1.Row <> -1 Then
        Rs.Bookmark = Grid1.Bookmark
        TxtImporte.text = Format(Rs!precio, nCantDecimales)
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub OptImporte_Click()
  nOp = 2
  LblTexto.Caption = "Importe :"
  TxtImporte.text = "0.00"
End Sub


Private Sub OptPorcentaje_Click()
  nOp = 3
  LblTexto.Caption = "Porcentaje :"
  TxtImporte.text = "0.00"
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
  If TxtImporte.text = "" Then TxtImporte.text = 0
  TxtImporte.text = Format(TxtImporte.text, nCantDecimales)
End Sub
Private Sub Limpiar()
  TxtImporte.text = "0.00"
End Sub

Private Sub LLenarGrid()
 On Error GoTo errHandler
 
 If Rs1.RecordCount <> 0 Then
    If sw = False Then
       CrearRs
       LlenarRs
       TxtImporte.text = Format(0, nCantDecimales)
    End If
 End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
