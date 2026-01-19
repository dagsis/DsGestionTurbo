VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form MtrProductoProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Productos Proveedores"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   8745
   Begin VB.CommandButton CmdModificar 
      Caption         =   "Modificar"
      Height          =   315
      Left            =   2865
      TabIndex        =   13
      Top             =   1665
      Width           =   1020
   End
   Begin VB.CommandButton CmdProveedor 
      Caption         =   "Buscar"
      Height          =   315
      Left            =   5805
      TabIndex        =   1
      Top             =   330
      Width           =   900
   End
   Begin VB.TextBox TxtProveedor 
      Height          =   300
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "TxtProveedor"
      Top             =   330
      Width           =   1140
   End
   Begin VB.CommandButton CmdSarlir 
      Caption         =   "Terminar"
      Height          =   390
      Left            =   7425
      TabIndex        =   11
      Top             =   5865
      Width           =   1155
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   3660
      Left            =   120
      TabIndex        =   10
      Top             =   2115
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   6456
      _Version        =   393216
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
            LCID            =   11274
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
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtProducto 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Text            =   "TxtProducto"
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton CmdProducto 
      Caption         =   "Buscar"
      Height          =   315
      Left            =   5805
      TabIndex        =   3
      Top             =   960
      Width           =   900
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Grabar"
      Height          =   315
      Left            =   1725
      TabIndex        =   5
      Top             =   1665
      Width           =   1020
   End
   Begin VB.TextBox TxtProduProvee 
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Text            =   "TxtProduProvee"
      Top             =   1650
      Width           =   1515
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
      Left            =   1350
      TabIndex        =   12
      Top             =   330
      Width           =   4380
   End
   Begin VB.Label Label4 
      Caption         =   "Codigo Producto :"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   705
      Width           =   1500
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
      Height          =   300
      Left            =   1350
      TabIndex        =   8
      Top             =   960
      Width           =   4380
   End
   Begin VB.Label Label2 
      Caption         =   "Producto Proveedor :"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   1350
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor :"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   75
      Width           =   1845
   End
End
Attribute VB_Name = "MtrProductoProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura, sw As Boolean
Dim Rsd As ADODB.Recordset

Private Sub CmdAplicar_Click()
  Dim cPro As ClsProductoE, cProb As ClsProductoL
  
  Set cPro = New ClsProductoE
  Set cProb = New ClsProductoL
  If cProb.EstaElProdProv(TxtProveedor.text, TxtProducto.text) = False Then
     If cProb.EstaElCodigoProv(TxtProveedor.text, TxtProduProvee.text) = False Then
        cPro.GrabarProductoProveedor TxtProveedor.text, TxtProducto.text, TxtProduProvee.text
        sw = True
        Rsd.AddNew
        GrabarRsd
        sw = False
        TxtProduProvee.text = ""
        TxtProducto.SetFocus
     End If
 End If
End Sub

Private Sub CmdProducto_Click()
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
     TxtProducto.text = RsP!Producto
     LblProducto.Caption = RsP!Descripcion
     TxtProduProvee.text = ""
     TxtProduProvee.SetFocus
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CmdProveedor_Click()
  Dim RsC As ADODB.Recordset, cRiL As ClsClienteL, cRcL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRiL = New ClsClienteL
  Set cRcL = New ClsClienteL
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 2
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ProveedorTraerUno", nDat, True)
     TxtProveedor.text = RsC!Proveedor
     LblProveedor.Caption = RsC!RazonSocial
     TxtProducto.SetFocus
     RsC.Close
     sw = False
     CrearRs
     LlenarGrilla
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CmdSarlir_Click()
  Unload Me
End Sub

Private Sub CmdModificar_Click()
  Dim cPro As ClsProductoE
  Set cPro = New ClsProductoE
  cPro.ModificarProductoProveedor TxtProveedor.text, TxtProducto.text, TxtProduProvee.text
  CrearRs
  LlenarGrilla
  TxtProduProvee.text = ""
  TxtProducto.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
     End If
     KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  LblProducto.Caption = ""
  LblProveedor.Caption = ""
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
End Sub
Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  Limpiar
  CrearRs
  CmdAplicar.Enabled = False
End Sub
Private Sub CrearRs()
  Set Rsd = New ADODB.Recordset
  Rsd.Fields.Append "Producto", adVarChar, 15
  Rsd.Fields.Append "Proveedor", adVarChar, 15
  Rsd.Fields.Append "Descripcion", adVarChar, 50
  Rsd.Open
  CabGrid
End Sub

Private Sub CabGrid()
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.size = 10
  Grid1.HeadFont.bold = True
 With Grid1
      .Columns(0).Width = 1400
      .Columns(0).Caption = "C.Producto"
      .Columns(1).Width = 1400
      .Columns(1).Caption = "C.Proveedor"
      .Columns(2).Width = 4850
      .Columns(2).Caption = "Descripción"
 End With
End Sub
Private Sub LlenarGrilla()
  Dim cPr As ClsProductoL, rPr As ADODB.Recordset
  Set cPr = New ClsProductoL
  
  Set rPr = cPr.TraerProductoProveedor(TxtProveedor.text)
  Do While Not rPr.EOF
     Rsd.AddNew
     Rsd!Producto = rPr!Producto
     Rsd!Proveedor = rPr!CodProvee
     Rsd!Descripcion = rPr!Descripcion
     Rsd.Update
     rPr.MoveNext
  Loop
  CabGrid
End Sub

Private Sub Grid1_Click()
  sw = True
End Sub
Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler
     If sw = False Then
        If Grid1.Row <> -1 Then
           Rsd.Bookmark = Grid1.Bookmark
           LinkearTexto
        End If
     End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub LinkearTexto()
  Dim cPro As ClsProductoL
  
  Set cPro = New ClsProductoL
  
  TxtProducto.text = Rsd!Producto
  TxtProduProvee.text = Rsd!Proveedor
  LblProducto.Caption = cPro.BuscarNombreProducto(TxtProducto.text)
End Sub
Private Sub TxtProducto_GotFocus()
  TxtProducto.SelStart = 0
  TxtProducto.SelLength = Len(TxtProducto.text)
End Sub
Private Sub TxtProduProvee_Change()
  CmdAplicar.Enabled = False
  CmdModificar.Enabled = False
  If LblProducto.Caption <> "" And Len(TxtProduProvee) <> 0 Then
     CmdAplicar.Enabled = True
     CmdModificar.Enabled = True
  End If
End Sub

Private Sub TxtProduProvee_GotFocus()
  TxtProduProvee.SelStart = 0
  TxtProduProvee.SelLength = Len(TxtProduProvee.text)
End Sub
Private Sub TxtProduProvee_LostFocus()
  Dim cPro As ClsProductoL
  Set cPro = New ClsProductoL
  If TxtProduProvee.text <> "" Then
  End If
End Sub
Private Sub TxtProveedor_GotFocus()
  TxtProveedor.SelStart = 0
  TxtProveedor.SelLength = Len(TxtProveedor.text)
End Sub
Private Sub GrabarRsd()
  Dim cPro As ClsProductoL
  
  Set cPro = New ClsProductoL
  
  Rsd!Producto = TxtProducto.text
  Rsd!Proveedor = TxtProduProvee.text
  Rsd!Descripcion = cPro.BuscarNombreProducto(TxtProducto.text)
  Rsd.Update
  CabGrid
End Sub
