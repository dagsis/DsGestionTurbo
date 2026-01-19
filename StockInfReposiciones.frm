VERSION 5.00
Begin VB.Form StockInfReposiciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reposiciones de Productos"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   7665
   Begin VB.Frame Frame1 
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
      Height          =   3015
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   7515
      Begin VB.ComboBox CmbProveedor 
         Height          =   315
         Index           =   0
         Left            =   1845
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   390
         Width           =   3990
      End
      Begin VB.Frame Frame2 
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
         Height          =   1875
         Left            =   165
         TabIndex        =   2
         Top             =   945
         Width           =   7080
         Begin VB.OptionButton OptDescripcion 
            Caption         =   "Descripción"
            Height          =   210
            Left            =   1440
            TabIndex        =   13
            Top             =   1575
            Width           =   1245
         End
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Codigo"
            Height          =   195
            Left            =   165
            TabIndex        =   12
            Top             =   1575
            Width           =   1050
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   270
            Index           =   1
            Left            =   6150
            TabIndex        =   6
            Top             =   1095
            Width           =   765
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Index           =   1
            Left            =   855
            MaxLength       =   25
            TabIndex        =   5
            Text            =   "TxtCodigo"
            Top             =   1095
            Width           =   1470
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   270
            Index           =   0
            Left            =   6150
            TabIndex        =   4
            Top             =   540
            Width           =   765
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Index           =   0
            Left            =   855
            MaxLength       =   25
            TabIndex        =   3
            Text            =   "TxtCodigo"
            Top             =   540
            Width           =   1470
         End
         Begin VB.Label LblDescripcion 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescripcion"
            Height          =   285
            Index           =   1
            Left            =   2400
            TabIndex        =   10
            Top             =   1095
            Width           =   3705
         End
         Begin VB.Label LblDescripcion 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescripcion"
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   9
            Top             =   540
            Width           =   3705
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta :"
            Height          =   240
            Left            =   195
            TabIndex        =   8
            Top             =   1050
            Width           =   600
         End
         Begin VB.Label Label2 
            Caption         =   "Desde :"
            Height          =   195
            Left            =   195
            TabIndex        =   7
            Top             =   495
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         Caption         =   " Proveedor :"
         Height          =   195
         Left            =   330
         TabIndex        =   11
         Top             =   390
         Width           =   1470
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   6315
      TabIndex        =   0
      Top             =   3300
      Width           =   1230
   End
End
Attribute VB_Name = "StockInfReposiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura, cPro As ClsProductoL, Rs As ADODB.Recordset
Dim cDes As ClsProductoL


Private Sub CmbProveedor_Click(Index As Integer)
  Dim i As Integer
  On Error GoTo errHandler
  Limpiar
  If CmbProveedor(0).text <> "" Then
     Set Rs = cPro.DameProductoProvee(CmbProveedor(0).ItemData(CmbProveedor(0).ListIndex))
     If Rs.RecordCount <> 0 Then
        Rs.MoveFirst
        For i = 0 To 1
            TxtCodigo(i).text = Rs!Producto
            LblDescripcion(i) = cDes.BuscarNombreProducto(Rs!Producto)
            Rs.MoveLast
        Next i
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores
End Sub

Private Sub CmdAplicar_Click()
  nImpr = 7
  If OptCodigo.Value = False Then
     nImpr = 8
  End If
  FrmImpresor.Show
End Sub
Private Sub CmdBuscar_Click(Index As Integer)
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
     TxtCodigo(Index) = RsP!Producto
     LblDescripcion(Index).Caption = RsP!Descripcion
     If Index = 0 Then
        TxtCodigo(1).text = RsP!Producto
        LblDescripcion(1).Caption = RsP!Descripcion
     End If
     RsP.Close
  End If
  Set cRsl = Nothing
 
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
  '   KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  LblDescripcion(0).Caption = ""
  LblDescripcion(1).Caption = ""
End Sub

Private Sub Form_Load()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  Set cPro = New ClsProductoL
  Set cDes = New ClsProductoL
  Set Rs = New ADODB.Recordset
  
  Limpiar
  OptDescripcion.Value = True
  CargarCombos
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CargarCombos()
  Dim i As Integer
  On Error GoTo errHandler
  
  cRsl.CargaCombo CmbProveedor(0), "Proveedores", "Id", "RazonSocial", ""
    
  If CmbProveedor(0).ListCount <> 0 Then
     CmbProveedor(0).AddItem "TODOS"
'     CmbProveedor(0).ItemData(CmbProveedor(0).NewIndex) = 0
     CmbProveedor(0).text = "TODOS"
  Else
     MsgBox "No hay Proveedores Cargados", vbCritical, "Atención"
     Unload Me
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub TxtCodigo_Change(Index As Integer)
  TodoBien (Index)
End Sub

Private Sub TxtCodigo_GotFocus(Index As Integer)
  TxtCodigo(Index).SelStart = 0
  TxtCodigo(Index).SelLength = Len(TxtCodigo(Index))
End Sub

Private Sub TxtCodigo_LostFocus(Index As Integer)
  LblDescripcion(Index).Caption = cDes.BuscarNombreProducto(TxtCodigo(Index).text)
  
  If LblDescripcion(Index).Caption = "" Then
     TxtCodigo(Index).text = ""
  Else
     TxtCodigo(1).text = TxtCodigo(Index).text
     LblDescripcion(1).Caption = LblDescripcion(Index).Caption
  End If
End Sub

Private Sub TodoBien(Index As Integer)
  CmdAplicar.Enabled = False
  If Len(TxtCodigo(Index)) <> 0 Then
     CmdAplicar.Enabled = True
  End If
End Sub


