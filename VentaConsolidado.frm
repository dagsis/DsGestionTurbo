VERSION 5.00
Begin VB.Form VentaConsolidado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidado de Mercaderias"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7230
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   7020
      Begin VB.TextBox TxtSuc 
         Height          =   315
         Index           =   1
         Left            =   4665
         TabIndex        =   6
         Text            =   "0000"
         Top             =   1995
         Width           =   495
      End
      Begin VB.TextBox TxtNum 
         Height          =   315
         Index           =   1
         Left            =   5460
         TabIndex        =   7
         Text            =   "00000000"
         Top             =   1995
         Width           =   855
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Index           =   0
         Left            =   825
         MaxLength       =   25
         TabIndex        =   1
         Text            =   "TxtCodigo"
         Top             =   405
         Width           =   1470
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Index           =   0
         Left            =   6120
         TabIndex        =   0
         Top             =   405
         Width           =   765
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Index           =   1
         Left            =   825
         MaxLength       =   25
         TabIndex        =   2
         Text            =   "TxtCodigo"
         Top             =   960
         Width           =   1470
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Index           =   1
         Left            =   6120
         TabIndex        =   10
         Top             =   960
         Width           =   765
      End
      Begin VB.ComboBox CmbComprobante 
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1500
         Width           =   3450
      End
      Begin VB.TextBox TxtNum 
         Height          =   315
         Index           =   0
         Left            =   2310
         TabIndex        =   5
         Text            =   "00000000"
         Top             =   1980
         Width           =   855
      End
      Begin VB.TextBox TxtSuc 
         Height          =   315
         Index           =   0
         Left            =   1515
         TabIndex        =   4
         Text            =   "0000"
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "--"
         Height          =   195
         Left            =   5310
         TabIndex        =   19
         Top             =   2055
         Width           =   150
      End
      Begin VB.Label Label11 
         Caption         =   "Desde :"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   165
         TabIndex        =   17
         Top             =   915
         Width           =   600
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescripcion"
         Height          =   285
         Index           =   0
         Left            =   2370
         TabIndex        =   16
         Top             =   405
         Width           =   3705
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescripcion"
         Height          =   285
         Index           =   1
         Left            =   2370
         TabIndex        =   15
         Top             =   960
         Width           =   3705
      End
      Begin VB.Label Label9 
         Caption         =   "Desde Nùmero :"
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   1980
         Width           =   1200
      End
      Begin VB.Label Label8 
         Caption         =   "Hasta Nùmero :"
         Height          =   255
         Left            =   3405
         TabIndex        =   13
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label Label7 
         Caption         =   "Comprobante :"
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label Label6 
         Caption         =   "--"
         Height          =   195
         Left            =   2100
         TabIndex        =   11
         Top             =   2040
         Width           =   150
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   5850
      TabIndex        =   9
      Top             =   2715
      Width           =   1245
   End
End
Attribute VB_Name = "VentaConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura, cPro As ClsPrecios, Rs As ADODB.Recordset
Dim cDes As ClsProductoL

Private Sub CargarCombos()
  Dim i As Integer, cRsl As ClsLectura
  
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set Rs = cPro.DameProductos(0, "", "")
  Rs.MoveFirst
  For i = 0 To 1
      TxtCodigo(i).text = Rs!Producto
      LblDescripcion(i) = cDes.BuscarNombreProducto(Rs!Producto)
      Rs.MoveLast
  Next i
  
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=1 and Suc=" & nSucursal
  CmbComprobante.ListIndex = 0
  
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmdAplicar_Click()
  nImpr = 43
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
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub TxtNum_GotFocus(Index As Integer)
  TxtNum(Index).SelStart = 0
  TxtNum(Index).SelLength = Len(TxtNum(Index).text)
End Sub

Private Sub TxtNum_KeyPress(Index As Integer, KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtNum_LostFocus(Index As Integer)
  If TxtNum(Index).text = "" Then TxtNum(Index).text = 0
  TxtNum(Index).text = Format(TxtNum(Index).text, "00000000")
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
  TxtSuc(0).text = "0001"
  TxtNum(0).text = "00000000"
  TxtSuc(1).text = "0001"
  TxtNum(1).text = "00000000"
  LblDescripcion(0).Caption = ""
  LblDescripcion(1).Caption = ""
End Sub

Private Sub Form_Load()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  Set cPro = New ClsPrecios
  Set cDes = New ClsProductoL
  Set Rs = New ADODB.Recordset
  
  Limpiar
  CargarCombos
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
  End If
End Sub

Private Sub TodoBien(Index As Integer)
  CmdAplicar.Enabled = False
  If Len(TxtCodigo(Index)) <> 0 Then
     CmdAplicar.Enabled = True
  End If
End Sub




