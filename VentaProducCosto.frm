VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaProducCosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Venta Productos con Costos"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   7380
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   5955
      TabIndex        =   9
      Top             =   2910
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   7080
      Begin VB.CheckBox Check1 
         Caption         =   "Alfabetico"
         Height          =   240
         Left            =   300
         TabIndex        =   14
         Top             =   2085
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   315
         Left            =   3180
         TabIndex        =   13
         Top             =   1485
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         Format          =   708771841
         CurrentDate     =   37156
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   885
         TabIndex        =   12
         Top             =   1485
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         Format          =   708771841
         CurrentDate     =   37156
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Index           =   1
         Left            =   6120
         TabIndex        =   4
         Top             =   960
         Width           =   765
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Index           =   1
         Left            =   825
         MaxLength       =   25
         TabIndex        =   3
         Text            =   "TxtCodigo"
         Top             =   960
         Width           =   1470
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Index           =   0
         Left            =   6120
         TabIndex        =   2
         Top             =   405
         Width           =   765
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Index           =   0
         Left            =   855
         MaxLength       =   25
         TabIndex        =   1
         Text            =   "TxtCodigo"
         Top             =   405
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   2460
         TabIndex        =   11
         Top             =   1485
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   225
         Left            =   180
         TabIndex        =   10
         Top             =   1485
         Width           =   675
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescripcion"
         Height          =   285
         Index           =   1
         Left            =   2370
         TabIndex        =   8
         Top             =   960
         Width           =   3705
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescripcion"
         Height          =   285
         Index           =   0
         Left            =   2370
         TabIndex        =   7
         Top             =   405
         Width           =   3705
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   165
         TabIndex        =   6
         Top             =   915
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Desde :"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "VentaProducCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura, cPro As ClsPrecios, Rs As ADODB.Recordset
Dim cDes As ClsProductoL

Private Sub CargarCombos()
  Dim i As Integer
  On Error GoTo errHandler
  
  Set Rs = cPro.DameProductos(0, "", "")
  If Rs.RecordCount <> 0 Then
    Rs.MoveFirst
    For i = 0 To 1
        TxtCodigo(i).text = Rs!Producto
        LblDescripcion(i) = cDes.BuscarNombreProducto(Rs!Producto)
        Rs.MoveLast
    Next i
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmdAplicar_Click()
  If Len(LblDescripcion(0).Caption) <> 0 And Len(LblDescripcion(1).Caption) <> 0 Then
     nImpr = 22
     FrmImpresor.Show
  End If
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
  DtpDesde.Value = Date
  DtpHasta.Value = Date
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


