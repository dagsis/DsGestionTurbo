VERSION 5.00
Begin VB.Form StockInfSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos de Stock"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   7935
   Begin VB.CheckBox ChkSaldo 
      Caption         =   "Mostrar Solos Con Stock"
      Height          =   225
      Left            =   270
      TabIndex        =   16
      Top             =   3600
      Width           =   2355
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   6540
      TabIndex        =   6
      Top             =   3555
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saldos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   225
      TabIndex        =   0
      Top             =   105
      Width           =   7515
      Begin VB.ComboBox CmbFamilia 
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   825
         Width           =   2835
      End
      Begin VB.OptionButton OptTrams 
         Caption         =   "Transito"
         Height          =   210
         Left            =   6210
         TabIndex        =   17
         Top             =   360
         Width           =   1245
      End
      Begin VB.ComboBox CmbDeposito 
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   2835
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripción"
         Height          =   210
         Left            =   5010
         TabIndex        =   12
         Top             =   360
         Width           =   1245
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   1050
      End
      Begin VB.Frame Frame2 
         Caption         =   "Articulos"
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
         Left            =   300
         TabIndex        =   1
         Top             =   1290
         Width           =   7080
         Begin VB.CheckBox Check1 
            Caption         =   "Totales en los Depositos"
            Height          =   240
            Left            =   90
            TabIndex        =   15
            Top             =   1545
            Width           =   2250
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Index           =   0
            Left            =   855
            MaxLength       =   25
            TabIndex        =   4
            Text            =   "TxtCodigo"
            Top             =   540
            Width           =   1470
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   270
            Index           =   0
            Left            =   6150
            TabIndex        =   3
            Top             =   540
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
            Index           =   1
            Left            =   6150
            TabIndex        =   2
            Top             =   1095
            Width           =   765
         End
         Begin VB.Label Label2 
            Caption         =   "Desde :"
            Height          =   195
            Left            =   195
            TabIndex        =   10
            Top             =   495
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta :"
            Height          =   240
            Left            =   195
            TabIndex        =   9
            Top             =   1050
            Width           =   600
         End
         Begin VB.Label LblDescripcion 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescripcion"
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   8
            Top             =   540
            Width           =   3705
         End
         Begin VB.Label LblDescripcion 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescripcion"
            Height          =   285
            Index           =   1
            Left            =   2400
            TabIndex        =   7
            Top             =   1095
            Width           =   3705
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Familia :"
         Height          =   240
         Left            =   330
         TabIndex        =   19
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Depósito :"
         Height          =   240
         Left            =   330
         TabIndex        =   13
         Top             =   345
         Width           =   765
      End
   End
End
Attribute VB_Name = "StockInfSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura, cPro As ClsProductoL, Rs As ADODB.Recordset
Dim cDes As ClsProductoL

Private Sub CmbDeposito_Click()
  Dim i As Integer
  On Error GoTo errHandler
  
  Set Rs = cPro.DameProductos(CmbDeposito.ItemData(CmbDeposito.ListIndex), "", "", 1)
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
  nImpr = 5
  If OptCodigo.Value = False And OptDescripcion.Value = True Then
     nImpr = 6
  End If
  If OptTrams.Value = True Then
      nImpr = 68
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
  On Error GoTo errHandler
    
  cRsl.CargaCombo CmbDeposito, "Depositos", "Deposito", "Descripcion", ""
  If CmbDeposito.ListCount <> 0 Then
     CmbDeposito.text = cRsl.TraerValorDeUnCampo("Depositos", "Descripcion", "Deposito=" & cRsl.RegPorDefecto("Deposito"))
     CmbDeposito.AddItem "TODOS"
     CmbDeposito.ItemData(CmbDeposito.NewIndex) = 0
  Else
     MsgBox "No hay Depósitos Cargados", vbCritical, "Atención"
     Unload Me
  End If
  
    cRsl.CargaCombo CmbFamilia, "Familias", "Familia", "Descripcion", ""
  CmbFamilia.AddItem "Todos", 0
  
  CmbFamilia.text = "Todos"
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


