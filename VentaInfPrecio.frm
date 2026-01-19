VERSION 5.00
Begin VB.Form VentaInfPrecio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Precios"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   7635
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   390
      Left            =   6240
      TabIndex        =   8
      Top             =   4035
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Left            =   135
      TabIndex        =   1
      Top             =   150
      Width           =   7365
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripción"
         Height          =   210
         Left            =   4875
         TabIndex        =   3
         Top             =   435
         Width           =   1245
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   3600
         TabIndex        =   2
         Top             =   435
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
         Height          =   2880
         Left            =   150
         TabIndex        =   10
         Top             =   765
         Width           =   7080
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1440
            Left            =   3720
            TabIndex        =   19
            Top             =   1335
            Width           =   3195
            Begin VB.ComboBox CmbHasta 
               Height          =   315
               Left            =   105
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1020
               Width           =   2985
            End
            Begin VB.ComboBox CmbDesde 
               Height          =   315
               Left            =   105
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   435
               Width           =   2985
            End
            Begin VB.Label Label5 
               Caption         =   "Hasta :"
               Height          =   195
               Left            =   135
               TabIndex        =   22
               Top             =   795
               Width           =   585
            End
            Begin VB.Label Label4 
               Caption         =   "Desde :"
               Height          =   195
               Left            =   105
               TabIndex        =   20
               Top             =   210
               Width           =   675
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Gupos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   150
            TabIndex        =   15
            Top             =   1335
            Width           =   3540
            Begin VB.OptionButton OptFamilias 
               Caption         =   "Familias"
               Height          =   270
               Left            =   135
               TabIndex        =   18
               Top             =   315
               Width           =   1035
            End
            Begin VB.OptionButton OptNinguno 
               Caption         =   "Ninguno"
               Height          =   270
               Left            =   2355
               TabIndex        =   17
               Top             =   315
               Width           =   1065
            End
            Begin VB.OptionButton OptMarcas 
               Caption         =   "Marca"
               Height          =   270
               Left            =   1245
               TabIndex        =   16
               Top             =   315
               Width           =   885
            End
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   270
            Index           =   1
            Left            =   6120
            TabIndex        =   7
            Top             =   960
            Width           =   765
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Index           =   1
            Left            =   825
            MaxLength       =   25
            TabIndex        =   6
            Text            =   "TxtCodigo"
            Top             =   960
            Width           =   1470
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   270
            Index           =   0
            Left            =   6120
            TabIndex        =   5
            Top             =   405
            Width           =   765
         End
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Index           =   0
            Left            =   825
            MaxLength       =   25
            TabIndex        =   4
            Text            =   "TxtCodigo"
            Top             =   405
            Width           =   1470
         End
         Begin VB.Label LblDescripcion 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescripcion"
            Height          =   285
            Index           =   1
            Left            =   2370
            TabIndex        =   14
            Top             =   960
            Width           =   3705
         End
         Begin VB.Label LblDescripcion 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescripcion"
            Height          =   285
            Index           =   0
            Left            =   2370
            TabIndex        =   13
            Top             =   405
            Width           =   3705
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta :"
            Height          =   240
            Left            =   165
            TabIndex        =   12
            Top             =   915
            Width           =   600
         End
         Begin VB.Label Label2 
            Caption         =   "Desde :"
            Height          =   195
            Left            =   165
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.ComboBox CmbLista 
         Height          =   315
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   375
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   "Lista :"
         Height          =   195
         Left            =   420
         TabIndex        =   9
         Top             =   345
         Width           =   585
      End
   End
End
Attribute VB_Name = "VentaInfPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura, cPro As ClsPrecios, Rs As ADODB.Recordset
Dim cDes As ClsProductoL


Private Sub CmbLista_Click()
  Dim i As Integer
  On Error GoTo errHandler
  
  Set Rs = cPro.DameProductos(CmbLista.ItemData(CmbLista.ListIndex), "", "")
  If Rs.RecordCount <> 0 Then
     Rs.MoveFirst
     For i = 0 To 1
         TxtCodigo(i).Text = Rs!Producto
         LblDescripcion(i) = cDes.BuscarNombreProducto(Rs!Producto)
         Rs.MoveLast
     Next i
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmdAplicar_Click()
  nImpr = 3
  If OptCodigo.Value = False Then
     nImpr = 4
  End If
  FrmImpresor.Show
End Sub

Private Sub Exportar()
'  Dim ab As String
'  Set ex = New DatabaseExport                 'open database file
'  Set ex.ProgressBar = ProgressBar1
'
'  Set rs = cPro.DameProductos(CmbLista.ItemData(CmbLista.ListIndex), "", "")
'  Set ex.ADORecordset = rs
'
'  If ex.ShowSave(CommonDialog1, 2) Then
'      ab = ex.FilePath
'  End If
'  ex.ExportToExcel True, True
  
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  
  nLista = CmbLista.ItemData(CmbLista.ListIndex)

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
         ctl.Text = ""
      End If
  Next ctl
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
  OptDescripcion.Value = True
  OptNinguno.Value = True
  CargarCombos
  Habilitar False, "Ninguno"
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CargarCombos()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  
  cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
  If CmbLista.ListCount <> 0 Then
     CmbLista.ListIndex = 0
  Else
     MsgBox "No hay lista de Precios Cargados", vbCritical, "Atención"
     Unload Me
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub OptMarcas_Click()
  Set cRsl = New ClsLectura
  Habilitar True, "Marcas"
  cRsl.CargaCombo CmbDesde, "Marcas", "Marca", "Descripcion", ""
  cRsl.CargaCombo CmbHasta, "Marcas", "Marca", "Descripcion", ""
  CmbDesde.ListIndex = 0
  CmbHasta.ListIndex = CmbHasta.ListCount - 1
  Set cRsl = Nothing
End Sub

Private Sub OptNinguno_Click()
  Habilitar False, "Ninguno"
  CmbDesde.Clear
  CmbHasta.Clear
End Sub

Private Sub OptFamilias_Click()
  Set cRsl = New ClsLectura
  Habilitar True, "Familias"
  cRsl.CargaCombo CmbDesde, "Familias", "Familia", "Descripcion", ""
  cRsl.CargaCombo CmbHasta, "Familias", "Familia", "Descripcion", ""
  CmbDesde.ListIndex = 0
  CmbHasta.ListIndex = CmbHasta.ListCount - 1
  Set cRsl = Nothing
End Sub

Private Sub TxtCodigo_Change(Index As Integer)
  TodoBien (Index)
End Sub

Private Sub TxtCodigo_GotFocus(Index As Integer)
  TxtCodigo(Index).SelStart = 0
  TxtCodigo(Index).SelLength = Len(TxtCodigo(Index))
End Sub

Private Sub TxtCodigo_LostFocus(Index As Integer)
  LblDescripcion(Index).Caption = cDes.BuscarNombreProducto(TxtCodigo(Index).Text)
  If LblDescripcion(Index).Caption = "" Then
     TxtCodigo(Index).Text = ""
  End If
End Sub

Private Sub TodoBien(Index As Integer)
  CmdAplicar.Enabled = False
  If Len(TxtCodigo(Index)) <> 0 Then
     CmdAplicar.Enabled = True
  End If
End Sub

Private Sub Habilitar(bHabi As Boolean, bCapt As String)
 Frame4.Caption = bCapt
 Frame4.Enabled = bHabi
 CmbDesde.Enabled = bHabi
 CmbHasta.Enabled = bHabi
End Sub
