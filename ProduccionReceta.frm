VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ProduccionReceta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Recetas"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   10965
   Begin VB.TextBox TxtObservacion 
      Height          =   525
      Left            =   1170
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "ProduccionReceta.frx":0000
      Top             =   555
      Width           =   6630
   End
   Begin VB.CommandButton CmdProducto 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   7905
      TabIndex        =   25
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TxtProducto 
      Height          =   285
      Left            =   945
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "TxtProducto"
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      Height          =   690
      Index           =   5
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4350
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   3
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1620
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   2
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   855
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Anular"
      Height          =   690
      Index           =   1
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   855
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      Height          =   690
      Index           =   0
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   105
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3870
      Left            =   60
      TabIndex        =   13
      Top             =   1185
      Width           =   8580
      Begin VB.TextBox TxtCosto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "TxtCosto"
         Top             =   3435
         Width           =   930
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         Height          =   645
         Index           =   1
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3075
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         Height          =   645
         Index           =   2
         Left            =   3765
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3075
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         Height          =   645
         Index           =   3
         Left            =   5250
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3075
         Width           =   870
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   285
         Left            =   1875
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "TxtDetalle"
         Top             =   2700
         Width           =   4395
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         Height          =   645
         Index           =   0
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3075
         Width           =   870
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "TxtCa"
         Top             =   3075
         Width           =   930
      End
      Begin VB.CommandButton CmdMateria 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   6330
         TabIndex        =   14
         Top             =   2700
         Width           =   705
      End
      Begin VB.TextBox TxtMateria 
         Height          =   285
         Left            =   675
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "TxtProducto"
         Top             =   2700
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2280
         Left            =   165
         TabIndex        =   17
         Top             =   255
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   4022
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
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   7500
         TabIndex        =   31
         Top             =   3510
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Total Costo :"
         Height          =   225
         Left            =   6495
         TabIndex        =   30
         Top             =   3510
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "P.Costo :"
         Height          =   255
         Left            =   75
         TabIndex        =   29
         Top             =   3435
         Width           =   780
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   3075
         Width           =   780
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         Height          =   210
         Left            =   90
         TabIndex        =   18
         Top             =   2700
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      Height          =   690
      Index           =   4
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2370
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   690
      Index           =   6
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   7
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   8
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2370
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Modificar"
      Height          =   690
      Index           =   9
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1620
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Obsercación :"
      Height          =   255
      Left            =   135
      TabIndex        =   28
      Top             =   525
      Width           =   1065
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
      Left            =   2010
      TabIndex        =   27
      Top             =   120
      Width           =   5790
   End
   Begin VB.Label Label4 
      Caption         =   "Producto :"
      Height          =   255
      Left            =   135
      TabIndex        =   26
      Top             =   150
      Width           =   900
   End
End
Attribute VB_Name = "ProduccionReceta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean                     ' Suich Cabecera Stock
Dim sw1 As Boolean                    ' Suich Detalles Stock
Dim sw2 As Boolean                    ' Suich Valores
Dim bProveedor As Boolean
Dim cImpor As Currency                ' Importe Total de Stock
Dim Rs_Sto As ADODB.Recordset         ' Cabecera Stock Tabla Uno
Dim RsCv As ADODB.Recordset           ' Cabecera Virtual Stock
Dim Rsd As ADODB.Recordset            ' Cabecera Vitual StockDetalle
Dim Rs As ADODB.Recordset             ' Cabecera Stock Tabla Todos
Dim Rs_Sis As ADODB.Recordset         ' Tabla Sistema, Ultima Stock
Dim nTip As Integer                   ' Valor de tipo de Movimiento
Public nMovi As Long
Public nCas As Integer
Public rImp As ADODB.Recordset
Dim bRemi As Boolean
Dim rRem As ADODB.Recordset
Dim nReg As Variant
Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Anular
         Case 2
              Anterior
         Case 3
              Siguiente
         Case 4
              Buscar
         Case 5
              Salir
         Case 6
              Imprimir
         Case 7
              Primero
         Case 8
              Ultimo
         Case 9
              Modificar
  End Select
End Sub
Private Sub CmdProducto_Click()
  Dim cRsl As ClsLectura
  
  Dim RsC As ADODB.Recordset
  On Error GoTo errHandler
  
  Set RsC = New ADODB.Recordset
  Set cRsl = New ClsLectura
  
  nBuscar = 3
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtProducto.Text = RsC!Producto
     LblProducto.Caption = RsC!Descripcion
     CmdDetalle(0).Enabled = True
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CmdMateria_Click()
  Dim RsP As ADODB.Recordset, cRsl As ClsLectura
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nBuscar = 3
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     TxtMateria.Text = RsP!Producto
     TxtDetalle.Text = RsP!Descripcion
     TxtCantidad.SetFocus
     RsP.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub DtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
        Exit Sub
     End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
     End If
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.Text = ""
      End If
  Next ctl
  LblTotal.Caption = "0.00"
  LblProducto.Caption = ""
  TxtCantidad.Text = "0.000"
  TxtCosto.Text = "0.000"
End Sub
Private Sub HabilitarTodo(bEstado As Boolean, bGrid As Boolean)
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
         Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
         Or TypeOf ctl Is ComboBox Then
         ctl.Enabled = bEstado
      End If
  Next ctl
  Grid1.Enabled = bGrid
End Sub
Private Sub LinkearTexto()
 Dim cRsl As ClsLectura
 
 Set cRsl = New ClsLectura
 
 TxtProducto.Text = Rs!Producto
 LblProducto.Caption = cRsl.TraerValorDeUnCampo("Productos", "Descripcion", "Producto='" & TxtProducto.Text & "'")
 TxtObservacion.Text = Rs!Observacion
 
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub CmdDetalle_Click(Index As Integer)
  Select Case Index
         Case 0
              NuevoRsD
         Case 1
              BorrarRsD
         Case 2
              ModificarRsD
         Case 3
              CancelarRsd
  End Select
End Sub
Private Sub Form_Load()
 On Error GoTo errHandler

 Set Rs = New ADODB.Recordset
 Set Rs_Sis = New ADODB.Recordset
 
 sw = True
 sw1 = True
 sw2 = True
 nCas = 0
 nLlama = 2
 
 Botones True, False, False, False, False, True, False
 BotonDetalles False, False, False, False
 BotonBuscar False, False

 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(4).Picture = LoadResPicture("Buscar", 0)
 CmdBotones(5).Picture = LoadResPicture("Salir", 0)
 CmdBotones(6).Picture = LoadResPicture("Imprimir", 0)
 CmdBotones(9).Picture = LoadResPicture("Modificar", 0)
 
 
 CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
 CmdDetalle(1).Picture = LoadResPicture("Borrar", 0)
 CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
 CmdDetalle(3).Picture = LoadResPicture("Cancelar", 0)
 
 CrearRsDetalles
 Limpiar
 HabilitarTodo False, False
 HabilitarDetalles False
 ArmaRSparaMovimiento
 Me.Top = 0
 Me.Left = 0
 
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bMod As Boolean, bBus As Boolean, _
  bVal As Boolean, bSal As Boolean, bModi As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bBorr
  CmdBotones(2).Enabled = bBus
  CmdBotones(3).Enabled = bBus
  CmdBotones(4).Enabled = bVal
  CmdBotones(5).Enabled = bSal
  CmdBotones(6).Enabled = bBus
  CmdBotones(7).Enabled = bBus
  CmdBotones(8).Enabled = bBus
  CmdBotones(9).Enabled = bModi
End Sub
Private Sub BotonDetalles(bAgr As Boolean, bBorr As Boolean, bMod As Boolean, bCan As Boolean)
  CmdDetalle(0).Enabled = bAgr
  CmdDetalle(1).Enabled = bBorr
  CmdDetalle(2).Enabled = bMod
  CmdDetalle(3).Enabled = bCan
End Sub
Private Sub BotonBuscar(bProv As Boolean, bPro As Boolean)
  CmdProducto.Enabled = bProv
  CmdMateria.Enabled = bPro
End Sub
Private Sub Form_Resize()
  On Error Resume Next
  CmdBotones(0).SetFocus
End Sub
Private Sub CrearRsStock()
 Dim cRsl As ClsLectura
 
 On Error GoTo errHandler

 Set RsCv = New ADODB.Recordset
 Set cRsl = New ClsLectura
 
 Set RsCv = cRsl.RsVacio("CabReceta", "Id", "N")
 
 Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub CrearRsDetalles()
 Dim cRsl As ClsLectura
 
 On Error GoTo errHandler

 Set Rsd = New ADODB.Recordset
 Set cRsl = New ClsLectura
 
 Set Rsd = cRsl.RsVacio("DetalleReceta", "Id", "N")
 CabGrid
 
 Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
 
End Sub
Private Sub CabGrid()
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.Bold = True
  With Grid1
      .Columns(0).Visible = False
      .Columns(1).Visible = False
      .Columns(2).Caption = "Producto"
      .Columns(2).Width = 1000
      .Columns(3).Caption = "Descripción"
      .Columns(3).Width = 5000
      .Columns(4).Caption = "Cant."
      .Columns(4).Width = 800
      .Columns(4).Alignment = dbgRight
      .Columns(4).NumberFormat = "#0.000"
      .Columns(5).Caption = "Costo"
      .Columns(5).Width = 800
      .Columns(5).Alignment = dbgRight
      .Columns(5).NumberFormat = "#0.000"
 End With
End Sub
Private Sub NuevoRsD()
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     CmdBotones(5).Enabled = False
     BotonBuscar False, True
     HabilitarTodo False, False
     HabilitarDetalles True
     LimpiarDetalles
     BotonDetalles False, False, False, True
     CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(0).Caption = "Grabar"
     TxtMateria.SetFocus
  Else
     CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
     CmdDetalle(0).Caption = "Agregar"
     Rsd.AddNew
     GrabarRsd
     BotonDetalles True, True, True, False
     BotonBuscar True, False
     sw1 = True
     HabilitarDetalles False
     HabilitarTodo True, True
     If CmdBotones(0).Caption = "Grabar" Then
        Botones True, False, False, False, False, True, False
     Else
        Botones False, False, False, False, False, True, True
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub BorrarRsD()
  On Error GoTo errHandler
  If MsgBox("Borra el Producto Seleccionado ?", 20, "Atención") = 6 Then
     LimpiarDetalles
     BotonDetalles True, False, False, False
     Rsd.Delete
     Rsd.MovePrevious
     If Rsd.BOF = False And Rsd.EOF = False Then
        If Rsd.BOF Then
           Rsd.MoveFirst
        End If
        LinkearDetalles
        Grid1.Bookmark = Rsd.Bookmark
        BotonDetalles True, True, True, False
'        CmdBotones(4).Enabled = HabiValores()
        CmdBotones(0).Enabled = False
     Else
        Botones False, False, False, False, False, True, False
     End If
  End If
Exit Sub
  
errHandler:
   ManejaErrores

End Sub
Private Sub ModificarRsD()
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     nReg = Rsd.Bookmark
     HabilitarTodo True, False
     HabilitarDetalles True
     BotonBuscar True, True
     BotonDetalles False, False, True, True
     CmdDetalle(2).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(2).Caption = "Grabar"
  Else
     CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
     CmdDetalle(2).Caption = "Actualizar"
     GrabarRsd
     BotonDetalles True, True, True, False
     BotonBuscar True, False
     HabilitarTodo True, True
     HabilitarDetalles False
     Rsd.Bookmark = nReg
     Botones False, False, False, False, False, True, False
     If Rsd.RecordCount <> 0 Then
        Botones False, False, False, False, False, True, True
     End If
     sw1 = True
  End If
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CancelarRsd()
  On Error GoTo errHandler

  CmdBotones(5).Enabled = True
  If sw1 = False Then
     If CmdDetalle(0).Caption = "Grabar" Then
        If CmdDetalle(0).Enabled = True Then
           If MsgBox("El Detalle no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        End If
        CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
        CmdDetalle(0).Caption = "Nuevo"
        BotonDetalles True, False, False, False
        HabilitarTodo True, True
        HabilitarDetalles False
        BotonBuscar True, False
        LimpiarDetalles
        If Rsd.EOF = False Or Rsd.BOF = False Then
           BotonDetalles True, True, True, False
           LinkearDetalles
        End If
        sw1 = True
     Else
        CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
        CmdDetalle(2).Caption = "Actualizar"
        BotonDetalles True, True, True, False
        HabilitarTodo True, True
        HabilitarDetalles False
        BotonBuscar True, False
        LinkearDetalles
        sw1 = True
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub GrabarRsd()
  Dim cRsl As ClsLectura
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  
  Rsd!Id = 0
  Rsd!Movimiento = cRsl.TraerValorDeUnCampo("Sistema", "Movimiento", "")
  Rsd!Producto = TxtMateria.Text
  Rsd!Descripcion = TxtDetalle.Text
  Rsd!Cantidad = TxtCantidad.Text
  Rsd!Costo = TxtCosto.Text * TxtCantidad.Text
  Rsd.Update
  VerEstado
  CalcularTotal
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub LimpiarDetalles()
 TxtMateria.Text = ""
 TxtCantidad.Text = "1.00"
 TxtCosto.Text = "0.00"
 TxtDetalle.Text = ""
End Sub
Private Sub LinkearDetalles()
  On Error GoTo errHandler
    
  TxtMateria.Text = Rsd!Producto
  TxtCantidad.Text = Format(Rsd!Cantidad, nCantDecimales)
  TxtDetalle.Text = Rsd!Descripcion
  TxtCosto.Text = Format(Rsd!Costo, nCantDecimales)
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub TxtCantidad_GotFocus()
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimal(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCantidad_LostFocus()
  If TxtCantidad.Text = "" Then TxtCantidad.Text = 1
  TxtCantidad.Text = Format(TxtCantidad.Text, "0.000")
End Sub

Private Sub TxtCosto_GotFocus()
  TxtCosto.SelStart = 0
  TxtCosto.SelLength = Len(TxtCosto.Text)
End Sub

Private Sub TxtCosto_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimal(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCosto_LostFocus()
  If TxtCosto.Text = "" Then TxtCosto.Text = 0
  TxtCosto.Text = Format(TxtCosto.Text, "#0.000")
End Sub

Private Sub TxtDetalle_Change()
  TxtDetalle.Enabled = False
  If TxtDetalle.Text <> "" Then
      CmdDetalle(0).Enabled = True
  End If
End Sub

Private Sub TxtDetalle_LostFocus()
  If TxtDetalle.Text <> "" Then
     If TxtCantidad.Enabled = True Then
        TxtCantidad.SetFocus
     End If
  Else
     If CmdMateria.Enabled = True Then
        CmdMateria.SetFocus
     End If
  End If
End Sub
Private Sub TxtMateria_LostFocus()
  Dim cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set cRpl = New ClsProductoL
  If TxtMateria.Text <> "" Then
     If cRpl.BuscarNombreProducto(TxtMateria.Text) <> "" Then
        TxtDetalle.Text = cRpl.BuscarNombreProducto(TxtMateria.Text)
        If TxtCantidad.Enabled = True Then
           TxtCantidad.SetFocus
        End If
     End If
  End If
  Set cRpl = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtMateria.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  TxtCosto.Enabled = bHabi
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo errHandler

  If sw = False And sw1 = True Then
     If Grid1.Row <> -1 Then
        Rsd.Bookmark = Grid1.Bookmark
        LinkearDetalles
     End If
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub ArmaRSparaMovimiento()
  Dim cRsl As ClsLectura, cComp As ClsComprobantesL
  
  On Error GoTo errHandler

  Set cComp = New ClsComprobantesL
  Set cRsl = New ClsLectura
  
  Set Rs_Sto = cComp.TraerTodosSin("CabReceta", "Id")
  Botones True, False, False, False, False, True, False
  If Rs_Sto.BOF = False And Rs_Sto.EOF = False Then
     Rs_Sto.MoveLast
     Set Rs = cRsl.TraerRS("CabRecetaTraerUno", Rs_Sto!Id, True)
     LinkearTexto
     LLenarDetalles
     nCas = 0
     Botones True, True, True, True, True, True, True
  End If
  Set cRsl = Nothing
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub Siguiente()
  Dim cRsl As ClsLectura
  On Error GoTo errHandler
  Rs_Sto.MoveNext
  If Rs_Sto.EOF Then
     Rs_Sto.MoveLast
     Beep
  End If
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabRecetaTraerUno", Rs_Sto!Id, True)
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Set cRsl = Nothing
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Anterior()
  Dim cRsl As ClsLectura
  On Error GoTo errHandler
  Rs_Sto.MovePrevious
  If Rs_Sto.BOF Then
     Rs_Sto.MoveFirst
     Beep
  End If
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabRecetaTraerUno", Rs_Sto!Id, True)
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Set cRsl = Nothing
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Primero()
  Dim cRsl As ClsLectura
  On Error GoTo errHandler
  Rs_Sto.MoveFirst
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabRecetaTraerUno", Rs_Sto!Id, True)
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Set cRsl = Nothing
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub
Private Sub Ultimo()
  Dim cRsl As ClsLectura
  On Error GoTo errHandler
  Rs_Sto.MoveLast
  Set cRsl = New ClsLectura
  Set Rs = cRsl.TraerRS("CabRecetaTraerUno", Rs_Sto!Id, True)
  CrearRsDetalles
  Limpiar
  LinkearTexto
  LLenarDetalles
  Botones True, True, True, True, True, True, True
  HabilitarTodo False, True
  BotonDetalles False, False, False, False
  Set cRsl = Nothing
  Exit Sub
  
errHandler:
   ManejaErrores
End Sub

Private Sub LLenarDetalles()
  Dim rAuxD As ADODB.Recordset, cRsl As ClsLectura
  On Error GoTo errHandler
 
  Set rAuxD = New ADODB.Recordset
  Set cRsl = New ClsLectura
  Set rAuxD = cRsl.TraerRS("DetalleRecetaTraerUnos", Rs!Movimiento, True)
  rAuxD.MoveFirst
  Do While Not rAuxD.EOF
     Rsd.AddNew
     Rsd!Movimiento = rAuxD!Movimiento
     Rsd!Producto = rAuxD!Producto
     Rsd!Descripcion = rAuxD!Descripcion
     Rsd!Cantidad = rAuxD!Cantidad
     Rsd!Costo = rAuxD!Costo
     Rsd.Update
     rAuxD.MoveNext
  Loop
'  CabGrid
  CalcularTotal
  Set cRsl = Nothing
  rAuxD.Close
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub Nuevo()
  On Error GoTo errHandler

  If sw = True Then
     Limpiar
     HabilitarTodo True, False
     HabilitarDetalles False
     Botones False, False, False, False, False, True, False
     BotonDetalles False, False, False, False
     BotonBuscar True, False
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(5).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(5).Caption = "Cancelar"
     sw = False
     CrearRsStock
     CrearRsDetalles
     nMovi = 0
     nCas = 1
  Else
     sw = True
     Botones False, False, False, False, False, False, False
     GrabarTodo
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
     CmdBotones(5).Caption = "Salir"
     HabilitarTodo False, True
     Botones True, True, True, True, True, True, True
     BotonDetalles False, False, False, False
     BotonBuscar False, False
     nMovi = RsCv!Movimiento
     nCas = 0
  End If
Exit Sub

errHandler:
     ManejaErrores
End Sub
Private Sub Salir()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        If CmdBotones(0).Enabled = True Then
           If MsgBox("El Comprobante no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        Else
           If MsgBox("Cancela la Crearción del Comprobante ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        End If
        CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
        CmdBotones(0).Caption = "Nuevo"
        CmdBotones(5).Picture = LoadResPicture("Salir", 0)
        CmdBotones(5).Caption = "Salir"
     Else
        If CmdBotones(9).Enabled = True Then
           If MsgBox("El Comprobante no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        Else
           If MsgBox("Cancela la Modificación del Comprobante ?", 20, "Atención") = vbNo Then
              Exit Sub
           End If
        End If
        CmdBotones(9).Picture = LoadResPicture("Modificar", 0)
        CmdBotones(9).Caption = "Modificar"
        CmdBotones(5).Picture = LoadResPicture("Salir", 0)
        CmdBotones(5).Caption = "Salir"
     End If
     HabilitarTodo False, False
     sw = True
     nTip = 0
     Limpiar
     LimpiarDetalles
     CrearRsDetalles
     Botones True, False, False, False, False, True, False
     BotonDetalles False, False, False, False
     BotonBuscar False, False
     If Rs_Sto.BOF = False Or Rs_Sto.EOF = False Then
        Rs_Sto.MoveLast
        Set Rs = cRsl.TraerRS("CabRecetaTraerUno", Rs_Sto!Id, True)
        LinkearTexto
        LLenarDetalles
        nCas = 0
        BotonDetalles False, False, False, False
        Botones True, True, True, True, True, True, True
     End If
  Else
     Unload Me
  End If
  Set cRsl = Nothing
End Sub
Private Sub CalcularTotal()
  Dim rReg As Variant
  On Error GoTo errHandler
  
  LblTotal.Caption = "0.00"
  
  ' Calcular Precio Neto
  If Rsd.BOF = False And Rsd.EOF = False Then
     rReg = Rsd.Bookmark
     Rsd.MoveFirst
        ' Neto sin Iva
     Do While Not Rsd.EOF
        LblTotal.Caption = CCur(LblTotal.Caption) + Rsd!Costo
        Rsd.MoveNext
     Loop
           ' Calculo el descuento en base al neto con iva incluido
     LblTotal.Caption = Format(LblTotal.Caption, nCantDecimales)
     Rsd.Bookmark = rReg
  End If
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub VerEstado()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        CmdBotones(4).Enabled = False
        If Len(TxtProducto.Text) <> 0 And Rsd.RecordCount <> 0 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(9).Enabled = False
        CmdBotones(4).Enabled = False
        If Len(TxtProducto.Text) <> 0 And Rsd.RecordCount <> 0 Then
           CmdBotones(9).Enabled = True
        End If
     End If
  End If
End Sub
Private Sub GrabarTodo()
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL
  Dim cRsl As ClsLectura
  Dim I As Byte

  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cRsl = New ClsLectura


  ' Grabar Cabecera Stock

  RsCv.AddNew
  RsCv!Id = 0
  If CmdBotones(0).Caption = "Grabar" Then
     RsCv!Movimiento = cNum.TraerUltimoMovimiento
  Else
     RsCv!Movimiento = Rs!Movimiento
  End If
  RsCv!Producto = TxtProducto.Text
  RsCv!Observacion = TxtObservacion.Text
  RsCv.Update
  ' Actualizo Stock

  Rsd.MoveFirst

  If CmdBotones(0).Caption = "Grabar" Then
     cGrab.GrabarReceta RsCv, Rsd
     CrearRsDetalles
     Limpiar
     ArmaRSparaMovimiento
    Else
       cGrab.ModificarReceta RsCv, Rsd
    End If
 ' End If
  nTip = 0

  Set cGrab = Nothing
  Set cNum = Nothing
  Set cRsl = Nothing
End Sub
Private Sub Buscar()
'  ComproBuscar.Show 1
'  If ComproBuscar.nId <> 0 Then
'     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
'     CmdBotones(0).Caption = "Nuevo"
'     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
'     CmdBotones(5).Caption = "Salir"
'     HabilitarTodo False, False
'     sw = True
'     Limpiar
'     LimpiarDetalles
'     CrearRsDetalles
'     Botones True, False, False, False, False, True, False
'     BotonDetalles False, False, False, False
'     BotonBuscar False, False
'     If Rs_Sto.BOF = False Or Rs_Sto.EOF Then
'        Rs_Sto.MoveLast
'        Set Rs = cRsl.TraerRS("CabezaTraerUno", ComproBuscar.nId, True)
'        LinkearTexto
'        LLenarDetalles
'        nCas = 0
'        BotonDetalles False, False, False, False
'        Botones True, True, True, True, True, True, True
'     End If
'  End If
End Sub

Private Sub Anular()
  Dim cBor As ClsComprobantesE
  
  Set cBor = New ClsComprobantesE
  If MsgBox("Anula el Comprobante Seleccionado ?", 20, "Atención") = 6 Then
     cBor.AnularComprobanteCantidades (Rs!Movimiento)
     LinkearTexto
     Botones True, True, True, True, True, True, True
  End If
End Sub

Private Sub Imprimir()
'  Dim cClie As ClsClienteL
'  Dim rCli As ADODB.Recordset
'
'  Set rImp = New ADODB.Recordset
'  Set cClie = New ClsClienteL
'
'  rImp.Fields.Append "Fecha", adDate
'  rImp.Fields.Append "Numero", adVarChar, 15, adFldIsNullable
'  rImp.Fields.Append "Cliente", adVarChar, 15, adFldIsNullable
'  rImp.Fields.Append "RazonSocial", adVarChar, 50, adFldIsNullable
'  rImp.Fields.Append "Direccion", adVarChar, 50, adFldIsNullable
'  rImp.Fields.Append "Iva", adVarChar, 30, adFldIsNullable
'  rImp.Fields.Append "CPostal", adVarChar, 8, adFldIsNullable
'  rImp.Fields.Append "Localidad", adVarChar, 30, adFldIsNullable
'  rImp.Fields.Append "Provincia", adVarChar, 30, adFldIsNullable
'  rImp.Fields.Append "Cuit", adVarChar, 13, adFldIsNullable
'  rImp.Fields.Append "Codigo", adVarChar, 15, adFldIsNullable
'  rImp.Fields.Append "Cantidad", adDouble
'  rImp.Fields.Append "Descripcion", adVarChar, 80, adFldIsNullable
'  rImp.Fields.Append "Corte", adInteger, , adFldIsNullable
'  rImp.Fields.Append "PUnitario", adCurrency
'  rImp.Open
'  Rsd.MoveFirst
'  Set rCli = cClie.BuscarDatosCli(TxtProducto.Text)
'
'  Do While Not Rsd.EOF
'     rImp.AddNew
'     rImp!Fecha = DTPFecha.Value
'     rImp!Numero = LblNumero.Caption
'     rImp!Cliente = TxtProveedor.Text
'     rImp!RazonSocial = LblProveedor.Caption
'     If rCli.RecordCount <> 0 Then
'        rImp!Direccion = rCli!Domicilio
'        rImp!Iva = rCli![TipoIva.Descripcion]
'        rImp!CPostal = rCli!CodigoPostal
'        rImp!Localidad = rCli![Localidades.Descripcion]
'        rImp!Provincia = rCli![Provincias.Descripcion]
'        rImp!Cuit = rCli!Cuit
'     End If
'     rImp!Codigo = Rsd!Producto
'     rImp!Cantidad = Rsd!Cantidad
'     rImp!Descripcion = Rsd!Descripcion
'     rImp!Corte = Rsd!Corte
'     rImp!PUnitario = 0
'     rImp.Update
'     Rsd.MoveNext
'  Loop
'  nImpr = 26
'  FrmImpresor.Show
End Sub
Private Sub Modificar()
  On Error GoTo errHandler

  If sw = True Then
     HabilitarTodo True, True
     HabilitarDetalles False
     BotonBuscar True, False
     CrearRsStock
     Rsd.MoveLast
     Grid1.Bookmark = Rsd.Bookmark
     LinkearDetalles
     Botones False, False, True, False, False, True, True
     BotonDetalles True, True, True, False
     CmdBotones(9).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(9).Caption = "Grabar"
     CmdBotones(5).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(5).Caption = "Cancelar"
     sw = False
     sw1 = True
  Else
     Botones False, False, False, False, False, False, False
     GrabarTodo
     sw = True
     CmdBotones(9).Picture = LoadResPicture("Modificar", 0)
     CmdBotones(9).Caption = "Modificar"
     CmdBotones(5).Picture = LoadResPicture("Salir", 0)
     CmdBotones(5).Caption = "Salir"
     HabilitarTodo False, True
     BotonBuscar False, False
     Botones True, True, True, True, True, True, True
     BotonDetalles False, False, False, False
     CmdBotones(0).SetFocus
  End If
Exit Sub

errHandler:
  MsgBox Err.Description, vbCritical, "Modificar"

End Sub

Private Sub TxtProducto_LostFocus()
  Dim cRpl As ClsProductoL
  On Error GoTo errHandler
  
  Set cRpl = New ClsProductoL
  If TxtProducto.Text <> "" Then
     If cRpl.BuscarNombreProducto(TxtProducto.Text) <> "" Then
        LblProducto.Caption = cRpl.BuscarNombreProducto(TxtProducto.Text)
        CmdDetalle(0).Enabled = True
     End If
  End If
  Set cRpl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub
