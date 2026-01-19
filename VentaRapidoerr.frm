VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form VentaRapido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas a Consumidores Finales"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   9375
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   960
      TabIndex        =   26
      Top             =   720
      Width           =   5895
      Begin VB.Frame Frame3 
         Caption         =   "Su pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   4815
         Begin VB.CommandButton CmdVuelto 
            Caption         =   "Anular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   30
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton CmdVuelto 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   29
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox TxtPago 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   3000
            TabIndex        =   28
            Text            =   "TxtPago"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label LblVuelto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblVuelto"
            Height          =   330
            Left            =   3000
            TabIndex        =   35
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label LblCompra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCompra"
            Height          =   330
            Left            =   3000
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Su Vuelto :"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Su Pago :"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Total de su Compra :"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   2415
         End
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   135
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   1
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
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
      Height          =   3810
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   8010
      Begin VB.TextBox TxtProducto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   690
         MaxLength       =   15
         TabIndex        =   1
         Top             =   2550
         Width           =   1470
      End
      Begin VB.CommandButton CmdProductos 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5445
         TabIndex        =   5
         Top             =   2565
         Width           =   705
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   855
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "TxtCa"
         Top             =   2880
         Width           =   930
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   690
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "TxtPrecio"
         Top             =   3195
         Width           =   1095
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   1905
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2895
         Width           =   870
      End
      Begin VB.TextBox TxtDetalle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2235
         MaxLength       =   40
         TabIndex        =   2
         Text            =   "TxtDetalle"
         Top             =   2550
         Width           =   3135
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   3
         Left            =   5265
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2895
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   2
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2895
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   2835
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2895
         Visible         =   0   'False
         Width           =   870
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2130
         Left            =   165
         TabIndex        =   10
         Top             =   255
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   3757
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
      Begin MSCommLib.MSComm MSComm1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   -1  'True
         RThreshold      =   1
         ParitySetting   =   3
         SThreshold      =   1
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   17
         Top             =   2565
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Cantidad :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   2895
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Precio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   3195
         Width           =   750
      End
      Begin VB.Label Label10 
         Caption         =   "Neto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6270
         TabIndex        =   14
         Top             =   2550
         Width           =   525
      End
      Begin VB.Label LblNeto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNeto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6840
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6270
         TabIndex        =   12
         Top             =   3375
         Width           =   495
      End
      Begin VB.Label LblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTotal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6840
         TabIndex        =   11
         Top             =   3375
         Width           =   1095
      End
   End
   Begin VB.Label LblNumero 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblNumero"
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
      Left            =   6600
      TabIndex        =   25
      Top             =   150
      Width           =   1455
   End
   Begin VB.Label LblCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblCliente"
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
      Left            =   1440
      TabIndex        =   24
      Top             =   540
      Width           =   3255
   End
   Begin VB.Label LblComprobante 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblComprobante"
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
      Left            =   1440
      TabIndex        =   23
      Top             =   150
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   22
      Top             =   540
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   21
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   20
      Top             =   150
      Width           =   1140
   End
End
Attribute VB_Name = "VentaRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_Ven As ADODB.Recordset         ' Cabecera Venta Tabla Uno
Dim RsCv As ADODB.Recordset           ' Cabecera Virtual Venta
Dim Rsd As ADODB.Recordset            ' Cabecera Vitual VentaDetalle
Dim Rs As ADODB.Recordset             ' Cabecera Venta Tabla Todos

Dim cRsl As ClsLectura
Dim sw1 As Boolean, sw As Boolean
Dim Cadena As String, ic As Byte
Dim nCompro As Integer, nCaja As Integer
Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Salir
  End Select
End Sub
Private Sub CmdProductos_Click()
  Dim RsP As ADODB.Recordset, cRpl As ClsProductoL, nDigi As Byte, nLen As Byte, cPrec As Double, cPrecAux As Double
  On Error GoTo errHandler
  
  Set RsP = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  Set cRpl = New ClsProductoL
  nBuscar = 3
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Set RsP = cRsl.TraerRS("ProductoTraerUno", nDat, True)
     If nDigito <> 0 Then
        nDigi = nDigito
     End If
     If Not cRpl.ProductoCompuesto(RsP!Producto) = True Then
         TxtProducto.Text = RsP!Producto
         TxtDetalle.Text = RsP!Descripcion
         TxtPrecio.Text = Format(cRpl.TraerPrecio(cRsl.RegPorDefecto("ListaPrecio"), RsP!Producto), "#0.00")
     End If
     TxtCantidad_LostFocus
     RsP.Close
     If TxtPrecio.Text <> 0 Then
         NuevoRsD
         SendKeys "{ENTER}"
     Else
        TxtPrecio.SetFocus
     End If
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   MsgBox Err.Description, vbCritical, "CmdProductos_Click"
  
End Sub
Private Sub CmdVuelto_Click(Index As Integer)
  Select Case Index
         Case 0
              GrabarTodo
              Frame2.Visible = False
         Case 1
              Frame2.Visible = False
              LblVuelto.Caption = "0.00"
              Botones True, True
              BotonDetalles True, True, True, False
  End Select
End Sub
Private Sub Form_Load()

  Set cRsl = New ClsLectura
  CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
  CmdBotones(1).Picture = LoadResPicture("Salir", 0)

  CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
  CmdDetalle(1).Picture = LoadResPicture("Borrar", 0)
  CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
  CmdDetalle(3).Picture = LoadResPicture("Cancelar", 0)

 Set Rs_Ven = New ADODB.Recordset
 CrearRsDetalles
 Limpiar
 sw = True
 Frame2.Visible = False
 HabilitarTodo False, False
 BotonDetalles False, False, False, False
 BotonBuscar False
 HabilitarDetalles False

' *******************************************************************
' Atención Si se trabaja con dos Caja Tener cuidado con los codigos
' Comprobante Caja 1 = 6
' Caja Caja 1 = 1
' Comprobante Caja 2 = 45
' Caja Caja 2 = 2
' ******************************************************************

 nCompro = 45
 nCaja = 2
 Me.Top = 0
 Me.Left = 0

End Sub
Private Sub Botones(bNuevo As Boolean, bSalir As Boolean)
  CmdBotones(0).Enabled = bNuevo
  CmdBotones(1).Enabled = bSalir
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
     End If
End Sub
Private Sub Limpiar()
  On Error GoTo Errores
  Dim ctl As Control
  
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.Text = ""
      End If
  Next ctl
  Set cRsl = New ClsLectura
  LblCliente.Caption = cRsl.DatoCombo("Clientes", "Id", "RazonSocial", cRsl.RegPorDefecto("ClienteConsFinal"))
  LblComprobante.Caption = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", nCompro)
 ' LblLista.Caption = cRsl.DatoCombo("ListadePrecio", "ListaPrecio", "Descripcion", cRsl.RegPorDefecto("ListaPrecio"))
 ' LblPago.Caption = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", cRsl.RegPorDefecto("Cobro"))
  LblCompra.Caption = "0.00"
  TxtPago.Text = "0.00"
  LblVuelto.Caption = "0.00"
  LblNumero.Caption = "0000-00000000"
  LblNeto.Caption = "0.00"
  LblTotal.Caption = "0.00"
  Exit Sub
Errores:
     MsgBox Err.Description, vbCritical, "Limpiar"
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
Private Sub BotonDetalles(bAgr As Boolean, bBorr As Boolean, bMod As Boolean, bCan As Boolean)
  CmdDetalle(0).Enabled = bAgr
  CmdDetalle(1).Enabled = bBorr
  CmdDetalle(2).Enabled = bMod
  CmdDetalle(3).Enabled = bCan
End Sub
Private Sub BotonBuscar(bPro As Boolean)
  CmdProductos.Enabled = bPro
End Sub
Private Sub Form_Resize()
  CmdBotones(0).SetFocus
End Sub
Private Sub CrearRsDetalles()
 On Error GoTo errHandler

 Set Rsd = New ADODB.Recordset
 Set cRsl = New ClsLectura

 Set Rsd = cRsl.RsVacio("DetallesComprobantes", "Id", "N")
 CabGrid

Exit Sub

errHandler:
      MsgBox Err.Description, vbCritical, "CrearRsDetalles"

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF12 And Shift = vbCtrlMask Then
      If nPuerto <> 0 Then
         nPuerto = 0
      End If
   End If
   If KeyCode = vbKeyF12 Then
      If TxtCantidad.Enabled = True Then
         TxtCantidad.SetFocus
      End If
   End If
   If KeyCode = vbKeyF5 Then
      If sw = True Then
         Nuevo
         NuevoRsD
      Else
         If CmdBotones(0).Enabled = True Then
            If CmdDetalle(0).Enabled = False Then
               CancelarRsd
            End If
            Nuevo
         End If
      End If
   End If
End Sub
Private Sub CabGrid()
  Set Grid1.DataSource = Rsd
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.Bold = True
  With Grid1
      .Columns(0).Visible = False
      .Columns(1).Visible = False
      .Columns(2).Caption = "Producto"
      .Columns(2).Width = 1400
      .Columns(3).Caption = "Descripción"
      .Columns(3).Width = 3300
      .Columns(4).Caption = "Cant."
      .Columns(4).Width = 650
      .Columns(4).Alignment = dbgRight
      .Columns(4).NumberFormat = "#0.00"
      .Columns(5).Caption = "P.Uni."
      .Columns(5).Width = 800
      .Columns(5).NumberFormat = "#0.00"
      .Columns(5).Alignment = dbgRight
      .Columns(6).Caption = "Total"
      .Columns(6).Width = 1000
      .Columns(6).NumberFormat = "#0.00"
      .Columns(6).Alignment = dbgRight
 End With
End Sub
Private Sub HabilitarDetalles(bHabi As Boolean)
  TxtProducto.Enabled = bHabi
  TxtDetalle.Enabled = bHabi
  TxtCantidad.Enabled = bHabi
  TxtPrecio.Enabled = bHabi
End Sub
Private Sub NuevoRsD()
  On Error GoTo errHandler

  If sw1 = True Then
     sw1 = False
     If nScaner <> 0 Then
        MSComm1.CommPort = nScaner
        MSComm1.PortOpen = True
     End If
     Cadena = ""
     ic = 1
     BotonBuscar True
     HabilitarTodo False, False
     HabilitarDetalles True
     LimpiarDetalles
     BotonDetalles False, False, False, True
     CmdDetalle(0).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(0).Caption = "Grabar"
     TxtProducto.SetFocus
  Else
      If nScaner <> 0 Then
         MSComm1.PortOpen = False
      End If
     CmdDetalle(0).Picture = LoadResPicture("Nuevo", 0)
     CmdDetalle(0).Caption = "Agregar"
     Rsd.AddNew
     GrabarRsd
     BotonDetalles True, True, True, False
     BotonBuscar False
     CalcularTotales
     sw1 = True
     HabilitarTodo True, True
     HabilitarDetalles False
     CmdBotones(0).Enabled = False
     CmdBotones(1).Enabled = True
     VerEstado
  End If
Exit Sub

errHandler:
      MsgBox Err.Description, vbCritical, "NuevoRsD"
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
        Botones False, True
     End If
  End If
  CalcularTotales
Exit Sub

errHandler:
   MsgBox Err.Description, vbCritical, "BorrarRsD"

End Sub
Private Sub ModificarRsD()
  On Error GoTo errHandler

  If sw1 = True Then
     If nScaner <> 0 Then
        MSComm1.PortOpen = True
     End If
     Cadena = ""
     ic = 1
     sw1 = False
     HabilitarTodo True, False
     HabilitarDetalles True
     BotonBuscar True
     CmdBotones(0).Enabled = False
     BotonDetalles False, False, True, True
     CmdDetalle(2).Picture = LoadResPicture("Grabar", 0)
     CmdDetalle(2).Caption = "Grabar"
  Else
     If nScaner <> 0 Then
        MSComm1.PortOpen = False
     End If
     CmdDetalle(2).Picture = LoadResPicture("Modificar", 0)
     CmdDetalle(2).Caption = "Actualizar"
     GrabarRsd
     BotonDetalles True, True, True, False
     BotonBuscar False
     CalcularTotales
     sw1 = True
     HabilitarTodo True, True
     HabilitarDetalles False
     CalcularTotales
     CmdBotones(0).Enabled = True
  End If
Exit Sub

errHandler:
  MsgBox Err.Description, vbCritical, "ModificarRsD"

End Sub
Private Sub CancelarRsd()
  On Error GoTo errHandler

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
        BotonBuscar False
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
        BotonBuscar False
        LinkearDetalles
        CmdBotones(0).Enabled = True
        sw1 = True
     End If
  End If
  If nScaner <> 0 Then
     MSComm1.PortOpen = False
  End If
Exit Sub

errHandler:
   MsgBox Err.Description, vbCritical, "CancelarRsd"

End Sub
Private Sub GrabarRsd()
  On Error GoTo errHandler

  Rsd!Id = 0
  Rsd!Producto = TxtProducto.Text
  Rsd!Descripcion = TxtDetalle.Text
  Rsd!Cantidad = TxtCantidad.Text
  Rsd!PrecioUnitario = TxtPrecio.Text
  Rsd!PrecioTotal = TxtCantidad.Text * TxtPrecio.Text
  Rsd.Update
  CabGrid
'  VerEstado
Exit Sub

errHandler:
   MsgBox Err.Description, vbCritical, "GrabarRsd"
End Sub
Private Sub LimpiarDetalles()
 TxtProducto.Text = ""
 TxtCantidad.Text = "1.00"
 TxtDetalle.Text = ""
 TxtPrecio.Text = "0.00"
End Sub
Private Sub LinkearDetalles()
  Dim cRsPl As ClsProductoL
  On Error GoTo errHandler

  Set cRsPl = New ClsProductoL

  TxtProducto.Text = Rsd!Producto
  TxtCantidad.Text = Format(Rsd!Cantidad, "#0.00")
  TxtDetalle.Text = cRsPl.BuscarNombreProducto(TxtProducto.Text)
  TxtPrecio.Text = Format(Rsd!PrecioUnitario, "#0.00")
Exit Sub

errHandler:
     MsgBox Err.Description, vbCritical, "LinkearDetalles"

End Sub
Private Sub Nuevo()
  On Error GoTo errHandler

  If sw = True Then
     Limpiar
     HabilitarTodo True, False
     HabilitarDetalles False
     Botones False, True
     BotonDetalles True, False, False, False
     BotonBuscar False
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(0).Caption = "&Grabar"
     CmdBotones(1).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(1).Caption = "&Cancelar"
     sw = False
     sw1 = True
     CrearRsVenta
     CrearRsDetalles
     Numero
     CmdDetalle(0).SetFocus
  Else
     Frame2.Visible = True
     TxtPago.SetFocus
     Botones False, False
     BotonDetalles False, False, False, False
     ' GrabarTodo
  End If
Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Nuevo"
End Sub
Private Sub Salir()
  If sw = False Then
     If CmdBotones(0).Enabled = True Then
        If MsgBox("El Comprobante no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
           Exit Sub
        End If
     Else
        If MsgBox("Cancela la Crearción del Comprobante ?", 20, "Atención") = vbNo Then
           Exit Sub
        End If
     End If
     If CmdDetalle(0).Caption = "Grabar" Then
        CancelarRsd
     End If
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(0).Caption = "&Nuevo"
     CmdBotones(1).Picture = LoadResPicture("Salir", 0)
     CmdBotones(1).Caption = "&Salir"
     HabilitarTodo False, True
     sw = True
     Limpiar
     LimpiarDetalles
     CrearRsDetalles
     Botones True, True
     BotonDetalles False, False, False, False
     BotonBuscar False
     CmdBotones(0).SetFocus
  Else
     Unload Me
  End If
End Sub
Private Sub CrearRsVenta()
 On Error GoTo errHandler

 Set RsCv = New ADODB.Recordset
 Set cRsl = New ClsLectura
 
 Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
 
Exit Sub

errHandler:
     MsgBox Err.Description, vbCritical, "CrearRsVenta"
End Sub
Private Sub GrabarTodo()
  On Error GoTo Errores
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cTipo As ClsComprobantesL
  Dim rCuen As ADODB.Recordset, rValo As ADODB.Recordset, rApli As ADODB.Recordset
  Dim rReci As ADODB.Recordset, rCant As ADODB.Recordset, rCaja As ADODB.Recordset
  Dim cVeri As ClsComprobantesL, cCuota As ClsComprobantesL, cDias As ClsComprobantesL, cStock As ClsComprobantesL
  Dim i As Byte, nCuotas As Byte, nDias As Integer, rValoAux As ADODB.Recordset
  ' declaracion para clientes que emiten pagare
  Dim cTotalPagare As Currency, nCuotaPagare As Integer, fFechaPagare As Date
  
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cTipo = New ClsComprobantesL
  Set rCuen = New ADODB.Recordset
  Set rValo = New ADODB.Recordset
  Set rValoAux = New ADODB.Recordset
  Set rApli = New ADODB.Recordset
  Set cVeri = New ClsComprobantesL
  Set cCuota = New ClsComprobantesL
  Set cDias = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  Set rReci = New ADODB.Recordset
  Set rCant = New ADODB.Recordset
  Set rCaja = New ADODB.Recordset
  Set cRsl = New ClsLectura
  Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
  Set rValo = cRsl.RsVacio("Valores", "Id", "N")
  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
  Set rCant = cRsl.RsVacio("Cantidades", "Id", "N")
  Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
    
  ' Grabar Cabecera Venta
  
  sw = True
  Botones False, False
    
  RsCv.AddNew
  RsCv!Id = 0
  RsCv!Tipo = 1
  RsCv!Movimiento = cNum.TraerUltimoMovimiento
  RsCv!Comprobante = nCompro
  RsCv!Fecha = Date
  RsCv!Numero = Right(LblNumero.Caption, 8)
  RsCv!Sucursal = Left(LblNumero.Caption, 4)
  RsCv!Cliente = cRsl.ClientePorDefecto(cRsl.RegPorDefecto("ClienteConsFinal"))
  RsCv!Lista = cRsl.RegPorDefecto("ListaPrecio")
  RsCv!Vendedor = cRsl.RegPorDefecto("Vendedor")
  RsCv!Caja = nCaja
  RsCv!CondVenta = 1
  RsCv!Neto = CCur(LblNeto.Caption)
  RsCv!Iva1 = 0
  RsCv!Iva2 = 0
  RsCv!Descuentos = 0
  RsCv!Financiacion = 0
  Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
         Case 1
              RsCv!Debe = CCur(LblTotal.Caption)
              RsCv!Haber = 0
         Case 2
              RsCv!Debe = 0
              RsCv!Haber = CCur(LblTotal.Caption)
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = ""
  RsCv.Update
  ' Detalle Venta ya esta Armado en el recordset Rsd
  ' Forma de Pagos,Valores,Aplicaciones,Cantidades,Caja
  ' Actualizo Stock
  
  Set rValoAux = VentaValores.RsV
  Set cRsl = New ClsLectura
  If cStock.ActualizaCtaCte(nCompro) = True Then
      rValo.AddNew
      rValo!Id = 0
      rValo!Venta = cNum.TraerUltimoMovimiento
      rValo!CodPago = 1
      rValo!FormaPago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", cRsl.RegPorDefecto("Cobro"))
      rValo!Importe = CCur(LblTotal.Caption)
      rValo.Update
  
      rReci.AddNew
      rReci!Id = 0
      rReci!Fecha = Date
      rReci!Cliente = cRsl.ClientePorDefecto(cRsl.RegPorDefecto("ClienteConsFinal"))
      rReci!Venta = RsCv!Movimiento
      rReci!Comprobante = nCompro
      rReci!Numero = cNum.TraerUltimoNumero(nCompro)
      rReci!Sucursal = cNum.TraerSucursalNumero(nCompro)
      rReci!Cuota = 0
      rReci!Debe = 0
      rReci!Haber = CCur(LblTotal.Caption)
      rReci.Update
  End If
  ' Actualizo Stock
  
 ' Grabar Valores
  db.Execute "UPDATE Sistema SET Movimiento = Movimiento + 1"
 

  If cStock.ActualizaStock(nCompro) = True Then
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        rCant.AddNew
        rCant!Producto = Rsd!Producto
        rCant!Deposito = cStock.DepositoVenta
        Select Case cTipo.QueMoviStock(nCompro)
               Case "Suma"
                     rCant!CantidadDebe = Rsd!Cantidad
                     rCant!CantidadHaber = 0
               Case "Resta"
                     rCant!CantidadDebe = 0
                     rCant!CantidadHaber = Rsd!Cantidad
        End Select
        rCant.Update
        Rsd.MoveNext
     Loop
     Rsd.MoveFirst
  End If
  ' Actualizo Caja
  If cStock.ActualizaCaja(nCompro) = True Then
     Select Case cTipo.TraerTipoMovimiento(RsCv!Comprobante)
            Case 1
                 If rValo.EOF = False Or rValo.BOF = False Then
                    rValo.MoveFirst
                    Do While Not rValo.EOF
                       If cTipo.TraerTipoCondicion(rValo!CodPago) <> 6 Then
                          rCaja.AddNew
                          rCaja!Movimiento = RsCv!Movimiento
                          rCaja!Comprobante = cTipo.TraerTipoRecibo(RsCv!Comprobante)
                          rCaja!Numero = cTipo.TraerUltimoNumero(rCaja!Comprobante)
                          rCaja!Sucursal = cTipo.TraerSucursalNumero(rCaja!Comprobante)
                          rCaja!Fecha = Date
                          rCaja!Hora = Time
                          rCaja!Caja = nCaja
                          rCaja!FormaPago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", cRsl.RegPorDefecto("Cobro"))
                          rCaja!Nombre = rValo!Nombre
                          rCaja!Banco = rValo!Banco
                          rCaja!NumeroCheque = rValo!NumeroCheque
                          rCaja!FechaAcreditacion = rValo!FechaAcreditacion
                          rCaja!NumeroTarjeta = rValo!NumeroTarjeta
                          rCaja!FechaVencimiento = rValo!FechaVencimiento
                          rCaja!Autorizacion = rValo!Autorizacion
                          rCaja!Cupon = rValo!Cupon
                          If cStock.SumaCaja(nCompro) = True Then
                             rCaja!Debe = rValo!Importe
                             rCaja!Haber = 0
                          Else
                             rCaja!Debe = 0
                             rCaja!Haber = rValo!Importe
                          End If
                          rCaja!Abierta = "S"
                          rCaja.Update
                    End If
                    rValo.MoveNext
                    Loop
                 End If
            Case 2
                 rCaja.AddNew
                 rCaja!Movimiento = RsCv!Movimiento
                 rCaja!Comprobante = cTipo.TraerTipoRecibo(RsCv!Comprobante)
                 rCaja!Numero = cTipo.TraerUltimoNumero(rCaja!Comprobante)
                 rCaja!Sucursal = cTipo.TraerSucursalNumero(rCaja!Comprobante)
                 rCaja!Fecha = Date
                 rCaja!Hora = Time
                 rCaja!Caja = nCaja
                 rCaja!FormaPago = rValo!FormaPago
                 rCaja!Nombre = ""
                 rCaja!Banco = ""
                 rCaja!NumeroCheque = ""
                 rCaja!FechaAcreditacion = Null
                 rCaja!NumeroTarjeta = ""
                 rCaja!FechaVencimiento = ""
                 rCaja!Autorizacion = ""
                 rCaja!Cupon = ""
                 If cStock.SumaCaja(nCompro) = True Then
                    rCaja!Debe = rValo!Importe
                    rCaja!Haber = 0
                 Else
                    rCaja!Debe = 0
                    rCaja!Haber = rValo!Importe
                 End If
                 rCaja!Abierta = "S"
                 rCaja.Update
     End Select
  End If
  
 ' proceso de facturacion fiscal
' El siguiente codigo es segun la condicion de iva del Comerciante
  
  Dim cDisc As ClsComprobantesL, rIvaDis As ADODB.Recordset
  Dim cCuit As ClsClienteL
 
volver:
  On Error GoTo Errores:
  Set cDisc = New ClsComprobantesL
  Set cCuit = New ClsClienteL
  
 Imprimir
' Si el sistema no usa controlador fiscal sacar todas estas lines
  ' Actualizo Ultimos Numero de Comrprobantes
  ' Grabo Todo en los Store Procedure
  
  
 cGrab.Agregar RsCv, Rsd, rReci, rValo, rApli, rCant, rCaja
  
 CrearRsDetalles
 Limpiar
  
  ' Fin del proceso de facturacion fiscal
  
  Set cGrab = Nothing
  Set cNum = Nothing
  Set cTipo = Nothing
  Set cVeri = Nothing
  Set cCuota = Nothing
  Set cDias = Nothing
  Set cStock = Nothing
  CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
  CmdBotones(0).Caption = "&Nuevo"
  CmdBotones(1).Picture = LoadResPicture("Salir", 0)
  CmdBotones(1).Caption = "&Salir"
  HabilitarTodo False, True
  Botones True, True
  BotonDetalles False, False, False, False
  BotonBuscar False
  CmdBotones(0).SetFocus

Exit Sub
 
Errores:
     MsgBox Err.Description, vbCritical, "GrabarTodo"
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
Public Sub CalcularTotales()
  Dim rReg As Variant, cDescTotal As Currency, cDescParcial As Currency, sClien As String
  Dim cImpIva1 As Currency, cImpIva2 As Currency, i As Byte, cTotPar As Currency
  Dim rAux As ADODB.Recordset
  
  On Error GoTo errHandler
  
  LblNeto.Caption = "0.00"
  ' Calcular Precio Neto
  If Rsd.BOF = False And Rsd.EOF = False Then
     rReg = Rsd.Bookmark
     Rsd.MoveFirst
     Do While Not Rsd.EOF
        LblNeto.Caption = CCur(LblNeto.Caption) + Rsd!PrecioTotal
        LblNeto.Caption = Format(LblNeto.Caption, "#0.00")
        Rsd.MoveNext
     Loop
     Rsd.Bookmark = rReg
     ' Calcular Descuento
     Set cRsl = New ClsLectura
     Set rAux = New ADODB.Recordset
     sClien = cRsl.RegPorDefecto("ClienteConsFinal")
     If sClien <> "" Then
        Set rAux = cRsl.TraerRS("ClienteDescuentoTraer", sClien, True)
        cDescTotal = 0
        cDescParcial = 0
        Do While Not rAux.EOF
           cDescParcial = CCur(LblNeto.Caption) * rAux!Porcentaje / 100
           cDescTotal = cDescTotal + cDescParcial
           rAux.MoveNext
        Loop
     Else
        MsgBox "No se Puede Calcular Los Descuentos el Campo Cliente está Vacio", vbCritical, "Atención"
     End If
     cTotPar = CCur(LblNeto.Caption)
     ' Calcular Financiación
     ' CalcularImpuestos
     Set cRsl = New ClsLectura
     Set rAux = New ADODB.Recordset
     sClien = cRsl.RegPorDefecto("ClienteConsFinal")
     Set rAux = cRsl.TraerRS("ClienteImpuestoTraer", sClien, True)
     i = 1
     cImpIva1 = 0
     cImpIva2 = 0
     Dim cLista As ClsComprobantesL
     Set cLista = New ClsComprobantesL
     If cLista.ListaConIva(cRsl.RegPorDefecto("ListaPrecio")) = False Then
'        Do While Not rAux.EOF
'           If i = 1 Then
'              cImpIva1 = cTotPar
'              LblIva1.Caption = Format(cImpIva1 * rAux!Porcentaje / 100, "#0.00")
'              i = 2
'           Else
'              cImpIva2 = cTotPar
'              LblIva2.Caption = Format(cImpIva2 * rAux!Porcentaje / 100, "#0.00")
'           End If
'           rAux.MoveNext
'        Loop
    End If
    ' Calcular Total General
    LblTotal.Caption = CCur(LblNeto.Caption)
    LblTotal.Caption = Format(LblTotal.Caption, "#0.00")
 End If
Exit Sub

errHandler:
     MsgBox Err.Description, vbCritical, "CalcularTotales"
 
End Sub
Private Sub TxtDetalle_Change()
  If sw1 = False Then
     CmdDetalle(0).Enabled = False
     If Len(TxtDetalle.Text) <> 0 Then
        CmdDetalle(0).Enabled = True
     End If
  End If
End Sub
Private Sub TxtPago_GotFocus()
  LblCompra.Caption = Format(LblTotal.Caption, "0.00")
  TxtPago.SelStart = 0
  TxtPago.SelLength = Len(TxtPago.Text)
End Sub
Private Sub TxtPago_LostFocus()
  If TxtPago.Text = "" Then TxtPago.Text = 0
  TxtPago.Text = Format(TxtPago.Text, "0.00")
  LblVuelto.Caption = Format(TxtPago.Text - LblCompra.Caption, "#0.00")
End Sub

Private Sub TxtProducto_LostFocus()
  Dim cRpl As ClsProductoL, nDigi As Byte, nLen As Byte, cPrec As Double, cPrecAux As Double
  On Error GoTo errHandler
  
  Set cRpl = New ClsProductoL
  Set cRsl = New ClsLectura
  If TxtProducto.Text <> "" Then
     If nDigito <> 0 Then
        nDigi = nDigito
     End If
     If cRpl.ProductoCompuesto(Left(TxtProducto.Text, nDigi)) = True Then
        TxtDetalle.Text = cRpl.BuscarNombreProducto(Left(TxtProducto.Text, nDigi))
        nLen = Len(TxtProducto.Text) - nDigi
        TxtPrecio.Text = Mid(TxtProducto.Text, nDigi + 1, nLen)
        cPrec = Format(TxtPrecio.Text / 1000, "#0.00")
        cPrecAux = Format(cRpl.TraerPrecio(cRsl.RegPorDefecto("ListaPrecio"), Left(TxtProducto.Text, nDigi)), "#0.00")
        If cPrecAux <> 0 Then
           TxtCantidad.Text = cPrec / cPrecAux
           TxtPrecio.Text = Format(cPrecAux, "#0.00")
        Else
           MsgBox "Producto sin Precio", vbCritical, "Atención"
        End If
     Else
         TxtDetalle.Text = cRpl.BuscarNombreProducto(TxtProducto.Text)
         If TxtDetalle.Text <> "" Then
            TxtPrecio.Text = Format(cRpl.TraerPrecio(cRsl.RegPorDefecto("ListaPrecio"), TxtProducto.Text), "#0.00")
         Else
            TxtDetalle.SetFocus
            Exit Sub
         End If
     End If
     TxtCantidad_LostFocus
     If TxtCantidad.Enabled = True Then
        If TxtPrecio.Text <> 0 Then
            NuevoRsD
            SendKeys "{ENTER}"
        Else
           TxtPrecio.SetFocus
       End If
     End If
  End If
  Set cRpl = Nothing
Exit Sub

errHandler:
      MsgBox Err.Description, vbCritical, "TxtProducto_LostFocus"
  
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
  On Error GoTo errHandler

  Dim cStock As ClsComprobantesL, cDepo As ClsProductoL
  Dim nDeposito As Integer
  
  Set cStock = New ClsComprobantesL
  Set cDepo = New ClsProductoL
  Set cRsl = New ClsLectura
  ' Si comprobante controla Stock, Verifica las Cantidades Disponibles
  If TxtProducto.Text <> "" Then
     If cStock.ActualizaStock(nCompro) = True Then
        If cStock.DepositoVenta <> 0 Then
           nDeposito = cStock.DepositoVenta
           If cDepo.MueveStock(TxtProducto.Text) = True Then
              If TxtCantidad.Text > cDepo.DameCantidad(TxtProducto.Text, nDeposito) Then
                 MsgBox "Cantidad Superior a la del Deposito de Venta", vbCritical, "Atención"
                 TxtCantidad.Text = cDepo.DameCantidad(TxtProducto.Text, nDeposito)
                 TxtCantidad.SetFocus
              End If
           End If
        Else
           MsgBox "No hay Deposito de Venta Definido", vbCritical, "Atención"
           TxtCantidad.SetFocus
        End If
     End If
  End If
  If TxtCantidad.Text = "" Then TxtCantidad.Text = 1
  TxtCantidad.Text = Format(TxtCantidad.Text, "#0.00")
 Exit Sub
errHandler:
      MsgBox Err.Description, vbCritical, "TxtCantidad_LostFocus"

End Sub
Private Sub TxtPrecio_GotFocus()
  TxtPrecio.SelStart = 0
  TxtPrecio.SelLength = Len(TxtPrecio.Text)
End Sub
Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimal(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtPrecio_LostFocus()
  If TxtPrecio.Text = "" Then TxtPrecio.Text = 0
  TxtPrecio.Text = Format(TxtPrecio.Text, "#0.00")
  If TxtPrecio.Text = 0 Then
     TxtPrecio.SetFocus
  End If
End Sub
Private Sub VerEstado()
  CmdBotones(0).Enabled = False
  If Rsd.RecordCount <> 0 Then
     CmdBotones(0).Enabled = True
  End If
End Sub
Private Sub Numero()
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4

  If sw = False Then
     Set cNum = New ClsComprobantesL
     nNum = Format(cNum.TraerUltimoNumero(nCompro), "00000000")
     nSuc = Format(cNum.TraerSucursalNumero(nCompro), "0000")
     LblNumero.Caption = nSuc & "-" & nNum
  End If
  Set cNum = Nothing
End Sub
Private Sub MSComm1_OnComm()
  On Error GoTo Errores

  Dim sStr As String
  If MSComm1.CommEvent = comEvReceive _
     And MSComm1.InBufferCount <> 0 Then
     Cadena = Cadena & MSComm1.Input
     If Asc(Right(Cadena, 1)) <> 10 Or Asc(Right(Cadena, 1)) <> 13 Then
        TxtProducto.Text = Left(Cadena, Len(Cadena) - 2)
        If ic = 1 Then
           ic = ic + 1
        Else
           ic = 1
           SendKeys "{TAB}"
           Cadena = ""
        End If
     End If
  ElseIf MSComm1.CommEvent <> comEvReceive Then
         MsgBox "Error en la comunicación " & MSComm1.CommEvent & _
                " durante la comunicación"
  End If
  Exit Sub
Errores:
  MsgBox Err.Description, vbCritical, "MsComm1_OnComm"
End Sub
Private Sub Imprimir()
 On Error GoTo Errores
 If nPuerto <> 0 Then
      'Imprimir en la controladora fiscal Factura B
    Principal.HASAR1.Puerto = nPuerto
    Principal.HASAR1.Comenzar
    Principal.HASAR1.DescripcionesLargas = True
    Principal.HASAR1.TratarDeCancelarTodo
    Principal.HASAR1.EspecificarNombreDeFantasia " ", " "
    Principal.HASAR1.AbrirComprobanteFiscal TICKET_C
    Rsd.MoveFirst
    Do While Not Rsd.EOF
       Dim cIva As Double, cInter As Double, cImpu As ClsComprobantesL, rImp As ADODB.Recordset
       Dim J As Byte
       Set cImpu = New ClsComprobantesL
       Set rImp = New ADODB.Recordset
       cIva = 21
       cInter = 0
       Set rImp = cImpu.ProductoImpuesto(Rsd!Producto)
       J = 1
       Do While Not rImp.EOF
          Select Case J
                 Case 1
                      If rImp!Interno = 1 Then
                         cInter = rImp!Porcentaje
                      Else
                         cIva = rImp!Porcentaje
                      End If
                      J = 2
                Case 2
                     If rImp!Interno = 1 Then
                        cInter = rImp!Porcentaje
                     Else
                        cIva = rImp!Porcentaje
                     End If
                     J = 1
             End Select
             rImp.MoveNext
          Loop
          Principal.HASAR1.ImprimirItem Left(Rsd!Descripcion, 25), Rsd!Cantidad, Rsd!PrecioUnitario, cIva, cInter
          Rsd.MoveNext
      Loop
      Principal.HASAR1.Subtotal False
      Principal.HASAR1.ImprimirPago "Su Pago :", TxtPago.Text
      Principal.HASAR1.CerrarComprobanteFiscal
      Principal.HASAR1.Finalizar
 End If
Exit Sub
Errores:
  MsgBox Err.Description, vbCritical, "Impimir"
End Sub
