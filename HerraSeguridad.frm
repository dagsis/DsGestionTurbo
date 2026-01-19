VERSION 5.00
Begin VB.Form HerraSeguridad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios y Permisos"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBotones 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   5
      Left            =   8310
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3975
      Width           =   1005
   End
   Begin VB.Frame Frame7 
      Caption         =   "Tablas y Maestros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   180
      TabIndex        =   36
      Top             =   2775
      Width           =   4290
      Begin VB.CheckBox ChkTabla 
         Caption         =   "Editar"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   52
         Top             =   675
         Width           =   855
      End
      Begin VB.CheckBox ChkTabla 
         Caption         =   "Borrar"
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   51
         Top             =   675
         Width           =   855
      End
      Begin VB.CheckBox ChkTabla 
         Caption         =   "Ver"
         Height          =   195
         Index           =   2
         Left            =   2055
         TabIndex        =   50
         Top             =   675
         Width           =   855
      End
      Begin VB.CheckBox ChkCliente 
         Caption         =   "Editar"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   49
         Top             =   1410
         Width           =   750
      End
      Begin VB.CheckBox ChkCliente 
         Caption         =   "Borrar"
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   48
         Top             =   1410
         Width           =   750
      End
      Begin VB.CheckBox ChkCliente 
         Caption         =   "Ver"
         Height          =   195
         Index           =   2
         Left            =   2055
         TabIndex        =   47
         Top             =   1410
         Width           =   750
      End
      Begin VB.CheckBox ChkCliente 
         Caption         =   "Informes"
         Height          =   195
         Index           =   3
         Left            =   3060
         TabIndex        =   46
         Top             =   1410
         Width           =   960
      End
      Begin VB.CheckBox ChkProductos 
         Caption         =   "Editar"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   45
         Top             =   2220
         Width           =   765
      End
      Begin VB.CheckBox ChkProductos 
         Caption         =   "Borrar"
         Height          =   240
         Index           =   1
         Left            =   1050
         TabIndex        =   44
         Top             =   2220
         Width           =   765
      End
      Begin VB.CheckBox ChkProductos 
         Caption         =   "Ver"
         Height          =   240
         Index           =   2
         Left            =   2055
         TabIndex        =   43
         Top             =   2220
         Width           =   765
      End
      Begin VB.CheckBox ChkProductos 
         Caption         =   "Informes"
         Height          =   240
         Index           =   3
         Left            =   3060
         TabIndex        =   42
         Top             =   2205
         Width           =   915
      End
      Begin VB.CheckBox ChkProveedores 
         Caption         =   "Editar"
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   41
         Top             =   3030
         Width           =   780
      End
      Begin VB.CheckBox ChkProveedores 
         Caption         =   "Borrar"
         Height          =   210
         Index           =   1
         Left            =   1050
         TabIndex        =   40
         Top             =   3030
         Width           =   780
      End
      Begin VB.CheckBox ChkProveedores 
         Caption         =   "Ver"
         Height          =   210
         Index           =   2
         Left            =   2055
         TabIndex        =   39
         Top             =   3030
         Width           =   780
      End
      Begin VB.CheckBox ChkProveedores 
         Caption         =   "Informes"
         Height          =   210
         Index           =   3
         Left            =   3060
         TabIndex        =   38
         Top             =   3015
         Width           =   945
      End
      Begin VB.CheckBox ChkTabla 
         Caption         =   "Informes"
         Height          =   195
         Index           =   3
         Left            =   3060
         TabIndex        =   37
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Tablas :"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   135
         TabIndex        =   56
         Top             =   315
         Width           =   660
      End
      Begin VB.Label Label9 
         Caption         =   "Clientes :"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   165
         TabIndex        =   55
         Top             =   1020
         Width           =   660
      End
      Begin VB.Label Label8 
         Caption         =   "Productos :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   135
         TabIndex        =   54
         Top             =   1785
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedores :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   165
         TabIndex        =   53
         Top             =   2610
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4545
      TabIndex        =   34
      Top             =   180
      Width           =   3600
      Begin VB.ListBox List5 
         Height          =   645
         Left            =   150
         TabIndex        =   35
         Top             =   210
         Width           =   3345
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Otros Permisos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4530
      TabIndex        =   24
      Top             =   4320
      Width           =   3660
      Begin VB.CheckBox Check1 
         Caption         =   "A.Precio"
         Height          =   210
         Left            =   105
         TabIndex        =   33
         Top             =   480
         Width           =   930
      End
      Begin VB.CheckBox Check2 
         Caption         =   "I.de Ventas"
         Height          =   210
         Left            =   1215
         TabIndex        =   32
         Top             =   1350
         Width           =   1245
      End
      Begin VB.CheckBox Check3 
         Caption         =   "I.de Caja"
         Height          =   210
         Left            =   2490
         TabIndex        =   31
         Top             =   1350
         Width           =   930
      End
      Begin VB.CheckBox Check4 
         Caption         =   "I.de Stock"
         Height          =   210
         Left            =   105
         TabIndex        =   30
         Top             =   1350
         Width           =   1035
      End
      Begin VB.CheckBox Check5 
         Caption         =   "A.Cobranzas"
         Height          =   240
         Left            =   1215
         TabIndex        =   29
         Top             =   885
         Width           =   1245
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Cobranzas"
         Height          =   240
         Left            =   105
         TabIndex        =   28
         Top             =   885
         Width           =   1065
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Cierres Caja"
         Height          =   210
         Left            =   1215
         TabIndex        =   27
         Top             =   480
         Width           =   1185
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Nota C/D"
         Height          =   210
         Left            =   2490
         TabIndex        =   26
         Top             =   480
         Width           =   1020
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Remitos"
         Height          =   240
         Left            =   2490
         TabIndex        =   25
         Top             =   885
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4545
      TabIndex        =   22
      Top             =   1200
      Width           =   3600
      Begin VB.ListBox List2 
         Height          =   645
         Left            =   150
         TabIndex        =   23
         Top             =   210
         Width           =   3345
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4545
      TabIndex        =   20
      Top             =   2235
      Width           =   3600
      Begin VB.ListBox List3 
         Height          =   645
         Left            =   150
         TabIndex        =   21
         Top             =   195
         Width           =   3345
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4545
      TabIndex        =   18
      Top             =   3240
      Width           =   3600
      Begin VB.ListBox List4 
         Height          =   645
         Left            =   120
         TabIndex        =   19
         Top             =   210
         Width           =   3390
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   810
      Index           =   0
      Left            =   8310
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   270
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   8310
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1196
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   8310
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2122
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   3
      Left            =   8310
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3048
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   4
      Left            =   8310
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5430
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   195
      TabIndex        =   8
      Top             =   180
      Width           =   4275
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1395
         Width           =   1965
      End
      Begin VB.TextBox TxtContra2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1785
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "TxtContra2"
         Top             =   2055
         Width           =   1020
      End
      Begin VB.TextBox TxtContra1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   165
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "TxtContra1"
         Top             =   2055
         Width           =   1020
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   315
         Left            =   180
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "TxtUsuario"
         Top             =   1395
         Width           =   1860
      End
      Begin VB.TextBox TxtNombre 
         Height          =   315
         Left            =   180
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "TxtNombre"
         Top             =   765
         Width           =   3915
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo :"
         Height          =   240
         Left            =   2175
         TabIndex        =   17
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Verificar  :"
         Height          =   180
         Left            =   1755
         TabIndex        =   14
         Top             =   1815
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "Contraseña :"
         Height          =   210
         Left            =   195
         TabIndex        =   13
         Top             =   1815
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario :"
         Height          =   210
         Left            =   180
         TabIndex        =   12
         Top             =   1140
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre :"
         Height          =   225
         Left            =   180
         TabIndex        =   11
         Top             =   525
         Width           =   750
      End
      Begin VB.Label LblCodigo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCodigo"
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
         Left            =   3255
         TabIndex        =   10
         Top             =   345
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   240
         Left            =   2655
         TabIndex        =   9
         Top             =   330
         Width           =   660
      End
   End
End
Attribute VB_Name = "HerraSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim sCod As Integer, sw As Boolean
'Dim Rs As ADODB.Recordset
'Dim cRsl As ClsLectura
'Dim cRsE As ClsEscritura
'Private Sub CargarCombos()
'  CmbTipo.AddItem "Administrador"
'  CmbTipo.ItemData(CmbTipo.NewIndex) = 1
'  CmbTipo.AddItem "Operador"
'  CmbTipo.ItemData(CmbTipo.NewIndex) = 2
'End Sub
'Private Sub Botones(bNue As Boolean, bBorr As Boolean, bModi As Boolean, bImp As Boolean, bSal As Boolean)
'  CmdBotones(0).Enabled = bNue
'  CmdBotones(1).Enabled = bBorr
'  CmdBotones(2).Enabled = bModi
'  CmdBotones(3).Enabled = bImp
'  CmdBotones(4).Enabled = bSal
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'     If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'        Exit Sub
'     End If
'     KeyAscii = Upper(KeyAscii)
'End Sub
'Private Sub Limpiar()
'  Dim ctl As Control
'  For Each ctl In Controls
'      If TypeOf ctl Is TextBox Then
'         ctl.Text = ""
'      End If
'  Next ctl
'  LblCodigo.Caption = ""
'  CmbTipo.ListIndex = -1
'End Sub
'
'Private Sub Habilitar(bEstado As Boolean, bList As Boolean)
'  Dim ctl As Control
'  For Each ctl In Controls
'      If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
'         Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
'         Or TypeOf ctl Is ComboBox Then
'         ctl.Enabled = bEstado
'      End If
'  Next ctl
'  List1.Enabled = bList
'End Sub
'
'Private Sub LinkearTexto()
'  Label3.Caption = Format(Rs!Caja, "00")
'  TxtDescripcion.Text = Rs!Descripcion
'End Sub
'
'Private Sub Form_Load()
' On Error GoTo errHandler
'
' Set Rs = New ADODB.Recordset
' Set cRsl = New ClsLectura
'
' sw = True
' Limpiar
' Habilitar False, False
'' Botones True, False, False, False, True
' CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
' CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
' CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
' CmdBotones(3).Picture = LoadResPicture("Valores", 0)
' CmdBotones(4).Picture = LoadResPicture("Salir", 0)
' CargarCombos
'' CargarLista
''  If Rs.EOF = False Or Rs.BOF = False Then
''    Botones True, True, True, True, True
''    Habilitar False, True
''    Rs.MoveFirst
''    LinkearTexto
'' End If
' Set cRsl = Nothing
'Exit Sub
'
'errHandler:
'   ManejaErrores
'
'End Sub
'
'Private Sub List1_Click()
' On Error GoTo errHandler
' Dim sCond As String, sCod As Integer
'
' Set cRsl = New ClsLectura
' sCod = List1.ItemData(List1.ListIndex)
' Set Rs = cRsl.TraerRS("CajaTraerUno", sCod, True)
' LinkearTexto
' Set cRsl = Nothing
'Exit Sub
'errHandler:
'   ManejaErrores
'End Sub
'Private Sub TxtDescripcion_Change()
'  If CmdBotones(0).Caption = "Grabar" Then
'     CmdBotones(0).Enabled = False
'     If Len(TxtDescripcion.Text) <> 0 Then
'        CmdBotones(0).Enabled = True
'     End If
'  Else
'     CmdBotones(2).Enabled = False
'     If Len(TxtDescripcion.Text) <> 0 Then
'        CmdBotones(2).Enabled = True
'     End If
'  End If
'End Sub
'Private Sub CmdBotones_Click(Index As Integer)
'  Select Case Index
'         Case 0
'              Nuevo
'         Case 1
'              Borrar
'         Case 2
'              Modificar
'         Case 3
'              Permisos
'         Case 4
'              Salir
'  End Select
'End Sub
'
'Private Sub Nuevo()
'  On Error GoTo Errores
'
'  If sw = True Then
'     sw = False
'     Botones False, False, False, False, True
'     Habilitar True, False
'     Limpiar
'     CmdBotones(0).Caption = "Grabar"
'     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
'     CmdBotones(4).Caption = "Cancelar"
'     CmdBotones(4).Picture = LoadResPicture("Cancelar", 0)
'     TxtDescripcion.SetFocus
'  Else
'     Botones True, True, True, True, True
'     Habilitar False, True
'     CmdBotones(0).Caption = "Nuevo"
'     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
'     CmdBotones(4).Caption = "Salir"
'     CmdBotones(4).Picture = LoadResPicture("Salir", 0)
'     sw = True
'     Grabar
'     CargarLista
'     List1.ListIndex = List1.ListCount - 1
'  End If
'Exit Sub
'Errores:
'   ManejaErrores
'End Sub
'
'Private Sub Salir()
' On Error GoTo errHandler
'
'  Botones True, False, False, False, True
'  If sw = False Then
'     Limpiar
'     If CmdBotones(0).Caption = "Grabar" Then
'        sw = True
'        Habilitar False, True
'        CmdBotones(0).Caption = "Nuevo"
'        CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
'        CmdBotones(4).Caption = "Salir"
'        CmdBotones(4).Picture = LoadResPicture("Salir", 0)
'     Else
'        sw = True
'        Habilitar False, True
'        CmdBotones(2).Caption = "Modificar"
'        CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
'        CmdBotones(4).Caption = "Salir"
'        CmdBotones(4).Picture = LoadResPicture("Salir", 0)
'     End If
'      If Rs.EOF = False Or Rs.BOF = False Then
'        Botones True, True, True, True, True
'        LinkearTexto
'     End If
'  Else
'     Unload Me
'  End If
'  Exit Sub
'errHandler:
'   ManejaErrores
'End Sub
'Private Sub Modificar()
'  On Error GoTo Errores
'
'  Static rList As Integer
'
'  If sw = True Then
'     sw = False
''     rList = List1.ListIndex
'     Botones False, False, True, False, True
'     Habilitar True, False
'     CmdBotones(2).Caption = "Grabar"
'     CmdBotones(2).Picture = LoadResPicture("Grabar", 0)
'     CmdBotones(4).Caption = "Cancelar"
'     CmdBotones(4).Picture = LoadResPicture("Cancelar", 0)
'  Else
'     Botones True, True, True, True, True
'     Habilitar False, True
'     Grabar
'     CmdBotones(2).Caption = "Modificar"
'     CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
'     CmdBotones(4).Caption = "Salir"
'     CmdBotones(4).Picture = LoadResPicture("Salir", 0)
'     sw = True
'     CargarLista
''     List1.ListIndex = rList
'  End If
'Exit Sub
'Errores:
'   ManejaErrores
'
'End Sub
'Private Sub Borrar()
'  On Error GoTo errHandler
'
'  Dim nList As Integer
'
'  Set cRsE = New ClsEscritura
'
'  nList = List1.ListIndex
'  If MsgBox("Borra La Caja Seleccionada ?", 20, "Atención") = 6 Then
'     sCod = Label3.Caption
'     cRsE.Actualizar "CajaBorrar", Rs, Val(Label3.Caption), "Caja"
'     CargarLista
'     Limpiar
'     Botones True, False, False, False, True
'      If Rs.EOF = False Or Rs.BOF = False Then
'        Botones True, True, True, True, True
'        If nList = 0 Then nList = 1
'        List1.ListIndex = nList - 1
'        LinkearTexto
'     End If
'  End If
'  Exit Sub
'errHandler:
'  ManejaErrores
'End Sub
'Private Sub Permisos()
'  HerraPermisos.Show 1
'End Sub
'Private Sub CargarLista()
'  On Error GoTo errHandler
'
'  Set cRsl = New ClsLectura
'  Set Rs = New ADODB.Recordset
'
'  List1.Clear
'  Set Rs = cRsl.TraerRS("CajaTraerTodos")
'
'  Do While Not Rs.EOF
'     List1.AddItem Rs!Descripcion
'     List1.ItemData(List1.NewIndex) = Rs!Caja
'     Rs.MoveNext
'  Loop
'  Set cRsl = Nothing
'  Exit Sub
'errHandler:
'  ManejaErrores
'End Sub
'
'Private Sub Grabar()
'  On Error GoTo errHandler
'
'  Dim Rs1 As ADODB.Recordset
'  Set Rs1 = New ADODB.Recordset
'
'  Rs1.Fields.Append "Descripcion", adVarChar, 40
'  Rs1.Open
'  Rs1.AddNew
'  Rs1!Descripcion = TxtDescripcion.Text
'  Rs1.Update
'
'  Set cRsE = New ClsEscritura
'
'  If CmdBotones(2).Caption = "Grabar" Then
'     cRsE.Actualizar "CajaModif", Rs1, Val(Label3.Caption)
'  Else
'     cRsE.Actualizar "CajaAgregar", Rs1
'  End If
'  Rs1.Close
'  Set cRsE = Nothing
'  Exit Sub
'errHandler:
'  ManejaErrores
'End Sub
'
'
'
'
