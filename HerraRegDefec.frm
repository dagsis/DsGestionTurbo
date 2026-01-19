VERSION 5.00
Begin VB.Form HerraRegDefec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros por defecto"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   9945
   Begin VB.ComboBox CmbTabla 
      Height          =   315
      Index           =   3
      Left            =   390
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1635
      Width           =   2295
   End
   Begin VB.ComboBox CmbUser 
      Height          =   315
      Left            =   1125
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   135
      Width           =   2655
   End
   Begin VB.ComboBox CmbTabla 
      Height          =   315
      Index           =   18
      Left            =   5175
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3435
      Width           =   2295
   End
   Begin VB.ComboBox CmbTabla 
      Height          =   315
      Index           =   2
      Left            =   5175
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1035
      Width           =   2295
   End
   Begin VB.ComboBox CmbTabla 
      Height          =   315
      Index           =   1
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1035
      Width           =   2295
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   7245
      TabIndex        =   39
      Top             =   8295
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8580
      TabIndex        =   0
      Top             =   8295
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comprobantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   195
      TabIndex        =   35
      Top             =   6030
      Width           =   9615
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   25
         Left            =   7260
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1680
         Width           =   2235
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   24
         Left            =   7260
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1080
         Width           =   2235
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   17
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   22
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1680
         Width           =   2325
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   23
         Left            =   7260
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   21
         Left            =   4875
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   20
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   14
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   15
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   16
         Left            =   4875
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label28 
         Caption         =   "Venta - Otra Condición :"
         Height          =   255
         Left            =   7260
         TabIndex        =   65
         Top             =   1425
         Width           =   2055
      End
      Begin VB.Label Label27 
         Caption         =   "Venta - Resp.Inscr.:"
         Height          =   255
         Left            =   7260
         TabIndex        =   62
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label18 
         Caption         =   "Cliente - Cons.Final :"
         Height          =   255
         Left            =   180
         TabIndex        =   61
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label25 
         Caption         =   "Cond. de Pago:"
         Height          =   210
         Left            =   180
         TabIndex        =   60
         Top             =   1455
         Width           =   1395
      End
      Begin VB.Label Label26 
         Caption         =   "Venta - Cons.Final N :"
         Height          =   255
         Left            =   7260
         TabIndex        =   59
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "Cobranza Por Defecto :"
         Height          =   240
         Left            =   4875
         TabIndex        =   53
         Top             =   840
         Width           =   2160
      End
      Begin VB.Label Label23 
         Caption         =   "Cobranza Sobre Ventas :"
         Height          =   240
         Left            =   2535
         TabIndex        =   51
         Top             =   840
         Width           =   2100
      End
      Begin VB.Label Label17 
         Caption         =   "Venta - Cons.Final :"
         Height          =   255
         Left            =   4875
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Venta - Facturación :"
         Height          =   255
         Left            =   2535
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Cobro por Defecto :"
         Height          =   255
         Left            =   195
         TabIndex        =   36
         Top             =   855
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tablas Comunes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5385
      Left            =   150
      TabIndex        =   20
      Top             =   555
      Width           =   9750
      Begin VB.CheckBox ChkAnulaPres 
         Caption         =   "Anula Presupuesto"
         Height          =   195
         Left            =   2625
         TabIndex        =   83
         Top             =   3930
         Width           =   2205
      End
      Begin VB.CheckBox ChkRedondea 
         Caption         =   "Redondea Precios"
         Height          =   195
         Left            =   255
         TabIndex        =   82
         Top             =   3915
         Width           =   2205
      End
      Begin VB.CheckBox ChkDescCaja 
         Caption         =   "Desc.Caja Rapida"
         Height          =   195
         Left            =   7245
         TabIndex        =   81
         Top             =   3615
         Width           =   1860
      End
      Begin VB.CheckBox ChkAutoNumera 
         Caption         =   "Autonumera Cli./Prov."
         Height          =   210
         Left            =   7245
         TabIndex        =   80
         Top             =   3300
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Restaurant y Delivery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   210
         TabIndex        =   75
         Top             =   4305
         Width           =   9420
         Begin VB.ComboBox CmbTabla 
            Height          =   315
            Index           =   29
            Left            =   7335
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   465
            Width           =   1965
         End
         Begin VB.ComboBox CmbTabla 
            Height          =   315
            Index           =   28
            Left            =   4935
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   465
            Width           =   2295
         End
         Begin VB.ComboBox CmbTabla 
            Height          =   315
            Index           =   27
            Left            =   2550
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   465
            Width           =   2295
         End
         Begin VB.ComboBox CmbTabla 
            Height          =   315
            Index           =   26
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   465
            Width           =   2295
         End
         Begin VB.Label Label32 
            Caption         =   "C.C. Delivery :"
            Height          =   210
            Left            =   7335
            TabIndex        =   79
            Top             =   210
            Width           =   1230
         End
         Begin VB.Label Label31 
            Caption         =   "C.C.  Restaurant :"
            Height          =   195
            Left            =   4935
            TabIndex        =   78
            Top             =   210
            Width           =   1560
         End
         Begin VB.Label Label30 
            Caption         =   "Lista Delivery"
            Height          =   225
            Left            =   2550
            TabIndex        =   77
            Top             =   195
            Width           =   1185
         End
         Begin VB.Label Label29 
            Caption         =   "Lista C.Rápida/Mesas :"
            Height          =   225
            Left            =   150
            TabIndex        =   76
            Top             =   210
            Width           =   1845
         End
      End
      Begin VB.CheckBox ChkCaja 
         Caption         =   "Mueve Caja"
         Height          =   255
         Left            =   7440
         TabIndex        =   74
         Top             =   2835
         Width           =   2055
      End
      Begin VB.CheckBox ChkMueveStock 
         Caption         =   "Mueve Stock"
         Height          =   255
         Left            =   7440
         TabIndex        =   73
         Top             =   1431
         Width           =   1695
      End
      Begin VB.CheckBox ChkVenta 
         Caption         =   "Venta Segun Condición"
         Height          =   255
         Left            =   7440
         TabIndex        =   72
         Top             =   2484
         Width           =   2055
      End
      Begin VB.CheckBox ChkComanda 
         Caption         =   "Recibos sin Comanderas"
         Height          =   255
         Left            =   7440
         TabIndex        =   71
         Top             =   2133
         Width           =   2055
      End
      Begin VB.CheckBox ChkStock 
         Caption         =   "Stock Negativo"
         Height          =   255
         Left            =   7440
         TabIndex        =   70
         Top             =   1782
         Width           =   1695
      End
      Begin VB.CheckBox chkTrans 
         Caption         =   "C.Transferencias"
         Height          =   255
         Left            =   7440
         TabIndex        =   69
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox ChkBAvanzada 
         Caption         =   "Busqueda Avanzada"
         Height          =   210
         Left            =   5355
         TabIndex        =   68
         Top             =   3615
         Width           =   1935
      End
      Begin VB.CheckBox ChkModoSuper 
         Caption         =   "Modo Supermercado"
         Height          =   225
         Left            =   5355
         TabIndex        =   67
         Top             =   3300
         Width           =   2100
      End
      Begin VB.CheckBox ChkDirecto 
         Caption         =   "Imprime Directo el Comprobante"
         Height          =   210
         Left            =   2625
         TabIndex        =   66
         Top             =   3615
         Width           =   2925
      End
      Begin VB.CheckBox ChkNumera 
         Caption         =   "Sin Barra de Herramienta"
         Height          =   210
         Left            =   2625
         TabIndex        =   55
         Top             =   3300
         Width           =   2565
      End
      Begin VB.CheckBox ChkActualizoVentas 
         Caption         =   "Actualizo Precios Ventas"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   3300
         Width           =   2205
      End
      Begin VB.CheckBox ChkActualizo 
         Caption         =   "Actualizo Precios Compras"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   3615
         Width           =   2205
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   19
         Left            =   7395
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   13
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   12
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2876
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   11
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   10
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   9
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   8
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   7
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   6
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   5
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   4
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox CmbTabla 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label22 
         Caption         =   "Descuentos :"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1425
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Moneda :"
         Height          =   255
         Index           =   0
         Left            =   7410
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   5040
         TabIndex        =   40
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Zona :"
         Height          =   255
         Left            =   5040
         TabIndex        =   34
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Vendedores :"
         Height          =   255
         Left            =   5040
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "U.de Medida :"
         Height          =   255
         Left            =   5040
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Provincia :"
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Iva :"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Marca :"
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Localidades :"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "L.dePrecio :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Impuestos :"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Familia :"
         Height          =   255
         Left            =   2640
         TabIndex        =   25
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Depósito :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Descuento :"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1221
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Caja :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Banco :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label20 
      Caption         =   "Usuario :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   45
      Top             =   135
      Width           =   840
   End
End
Attribute VB_Name = "HerraRegDefec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura, rDef As ADODB.Recordset, bPrime As Boolean

Private Sub ChkActualizo_Click()
  TodoBien
End Sub

Private Sub ChkActualizoVentas_Click()
  TodoBien
End Sub

Private Sub ChkAnulaPres_Click()
  TodoBien
End Sub

Private Sub ChkAutoNumera_Click()
  TodoBien
End Sub

Private Sub ChkBAvanzada_Click()
  TodoBien
End Sub

Private Sub ChkCaja_Click()
  TodoBien
End Sub

Private Sub ChkComanda_Click()
  TodoBien
End Sub

Private Sub ChkDescCaja_Click()
  TodoBien
End Sub

Private Sub ChkModoSuper_Click()
  TodoBien
End Sub

Private Sub ChkMueveStock_Click()
  TodoBien
End Sub

Private Sub ChkNumera_Click()
  TodoBien
End Sub

Private Sub ChkRedondea_Click()
  TodoBien
End Sub

Private Sub ChkStock_Click()
  TodoBien
End Sub

Private Sub chkTrans_Click()
  TodoBien
End Sub

Private Sub ChkVenta_Click()
  TodoBien
End Sub

Private Sub CmbTabla_Click(Index As Integer)
  TodoBien
End Sub

Private Sub CmbUser_Click()
  CagarDefectos
End Sub

Private Sub CmdAplicar_Click()
  Dim cRle As ClsEscritura
  Dim i As Integer
  Dim pPar(29) As Long, pByte(15) As Byte
  Dim pPar21 As Byte, pPar22 As Byte, pPar26 As Byte
  
  For i = 0 To 29
      If CmbTabla(i).ListIndex <> -1 Then
        pPar(i) = CmbTabla(i).ItemData(CmbTabla(i).ListIndex)
      End If
  Next i
  
  pByte(0) = ChkActualizo.Value
  pByte(1) = ChkActualizoVentas.Value
  pByte(2) = ChkNumera.Value
  pByte(3) = ChkDirecto.Value
  pByte(4) = ChkModoSuper.Value
  pByte(5) = ChkBAvanzada.Value
  pByte(6) = chkTrans.Value
  pByte(7) = ChkStock.Value
  pByte(8) = ChkComanda.Value
  pByte(9) = ChkVenta.Value
  pByte(10) = ChkMueveStock.Value
  pByte(11) = ChkCaja.Value
  pByte(12) = ChkAutoNumera.Value
  pByte(13) = ChkDescCaja.Value
  pByte(14) = ChkRedondea.Value
  pByte(15) = ChkAnulaPres.Value
    
  Set cRle = New ClsEscritura
  
  cRle.ActualizarRegistro CmbUser.ItemData(CmbUser.ListIndex), pPar(), pByte()
  
  Botones False, True
  MsgBox "Deberá Reiniciar el Sistema para Aplicar los Cambios", vbInformation, "Atención"
  
  End
  
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub
Private Sub Form_Load()
  Set rDef = New ADODB.Recordset
  bPrime = False
  CargarCombos
  bPrime = True
  CagarDefectos
  
  Botones False, True
  Me.Top = 0
  Me.Left = 0

End Sub

Private Sub Botones(bApli As Boolean, bSal As Boolean)
  CmdAplicar.Enabled = bApli
  CmdSalir.Enabled = bSal
End Sub

Private Sub CargarCombos()
  Dim i As Long, sCond As String
  
  On Error GoTo errHandler

  If nCatUser = 0 Then sCond = ""
  If nCatUser = 1 Then sCond = "Categoria<>0"
  If nCatUser = 2 Then sCond = "Categoria=2 and Id_Usuario=" & nUsuario
  
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbUser, "Usuarios", "Id_Usuario", "Usuario", sCond
  CmbUser.ListIndex = 0
  
  cRsl.CargaCombo CmbTabla(0), "Bancos", "Banco", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(1), "Marcas", "Marca", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(2), "UnidadMedida", "Unidad", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(3), "Cajas", "Caja", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(4), "Provincias", "Provincia", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(5), "Vendedores", "Vendedor", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(6), "Descuentos", "Descuento", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(7), "Impuestos", "Impuesto", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(8), "Zonas", "Zona", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(9), "Depositos", "Deposito", "Descripcion", "Suc=" & nSucursal
  cRsl.CargaCombo CmbTabla(10), "Familias", "Familia", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(11), "TipoIva", "TipoIva", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(12), "ListadePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(13), "Localidades", "Localidad", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(14), "CondVenta", "CondVta", "Descripcion", "Tipo<6"
  If nCatUser < 2 Then
     cRsl.CargaCombo CmbTabla(15), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
     cRsl.CargaCombo CmbTabla(16), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
     cRsl.CargaCombo CmbTabla(23), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
     cRsl.CargaCombo CmbTabla(24), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
     cRsl.CargaCombo CmbTabla(25), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & nUsuario
  Else
     cRsl.CargaCombo CmbTabla(15), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
     cRsl.CargaCombo CmbTabla(16), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
     cRsl.CargaCombo CmbTabla(23), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
     cRsl.CargaCombo CmbTabla(24), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
     cRsl.CargaCombo CmbTabla(25), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
  End If
  cRsl.CargaCombo CmbTabla(17), "Clientes", "Id", "RazonSocial", ""
  cRsl.CargaCombo CmbTabla(18), "Proveedores", "Id", "RazonSocial", ""
  cRsl.CargaCombo CmbTabla(19), "Monedas", "Id", "Descripcion", ""
  
  CmbTabla(20).AddItem "Aplica Automaticamente"
  CmbTabla(20).ItemData(CmbTabla(20).NewIndex) = 0
  CmbTabla(20).AddItem "Pide Forma de Pago"
  CmbTabla(20).ItemData(CmbTabla(20).NewIndex) = 1
  CmbTabla(20).AddItem "Aplica a la Caja"
  CmbTabla(20).ItemData(CmbTabla(20).NewIndex) = 2
  
  
  cRsl.CargaCombo CmbTabla(21), "Comprobantes", "Id", "Descripcion", "TipoOperacion=4"
 
  CmbTabla(22).AddItem "CUENTA CORRIENTE"
  CmbTabla(22).ItemData(CmbTabla(22).NewIndex) = 0
  CmbTabla(22).AddItem "CONTADO"
  CmbTabla(22).ItemData(CmbTabla(22).NewIndex) = 1
  
  cRsl.CargaCombo CmbTabla(26), "ListadePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(27), "ListadePrecio", "ListaPrecio", "Descripcion", ""
  cRsl.CargaCombo CmbTabla(28), "PlanDeCuenta", "Id", "Descripcion", "Grupo=4"
  cRsl.CargaCombo CmbTabla(29), "PlanDeCuenta", "Id", "Descripcion", "Grupo=4"



  Set cRsl = Nothing

Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CagarDefectos()
  Dim sTextCob  As String
  
  On Error GoTo errHandler

  If bPrime = True Then
      Set cRsl = New ClsLectura
            
      cRsl.CargaCombo CmbTabla(15), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
      cRsl.CargaCombo CmbTabla(16), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
      cRsl.CargaCombo CmbTabla(23), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
      cRsl.CargaCombo CmbTabla(24), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
      cRsl.CargaCombo CmbTabla(25), "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Valor=1 and Usuario=" & CmbUser.ItemData(CmbUser.ListIndex)
      
      
      Set rDef = cRsl.TraerDefectos(CmbUser.ItemData(CmbUser.ListIndex))
      If rDef!Banco <> 0 Then
         CmbTabla(0).text = cRsl.DatoCombo("Bancos", "Banco", "Descripcion", rDef!Banco)
      End If
      If rDef!Marca <> 0 Then
         CmbTabla(1).text = cRsl.DatoCombo("Marcas", "Marca", "Descripcion", rDef!Marca)
      End If
      If rDef!Medida <> 0 Then
         CmbTabla(2).text = cRsl.DatoCombo("UnidadMedida", "Unidad", "Descripcion", rDef!Medida)
      End If
      If rDef!Caja <> 0 Then
         CmbTabla(3).text = cRsl.DatoCombo("Cajas", "Caja", "Descripcion", rDef!Caja)
      End If
      If Not IsNull(rDef!Provincia) Then
         CmbTabla(4).text = cRsl.DatoCombo("Provincias", "Provincia", "Descripcion", rDef!Provincia)
      End If
      If Not IsNull(rDef!Vendedor) Then
         CmbTabla(5).text = cRsl.DatoCombo("Vendedores", "Vendedor", "Descripcion", rDef!Vendedor)
      End If
      If Not IsNull(rDef!Descuento) Then
         CmbTabla(6).text = cRsl.DatoCombo("Descuentos", "Descuento", "Descripcion", rDef!Descuento)
      End If
      If Not IsNull(rDef!Pais) Then
         CmbTabla(7).text = cRsl.DatoCombo("Impuestos", "Impuesto", "Descripcion", rDef!Pais)
      End If
      If Not IsNull(rDef!Zona) Then
         CmbTabla(8).text = cRsl.DatoCombo("Zonas", "Zona", "Descripcion", rDef!Zona)
      End If
      If Not IsNull(rDef!Deposito) Then
         CmbTabla(9).text = cRsl.DatoCombo("Depositos", "Deposito", "Descripcion", rDef!Deposito)
      End If
      If Not IsNull(rDef!Familia) Then
         CmbTabla(10).text = cRsl.DatoCombo("Familias", "Familia", "Descripcion", rDef!Familia)
      End If
      If Not IsNull(rDef!Impuesto) Then
         CmbTabla(11).text = cRsl.DatoCombo("TipoIva", "TipoIva", "Descripcion", rDef!Impuesto)
      End If
      If Not IsNull(rDef!ListaPrecio) Then
         CmbTabla(12).text = cRsl.DatoCombo("ListaDePrecio", "ListaPrecio", "Descripcion", rDef!ListaPrecio)
      End If
      If Not IsNull(rDef!Localidad) Then
         CmbTabla(13).text = cRsl.DatoCombo("Localidades", "Localidad", "Descripcion", rDef!Localidad)
      End If
      If Not IsNull(rDef!Cobro) Then
         CmbTabla(14).text = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", rDef!Cobro)
      End If
      If Not IsNull(rDef!VentaFactura) Then
         CmbTabla(15).text = cRsl.DatoCombo("UsuariosComprobantesVentas", "Comprobante", "Descripcion", rDef!VentaFactura)
      End If
      If Not IsNull(rDef!VentaConsFinal) Then
         CmbTabla(16).text = cRsl.DatoCombo("UsuariosComprobantesVentas", "Comprobante", "Descripcion", rDef!VentaConsFinal)
      End If
      If Not IsNull(rDef!ClienteConsFinal) Then
         CmbTabla(17).text = cRsl.DatoCombo("Clientes", "Id", "RazonSocial", rDef!ClienteConsFinal)
      End If
      If Not IsNull(rDef!Proveedor) Then
         CmbTabla(18).text = cRsl.DatoCombo("Proveedores", "Id", "RazonSocial", rDef!Proveedor)
      End If
      If Not IsNull(rDef!Moneda) Then
         CmbTabla(19).text = cRsl.DatoCombo("Monedas", "Id", "Descripcion", rDef!Moneda)
      End If
      ChkActualizo.Value = IIf(rDef!precio = True, 1, 0)
      ChkActualizoVentas.Value = IIf(rDef!Venta = True, 1, 0)
      ChkNumera.Value = IIf(rDef!NumeraContr = True, 1, 0)
      ChkDirecto.Value = IIf(rDef!ImprimeDirecto = True, 1, 0)
      ChkModoSuper.Value = IIf(cRsl.TraerValorDeUnCampo("Sistema", "Pagare", "") = True, 1, 0)
      ChkBAvanzada.Value = IIf(cRsl.TraerValorDeUnCampo("Sistema", "bPagare", "") = True, 1, 0)
      chkTrans.Value = IIf(cRsl.TraerValorDeUnCampo("Sistema", "bRecibo", "") = True, 1, 0)
      ChkStock.Value = IIf(cRsl.TraerValorDeUnCampo("Sistema", "bRapido", "") = True, 1, 0)
      ChkComanda.Value = cRsl.TraerValorDeUnCampo("Sistema", "Receta", "")
      ChkVenta.Value = IIf(cRsl.TraerValorDeUnCampo("Registros", "TipoVenta", "") = True, 1, 0)
      ChkMueveStock.Value = IIf(cRsl.TraerValorDeUnCampo("Sistema", "MueveStock", "") = True, 1, 0)
      ChkCaja.Value = IIf(cRsl.TraerValorDeUnCampo("Sistema", "MueveCaja", "") = True, 1, 0)
      ChkAutoNumera.Value = IIf(rDef!CliProvAut = True, 1, 0)
      ChkDescCaja.Value = IIf(rDef!DescCaja = True, 1, 0)
      ChkRedondea.Value = IIf(rDef!Redondea = True, 1, 0)
      ChkAnulaPres.Value = IIf(rDef!AnulaPres = True, 1, 0)
      
      If Not IsNull(rDef!FormaCob) Then
         Select Case rDef!FormaCob
                Case 0
                     sTextCob = "Aplica Automaticamente"
                Case 1
                     sTextCob = "Pide Forma de Pago"
                Case 2
                     sTextCob = "Aplica a la caja"
         End Select
         CmbTabla(20).text = sTextCob
      End If
      
      If Not IsNull(rDef!Cobranza) Then
         CmbTabla(21).text = cRsl.DatoCombo("Comprobantes", "Id", "Descripcion", rDef!Cobranza)
      End If
      
      If Not IsNull(rDef!CPago) Then
         Select Case rDef!CPago
                Case 0
                     sTextCob = "CUENTA CORRIENTE"
                Case 1
                     sTextCob = "CONTADO"
         End Select
         CmbTabla(22).text = sTextCob
      End If
      

      If Not IsNull(rDef!VentaConsFinalN) Then
         CmbTabla(23).text = cRsl.DatoCombo("UsuariosComprobantesVentas", "Comprobante", "Descripcion", rDef!VentaConsFinalN)
      End If

      If Not IsNull(rDef!VentaResponsable) Then
         CmbTabla(24).text = cRsl.DatoCombo("UsuariosComprobantesVentas", "Comprobante", "Descripcion", rDef!VentaResponsable)
      End If

      If Not IsNull(rDef!VentaOtraCondicion) Then
         CmbTabla(25).text = cRsl.DatoCombo("UsuariosComprobantesVentas", "Comprobante", "Descripcion", rDef!VentaOtraCondicion)
      End If

      If Not IsNull(rDef!ListaRes) Then
         CmbTabla(26).text = cRsl.DatoCombo("ListadePrecio", "ListaPrecio", "Descripcion", rDef!ListaRes)
      End If

      If Not IsNull(rDef!ListaDel) Then
         CmbTabla(27).text = cRsl.DatoCombo("ListadePrecio", "ListaPrecio", "Descripcion", rDef!ListaDel)
      End If

      If Not IsNull(rDef!CuentaRes) Then
         CmbTabla(28).text = cRsl.DatoCombo("PlandeCuenta", "Id", "Descripcion", rDef!CuentaRes)
      End If

      If Not IsNull(rDef!CuentaDel) Then
         CmbTabla(29).text = cRsl.DatoCombo("PlandeCuenta", "Id", "Descripcion", rDef!CuentaDel)
      End If

  End If
  Exit Sub

errHandler:
   If err.Number = 383 Then
      Resume Next
   Else
      MsgBox err.Description
   End If
   

End Sub

Private Sub TodoBien()
  Botones True, True
End Sub
