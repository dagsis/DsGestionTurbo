VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaCobAuto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobranza Automática"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10470
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   720
      Index           =   2
      Left            =   9495
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Valores"
      Height          =   720
      Index           =   1
      Left            =   8565
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Grabar"
      Height          =   720
      Index           =   0
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuenta Corriente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5400
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   10290
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Detalle"
         Height          =   285
         Left            =   7605
         TabIndex        =   14
         Top             =   1080
         Width           =   810
      End
      Begin VB.TextBox TxtEfectivo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3615
         TabIndex        =   12
         Text            =   "TxtEfectivo"
         Top             =   1455
         Width           =   915
      End
      Begin VB.ComboBox CmbCaja 
         Height          =   315
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1470
         Width           =   2190
      End
      Begin VB.ComboBox CmbComprobante 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   270
         Width           =   2445
      End
      Begin VB.CommandButton CmdClientes 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   7605
         TabIndex        =   9
         Top             =   735
         Width           =   810
      End
      Begin VB.TextBox TxtCliente 
         Height          =   315
         Left            =   1110
         TabIndex        =   8
         Text            =   "TxtCliente"
         Top             =   750
         Width           =   735
      End
      Begin VB.TextBox TxtCancelar 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5790
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "TxtCancela"
         Top             =   1455
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "Aplicar"
         Height          =   285
         Left            =   6675
         TabIndex        =   6
         Top             =   1455
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   8610
         TabIndex        =   2
         Top             =   225
         Width           =   1410
         Begin VB.OptionButton Option1 
            Caption         =   "Normal"
            Height          =   225
            Left            =   165
            TabIndex        =   5
            Top             =   300
            Width           =   825
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Dar Vuelto"
            Height          =   210
            Left            =   165
            TabIndex        =   4
            Top             =   675
            Width           =   1125
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Adelanto"
            Height          =   195
            Left            =   165
            TabIndex        =   3
            Top             =   1065
            Width           =   1065
         End
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   315
         Left            =   7140
         TabIndex        =   15
         Top             =   270
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   707395585
         CurrentDate     =   36821
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   3000
         Left            =   180
         TabIndex        =   16
         Top             =   1965
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5292
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
      Begin VB.Label LblCancelado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCancelado"
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
         Left            =   7155
         TabIndex        =   32
         Top             =   5040
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "A Cobrar :"
         Height          =   195
         Left            =   2835
         TabIndex        =   31
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label LblSeleccionado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblSeleccionado"
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
         Left            =   1365
         TabIndex        =   30
         Top             =   5025
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Seleccionado :"
         Height          =   210
         Left            =   240
         TabIndex        =   29
         Top             =   5025
         Width           =   1065
      End
      Begin VB.Label LblImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblImporte"
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
         Left            =   4995
         TabIndex        =   28
         Top             =   5040
         Width           =   1125
      End
      Begin VB.Label LblCobrado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCobrado"
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
         Left            =   6150
         TabIndex        =   27
         Top             =   5040
         Width           =   960
      End
      Begin VB.Label LblSaldo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblSaldo"
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
         Left            =   8355
         TabIndex        =   26
         Top             =   5040
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   "Totales :"
         Height          =   240
         Left            =   4230
         TabIndex        =   25
         Top             =   5040
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "Caja :"
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   1455
         Width           =   540
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   6345
         TabIndex        =   23
         Top             =   270
         Width           =   630
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
         Height          =   315
         Left            =   4215
         TabIndex        =   22
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Nº :"
         Height          =   225
         Left            =   3840
         TabIndex        =   21
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label3 
         Caption         =   "Comprobante :"
         Height          =   240
         Left            =   150
         TabIndex        =   20
         Top             =   270
         Width           =   1155
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
         Height          =   315
         Left            =   1935
         TabIndex        =   19
         Top             =   750
         Width           =   5505
      End
      Begin VB.Label Label1 
         Caption         =   "Clientes :"
         Height          =   240
         Left            =   150
         TabIndex        =   18
         Top             =   735
         Width           =   900
      End
      Begin VB.Label Label10 
         Caption         =   "Sdo a Cancelar :"
         Height          =   210
         Left            =   4560
         TabIndex        =   17
         Top             =   1455
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7425
         Y1              =   1335
         Y2              =   1335
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   720
      Index           =   3
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Si el pago es solamente en Efectivo grabar sin valores."
      Height          =   420
      Left            =   165
      TabIndex        =   35
      Top             =   5550
      Width           =   2685
   End
End
Attribute VB_Name = "VentaCobAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura
Dim sw As Boolean, cImporte As Currency
Dim rReci As ADODB.Recordset, rCaja As ADODB.Recordset
Dim rValo As ADODB.Recordset, rApli As ADODB.Recordset
Dim rAde As ADODB.Recordset

Public cSaldo As Currency
Public RsC As ADODB.Recordset
Public nCas As Integer, sNumero As String
Private Sub CmbComprobante_Click()
  Dim cNum As ClsLectura, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler

  If sw = False Then
     Set cNum = New ClsLectura
     nNum = Format(cNum.TraerValorDeUnCampo("Comprobantes", "Numero", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
     nSuc = Format(cNum.TraerValorDeUnCampo("Comprobantes", "Sucursal", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
     LblNumero.Caption = nSuc & "-" & nNum
  End If
  Set cNum = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CmdAplicar_Click()
'  TxtEfectivo.Text = TxtEfectivo.Text - TxtCancelar.Text
'  LblCancelado.Caption = CDbl(LblCancelado.Caption) + CDbl(TxtCancelar.Text)
'  RsC!Cancelado = TxtCancelar.Text
'  RsC!Cobrar = RsC!Importe - RsC!Cobrado - RsC!Cancelado
'  RsC.Update
'  TxtEfectivo.Text = Format(TxtEfectivo.Text, "#0.00")
'  TxtCancelar.Text = Format(0, "#0.00")
'  LblCancelado.Caption = Format(LblCancelado, "#0.00")
   CmdBotones(0).Enabled = True
End Sub

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              GrabarTodo
              CrearRs
              LlenarGrilla
              CmdBotones(3).Enabled = True
              CmdBotones(3).SetFocus
         Case 1
              Valor
         Case 2
              Salir
         Case 3
              Imprimir
  End Select
End Sub
Private Sub CmdClientes_Click()

  Dim RsC As ADODB.Recordset
  On Error GoTo errHandler
  
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  CrearRs
  CmdBotones(0).Enabled = False
  
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ClienteTraerUno", nDat, True)
     TxtCliente.text = RsC!Cliente
     LblCliente.Caption = RsC!RazonSocial
     Option1.Value = True
     LlenarGrilla
     nLlama = 1
     TxtEfectivo.text = LblSaldo.Caption
     CmbComprobante_Click
     TxtEfectivo.SetFocus
     If TxtEfectivo.text > 0 Then
        CmdBotones(0).Enabled = True
     Else
        TxtCliente.text = ""
        LblCliente.Caption = ""
        TxtCliente.SetFocus
     End If
     RsC.Close
  End If
  Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CmdDetalle_Click()
  VentaCobranzasDetalles.Show 1
End Sub
Private Sub DtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Activate()
  TxtCliente.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      Unload Me
      Exit Sub
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  If CmbComprobante.ListIndex <> -1 Then
     CmbComprobante.ListIndex = 0
  End If
  LblCliente.Caption = ""
  If CmbCaja.ListIndex <> -1 Then
     CmbCaja.ListIndex = 0
  End If
  DtpFecha.Value = Date
  TxtEfectivo.text = "0.00"
  TxtCancelar.text = "0.00"
  LblNumero.Caption = "0000-00000000"
  LblImporte.Caption = "0.00"
  LblCobrado.Caption = "0.00"
  LblCancelado.Caption = "0.00"
  LblSaldo.Caption = "0.00"
  
  LblSeleccionado.Caption = "0.00"
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
Private Sub Form_Load()

 CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
 CmdBotones(1).Picture = LoadResPicture("Valores", 0)
 CmdBotones(2).Picture = LoadResPicture("Salir", 0)
 CmdBotones(3).Picture = LoadResPicture("Imprimir", 0)

 Limpiar
 nLlama = 1
 Option1.Value = True
 HabilitarTodo True, True
 CmdDetalle.Enabled = False
 CargarCombos
 CrearRs
 Botones False, False, True
 Me.Top = 0
 Me.Left = 0
 
End Sub
Private Sub Botones(bGrabar As Boolean, bValor As Boolean, bSalir As Boolean)
  CmdBotones(3).Enabled = False
  CmdBotones(0).Enabled = bGrabar
  CmdBotones(1).Enabled = bValor
  CmdBotones(2).Enabled = bSalir
End Sub
Private Sub CargarCombos()
  On Error GoTo errHandler
  Dim cRsv As ClsLectura
  
  Set cRsv = New ClsLectura
  
  cRsv.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=4"
  Set cRsl = New ClsLectura
  cRsv.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", "Caja=" & cRsl.RegPorDefecto("Caja")
  Set cRsl = Nothing
    
  CmbComprobante.ListIndex = 0
  CmbCaja.ListIndex = 0
  Set cRsl = New ClsLectura
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub Grid1_Click()
'  Dim i As Integer
'  On Error GoTo errHandler
'
'  LblSeleccionado.Caption = 0
'
'  If TxtEfectivo.Text <> 0 Then
'     Botones False, False, True
'     CmdAplicar.Enabled = False
'      RsC.Bookmark = Grid1.Bookmark
'      If RsC!Cuota <> 1000 Then
'         CmdAplicar.Enabled = True
'         LblSeleccionado.Caption = LblSeleccionado.Caption - RsC!Cobrar
'         If Grid1.SelBookmarks.Count = 1 Then
'            Botones True, True, True
'         End If
'         If CDbl(TxtEfectivo.Text) > CDbl(RsC!Cobrar) Then
'            TxtCancelar.Text = Format(RsC!Cobrar, "#0.00")
'         Else
'            TxtCancelar.Text = Format(TxtEfectivo.Text, "#0.00")
'         End If
'    End If
'    LblSeleccionado.Caption = Format(LblSeleccionado.Caption, "0.00")
'    cImporte = CCur(TxtEfectivo.Text)
'    nCas = 1
'  End If
'Exit Sub
'
'errHandler:
'   ManejaErrores
End Sub
Private Sub TxtCancelar_GotFocus()
  TxtCancelar.SelStart = 0
  TxtCancelar.SelLength = Len(TxtCancelar.text)
End Sub

Private Sub TxtCancelar_LostFocus()
  If TxtCancelar.text = "" Then TxtCancelar.text = 0
  TxtCancelar.text = Format(TxtCancelar.text, "#0.00")
End Sub

Private Sub TxtCliente_GotFocus()
  TxtCliente.SelStart = 0
  TxtCliente.SelLength = Len(TxtCliente.text)
End Sub
Private Sub TxtCliente_LostFocus()
  Dim cRsv As ClsLectura
  On Error GoTo errHandler
  
  Set cRsv = New ClsLectura
  
  LblCliente.Caption = ""
  CrearRs
  CmdBotones(3).Enabled = False
  If TxtCliente.text <> "" Then
     If cRsv.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & TxtCliente.text & "'") <> "" Then
        LblCliente.Caption = cRsv.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & TxtCliente.text & "'")
        Option1.Value = True
        LlenarGrilla
        nLlama = 1
        TxtEfectivo.text = LblSaldo.Caption
        CmbComprobante_Click
        TxtEfectivo.SetFocus
        If TxtEfectivo.text > 0 Then
           CmdBotones(0).Enabled = True
        Else
           TxtCliente.text = ""
           LblCliente.Caption = ""
           TxtCliente.SetFocus
        End If
     Else
        MsgBox "El Cliente no es el Correcto", vbCritical, "Atención"
        CmdClientes.SetFocus
     End If
  End If
  Set cRsv = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub CrearRs()
 On Error GoTo errHandler

 Set RsC = New ADODB.Recordset
  
 RsC.Fields.Append "Fecha", adDate
 RsC.Fields.Append "Comprobante", adVarChar, 30
 RsC.Fields.Append "Numero", adDouble
 RsC.Fields.Append "Cuota", adInteger
 RsC.Fields.Append "Importe", adCurrency
 RsC.Fields.Append "Cobrado", adCurrency
 RsC.Fields.Append "Cancelado", adCurrency
 RsC.Fields.Append "Cobrar", adCurrency
 RsC.Fields.Append "Venta", adDouble
 RsC.Fields.Append "Sucursal", adInteger
 RsC.CursorType = adOpenKeyset
 RsC.LockType = adLockOptimistic
 RsC.Open
 
 CabGrid
 
 Exit Sub
errHandler:
   ManejaErrores

End Sub
Private Sub CabGrid()
  Set Grid1.DataSource = RsC
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
 With Grid1
      .Columns(0).Width = 1400
      .Columns(0).Caption = "Vencimiento"
      .Columns(1).Width = 1600
      .Columns(1).Caption = "Comprobante"
      .Columns(2).Width = 950
      .Columns(2).Caption = "Numero"
      .Columns(2).NumberFormat = "000000"
      .Columns(2).Alignment = dbgRight
      .Columns(3).Width = 700
      .Columns(3).Caption = "Cuota"
      .Columns(3).NumberFormat = "00"
      .Columns(3).Alignment = dbgRight
      .Columns(4).Width = 950
      .Columns(4).Caption = "Importe"
      .Columns(4).NumberFormat = "#0.00"
      .Columns(4).Alignment = dbgRight
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Cobrado"
      .Columns(5).NumberFormat = "#0.00"
      .Columns(5).Alignment = dbgRight
      .Columns(6).Width = 1200
      .Columns(6).Caption = "Cancelado"
      .Columns(6).NumberFormat = "#0.00"
      .Columns(6).Alignment = dbgRight
      .Columns(7).Width = 1300
      .Columns(7).Caption = "A Cobrar"
      .Columns(7).NumberFormat = "#0.00"
      .Columns(7).Alignment = dbgRight
      .Columns(8).Visible = False
      .Columns(9).Visible = False
 End With
End Sub
Private Sub LlenarGrilla()
 Dim cCom As ClsLectura, rCuen As ADODB.Recordset
 Dim sDebe As Currency, sCobrado As Currency, sCobrar As Currency
 Dim sHaber As Currency, hCobrado As Currency, hCobrar As Currency
 Dim cRx As ClsLectura
 
 Dim cCobrado As Currency, cAux As Currency, cAdelanto As Currency, nReg As Long
 Dim sCon As String, nMovi As Long
 
 Set cCom = New ClsLectura
 Set rCuen = New ADODB.Recordset
 Set cRx = New ClsLectura
 
 Set rCuen = cCom.DameCuentaCorriente(TxtCliente.text)
 CmdDetalle.Enabled = False
 LblSeleccionado.Caption = "0.00"
 nReg = 1
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = Left(cCom.DameCodComprobante(rCuen!Comprobante), 15)
    RsC!Numero = rCuen!Numero
    RsC!Cuota = rCuen!Cuota
    RsC!importe = Format(rCuen!Debe, "#0.00")
    cCobrado = cCom.DamePagos(rCuen!Cuota, rCuen!Venta)
    RsC!Cobrado = cCobrado - cAdelanto
    RsC!Cancelado = 0
    RsC!Cobrar = Format((rCuen!Debe), "#0.00") - RsC!Cobrado
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    RsC.Update
    If RsC!Cobrar <= 0 And nReg + 1 <= rCuen.RecordCount Then
       cAdelanto = RsC!Cobrar
       nReg = nReg + 1
       If rCuen.RecordCount <> 1 Then
          db.Open sDb
          db.Execute "UPDATE CuentaCorriente SET Anulado=1 where Venta=" & rCuen!Venta & " AND Cuota=" & rCuen!Cuota
          db.Close
          RsC.Delete
       End If
    Else
       sDebe = sDebe + RsC!importe
       sCobrado = sCobrado + RsC!Cobrado
       sCobrar = sCobrar + RsC!Cobrar
       If RsC!Cobrar = 0 Then
          db.Open sDb
          db.Execute "UPDATE CuentaCorriente SET Anulado=1 where Venta=" & rCuen!Venta & " AND Cuota=" & rCuen!Cuota
          db.Close
          RsC.Delete
       End If
       cAdelanto = 0
    End If
    rCuen.MoveNext
 Loop
 
 If RsC.RecordCount <> 0 Then
    CmdDetalle.Enabled = True
 End If
 Botones False, False, True

 LblImporte.Caption = Format(sDebe - sHaber, "#0.00")
 LblCobrado.Caption = Format(sCobrado - hCobrado, "#0.00")
 LblSaldo.Caption = Format(sCobrar - hCobrar, "#0.00")
 
 Set cRx = Nothing
End Sub
Private Sub TxtEfectivo_GotFocus()
  TxtEfectivo.SelStart = 0
  TxtEfectivo.SelLength = Len(TxtEfectivo.text)
End Sub
Private Sub TxtEfectivo_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtEfectivo) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtEfectivo_LostFocus()
  If TxtEfectivo.text = "" Then TxtEfectivo.text = 0
  TxtEfectivo.text = Format(TxtEfectivo.text, "#0.00")
End Sub

Private Sub GrabarTodo()
  On Error GoTo Errores
  
  Dim cGrab As ClsComprobantesE, cRsl As ClsLectura, nMovimiento As Long, cRcL As ClsComprobantesL
  
  Dim RsCv As ADODB.Recordset, rRet As ADODB.Recordset, rBanco As ADODB.Recordset
  Dim i As Byte, nNum As String * 8, nSuc As String * 4
  
  Dim Rx As ADODB.Recordset
  Dim Rx1 As ADODB.Recordset

  Set cGrab = New ClsComprobantesE
  Set cRsl = New ClsLectura
  Set cRcL = New ClsComprobantesL
  
  Set rValo = New ADODB.Recordset
  Set rApli = New ADODB.Recordset
  Set rReci = New ADODB.Recordset
  Set rCaja = New ADODB.Recordset
  Set rBanco = New ADODB.Recordset
  
  cSaldo = 0

  ' Forma de Pagos,Valores,Aplicaciones,Cantidades,Caja
  Set RsCv = cRsl.RsVacio("CabComprobantes", "Id", "N")
  Set rReci = cRsl.RsVacio("CuentaCorriente", "Id", "N")
  Set rValo = cRsl.RsVacio("Valores", "Id", "N")
  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
  Set rCaja = cRsl.RsVacio("Caja", "Id", "N")
  Set rBanco = cRsl.RsVacio("CtaBanco", "Id", "N")
  Set rAde = cRsl.RsVacio("Aplicaciones", "Id", "N")
  Set rRet = cRsl.RsVacio("Retenciones", "Id", "N")

  
  nMovimiento = cRsl.ActualizarMovimientoTraer

  Dim sSaldo As Currency
  sSaldo = TxtEfectivo.text

  If RsC.RecordCount <> 0 Then
     RsC.MoveFirst
     Do While Not RsC.EOF
        If RsC!Cuota <> 1000 Then
           If CDbl(sSaldo) > CDbl(RsC!Cobrar) Then
              RsC!Cancelado = Format(RsC!Cobrar, "#0.00")
           Else
              RsC!Cancelado = Format(sSaldo, "#0.00")
           End If
           sSaldo = sSaldo - RsC!Cancelado
           RsC!Cobrar = RsC!importe - RsC!Cobrado - RsC!Cancelado
           RsC.Update
           TxtEfectivo.text = Format(sSaldo, "#0.00")
           cSaldo = cSaldo + RsC!Cancelado
        End If
        RsC.MoveNext
     Loop
  Else
'      RsCv.AddNew
'      RsCv!Id = 0
'      RsCv!tipo = 1
'      RsCv!Movimiento = nMovimiento
'      RsCv!Comprobante = 8
'      RsCv!fecha = DTPFecha.Value
'      RsCv!Numero = 0
'      RsCv!Sucursal = 0
'      RsCv!Cliente = TxtCliente.text
'      RsCv!Lista = 0
'      RsCv!Vendedor = 0
'      RsCv!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
'      RsCv!CondVenta = 0
'      RsCv!Neto = 0
'      RsCv!Iva1 = 0
'      RsCv!Descuentos = 0
'      RsCv!Financiacion = 0
'      RsCv!NoGravados = 0
'      RsCv!Debe = 0
'      RsCv!Haber = 0
'      RsCv!Desde = 0
'      RsCv!Hasta = 0
'      RsCv!factura = 0
'      RsCv!Motivo = ""
'      RsCv!COTIZACION = 0
'      RsCv.Update
'
'      nMovimiento = cRsl.ActualizarMovimientoTraer
'
'      rReci.AddNew
'      rReci!Id = 0
'      rReci!fecha = DTPFecha.Value
'      rReci!Cliente = TxtCliente.text
'      rReci!Venta = nMovimiento
'      rReci!Comprobante = 8
'      rReci!Numero = 0
'      rReci!Sucursal = 0
'      rReci!Cuota = 0
'      rReci!Debe = 0
'      rReci!Haber = sSaldo
'      rReci.Update
'
'      rApli.AddNew
'      rApli!Venta = nMovimiento
'      rApli!MoviCta = nMovimiento
'      rApli!Comprobante = 8
'      rApli!Numero = 0
'      rApli!Sucursal = 0
'      rApli!Cuota = 0
'      rApli!importe = sSaldo
'      rApli.Update

  End If
  
  If RsC.RecordCount <> 0 Then
     RsC.MoveFirst
  Else
     cSaldo = TxtEfectivo.text
     TxtEfectivo.text = 0
  End If
  
  RsCv.AddNew
  RsCv!Suc = nSucursal
  RsCv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
  RsCv!Movimiento = nMovimiento
  RsCv!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  RsCv!fecha = DtpFecha.Value
  RsCv!FechaIva = DtpFecha.Value
  RsCv!Numero = Right(LblNumero.Caption, 8)
  RsCv!Sucursal = Left(LblNumero.Caption, 4)
  RsCv!Cliente = TxtCliente.text
  RsCv!Lista = 0
  RsCv!Vendedor = 0
  RsCv!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
  RsCv!CondVenta = 0
  RsCv!Neto = 0
  RsCv!Iva1 = 0
  RsCv!Descuentos = 0
  RsCv!Financiacion = 0
  RsCv!NoGravados = 0
  Select Case cRcL.TraerTipoMovimiento(RsCv!Comprobante)
         Case 13
              RsCv!Debe = cSaldo + sSaldo
              RsCv!Haber = 0
         Case 10
              RsCv!Debe = 0
              RsCv!Haber = cSaldo + sSaldo
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  RsCv!factura = 0
  RsCv!Motivo = ""
  RsCv!COTIZACION = 1
  RsCv!Anulado = 0
  RsCv.Update

     ' Grabo Recibo de Pago
  If RsC.RecordCount <> 0 Then
     rReci.AddNew
     rReci!fecha = DtpFecha.Value
     rReci!Cliente = TxtCliente.text
     rReci!Venta = RsCv!Movimiento
     rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
     rReci!Numero = cRcL.TraerUltimoNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex))
     rReci!Sucursal = cRcL.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex))
     Select Case cRcL.TraerTipoMovimiento(rReci!Comprobante)
            Case 10
                 rReci!Cuota = 0
                 rReci!Debe = 0
                 rReci!Haber = cSaldo + sSaldo
            Case 13
                 rReci!Cuota = 1000
                 rReci!Debe = cSaldo + sSaldo
                 rReci!Haber = 0
      End Select
      rReci!Anulado = 0
      rReci.Update
      RsC.MoveFirst
      If TxtEfectivo.text > 0 Then
         rApli.AddNew
          rApli!Venta = RsC!Venta
          rApli!MoviCta = RsCv!Movimiento
          rApli!Comprobante = rReci!Comprobante
          rApli!Numero = rReci!Numero
          rApli!Sucursal = rReci!Sucursal
          rApli!Cuota = RsC!Cuota
          rApli!importe = TxtEfectivo.text
          rApli!Retencion = 0
          rApli!GAnancia = 0
          rApli!Anulado = 0
          rApli.Update
      End If
  End If
 ' Grabar Valores
   Dim rValoAux As ADODB.Recordset
   If RsC.RecordCount <> 0 Then
      RsC.MoveFirst
   End If
   rValo.AddNew
   rValo!Venta = RsCv!Movimiento
   rValo!CodPago = 1
   rValo!formapago = cRsl.TraerValorDeUnCampo("CondVenta", "Descripcion", "CondVta=1")
   rValo!Nombre = ""
   rValo!importe = cSaldo + CCur(TxtEfectivo.text)
   rValo.Update
   Dim cTotal As Currency, cDife As Currency
   cDife = CCur(TxtEfectivo.text)
  
   If RsC.RecordCount <> 0 Then
      RsC.MoveFirst
      Do While Not RsC.EOF
         If RsC!Cancelado > 0 Then
            rApli.AddNew
            rApli!Venta = RsC!Venta
            rApli!MoviCta = RsCv!Movimiento
            rApli!Comprobante = cRsl.TraerValorDeUnCampo("Comprobantes", "Id", "Descripcion='" & RsC!Comprobante & "'")
            rApli!Numero = RsC!Numero
            rApli!Sucursal = RsC!Sucursal
            rApli!Cuota = RsC!Cuota
            rApli!importe = RsC!Cancelado
            rApli!Retencion = 0
            rApli!GAnancia = 0
            rApli!Anulado = 0
            rApli.Update
        End If
        RsC.MoveNext
      Loop
      RsC.MoveFirst
  End If
 ' Actualizo Caja

  If cRcL.ActualizaCaja(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
     If rValo.EOF = False Or rValo.BOF = False Then
        rValo.MoveFirst
        Do While Not rValo.EOF
           rCaja.AddNew
           rCaja!Movimiento = RsCv!Movimiento
           rCaja!Comprobante = rReci!Comprobante
           rCaja!Numero = rReci!Numero
           rCaja!Sucursal = rReci!Sucursal
           rCaja!fecha = Date
           rCaja!Hora = Time
           rCaja!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
           rCaja!formapago = rValo!formapago
           rCaja!Nombre = rValo!Nombre
           rCaja!Banco = rValo!Banco
           rCaja!NumeroCheque = rValo!NumeroCheque
           rCaja!FechaAcreditacion = rValo!FechaAcreditacion
           rCaja!NumeroTarjeta = rValo!NumeroTarjeta
           rCaja!FechaVencimiento = rValo!FechaVencimiento
           rCaja!Autorizacion = rValo!Autorizacion
           rCaja!Cupon = rValo!Cupon
           rCaja!Debe = rValo!importe
           rCaja!Haber = 0
           rCaja!Abierta = "S"
           rCaja.Update
           rValo.MoveNext
       Loop
     End If
  End If
'  ' Actualizo Ultimos Numero de Comrprobantes
 ' Grabo Todo en los Store Procedure
'
  

  CmdBotones(0).Enabled = False
  CmdBotones(1).Enabled = False
' Imprimir Recibo de Cobro
  cGrab.AgregarCobranzas RsCv, rReci, rValo, rApli, rCaja, rBanco, rRet
  
  sNumero = LblNumero.Caption
  
  nNum = Format(cRsl.TraerValorDeUnCampo("Comprobantes", "Numero", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
  nSuc = Format(cRsl.TraerValorDeUnCampo("Comprobantes", "Sucursal", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
  LblNumero.Caption = nSuc & "-" & nNum

  Set cGrab = Nothing
  Set RsC = Nothing
Exit Sub
Errores:
  MsgBox err.Description, vbCritical, "GrabarTodo"

End Sub

Private Sub Salir()
  If CmdBotones(0).Enabled = True Then
     If MsgBox("El Cobro no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
        Exit Sub
     End If
  End If
  Unload Me
End Sub
Private Sub Valor()
  nLlama = 4
  cSaldo = 0
  
  RsC.MoveFirst
  Do While Not RsC.EOF
     cSaldo = cSaldo + RsC!Cancelado
     RsC.MoveNext
  Loop
  
End Sub

Private Sub Imprimir()
   Dim Respuesta As Boolean, cRsv As ClsLectura
   
   Set cRsv = New ClsLectura
   
   cSaldo = 0
   
   If rApli.RecordCount <> 0 Then
      rApli.MoveFirst
      Do While Not rApli.EOF
         cSaldo = cSaldo + rApli!importe
         rApli.MoveNext
      Loop
   End If
   
   If rAde.RecordCount <> 0 Then
      rAde.MoveFirst
      cSaldo = cSaldo + rAde!importe
   End If
   
   Dim cSal As ClsClienteL
   
   Set cSal = New ClsClienteL
   
   ' Aca Va el recibo de Pago
      
    Dim Y As Double, X As Double, sImpre As String, P As Variant
                    
    sImpre = cRsv.TraerValorDeUnCampo("Impresoras", "Impresora", "Computer='" & sComputadora & "' AND Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         
    If sImpre <> "Ninguna" And sImpre <> "0" And bAnulaPres <> True Then
                                            
     For Each P In Printers
         If P.DeviceName = sImpre Then
             Set Printer = P
             Printer.FontName = "Courier New"
             Printer.ScaleMode = 7
             Printer.FontSize = 9
             Exit For
          End If
      Next
                                 
      Printer.FontSize = 14
                                 
      Y = 1
      X = 0
              
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.FontBold = True
      Printer.Print sNomFan
      Printer.FontBold = False
      
      Printer.FontSize = 7
      
      Y = Y + 0.5
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print Date & " - " & Time
      
       Y = Y + 0.8
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.FontBold = True
      Printer.Print "RECIBO DE PAGO " & LblNumero.Caption
      Printer.FontBold = False

      Y = Y + 0.4
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print "Recibimos de :"

      Y = Y + 0.4
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print Trim(LblCliente.Caption)
      
      Y = Y + 0.4
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print "IMPORTE :" & Format(cSaldo, "#0.00")
      
      Y = Y + 0.5
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print "----------------------------------------"
           
      Y = Y + 0.4
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print "Saldo Anterior :" & Format(cSal.TraerSaldoCtaCte(TxtCliente.text), "#0.00")
       
       Y = Y + 0.8
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print " "
       
      Y = Y + 0.8
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print " "
                            
      Y = Y + 0.8
      Printer.CurrentY = Y
      Printer.CurrentX = X
      Printer.Print "----------------------------------------"
                                  
     Printer.EndDoc
    End If
End Sub


