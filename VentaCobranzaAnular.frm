VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentaCobranzaAnular 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anular Cobranzas"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   7035
   Begin VB.CommandButton CmdBotones 
      Caption         =   "DesImputar"
      Height          =   720
      Index           =   2
      Left            =   3585
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4455
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Recibos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Left            =   90
      TabIndex        =   2
      Top             =   135
      Width           =   6750
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   5655
         TabIndex        =   4
         Top             =   495
         Width           =   915
      End
      Begin VB.TextBox TxtCliente 
         Height          =   315
         Left            =   930
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "TxtCliente"
         Top             =   495
         Width           =   750
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2925
         Left            =   135
         TabIndex        =   5
         Top             =   1110
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   5159
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
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   495
         Width           =   645
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
         Height          =   300
         Left            =   1725
         TabIndex        =   6
         Top             =   495
         Width           =   3900
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   720
      Index           =   1
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4455
      Width           =   975
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Anular"
      Height          =   720
      Index           =   0
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4455
      Width           =   975
   End
End
Attribute VB_Name = "VentaCobranzaAnular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura
Dim sw As Boolean, cImporte As Currency
Dim RsC As ADODB.Recordset
Public nCas As Integer, RsAp As Recordset

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Anular
         Case 1
              Salir
         Case 2
              Desimputar
  End Select
End Sub

Private Sub CmdBuscar_Click()
  Dim RsC As ADODB.Recordset
  On Error GoTo errHandler
  
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  CrearRs
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ClienteTraerUno", nDat, True)
     TxtCliente.text = RsC!Cliente
     LblCliente.Caption = RsC!RazonSocial
     LlenarGrilla
     nLlama = 1
     RsC.Close
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
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  LblCliente.Caption = ""
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

 CmdBotones(0).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(1).Picture = LoadResPicture("Salir", 0)
 CmdBotones(2).Picture = LoadResPicture("Valores", 0)

 Limpiar
 nLlama = 1
 HabilitarTodo True, True
 CrearRs
 Botones False, True
 Me.Top = 0
 Me.Left = 0
 
End Sub
Private Sub Botones(bGrabar As Boolean, bSalir As Boolean)
  CmdBotones(0).Enabled = bGrabar
  CmdBotones(1).Enabled = bSalir
  CmdBotones(2).Enabled = bGrabar
End Sub

Private Sub Grid1_Click()
  Dim i As Integer
  On Error GoTo errHandler
  
  
  Botones False, True
  For i = 0 To Grid1.SelBookmarks.Count - 1
      RsC.Bookmark = Grid1.SelBookmarks(i)
      Botones True, True
  Next i
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub Grid1_DblClick()
 Set RsAp = RsC
 VentaVerCobranzas.Show 1
End Sub

Private Sub TxtCliente_GotFocus()
  TxtCliente.SelStart = 0
  TxtCliente.SelLength = Len(TxtCliente.text)
End Sub
Private Sub TxtCliente_LostFocus()
  Dim cRcL As ClsClienteL, cRiL As ClsClienteL
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set cRcL = New ClsClienteL
  LblCliente.Caption = ""
  CrearRs
  If TxtCliente.text <> "" Then
     If cRcL.BuscarNombreCliente(TxtCliente.text) <> "" Then
        LblCliente.Caption = cRcL.BuscarNombreCliente(TxtCliente.text)
        LlenarGrilla
        nLlama = 1
     Else
        MsgBox "El Cliente no es el Correto", vbCritical, "Atención"
        CmdBuscar.SetFocus
     End If
  End If
  Set cRcL = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
End Sub
Private Sub CrearRs()
 On Error GoTo errHandler

 Set RsC = New ADODB.Recordset
  
 RsC.Fields.Append "Movimiento", adDouble
 RsC.Fields.Append "Fecha", adDate
 RsC.Fields.Append "Comprobante", adVarChar, 20
 RsC.Fields.Append "Sucursal", adInteger
 RsC.Fields.Append "Numero", adDouble
 RsC.Fields.Append "Importe", adCurrency
 RsC.Fields.Append "nCompro", adInteger
 RsC.Fields.Append "Cuota", adInteger

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
      .Columns(0).Visible = False
      .Columns(1).Width = 1000
      .Columns(1).Caption = "F.Cobro"
      .Columns(2).Width = 1850
      .Columns(2).Caption = "Comprobante"
      .Columns(3).Caption = "Sucursal"
      .Columns(3).NumberFormat = "000000"
      .Columns(3).Width = 1000
      .Columns(3).Alignment = dbgRight
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Numero"
      .Columns(4).NumberFormat = "000000"
      .Columns(4).Alignment = dbgRight
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Importe"
      .Columns(5).NumberFormat = nCantDecimales
      .Columns(5).Alignment = dbgRight
      .Columns(6).Visible = False
      .Columns(7).Visible = False
 End With
End Sub

Private Sub LlenarGrilla()
 Dim cCom As ClsComprobantesL, rCuen As ADODB.Recordset
 Dim sDebe As Currency, sCobrado As Currency, sCobrar As Currency
 Dim cCobrado As Currency, cAux As Currency
 
 Set cCom = New ClsComprobantesL
 Set rCuen = New ADODB.Recordset
 
 Set rCuen = cCom.TraerRecibosCobros(TxtCliente.text)
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!Movimiento = rCuen!Venta
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = Left(rCuen!Descripcion, 20)
    RsC!Sucursal = rCuen!Sucursal
    RsC!Numero = rCuen!Numero
    RsC!importe = rCuen!Haber
    RsC!nCompro = rCuen!Comprobante
    RsC.Update
    rCuen.MoveNext
 Loop

End Sub

Private Sub Anular()
  Dim i As Integer, cRce As ClsComprobantesE, cRsl As ClsLectura
  
  Dim nOpe As Integer, nTipo As Integer, cLer As ClsComprobantesL
   
  Set cLer = New ClsComprobantesL
  Set cRce = New ClsComprobantesE
  Set cRsl = New ClsLectura

  If MsgBox("Está seguro de Anular el Cobro ?", 20, "Atención") = vbYes Then
     For i = 0 To Grid1.SelBookmarks.Count - 1
         RsC.Bookmark = Grid1.SelBookmarks(i)
         nOpe = cLer.TraerTipoOperacion(RsC!nCompro)
         nTipo = cLer.TraerTipoMovimiento(RsC!nCompro)
         If cRsl.TraerCantidad("Cheques", "MoviProv", "Movimiento=" & RsC!Movimiento & " and MoviProv<>0") = 0 Then
            cRce.AnularComprobantes RsC!Movimiento, nOpe, nTipo
         Else
            MsgBox "El comprobante no se puede anular porque tiene Cheques Imputados", vbCritical, "Atención"
         End If
    Next i
    Botones False, True
    CrearRs
    LlenarGrilla
  End If
End Sub


Private Sub Desimputar()
  Dim i As Integer, cRce As ClsComprobantesE
  
  If MsgBox("Está seguro de Desimputar el Cobro ?", 20, "Atención") = vbYes Then
  
     Set cRce = New ClsComprobantesE
     For i = 0 To Grid1.SelBookmarks.Count - 1
         RsC.Bookmark = Grid1.SelBookmarks(i)
         cRce.DesImputarComprobante RsC!Movimiento
     Next i
     Botones False, True
     CrearRs
     LlenarGrilla
  End If
End Sub

Private Sub Salir()
  If CmdBotones(0).Enabled = True Then
     If MsgBox("El Cobro no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
        Exit Sub
     End If
  End If
  Unload Me
End Sub

