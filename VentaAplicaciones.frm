VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentaAplicaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicaciones"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Anular"
      Height          =   720
      Index           =   1
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4785
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Grabar"
      Height          =   720
      Index           =   0
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4785
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   720
      Index           =   2
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4785
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
      Height          =   3720
      Left            =   120
      TabIndex        =   2
      Top             =   915
      Width           =   6555
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2865
         Left            =   210
         TabIndex        =   3
         Top             =   270
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   5054
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
      Begin VB.Label Label2 
         Caption         =   "Total :"
         Height          =   240
         Left            =   4425
         TabIndex        =   9
         Top             =   3315
         Width           =   675
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
         Left            =   5115
         TabIndex        =   8
         Top             =   3315
         Width           =   1005
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
         Left            =   3795
         TabIndex        =   7
         Top             =   3615
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   2730
         TabIndex        =   6
         Top             =   3615
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Seleccionado :"
         Height          =   210
         Left            =   330
         TabIndex        =   5
         Top             =   3315
         Width           =   1065
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
         Left            =   1455
         TabIndex        =   4
         Top             =   3315
         Width           =   1095
      End
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
      Left            =   195
      TabIndex        =   13
      Top             =   435
      Width           =   4950
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   195
      Width           =   870
   End
   Begin VB.Label Label3 
      Caption         =   "Total Credito :"
      Height          =   180
      Left            =   5250
      TabIndex        =   11
      Top             =   195
      Width           =   1095
   End
   Begin VB.Label LblCredito 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblCredito"
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
      Left            =   5265
      TabIndex        =   10
      Top             =   435
      Width           =   1080
   End
End
Attribute VB_Name = "VentaAplicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura
Dim sw As Boolean, sw1 As Boolean, cImporte As Currency
Dim RsC As ADODB.Recordset
Dim rApli As ADODB.Recordset
Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              GrabarTodo
              CrearRs
              sw = True
              LlenarGrilla
         Case 1
              Anular
         Case 2
              Salir
  End Select
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
  LblCliente.Caption = ""
  LblCobrado.Caption = "0.00"
  LblSaldo.Caption = "0.00"
  LblSeleccionado.Caption = "0.00"
End Sub
Private Sub HabilitarTodo(bGrid As Boolean)
  Grid1.Enabled = bGrid
End Sub
Private Sub Form_Load()
 
 On Error GoTo errores
 CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(2).Picture = LoadResPicture("Salir", 0)

 Limpiar
 HabilitarTodo True
 Botones False, False, True
 sw1 = False
 If rApli.RecordCount <> 0 Then
    sw = True
    HabilitarTodo False
    CmdBotones(1).Enabled = True
 Else
volver:
    BorrarTabla
    sw = False
 End If
 CrearRs
 LblCliente.Caption = VentaNotaCredito.LblCliente.Caption
 LblCredito.Caption = Format(VentaNotaCredito.LblTotal.Caption, "#0.00")
 LlenarGrilla
 Me.Top = 0
 Me.Left = 0
 Exit Sub
errores:
  Resume volver
End Sub
Private Sub Botones(bGrabar As Boolean, bAnu As Boolean, bSalir As Boolean)
  CmdBotones(0).Enabled = bGrabar
  CmdBotones(1).Enabled = bAnu
  CmdBotones(2).Enabled = bSalir
End Sub

Private Sub Grid1_Click()
  Dim i As Integer
  On Error GoTo errHandler
  
  LblSeleccionado.Caption = 0
  
  Botones False, False, True
  For i = 0 To Grid1.SelBookmarks.Count - 1
      RsC.Bookmark = Grid1.SelBookmarks(i)
      If LblSeleccionado.Caption < LblCredito.Caption Then
         sw1 = False
      End If
      If sw1 = False Then
         LblSeleccionado.Caption = LblSeleccionado.Caption + RsC!Cobrar
         If LblSeleccionado.Caption >= LblCredito.Caption Then
            Botones True, False, True
            sw1 = True
         End If
      Else
          MsgBox "Limite Exedido de Aplicaciones", vbCritical, "Atención"
          Grid1.SelBookmarks.Remove (i)
      End If
  Next i
  LblSeleccionado.Caption = Format(LblSeleccionado.Caption, "0.00")
  cImporte = CCur(LblCredito.Caption)
Exit Sub

errHandler:
   ManejaErrores

End Sub
Private Sub CrearRs()
 On Error GoTo errHandler

 Set RsC = New ADODB.Recordset
  
 RsC.Fields.Append "Fecha", adDate
 RsC.Fields.Append "Comprobante", adVarChar, 4
 RsC.Fields.Append "Numero", adDouble
 RsC.Fields.Append "Cuota", adInteger
 RsC.Fields.Append "Importe", adCurrency
 RsC.Fields.Append "Cobrado", adCurrency
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
  Grid1.HeadFont.Bold = True
 With Grid1
      .Columns(0).Width = 1400
      .Columns(0).Caption = "Vencimiento"
      .Columns(1).Width = 1500
      .Columns(1).Caption = "Comprobante"
      .Columns(2).Width = 950
      .Columns(2).Caption = "Numero"
      .Columns(2).NumberFormat = "000000"
      .Columns(2).Alignment = dbgRight
      .Columns(3).Width = 700
      .Columns(3).Caption = "Cuota"
      .Columns(3).NumberFormat = "00"
      .Columns(3).Alignment = dbgRight
      .Columns(4).Visible = False
      .Columns(5).Visible = False
      .Columns(6).Width = 1050
      .Columns(6).Caption = "Saldo"
      .Columns(6).NumberFormat = "#0.00"
      .Columns(6).Alignment = dbgRight
      .Columns(7).Visible = False
      .Columns(8).Visible = False
 End With
End Sub

Private Sub LlenarGrilla()
 Dim cCom As ClsComprobantesL, rCuen As ADODB.Recordset
 Dim sDebe As Currency, sCobrado As Currency, sCobrar As Currency
 Dim cCobrado As Currency, cAux As Currency
 
 Set cCom = New ClsComprobantesL
 Set rCuen = New ADODB.Recordset
 
 Set rCuen = cCom.DameCuentaCorriente(VentaNotaCredito.TxtCliente.Text)
 LblSeleccionado.Caption = "0.00"
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!Fecha = rCuen!Fecha
    RsC!Comprobante = cCom.DameCodComprobante(rCuen!Comprobante)
    RsC!Numero = rCuen!Numero
    RsC!Cuota = rCuen!Cuota
    RsC!Importe = Format(rCuen!Debe, "#0.00")
    cCobrado = cCom.DamePagos(rCuen!Cuota, rCuen!Venta, sw)
    RsC!Cobrado = cCobrado
    RsC!Cobrar = Format(rCuen!Debe, "#0.00") - RsC!Cobrado
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    RsC.Update
    If RsC!Cobrar = 0 Then
       RsC.Delete
    Else
        sDebe = sDebe + RsC!Importe
        sCobrado = sCobrado + RsC!Cobrado
        sCobrar = sCobrar + RsC!Cobrar
    End If
    rCuen.MoveNext
 Loop
 LblImporte.Caption = Format(sDebe, "#0.00")
 LblCobrado.Caption = Format(sCobrado, "#0.00")
 LblSaldo.Caption = Format(sCobrar, "#0.00")
End Sub

Private Sub GrabarTodo()
  Dim cTotal As Currency, cDife As Currency, cNum As ClsComprobantesL
  Dim i As Integer
  Dim cGrab As ClsComprobantesE
  
  Set cRsl = New ClsLectura
  Set rApli = New ADODB.Recordset
  Set cNum = New ClsComprobantesL
  Set cGrab = New ClsComprobantesE
  
  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
   
  cDife = CCur(LblCredito.Caption)
  For i = 0 To Grid1.SelBookmarks.Count - 1
     RsC.Bookmark = Grid1.SelBookmarks(i)
     cTotal = cDife - RsC!Cobrar
     rApli.AddNew
     rApli!Venta = RsC!Venta
     rApli!Comprobante = VentaNotaCredito.CmbComprobante.ItemData(VentaNotaCredito.CmbComprobante.ListIndex)
     rApli!Numero = Right(VentaNotaCredito.LblNumero.Caption, 6)
     rApli!Sucursal = Left(VentaNotaCredito.LblNumero.Caption, 6)
     rApli!Cuota = RsC!Cuota
     If cTotal >= 0 Then
        rApli!Importe = RsC!Cobrar
     Else
        rApli!Importe = Abs(cDife)
     End If
     cDife = cTotal
     rApli.Update
  Next i
'
  CmdBotones(0).Enabled = False
  CmdBotones(1).Enabled = True
  cGrab.GrabarAplicaciones rApli
  LblCredito.Caption = "0.00"
  HabilitarTodo False
  Set RsC = Nothing
Exit Sub

End Sub
Private Sub Salir()
  If CmdBotones(0).Enabled = True Then
     If MsgBox("La Aplicación no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
        Exit Sub
     End If
  End If
  Unload Me
End Sub
Private Sub BorrarTabla()
  db.Execute "DELETE FROM AplicacionAuxiliar"
End Sub
Private Sub Anular()
 If MsgBox("Anula la Aplicación, recuperando Aplicaciones Anteriores ?", 20, "Atención") = vbYes Then
    HabilitarTodo True
    BorrarTabla
    sw = False
    sw1 = False
    CrearRs
    LlenarGrilla
    LblCredito.Caption = Format(VentaNotaCredito.LblTotal.Caption, "#0.00")
    CmdBotones(1).Enabled = False
 End If
End Sub
