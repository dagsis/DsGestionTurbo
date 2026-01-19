VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ComprasAdeAplica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicaciones"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      Height          =   3525
      Left            =   75
      TabIndex        =   4
      Top             =   1335
      Width           =   10110
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2835
         Left            =   165
         TabIndex        =   5
         Top             =   225
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   5001
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
               LCID            =   11274
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
               LCID            =   11274
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
         Left            =   7365
         TabIndex        =   10
         Top             =   3135
         Width           =   1185
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
         Left            =   5205
         TabIndex        =   9
         Top             =   3135
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
         Left            =   6360
         TabIndex        =   8
         Top             =   3135
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
         Left            =   8565
         TabIndex        =   7
         Top             =   3135
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   "Totales :"
         Height          =   240
         Left            =   4440
         TabIndex        =   6
         Top             =   3135
         Width           =   675
      End
   End
   Begin VB.TextBox TxtCancelar 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2730
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "TxtCancela"
      Top             =   5085
      Width           =   810
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   285
      Left            =   3615
      TabIndex        =   2
      Top             =   5085
      Width           =   870
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   720
      Index           =   0
      Left            =   9285
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4965
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Grabar"
      Height          =   720
      Index           =   1
      Left            =   8355
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4965
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   315
      Left            =   2430
      TabIndex        =   11
      Top             =   960
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   556
      _Version        =   393216
      Format          =   708378625
      CurrentDate     =   38050
   End
   Begin VB.Label Label1 
      Caption         =   "Total Nota de Debito/Credito :"
      Height          =   480
      Left            =   195
      TabIndex        =   18
      Top             =   4995
      Width           =   1320
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   1560
      TabIndex        =   17
      Top             =   5085
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   255
      TabIndex        =   16
      Top             =   75
      Width           =   1155
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
      Height          =   300
      Left            =   1470
      TabIndex        =   15
      Top             =   90
      Width           =   2520
   End
   Begin VB.Label LblNumero 
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
      Height          =   285
      Left            =   2490
      TabIndex        =   14
      Top             =   510
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "Número :"
      Height          =   240
      Left            =   270
      TabIndex        =   13
      Top             =   525
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha :"
      Height          =   210
      Left            =   300
      TabIndex        =   12
      Top             =   960
      Width           =   645
   End
End
Attribute VB_Name = "ComprasAdeAplica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura
Dim sw As Boolean, cImporte As Currency
Dim rReci As ADODB.Recordset, rCaja As ADODB.Recordset
Dim rValo As ADODB.Recordset, rApli As ADODB.Recordset
Dim rAde As ADODB.Recordset, bApli As Boolean

Public cSaldo As Currency
Public RsC As ADODB.Recordset
Public nCas As Integer

Private Sub CmdBotones_Click(Index As Integer)
  Dim cRx As ClsComprobantesL
  
  Set cRx = New ClsComprobantesL
  Select Case Index
         Case 0
              Salir
         Case 1
              GrabarTodo
              CrearRs
              LlenarGrilla
              CmdBotones(1).Enabled = False
              bApli = False
              If cRx.TieneAplicacion(ComprasAdelantos.CmbComprobante.ItemData(ComprasAdelantos.CmbComprobante.ListIndex), Right(LblNumero.Caption, 8), ComprasAdelantos.TxtCliente.text) = True Then
                 bApli = True
                 Me.Caption = "Adelanto a Proveedor Aplicado"
                 CmdAplicar.Enabled = False
              End If
  End Select

End Sub
Private Sub Salir()
  If CmdBotones(1).Enabled = True Then
     If MsgBox("La Aplicación no se Guardo, Continua Sin Grabar ?", 20, "Atención") = vbNo Then
        Exit Sub
     End If
  End If
  Unload Me
End Sub

Private Sub TxtCancelar_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtCancelar) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub Grid1_Click()
  Dim i As Integer
  On Error GoTo errHandler
  
  
  CmdAplicar.Enabled = False
 ' For i = 0 To Grid1.SelBookmarks.Count - 1
  If bApli = False Then
     RsC.Bookmark = Grid1.Bookmark
     If RsC!Cuota <> 1000 Then
        CmdAplicar.Enabled = True
        If Grid1.SelBookmarks.Count = 1 Then
            Botones True, True
         End If
         If CDbl(lblTotal.Caption) > CDbl(RsC!Cobrar) Then
            TxtCancelar.text = Format(RsC!Cobrar, "#0.00")
         Else
            TxtCancelar.text = Format(lblTotal.Caption, "#0.00")
         End If
      End If
 ' Next i
 ' TxtEfectivo.Text = Format(Abs(LblSeleccionado.Caption), "#0.00")
 ' TxtCancelar.Text = Format(Abs(LblSeleccionado.Caption), "#0.00")
    cImporte = CCur(lblTotal.Caption)
    nCas = 1
  End If
Exit Sub

errHandler:
   ManejaErrores
End Sub

Private Sub TxtCancelar_GotFocus()
  TxtCancelar.SelStart = 0
  TxtCancelar.SelLength = Len(TxtCancelar.text)
End Sub

Private Sub TxtCancelar_LostFocus()
  If TxtCancelar.text = "" Then TxtCancelar.text = 0
  TxtCancelar.text = Format(TxtCancelar.text, "#0.00")
End Sub

Private Sub CmdAplicar_Click()
  Dim nCotizacion As Currency, cRx As ClsLectura
  
  Set cRx = New ClsLectura
  nCotizacion = cRx.TraerValorDeUnCampo("CabComprobantes", "Cotizacion", "Movimiento=" & RsC!Venta)
'    RsC!Importe = Format(rCuen!Haber * nCotizacion, "#0.00")

  lblTotal.Caption = lblTotal.Caption - TxtCancelar.text
  LblCancelado.Caption = CDbl(LblCancelado.Caption) + CDbl(TxtCancelar.text)
  RsC!Cancelado = TxtCancelar.text
  RsC!Cobrar = (RsC!importe * nCotizacion) - RsC!Cobrado - RsC!Cancelado
  RsC.Update
  lblTotal.Caption = Format(lblTotal.Caption, "#0.00")
  TxtCancelar.text = Format(0, "#0.00")
  LblCancelado.Caption = Format(LblCancelado, "#0.00")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
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
 Dim cCom As ClsComprobantesL, rCuen As ADODB.Recordset
 Dim sDebe As Currency, sCobrado As Currency, sCobrar As Currency
 Dim sHaber As Currency, hCobrado As Currency, hCobrar As Currency
 Dim cRx As ClsLectura, nCotizacion As Currency

 Set cRx = New ClsLectura

 Dim cCobrado As Currency, cAux As Currency
 
 Set cCom = New ClsComprobantesL
 Set rCuen = New ADODB.Recordset
 
 Set rCuen = cCom.DameCuentaCorrienteP(ComprasAdelantos.TxtCliente.text)
 Set db1 = New Connection
 db1.Open sDb
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = Left(cCom.DameCodComprobante(rCuen!Comprobante), 15)
    RsC!Numero = rCuen!Numero
    RsC!Cuota = rCuen!Cuota
    nCotizacion = cRx.TraerValorDeUnCampo("CabComprobantes", "Cotizacion", "Movimiento=" & rCuen!Venta)
    RsC!importe = Format(rCuen!Haber * nCotizacion, "#0.00")
    cCobrado = cCom.DamePagos(rCuen!Cuota, rCuen!Venta)
    RsC!Cobrado = cCobrado
    RsC!Cancelado = 0
    RsC!Cobrar = Format((rCuen!Haber * nCotizacion), "#0.00") - RsC!Cobrado
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    RsC.Update
    If RsC!Cobrar = 0 Then
       RsC.Delete
    Else
        sDebe = sDebe + RsC!importe
        sCobrado = sCobrado + RsC!Cobrado
        sCobrar = sCobrar + RsC!Cobrar
    End If
    rCuen.MoveNext
 Loop
 db1.Close
' Cuenta Corriente en el Haber

 LblImporte.Caption = Format(sDebe - sHaber, "#0.00")
 LblCobrado.Caption = Format(sCobrado - hCobrado, "#0.00")
 LblSaldo.Caption = Format(sCobrar - hCobrar, "#0.00")
 
End Sub

Private Sub Form_Load()
  Dim cRx As ClsComprobantesL, cApli As Currency
  
  Set cRx = New ClsComprobantesL
  CrearRs
  Limpiar
  CargarCombos
  LlenarGrilla
  Botones True, False
  bApli = False
  cApli = cRx.TieneAplicacion(ComprasAdelantos.CmbComprobante.ItemData(ComprasAdelantos.CmbComprobante.ListIndex), Right(LblNumero.Caption, 8), ComprasAdelantos.TxtCliente.text)
  If cApli <> 0 Then
     bApli = False
     Me.Caption = "Adelanto de Proveedor *****Aplicado*****"
  End If
  CmdBotones(0).Picture = LoadResPicture("Salir", 0)
  CmdBotones(1).Picture = LoadResPicture("Grabar", 0)
  lblTotal.Caption = Format(ComprasAdelantos.lblTotal.Caption - cApli, "#0.00")
  If lblTotal.Caption > 0 Then
     bApli = False
     Me.Caption = "Adelanto a Proveedor Aplicado Parcialmente o Sin Aplicación"
  End If
End Sub
Private Sub Botones(bSal As Boolean, bGra As Boolean)
  CmdBotones(0).Enabled = bSal
  CmdBotones(1).Enabled = bGra
End Sub
Private Sub CargarCombos()
  On Error GoTo errHandler

  LblComprobante.Caption = ComprasAdelantos.CmbComprobante.text
  LblNumero.Caption = ComprasAdelantos.LblNumero.Caption
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub Limpiar()
  TxtCancelar.text = "0.00"
  LblImporte.Caption = "0.00"
  LblCobrado.Caption = "0.00"
  LblCancelado.Caption = "0.00"
  LblSaldo.Caption = "0.00"
End Sub

Private Sub GrabarTodo()
  On Error GoTo Errores
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cStock As ClsComprobantesL
  
  Dim i As Byte
  
  Set cRsl = New ClsLectura
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
  Set rValo = New ADODB.Recordset
  Set rApli = New ADODB.Recordset
  Set rReci = New ADODB.Recordset
  Set rCaja = New ADODB.Recordset
  
  
  Set rApli = cRsl.RsVacio("Aplicaciones", "Id", "N")
  
  cSaldo = 0
  
  RsC.MoveFirst
  Do While Not RsC.EOF
     cSaldo = cSaldo + RsC!Cancelado
     RsC.MoveNext
  Loop
  RsC.MoveFirst
  

 ' Grabar Valores
 
  Dim cTotal As Currency, cDife As Currency
  cDife = CCur(lblTotal.Caption)
  RsC.MoveFirst
  Do While Not RsC.EOF
     If RsC!Cancelado > 0 Then
        rApli.AddNew
        rApli!Venta = RsC!Venta
        rApli!MoviCta = nMoviAde
        rApli!Comprobante = ComprasAdelantos.CmbComprobante.ItemData(ComprasAdelantos.CmbComprobante.ListIndex)
        rApli!Numero = Right(LblNumero.Caption, 8)
        rApli!Sucursal = cNum.TraerSucursalNumero(ComprasAdelantos.CmbComprobante.ItemData(ComprasAdelantos.CmbComprobante.ListIndex))
        rApli!Cuota = RsC!Cuota
        rApli!importe = RsC!Cancelado
        rApli!Anulado = 0
        rApli.Update
     End If
     RsC.MoveNext
  Loop
  
  nMoviAde = 0
  
' Actualizo Caja
'  ' Actualizo Ultimos Numero de Comrprobantes
 ' Grabo Todo en los Store Procedure
'
  CmdBotones(0).Enabled = True
  CmdBotones(1).Enabled = False
' Imprimir Recibo de Cobro
  cGrab.AgregarAplicacion rApli, 1
    
  LblCancelado.Caption = "0.00"
  
  Set cGrab = Nothing
  Set cNum = Nothing
  Set cStock = Nothing
  Set RsC = Nothing
Exit Sub
Errores:
 MsgBox err.Description, vbCritical, "Atención"

End Sub




