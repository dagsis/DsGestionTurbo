VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaInfFamilias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadistica de Ventas por Familias"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   7095
   Begin VB.Frame Frame1 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   90
      TabIndex        =   1
      Top             =   15
      Width           =   6840
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1890
         Left            =   225
         TabIndex        =   6
         Top             =   1095
         Width           =   6420
         Begin VB.CheckBox ChkResumen 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   210
            TabIndex        =   17
            Top             =   1590
            Width           =   1245
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   240
            Index           =   1
            Left            =   5220
            TabIndex        =   11
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox TxtCliente 
            Height          =   285
            Index           =   1
            Left            =   495
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "TxtCliente"
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   240
            Index           =   0
            Left            =   5220
            TabIndex        =   9
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox TxtCliente 
            Height          =   285
            Index           =   0
            Left            =   495
            MaxLength       =   15
            TabIndex        =   8
            Text            =   "TxtCliente"
            Top             =   540
            Width           =   855
         End
         Begin VB.CheckBox ChkCliente 
            Caption         =   "Por Clientes"
            Height          =   210
            Left            =   90
            TabIndex        =   7
            Top             =   0
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "Desde Cliente :"
            Height          =   195
            Left            =   525
            TabIndex        =   16
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta Cliente :"
            Height          =   180
            Left            =   495
            TabIndex        =   15
            Top             =   990
            Width           =   1260
         End
         Begin VB.Label Label3 
            Caption         =   "Desde Cliente :"
            Height          =   195
            Left            =   90
            TabIndex        =   14
            Top             =   -15
            Width           =   1350
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
            Height          =   270
            Index           =   1
            Left            =   1395
            TabIndex        =   13
            Top             =   1200
            Width           =   3705
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
            Height          =   270
            Index           =   0
            Left            =   1395
            TabIndex        =   12
            Top             =   540
            Width           =   3705
         End
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   360
         Left            =   1470
         TabIndex        =   2
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   50724865
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   360
         Left            =   3645
         TabIndex        =   4
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   125632513
         CurrentDate     =   37021
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta :"
         Height          =   165
         Left            =   3090
         TabIndex        =   5
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label6 
         Caption         =   "Desde :"
         Height          =   210
         Left            =   840
         TabIndex        =   3
         Top             =   390
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   330
      Left            =   5880
      TabIndex        =   0
      Top             =   3180
      Width           =   1080
   End
End
Attribute VB_Name = "VentaInfFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As Recordset

Private Sub ChkCliente_Click()
  TxtCliente(0).Enabled = False
  TxtCliente(1).Enabled = False
  LblCliente(0).Enabled = False
  LblCliente(1).Enabled = False
  CmdBuscar(0).Enabled = False
  CmdBuscar(1).Enabled = False
  If ChkCliente.Value = 1 Then
     TxtCliente(0).Enabled = True
     TxtCliente(1).Enabled = True
     LblCliente(0).Enabled = True
     LblCliente(1).Enabled = True
     CmdBuscar(0).Enabled = True
     CmdBuscar(1).Enabled = True
  End If
End Sub

Private Sub CmdAplicar_Click()
 Select Case pPlani
        Case 1
             nImpr = 27
        Case 2
             nImpr = 28
        Case 3
             nImpr = 29
 End Select
 If pPlani = 1 Then
     Set Rs = New Recordset
     

     Rs.Fields.Append "Cliente", adVarChar, 15
     Rs.Fields.Append "Nombre", adVarChar, 50
     Rs.Fields.Append "Familia", adVarChar, 30
     Rs.Fields.Append "PrecioTotal", adCurrency
     Rs.Fields.Append "Impuesto", adCurrency
     Rs.Fields.Append "Cotizacion", adCurrency
     Rs.Open
     Dim cCaja As ClsProductoL, Rx As Recordset, fDesde As Date, fHasta As Date, cRsl As ClsLectura, nIva As Integer, nCosto As Integer
     Dim sCliente As String
     
     Set cCaja = New ClsProductoL
     Set cRsl = New ClsLectura
  
     If ChkResumen.Value = 0 Then
        Set Rx = cCaja.TraerVentaFamilia(DtpDesde.Value, DtpHasta.Value, TxtCliente(0).text, TxtCliente(1).text, ChkCliente.Value, ChkResumen.Value)
     Else
        Set Rx = cCaja.TraerVentaFamilia(DtpDesde.Value, DtpHasta.Value, TxtCliente(0).text, TxtCliente(1).text, ChkCliente.Value, ChkResumen.Value)
     End If
     Dim sFam As String, nTotal As Currency
     
     If Rx.RecordCount <> 0 Then
        sFam = "" & Rx!Descripcion
        If ChkResumen.Value = 0 And ChkCliente.Value = 1 Then
           sCliente = "" & Rx!Cliente
        End If
        Do While Not Rx.EOF
            nTotal = 0
            Rs.AddNew
            If ChkCliente.Value = 0 Then
               Rs!Cliente = "000"
               Rs!Nombre = ""
            Else
               If ChkResumen.Value = 0 Then
                  Rs!Cliente = "" & Rx!Cliente
                  Rs!Nombre = "" & Rx!RazonSocial
               Else
               Rs!Cliente = "000"
               Rs!Nombre = "Resumen de los Clientes Emitidos"
               End If
            End If
            Rs!Familia = Rx!Descripcion
            If ChkResumen.Value = 0 And ChkCliente.Value = 1 Then
               Do While sFam = Rx!Descripcion And sCliente = Rx!Cliente
                 nIva = cRsl.TraerValorDeUnCampo("Comprobantes", "Iva", "Id=" & Rx!Comprobante)
                 nCosto = cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Comprobante)
                 If nIva = 0 Or (nIva = 1 And nCosto = 0) Then
                    nTotal = nTotal + Rx!SumaDePrecioTotal
                 Else
                     nTotal = nTotal + Rx!SumaDePrecioTotal + Rx!Impuesto
                 End If
                 Rs!PrecioTotal = nTotal
                 Rx.MoveNext
                 If Rx.EOF Then
                    Exit Do
                 End If
               Loop
           Else
              Do While sFam = Rx!Descripcion
                 nIva = cRsl.TraerValorDeUnCampo("Comprobantes", "Iva", "Id=" & Rx!Comprobante)
                 nCosto = cRsl.TraerValorDeUnCampo("Comprobantes", "Costo", "Id=" & Rx!Comprobante)
                 If nIva = 0 Or (nIva = 1 And nCosto = 0) Then
                    nTotal = nTotal + Rx!SumaDePrecioTotal
                 Else
                    nTotal = nTotal + (Rx!SumaDePrecioTotal + Rx!Impuesto)
                 End If
                 Rs!PrecioTotal = nTotal
                 Rx.MoveNext
                 If Rx.EOF Then
                    Exit Do
                 End If
              Loop
           End If
           Rs.Update
           If Rx.EOF Then
              Exit Do
'           Else
'              Rx.MoveNext
           End If
           sFam = Rx!Descripcion
           If ChkResumen.Value = 0 And ChkCliente.Value = 1 Then
              sCliente = Rx!Cliente
           End If
        Loop
     End If
 End If
 FrmImpresor.Show
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
 Dim crs As ClsLectura
  Set crs = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  If nDat <> 0 Then
     Set Rs = crs.TraerTodos("Clientes", "Id")
     Rs.MoveFirst
     Rs.Find "Id=" & nDat
     Set Rs = crs.TraerRS("ClienteTraerUno", nDat, True)
     LinkearTexto Index
     If Index = 0 Then
        TxtCliente(1).text = Rs!Cliente
        LblCliente(1).Caption = Rs!RazonSocial
     End If
  End If
End Sub

Private Sub Form_Activate()
  Select Case pPlani
         Case 1
              Me.Caption = "Estadistica de Venta"
              ChkCliente.Value = 0
              ChkCliente.Enabled = True
              TxtCliente(0).Enabled = False
              TxtCliente(1).Enabled = False
              LblCliente(0).Enabled = False
              LblCliente(1).Enabled = False
              CmdBuscar(0).Enabled = False
              CmdBuscar(1).Enabled = False
         Case 2
              Me.Caption = "Planilla de Utilidades"
              ChkCliente.Value = 0
              ChkCliente.Enabled = False
              TxtCliente(0).Enabled = False
              TxtCliente(1).Enabled = False
              LblCliente(0).Enabled = False
              LblCliente(1).Enabled = False
              CmdBuscar(0).Enabled = False
              CmdBuscar(1).Enabled = False
         Case 3
              Me.Caption = "Control de Gastos"
              ChkCliente.Value = 0
              ChkCliente.Enabled = False
              TxtCliente(0).Enabled = False
              TxtCliente(1).Enabled = False
              LblCliente(0).Enabled = False
              LblCliente(1).Enabled = False
              CmdBuscar(0).Enabled = False
              CmdBuscar(1).Enabled = False
  End Select
End Sub

Private Sub Form_Load()
  Limpiar
  CargarEtiquetas
End Sub

Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub

Private Sub TxtCliente_LostFocus(Index As Integer)
  Dim cRsl As ClsClienteL
  CmdAplicar.Enabled = False
  LblCliente(Index).Caption = ""
  Set cRsl = New ClsClienteL
  If TxtCliente(Index).text <> "" Then
     If cRsl.BuscarNombreCliente(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cRsl.BuscarNombreCliente(TxtCliente(Index).text)
        CmdAplicar.Enabled = True
     End If
  End If
End Sub

Private Sub LinkearTexto(Index As Integer)
  If Rs.RecordCount <> 0 Then
     TxtCliente(Index).text = Rs!Cliente
     LblCliente(Index).Caption = Rs!RazonSocial
  End If
End Sub
Private Sub DtpFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim WshShell As Object
  If KeyCode = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
     KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
End Sub

Private Sub CargarEtiquetas()
  Dim cRsl As ClsClienteL
  Dim Rz As ADODB.Recordset

  Set Rz = New ADODB.Recordset
  Set cRsl = New ClsClienteL
  Set Rs = cRsl.ClientesAlfabetico(0, False)
  
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Cliente
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Cliente
  LblCliente(1).Caption = Rs!RazonSocial
  
End Sub
