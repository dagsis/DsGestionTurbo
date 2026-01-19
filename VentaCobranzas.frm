VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaCobranzas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobranzas"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10470
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   720
      Index           =   3
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5880
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
      Height          =   5790
      Left            =   60
      TabIndex        =   11
      Top             =   -15
      Width           =   10290
      Begin VB.TextBox TxtSucursal 
         Height          =   315
         Left            =   4665
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "0000"
         Top             =   645
         Width           =   495
      End
      Begin VB.TextBox TxtNumero 
         Height          =   315
         Left            =   5460
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "00000000"
         Top             =   645
         Width           =   855
      End
      Begin VB.ComboBox CmbEmpresa 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   240
         Width           =   3390
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
         Height          =   1665
         Left            =   8535
         TabIndex        =   32
         Top             =   600
         Width           =   1635
         Begin VB.OptionButton Option1 
            Caption         =   "Normal"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Dar Vuelto"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   1305
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Adelanto"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "Aplicar"
         Height          =   285
         Left            =   6495
         TabIndex        =   14
         Top             =   1875
         Width           =   870
      End
      Begin VB.TextBox TxtCancelar 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5565
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "TxtCancela"
         Top             =   1875
         Width           =   810
      End
      Begin VB.TextBox TxtCliente 
         Height          =   315
         Left            =   1125
         TabIndex        =   4
         Text            =   "TxtCliente"
         Top             =   1170
         Width           =   735
      End
      Begin VB.CommandButton CmdClientes 
         Caption         =   "Buscar"
         Height          =   285
         Left            =   7605
         TabIndex        =   13
         Top             =   1170
         Width           =   810
      End
      Begin VB.ComboBox CmbComprobante 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   660
         Width           =   2445
      End
      Begin VB.ComboBox CmbCaja 
         Height          =   315
         Left            =   645
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1890
         Width           =   1950
      End
      Begin VB.TextBox TxtEfectivo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3390
         TabIndex        =   6
         Text            =   "TxtEfectivo"
         Top             =   1875
         Width           =   915
      End
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Detalle"
         Height          =   285
         Left            =   7620
         TabIndex        =   12
         Top             =   1500
         Width           =   810
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   315
         Left            =   7140
         TabIndex        =   3
         Top             =   660
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   709951489
         CurrentDate     =   36821
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   3000
         Left            =   180
         TabIndex        =   15
         Top             =   2355
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
         Left            =   6885
         TabIndex        =   41
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Nº :"
         Height          =   225
         Left            =   6510
         TabIndex        =   40
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label12 
         Caption         =   "Número :"
         Height          =   240
         Left            =   3810
         TabIndex        =   39
         Top             =   645
         Width           =   765
      End
      Begin VB.Label Label17 
         Caption         =   "--"
         Height          =   195
         Left            =   5235
         TabIndex        =   38
         Top             =   630
         Width           =   150
      End
      Begin VB.Label Label11 
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   37
         Top             =   270
         Width           =   885
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   7440
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Label Label10 
         Caption         =   "Sdo a Cancelar :"
         Height          =   210
         Left            =   4335
         TabIndex        =   29
         Top             =   1875
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Clientes :"
         Height          =   240
         Left            =   165
         TabIndex        =   28
         Top             =   1155
         Width           =   900
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
         Left            =   1950
         TabIndex        =   27
         Top             =   1170
         Width           =   5505
      End
      Begin VB.Label Label3 
         Caption         =   "Comprobante :"
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   6480
         TabIndex        =   25
         Top             =   660
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Caja :"
         Height          =   255
         Left            =   195
         TabIndex        =   24
         Top             =   1875
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Totales :"
         Height          =   240
         Left            =   4230
         TabIndex        =   23
         Top             =   5430
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
         Left            =   8355
         TabIndex        =   22
         Top             =   5430
         Width           =   1260
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
         TabIndex        =   21
         Top             =   5430
         Width           =   960
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
         TabIndex        =   20
         Top             =   5430
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "Seleccionado :"
         Height          =   210
         Left            =   240
         TabIndex        =   19
         Top             =   5415
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
         Left            =   1365
         TabIndex        =   18
         Top             =   5415
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "A Cobrar :"
         Height          =   195
         Left            =   2610
         TabIndex        =   17
         Top             =   1875
         Width           =   720
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
         TabIndex        =   16
         Top             =   5430
         Width           =   1185
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Grabar"
      Height          =   720
      Index           =   0
      Left            =   7545
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Valores"
      Height          =   720
      Index           =   1
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   720
      Index           =   2
      Left            =   9420
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Si el pago es solamente en Efectivo grabar sin valores."
      Height          =   420
      Left            =   180
      TabIndex        =   31
      Top             =   6180
      Width           =   2685
   End
End
Attribute VB_Name = "VentaCobranzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsLectura
Dim sw As Boolean, cImporte As Double
Dim rReci As ADODB.Recordset, rCaja As ADODB.Recordset
Dim rValo As ADODB.Recordset, rApli As ADODB.Recordset
Dim rBanco As ADODB.Recordset, sw1 As Boolean

Dim rAde As ADODB.Recordset, nMarca As Byte

Public cSaldo As Double
Public RsC As ADODB.Recordset
Public nCas As Integer, sNumero As String


Private Sub CmbComprobante_Click()
  Dim cNum As ClsComprobantesL, nNum As String * 8, nSuc As String * 4
  On Error GoTo errHandler

  If sw = False Then
     Set cNum = New ClsComprobantesL
     nNum = "00000000" '
     nSuc = "0000"
     LblNumero.Caption = nSuc & "-" & nNum
  End If
  Set cNum = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CmdAplicar_Click()
 If Grid1.SelBookmarks.Count = 1 And TxtCancelar.text <> 0 Then
    TxtEfectivo.text = TxtEfectivo.text - TxtCancelar.text
    LblCancelado.Caption = CDbl(LblCancelado.Caption) + CDbl(TxtCancelar.text)
    RsC!Cancelado = TxtCancelar.text
    RsC!Cobrar = RsC!importe - RsC!Cobrado - RsC!Cancelado
    RsC.Update
    TxtEfectivo.text = Format(TxtEfectivo.text, "#0.00")
    TxtCancelar.text = Format(0, "#0.00")
    LblCancelado.Caption = Format(LblCancelado, "#0.00")
    If TxtEfectivo.text <= 0 Or Option1.Value = False Then
       Botones True, True, True
    End If
  End If
End Sub

Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              GrabarTodo
              CrearRs
              LlenarGrilla
              CmdBotones(3).Enabled = True
            '  ChkVuelto.Value = 0
         Case 1
              Valor
         Case 2
              Salir
         Case 3
              If bRecibo = True Then
                 Imprimir
              Else
                 Comanda
              End If
  End Select
End Sub

Private Sub Comanda()
  Dim cCli As ClsClienteL, pY As Single, cRsl As ClsLectura
  Dim cNum As ClsComprobantesL, nTotalCuotas As Byte, nI As Byte
  
  
  Dim Rp As ADODB.Recordset
  
  Set cCli = New ClsClienteL
  Set cRsl = New ClsLectura
  Set cNum = New ClsComprobantesL
  
  
  Dim P As Printer
  
  For Each P In Printers
      If P.DeviceName = cRsl.TraerValorDeUnCampo("Impresion", "Impresora", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) Then
         Set Printer = P
      '   Printer.FontName = "Courier New"
         Printer.ScaleMode = 7
         Printer.FontBold = cRsl.TraerValorDeUnCampo("Impresion", "Negrita", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Printer.FontSize = cRsl.TraerValorDeUnCampo("Impresion", "Tamaño", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Exit For
      End If
  Next

  For nI = 1 To nRecibo
      cSaldo = 0
      
      Printer.font.Name = "Ms Sans Seriff"
      Printer.font.Size = 10
      Printer.font.bold = True
      Printer.ScaleMode = 7
      Printer.CurrentY = 1.5
      Printer.CurrentX = 0
      Printer.Print cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
      Printer.font.Size = 10
      Printer.font.bold = False
      Printer.CurrentY = 1
      Printer.CurrentX = 0
      Printer.Print "RECIBO DE COBRANZAS"
      Printer.font.Size = 10
      Printer.CurrentY = 3.5
      Printer.CurrentX = 0
      Printer.Print "Nro.de Documento :" & sNumero
      Printer.CurrentY = 4
      Printer.CurrentX = 0
      Printer.Print "Fecha de Cobro :" & DtpFecha.Value
      Printer.CurrentY = 4.5
      Printer.CurrentX = 0
      Printer.Print "Cliente :" & Trim(TxtCliente.text) & " - " & LblCliente.Caption
      Printer.CurrentY = 5
      Printer.CurrentX = 0
      Printer.Print "Domicilio :" & cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & TxtCliente.text & "'")
      Printer.CurrentY = 5.5
      Printer.CurrentX = 0
      Printer.Print "Localidad :" & cRsl.TraerValorDeUnCampo("Clientes", "Ciudad", "Cliente='" & TxtCliente.text & "'")
      Printer.CurrentY = 6
      Printer.CurrentX = 0
      Printer.Print "CUIT :" & cRsl.TraerValorDeUnCampo("Clientes", "Cuit", "Cliente='" & TxtCliente.text & "'")
      Printer.CurrentY = 6.5
      Printer.CurrentX = 0
      Printer.font.Size = 8
      Printer.Print "-----------------------------------------------------------------------"
      Printer.CurrentY = 7
      Printer.CurrentX = 0
      Printer.Print "FECHA       TIPO     NUMERO   CUOTA       IMPORTE"
      Printer.CurrentY = 7.5
      Printer.CurrentX = 0
      Printer.Print "------------------------------------------------------------------------"
      rApli.MoveFirst
      Printer.CurrentY = 8
      pY = 8
      rApli.MoveFirst
      Do While Not rApli.EOF
         Set Rp = cRsl.TraerComproPago(rApli!Venta)
        Printer.CurrentY = pY
        Printer.CurrentX = 0
        Printer.Print Rp!fecha
        Printer.CurrentY = pY
        Printer.CurrentX = 1.5
        Printer.Print Rp!Comprobante
        Printer.CurrentY = pY
        Printer.CurrentX = 2.7
        Printer.Print Format(Rp!Numero, "000000")
        Printer.CurrentY = pY
        Printer.CurrentX = 4.2
        Printer.Print rApli!Cuota & "/" & cRsl.TraerCantidadCuotas(rApli!Venta)
        Printer.CurrentY = pY
        Printer.CurrentX = 7 - Printer.TextWidth(Format(rApli!importe, "#0.00"))
        Printer.Print Format(rApli!importe, "#0.00")
        cSaldo = cSaldo + rApli!importe
        rApli.MoveNext
        pY = pY + 0.5
      Loop
      If rAde.RecordCount <> 0 Then
         rAde.MoveFirst
         Printer.CurrentX = 1
         Printer.Print Format(rReci!Asiento, "00000000")
         Printer.CurrentY = pY
         Printer.CurrentX = 0
         Printer.Print DtpFecha.Value
         Printer.CurrentY = pY
         Printer.CurrentX = 2
         Printer.Print "A Cuenta"
         Printer.CurrentY = pY
         Printer.CurrentX = 2.7
         Printer.Print Format(rAde!Numero, "000000")
         Printer.CurrentY = pY
         Printer.CurrentX = 6.5 - Printer.TextWidth(Format(rAde!importe, "#0.00"))
         Printer.Print Format(rAde!importe, "#0.00")
         cSaldo = cSaldo + rAde!importe
      End If
      pY = pY + 0.5
     Printer.CurrentY = pY
     Printer.CurrentX = 0
     Printer.Print "T O T A L   D E L  C O B R O :"
     Printer.CurrentY = pY
     Printer.CurrentX = 6.5 - Printer.TextWidth(Format(cSaldo, "#0.00"))
     Printer.Print Format(cSaldo, "#0.00")
     rValo.MoveFirst
     pY = pY + 1
     Do While Not rValo.EOF
        Printer.CurrentY = pY
        Printer.CurrentX = 0
        Printer.Print rValo!formapago
        Printer.CurrentY = pY
        Printer.CurrentX = 0
        If rValo!NumeroCheque <> "" Then
           Printer.Print Left(rValo!Nombre, 10) & " " & rValo!Banco & " " & rValo!NumeroCheque & " " & rValo!FechaAcreditacion
        Else
           If rValo!Cupon <> "" Then
              Printer.Print Left(rValo!Nombre, 10) & " " & rValo!Banco & " " & rValo!NumeroTarjeta & " " & rValo!Cupon
           End If
        End If
        Printer.CurrentY = pY
        Printer.CurrentX = 6.5 - Printer.TextWidth(Format(rValo!importe, "#0.00"))
        Printer.Print Format(rValo!importe, "#0.00")
        rValo.MoveNext
        pY = pY + 0.5
     Loop
     pY = pY + 1
     Printer.CurrentY = pY
     Printer.CurrentX = 1
     Printer.Print "CONFECCIONO :" & rUsu!Nombre
     Printer.CurrentY = pY
     Printer.CurrentX = 1
     pY = pY + 0.5
     Printer.CurrentY = pY
     Printer.CurrentX = 1
     Printer.Print "SALDO PENDIENTE :" & LblSaldo.Caption
     pY = pY + 8
     Printer.CurrentY = pY
     Printer.CurrentX = 1
     Printer.Print "------"
     Printer.EndDoc
  Next nI
 ' Limpiar

End Sub

Private Sub CmdClientes_Click()
  Dim RsC As ADODB.Recordset
  On Error GoTo errHandler
  
  Set RsC = New ADODB.Recordset
  
  Set cRsl = New ClsLectura
  nDat = 0
  nBuscar = 1
  FrmBuscarEx.Show 1
  Limpiar
  CrearRs
  If nDat <> 0 Then
     Set RsC = cRsl.TraerRS("ClienteTraerUno", nDat, True)
     TxtCliente.text = RsC!Cliente
     LblCliente.Caption = RsC!RazonSocial
     Option1.Value = True
     LlenarGrilla
     nLlama = 1
     CmbComprobante_Click
     CmbCaja.SetFocus
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
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub Limpiar()
  TxtCliente.text = ""
  If CmbComprobante.ListIndex <> -1 Then
     CmbComprobante.ListIndex = 0
  End If
  LblCliente.Caption = ""
'  If CmbCaja.ListIndex <> -1 Then
'     CmbCaja.ListIndex = 0
'  End If
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
 HabilitarTodo True, True
 CmdDetalle.Enabled = False
 CargarCombos
 CrearRs
 Botones False, False, True
 nMarca = 0
 DtpFecha.Value = Date
 Option1.Value = True
 If bAbiertoVenta = True Then
    TxtCliente.text = VentaComprobantes.TxtCliente.text
    LblCliente.Caption = VentaComprobantes.LblCliente.Caption
    TxtEfectivo.text = VentaComprobantes.lblTotal.Caption
    LlenarGrilla
    nMarca = 1
    CmdClientes.Enabled = False
 End If
 
 CargarEmpresas
 If bAbiertoVenta = True Then
    CmbEmpresa.ListIndex = VentaComprobantes.CmbEmpresa.ListIndex
 End If
 Me.Left = 0
 Me.Top = 1400
End Sub

Private Sub CargarEmpresas()
  Dim db As Connection, cRx As ClsLectura
  Dim sAux As String
  
  
  Set db = New Connection
  Set cRx = New ClsLectura
  
  sAux = sDb
    
  sDb = sUdl
  cRx.CargaCombo CmbEmpresa, "Empresa", "Id", "Razon", ""
  CmbEmpresa.ListIndex = nMenu - 1
  
  sDb = sAux
  
  Set db = Nothing
  Set cRx = Nothing
End Sub
Private Sub Botones(bGrabar As Boolean, bValor As Boolean, bSalir As Boolean)
  CmdBotones(3).Enabled = False
  CmdBotones(0).Enabled = bGrabar
  CmdBotones(1).Enabled = bValor
  CmdBotones(2).Enabled = bSalir
End Sub

Private Sub CmbEmpresa_Click()
  Dim cRsl As ClsLectura, sAux As String
  
  Set cRsl = New ClsLectura
  If sw1 = True Then
   '  db.Close
     db.Open sUdl
     sDb = sUdl
     sAux = "FILE NAME=" & App.Path & "\" & cRsl.TraerValorDeUnCampo("Empresa", "Udl", "Id=" & CmbEmpresa.ListIndex)
     db.Close
     sDb = sAux
  '   db.Open sAux
     CargarCombos
     CrearRs
     Botones False, False, True
     nMarca = 0
     Option1.Value = True
     LlenarGrilla
     nMarca = 1
  End If
  sw1 = True
  
  Set cRsl = Nothing
  
End Sub

Private Sub CargarCombos()
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  
  cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=4 and Suc=" & nSucursal
  
  cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
  CmbComprobante.text = cRsl.TraerValorDeUnCampo("Comprobantes", "Descripcion", "Id=" & cRsl.RegPorDefecto("Cobranza"))
  CmbCaja.text = cRsl.DatoCombo("Cajas", "Caja", "Descripcion", cRsl.RegPorDefecto("Caja"))
  
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub Grid1_Click()
  Dim i As Integer
  On Error GoTo errHandler
  
  LblSeleccionado.Caption = 0
  
  If TxtEfectivo.text <> 0 Then
     Botones False, False, True
     CmdAplicar.Enabled = False
      RsC.Bookmark = Grid1.Bookmark
      If RsC!Cuota <> 1000 Then
         CmdAplicar.Enabled = True
         LblSeleccionado.Caption = LblSeleccionado.Caption - RsC!Cobrar
'         If Grid1.SelBookmarks.Count = 1 Then
'            Botones True, True, True
'         End If
         If CDbl(TxtEfectivo.text) > CDbl(RsC!Cobrar) Then
            TxtCancelar.text = Format(RsC!Cobrar, "#0.00")
         Else
            TxtCancelar.text = Format(TxtEfectivo.text, "#0.00")
         End If
    End If
    LblSeleccionado.Caption = Format(LblSeleccionado.Caption, "0.00")
    cImporte = CCur(TxtEfectivo.text)
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
  CmdBotones(3).Enabled = False
  If TxtCliente.text <> "" Then
     If cRcL.BuscarNombreCliente(TxtCliente.text) <> "" Then
        LblCliente.Caption = cRcL.BuscarNombreCliente(TxtCliente.text)
        Option1.Value = True
        LlenarGrilla
        nLlama = 1
        CmbComprobante_Click
        CmbCaja.SetFocus
     Else
        MsgBox "El Cliente no es el Correcto", vbCritical, "Atención"
        CmdClientes.SetFocus
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
  
 RsC.Fields.Append "Fecha", adDate
 RsC.Fields.Append "Comprobante", adVarChar, 40
 RsC.Fields.Append "Numero", adDouble
 RsC.Fields.Append "Cuota", adInteger
 RsC.Fields.Append "Importe", adDouble
 RsC.Fields.Append "Cobrado", adDouble
 RsC.Fields.Append "Cancelado", adDouble
 RsC.Fields.Append "Cobrar", adDouble
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

 On Error GoTo errHandler
 
 Dim cCom As ClsComprobantesL, rCuen As ADODB.Recordset
 Dim sDebe As Double, sCobrado As Double, sCobrar As Double
 Dim sHaber As Double, hCobrado As Double, hCobrar As Double
 Dim cRx As ClsLectura, nCotizacion As Double
 
 Dim cCobrado As Double, cAux As Double, cAdelanto As Double, nReg As Long
 Dim sCon As String, nMovi As Long
 
 Set cCom = New ClsComprobantesL
 Set rCuen = New ADODB.Recordset
 Set cRx = New ClsLectura
 
 Set rCuen = cCom.DameCuentaCorriente(TxtCliente.text)
 CmdDetalle.Enabled = False
 LblSeleccionado.Caption = "0.00"
 nReg = 1
 Set db1 = New Connection
 db1.Open sDb
 Do While Not rCuen.EOF
    RsC.AddNew
    RsC!fecha = rCuen!fecha
    RsC!Comprobante = Left(cCom.DameCodComprobante(rCuen!Comprobante), 15) & " " & cRsl.TraerValorDeUnCampo("CabComprobantes", "Motivo", "Movimiento= " & rCuen!Venta)
    RsC!Numero = rCuen!Numero
'    If rCuen!Numero = 4795 Then
'       MsgBox "Llegue"
'    End If
    RsC!Cuota = rCuen!Cuota
    nCotizacion = cRx.TraerValorDeUnCampo("CabComprobantes", "Cotizacion", "Movimiento=" & rCuen!Venta)
    RsC!importe = Format(rCuen!Debe * nCotizacion, "#0.00")
    cCobrado = cCom.DamePagos(rCuen!Cuota, rCuen!Venta)
'    If cCobrado <> 0 Then
'       MsgBox rCuen!Venta
'    End If
    RsC!Cobrado = cCobrado - cAdelanto
    RsC!Cancelado = 0
    RsC!Cobrar = Round((rCuen!Debe * nCotizacion), 2) - Round(RsC!Cobrado, 2)
    RsC!Venta = rCuen!Venta
    RsC!Sucursal = rCuen!Sucursal
    RsC.Update
    If RsC!Cobrar <= 0 And nReg + 1 <= rCuen.RecordCount Then
       cAdelanto = RsC!Cobrar
       nReg = nReg + 1
       If rCuen.RecordCount <> 1 Then
          If RsC!Cobrar <= 0 Then
'                 db.Open sDb
'                 db.Execute "UPDATE CuentaCorriente SET Anulado=1 where Venta=" & rCuen!Venta & " AND Cuota=" & rCuen!Cuota
'                 db.Close
             RsC.Delete
          End If
       End If
    Else
       sDebe = sDebe + Round(RsC!importe, 2)
       sCobrado = sCobrado + Round(RsC!Cobrado, 2)
       sCobrar = sCobrar + Round(RsC!Cobrar, 2)
       If RsC!Cobrar = 0 Then
'             db.Open sDb
'             db.Execute "UPDATE CuentaCorriente SET Anulado=1 where Venta=" & rCuen!Venta & " AND Cuota=" & rCuen!Cuota
'             db.Close
          RsC.Delete
       End If

       cAdelanto = 0
    End If
    rCuen.MoveNext
 Loop
 db1.Close
 If RsC.RecordCount <> 0 Then
    CmdDetalle.Enabled = True
 End If
 Botones False, False, True

 LblImporte.Caption = Format(sDebe - sHaber, "#0.00")
 LblCobrado.Caption = Format(sCobrado - hCobrado, "#0.00")
 LblSaldo.Caption = Format(sCobrar - hCobrar, "#0.00")
 
 Exit Sub
errHandler:
   MsgBox "Consultar Programador :" & rCuen!Venta, vbCritical

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
  Dim cGrab As ClsComprobantesE, cNum As ClsComprobantesL, cStock As ClsComprobantesL
  Dim RsCv As ADODB.Recordset, rRet As ADODB.Recordset
  Dim i As Byte, nNum As String * 8, nSuc As String * 4
    
  Dim Rx As ADODB.Recordset
  Dim Rx1 As ADODB.Recordset

  Set cRsl = New ClsLectura
  Set cGrab = New ClsComprobantesE
  Set cNum = New ClsComprobantesL
  Set cStock = New ClsComprobantesL
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

  
  If TxtNumero.text = 0 Then
     If cRsl.TraerValorDeUnCampo("Impresion", "Numero", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) = False Then
        nNum = Format(cNum.TraerUltimoNumeroVenta(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
        nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
     Else
        nNum = Format(cNum.TraerUltimoNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "00000000")
        nSuc = Format(cNum.TraerSucursalNumero(CmbComprobante.ItemData(CmbComprobante.ListIndex)), "0000")
    End If
  Else
     nNum = TxtNumero.text
     nSuc = TxtSucursal.text
  End If
  
  LblNumero.Caption = nSuc & "-" & nNum
  
  TxtNumero.text = nNum
  TxtSucursal.text = nSuc
  
  RsC.MoveFirst
  Do While Not RsC.EOF
     cSaldo = cSaldo + RsC!Cancelado
     RsC.MoveNext
  Loop
  
  RsC.MoveFirst
      
  RsCv.AddNew
  RsCv!Suc = nSucursal
 ' RsCv!Id = 0
  RsCv!tipo = cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Id=" & CmbComprobante.ItemData(CmbComprobante.ListIndex))
  RsCv!Movimiento = cNum.TraerUltimoMovimiento
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
  Select Case cStock.TraerTipoMovimiento(RsCv!Comprobante)
         Case 13
              RsCv!Debe = cSaldo
              RsCv!Haber = 0
         Case 10
              RsCv!Debe = 0
              If Option3.Value = True Then
                 RsCv!Haber = cSaldo + TxtEfectivo.text
              Else
                 RsCv!Haber = cSaldo
              End If
  End Select
  RsCv!Desde = 0
  RsCv!Hasta = 0
  
  nMarca = 0
  If bAbiertoVenta = True Then nMarca = 1
  RsCv!factura = nMarca
  RsCv!Motivo = ""
  RsCv!COTIZACION = 1
  RsCv!Anulado = 0
  RsCv.Update

  nMarca = 0
  bAbiertoVenta = False
     ' Grabo Recibo de Pago
  rReci.AddNew
'  rReci!Id = 0
  rReci!fecha = DtpFecha.Value
  rReci!Cliente = TxtCliente.text
  rReci!Venta = RsCv!Movimiento
  rReci!Comprobante = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  rReci!Numero = RsCv!Numero
  rReci!Sucursal = RsCv!Sucursal
  Select Case cStock.TraerTipoMovimiento(rReci!Comprobante)
         Case 10
              rReci!Cuota = 0
              rReci!Debe = 0
              If Option3.Value = True Then
                 rReci!Haber = cSaldo + TxtEfectivo.text
              Else
                 rReci!Haber = cSaldo
              End If
         Case 13
              rReci!Cuota = 1000
              If Option3.Value = True Then
                 rReci!Debe = cSaldo + TxtEfectivo.text
                 rReci!Haber = 0
              Else
                 rReci!Debe = cSaldo
                 rReci!Haber = 0
              End If
  End Select
  rReci!Anulado = 0
  rReci.Update
  
  RsC.MoveFirst
  
  If TxtEfectivo.text > 0 And Option2.Value = False Then
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
  
 ' Grabar Valores
 Dim rValoAux As ADODB.Recordset
 RsC.MoveFirst
 If nLlama = 4 Then
    Set rValoAux = VentaValores.RsV
    rValoAux.MoveFirst
    Do While Not rValoAux.EOF
       rValo.AddNew
      ' rValo!Id = 0
       rValo!Venta = RsCv!Movimiento
       rValo!CodPago = rValoAux!CodPago
       rValo!formapago = rValoAux!formapago
       rValo!Nombre = rValoAux!Nombre
       rValo!Banco = rValoAux!Banco
       rValo!IdCheque = rValoAux!IdCheque
       rValo!NumeroCheque = rValoAux!NCheque
       rValo!FechaAcreditacion = rValoAux!FAcredi
       rValo!NumeroTarjeta = rValoAux!NTarjet
       rValo!FechaVencimiento = rValoAux!FVencim
       rValo!Autorizacion = rValoAux!NAutori
       rValo!Cupon = rValoAux!NCupon
       rValo!importe = Format(rValoAux!importe, "#0.00")
       rValo!PorVenta = " "
       rValo.Update
       rValoAux.MoveNext
   Loop
    rValoAux.MoveFirst
    RsC.MoveFirst
    Do While Not rValoAux.EOF
       If cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "CondVta=" & rValoAux!CodPago) > 8 Then
          If cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "CondVta=" & rValoAux!CodPago) < 14 Then
            rRet.AddNew
            rRet!Suc = nSucursal
            rRet!Movimiento = RsCv!Movimiento
            rRet!fecha = rValoAux!FAcredi
            rRet!Cliente = TxtCliente.text
            rRet!tipo = cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "CondVta=" & rValoAux!CodPago)
            rRet!Numero = rValoAux!NCheque
            rRet!importe = Format(rValoAux!importe, "#0.00")
            rRet.Update
         End If
       End If
       rValoAux.MoveNext
    Loop
  Else
      rValo.AddNew
     ' rValo!Id = 0
      rValo!Venta = RsCv!Movimiento
      rValo!CodPago = 1
      rValo!formapago = cRsl.DatoCombo("CondVenta", "CondVta", "Descripcion", cRsl.RegPorDefecto("Cobro"))
      rValo!Nombre = ""
      rValo!importe = cSaldo + CCur(TxtEfectivo.text)
      rValo.Update
  End If
  Dim cTotal As Double, cDife As Double
  cDife = CCur(TxtEfectivo.text)
  
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
    
' Actualizo Caja
  RsC.MoveFirst
  If cStock.ActualizaCaja(CmbComprobante.ItemData(CmbComprobante.ListIndex)) = True Then
     If rValo.EOF = False Or rValo.BOF = False Then
        rValo.MoveFirst
        Do While Not rValo.EOF
           If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Trim(rValo!formapago) & "'") = "Caja" Then
              rCaja.AddNew
              rCaja!Suc = nSucursal
              rCaja!Movimiento = RsCv!Movimiento
              rCaja!Comprobante = rReci!Comprobante
              rCaja!Numero = rReci!Numero
              rCaja!Sucursal = rReci!Sucursal
              rCaja!fecha = DtpFecha.Value
              rCaja!Hora = Time
              rCaja!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
              rCaja!formapago = rValo!formapago
              rCaja!Nombre = LblCliente.Caption
              rCaja!Banco = rValo!Banco
              rCaja!NumeroCheque = rValo!NumeroCheque
              rCaja!FechaAcreditacion = rValo!FechaAcreditacion
              rCaja!NumeroTarjeta = rValo!NumeroTarjeta
              rCaja!FechaVencimiento = rValo!FechaVencimiento
              rCaja!Autorizacion = rValo!Autorizacion
              rCaja!Cupon = rValo!Cupon
              If cStock.SumaCaja(rCaja!Comprobante) Then
                 rCaja!Debe = rValo!importe
                 rCaja!Haber = 0
              Else
                 rCaja!Debe = 0
                 rCaja!Haber = rValo!importe
              End If
              rCaja!Abierta = "S"
              rCaja!Anulado = 0
              rCaja.Update
              If Option2.Value = True Then
                 rCaja.AddNew
                 rCaja!Suc = nSucursal
                 rCaja!Movimiento = RsCv!Movimiento
                 rCaja!Comprobante = rReci!Comprobante
                 rCaja!Numero = rReci!Numero
                 rCaja!Sucursal = rReci!Sucursal
                 rCaja!fecha = DtpFecha.Value
                 rCaja!Hora = Time
                 rCaja!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
                 rCaja!formapago = "EFECTIVO"
                 rCaja!Nombre = LblCliente.Caption
                 rCaja!Banco = 0
                 rCaja!NumeroCheque = Null
                 rCaja!FechaAcreditacion = Null
                 rCaja!NumeroTarjeta = Null
                 rCaja!FechaVencimiento = Null
                 rCaja!Autorizacion = Null
                 rCaja!Cupon = Null
                 rCaja!Debe = 0
                 rCaja!Haber = TxtEfectivo.text
                 rCaja!Abierta = "S"
                 rCaja!Anulado = 0
                 rCaja.Update
              End If
           Else
            If cRsl.TraerValorDeUnCampo("CondVenta", "Mueve", "Descripcion='" & Trim(rValo!formapago) & "'") = "Banco" Then
               rBanco.AddNew
               rBanco!Movimiento = RsCv!Movimiento
               rBanco!Comprobante = rReci!Comprobante
               rBanco!Numero = rReci!Numero
               rBanco!Sucursal = rReci!Sucursal
               rBanco!fecha = DtpFecha.Value
               rBanco!Hora = Time
               rBanco!Caja = CmbCaja.ItemData(CmbCaja.ListIndex)
               rBanco!formapago = rValo!formapago
               rBanco!Nombre = rValo!Nombre
               rBanco!Banco = rValo!Banco
               rBanco!NumeroCheque = rValo!NumeroCheque
               rBanco!FechaAcreditacion = rValo!FechaAcreditacion
               rBanco!NumeroTarjeta = rValo!NumeroTarjeta
               rBanco!FechaVencimiento = rValo!FechaVencimiento
               rBanco!Autorizacion = rValo!Autorizacion
               rBanco!Cupon = rValo!Cupon
               If cStock.SumaCaja(rBanco!Comprobante) Then
                  rBanco!Debe = rValo!importe
                  rBanco!Haber = 0
               Else
                  rBanco!Debe = 0
                  rBanco!Haber = rValo!importe
               End If
               rBanco!Abierta = "S"
               rBanco!Anulado = 0
               rBanco.Update
            End If
           End If
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
  
  sNumero = TxtSucursal.text & "-" & TxtNumero.text
  TxtNumero.text = "00000000"
  TxtSucursal.text = "0000"
  
  CmdClientes.Enabled = True

  Set cGrab = Nothing
  Set cNum = Nothing
  Set cStock = Nothing
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

  VentaValores.Show 1
End Sub

Private Sub Imprimir()
  Dim cCli As ClsClienteL, pY As Single, cRsl As ClsLectura
  Dim cNum As ClsComprobantesL, nTotalCuotas As Byte, nI As Byte
  
  
  Dim Rp As ADODB.Recordset
  
  Set cCli = New ClsClienteL
  Set cRsl = New ClsLectura
  Set cNum = New ClsComprobantesL
  
  
  Dim P As Printer
  
  For Each P In Printers
      If P.DeviceName = cRsl.TraerValorDeUnCampo("Impresion", "Impresora", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario) Then
         Set Printer = P
      '   Printer.FontName = "Courier New"
         Printer.ScaleMode = 7
         Printer.FontBold = cRsl.TraerValorDeUnCampo("Impresion", "Negrita", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Printer.FontSize = cRsl.TraerValorDeUnCampo("Impresion", "Tamaño", "Comprobante=" & CmbComprobante.ItemData(CmbComprobante.ListIndex) & " AND Usuario=" & nUsuario)
         Exit For
      End If
  Next

  For nI = 1 To nRecibo
      cSaldo = 0
      
      Printer.font.Name = "Ms Sans Seriff"
      Printer.font.Size = 14
      Printer.font.bold = True
      Printer.ScaleMode = 7
      Printer.CurrentY = 1.5
      Printer.CurrentX = 1
      Printer.Print sNomFan  '  cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
      Printer.font.Size = 14
      Printer.font.bold = False
      Printer.CurrentY = 2.5
      Printer.CurrentX = 1
      Printer.Print "RECIBO DE COBRANZAS"
      Printer.font.Size = 10
      Printer.CurrentY = 3.5
      Printer.CurrentX = 13
      Printer.Print "Nro.de Documento :" & sNumero
      Printer.CurrentY = 4
      Printer.CurrentX = 13
      Printer.Print "Fecha de Cobro :" & DtpFecha.Value
      Printer.CurrentY = 4.5
      Printer.CurrentX = 1
      Printer.Print "Cliente :" & Trim(TxtCliente.text) & " - " & LblCliente.Caption
      Printer.CurrentY = 5
      Printer.CurrentX = 1
      Printer.Print "Domicilio :" & cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & TxtCliente.text & "'")
      Printer.CurrentY = 5.5
      Printer.CurrentX = 1
      Printer.Print "Localidad :" & cRsl.TraerValorDeUnCampo("Clientes", "Ciudad", "Cliente='" & TxtCliente.text & "'")
      Printer.CurrentY = 6
      Printer.CurrentX = 1
      Printer.Print "CUIT :" & cRsl.TraerValorDeUnCampo("Clientes", "Cuit", "Cliente='" & TxtCliente.text & "'")
      Printer.CurrentY = 6.5
      Printer.CurrentX = 1
      Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
      Printer.CurrentY = 7
      Printer.CurrentX = 1
      Printer.Print "                           FECHA                      TIPO                        NUMERO              CUOTA               IMPORTE"
      Printer.CurrentY = 7.5
      Printer.CurrentX = 1
      Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
      rApli.MoveFirst
      Printer.CurrentY = 8
      pY = 8
      rApli.MoveFirst
      Do While Not rApli.EOF
         Set Rp = cRsl.TraerComproPago(rApli!Venta)
        Printer.CurrentY = pY
        Printer.CurrentX = 3
        Printer.Print Rp!fecha
        Printer.CurrentY = pY
        Printer.CurrentX = 6.5
        Printer.Print Rp!Comprobante
        Printer.CurrentY = pY
        Printer.CurrentX = 10
        Printer.Print Format(Rp!Sucursal, "0000") & "-"
        Printer.CurrentY = pY
        Printer.CurrentX = 11
        Printer.Print Format(Rp!Numero, "00000000")
        Printer.CurrentY = pY
        Printer.CurrentX = 13
        Printer.Print rApli!Cuota & "/" & cRsl.TraerCantidadCuotas(rApli!Venta)
        Printer.CurrentY = pY
        Printer.CurrentX = 17.5 - Printer.TextWidth(Format(rApli!importe, "#0.00"))
        Printer.Print Format(rApli!importe, "#0.00")
        cSaldo = cSaldo + rApli!importe
        rApli.MoveNext
        pY = pY + 0.5
      Loop
      If rAde.RecordCount <> 0 Then
         rAde.MoveFirst
         Printer.CurrentX = 1
         Printer.Print Format(rReci!Asiento, "00000000")
         Printer.CurrentY = pY
         Printer.CurrentX = 4
         Printer.Print DtpFecha.Value
         Printer.CurrentY = pY
         Printer.CurrentX = 6.5
         Printer.Print "A Cuenta"
         Printer.CurrentY = pY
         Printer.CurrentX = 11
         Printer.Print Format(rAde!Sucursal, "0000") & "-"
         Printer.CurrentY = pY
         Printer.CurrentX = 12
         Printer.Print Format(rAde!Numero, "00000000")
         Printer.CurrentY = pY
         Printer.CurrentX = 15
         Printer.Print ""
         Printer.CurrentY = pY
         Printer.CurrentX = 18 - Printer.TextWidth(Format(rAde!importe, "#0.00"))
         Printer.Print Format(rAde!importe, "#0.00")
         cSaldo = cSaldo + rAde!importe
      End If
      pY = pY + 0.5
     Printer.CurrentY = pY
     Printer.CurrentX = 8
     Printer.Print "T O T A L   D E L  C O B R O :"
     Printer.CurrentY = pY
     Printer.CurrentX = 17.5 - Printer.TextWidth(Format(cSaldo, "#0.00"))
     Printer.Print Format(cSaldo, "#0.00")
     rValo.MoveFirst
     pY = pY + 1
     Do While Not rValo.EOF
        Printer.CurrentY = pY
        Printer.CurrentX = 1
        Printer.Print rValo!formapago
        Printer.CurrentY = pY
        Printer.CurrentX = 7
        If rValo!NumeroCheque <> "" And Not IsNull(rValo!Banco) Then
           Printer.Print cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & rValo!Banco) & " " & rValo!NumeroCheque & " " & rValo!FechaAcreditacion
        Else
           If rValo!Cupon <> "" Then
              Printer.Print cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & rValo!Banco) & " " & rValo!NumeroTarjeta & " " & rValo!Cupon
           End If
        End If
        Printer.CurrentY = pY
        Printer.CurrentX = 18 - Printer.TextWidth(Format(rValo!importe, "#0.00"))
        Printer.Print Format(rValo!importe, "#0.00")
        rValo.MoveNext
        pY = pY + 0.5
     Loop
     pY = pY + 1
     Printer.CurrentY = pY
     Printer.CurrentX = 1
     Printer.Print "CONFECCIONO :" & rUsu!Nombre
     Printer.CurrentY = pY
     Printer.CurrentX = 10
     pY = pY + 0.5
     Printer.CurrentY = pY
     Printer.CurrentX = 1
     Printer.Print "SALDO PENDIENTE :" & LblSaldo.Caption
     Printer.EndDoc
  Next nI
 ' Limpiar

End Sub

Private Sub TxtSucursal_GotFocus()
  TxtSucursal.SelStart = 0
  TxtSucursal.SelLength = Len(TxtSucursal.text)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtSucursal_LostFocus()
  If TxtSucursal.text = "" Then TxtSucursal.text = 1
  TxtSucursal.text = Format(TxtSucursal.text, "0000")
End Sub

Private Sub TxtNumero_GotFocus()
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtNumero_LostFocus()
  If TxtNumero.text = "" Then TxtNumero.text = 0
  TxtNumero.text = Format(TxtNumero.text, "00000000")
End Sub

