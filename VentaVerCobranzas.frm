VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VentaVerCobranzas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Cobranzas"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Terminar"
      Height          =   495
      Left            =   6705
      TabIndex        =   15
      Top             =   5400
      Width           =   1200
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   5445
      TabIndex        =   14
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Formas de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   225
      TabIndex        =   3
      Top             =   3060
      Width           =   7710
      Begin MSDataGridLib.DataGrid Grid2 
         Height          =   1575
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2778
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
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
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
         Height          =   255
         Left            =   6045
         TabIndex        =   7
         Top             =   1890
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "Total Cancelado :"
         Height          =   240
         Left            =   4710
         TabIndex        =   6
         Top             =   1905
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comprobantes Afectados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   225
      TabIndex        =   2
      Top             =   1035
      Width           =   7740
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   225
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   2778
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
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label LblFecha 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblFecha"
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
      Left            =   6735
      TabIndex        =   13
      Top             =   150
      Width           =   1140
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha :"
      Height          =   240
      Left            =   6120
      TabIndex        =   12
      Top             =   150
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
      Left            =   4680
      TabIndex        =   11
      Top             =   150
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "Nº :"
      Height          =   225
      Left            =   4305
      TabIndex        =   10
      Top             =   150
      Width           =   390
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
      Height          =   315
      Left            =   1770
      TabIndex        =   9
      Top             =   150
      Width           =   2460
   End
   Begin VB.Label Label3 
      Caption         =   "Comprobante :"
      Height          =   210
      Left            =   195
      TabIndex        =   8
      Top             =   150
      Width           =   1125
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
      Left            =   945
      TabIndex        =   1
      Top             =   555
      Width           =   3300
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   225
      Left            =   225
      TabIndex        =   0
      Top             =   555
      Width           =   675
   End
End
Attribute VB_Name = "VentaVerCobranzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As Recordset
Dim Rs1 As Recordset
Dim Rs2 As Recordset

Private Sub CmdImprimir_Click()
  Dim cCli As ClsClienteL, pY As Single, cRsl As ClsLectura
  Dim cNum As ClsComprobantesL, rValo As Recordset
  Dim cImpo As Currency, cSaldo As Currency
  
  
  Dim Rp As ADODB.Recordset
  
  Set cCli = New ClsClienteL
  Set cRsl = New ClsLectura
  Set cNum = New ClsComprobantesL
  
  cSaldo = 0
  Printer.font.Name = "Ms Sans Seriff"
  Printer.font.size = 14
  Printer.font.bold = True
  Printer.ScaleMode = 7
  Printer.CurrentY = 1.5
  Printer.CurrentX = 1
  Printer.Print cRsl.TraerValorDeUnCampo("Empresa", "Razon", "") & " - " & cRsl.TraerValorDeUnCampo("Sucursales", "Descripcion", "Id=" & nSucursal)
  Printer.font.size = 14
  Printer.font.bold = False
  Printer.CurrentY = 2.5
  Printer.CurrentX = 1
  Printer.Print "RECIBO DE COBRANZAS"
  Printer.font.size = 10
  Printer.CurrentY = 3.5
  Printer.CurrentX = 13
  Printer.Print "Nro.de Documento :" & LblNumero.Caption
  Printer.CurrentY = 4
  Printer.CurrentX = 13
  Printer.Print "Fecha de Pago :" & LblFecha.Caption
  Printer.CurrentY = 4.5
  Printer.CurrentX = 1
  Printer.Print "Cliente :" & LblCliente.Caption
  Printer.CurrentY = 5
  Printer.CurrentX = 1
  Printer.Print "Domicilio :" & cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & VentaCobranzaAnular.TxtCliente.text & "'")
  Printer.CurrentY = 5.5
  Printer.CurrentX = 1
  Printer.Print "Localidad :"
  Printer.CurrentY = 6
  Printer.CurrentX = 1
  Printer.Print "CUIT :" & cRsl.TraerValorDeUnCampo("Proveedores", "Cuit", "Proveedor='" & VentaCobranzaAnular.TxtCliente.text & "'")
  Printer.CurrentY = 6.5
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  Printer.CurrentY = 7
  Printer.CurrentX = 1
  Printer.Print "                           FECHA                      TIPO                        NUMERO              CUOTA               IMPORTE"
  Printer.CurrentY = 7.5
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  Printer.CurrentY = 8
  pY = 8
  Rs1.MoveFirst

  Do While Not Rs1.EOF
     Set Rp = cRsl.TraerComproPago(Rs1!Movimiento)
'     Printer.CurrentX = 1
'     Printer.Print cNum.NumeroAsiento
     Printer.CurrentY = pY
     Printer.CurrentX = 4
     Printer.Print Rs1!fecha
     Printer.CurrentY = pY
     Printer.CurrentX = 6.5
     Printer.Print Rs1!Comprobante
     Printer.CurrentY = pY
     Printer.CurrentX = 11
     Printer.Print Format(Rs1!Sucursal, "0000") & "-"
     Printer.CurrentY = pY
     Printer.CurrentX = 12
     Printer.Print Format(Rs1!Numero, "00000000")
     Printer.CurrentY = pY
     Printer.CurrentX = 15
     Printer.Print cRsl.TraerValorDeUnCampo("Aplicaciones", "Cuota", "Venta=" & Rs1!Movimiento & " And Sucursal=" & Rp!Sucursal & " AND Numero=" & Rp!Numero)
     Printer.CurrentY = pY
     
   ' cImpo = cRsl.TraerValorDeUnCampo("Aplicaciones", "Importe", "Venta=" & Rs1!Movimiento & " And Sucursal=" & Rp!Sucursal & " AND Numero=" & Rp!Numero)
     
     Printer.CurrentX = 18 - Printer.TextWidth(Format(Rs1!Cancelado))
     Printer.Print Format(Rs1!Cancelado)
     cSaldo = cSaldo + Rs1!Cancelado
     Rs1.MoveNext
     pY = pY + 0.5
  Loop
  
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 8
  Printer.Print "T O T A L   A  C O B R A R :"
  Printer.CurrentY = pY
  Printer.CurrentX = 18 - Printer.TextWidth(Format(cSaldo, "#0.00"))
  Printer.Print Format(cSaldo, "#0.00")

  Set rValo = cRsl.TraerRsCondi("Valores", "Id", "Venta=" & Rs!Movimiento)

  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "CANCELACION :"
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
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 8
  Printer.Print "T O T A L   D E L  C O B R O :"
  Printer.CurrentY = pY
  Printer.CurrentX = 18 - Printer.TextWidth(Format(LblTotal.Caption, "#0.00"))
  Printer.Print Format(cSaldo, "#0.00")
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "CONFECCIONO :"
  Printer.CurrentY = pY
  Printer.CurrentX = 10
  Printer.Print "AUTORIZO :_____________________________________________"
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "Recibimos de " & Trim(LblCliente.Caption) & " la cantidad de $ " & Format(LblTotal.Caption, "#0.00")
  pY = pY + 0.5
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "En todo de acuerdo con la presente liquidación."
  pY = pY + 1
  Printer.CurrentY = pY
  Printer.CurrentX = 1
  Printer.Print "ACLARACION_______________________________  DOC:_____________________ FIRMA_________________________"
  Printer.EndDoc

End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Limpiar
  CrearRsComprobantes
  CrearRsFormaPago
  Set Rs = VentaCobranzaAnular.RsAp
  LinkearTexto
End Sub

Private Sub LinkearTexto()
  LblFecha.Caption = Rs!fecha
  LblComprobante.Caption = Rs!Comprobante
  LblNumero.Caption = Format(Rs!Sucursal, "0000") & "-" & Format(Rs!Numero, "00000000")
  LblTotal.Caption = Format(Rs!importe, "0.00")
  LblCliente.Caption = VentaCobranzaAnular.LblCliente.Caption
  LlenarRs2
  LlenarRs1
  
End Sub

Private Sub LlenarRs1()
  Dim cRsl As ClsLectura, rAux As Recordset
  
  Set cRsl = New ClsLectura
  Set rAux = cRsl.TraerRsCondi("DetallesComprobantes", "Id", "Movimiento=" & Rs!Movimiento)
  Do While Not rAux.EOF
     Rs1.AddNew
     Rs1!fecha = cRsl.TraerValorDeUnCampo("CabComprobantes", "Fecha", "Movimiento=" & rAux!Tasa)
     Rs1!Comprobante = cRsl.TraerValorDeUnCampo("Comprobantes", "Descripcion", "Id=" & cRsl.TraerValorDeUnCampo("CabComprobantes", "Comprobante", "Movimiento=" & rAux!Tasa))
     Rs1!Sucursal = Format(cRsl.TraerValorDeUnCampo("CabComprobantes", "Sucursal", "Movimiento=" & rAux!Tasa), "0000")
     Rs1!Numero = Format(cRsl.TraerValorDeUnCampo("CabComprobantes", "Numero", "Movimiento=" & rAux!Tasa), "00000000")
     Rs1!importe = Format(cRsl.TraerValorDeUnCampo("CabComprobantes", "Debe", "Movimiento=" & rAux!Tasa), "0.00")
     Rs1!Cancelado = Format(rAux!PrecioTotal, "0.00")
     Rs1!Movimiento = rAux!Tasa
     Rs1.Update
     rAux.MoveNext
  Loop
  
End Sub

Private Sub LlenarRs2()
  Dim cRsl As ClsLectura, rAux As Recordset
  
  Set cRsl = New ClsLectura
  Set rAux = cRsl.TraerRsCondi("Valores", "Id", "Venta=" & Rs!Movimiento)
  
  Do While Not rAux.EOF
     Rs2.AddNew
     Rs2!Forma = rAux!formapago
     Rs2!Numero = rAux!NumeroCheque
     If Not IsNull(rAux!Banco) Then
        Rs2!Banco = cRsl.TraerValorDeUnCampo("Bancos", "Descripcion", "Banco=" & rAux!Banco)
     End If
     Rs2!fecha = rAux!FechaAcreditacion
     Rs2!importe = rAux!importe
     Rs2.Update
     rAux.MoveNext
  Loop
  
End Sub
Private Sub Limpiar()
  LblCliente.Caption = ""
  LblTotal.Caption = "0.00"
  LblComprobante.Caption = ""
  LblNumero.Caption = "0000-00000000"
  LblFecha.Caption = ""
End Sub
Private Sub CrearRsComprobantes()
  Set Rs1 = New Recordset
  Rs1.Fields.Append "Fecha", adDate
  Rs1.Fields.Append "Comprobante", adVarChar, 30
  Rs1.Fields.Append "Sucursal", adInteger
  Rs1.Fields.Append "Numero", adInteger
  Rs1.Fields.Append "Importe", adCurrency
  Rs1.Fields.Append "Cancelado", adCurrency
  Rs1.Fields.Append "Movimiento", adInteger
  Rs1.Open
  CabGrid1
End Sub

Private Sub CrearRsFormaPago()
  Set Rs2 = New Recordset
  Rs2.Fields.Append "Forma", adVarChar, 30
  Rs2.Fields.Append "Numero", adVarChar, 20, adFldIsNullable
  Rs2.Fields.Append "Banco", adVarChar, 50, adFldIsNullable
  Rs2.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs2.Fields.Append "Importe", adCurrency
  Rs2.Open
  CabGrid2
End Sub

Private Sub CabGrid1()
  Set Grid1.DataSource = Rs1
  Grid1.HeadFont.size = 10
  Grid1.HeadFont.bold = True
  With Grid1
      .Columns(0).Width = 1000
      .Columns(0).Caption = "Fecha"
      .Columns(1).Width = 1800
      .Columns(1).Caption = "Comprobante"
      .Columns(2).Caption = "Sucursal"
      .Columns(2).NumberFormat = "000000"
      .Columns(2).Width = 1000
      .Columns(2).Alignment = dbgRight
      .Columns(3).Width = 1000
      .Columns(3).Caption = "Numero"
      .Columns(3).NumberFormat = "000000"
      .Columns(3).Alignment = dbgRight
      .Columns(4).Width = 1150
      .Columns(4).Caption = "Importe"
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(4).Alignment = dbgRight
      .Columns(5).Width = 1200
      .Columns(5).Caption = "Cancelado"
      .Columns(5).NumberFormat = nCantDecimales
      .Columns(5).Alignment = dbgRight
      .Columns(6).Visible = False
  End With
End Sub

Private Sub CabGrid2()
  Set Grid2.DataSource = Rs2
  Grid2.HeadFont.size = 10
  Grid2.HeadFont.bold = True
  With Grid2
      .Columns(0).Width = 2200
      .Columns(0).Caption = "F.de Pago"
      .Columns(1).Width = 1100
      .Columns(1).Caption = "Número"
      .Columns(2).Caption = "Banco"
      .Columns(2).Width = 1600
      .Columns(3).Width = 1000
      .Columns(3).Caption = "Fecha"
      .Columns(4).Width = 1200
      .Columns(4).Caption = "Importe"
      .Columns(4).NumberFormat = nCantDecimales
      .Columns(4).Alignment = dbgRight
 End With
End Sub

Private Sub Label5_Click()

End Sub
