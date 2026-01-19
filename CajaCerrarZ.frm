VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CajaCerrarZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cerrar Caja en Modo Z"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   7680
   Begin VB.CheckBox ChkTurno 
      Caption         =   "Turno Abierto"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2415
   End
   Begin VB.ComboBox CmbCaja 
      Height          =   315
      Left            =   5475
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   135
      Width           =   2070
   End
   Begin VB.CommandButton BtnBuscar 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7200
      TabIndex        =   6
      Top             =   600
      Width           =   345
   End
   Begin VB.CommandButton CmbBotones 
      Caption         =   "Salir"
      Height          =   855
      Index           =   4
      Left            =   6480
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmbBotones 
      Caption         =   "Listado de Cheques y Tarjetas"
      Height          =   855
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmbBotones 
      Caption         =   "Listado de Movimientos"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   360
      Left            =   5475
      TabIndex        =   0
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52822017
      CurrentDate     =   39377
   End
   Begin VB.Label Label2 
      Caption         =   "FECHA :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "CAJA DIARIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      TabIndex        =   1
      Top             =   165
      Width           =   2355
   End
End
Attribute VB_Name = "CajaCerrarZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rPrin As Recordset
Dim sw As Boolean

Private Sub Botones(pBtn1 As Boolean, pBtn2 As Boolean, pBtn3 As Boolean, pBtn4 As Boolean, pBtn5 As Boolean, pBtn6 As Boolean)
  CmbBotones(0).Enabled = pBtn1
  CmbBotones(1).Enabled = pBtn2
End Sub

Private Sub BtnBuscar_Click()
   CmbBotones(0).Enabled = True
   Botones True, True, True, True, True, False
End Sub

Private Sub CmbBotones_Click(Index As Integer)
  Select Case Index
     Case 0
        ListadoMovimientos
     Case 1
        ListadoChequesTarjetas
     Case 2
        Grabar
        CajaDiaria
     Case 3
         CajaApertura.Show 1
         If nDat <> 0 Then
            CmbBotones(3).Enabled = False
            Botones True, True, True, False, True, False
         End If
     Case 4
       Salir
  End Select
  
End Sub
Private Sub CajaDiaria()
  Dim Rx As Recordset, cRle As ClsComprobantesL, cRsl As ClsLectura, i As Integer
  
  Set cRsl = New ClsLectura
  
  Set Rx = cRsl.TraerRsCondi("CajaDiaria", "Id", "Fecha='" & FechaAmericana(DtpFecha.Value) & "' and Suc=" & nSucursal)
  Set rPrin = Rx
  
  nImpr = 52
  frmImpresor.Show

End Sub

Private Sub ListadoChequesTarjetas()
  Dim Rx As Recordset, cRle As ClsComprobantesL, cRsl As ClsLectura, i As Integer

  Set rPrin = New Recordset
  Set Rx = New Recordset
  Set cRle = New ClsComprobantesL
  Set cRsl = New ClsLectura
  
  rPrin.Fields.Append "Id", adInteger
  rPrin.Fields.Append "Nombre", adVarChar, 50, adFldIsNullable
  rPrin.Fields.Append "Tipo", adVarChar, 100, adFldIsNullable
  rPrin.Fields.Append "Banco", adVarChar, 30, adFldIsNullable
  rPrin.Fields.Append "Numero", adVarChar, 15, adFldIsNullable
  rPrin.Fields.Append "Acreditacion", adDate, , adFldIsNullable
  rPrin.Fields.Append "Tarjeta", adVarChar, 30, adFldIsNullable
  rPrin.Fields.Append "Vencimiento", adDate, , adFldIsNullable
  rPrin.Fields.Append "Autorizacion", adVarChar, 30, adFldIsNullable
  rPrin.Fields.Append "Cupon", adVarChar, 30, adFldIsNullable
  rPrin.Fields.Append "Importe", adCurrency
  rPrin.Open
  
  i = 1
  Set Rx = cRle.TraerChequeTarjetas(DtpFecha.Value, CmbCaja.ItemData(CmbCaja.ListIndex), ChkTurno.Value)
  Do While Not Rx.EOF
     If cRsl.TraerValorDeUnCampo("Caja", "Abierta", "Movimiento=" & Rx!Venta) = "S" Then
     rPrin.AddNew
     rPrin!ID = i
     rPrin!Nombre = Rx!Nombre
     rPrin!tipo = Rx!formapago & " - " & Rx!Comprobante
     If cRsl.TraerValorDeUnCampo("CondVenta", "Tipo", "CondVta=" & Rx!CodPago) = 2 Then
        rPrin!Banco = Format(Rx!sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
     Else
        rPrin!Banco = Rx!Descripcion
     End If
    ' rPrin!Banco = Rx!Descripcion
     rPrin!Numero = Rx!NumeroCheque
     rPrin!Acreditacion = Rx!FechaAcreditacion
     rPrin!Tarjeta = Rx!NumeroTarjeta
     rPrin!Vencimiento = Rx!FechaVencimiento
     rPrin!Autorizacion = Rx!Autorizacion
     rPrin!Cupon = Rx!Cupon
     rPrin!importe = Rx!importe
     rPrin.Update
     End If
     Rx.MoveNext
  Loop
  
  nImpr = 51
  frmImpresor.Show

End Sub

Private Sub ListadoMovimientos()
  Dim Rx As Recordset, cRle As ClsComprobantesL, cRsl As ClsLectura, i As Integer, b As Integer
  
  Set rPrin = New Recordset
  Set Rx = New Recordset
  Set cRle = New ClsComprobantesL
  Set cRsl = New ClsLectura
  
  rPrin.Fields.Append "Id", adInteger
  rPrin.Fields.Append "Comprobante", adVarChar, 20, adFldIsNullable
  rPrin.Fields.Append "Observacion", adVarChar, 60, adFldIsNullable
  rPrin.Fields.Append "Numero", adVarChar, 14, adFldIsNullable
  rPrin.Fields.Append "Ingreso", adCurrency
  rPrin.Fields.Append "Egreso", adCurrency
  rPrin.Open
  

  rPrin.AddNew
  rPrin!ID = 1
  rPrin!Comprobante = "SALDO INICIAL :"
  rPrin!Numero = ""
  rPrin!Observacion = ""
  rPrin!Ingreso = cRsl.TraerValorDeUnCampo("SaldoInicial", "Saldo", "Fecha='" & FechaAmericana(DtpFecha.Value) & "' and Suc=" & nSucursal)
  rPrin!Egreso = 0
  rPrin.Update
  
  
  Set Rx = cRle.IngresoCajaRes(DtpFecha.Value, CmbCaja.ItemData(CmbCaja.ListIndex))
  i = i + 1
  Do While Not Rx.EOF
     rPrin.AddNew
     rPrin!ID = i
     rPrin!Comprobante = Rx!Comprobante
     rPrin!Observacion = cRsl.TraerValorDeUnCampo("CabComprobantes", "Motivo", "Movimiento=" & Rx!Movimiento)
     rPrin!Numero = Format(Rx!sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
     rPrin!Ingreso = Rx!Debe
     rPrin!Egreso = Rx!Haber
     rPrin.Update
     Rx.MoveNext
     i = i + 1
  Loop
  
  Set Rx = cRle.MovimientosDeCaja(DtpFecha, CmbCaja.ItemData(CmbCaja.ListIndex), ChkTurno.Value)

  i = i + 1
  Dim sHora As String
  
  Do While Not Rx.EOF
     If cRsl.TraerValorDeUnCampo("Comprobantes", "Caja", "Comprobante='" & Rx!Comprobante & "'") <> "No" Then
        If cRsl.TraerValorDeUnCampo("Caja", "Abierta", "Movimiento=" & Rx!Movimiento) = "S" Then
           If cRsl.TraerValorDeUnCampo("Valores", "CodPago", "Venta=" & Rx!Movimiento) = 1 Then
              rPrin.AddNew
              rPrin!ID = i
              sHora = Left(cRsl.TraerValorDeUnCampo("Caja", "Hora", "Movimiento=" & Rx!Movimiento), 5)
              rPrin!Comprobante = Rx!Comprobante & " - " & sHora
              rPrin!Observacion = cRsl.TraerValorDeUnCampo("Caja", "Nombre", "Movimiento=" & Rx!Movimiento)
              rPrin!Numero = Format(Rx!sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
              rPrin!Ingreso = cRsl.TraerValorDeUnCampo("Caja", "Debe", "FormaPago >= 'EFECTIVO'  AND Movimiento=" & Rx!Movimiento) + IIf(cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & Rx!Comprobante & "'") = 5, 0, 0)
              rPrin!Egreso = IIf(cRsl.TraerValorDeUnCampo("Comprobantes", "TipoOperacion", "Comprobante='" & Rx!Comprobante & "'") = 5, Rx!Haber, 0)
              rPrin.Update
              i = i + 1
          End If
        End If
     End If
     Rx.MoveNext
  Loop
  
  Set Rx = cRle.EgresoCajaRes(DtpFecha.Value, CmbCaja.ItemData(CmbCaja.ListIndex))
  i = i + 1
  Do While Not Rx.EOF
     rPrin.AddNew
     rPrin!ID = i
     rPrin!Comprobante = Rx!Comprobante
     rPrin!Observacion = cRsl.TraerValorDeUnCampo("CabComprobantes", "Motivo", "Movimiento=" & Rx!Movimiento)
     rPrin!Numero = Format(Rx!sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
     rPrin!Ingreso = Rx!Debe
     rPrin!Egreso = Rx!Haber
     rPrin.Update
     Rx.MoveNext
     i = i + 1
  Loop
  
  Set Rx = cRle.EgresoCajaPagos(DtpFecha.Value, CmbCaja.ItemData(CmbCaja.ListIndex))
  i = i + 1
  Do While Not Rx.EOF
     rPrin.AddNew
     rPrin!ID = i
     rPrin!Comprobante = Rx!Comprobante
     rPrin!Observacion = Rx!Nombre
     rPrin!Numero = Format(Rx!sucursal, "0000") & "-" & Format(Rx!Numero, "00000000")
     rPrin!Ingreso = Rx!Debe
     rPrin!Egreso = Rx!Haber
     rPrin.Update
     Rx.MoveNext
     i = i + 1
  Loop
  
  nImpr = 50
  frmImpresor.Show

End Sub
Private Sub Salir()
  Unload Me
End Sub

Private Sub Grabar()
'  Dim Rx1 As Recordset, cRsl As ClsLectura, cRse As ClsEscritura
'
'
'  Set Rx1 = New Recordset
'  Set cRsl = New ClsLectura
'  Set cRse = New ClsEscritura
'
'  Set Rx1 = cRsl.RsVacio("CajaDiaria", "Id", "N")
'  Rx1.AddNew
'  Rx1!Suc = nSucursal
'  Rx1!fecha = DtpFecha.Value
'  Rx1!Monedas = TxtValor(0).text
'  Rx1!x2 = TxtValor(1).text
'  Rx1!X5 = TxtValor(2).text
'  Rx1!X10X20 = TxtValor(3).text
'  Rx1!X50 = TxtValor(4).text
'  Rx1!X100 = TxtValor(5).text
'  Rx1!X10X20b = TxtValor(11).text
'  Rx1!Dolares = TxtDolares(0).text
'  Rx1!COTIZACION = TxtDolares(1).text
'  Rx1!TotalDolares = TxtValor(6).text
'  Rx1!Cheques = TxtValor(7).text
'  Rx1!Tarjetas = TxtValor(8).text
'  Rx1!Vales = TxtValor(9).text
'  Rx1!otros = TxtValor(10).text
'  Rx1.Update

  db.Open sDb
  db.Execute "DELETE  FROM CajaDiaria WHERE Fecha='" & FechaAmericana(DtpFecha.Value) & "' and Suc=" & nSucursal
  db.Close
'  cRse.Actualizar "CajaDiariaAgregar", Rx1
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

Private Sub Form_Load()
  sw = False
  Botones False, False, False, False, False, True
  DtpFecha.Value = Date
  CargarCombos
End Sub
Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
    
  If nCatUser < 2 Then
     cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", ""
     CmbCaja.AddItem "TODAS"
     CmbCaja.ItemData(CmbCaja.NewIndex) = 0
     CmbCaja.text = "TODAS"
  Else
     cRsl.CargaCombo CmbCaja, "Cajas", "Caja", "Descripcion", "Caja=" & cRsl.RegPorDefecto("Caja")
     CmbCaja.ListIndex = 0
  End If
  Set cRsl = New ClsLectura
End Sub


