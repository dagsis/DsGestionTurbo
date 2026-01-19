VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmInformeCobranzas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Cobranzas"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   3825
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
      Height          =   2250
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   3510
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   450
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   53542913
         CurrentDate     =   36983
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   345
         Left            =   1140
         TabIndex        =   3
         Top             =   1185
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   53542913
         CurrentDate     =   36983
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   435
         TabIndex        =   5
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   420
         TabIndex        =   4
         Top             =   1185
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Aplicar"
      Height          =   420
      Left            =   2760
      TabIndex        =   0
      Top             =   2505
      Width           =   945
   End
End
Attribute VB_Name = "FrmInformeCobranzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset
Public Rs3 As ADODB.Recordset
Private Sub CmdImprimir_Click()
  Me.MousePointer = 11
  CrearCabeza
  CrearRsComprobantes
  CrearRsFormaPago
  LlenarRs3
  nImpr = 53
  FrmImpresor.Show
  Me.MousePointer = 0
End Sub

Private Sub LlenarRs3()
  Dim rAux As ADODB.Recordset, cX1 As ClsComprobantesL, cRsl As ClsLectura, x As Integer
  
  Set cX1 = New ClsComprobantesL
  Set rAux = cX1.TraerCobranzaLista(DtpDesde.Value, DtpHasta.Value)
  
  Set cRsl = New ClsLectura
  
  Do While Not rAux.EOF
      ' Cabeza
      LlenarRs1 rAux!Movimiento
      If Rs1.RecordCount <> 0 Then
         Rs1.MoveFirst
         Do While Not Rs1.EOF
            Rs3.AddNew
            Rs3!Documento = Format(rAux!Sucursal, "0000") & "-" & Format(rAux!Numero, "00000000")
            Rs3!FechaCob = rAux!fecha
            Rs3!Cliente = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & rAux!Cliente & "'")
            Rs3!Domicilio = cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & rAux!Cliente & "'")
            Rs3!Localidad = ""
      ' Facturas
            Rs3!fecha = Rs1!fecha
            Rs3!Comprobante = Rs1!Comprobante
            Rs3!Sucursal = Rs1!Sucursal
            Rs3!Numero = Rs1!Numero
            Rs3!importe = Rs1!importe
            Rs3!Cancelado = Rs1!Cancelado
            Rs3!Movimiento = Rs1!Movimiento
      ' Formas
            Rs3!Forma = ""
            Rs3!NumeroCom = 0
            Rs3!Banco = ""
            Rs3!FechaCom = Date
            Rs3!ImportePago = 0
            Rs3!Aplicacion = True
            Rs3.Update
            Rs1.MoveNext
         Loop
         LlenarRs2 rAux!Movimiento
         Rs2.MoveFirst
         Do While Not Rs2.EOF
            Rs3.AddNew
            Rs3!Documento = Format(rAux!Sucursal, "0000") & "-" & Format(rAux!Numero, "00000000")
            Rs3!FechaCob = rAux!fecha
            Rs3!Cliente = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & rAux!Cliente & "'")
            Rs3!Domicilio = cRsl.TraerValorDeUnCampo("Clientes", "Domicilio", "Cliente='" & rAux!Cliente & "'")
            Rs3!Localidad = ""
          ' Facturas
            Rs3!fecha = Date
            Rs3!Comprobante = ""
            Rs3!Sucursal = 0
            Rs3!Numero = 0
            Rs3!importe = 0
            Rs3!Cancelado = 0
            Rs3!Movimiento = 0
         ' Formas
            Rs3!Forma = Rs2!Forma
            Rs3!NumeroCom = Rs2!Numero
            Rs3!Banco = Rs2!Banco
            Rs3!FechaCom = Rs2!fecha
            Rs3!ImportePago = Rs2!importe
            Rs3!Aplicacion = False
            Rs3.Update
           Rs2.MoveNext
         Loop
         If Rs1.RecordCount <> 0 Then
            Rs1.MoveLast
            For x = Rs1.RecordCount - 1 To 0 Step -1
               Rs1.Delete
               Rs1.MovePrevious
            Next x
        End If
        If Rs2.RecordCount <> 0 Then
           Rs2.MoveLast
           For x = Rs2.RecordCount - 1 To 0 Step -1
               Rs2.Delete
               Rs2.MovePrevious
          Next x
       End If
     End If
     rAux.MoveNext
  Loop
 
End Sub

Private Sub Form_Load()
  DtpDesde.Value = Date
  DtpHasta.Value = Date
End Sub

Private Sub CrearCabeza()
  Set Rs3 = New ADODB.Recordset
  
  Rs3.Fields.Append "Documento", adVarChar, 13
  Rs3.Fields.Append "FechaCob", adDBDate
  Rs3.Fields.Append "Cliente", adVarChar, 50
  Rs3.Fields.Append "Domicilio", adVarChar, 50, adFldIsNullable
  Rs3.Fields.Append "Localidad", adVarChar, 50, adFldIsNullable
  Rs3.Fields.Append "Fecha", adDate
  Rs3.Fields.Append "Comprobante", adVarChar, 30
  Rs3.Fields.Append "Sucursal", adInteger
  Rs3.Fields.Append "Numero", adInteger
  Rs3.Fields.Append "Importe", adCurrency
  Rs3.Fields.Append "Cancelado", adCurrency
  Rs3.Fields.Append "Movimiento", adInteger
  Rs3.Fields.Append "Forma", adVarChar, 30
  Rs3.Fields.Append "NumeroCom", adVarChar, 40, adFldIsNullable
  Rs3.Fields.Append "Banco", adVarChar, 30, adFldIsNullable
  Rs3.Fields.Append "FechaCom", adDate, , adFldIsNullable
  Rs3.Fields.Append "ImportePago", adCurrency
  Rs3.Fields.Append "Aplicacion", adBoolean
  Rs3.Open
  
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
End Sub

Private Sub CrearRsFormaPago()
  Set Rs2 = New Recordset
  Rs2.Fields.Append "Forma", adVarChar, 30
  Rs2.Fields.Append "Numero", adVarChar, 40, adFldIsNullable
  Rs2.Fields.Append "Banco", adVarChar, 30, adFldIsNullable
  Rs2.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs2.Fields.Append "Importe", adCurrency
  Rs2.Open
End Sub

Private Sub LlenarRs1(pMovi As Long)
  Dim cRsl As ClsLectura, rAux As Recordset, x As Integer

  Set cRsl = New ClsLectura
  Set rAux = cRsl.TraerRsCondi("DetallesComprobantes", "Id", "Movimiento=" & pMovi)
  
  If rAux.RecordCount <> 0 Then
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
  End If
End Sub

Private Sub LlenarRs2(pMovi As Long)
  Dim cRsl As ClsLectura, rAux As Recordset, x As Integer

  Set cRsl = New ClsLectura
  Set rAux = cRsl.TraerRsCondi("Valores", "Id", "Venta=" & pMovi)
   
   
'  For X = rAux.RecordCount To 1 Step -1
'      rAux.Delete
'      rAux.MovePrevious
'      If rAux.BOF Then
'         Exit For
'      End If
'  Next X
  
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

