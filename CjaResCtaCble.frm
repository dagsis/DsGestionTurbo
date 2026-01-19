VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form CjaResCtaCble 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Cuentas Contables"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   4650
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   165
      TabIndex        =   11
      Top             =   3210
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen"
      Height          =   3345
      Left            =   105
      TabIndex        =   1
      Top             =   75
      Width           =   4410
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1110
         Left            =   150
         TabIndex        =   7
         Top             =   1545
         Width           =   4035
         Begin VB.CheckBox ChkTodas 
            Caption         =   "Todas las Cuentas"
            Height          =   195
            Left            =   105
            TabIndex        =   10
            Top             =   810
            Width           =   2145
         End
         Begin VB.ComboBox CmbCuenta 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   330
            Width           =   3720
         End
         Begin VB.CheckBox ChkDetalle 
            Caption         =   "Cuenta Con Detalle"
            Height          =   240
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   2250
         End
      End
      Begin VB.CheckBox ChkTodos 
         Caption         =   "Todas las Operaciones"
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   2760
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   285
         Left            =   1890
         TabIndex        =   2
         Top             =   930
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   52363265
         CurrentDate     =   38593
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   285
         Left            =   1890
         TabIndex        =   3
         Top             =   375
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   52363265
         CurrentDate     =   38593
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha :"
         Height          =   210
         Left            =   570
         TabIndex        =   5
         Top             =   375
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha :"
         Height          =   240
         Left            =   555
         TabIndex        =   4
         Top             =   915
         Width           =   1110
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   450
      Left            =   3585
      TabIndex        =   0
      Top             =   3495
      Width           =   900
   End
End
Attribute VB_Name = "CjaResCtaCble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As ADODB.Recordset

Private Sub ChkDetalle_Click()
  CmbCuenta.Enabled = False
  ChkTodas.Enabled = False
  If ChkDetalle.Value = 1 Then
     CmbCuenta.Enabled = True
     ChkTodas.Enabled = True
     ChkTodas.Value = 1
  End If
    CmbCuenta.ListIndex = -1
End Sub

Private Sub ChkTodas_Click()
  CmbCuenta.Enabled = True
  CmbCuenta.ListIndex = 0
  If ChkTodas.Value = 1 Then
     CmbCuenta.ListIndex = -1
     CmbCuenta.Enabled = False
     ChkTodos.Value = 1
     ChkTodos.Enabled = False
  End If
End Sub

Private Sub Command1_Click()
  Dim Rxa As ADODB.Recordset, cRx As ClsComprobantesL, cRl As ClsLectura, j As Integer
  
  j = 1
  If ChkDetalle.Value = 0 Then
  CrearRs1
  
  Set cRx = New ClsComprobantesL
  Set cRl = New ClsLectura
  
  ProgressBar1.Value = 0

  Set Rxa = cRx.TraerResPlanCta(DTPDesde.Value, DTPHasta.Value, ChkTodos.Value)
  Dim bModi As Boolean
  If Rxa.RecordCount <> 0 Then
     ProgressBar1.Max = Rxa.RecordCount
     Do While Not Rxa.EOF
        If Not IsNull(Rxa!Grupo) Then
           If (Rxa!Movi = "Debito" Or Rxa!Movi = "Resta") And Rxa!Grupo = "EGRESOS" Then
              nTotal = -Rxa!SumaDePrecioTotal
              bModi = True
           End If
              If Rxa!Movi = "Resta" And Rxa!Grupo = "EGRESOS" Then
                 nTotal = -Rxa!SumaDePrecioTotal
                 bModi = True
           End If

           If Rxa!Movi = "Credito" And Rxa!Grupo = "INGRESOS" Then
              nTotal = -Rxa!SumaDePrecioTotal
              bModi = True
           End If
           If Rxa!Movi = "Debito" And Rxa!Grupo = "INGRESOS" Then
              nTotal = Rxa!SumaDePrecioTotal
              bModi = True
           End If
           If (Rxa!Movi = "Credito" Or Rxa!Movi = "Resta") And Rxa!Grupo = "EGRESOS" Then
              nTotal = Rxa!SumaDePrecioTotal
              bModi = True
           End If
           If bModi = True Then
              rs.AddNew
              bModi = False
              rs!Grupo = Rxa!Grupo
              rs!NCta = Rxa!Id
              rs!Cuenta = Rxa!Cuenta
              rs!importe = nTotal
              rs.Update
            End If
        End If
        Rxa.MoveNext
        ProgressBar1.Value = i
        i = i + 1
     Loop
  End If
  
  nImpr = 44
  Else
     CrearRs2
       Set cRx = New ClsComprobantesL
       Set cRl = New ClsLectura
          Dim pCuenta As Integer
          
          pCuenta = -1
          If CmbCuenta.ListIndex <> -1 Then
             pCuenta = CmbCuenta.ItemData(CmbCuenta.ListIndex)
          End If
          Set Rxa = cRx.TraerResPlanCtaDet(DTPDesde.Value, DTPHasta.Value, pCuenta)
          If Rxa.RecordCount <> 0 Then
            ProgressBar1.Max = Rxa.RecordCount
             Do While Not Rxa.EOF
                rs.AddNew
                rs!fecha = Rxa!fechaIva
                If Rxa!Grupo = "EGRESOS" Then
                   rs!Cliente = cRl.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Proveedor='" & Rxa!Cliente & "'")
                Else
                   rs!Cliente = cRl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Cliente='" & Rxa!Cliente & "'")
                End If
                rs!Sucursal = Rxa!Sucursal
                rs!Numero = Rxa!Numero
                rs!Comprobante = Rxa!DesCompr
                rs!Cuenta = Rxa!Cuenta
                rs!Producto = Rxa!Producto
                rs!Cantidad = Rxa!Cantidad
                rs!Detalle = Left(Rxa!DesDet, 50)
                rs!precio = Rxa!PrecioTotal
                rs.Update
                Rxa.MoveNext
                ProgressBar1.Value = i
                i = i + 1
             Loop
          End If
          nImpr = 60
  End If
  FrmImpresor.Show

End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub CrearRs1()
  Set rs = New ADODB.Recordset
  rs.Fields.Append "Grupo", adVarChar, 30, adFldIsNullable
  rs.Fields.Append "NCta", adInteger
  rs.Fields.Append "Cuenta", adVarChar, 30, adFldIsNullable
  rs.Fields.Append "Importe", adCurrency
  rs.Open
End Sub

Private Sub CrearRs2()
  Set rs = New ADODB.Recordset
  rs.Fields.Append "Fecha", adDate
  rs.Fields.Append "Cliente", adVarChar, 50, adFldIsNullable
  rs.Fields.Append "Sucursal", adInteger
  rs.Fields.Append "Numero", adInteger
  rs.Fields.Append "Comprobante", adVarChar, 30, adFldIsNullable
  rs.Fields.Append "Cuenta", adVarChar, 30, adFldIsNullable
  rs.Fields.Append "Producto", adVarChar, 30, adFldIsNullable
  rs.Fields.Append "Cantidad", adSingle
  rs.Fields.Append "Detalle", adVarChar, 50, adFldIsNullable
  rs.Fields.Append "Precio", adCurrency
  rs.Open
End Sub

Private Sub Form_Load()
  DTPDesde.Value = Date
  DTPHasta.Value = Date
  ChkTodos.Value = 1
  CmbCuenta.Enabled = False
  ChkTodas.Enabled = False
  Cargarcombos
End Sub

Private Sub Cargarcombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbCuenta, "PlanDeCuenta", "Id", "Descripcion", ""

End Sub
