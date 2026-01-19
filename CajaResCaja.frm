VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CajaResCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Caja"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   4665
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      Height          =   450
      Left            =   3630
      TabIndex        =   1
      Top             =   1920
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen"
      Height          =   1740
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   4410
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   285
         Left            =   1890
         TabIndex        =   5
         Top             =   930
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   21233665
         CurrentDate     =   38593
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   285
         Left            =   1890
         TabIndex        =   4
         Top             =   375
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   21233665
         CurrentDate     =   38593
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha :"
         Height          =   240
         Left            =   555
         TabIndex        =   3
         Top             =   915
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha :"
         Height          =   210
         Left            =   570
         TabIndex        =   2
         Top             =   375
         Width           =   1230
      End
   End
End
Attribute VB_Name = "CajaResCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Rs As ADODB.Recordset

Private Sub Command1_Click()
  Dim Rxa As ADODB.Recordset, cRx As ClsComprobantesL
  
  CrearRs
  
  Set cRx = New ClsComprobantesL
  Set Rxa = cRx.TraerOtrasCajas()
  Do While Not Rxa.EOF
     If Rxa!formapago <> "CHEQUES PROPIOS" Then
        Rs.AddNew
        Rs!Forma = Rxa!formapago
        Rs!Banco = ""
        Rs!Numero = ""
        Rs!Titular = ""
        Rs!Fecha = Null
        Rs!Importe = Rxa!Total
        Rs.Update
     End If
     Rxa.MoveNext
  Loop
  
  Set Rxa = cRx.TraerCajacheque()
    Do While Not Rxa.EOF
     If Not IsNull(Rxa!formapago) Then
        Rs.AddNew
        Rs!Forma = Rxa!formapago
        Rs!Banco = Rxa!Descripcion
        Rs!Numero = Rxa!Numero
        Rs!Titular = Rxa!Titular
        Rs!Fecha = Rxa!FechaAcre
        Rs!Importe = Rxa!Importe
        Rs.Update
     End If
     Rxa.MoveNext
  Loop
  
  nImpr = 36
  FrmImpresor.Show

End Sub

Private Sub Command2_Click()
  Unload Me
End Sub


Private Sub CrearRs()
  Set Rs = New ADODB.Recordset
  Rs.Fields.Append "Forma", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Banco", adVarChar, 30, adFldIsNullable
  Rs.Fields.Append "Numero", adVarChar, 20, adFldIsNullable
  Rs.Fields.Append "Titular", adVarChar, 60, adFldIsNullable
  Rs.Fields.Append "Fecha", adDate, , adFldIsNullable
  Rs.Fields.Append "Importe", adCurrency
  Rs.Open
  
End Sub

Private Sub Form_Load()
  DtpDesde.Value = Date
  DtpHasta.Value = Date
End Sub
