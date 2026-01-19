VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaRepartoRemito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reparto por Remitos"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   4275
   Begin VB.CommandButton CmdEstado 
      Caption         =   "Cambiar a Pendientes"
      Height          =   420
      Left            =   195
      TabIndex        =   10
      Top             =   2790
      Width           =   1890
   End
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
      Height          =   2670
      Left            =   120
      TabIndex        =   5
      Top             =   30
      Width           =   4005
      Begin VB.ComboBox CmbEstado 
         Height          =   315
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   2205
      End
      Begin VB.ComboBox CmbZona 
         Height          =   315
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   795
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   345
         Left            =   1545
         TabIndex        =   2
         Top             =   1425
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   54198273
         CurrentDate     =   36983
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   345
         Left            =   1545
         TabIndex        =   3
         Top             =   2070
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   54198273
         CurrentDate     =   36983
      End
      Begin VB.Label Label3 
         Caption         =   "Estado :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Zona :"
         Height          =   255
         Left            =   255
         TabIndex        =   8
         Top             =   795
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   255
         TabIndex        =   7
         Top             =   1425
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   240
         Left            =   255
         TabIndex        =   6
         Top             =   2070
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Aplicar"
      Height          =   420
      Left            =   3165
      TabIndex        =   4
      Top             =   2790
      Width           =   945
   End
End
Attribute VB_Name = "VentaRepartoRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Private Sub CmbEstado_Change()
  CrearRs
End Sub

Private Sub CmdEstado_Click()
  On Error GoTo Errores

  If rs.RecordCount <> 0 Then
     If CmbEstado.text = "En Espera" Then
        If MsgBox("Esta Seguro de Realizar Esta Operación ?", 20, "Atención") = 6 Then

           rs.MoveFirst
        
           db.Open sDb
           db.BeginTrans
        
           Do While Not rs.EOF
              db.Execute "UPDATE RegTransporte set Estado=2 WHERE (RegTransporte.Estado = 1) AND RegTransporte.Fecha>='" & DtpDesde.Value & "' AND RegTransporte.Fecha<='" & DtpHasta.Value & "' "
              rs.MoveNext
           Loop
           db.CommitTrans
           db.Close
           CrearRs
           MsgBox "Cambio de Estado a Pendiente Realizado Con Exito", vbInformation, "Atención"
        End If
     Else
        MsgBox "Imprima Los Registros En Espera antes de Cambiar el Estado a Pendiente...", vbInformation, "Atención"
     End If
  End If
  
Exit Sub
Errores:
  MsgBox Err.Description, vbCritical, "Atención"
  db.RollbackTrans
  db.Close
End Sub

Private Sub CmdImprimir_Click()
  Dim RsAux As Recordset, cRsl As ClsLectura, cRe As ClsComprobantesL
  
  Set cRsl = New ClsLectura
  Set cRe = New ClsComprobantesL
  
  CrearRs
  Set RsAux = cRe.TraerRepartoRemi(DtpDesde.Value, DtpHasta.Value, CmbEstado.ItemData(CmbEstado.ListIndex), CmbZona.text)
  
  If RsAux.RecordCount <> 0 Then
     Do While Not RsAux.EOF
        rs.AddNew
        rs!fecha = RsAux!fecha
        rs!NRemito = Format(RsAux!Sucursal, "0000") & "-" & Format(RsAux!Numero, "00000000")
        rs!nFactura = RsAux!Motivo
        rs!Horario = RsAux!Hora
        rs!Destinatario = RsAux!Destinatario
        rs!Vendedor = cRsl.TraerValorDeUnCampo("Vendedores", "Descripcion", "Vendedor=" & RsAux!Vendedor)
        rs!Direccion = RsAux!Direccion
        rs!Telefono = RsAux!Telefono
        rs!Observaciones = RsAux!Obsrevaciones
        rs!Cantidad = RsAux!Cantidad
        rs!Producto = RsAux!Producto
        rs!Descripcion = RsAux!Descripcion
        rs.Update
        RsAux.MoveNext
     Loop
  End If
  
  nImpr = 69
  FrmImpresor.Show
  
End Sub

Private Sub Form_Load()
  CargarCombos
  CrearRs
End Sub

Private Sub CargarCombos()
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbZona, "Zonas", "Zona", "Descripcion", ""

  CmbZona.AddItem "TODAS"
  CmbZona.ItemData(CmbZona.NewIndex) = 0
  CmbZona.text = "TODAS"
  
  CmbEstado.AddItem "En Espera"
  CmbEstado.ItemData(CmbEstado.NewIndex) = 1
  
  CmbEstado.AddItem "Pendiente"
  CmbEstado.ItemData(CmbEstado.NewIndex) = 2

  CmbEstado.AddItem "Entregado"
  CmbEstado.ItemData(CmbEstado.NewIndex) = 3

  CmbEstado.text = "En Espera"
  
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  
End Sub

Private Sub CrearRs()
 Set rs = New ADODB.Recordset
  
  rs.Fields.Append "Fecha", adDate
  rs.Fields.Append "NRemito", adVarChar, 20
  rs.Fields.Append "NFactura", adVarChar, 100
  rs.Fields.Append "Horario", adVarChar, 50
  rs.Fields.Append "Destinatario", adVarChar, 80
  rs.Fields.Append "Vendedor", adVarChar, 50, adFldIsNullable
  rs.Fields.Append "Direccion", adVarChar, 80, adFldIsNullable
  rs.Fields.Append "Telefono", adVarChar, 80, adFldIsNullable
  rs.Fields.Append "Observaciones", adVarChar, 200, adFldIsNullable
  rs.Fields.Append "Cantidad", adSingle
  rs.Fields.Append "Producto", adVarChar, 25, adFldIsNullable
  rs.Fields.Append "Descripcion", adVarChar, 80, adFldIsNullable
  rs.Open
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If CmbEstado.text = "En Espera" And rs.RecordCount <> 0 Then
    If MsgBox("Tiene Registros En Espera para Cambiar de Estado. Sale sin Procesar ?", 20, "Atención") = 6 Then
       Exit Sub
    End If
    Cancel = 1
  End If
End Sub
