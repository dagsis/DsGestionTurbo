VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentasInfProy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proyección de Saldos de Cta.Cte"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   7605
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   60
      TabIndex        =   1
      Top             =   15
      Width           =   7470
      Begin VB.ComboBox CmbZona 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1890
         Width           =   2205
      End
      Begin VB.CheckBox ChkAbc 
         Caption         =   "Alfabetico"
         Height          =   210
         Left            =   180
         TabIndex        =   17
         Top             =   2505
         Width           =   1155
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   210
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "TxtCliente"
         Top             =   630
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   5115
         TabIndex        =   3
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   210
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "TxtCliente"
         Top             =   1290
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   5100
         TabIndex        =   2
         Top             =   1305
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   0
         Left            =   5985
         TabIndex        =   6
         Top             =   525
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   707788801
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   1
         Left            =   5985
         TabIndex        =   8
         Top             =   945
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   707788801
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   2
         Left            =   5985
         TabIndex        =   9
         Top             =   1350
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   707788801
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   3
         Left            =   5985
         TabIndex        =   10
         Top             =   1770
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   707788801
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   4
         Left            =   5985
         TabIndex        =   11
         Top             =   2190
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   707788801
         CurrentDate     =   37021
      End
      Begin VB.Label Label5 
         Caption         =   "Zona :"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1890
         Width           =   615
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
         Left            =   1110
         TabIndex        =   16
         Top             =   630
         Width           =   3930
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
         Left            =   1125
         TabIndex        =   15
         Top             =   1290
         Width           =   3930
      End
      Begin VB.Label Label1 
         Caption         =   "Fechas :"
         Height          =   210
         Left            =   5985
         TabIndex        =   14
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente :"
         Height          =   180
         Left            =   210
         TabIndex        =   7
         Top             =   1065
         Width           =   1260
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   435
      Left            =   6120
      TabIndex        =   13
      Top             =   2910
      Width           =   1380
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   45
      TabIndex        =   0
      Top             =   2910
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "VentasInfProy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rs As Recordset

Private Sub CmdAplicar_Click()
  Dim RsAux As ADODB.Recordset, cRsl As ClsLectura, cRx As ClsComprobantesL, j As Long, rRes As Recordset
  Dim sOrden As String
  
  Set RsAux = New ADODB.Recordset
  Set cRsl = New ClsLectura
  Set cRx = New ClsComprobantesL
  Set Rs = New ADODB.Recordset
  
  CmdAplicar.Enabled = False
  
  Rs.Fields.Append "Cliente", adVarChar, 15
  Rs.Fields.Append "RazonSocial", adVarChar, 50
  Rs.Fields.Append "Fecha1", adDouble
  Rs.Fields.Append "Fecha2", adDouble
  Rs.Fields.Append "Fecha3", adDouble
  Rs.Fields.Append "Fecha4", adDouble
  Rs.Fields.Append "Fecha5", adDouble
  Rs.Fields.Append "Total", adDouble
  
  Rs.Open
  
  If TxtCliente(0).text <> TxtCliente(1).text Then
     sOrden = "Cliente"
     If ChkAbc.Value = 1 Then
        sOrden = "RazonSocial"
     End If
     If CmbZona.text = "TODAS" Then
        Set RsAux = cRsl.TraerRsCondi("Clientes", sOrden, "RazonSocial>='" & LblCliente(0).Caption & "' AND RazonSocial<='" & LblCliente(1).Caption & "'")
     Else
        Set RsAux = cRsl.TraerRsCondi("Clientes", sOrden, "RazonSocial>='" & LblCliente(0).Caption & "' AND RazonSocial<='" & LblCliente(1).Caption & "' and Zona=" & CmbZona.ItemData(CmbZona.ListIndex))
     End If
  Else
     Set RsAux = cRsl.TraerRsCondi("Clientes", "Id", "Cliente='" & TxtCliente(0).text & "'")
  End If
  j = 1
  Dim nSaldo As Double
  ProgressBar1.Value = 1
  If RsAux.RecordCount <> 0 Then
     ProgressBar1.Max = RsAux.RecordCount
     Do While Not RsAux.EOF
        Rs.AddNew
        Rs!Cliente = RsAux!Cliente
        Rs!RazonSocial = RsAux!RazonSocial
        Rs!Fecha1 = cRx.TraerComprobantesProyeccion("01/01/1900", DtpFecha(0).Value, RsAux!Cliente)
        Rs!Fecha2 = cRx.TraerComprobantesProyeccion(DtpFecha(0).Value, DtpFecha(1).Value, RsAux!Cliente)
        Rs!Fecha3 = cRx.TraerComprobantesProyeccion(DtpFecha(1).Value, DtpFecha(2).Value, RsAux!Cliente)
        Rs!Fecha4 = cRx.TraerComprobantesProyeccion(DtpFecha(2).Value, DtpFecha(3).Value, RsAux!Cliente)
        Rs!Fecha5 = cRx.TraerComprobantesProyeccion(DtpFecha(3).Value, DtpFecha(4).Value, RsAux!Cliente)
        Rs!Total = Rs!Fecha1 + Rs!Fecha2 + Rs!Fecha3 + Rs!Fecha4 + Rs!Fecha5
        If Rs!Total <> 0 Then
           Rs.Update
        Else
           Rs.Delete
        End If
        ProgressBar1.Value = j
        j = j + 1
        RsAux.MoveNext
     Loop
     
     nImpr = 66
     FrmImpresor.Show

  End If
  CmdAplicar.Enabled = True

End Sub

Private Sub Form_Load()
  Limpiar
  CargarEtiquetas
End Sub

Private Sub CargarEtiquetas()
  Dim cRsl As ClsClienteL
  
  Dim Rz As ADODB.Recordset
  Dim cZona As ClsLectura

  Set cZona = New ClsLectura
  
  Set cRsl = New ClsClienteL
  Set Rz = New ADODB.Recordset
  Set Rs = cRsl.ClientesAlfabetico(0, False)
  
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Cliente
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Cliente
  LblCliente(1).Caption = Rs!RazonSocial
  ChkAbc.Value = 1
  
  CmbZona.AddItem "TODAS"
  CmbZona.ItemData(CmbZona.NewIndex) = 0
  Set Rz = cZona.TraerRS("ZonaTraerTodos", , True)
  Do While Not Rz.EOF
     CmbZona.AddItem Rz!Descripcion
     CmbZona.ItemData(CmbZona.NewIndex) = Rz!Zona
     Rz.MoveNext
  Loop
  CmbZona.ListIndex = 0
  
End Sub

Private Sub Limpiar()
  Dim ultimo As Integer, dDate1 As Date
  
  DtpFecha(0).Value = Date - 150
  DtpFecha(0).Value = Day(DateSerial(Year(DtpFecha(0).Value), Month(DtpFecha(0).Value) + 1, 0)) & "/" & Month(DtpFecha(0).Value) & "/" & Year(DtpFecha(0).Value)

  DtpFecha(1).Value = DtpFecha(0).Value + 1
  DtpFecha(1).Value = Day(DateSerial(Year(DtpFecha(1).Value), Month(DtpFecha(1).Value) + 1, 0)) & "/" & Month(DtpFecha(1).Value) & "/" & Year(DtpFecha(1).Value)

  DtpFecha(2).Value = DtpFecha(1).Value + 1
  DtpFecha(2).Value = Day(DateSerial(Year(DtpFecha(2).Value), Month(DtpFecha(2).Value) + 1, 0)) & "/" & Month(DtpFecha(2).Value) & "/" & Year(DtpFecha(2).Value)

  DtpFecha(3).Value = DtpFecha(2).Value + 1
  DtpFecha(3).Value = Day(DateSerial(Year(DtpFecha(3).Value), Month(DtpFecha(3).Value) + 1, 0)) & "/" & Month(DtpFecha(3).Value) & "/" & Year(DtpFecha(3).Value)

  DtpFecha(4).Value = DtpFecha(3).Value + 1
  DtpFecha(4).Value = Day(DateSerial(Year(DtpFecha(4).Value), Month(DtpFecha(4).Value) + 1, 0)) & "/" & Month(DtpFecha(4).Value) & "/" & Year(DtpFecha(4).Value)
  
  
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
