VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VentaInfResApli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Cuenta Con Aplicaciones"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   6270
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
      Height          =   2700
      Left            =   180
      TabIndex        =   6
      Top             =   90
      Width           =   5970
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   180
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "TxtCliente"
         Top             =   555
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4920
         TabIndex        =   8
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   180
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "TxtCliente"
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4905
         TabIndex        =   7
         Top             =   1215
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   0
         Left            =   825
         TabIndex        =   2
         Top             =   1755
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   707854337
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   1
         Left            =   3045
         TabIndex        =   3
         Top             =   1755
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   707854337
         CurrentDate     =   37021
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
         Left            =   1080
         TabIndex        =   14
         Top             =   555
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
         Index           =   1
         Left            =   1095
         TabIndex        =   13
         Top             =   1215
         Width           =   3705
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   210
         Left            =   195
         TabIndex        =   12
         Top             =   1755
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   165
         Left            =   2490
         TabIndex        =   11
         Top             =   1755
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   285
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente :"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   1005
         Width           =   1260
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5220
      TabIndex        =   5
      Top             =   2895
      Width           =   930
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   165
      TabIndex        =   4
      Top             =   2910
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "VentaInfResApli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsClienteL
Public Rs As ADODB.Recordset

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
  Set cRsl = Nothing

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
  DtpFecha(0).Value = Date
  DtpFecha(1).Value = Date
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
End Sub

Private Sub Form_Load()
  Dim Rz As ADODB.Recordset
  Dim cZona As ClsLectura
  Dim cRsl As New ClsClienteL
  
  Set Rz = New ADODB.Recordset
  Set Rs = cRsl.ClientesAlfabetico(0, False)
  Set cZona = New ClsLectura
  Limpiar
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Cliente
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Cliente
  LblCliente(1).Caption = Rs!RazonSocial
  
End Sub

Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub
Private Sub TxtCliente_LostFocus(Index As Integer)
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


