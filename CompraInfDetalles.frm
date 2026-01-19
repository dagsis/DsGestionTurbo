VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CompraInfDetalles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras por Detalles"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   6285
   Begin VB.Frame Frame1 
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   5970
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripción"
         Height          =   210
         Left            =   4410
         TabIndex        =   15
         Top             =   240
         Width           =   1245
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   3135
         TabIndex        =   14
         Top             =   240
         Width           =   1050
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   0
         Left            =   255
         MaxLength       =   15
         TabIndex        =   6
         Text            =   "TxtCliente"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   0
         Left            =   4980
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtCliente 
         Height          =   285
         Index           =   1
         Left            =   255
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "TxtCliente"
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   240
         Index           =   1
         Left            =   4980
         TabIndex        =   3
         Top             =   1260
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   0
         Left            =   975
         TabIndex        =   2
         Top             =   1815
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   208338945
         CurrentDate     =   37021
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   360
         Index           =   1
         Left            =   3195
         TabIndex        =   7
         Top             =   1815
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   208338945
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
         Left            =   1155
         TabIndex        =   13
         Top             =   600
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
         Left            =   1155
         TabIndex        =   12
         Top             =   1260
         Width           =   3705
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   210
         Left            =   345
         TabIndex        =   11
         Top             =   1815
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   165
         Left            =   2640
         TabIndex        =   10
         Top             =   1815
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente :"
         Height          =   195
         Left            =   255
         TabIndex        =   9
         Top             =   330
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente :"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   1005
         Width           =   1260
      End
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5130
      TabIndex        =   0
      Top             =   2640
      Width           =   930
   End
End
Attribute VB_Name = "CompraInfDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRsl As ClsClienteL, Rs As ADODB.Recordset

Private Sub CmdAplicar_Click()
  nImpr = 33
  FrmImpresor.Show
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
  Dim crs As ClsLectura
  Set crs = New ClsLectura
  nBuscar = 2
  FrmBuscar.Show 1
  If nDat <> 0 Then
     Set Rs = crs.TraerTodos("Proveedores", "Id")
     Rs.MoveFirst
     Rs.Find "Id=" & nDat
     Set Rs = crs.TraerRS("ProveedorTraerUno", nDat, True)
     LinkearTexto Index
  End If
  Set cRsl = Nothing

End Sub
Private Sub LinkearTexto(Index As Integer)
  TxtCliente(Index).text = Rs!Proveedor
  LblCliente(Index).Caption = Rs!RazonSocial
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
  OptDescripcion.Value = True
  DtpFecha(0).Value = Date
  DtpFecha(1).Value = Date
  LblCliente(0).Caption = ""
  LblCliente(1).Caption = ""
End Sub

Private Sub Form_Load()
  Set cRsl = New ClsClienteL
  Set Rs = New ADODB.Recordset

  Limpiar
  CargarEtiquetas
  
End Sub

Private Sub CargarEtiquetas()
  Set Rs = cRsl.ProveedorAlfabetico
  Rs.MoveFirst
  TxtCliente(0).text = Rs!Proveedor
  LblCliente(0).Caption = Rs!RazonSocial
  Rs.MoveLast
  TxtCliente(1).text = Rs!Proveedor
  LblCliente(1).Caption = Rs!RazonSocial
End Sub
Private Sub TxtCliente_GotFocus(Index As Integer)
  TxtCliente(Index).SelStart = 0
  TxtCliente(Index).SelLength = Len(TxtCliente(Index).text)
End Sub
Private Sub TxtCliente_LostFocus(Index As Integer)
  CmdAplicar.Enabled = False
  LblCliente(Index).Caption = ""
  If TxtCliente(Index).text <> "" Then
     If cRsl.BuscarNombreProveedor(TxtCliente(Index).text) <> "" Then
        LblCliente(Index).Caption = cRsl.BuscarNombreProveedor(TxtCliente(Index).text)
        CmdAplicar.Enabled = True
     End If
  End If
End Sub

