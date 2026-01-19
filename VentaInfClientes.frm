VERSION 5.00
Begin VB.Form VentaInfClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Clientes"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   7335
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      Top             =   2985
      Width           =   1140
   End
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
      Height          =   2715
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   7050
      Begin VB.Frame Frame2 
         Caption         =   "Zonas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   3450
         TabIndex        =   8
         Top             =   255
         Width           =   3465
         Begin VB.ComboBox CmbZonaDes 
            Height          =   315
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   555
            Width           =   3195
         End
         Begin VB.ComboBox CmbZonaHas 
            Height          =   315
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1215
            Width           =   3195
         End
         Begin VB.Label Label3 
            Caption         =   "Desde :"
            Height          =   210
            Left            =   165
            TabIndex        =   12
            Top             =   300
            Width           =   1125
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta :"
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   960
            Width           =   1110
         End
      End
      Begin VB.CheckBox ChkAlfabeticamente 
         Caption         =   "Alfabetico"
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   2280
         Width           =   1740
      End
      Begin VB.CheckBox ChkAgrupado 
         Caption         =   "Agrupado por Zona"
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   1770
         Width           =   1875
      End
      Begin VB.ComboBox CmbHasta 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1230
         Width           =   3090
      End
      Begin VB.ComboBox CmbDesde 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   3090
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Cliente :"
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   975
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Cliente :"
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1185
      End
   End
End
Attribute VB_Name = "VentaInfClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Rs As ADODB.Recordset
Dim cRsl As ClsLectura
Private Sub CmdAplicar_Click()
  Set cCliente = New ClsClienteL
  Dim report1 As RptClienteSinZona
    
  Set Rs = New ADODB.Recordset
  Set report1 = New RptClienteSinZona
  Set Rs = cCliente.ClientesListados(IIf(ChkAlfabeticamente.Value, True, False), CmbDesde.text, CmbHasta.text, CmbZonaDes.text, CmbZonaHas.text)
  nImpr = 11
  If ChkAgrupado.Value = 0 Then
     Set Rs = cCliente.ClientesListados(IIf(ChkAlfabeticamente.Value, True, False), CmbDesde.text, CmbHasta.text)
     nImpr = 10
 End If
  FrmImpresor.Show
End Sub

Private Sub ChkAgrupado_Click()
  Frame2.Enabled = False
  CmbZonaDes.Enabled = False
  CmbZonaHas.Enabled = False
  If ChkAgrupado.Value = 1 Then
     Frame2.Enabled = True
     CmbZonaDes.Enabled = True
     CmbZonaHas.Enabled = True
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
Private Sub Form_Load()

  On Error GoTo errHandler

  Set Rs = New ADODB.Recordset
  Set cCliente = New ClsClienteL
    
  Set Rs = cCliente.ClientesAlfabetico(0)
  
  Do While Not Rs.EOF
     CmbDesde.AddItem Rs!RazonSocial
     CmbDesde.ItemData(CmbDesde.NewIndex) = Rs!ID
     CmbHasta.AddItem Rs!RazonSocial
     CmbHasta.ItemData(CmbHasta.NewIndex) = Rs!ID
     Rs.MoveNext
  Loop
  CmbDesde.ListIndex = 0
  CmbHasta.ListIndex = CmbDesde.ListCount - 1
  CargarCombos
  Frame2.Enabled = False
  CmbZonaDes.Enabled = False
  CmbZonaHas.Enabled = False
  
Exit Sub

errHandler:
   ManejaErrores
  
End Sub

Private Sub CargarCombos()
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbZonaDes, "Zonas", "Zona", "Descripcion", ""
  Set cRsl = Nothing
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbZonaHas, "Zonas", "Zona", "Descripcion", ""
  CmbZonaDes.ListIndex = 0
  CmbZonaHas.ListIndex = CmbZonaHas.ListCount - 1
  Set cRsl = Nothing
End Sub
