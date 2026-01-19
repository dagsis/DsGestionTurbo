VERSION 5.00
Begin VB.Form AbrirMesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abrir Mesa"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmbOk 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   3165
      TabIndex        =   1
      Top             =   3015
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Apertura de Mesa"
      Height          =   2910
      Left            =   135
      TabIndex        =   2
      Top             =   30
      Width           =   4125
      Begin VB.ComboBox CmbLista 
         Enabled         =   0   'False
         Height          =   315
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1455
         Width           =   2370
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   315
         Left            =   1905
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   2145
         Width           =   885
      End
      Begin VB.ComboBox CmbMozo 
         Height          =   315
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   780
         Width           =   2370
      End
      Begin VB.Label Label5 
         Caption         =   "L.Precio :"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad de Cubietos :"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   2100
         Width           =   1725
      End
      Begin VB.Label Label3 
         Caption         =   "Mozo :"
         Height          =   240
         Left            =   210
         TabIndex        =   5
         Top             =   495
         Width           =   600
      End
      Begin VB.Label LblMesa 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1830
         TabIndex        =   4
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Mesa :"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   585
      End
   End
End
Attribute VB_Name = "AbrirMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbOk_Click()
  Dim rAux As Recordset, cRsl As ClsLectura, cRle As ClsEscritura
    
  Set rAux = New Recordset
  Set cRsl = New ClsLectura
  Set cRle = New ClsEscritura
  
  FrmAdicion.CmdMesa(nMesa).BackColor = &HFF&
    
  Set rAux = cRsl.RsVacio("Auxiliar", "Id", "N")
  
  rAux.AddNew
  rAux![a_codart] = "01-0099"
  rAux![A_DESCRI] = "SERVICIO DE MESA"
  rAux![a_cant] = TxtCantidad.text
  rAux![a_Mesa] = LblMesa.Caption
  rAux![a_mozo] = CmbMozo.ItemData(CmbMozo.ListIndex)
  rAux![a_impcom] = "N"
  rAux![a_hora] = Time
  rAux![a_fecha] = Date
  rAux![A_MENSAJE] = "ABIERTA"
  rAux![a_lista] = cRsl.RegPorDefecto("ListaRes")
  rAux![a_precio] = cRsl.TraerValorDeUnCampo("Precios", "Precio", "Producto='" & rAux![a_codart] & "' AND Lista=" & CmbLista.ItemData(CmbLista.ListIndex))
  rAux![a_puerto] = "LPT1"
  rAux![a_puerto2] = "LPT1"
  rAux![a_puerto3] = "LPT1"
  rAux![a_total] = rAux![a_cant] * rAux![a_precio]
  rAux![a_Observ] = ""
  rAux.Update
  
  cRle.Actualizar "Auxiliar_A", rAux
  
  nLista = CmbLista.ItemData(CmbLista.ListIndex)
  nMozo = CmbMozo.ItemData(CmbMozo.ListIndex)
  nMesa = LblMesa.Caption
  
  Unload Me
  frmMesa.Show 1
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
        Unload Me
        FrmAdicion.TxtMesa.SetFocus
        Exit Sub
     End If
  If KeyAscii = vbKeyReturn Then
     KeyAscii = 0
     SendKeys "{TAB}"
     Exit Sub
  End If
  KeyAscii = Upper(KeyAscii)
End Sub

Private Sub Form_Load()
  Limpiar
  CargarCombos
End Sub

Private Sub Limpiar()
  LblMesa.Caption = Format$(nMesa, "00000")
  TxtCantidad.text = "1"
End Sub
Private Sub CargarCombos()
  Dim cRsl As ClsLectura, nAMozo As Integer
  
  Set cRsl = New ClsLectura
  
  nAMozo = cRsl.TraerValorDeUnCampo("MozoMesa", "Mozo", "Mesa=" & nMesa)
  If nAMozo = 0 Then
     nAMozo = cRsl.RegPorDefecto("Vendedor")
  End If
  
  cRsl.CargaCombo CmbMozo, "Vendedores", "Vendedor", "Descripcion", "Vendedor=" & nAMozo
  
  cRsl.CargaCombo CmbLista, "ListadePrecio", "ListaPrecio", "Descripcion", ""
  CmbMozo.ListIndex = 0
  CmbLista.ListIndex = 0
  
End Sub

Private Sub TxtCantidad_GotFocus()
  
  TxtCantidad.SelStart = 0
  TxtCantidad.SelLength = Len(TxtCantidad.text)
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCantidad_LostFocus()
  If TxtCantidad.text = "" Then TxtCantidad.text = "1"

End Sub
