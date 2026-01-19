VERSION 5.00
Begin VB.Form ComproBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Comprobantes"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   330
      Left            =   4110
      TabIndex        =   6
      Top             =   953
      Width           =   1110
   End
   Begin VB.TextBox TxtNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4065
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   398
      Width           =   1140
   End
   Begin VB.ComboBox CmbSucursal 
      Height          =   315
      Left            =   2655
      TabIndex        =   2
      Text            =   "CmbSucursal"
      Top             =   398
      Width           =   1335
   End
   Begin VB.ComboBox CmbComprobante 
      Height          =   315
      Left            =   315
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   398
      Width           =   2265
   End
   Begin VB.Label Label3 
      Caption         =   "Número :"
      Height          =   195
      Left            =   4065
      TabIndex        =   5
      Top             =   158
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "Sucursal :"
      Height          =   195
      Left            =   2655
      TabIndex        =   4
      Top             =   158
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante :"
      Height          =   240
      Left            =   315
      TabIndex        =   0
      Top             =   158
      Width           =   1110
   End
End
Attribute VB_Name = "ComproBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nId As Long
Dim cRsl As ClsLectura, cBus As ClsComprobantesL

Private Sub CmdAplicar_Click()
  Dim nCompro As Integer, nSucu As Long, nNum As Long
  
  nCompro = CmbComprobante.ItemData(CmbComprobante.ListIndex)
  nSucu = CmbSucursal.text
  nNum = TxtNumero.text
  Set cBus = New ClsComprobantesL
  nId = cBus.TraerIdComprobante(nCompro, nSucu, nNum)
  If nId = 0 Then
     MsgBox "Comprobante no encontrado", vbCritical, "Atención"
     Exit Sub
  End If
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
End Sub
Private Sub Form_Load()
 CargarCombos
End Sub
Private Sub CargarCombos()
  Dim i As Byte
  On Error GoTo errHandler

  Set cRsl = New ClsLectura
  Select Case nLlama
         Case 1
             cRsl.CargaCombo CmbComprobante, "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Ver=1 and Usuario=" & nUsuario
         Case 2
             cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=2 and Suc=" & nSucursal
         Case 3
             cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "(TipoOperacion=8 Or TipoOperacion=12) and Suc=" & nSucursal
         Case 4
             cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=6 and Suc=" & nSucursal
         Case 5
             cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=7 and Suc=" & nSucursal
         Case 6
             cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=4 and Suc=" & nSucursal
         Case 7
             cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "TipoOperacion=9 and Suc=" & nSucursal
         Case 8
             cRsl.CargaCombo CmbComprobante, "Comprobantes", "Id", "Descripcion", "(TipoOperacion=14 Or TipoOperacion=3 or TipoOperacion=2) and Suc=" & nSucursal
         Case 9
             cRsl.CargaCombo CmbComprobante, "UsuariosComprobantesVentas", "Comprobante", "Descripcion", "Ver=1 and Usuario=" & nUsuario
          
  End Select
  
  For i = 1 To 99
      CmbSucursal.AddItem Format(i, "00000")
      CmbSucursal.ItemData(CmbSucursal.NewIndex) = i
  Next i
  
  Set cRsl = Nothing
  TxtNumero.text = Format(0, "000000")
  CmbComprobante.ListIndex = 0
  CmbSucursal.ListIndex = 0
Exit Sub

errHandler:
   ManejaErrores
End Sub
Private Sub TxtNumero_GotFocus()
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtNumero_LostFocus()
  If TxtNumero.text = "" Then TxtNumero.text = 0
  TxtNumero.text = Format(TxtNumero.text, "000000")
End Sub
