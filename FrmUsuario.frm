VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso al Sistema"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5055
      Begin MSComCtl2.DTPicker DTPfecha 
         Height          =   330
         Left            =   3285
         TabIndex        =   8
         Top             =   300
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   582
         _Version        =   393216
         Format          =   203423745
         CurrentDate     =   37158
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   330
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "TxtUsuario"
         Top             =   945
         Width           =   2940
      End
      Begin VB.TextBox TxtContraseña 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "TxtContras"
         Top             =   1440
         Width           =   2910
      End
      Begin VB.CommandButton CmdIngresar 
         Caption         =   "Ingresar"
         Height          =   345
         Left            =   3465
         TabIndex        =   5
         Top             =   1995
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   345
         Left            =   2340
         TabIndex        =   1
         Top             =   1995
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   210
         Left            =   2580
         TabIndex        =   7
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del Usuario :"
         Height          =   480
         Left            =   570
         TabIndex        =   6
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Contraseña :"
         Height          =   180
         Left            =   555
         TabIndex        =   3
         Top             =   1500
         Width           =   930
      End
   End
End
Attribute VB_Name = "FrmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bCanc As Boolean, sUsuario As String, sContra As String

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Sub RemoveCancelMenuItem(frm As Form)
  Dim hSysMenu As Long
  hSysMenu = GetSystemMenu(frm.hwnd, 0)
  Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
  Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

Private Sub CmdCancelar_Click()
  bCanc = True
  Unload Me
End Sub

Private Sub CmdIngresar_Click()
  Dim cRsl As ClsLectura, sUserComp As String
  Dim RsUsuario As Recordset
  
  sUsuario = TxtUsuario.text
  sContra = TxtContraseña.text
   
  Set cRsl = New ClsLectura
  Set RsUsuario = cRsl.TraerRsCondi("Usuarios", "Id_Usuario", "Usuario='" & sUsuario & "' and Contraseña='" & sContra & "'")
  
  If RsUsuario.RecordCount <> 0 Then
     'If cRsl.ValidarUsuario(TxtUsuario, TxtContraseña) = True Then
    
     nSucursal = RsUsuario!Suc 'cRsl.TraerValorDeUnCampo("Usuarios", "Suc", "Usuario='" & sUsuario & "'")
     bAdmin = IIf(RsUsuario!Categoria < 2, True, False)
     bProg = IIf(RsUsuario!Categoria = 0, True, False)
     
   '  Date = DTPfecha.Value
     
     sComputadora = ComputerName
     sUserComp = RsUsuario!Computadora 'TraerValorDeUnCampo("Usuarios", "Computadora", "Usuario='" & sUsuario & "'")
     
     If sComputadora <> sUserComp And sUserComp <> "TODAS" Then
        MsgBox "Usuario No asignado para Usar en Esta Computadora", vbCritical, "Atención"
        Limpiar
        CmdIngresar.Default = False
        TxtUsuario.SetFocus
     Else
        Unload Me
     End If
  Else
     MsgBox "Usuario o Contraseña Incorrectas", vbCritical, "Atención"
     Limpiar
     CmdIngresar.Default = False
     TxtUsuario.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  TxtUsuario.SetFocus
  DTPfecha.Value = Date
End Sub
Private Sub Limpiar()
  TxtUsuario.text = ""
  TxtContraseña.text = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim WshShell As Object
  If KeyAscii = vbKeyReturn Then
      Set WshShell = CreateObject("WScript.Shell")
      WshShell.SendKeys "{TAB}"
      KeyAscii = 0
   End If
End Sub

Private Sub Form_Load()
  Limpiar
 ' RemoveCancelMenuItem Me
End Sub
Private Sub TxtUsuario_LostFocus()
  CmdIngresar.Default = True
End Sub
