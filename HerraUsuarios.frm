VERSION 5.00
Begin VB.Form HerraUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios y Permisos"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   6975
   Begin VB.Frame Frame1 
      Caption         =   "ABM de Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   105
      TabIndex        =   12
      Top             =   0
      Width           =   3285
      Begin VB.ComboBox CmbComputadora 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1350
         Width           =   2940
      End
      Begin VB.ComboBox CmbControlador 
         Height          =   315
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3195
         Width           =   1605
      End
      Begin VB.ComboBox CmbCategoria 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   2325
      End
      Begin VB.TextBox TxtCom 
         Height          =   330
         Left            =   165
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "1"
         Top             =   3195
         Width           =   330
      End
      Begin VB.TextBox TxtPass 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1515
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "TxtPass"
         Top             =   2625
         Width           =   1095
      End
      Begin VB.TextBox TxtUser 
         Height          =   300
         Left            =   150
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "TxtUser"
         Top             =   2625
         Width           =   1095
      End
      Begin VB.TextBox TxtNombre 
         Height          =   285
         Left            =   150
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "TxtDescripcion"
         Top             =   2010
         Width           =   2910
      End
      Begin VB.Label Label9 
         Caption         =   "Computadora :"
         Height          =   240
         Left            =   165
         TabIndex        =   21
         Top             =   1125
         Width           =   1170
      End
      Begin VB.Label Label8 
         Caption         =   "Controlador Fiscal :"
         Height          =   210
         Left            =   1530
         TabIndex        =   20
         Top             =   2955
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "Categoria :"
         Height          =   240
         Left            =   135
         TabIndex        =   19
         Top             =   465
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Controladora :"
         Height          =   255
         Left            =   150
         TabIndex        =   18
         Top             =   2955
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "Contraseña :"
         Height          =   195
         Left            =   1515
         TabIndex        =   17
         Top             =   2370
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario :"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   2385
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1755
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2445
         TabIndex        =   13
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   3495
      TabIndex        =   11
      Top             =   90
      Width           =   3375
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   810
      Index           =   0
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   4
      Left            =   5865
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3825
      Width           =   1005
   End
End
Attribute VB_Name = "HerraUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCod As Integer, sw As Boolean
Dim Rs As ADODB.Recordset
Dim cRsl As ClsLectura
Dim cRse As ClsEscritura

Private Sub Botones(bNue As Boolean, bBorr As Boolean, bModi As Boolean, bImp As Boolean, bSal As Boolean)
  CmdBotones(0).Enabled = bNue
  CmdBotones(1).Enabled = bBorr
  CmdBotones(2).Enabled = bModi
  CmdBotones(4).Enabled = bSal
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim WshShell As Object
   If KeyAscii = vbKeyReturn Then
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{TAB}"
   End If
'     KeyAscii = Upper(KeyAscii)
End Sub
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
  Next ctl
  TxtCom.text = 0
End Sub

Private Sub Habilitar(bEstado As Boolean, bList As Boolean)
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Or TypeOf ctl Is CheckBox _
         Or TypeOf ctl Is DTPicker Or TypeOf ctl Is ListBox _
         Or TypeOf ctl Is ComboBox Then
         ctl.Enabled = bEstado
      End If
  Next ctl
  List1.Enabled = bList
End Sub

Private Sub LinkearTexto()
  On Error GoTo error
  
  Dim sCond As String

  Label3.Caption = Format(Rs!Id_Usuario, "00")
  
  Select Case Rs!Categoria
         Case 0
              sCond = "Programador"
         Case 1
              sCond = "Administrador"
         Case 2
              sCond = "Operador"
  End Select
  
  CmbCategoria.text = sCond
  TxtNombre.text = Rs!Nombre
  TxtUser.text = Rs!Usuario
  TxtPass.text = Rs!Contraseña
  TxtCom.text = Rs!Controladora
  
  Select Case Rs!Controlador
      Case 0
           sCond = "Sin Controlador"
      Case 1
           sCond = "Hasar"
      Case 2
           sCond = "Epson"
      Case 3
           sCond = "Epson N.G."
  End Select
  
  CmbControlador.text = sCond
  CmbComputadora.text = Rs!Computadora
Exit Sub
error:
 If err.Number = 383 Then
    CmbComputadora.AddItem Rs!Computadora
    Resume
 End If
End Sub

Private Sub Form_Load()
 On Error GoTo errHandler

 Set Rs = New ADODB.Recordset
 Set cRsl = New ClsLectura

 sw = True
 Limpiar
 Habilitar False, False
 Botones pEdit, pBorr, pEdit, False, True
 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
 CmdBotones(4).Picture = LoadResPicture("Salir", 0)
 
 CargarCombos
 CargarLista
  If Rs.EOF = False Or Rs.BOF = False Then
    Botones pEdit, pBorr, pEdit, True, True
    Habilitar False, True
    Rs.MoveFirst
    LinkearTexto
 End If
 Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub CargarCombos()
  Dim i As Integer
  Dim sw As CNetworkEnum 'Decalre class

  CmbCategoria.AddItem "Programador"
  CmbCategoria.ItemData(CmbCategoria.NewIndex) = 0
  CmbCategoria.AddItem "Administrador"
  CmbCategoria.ItemData(CmbCategoria.NewIndex) = 1
  CmbCategoria.AddItem "Operador"
  CmbCategoria.ItemData(CmbCategoria.NewIndex) = 2
  
  CmbControlador.AddItem "Sin Controlador"
  CmbControlador.ItemData(CmbControlador.NewIndex) = 0
  CmbControlador.AddItem "Hasar"
  CmbControlador.ItemData(CmbControlador.NewIndex) = 1
  CmbControlador.AddItem "Epson"
  CmbControlador.ItemData(CmbControlador.NewIndex) = 2
  CmbControlador.AddItem "Epson N.G."
  CmbControlador.ItemData(CmbControlador.NewIndex) = 3
  
  
  
'  Set sw = New CNetworkEnum 'Set class to SW
'  Call sw.SetResourceType(0) 'Set resource type
'  Call sw.Reset
'  Dim ST As String
'
'  ST = Replace(sw.GetServerList, "\\", "") 'Replace the \\ in the computer names
'
'  For i = 1 To Len(ST)
'      If Mid(ST, i, 1) = "," Then
'          CmbComputadora.AddItem (Left(ST, i - 1))
'          ST = Right(ST, Len(ST) - i)
'          i = 1
'      End If
'  Next i
  
  CmbComputadora.AddItem "TODAS"
  
End Sub

Private Sub List1_Click()
 On Error GoTo errHandler
 Dim sCond As String, sCod As Integer

 Set cRsl = New ClsLectura
 sCod = List1.ItemData(List1.ListIndex)
 
 sCond = "Id_Usuario=" & sCod
 Set Rs = cRsl.TraerRsCondi("Usuarios", "Id_Usuario", sCond)
 
 LinkearTexto
 Set cRsl = Nothing
Exit Sub
errHandler:
   ManejaErrores
End Sub

Private Sub TxtCom_Change()
  TodoBien
End Sub

Private Sub TxtCom_GotFocus()
  TxtCom.SelStart = 0
  TxtCom.SelLength = Len(TxtCom.text)
End Sub

Private Sub TxtCom_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtCom_LostFocus()
  If TxtCom.text = "" Then TxtCom.text = 0
End Sub

Private Sub TxtNombre_Change()
  TodoBien
End Sub
Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              If Val(Label3.Caption) <> nUsuario Then
                 Borrar
              End If
         Case 2
              Modificar
         Case 3
         Case 4
              Salir
  End Select
End Sub

Private Sub Nuevo()
  Dim cRsl As New ClsLectura
  
  Set cRsl = New ClsLectura
  
  On Error GoTo Errores

  If sw = True Then
     sw = False
     Botones False, False, False, False, True
     Habilitar True, False
     Limpiar
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(4).Caption = "Cancelar"
     CmdBotones(4).Picture = LoadResPicture("Cancelar", 0)
     Label3.Caption = Format(cRsl.TraerValorDeUnCampo("Sistema", "Usuario", ""), "00")
     TxtNombre.SetFocus
  Else
     Botones True, True, True, True, True
     Habilitar False, True
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(4).Caption = "Salir"
     CmdBotones(4).Picture = LoadResPicture("Salir", 0)
     sw = True
     Grabar
     CargarLista
     List1.ListIndex = List1.ListCount - 1
  End If
Exit Sub
Errores:
   ManejaErrores
End Sub

Private Sub Salir()
 On Error GoTo errHandler

  Botones True, False, False, False, True
  If sw = False Then
     Limpiar
     If CmdBotones(0).Caption = "Grabar" Then
        sw = True
        Habilitar False, True
        CmdBotones(0).Caption = "Nuevo"
        CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
        CmdBotones(4).Caption = "Salir"
        CmdBotones(4).Picture = LoadResPicture("Salir", 0)
     Else
        sw = True
        Habilitar False, True
        CmdBotones(2).Caption = "Modificar"
        CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
        CmdBotones(4).Caption = "Salir"
        CmdBotones(4).Picture = LoadResPicture("Salir", 0)
     End If
      If Rs.EOF = False Or Rs.BOF = False Then
        Botones True, True, True, True, True
        LinkearTexto
     End If
  Else
     Unload Me
  End If
  Exit Sub
errHandler:
   ManejaErrores
End Sub
Private Sub Modificar()
  On Error GoTo Errores

  Static rList As Integer

  If sw = True Then
     sw = False
     rList = List1.ListIndex
     Botones False, False, True, False, True
     Habilitar True, False
     CmdBotones(2).Caption = "Grabar"
     CmdBotones(2).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(4).Caption = "Cancelar"
     CmdBotones(4).Picture = LoadResPicture("Cancelar", 0)
  Else
     Botones True, True, True, True, True
     Habilitar False, True
     Grabar
     CmdBotones(2).Caption = "Modificar"
     CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
     CmdBotones(4).Caption = "Salir"
     CmdBotones(4).Picture = LoadResPicture("Salir", 0)
     sw = True
     CargarLista
     List1.ListIndex = rList
  End If
Exit Sub
Errores:
   ManejaErrores

End Sub
Private Sub Borrar()
  On Error GoTo errHandler

  Dim nList As Integer

  Set cRse = New ClsEscritura

  nList = List1.ListIndex
  If MsgBox("Borra el Usuario Seleccionado ?", 20, "Atención") = 6 Then
     sCod = Label3.Caption
     cRse.Actualizar "Usuarios_B", Nothing, Val(Label3.Caption), "Usuario"
     CargarLista
     Limpiar
     Botones True, False, False, False, True
      If Rs.EOF = False Or Rs.BOF = False Then
        Botones True, True, True, True, True
        If nList = 0 Then nList = 1
        List1.ListIndex = nList - 1
        LinkearTexto
     End If
  End If
  Exit Sub
errHandler:
  ManejaErrores
End Sub
Private Sub CargarLista()
  On Error GoTo errHandler
  Dim sCond As String

  Set cRsl = New ClsLectura
  Set Rs = New ADODB.Recordset

  List1.Clear
  sCond = "Id_Usuario<>1"
  If nUsuario = 1 Then
     sCond = "Id_Usuario<>0"
  End If
  Set Rs = cRsl.TraerRsCondi("Usuarios", "Id_Usuario", sCond)

  Do While Not Rs.EOF
     List1.AddItem Rs!Nombre
     List1.ItemData(List1.NewIndex) = Rs!Id_Usuario
     Rs.MoveNext
  Loop
  Set cRsl = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub Grabar()
  On Error GoTo errHandler

  Dim Rs1 As ADODB.Recordset, cRsl As ClsLectura

  Set cRsl = New ClsLectura
  
  Set Rs1 = cRsl.RsVacio("Usuarios", "Id_Usuario", "N")
  
  Rs1.AddNew
  Rs1!Id_Usuario = Val(Label3.Caption)
  Rs1!Suc = 1
  Rs1!Nombre = TxtNombre.text
  Rs1!Usuario = TxtUser.text
  Rs1!Contraseña = TxtPass.text
  Rs1!Categoria = CmbCategoria.ItemData(CmbCategoria.ListIndex)
  Rs1!Controlador = CmbControlador.ItemData(CmbControlador.ListIndex)
  Rs1!Controladora = TxtCom.text
  Rs1!Computadora = CmbComputadora.text
  Rs1!Scanner = 0
  Rs1!Comprobante = 6
  Rs1!Caja = 1
  Rs1.Update

  Set cRse = New ClsEscritura

  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Actualizar "Usuarios_M", Rs1
  Else
     cRse.Actualizar "Usuarios_A", Rs1
  End If
  Rs1.Close
  Set cRse = Nothing
  bFamilia = True
  Exit Sub
errHandler:
  ManejaErrores
End Sub


Private Sub TodoBien()
 If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtNombre.text) <> 0 And Len(TxtUser.text) <> 0 And Len(TxtPass.text) <> 0 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(2).Enabled = False
        If Len(TxtNombre.text) <> 0 And Len(TxtUser.text) <> 0 And Len(TxtPass.text) <> 0 Then
           CmdBotones(2).Enabled = True
        End If
     End If
  End If
End Sub

Private Sub TxtPass_Change()
  TodoBien
End Sub

Private Sub TxtUser_Change()
  TodoBien
End Sub
