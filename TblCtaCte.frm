VERSION 5.00
Begin VB.Form TblCtaCte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Corrientes Bancarias"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   6450
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   4
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3390
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   810
      Index           =   3
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3390
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   2235
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3390
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3390
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   810
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3390
      Width           =   1005
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   3480
      TabIndex        =   13
      Top             =   120
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Caption         =   "ABM de Cte Cte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   120
      TabIndex        =   8
      Top             =   30
      Width           =   3285
      Begin VB.TextBox TxtTitular 
         Height          =   285
         Left            =   195
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "TxtTitular"
         Top             =   1275
         Width           =   2910
      End
      Begin VB.TextBox TxtLimite 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   195
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "TxtLimite"
         Top             =   2895
         Width           =   1230
      End
      Begin VB.TextBox TxtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1710
         MaxLength       =   20
         TabIndex        =   4
         Text            =   "TxtNumero"
         Top             =   2355
         Width           =   1185
      End
      Begin VB.ComboBox CmbBanco 
         Height          =   315
         Left            =   195
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1800
         Width           =   2835
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   195
         MaxLength       =   30
         TabIndex        =   0
         Text            =   "TxtDescripcion"
         Top             =   705
         Width           =   2910
      End
      Begin VB.TextBox TxtCuenta 
         Height          =   285
         Left            =   195
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "TxtCuenta"
         Top             =   2355
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "Titular :"
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label Label7 
         Caption         =   "Límite :"
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   2655
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "U.Número :"
         Height          =   195
         Left            =   1725
         TabIndex        =   18
         Top             =   2115
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Banco :"
         Height          =   225
         Left            =   195
         TabIndex        =   17
         Top             =   1575
         Width           =   660
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
         TabIndex        =   12
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   450
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   195
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   2115
         Width           =   735
      End
   End
End
Attribute VB_Name = "TblCtaCte"
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
  CmdBotones(3).Enabled = bImp
  CmdBotones(4).Enabled = bSal
End Sub

Private Sub CmbBanco_Change()
  TodoBien
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
  Label3.Caption = ""
  TxtNumero.text = "1"
  TxtLimite.text = "0.00"
  CmbBanco.ListIndex = -1
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
  Dim cRsl As ClsLectura
  
  Set cRsl = New ClsLectura
  Label3.Caption = Format(Rs!ID, "00")
  TxtDescripcion.text = Rs!Descripcion
  TxtTitular.text = Rs!Titular
  CmbBanco.text = cRsl.TraerValorDeUnCampo("Proveedores", "RazonSocial", "Id=" & Rs!Banco)
  TxtCuenta.text = Rs!Cuenta
  TxtLimite.text = Format(Rs!Limite, nCantDecimales)
  TxtNumero.text = Rs!Numero
  Set cRsl = Nothing
End Sub

Private Sub Form_Load()
 On Error GoTo errHandler

 Set Rs = New ADODB.Recordset
 Set cRsl = New ClsLectura
 
 sw = True
 Limpiar
 Habilitar False, False
 Botones True, False, False, False, True
 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
 CmdBotones(3).Picture = LoadResPicture("Imprimir", 0)
 CmdBotones(4).Picture = LoadResPicture("Salir", 0)
 CargarLista
 If Rs.EOF = False Or Rs.BOF = False Then
    Botones True, True, True, True, True
    Habilitar False, True
    Rs.MoveFirst
    LinkearTexto
 End If
 Set cRsl = Nothing
Exit Sub

errHandler:
   ManejaErrores

End Sub

Private Sub List1_Click()
 On Error GoTo errHandler
 Dim sCond As String, sCod As Integer
 
 Set cRsl = New ClsLectura
 cRsl.CargaCombo CmbBanco, "Proveedores", "Id", "RazonSocial", "CtaCte=1"
 sCod = List1.ItemData(List1.ListIndex)
 Set Rs = cRsl.TraerRS("CuentasBancariasTraerUno", sCod, True)
 LinkearTexto
 Set cRsl = Nothing
Exit Sub
errHandler:
   ManejaErrores
End Sub

Private Sub TxtCuenta_Change()
  TodoBien
End Sub

Private Sub TxtDescripcion_Change()
  TodoBien
End Sub

Private Sub TodoBien()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtDescripcion.text) <> 0 And Len(TxtTitular.text) <> 0 And Len(CmbBanco.text) <> 0 And Len(TxtCuenta.text) <> 0 And Len(TxtNumero.text) <> 0 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(2).Enabled = False
        If Len(TxtDescripcion.text) <> 0 And Len(TxtTitular.text) <> 0 And Len(CmbBanco.text) <> 0 And Len(TxtCuenta.text) <> 0 And Len(TxtNumero.text) <> 0 Then
           CmdBotones(2).Enabled = True
        End If
     End If
  End If

End Sub
Private Sub CmdBotones_Click(Index As Integer)
  Select Case Index
         Case 0
              Nuevo
         Case 1
              Borrar
         Case 2
              Modificar
         Case 3
         Case 4
              Salir
  End Select
End Sub

Private Sub Nuevo()
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
     TxtDescripcion.SetFocus
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
  
  Dim nList As Integer, cRx As ClsLectura

  Set cRse = New ClsEscritura
  Set cRx = New ClsLectura
  
  nList = List1.ListIndex
  If cRx.TraerCantidad("CabComprobantes", "Vendedor", "Vendedor=" & Label3.Caption & " AND Tipo=3") = 0 Then
        If MsgBox("Borra La Cuenta Seleccionada ?", 20, "Atención") = 6 Then
           sCod = Label3.Caption
           cRse.Actualizar "CuentasBancariasBorrar", Nothing, Val(Label3.Caption)
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
   Else
     MsgBox "Cuenta Asociada con un Movimiento de Caja", vbCritical, "Atención"
   End If
  Exit Sub
errHandler:
  ManejaErrores
End Sub
Private Sub CargarLista()
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set Rs = New ADODB.Recordset
  
  List1.Clear
  cRsl.CargaCombo CmbBanco, "Proveedores", "Id", "RazonSocial", "CtaCte=1"
  Set Rs = cRsl.TraerRS("CuentasBancariasTraerTodas", , True)

  Do While Not Rs.EOF
     List1.AddItem Rs!Descripcion
     List1.ItemData(List1.NewIndex) = Rs!ID
     Rs.MoveNext
  Loop
  Set cRsl = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub Grabar()
  On Error GoTo errHandler
  
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  
  Rs1.Fields.Append "Banco", adInteger
  Rs1.Fields.Append "Descripcion", adVarChar, 50
  Rs1.Fields.Append "Cuenta", adVarChar, 20
  Rs1.Fields.Append "Limite", adCurrency
  Rs1.Fields.Append "Numero", adInteger
  Rs1.Fields.Append "Titular", adVarChar, 30
  
  Rs1.Open
  Rs1.AddNew
  Rs1!Banco = CmbBanco.ItemData(CmbBanco.ListIndex)
  Rs1!Descripcion = TxtDescripcion.text
  Rs1!Cuenta = TxtCuenta.text
  Rs1!Limite = TxtLimite.text
  Rs1!Numero = TxtNumero.text
  Rs1!Titular = TxtTitular.text
  Rs1.Update
  
  Set cRse = New ClsEscritura
  
  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Actualizar "CuentasBancariasModif", Rs1, Val(Label3.Caption)
  Else
     cRse.Actualizar "CuentasBancariasAgregar", Rs1
  End If
  Rs1.Close
  Set cRse = Nothing
  bFamilia = True
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub TxtLimite_Change()
  TodoBien
End Sub

Private Sub TxtLimite_GotFocus()
  TxtLimite.SelStart = 0
  TxtLimite.SelLength = Len(TxtLimite.text)
End Sub

Private Sub TxtLimite_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtLimite) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtLimite_LostFocus()
  If TxtLimite.text = "" Then TxtLimite.text = 0
  TxtLimite.text = Format(TxtLimite, nCantDecimales)
End Sub

Private Sub TxtNumero_Change()
  TodoBien
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
End Sub

