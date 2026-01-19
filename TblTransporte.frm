VERSION 5.00
Begin VB.Form TblTransporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transporte"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   6435
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   4
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3345
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   810
      Index           =   3
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3345
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3345
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3345
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   810
      Index           =   0
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3345
      Width           =   1005
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   3525
      TabIndex        =   11
      Top             =   120
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Caption         =   "ABM de Transporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   165
      TabIndex        =   4
      Top             =   30
      Width           =   3285
      Begin VB.TextBox TxtPacoplado 
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "TxtPac"
         Top             =   2835
         Width           =   900
      End
      Begin VB.TextBox TxtPChasis 
         Height          =   285
         Left            =   195
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "Txt"
         Top             =   2835
         Width           =   900
      End
      Begin VB.TextBox TxtCuit 
         Height          =   285
         Left            =   195
         MaxLength       =   13
         TabIndex        =   2
         Text            =   "TxtCuit"
         Top             =   2160
         Width           =   1290
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   285
         Left            =   195
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "TxtDireccion"
         Top             =   1515
         Width           =   2910
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   195
         MaxLength       =   30
         TabIndex        =   0
         Text            =   "TxtDescripcion"
         Top             =   855
         Width           =   2910
      End
      Begin VB.Label Label7 
         Caption         =   "P.Acoplado :"
         Height          =   195
         Left            =   1560
         TabIndex        =   18
         Top             =   2565
         Width           =   1110
      End
      Begin VB.Label Label6 
         Caption         =   "P.Chasis :"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   2565
         Width           =   1110
      End
      Begin VB.Label Label5 
         Caption         =   "Cuit :"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   1890
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   1245
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
         Left            =   2430
         TabIndex        =   10
         Top             =   330
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   585
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   1665
         TabIndex        =   6
         Top             =   300
         Width           =   690
      End
   End
End
Attribute VB_Name = "TblTransporte"
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
  Label3.Caption = Format(Rs!ID, "00")
  TxtDescripcion.text = Rs!Descripcion
  TxtDireccion.text = "" & Rs!Domicilio
  TxtCuit.text = "" & Rs!Cuit
  TxtPChasis.text = "" & Rs!PatenteChasis
  TxtPacoplado.text = "" & Rs!PatenteAcoplado
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
 CmdBotones(3).Picture = LoadResPicture("Imprimir", 0)
 CmdBotones(4).Picture = LoadResPicture("Salir", 0)
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

Private Sub List1_Click()
 On Error GoTo errHandler
 Dim sCond As String, sCod As Integer
 
 Set cRsl = New ClsLectura
 sCod = List1.ItemData(List1.ListIndex)
 Set Rs = cRsl.TraerRS("TransporteTraerUno", sCod, True)
 LinkearTexto
 Set cRsl = Nothing
Exit Sub
errHandler:
   ManejaErrores
End Sub
Private Sub TxtDescripcion_Change()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtDescripcion.text) <> 0 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(2).Enabled = False
        If Len(TxtDescripcion.text) <> 0 Then
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
  If cRx.TraerCantidad("Clientes", "Transporte", "Transporte=" & Label3.Caption) = 0 Then
    If MsgBox("Borra El Transporte Seleccionado ?", 20, "Atención") = 6 Then
       sCod = Label3.Caption
       cRse.Actualizar "TransporteBorrar", Nothing, Val(Label3.Caption), "Transporte"
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
     MsgBox "Transporte Asociado a un Cliente", vbCritical, "Atención"
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
  Set Rs = cRsl.TraerRS("TransporteTraerTodos", , True)

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
  Dim cRx As ClsLectura
  
  On Error GoTo errHandler
  
  Dim Rs1 As ADODB.Recordset
  
  Set cRx = New ClsLectura
  
  Set Rs1 = cRx.RsVacio("Transporte", "Id", "N")
  
  Rs1.AddNew
  Rs1!Descripcion = TxtDescripcion.text
  Rs1!Domicilio = TxtDireccion.text
  Rs1!Cuit = TxtCuit.text
  Rs1!PatenteChasis = TxtPChasis.text
  Rs1!PatenteAcoplado = TxtPacoplado.text
  Rs1.Update
  
  Set cRse = New ClsEscritura
  
  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Actualizar "TransporteModif", Rs1, Val(Label3.Caption)
  Else
     cRse.Actualizar "TransporteAgregar", Rs1
  End If
  Rs1.Close
  Set cRse = Nothing
  bFamilia = True
  Exit Sub
errHandler:
  ManejaErrores
End Sub

