VERSION 5.00
Begin VB.Form TblFamiliaSub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Sub Familias"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   6945
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   4
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3435
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   810
      Index           =   3
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3435
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   2175
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3435
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3435
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   810
      Index           =   0
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3435
      Width           =   1005
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   3405
      TabIndex        =   8
      Top             =   510
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "ABM de Sub Familias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3285
      Begin VB.ComboBox CmbGrupo 
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   810
         Width           =   2325
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   195
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "TxtDescripcion"
         Top             =   1530
         Width           =   2910
      End
      Begin VB.Label Label4 
         Caption         =   "Familia :"
         Height          =   210
         Left            =   195
         TabIndex        =   7
         Top             =   810
         Width           =   780
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
         TabIndex        =   6
         Top             =   330
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   1260
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   1665
         TabIndex        =   4
         Top             =   300
         Width           =   690
      End
   End
   Begin VB.ComboBox CmbOrden 
      Height          =   315
      Left            =   3405
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   3420
   End
End
Attribute VB_Name = "TblFamiliaSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sw As Boolean, sTabla As String, sCampo As String
Dim sCod As Integer
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

Private Sub CmbEstado_Click()
  TodoBien
End Sub
Private Sub CmbGrupo_Click()
  TodoBien
End Sub

Private Sub CmbOrden_Click()
   CargarLista
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
  CmbGrupo.ListIndex = -1
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
  CmbOrden.Enabled = bList
End Sub

Private Sub LinkearTexto()
  
  Set cRsl = New ClsLectura
  
  Label3.Caption = Format(Rs!ID, "000")
  TxtDescripcion.text = Rs!Descripcion
  CmbGrupo.text = cRsl.DatoCombo("Familias", "Familia", "Descripcion", Rs!Familia)
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
 
 cRsl.CargaCombo CmbGrupo, "Familias", "Familia", "Descripcion", ""
 cRsl.CargaCombo CmbOrden, "Familias", "Familia", "Descripcion", ""
 CmbOrden.ListIndex = 0
  
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
 
 Dim sCod As Long
 
 Set cRsl = New ClsLectura
 
  sCod = List1.ItemData(List1.ListIndex)
  Set Rs = cRsl.TraerRsCondi("FamiliasSub", "Id", "Id=" & sCod)
  LinkearTexto
  Set cRsl = Nothing
Exit Sub
errHandler:
   ManejaErrores

End Sub
Private Sub TxtDescripcion_Change()
  TodoBien
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
     CmbGrupo.text = CmbOrden.text
     TxtDescripcion.SetFocus
  Else
     Botones True, True, True, True, True
     Habilitar False, True
     Grabar
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(4).Caption = "Salir"
     CmdBotones(4).Picture = LoadResPicture("Salir", 0)
     sw = True
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

  Set cRse = New ClsEscritura

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
     Dim nEst As Integer
    
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
  
  Dim nList As Integer, cDep As ClsComprobantesE
  Set cRse = New ClsEscritura
  Set cDep = New ClsComprobantesE
  
  nList = List1.ListIndex
  If MsgBox("Borra la Sub Familia Seleccionada ?", 20, "Atención") = 6 Then
     sCod = Label3.Caption
     cRse.Actualizar "FamiliasSub_B", Nothing, sCod
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
  Set cRse = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub
Private Sub CargarLista()
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  
  List1.Clear
  Set Rs = cRsl.TraerRsCondi("FamiliasSub", "Descripcion", "Familia=" & CmbOrden.ItemData(CmbOrden.ListIndex))
  
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
  
  Rs1.Fields.Append "Familia", adInteger
  Rs1.Fields.Append "Descripcion", adVarChar, 30
  Rs1.Open
  Rs1.AddNew
  Rs1!Descripcion = TxtDescripcion.text
  If CmbGrupo.ListIndex <> -1 Then
     Rs1!Familia = CmbGrupo.ItemData(CmbGrupo.ListIndex)
  Else
     Rs1!Familia = Rs!Familia
  End If
  Rs1.Update
  
  Set cRse = New ClsEscritura
  
  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Actualizar "FamiliasSub_M", Rs1, Val(Label3.Caption)
  Else
     cRse.Actualizar "FamiliasSub_A", Rs1
  End If
  Rs1.Close
  bDeposito = True
  Set cRse = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub TodoBien()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtDescripcion.text) <> 0 And Len(CmbGrupo.text) <> 0 Then
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


