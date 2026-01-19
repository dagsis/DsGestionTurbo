VERSION 5.00
Begin VB.Form TblImpuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Impuestos"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   7095
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   4
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3960
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   810
      Index           =   3
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3975
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   2235
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3975
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3975
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   810
      Index           =   0
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3975
      Width           =   1005
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   3555
      TabIndex        =   10
      Top             =   225
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "ABM de Impuestos"
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
      Left            =   158
      TabIndex        =   6
      Top             =   135
      Width           =   3285
      Begin VB.CheckBox ChkRetencion 
         Caption         =   "Retenciones"
         Height          =   240
         Left            =   225
         TabIndex        =   19
         Top             =   3345
         Width           =   1305
      End
      Begin VB.CheckBox ChkPerce 
         Caption         =   "Percepciones"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   3075
         Width           =   1290
      End
      Begin VB.CheckBox ChkInterno 
         Caption         =   "No Gravados"
         Height          =   225
         Left            =   225
         TabIndex        =   5
         Top             =   2790
         Width           =   1275
      End
      Begin VB.TextBox TxtFijo 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2175
         MaxLength       =   7
         TabIndex        =   16
         Text            =   "TxtFijo"
         Top             =   1800
         Width           =   945
      End
      Begin VB.CheckBox ChkProducto 
         Caption         =   "Imp. a los Productos"
         Height          =   255
         Left            =   225
         TabIndex        =   4
         Top             =   2497
         Width           =   2040
      End
      Begin VB.CheckBox ChkDiscrimina 
         Caption         =   "Discrimina"
         Height          =   225
         Left            =   225
         TabIndex        =   3
         Top             =   2235
         Width           =   1140
      End
      Begin VB.TextBox TxtPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2175
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "TxtPorcentaje"
         Top             =   1395
         Width           =   945
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   195
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "TxtDescripcion"
         Top             =   915
         Width           =   2910
      End
      Begin VB.Label Label5 
         Caption         =   "Valor Fijo :"
         Height          =   240
         Left            =   225
         TabIndex        =   17
         Top             =   1770
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "Porcentaje :"
         Height          =   225
         Left            =   225
         TabIndex        =   15
         Top             =   1380
         Width           =   1215
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
         TabIndex        =   9
         Top             =   330
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   660
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   1665
         TabIndex        =   7
         Top             =   300
         Width           =   690
      End
   End
End
Attribute VB_Name = "TblImpuestos"
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
  TxtPorcentaje.text = "0.00"
  TxtFijo.text = "0.00"
  ChkDiscrimina.Value = 0
  ChkProducto.Value = 0
  ChkInterno.Value = 0
  ChkPerce.Value = 0
  ChkRetencion.Value = 0
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
  Label3.Caption = Format(Rs!Impuesto, "00")
  TxtDescripcion.text = Rs!Descripcion
  TxtPorcentaje.text = Format(Rs!Porcentaje, nCantDecimales)
  TxtFijo.text = Format(Rs!Fijo, nCantDecimales)
  ChkDiscrimina.Value = Rs!Discrimina
  ChkProducto.Value = Rs!Producto
  ChkInterno.Value = Rs!Interno
  ChkPerce.Value = Rs!Percepcion
  ChkRetencion.Value = Rs!Retencion
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
 
 Set cRsl = New ClsLectura
 sCod = List1.ItemData(List1.ListIndex)
 Set Rs = cRsl.TraerRS("ImpuestoTraerUno", sCod, True)
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
  
  Dim cGrab As ClsEscritura
  Static rList As Integer

  Set cGrab = New ClsEscritura

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
  Set cGrab = Nothing
Exit Sub
Errores:
   ManejaErrores

End Sub
Private Sub Borrar()
  MsgBox "Estas operaciones son solamente para el Programador", vbCritical, "Atención"
End Sub
Private Sub CargarLista()
  On Error GoTo errHandler
  
  Set cRsl = New ClsLectura
  Set Rs = New ADODB.Recordset
  
  List1.Clear
  Set Rs = cRsl.TraerRS("ImpuestoTraerTodos", , True)

  Do While Not Rs.EOF
     List1.AddItem Rs!Descripcion
     List1.ItemData(List1.NewIndex) = Rs!Impuesto
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
  
  Rs1.Fields.Append "Descripcion", adVarChar, 40
  Rs1.Fields.Append "Porcentaje", adSingle
  Rs1.Fields.Append "Fijo", adCurrency
  Rs1.Fields.Append "Discrimina", adVarChar, 1
  Rs1.Fields.Append "Producto", adVarChar, 1
  Rs1.Fields.Append "Interno", adVarChar, 1
  Rs1.Fields.Append "Percepcion", adVarChar, 1
  Rs1.Fields.Append "Retencion", adVarChar, 1
  Rs1.Open
  Rs1.AddNew
  Rs1!Descripcion = TxtDescripcion.text
  Rs1!Porcentaje = TxtPorcentaje.text
  Rs1!Fijo = TxtFijo.text
  Rs1!Discrimina = ChkDiscrimina.Value
  Rs1!Producto = ChkProducto.Value
  Rs1!Interno = ChkInterno.Value
  Rs1!Percepcion = ChkPerce.Value
  Rs1!Retencion = ChkRetencion.Value
  Rs1.Update
  Set cRse = New ClsEscritura
  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Actualizar "ImpuestoModif", Rs1, Val(Label3.Caption)
  Else
     cRse.Actualizar "ImpuestoAgregar", Rs1
  End If
  Rs1.Close
  Set cRse = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub
Private Sub TxtFijo_GotFocus()
  If TxtPorcentaje.text = 0 Then
     TxtFijo.SelStart = 0
     TxtFijo.SelLength = Len(TxtFijo.text)
  Else
     MsgBox "Debe Elegir un solo campo con Valor", vbCritical, "Atención"
     TxtPorcentaje.SetFocus
  End If
End Sub
Private Sub TxtFijo_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtFijo) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtFijo_LostFocus()
  If TxtFijo.text = "" Then TxtFijo.text = 0
  TxtFijo.text = Format(TxtFijo.text, nCantDecimales)
End Sub
Private Sub TxtPorcentaje_Change()
  TodoBien
End Sub
Private Sub TxtPorcentaje_GotFocus()
  If TxtFijo.text = 0 Then
     TxtPorcentaje.SelStart = 0
     TxtPorcentaje.SelLength = Len(TxtPorcentaje.text)
  Else
     MsgBox "Debe Elegir un solo campo con Valor", vbCritical, "Atención"
     TxtFijo.SetFocus
  End If
End Sub
Private Sub TxtPorcentaje_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtPorcentaje) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPorcentaje_LostFocus()
  TxtPorcentaje.text = Format(TxtPorcentaje.text, nCantDecimales)
End Sub

Private Sub TodoBien()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtDescripcion.text) <> 0 And Len(TxtPorcentaje.text) <> 0 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(2).Enabled = False
        If Len(TxtDescripcion.text) <> 0 And Len(TxtPorcentaje.text) <> 0 Then
           CmdBotones(2).Enabled = True
        End If
     End If
  End If
End Sub
