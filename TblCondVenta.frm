VERSION 5.00
Begin VB.Form TblCondVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Forma de Pago"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleMode       =   0  'User
   ScaleWidth      =   10407.69
   Begin VB.Frame Frame2 
      Caption         =   "Tarjetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   3540
      TabIndex        =   22
      Top             =   2085
      Width           =   3690
      Begin VB.TextBox TxtIBTarjeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1230
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   615
         Width           =   540
      End
      Begin VB.TextBox TxtRgTarjeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3015
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   225
         Width           =   540
      End
      Begin VB.TextBox TxtIvaTarjeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1230
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label10 
         Caption         =   "% IB. Tarjeta :"
         Height          =   240
         Left            =   150
         TabIndex        =   25
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label Label9 
         Caption         =   "% RG. Tarjeta :"
         Height          =   210
         Left            =   1905
         TabIndex        =   24
         Top             =   255
         Width           =   1125
      End
      Begin VB.Label Label8 
         Caption         =   "% Iva Tarjeta :"
         Height          =   210
         Left            =   150
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   4
      Left            =   6218
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3247
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   810
      Index           =   3
      Left            =   3308
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3247
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   2258
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3247
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   1193
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3247
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   810
      Index           =   0
      Left            =   143
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3247
      Width           =   1005
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   3578
      TabIndex        =   13
      Top             =   150
      Width           =   3645
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   158
      TabIndex        =   9
      Top             =   75
      Width           =   3285
      Begin VB.ComboBox CmbValor 
         Height          =   315
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2205
         Width           =   1185
      End
      Begin VB.ComboBox CmbPCuota 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2580
         Width           =   2265
      End
      Begin VB.TextBox TxtDias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "Tex"
         Top             =   1770
         Width           =   435
      End
      Begin VB.TextBox TxtCuotas 
         Height          =   285
         Left            =   1305
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1785
         Width           =   285
      End
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         Left            =   1305
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1350
         Width           =   1905
      End
      Begin VB.TextBox TxtPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1305
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "TxtPorcentaje"
         Top             =   2205
         Width           =   495
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   90
         MaxLength       =   30
         TabIndex        =   0
         Text            =   "TxtDescripcion"
         Top             =   915
         Width           =   3105
      End
      Begin VB.Label Label3 
         Caption         =   "P.Cuota :"
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   2580
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "Cant.Dias :"
         Height          =   210
         Left            =   1860
         TabIndex        =   21
         Top             =   1785
         Width           =   840
      End
      Begin VB.Label Label7 
         Caption         =   "% Financiación :"
         Height          =   210
         Left            =   105
         TabIndex        =   20
         Top             =   2190
         Width           =   1230
      End
      Begin VB.Label Label6 
         Caption         =   "Cant.Cuotas :"
         Height          =   210
         Left            =   105
         TabIndex        =   19
         Top             =   1785
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Valor :"
         Height          =   210
         Left            =   90
         TabIndex        =   18
         Top             =   1350
         Width           =   840
      End
      Begin VB.Label LblCodigo 
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
         TabIndex        =   12
         Top             =   330
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   660
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   1665
         TabIndex        =   10
         Top             =   300
         Width           =   690
      End
   End
End
Attribute VB_Name = "TblCondVenta"
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
  LblCodigo.Caption = ""
  TxtCuotas.text = "01"
  TxtDias.text = "30"
  CmbTipo.ListIndex = 0
  TxtPorcentaje.text = Format(0, "0.00")
  TxtIvaTarjeta.text = Format(0, "0.00")
  TxtRgTarjeta.text = Format(0, "0.00")
  TxtIBTarjeta.text = Format(0, "0.00")
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
  Set cRsl = New ClsLectura
  
  LblCodigo.Caption = Format(Rs!CondVta, "00")
  TxtDescripcion.text = Rs!Descripcion
  CmbTipo.text = cRsl.DatoCombo("TipoValor", "Tipo", "Descripcion", Rs!tipo)
  CmbPCuota.text = cRsl.DatoCombo("TipoCuotas", "Id", "Descripcion", Rs!TipoCuota)
  TxtCuotas.text = Format(Rs!Cuotas, "00")
  TxtDias.text = Format(Rs!Dias, "00")
  TxtPorcentaje.text = Format(Rs!Financiacion, nCantDecimales)
  TxtIvaTarjeta.text = Format(Rs!IvaTarjeta, nCantDecimales)
  TxtRgTarjeta.text = Format(Rs!RgTarjeta, nCantDecimales)
  TxtIBTarjeta.text = Format(Rs!IBTarjeta, nCantDecimales)
  CmbValor.text = Rs!Mueve
End Sub

Private Sub Form_Load()
 On Error GoTo errHandler

 
 Set Rs = New ADODB.Recordset
 Set cRsl = New ClsLectura
 
 sw = True
 Habilitar False, False
 Botones pEdit, pBorr, pEdit, False, True
 CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
 CmdBotones(1).Picture = LoadResPicture("Borrar", 0)
 CmdBotones(2).Picture = LoadResPicture("Modificar", 0)
 CmdBotones(3).Picture = LoadResPicture("Imprimir", 0)
 CmdBotones(4).Picture = LoadResPicture("Salir", 0)
 cRsl.CargaCombo CmbTipo, "TipoValor", "Tipo", "Descripcion", ""
 cRsl.CargaCombo CmbPCuota, "TipoCuotas", "Id", "Descripcion", ""
 CmbValor.AddItem "Caja"
 CmbValor.AddItem "Banco"
 CmbValor.AddItem "No"
 
 CmbValor.text = "No"
 
 Limpiar
 CargarLista
  If Rs.EOF = False Or Rs.BOF = False Then
     Botones pEdit, pBorr, pEdit, False, True
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
 Set Rs = cRsl.TraerRS("CondVentaTraerUno", sCod, True)
 LinkearTexto
 Set cRsl = Nothing
Exit Sub
errHandler:
   ManejaErrores
End Sub
Private Sub TxtCuotas_GotFocus()
  TxtCuotas.SelStart = 0
  TxtCuotas.SelLength = Len(TxtCuotas)
End Sub
Private Sub TxtCuotas_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtCuotas_LostFocus()
  If TxtCuotas.text = "" Then TxtCuotas.text = 1
  TxtCuotas.text = Format(TxtCuotas.text, "00")
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
  Set Rs = cRsl.TraerRS("CondVentaTraerTodos", , True)

  Do While Not Rs.EOF
     List1.AddItem Rs!Descripcion
     List1.ItemData(List1.NewIndex) = Rs!CondVta
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
  Rs1.Fields.Append "Tipo", adInteger
  Rs1.Fields.Append "Cuotas", adInteger
  Rs1.Fields.Append "Dias", adInteger
  Rs1.Fields.Append "Financiacion", adSingle
  Rs1.Fields.Append "IvaTarjeta", adSingle
  Rs1.Fields.Append "RgTarjeta", adSingle
  Rs1.Fields.Append "IbTarjeta", adSingle
  Rs1.Fields.Append "TipoCuota", adInteger
  Rs1.Fields.Append "Mueve", adVarChar, 10
  
  Rs1.Open
  Rs1.AddNew
  Rs1!Descripcion = TxtDescripcion.text
  If CmbTipo.ListIndex <> -1 Then
     Rs1!tipo = CmbTipo.ItemData(CmbTipo.ListIndex)
  Else
     Rs1!tipo = Rs!tipo
  End If
  Rs1!Cuotas = TxtCuotas.text
  Rs1!Dias = TxtDias.text
  Rs1!Financiacion = TxtPorcentaje.text
  Rs1!IvaTarjeta = TxtIvaTarjeta.text
  Rs1!RgTarjeta = TxtRgTarjeta.text
  Rs1!IBTarjeta = TxtIBTarjeta.text
  Rs1!TipoCuota = CmbPCuota.ItemData(CmbPCuota.ListIndex)
  Rs1!Mueve = CmbValor.text
  Rs1.Update
  Set cRse = New ClsEscritura
  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Actualizar "CondVentaModif", Rs1, Val(LblCodigo.Caption)
  Else
     cRse.Actualizar "CondVentaAgregar", Rs1
  End If
  Rs1.Close
  Set cRse = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub TxtPorcentaje_Change()
  TodoBien
End Sub

Private Sub TxtPorcentaje_GotFocus()
  TxtPorcentaje.SelStart = 0
  TxtPorcentaje.SelLength = Len(TxtPorcentaje.text)
End Sub
Private Sub TxtPorcentaje_KeyPress(KeyAscii As Integer)
  If SoloNumeroDecimalFinal(KeyAscii, TxtPorcentaje) = False Then
     KeyAscii = 0
  End If
End Sub

Private Sub TxtPorcentaje_LostFocus()
  If TxtPorcentaje.text = "" Then TxtPorcentaje.text = 0
  TxtPorcentaje.text = Format(TxtPorcentaje.text, nCantDecimales)
End Sub
Private Sub TxtDias_GotFocus()
  TxtDias.SelStart = 0
  TxtDias.SelLength = Len(TxtDias)
End Sub
Private Sub TxtDias_KeyPress(KeyAscii As Integer)
  If SoloNumero(KeyAscii) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtDias_LostFocus()
  If TxtDias.text = "" Then TxtDias.text = 30
  TxtDias.text = Format(TxtDias.text, "00")
End Sub

Private Sub TxtIBTarjeta_GotFocus()
  TxtIBTarjeta.SelStart = 0
  TxtIBTarjeta.SelLength = Len(TxtIBTarjeta)
End Sub
Private Sub TxtIBTarjeta_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtIBTarjeta) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtIBTarjeta_LostFocus()
  If TxtIBTarjeta.text = "" Then TxtIBTarjeta.text = 0
  TxtIBTarjeta.text = Format(TxtIBTarjeta.text, nCantDecimales)
End Sub

Private Sub TxtIvaTarjeta_GotFocus()
  TxtIvaTarjeta.SelStart = 0
  TxtIvaTarjeta.SelLength = Len(TxtIvaTarjeta)
End Sub
Private Sub TxtIvaTarjeta_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtIvaTarjeta) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtIvaTarjeta_LostFocus()
  If TxtIvaTarjeta.text = "" Then TxtIvaTarjeta.text = 0
  TxtIvaTarjeta.text = Format(TxtIvaTarjeta.text, nCantDecimales)
End Sub

Private Sub TxtRgTarjeta_GotFocus()
  TxtRgTarjeta.SelStart = 0
  TxtRgTarjeta.SelLength = Len(TxtRgTarjeta)
End Sub
Private Sub TxtRgTarjeta_KeyPress(KeyAscii As Integer)
 If SoloNumeroDecimalFinal(KeyAscii, TxtRgTarjeta) = False Then
     KeyAscii = 0
  End If
End Sub
Private Sub TxtRgTarjeta_LostFocus()
  If TxtRgTarjeta.text = "" Then TxtRgTarjeta.text = 0
  TxtRgTarjeta.text = Format(TxtRgTarjeta.text, nCantDecimales)
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
