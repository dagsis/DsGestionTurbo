VERSION 5.00
Begin VB.Form TblComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABM de Comprobantes"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   6870
   Begin VB.Frame Frame2 
      Caption         =   "Actualiza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   3495
      TabIndex        =   21
      Top             =   1800
      Width           =   3210
      Begin VB.ComboBox CmbCaja 
         Height          =   315
         Left            =   2205
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1005
         Width           =   930
      End
      Begin VB.ComboBox CmbCtaCte 
         Height          =   315
         Left            =   2205
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   630
         Width           =   930
      End
      Begin VB.ComboBox CmbStock 
         Height          =   315
         Left            =   2205
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   225
         Width           =   930
      End
      Begin VB.CheckBox ChkIva 
         Caption         =   "Iva"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   285
         Width           =   660
      End
      Begin VB.CheckBox ChkCosto 
         Caption         =   "Discrimina"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Caja :"
         Height          =   195
         Left            =   1500
         TabIndex        =   28
         Top             =   1035
         Width           =   570
      End
      Begin VB.Label Label8 
         Caption         =   "Cta.Cte :"
         Height          =   180
         Left            =   1500
         TabIndex        =   26
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label7 
         Caption         =   "Stock :"
         Height          =   210
         Left            =   1500
         TabIndex        =   24
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.ComboBox CmbMovimiento 
      Height          =   315
      Left            =   1755
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1845
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Caption         =   "ABM de Comprobantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   105
      TabIndex        =   14
      Top             =   75
      Width           =   3300
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2625
         Width           =   2085
      End
      Begin VB.ComboBox CmbCobranza 
         Height          =   315
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2205
         Width           =   2085
      End
      Begin VB.ComboBox CmbOperacion 
         Height          =   315
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1368
         Width           =   2025
      End
      Begin VB.TextBox TxtNumero 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2025
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "TxtNumero"
         Top             =   3465
         Width           =   1065
      End
      Begin VB.TextBox TxtSucursal 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "TxtSucursal"
         Top             =   3030
         Width           =   1065
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Left            =   2340
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "TxtCodigo"
         Top             =   345
         Width           =   780
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   300
         Left            =   210
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "TxtDescripcion"
         Top             =   945
         Width           =   2910
      End
      Begin VB.Label Label11 
         Caption         =   "Moneda :"
         Height          =   210
         Left            =   180
         TabIndex        =   31
         Top             =   2625
         Width           =   825
      End
      Begin VB.Label Label10 
         Caption         =   "Cobranza :"
         Height          =   210
         Left            =   210
         TabIndex        =   30
         Top             =   2205
         Width           =   825
      End
      Begin VB.Label Label6 
         Caption         =   "T. de Ope. :"
         Height          =   225
         Left            =   210
         TabIndex        =   20
         Top             =   1365
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Número :"
         Height          =   225
         Left            =   210
         TabIndex        =   19
         Top             =   3480
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal  :"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   3045
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Movimiento :"
         Height          =   270
         Left            =   210
         TabIndex        =   17
         Top             =   1770
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   345
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   690
         Width           =   1110
      End
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   3510
      TabIndex        =   13
      Top             =   165
      Width           =   3240
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Nuevo"
      Height          =   810
      Index           =   0
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4035
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Borrar"
      Height          =   810
      Index           =   1
      Left            =   1245
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4035
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Modificar"
      Height          =   810
      Index           =   2
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4035
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Imprimir"
      Height          =   810
      Index           =   3
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4035
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Salir"
      Height          =   810
      Index           =   4
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4035
      Width           =   1005
   End
End
Attribute VB_Name = "TblComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCod As Integer, sw As Boolean, nId As Integer
Dim sCod1 As String
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


Private Sub CmbMovimiento_Click()
  TodoBien
End Sub

Private Sub CmbOperacion_Click()
  If sw = False Then
     Set cRsl = New ClsLectura
        cRsl.CargaCombo CmbMovimiento, "TipoMovimiento", "TipoMovimiento", "Descripcion", "Operacion=" & CmbOperacion.ItemData(CmbOperacion.ListIndex)
        If CmdBotones(0).Caption = "Grabar" Then
           CmbMovimiento.ListIndex = 0
        Else
          CmbMovimiento.text = cRsl.DatoCombo("TipoMovimiento", "TipoMovimiento", "Descripcion", Rs!TipoMovimiento)
        End If
        Select Case CmbOperacion.ItemData(CmbOperacion.ListIndex)
               Case 1
                 Label10.Caption = "Cobranza"
                 HabiTipo True, True, True, True, False
               Case 2
                 HabiTipo False, False, False, False, True
               Case 3
                 HabiTipo False, False, False, False, False
               Case 4
                 HabiTipo False, False, True, False, False
               Case 5
                 HabiTipo True, False, True, True, False
               Case 6
                    HabiTipo False, False, False, False, False
               Case 8
                 Label10.Caption = "Pagos"
                 HabiTipo True, True, True, True, True
               Case 9
                 HabiTipo False, False, True, False, False
               Case 10
                  HabiTipo True, False, True, True, False
        End Select
        Set cRsl = Nothing
   End If
   TodoBien
End Sub
Private Sub HabiTipo(bCtaCte As Boolean, bIva As Boolean, bCaja As Boolean, bStock As Boolean, bCosto As Boolean)
  CmbCtaCte.Enabled = bCtaCte
  ChkIva.Enabled = bIva
  CmbCobranza.Enabled = bIva
  CmbCaja.Enabled = bCaja
  CmbStock.Enabled = bStock
  ChkCosto.Enabled = bCosto
  If CmdBotones(0).Caption = "Grabar" Then
     If bStock = True Then CmbStock.ListIndex = 0
     If bCtaCte = True Then CmbCtaCte.ListIndex = 0
     If bCaja = True Then CmbCaja.ListIndex = 0
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
Private Sub Limpiar()
  Dim ctl As Control
  For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         ctl.text = ""
      End If
      If TypeOf ctl Is CheckBox Then
         ctl.Value = 0
      End If
      If TypeOf ctl Is ComboBox Then
         ctl.ListIndex = -1
      End If
  Next ctl
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
  Dim cRop As ClsLectura, cRmo As ClsLectura
  
  Set cRop = New ClsLectura
  Set cRmo = New ClsLectura
  On Error GoTo errHandler

  
  TxtCodigo.text = Rs!Comprobante
  TxtDescripcion.text = Rs!Descripcion
  TxtNumero.text = Format(Rs!Numero, "000000")
  TxtSucursal.text = Format(Rs!Sucursal, "000000")
  ChkIva.Value = Rs!Iva
  CmbCtaCte.text = Rs!CtaCte
  CmbCaja.text = Rs!Caja
  ChkCosto.Value = Rs!Costo
  CmbOperacion.text = cRop.DatoCombo("TipoOperacion", "TipoOperacion", "Descripcion", Rs!TipoOperacion)
  CmbMoneda.text = cRop.DatoCombo("Monedas", "Id", "Descripcion", Rs!Moneda)
  Set cRop = New ClsLectura
  cRmo.CargaCombo CmbMovimiento, "TipoMovimiento", "TipoMovimiento", "Descripcion", ""
  CmbMovimiento.text = cRmo.DatoCombo("TipoMovimiento", "TipoMovimiento", "Descripcion", Rs!TipoMovimiento)
  If Rs!Cobranza <> 0 Then
     CmbCobranza.text = cRmo.DatoCombo("Comprobantes", "Id", "Descripcion", Rs!Cobranza)
  Else
     CmbCobranza.ListIndex = CmbCobranza.ListCount - 1
  End If
  CmbStock.text = Rs!MoviStock
  Label10.Caption = "Cobranzas"
  Select Case CmbOperacion.ItemData(CmbOperacion.ListIndex)
         Case 1
              Label10.Caption = "Cobranzas"
         Case 8
              Label10.Caption = "Pagos"
  End Select
  
  
  Set cRop = Nothing
  Set cRmo = Nothing
Exit Sub

errHandler:
   ManejaErrores
  
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
 CargarCombo
 CargarLista
 If Rs.EOF = False Or Rs.BOF = False Then
    Botones pEdit, pBorr, pEdit, False, True
    Habilitar False, True
    List1.ListIndex = 0
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
 Set Rs = cRsl.TraerRS("ComprobanteTraerUno", sCod, True)
 Set cRsl = Nothing
 nId = Rs!ID
 LinkearTexto
Exit Sub
errHandler:
   ManejaErrores
End Sub
Private Sub TxtCodigo_LostFocus()
  Dim cRsl As ClsLectura, Rs1 As ADODB.Recordset
  Set cRsl = New ClsLectura
  Set Rs1 = New ADODB.Recordset

  If CmdBotones(2).Caption <> "Grabar" Then
     If TxtCodigo.text <> "" Then
        Set Rs1 = cRsl.TraerRS("ComprobanteTraerUnoS", TxtCodigo.text, True)
        If Rs1.EOF = False Or Rs1.BOF = False Then
           MsgBox "Codigo de Comprobante Repetido", vbCritical, "Atención"
           TxtCodigo.SetFocus
           TxtCodigo.text = ""
        End If
     End If
  Else
      Set Rs1 = cRsl.TraerRS("ComprobanteTraerUnoS", TxtCodigo.text, True)
      If Rs1.EOF = False Or Rs1.BOF = False Then
         If sCod1 <> TxtCodigo.text Then
            MsgBox "Codigo de Comprobante Repetido", vbCritical, "Atención"
            TxtCodigo.SetFocus
            TxtCodigo.text = ""
           End If
      End If
      Rs1.Close
  End If
  Set cRsl = Nothing
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
     Habilitar True, False
     HabiTipo False, False, False, False, False
     Botones False, False, False, False, True
     Limpiar
     CmdBotones(0).Caption = "Grabar"
     CmdBotones(0).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(4).Caption = "Cancelar"
     CmdBotones(4).Picture = LoadResPicture("Cancelar", 0)
     TxtCodigo.SetFocus
     sw = False
  Else
     Botones True, True, True, True, True
     Habilitar False, True
     CmdBotones(0).Caption = "Nuevo"
     CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
     CmdBotones(4).Caption = "Salir"
     CmdBotones(4).Picture = LoadResPicture("Salir", 0)
     sw = True
     Grabar
     CargarCombo
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
     sw = True
     Limpiar
     If CmdBotones(0).Caption = "Grabar" Then
        Habilitar False, True
        CmdBotones(0).Caption = "Nuevo"
        CmdBotones(0).Picture = LoadResPicture("Nuevo", 0)
        CmdBotones(4).Caption = "Salir"
        CmdBotones(4).Picture = LoadResPicture("Salir", 0)
     Else
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
     rList = List1.ListIndex
     nId = Rs!ID
     sCod1 = TxtCodigo.text
     Botones False, False, True, False, True
     Habilitar True, False
     CmdBotones(2).Caption = "Grabar"
     CmdBotones(2).Picture = LoadResPicture("Grabar", 0)
     CmdBotones(4).Caption = "Cancelar"
     CmdBotones(4).Picture = LoadResPicture("Cancelar", 0)
     CmdBotones(4).SetFocus
     sw = False
     CmbOperacion_Click
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
     CargarCombo
     List1.ListIndex = rList
  End If
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
  Set Rs = cRsl.TraerRsCondi("Comprobantes", "Id", "Suc=" & nSucursal)

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
  Set cRsl = New ClsLectura
  
  Set Rs1 = cRsl.RsVacio("Comprobantes", "Id", "N")
  Rs1.AddNew
 ' Rs1!Id = 0
  Rs1!Suc = nSucursal
  Rs1!Comprobante = TxtCodigo.text
  Rs1!Descripcion = TxtDescripcion.text
  
  If CmbOperacion.ListIndex <> -1 Then
     Rs1!TipoOperacion = CmbOperacion.ItemData(CmbOperacion.ListIndex)
  Else
     Rs1!TipoOperacion = Rs!TipoOperacion
  End If
  If CmbMovimiento.ListIndex <> -1 Then
     Rs1!TipoMovimiento = CmbMovimiento.ItemData(CmbMovimiento.ListIndex)
  Else
     Rs1!TipoMovimiento = Rs!TipoMovimiento
  End If
  Rs1!Iva = ChkIva.Value
  Rs1!Numero = IIf(TxtNumero.text = "", 0, TxtNumero.text)
  
  Rs1!CtaCte = IIf(CmbCtaCte.text = "", "No", CmbCtaCte.text)
  Rs1!Caja = IIf(CmbCaja.text = "", "No", CmbCaja.text)
  Rs1!Costo = ChkCosto.Value
  Rs1!MoviStock = IIf(CmbStock.text = "", "No", CmbStock.text)
  Rs1!Sucursal = IIf(TxtSucursal.text = "", 0, TxtSucursal.text)
  If CmbCobranza.ListIndex <> -1 Then
     Rs1!Cobranza = CmbCobranza.ItemData(CmbCobranza.ListIndex)
  Else
     Rs1!Cobranza = Rs!Cobranza
  End If
  Rs1!Moneda = CmbMoneda.ItemData(CmbMoneda.ListIndex)
  Rs1!Limite = 0
  Rs1!Afip = 0
  Rs1.Update
  Set cRse = New ClsEscritura
  
  If CmdBotones(2).Caption = "Grabar" Then
     cRse.Actualizar "ComprobanteModif", Rs1, nId
  Else
     cRse.Actualizar "ComprobanteAgregar", Rs1
  End If
  Rs1.Close
  Set cRse = Nothing
  Exit Sub
errHandler:
  ManejaErrores
End Sub

Private Sub TxtNumero_GotFocus()
  TxtNumero.SelStart = 0
  TxtNumero.SelLength = Len(TxtNumero.text)
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
  SoloNumero KeyAscii
End Sub

Private Sub TxtNumero_LostFocus()
  If TxtNumero = "" Then TxtNumero.text = 1
  TxtNumero = Format(TxtNumero.text, "000000")
End Sub

Private Sub TxtSucursal_GotFocus()
  TxtSucursal.SelStart = 0
  TxtSucursal.SelLength = Len(TxtSucursal.text)
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
  SoloNumero KeyAscii
End Sub

Private Sub TxtSucursal_LostFocus()
  If TxtSucursal.text = "" Then TxtSucursal.text = 1
  TxtSucursal.text = Format(TxtSucursal.text, "000000")
End Sub

Private Sub CargarCombo()
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbMovimiento, "TipoMovimiento", "TipoMovimiento", "Descripcion", ""
  cRsl.CargaCombo CmbMoneda, "Monedas", "Id", "Descripcion", ""
  Set cRsl = Nothing
  Set cRsl = New ClsLectura
  cRsl.CargaCombo CmbOperacion, "TipoOperacion", "TipoOperacion", "Descripcion", ""
  
  CmbStock.Clear
  CmbStock.AddItem "No"
  CmbStock.AddItem "Suma"
  CmbStock.AddItem "Resta"
  CmbStock.ListIndex = 0
  
  CmbCtaCte.Clear
  CmbCtaCte.AddItem "No"
  CmbCtaCte.AddItem "Debe"
  CmbCtaCte.AddItem "Haber"
  CmbCtaCte.ListIndex = 0

  CmbCaja.Clear
  CmbCaja.AddItem "No"
  CmbCaja.AddItem "Suma"
  CmbCaja.AddItem "Resta"
  CmbCaja.ListIndex = 0
  cRsl.CargaCombo CmbCobranza, "Comprobantes", "Id", "Descripcion", "TipoOperacion=4 or TipoOperacion=9"
  CmbCobranza.AddItem "Ninguno"
  CmbCobranza.ItemData(CmbCobranza.NewIndex) = 0
  CmbCobranza.ListIndex = CmbCobranza.ListCount - 1
  
End Sub

Private Sub TodoBien()
  If sw = False Then
     If CmdBotones(0).Caption = "Grabar" Then
        CmdBotones(0).Enabled = False
        If Len(TxtDescripcion.text) <> 0 And Len(CmbOperacion.text) <> 0 And Len(CmbMovimiento.text) <> 0 Then
           CmdBotones(0).Enabled = True
        End If
     Else
        CmdBotones(2).Enabled = False
        If Len(TxtDescripcion.text) <> 0 And Len(CmbOperacion.text) <> 0 And Len(CmbMovimiento.text) <> 0 Then
           CmdBotones(2).Enabled = True
        End If
     End If
  End If
End Sub
