VERSION 5.00
Begin VB.Form FrmBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   210
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   6600
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   210
      Index           =   1
      Left            =   1335
      TabIndex        =   5
      Top             =   6600
      Width           =   1290
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   210
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   6600
      Width           =   1050
   End
   Begin VB.CheckBox Check1 
      Caption         =   "B.Avanzada"
      Height          =   270
      Left            =   195
      TabIndex        =   3
      Top             =   6930
      Width           =   1830
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   7020
   End
   Begin VB.ListBox List1 
      Height          =   5910
      ItemData        =   "FrmBuscar.frx":0000
      Left            =   195
      List            =   "FrmBuscar.frx":0002
      TabIndex        =   1
      Top             =   630
      Width           =   7005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6180
      TabIndex        =   2
      Top             =   6600
      Width           =   1065
   End
End
Attribute VB_Name = "FrmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bNo As Boolean, sBuscar As String, rBuscar As New ADODB.Recordset, ulti As Integer
Dim cBus As ClsLectura
Dim sTabla As String, sDescripcion As String, nBuscar1 As Byte
Dim nI As Integer

Private Sub Check1_Click()
  Text1.SetFocus
End Sub

Private Sub Command1_Click()
  nDat = 0
  Unload Me
End Sub

Private Sub Form_Activate()
  Option1(1).Value = True
  Cambiar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
     KeyAscii = 0
     SendKeys "{TAB}"
    Exit Sub
 End If
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
     KeyAscii = 0
     List1_DblClick
     Exit Sub
  End If
End Sub

Private Sub List1_DblClick()
  If List1.ListIndex <> -1 Then
     nDat = List1.ItemData(List1.ListIndex)
  End If
  Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
  Select Case Index
         Case 0
         If Option1(Index).Value = True Then nBuscar1 = 1
         Case 1
         If Option1(Index).Value = True Then nBuscar1 = 2
         Case 2
         If Option1(Index).Value = True Then nBuscar1 = 3
  End Select
  Cambiar
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim sCod As String, sSaldo As Currency, cDesc As String

  Set cBus = New ClsLectura
  bNo = False
  On Error GoTo Errores
  If KeyAscii <> 8 Then
     Set rBuscar = cBus.Buscar(sTabla, sDescripcion, "Id", Text1.text & Chr$(KeyAscii), Check1.Value)
     ulti = Len(Text1.text + Chr$(KeyAscii))
  Else
     If ulti <> 1 Then
        Set rBuscar = cBus.Buscar(sTabla, sDescripcion, "Id", Left$(Text1.text, ulti - 1), Check1.Value)
        ulti = ulti - 1
     Else
        Set rBuscar = cBus.BuscarTodos(sTabla, sDescripcion, "Id")
     End If
  End If
  If KeyAscii <> 8 Then
     List1.Clear
     If Not rBuscar.EOF Then
        For nI = 1 To rBuscar.RecordCount
            Select Case nBuscar
                 Case 1
                      sCod = cBus.TraerValorDeUnCampo("Clientes", "Cliente", "Id=" & rBuscar!codigo)
                      sSaldo = cBus.TraerSaldos(rBuscar!codigo)
                      List1.AddItem "(" & sCod & ") " & rBuscar!Descripcion & " - $ " & sSaldo & " " & ")"
                      List1.ItemData(List1.NewIndex) = rBuscar!codigo
                 Case 2
                      sCod = cBus.TraerValorDeUnCampo("Proveedores", "Proveedor", "Id=" & rBuscar!codigo)
                      List1.AddItem rBuscar!Descripcion
                      List1.ItemData(List1.NewIndex) = rBuscar!codigo
                 Case 3
                      sCod = cBus.TraerValorDeUnCampo("Productos", "Producto", "Id=" & rBuscar!codigo)
                      cDesc = cBus.TraerValorDeUnCampo("Productos", "Descripcion1", "Id=" & rBuscar!codigo)
                      List1.AddItem "(" & sCod & ") " & rBuscar!Descripcion & " - " & cDesc & " "
                      List1.ItemData(List1.NewIndex) = rBuscar!codigo
                 Case 4
            End Select
            rBuscar.MoveNext
            If rBuscar.EOF Then Exit For
        Next nI
     End If
  Else
     If Text1.text <> "" Then
        List1.Clear
        If rBuscar.RecordCount <> 0 Then
        For nI = 1 To rBuscar.RecordCount
            Select Case nBuscar
                 Case 1
                      sCod = cBus.TraerValorDeUnCampo("Clientes", "Cliente", "Id=" & rBuscar!codigo)
                      sSaldo = cBus.TraerSaldos(rBuscar!codigo)
                      List1.AddItem "(" & sCod & ") " & rBuscar!Descripcion & " - $ " & sSaldo & " " & ")"
                      List1.ItemData(List1.NewIndex) = rBuscar!codigo
                 Case 2
                      sCod = cBus.TraerValorDeUnCampo("Proveedores", "Proveedor", "Id=" & rBuscar!codigo)
                      List1.AddItem rBuscar!Descripcion
                      List1.ItemData(List1.NewIndex) = rBuscar!codigo
                 Case 3
                      sCod = cBus.TraerValorDeUnCampo("Productos", "Producto", "Id=" & rBuscar!codigo)
                      cDesc = cBus.TraerValorDeUnCampo("Productos", "Descripcion1", "Id=" & rBuscar!codigo)
                      List1.AddItem "(" & sCod & ") " & rBuscar!Descripcion & " - " & cDesc & " "
                      List1.ItemData(List1.NewIndex) = rBuscar!codigo
                 Case 4
            End Select
            rBuscar.MoveNext
            If rBuscar.EOF Then Exit For
        Next nI
        End If
      End If
  End If
  rBuscar.Close
  Exit Sub
Errores:
    If err.Number = 5 Then bNo = True: Exit Sub
    MsgBox err.Description
End Sub

Private Sub Cambiar()
  Text1.text = ""
  Text1.SetFocus
  List1.Clear
  Select Case nBuscar
         Case 1
              Option1(0).Caption = "Código"
              Option1(1).Caption = "Nombre"
              Option1(2).Visible = False
              Select Case nBuscar1
                     Case 1
                          sTabla = "Clientes"
                          sDescripcion = "Cliente"
                     Case 2
                          sTabla = "Clientes"
                          sDescripcion = "RazonSocial"
              End Select
         Case 2
              Option1(0).Caption = "Código"
              Option1(1).Caption = "Nombre"
              Option1(2).Visible = False
              Select Case nBuscar1
                     Case 1
                          sTabla = "Proveedores"
                          sDescripcion = "Proveedor"
                     Case 2
                          sTabla = "Proveedores"
                          sDescripcion = "RazonSocial"
              End Select
         Case 3
              Option1(0).Caption = "Código"
              Option1(1).Caption = "Descripción"
              Option1(2).Caption = "Desc. Adicional"
              Select Case nBuscar1
                     Case 1
                          sTabla = "Productos"
                          sDescripcion = "Producto"
                     Case 2
                          sTabla = "Productos"
                          sDescripcion = "Descripcion"
                     Case 3
                          sTabla = "Productos"
                          sDescripcion = "Descripcion1"
              End Select
  End Select
End Sub
