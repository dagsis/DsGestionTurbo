VERSION 5.00
Begin VB.Form VentaRemitosPendientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos Pendientes"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkAgruRemi 
      Caption         =   "Agrupar"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   2385
   End
   Begin VB.CheckBox ChkAgrupar 
      Caption         =   "Actualizar de Lista de Precio"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   4425
      Width           =   2385
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   360
      Left            =   2805
      TabIndex        =   1
      Top             =   4395
      Width           =   1170
   End
   Begin VB.ListBox LstRemitos 
      Height          =   4110
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   150
      Width           =   3885
   End
End
Attribute VB_Name = "VentaRemitosPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rRemitos As ADODB.Recordset
Dim rRem As ADODB.Recordset

Private Sub CmdAplicar_Click()
  Dim i As Byte, cRem As ClsComprobantesL
  
  Set cRem = New ClsComprobantesL
    
  ' Remitos
  For i = 0 To LstRemitos.ListCount - 1
      If LstRemitos.Selected(i) = True Then
         rRemitos.AddNew
         rRemitos!Movimiento = LstRemitos.ItemData(i)
         LstRemitos.ListIndex = i
         rRemitos!Descripcion = LstRemitos.text
         rRemitos.Update
      End If
  Next i
  bAgrupar = False
  bAgruRemito = False
  
  If ChkAgrupar.Value = 1 Then
     bAgrupar = True
  End If
  If ChkAgruRemi.Value = 1 Then
     bAgruRemito = True
  End If
  Unload Me
End Sub

Private Sub Form_Activate()
  If CmdAplicar.Enabled = True Then
     CmdAplicar.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Set rRemitos = New ADODB.Recordset
  rRemitos.Fields.Append "Movimiento", adDouble
  rRemitos.Fields.Append "Descripcion", adVarChar, 50
  rRemitos.Open
  LlenarGrilla
End Sub

Private Sub LlenarGrilla()
  On Error GoTo Errores
  Dim cRem As ClsComprobantesL, nCantRem As Single, nCantRemAux As Single, xs As Boolean
  Dim rAux As ADODB.Recordset
  
  Set cRem = New ClsComprobantesL
  Set rAux = New ADODB.Recordset
  
  Select Case nLlama
         Case 1
             Set rRem = cRem.TraerClientesRemitos(VentaComprobantes.TxtCliente.text)
         Case 2
              Set rRem = cRem.TraerProveedorRemitos(ComprasComprobantes.TxtCliente.text)
  End Select
  LstRemitos.Clear
  Do While Not rRem.EOF
     xs = True
     Set rAux = cRem.TraerDetallesMovimiento(rRem!Movimiento)
     rAux.MoveFirst
     Do While Not rAux.EOF
   '  If Not IsNull(rAux!Producto) And rAux!Producto <> "" Then
        nCantRem = Round(cRem.TraerCantidadRemito(rAux!Producto, rRem!Movimiento, rAux!ID), 2)
        nCantRemAux = Round(IIf(IsNull(rAux!Cantidad), 0, rAux!Cantidad) - Abs(nCantRem), 2)
        If nCantRemAux > 0 Or nCantRem = -1 Then
           LstRemitos.AddItem rRem!Descripcion & " " & Format(rRem!Numero, "00000")
           LstRemitos.ItemData(LstRemitos.NewIndex) = rRem!Movimiento
           LstRemitos.Selected(LstRemitos.ListCount - 1) = True
           xs = False
           Exit Do
        End If
    '  Else
    '    LstRemitos.AddItem rRem!Descripcion & " " & Format(rRem!Numero, "00000000")
    '    LstRemitos.ItemData(LstRemitos.NewIndex) = rRem!Movimiento
    '    LstRemitos.Selected(LstRemitos.ListCount - 1) = True
    '    Exit Do
    '  End If
       rAux.MoveNext
     Loop
     If xs = True Then
        db.Open sDb
        db.Execute "UPDATE CabComprobantes SET Anulado=1 where Movimiento=" & rRem!Movimiento
        db.Close
     End If
     rRem.MoveNext
  Loop
  If LstRemitos.ListCount = 0 Then
     Unload Me
  End If
Exit Sub
Errores:
   MsgBox err.Description
End Sub
