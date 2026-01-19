VERSION 5.00
Begin VB.Form PedidosPendientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Pendientes"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LstPedidos 
      Height          =   4110
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   3885
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      Height          =   360
      Left            =   2850
      TabIndex        =   0
      Top             =   4350
      Width           =   1170
   End
End
Attribute VB_Name = "PedidosPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rPedidos As ADODB.Recordset
Dim rPed As ADODB.Recordset

Private Sub CmdAplicar_Click()
  Dim i As Byte, cRem As ClsComprobantesL
  
  Set cRem = New ClsComprobantesL
    
  ' Pedidos
  For i = 0 To LstPedidos.ListCount - 1
      If LstPedidos.Selected(i) = True Then
         rPedidos.AddNew
         rPedidos!Movimiento = LstPedidos.ItemData(i)
         LstPedidos.ListIndex = i
         rPedidos!Descripcion = LstPedidos.Text
         rPedidos.Update
      End If
  Next i
  ' Devoluciones
  Unload Me
End Sub

Private Sub Form_Activate()
  If CmdAplicar.Enabled = True Then
     CmdAplicar.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Set rPedidos = New ADODB.Recordset
  rPedidos.Fields.Append "Movimiento", adDouble
  rPedidos.Fields.Append "Descripcion", adVarChar, 50
  rPedidos.Open
  LlenarGrilla
End Sub

Private Sub LlenarGrilla()
'  Dim cRem As ClsComprobantesL
'
'  Set cRem = New ClsComprobantesL
'  If nLlama = 2 Then
'     Set rPed = cRem.TraerClientesPedidos(StockComprobantes.TxtProveedor.Text)
'  Else
'     Set rPed = cRem.TraerClientesPedidos(VentaComprobantes.TxtCliente.Text)
'  End If
'  LstPedidos.Clear
'  Do While Not rPed.EOF
'     LstPedidos.AddItem rPed!Fecha & " " & rPed!Descripcion & " " & Format(rPed!Sucursal, "0000") & "-" & Format(rPed!Numero, "00000000")
'     LstPedidos.ItemData(LstPedidos.NewIndex) = rPed!Movimiento
'     LstPedidos.Selected(LstPedidos.ListCount - 1) = True
'     rPed.MoveNext
'  Loop
End Sub

