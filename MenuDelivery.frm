VERSION 5.00
Begin VB.Form MenuDelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   1425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   1425
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MenuDelivery 
      Caption         =   "Delivery"
      Begin VB.Menu MenuMovi 
         Caption         =   "Nuevo Pedido"
      End
      Begin VB.Menu MenuPedido 
         Caption         =   "Ver Pedido"
      End
      Begin VB.Menu MenuAnular 
         Caption         =   "Anular"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSalidaMoto 
         Caption         =   "Salida Moto"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "MenuDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MenuAnular_Click()
  Dim sSql As String
  If MsgBox("Anula el Delivery elegido  ?", 20, "Mensaje de A&C IT") = 6 Then
     Delivery.rPed.Bookmark = Delivery.Grid1.Bookmark
     db.Open sDb
     db.BeginTrans

     sSql = "DELETE FROM Auxiliar WHERE a_Mesa=" & Delivery.rPed![a_Mesa] & " AND a_mensaje='DELIVERY'"
     db.Execute sSql
    
     db.CommitTrans
     db.Close
     Delivery.LlenarGrilla
   End If

End Sub

Private Sub Menucobranzas_Click()
  ' CobDelivery.Show 1
End Sub

Private Sub MenuMovi_Click()
   NuevoDelivery.Show 1
End Sub

Private Sub MenuPedido_Click()
   bDelivery = True
   Delivery.rPed.Bookmark = Delivery.Grid1.Bookmark
   nMesa = Delivery.rPed![a_Mesa]
   frmMesa.Show 1
End Sub

Private Sub MenuSalidaMoto_Click()
  Dim sSql As String, sTime As String
  
  sTime = Time

  Delivery.rPed.Bookmark = Delivery.Grid1.Bookmark
  Delivery.rPed![a_puerto3] = sTime
  Delivery.rPed.Update
  
  db.Open sDb
  db.BeginTrans
  
  
  sSql = "UPDATE Auxiliar SET a_puerto3='" & sTime & "' WHERE Id=" & Delivery.rPed![Id]
  db.Execute sSql
    
  sSql = "UPDATE Auxiliar SET a_Mozo=" & Delivery.CmbMoto.ItemData(Delivery.CmbMoto.ListIndex) & " WHERE a_mesa=" & Delivery.rPed![a_Mesa]
  db.Execute sSql
  
  db.CommitTrans
  db.Close
End Sub
