VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsClienteE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Agregar(pRsCliente As ADODB.Recordset, pTip As Byte)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  If pTip = 0 Then
     Com.CommandText = "ClienteAgregar"
  Else
     Com.CommandText = "ProveedorAgregar"
  End If
    
  ' recorre la coleccion fields y carga los parametros de la tabla de Clientes
  
  iCampos = pRsCliente.Fields.Count - 1
  For i = 0 To iCampos
      If pRsCliente(i).Name <> "Id" Then
         If pRsCliente(i).Type = adDouble Or pRsCliente(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsCliente(i).Name, adDouble, adParamInput, pRsCliente(i).DefinedSize, pRsCliente(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsCliente(i).Name, pRsCliente(i).Type, adParamInput, pRsCliente(i).DefinedSize, pRsCliente(i).Value)
         End If
      End If
  Next
  Com.Execute
  Set Com = Nothing
  
  ' graba los impuestos asociados
  db.CommitTrans
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    
    MsgBox err.Description, vbCritical, "Atención"
    db.Close
End Sub

Public Sub Modificar(pRsCliente As ADODB.Recordset, Cod As String, pTip As Byte)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
    
  db.Open sDb
  db.BeginTrans
  
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  If pTip = 0 Then
     Com.CommandText = "ClienteModif"
  Else
     Com.CommandText = "ProveedorModif"
  End If
    
'  Com.Parameters.Append Com.CreateParameter("Param", adVarChar, adParamInput, Len(Cod), Cod)
' recorre la coleccion fields y carga los parametros de la tabla de cliente

  iCampos = pRsCliente.Fields.Count - 1
  For i = 0 To iCampos
      If pRsCliente(i).Name <> "Id" Then
         If pRsCliente(i).Type = adDouble Or pRsCliente(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsCliente(i).Name, adDouble, adParamInput, pRsCliente(i).DefinedSize, pRsCliente(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsCliente(i).Name, pRsCliente(i).Type, adParamInput, pRsCliente(i).DefinedSize, pRsCliente(i).Value)
         End If
      End If
  Next
  Com.Execute
  Set Com = Nothing
  
  db.CommitTrans
  
  db.Close
  Exit Sub
    
errHandler:
    db.RollbackTrans
    db.Close
    err.Raise err.Number
    
End Sub

Public Sub Borrar(pStoreProc As String, Optional pRsParam As ADODB.Recordset, Optional Cod As Variant = 0)
  Dim iCampos As Integer
  Dim Com As ADODB.Command
  Dim i As Integer
  On Error GoTo errHandler
  Set Com = New ADODB.Command
  
  db.Open sDb
  
  db.BeginTrans
  Com.ActiveConnection = db
  Com.CommandType = adCmdStoredProc
  Com.CommandText = pStoreProc
    
    ' recorre la coleccion fields y carga los parametros
  If Cod <> 0 Then
     If IsNumeric(Cod) Then
        Com.Parameters.Append Com.CreateParameter("Param", adInteger, adParamInput, , Cod)
     Else
        Com.Parameters.Append Com.CreateParameter("Param", adVarChar, adParamInput, Len(Cod), Cod)
     End If
  End If
  If Not pRsParam Is Nothing Then
     iCampos = pRsParam.Fields.Count - 1
     For i = 0 To iCampos
         If pRsParam(i).Name <> "Id" Then
         If pRsParam(i).Type = adDouble Or pRsParam(i).Type = adNumeric Then
            Com.Parameters.Append Com.CreateParameter(pRsParam(i).Name, adDouble, adParamInput, pRsParam(i).DefinedSize, pRsParam(i).Value)
         Else
            Com.Parameters.Append Com.CreateParameter(pRsParam(i).Name, pRsParam(i).Type, adParamInput, pRsParam(i).DefinedSize, pRsParam(i).Value)
         End If
      End If
     Next
  End If
  Com.Execute
  Set Com = Nothing
  
  
  db.CommitTrans
  db.Close
  
 Exit Sub
    
errHandler:
    err.Raise err.Number
    db.RollbackTrans
End Sub


