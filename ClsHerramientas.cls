VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsHerramientas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub BorrarLasTablas()
  Dim Ub As ADODB.Connection
    
  Set Ub = New ADODB.Connection
  
  Ub.Open Us
  Ub.Execute "DELETE  FROM PRODUCTOS"
  Ub.Close
End Sub

Public Function RsVacio(pTabla As String, pCampo As String, pTipo As String) As ADODB.Recordset
  Dim Ub As ADODB.Connection, rs As ADODB.Recordset
  Dim sSql As String
  
  
  Set Ub = New ADODB.Connection
  Set rs = New ADODB.Recordset
  
  Ub.Open Us
    
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  Select Case pTipo
    Case "N"
       sSql = "SELECT * FROM " & pTabla & " WHERE " & pCampo & "=0"
    Case "S"
       sSql = "SELECT * FROM " & pTabla & " WHERE " & pCampo & "=''"
  End Select
  rs.CursorLocation = adUseClient
  rs.Open sSql, Ub, adOpenStatic, adLockOptimistic
'  Rs.ActiveConnection = Nothing
  Set RsVacio = rs
' Ub.Close
 
Exit Function
errHandler:
  Err.Raise Err.Number
End Function

Public Function TraerValorDeUnCampo(pTabla As String, pDescri As String, pCondi As String) As Variant
  Dim Ub As ADODB.Connection, rs As ADODB.Recordset
  Dim sSql As String
  
  
  Set Ub = New ADODB.Connection
  Set rs = New ADODB.Recordset
  
  Ub.Open Us
    
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  If pCondi <> "" Then
     sSql = "SELECT " & pDescri & " as Descripcion FROM " & pTabla & " WHERE  " & pCondi
  Else
     sSql = "SELECT " & pDescri & " as Descripcion FROM " & pTabla
  End If
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, Ub, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  If rs.RecordCount <> 0 Then
     TraerValorDeUnCampo = IIf(IsNull(rs!Descripcion), "", Trim(rs!Descripcion))
  Else
     TraerValorDeUnCampo = 0
  End If
 Ub.Close
  
Exit Function
errHandler:
  Err.Raise Err.Number

End Function

Public Function TraerTodos(pTabla As String, pCampo As String) As ADODB.Recordset
  Dim Ub As ADODB.Connection, rs As ADODB.Recordset
  Dim sSql As String
  
  
  Set Ub = New ADODB.Connection
  Set rs = New ADODB.Recordset
  
  Ub.Open Us
    
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset

  If pCampo <> "*" Then
     sSql = "SELECT " & pCampo & " FROM " & pTabla & " ORDER BY " & pCampo
  Else
     sSql = "SELECT " & pCampo & " FROM " & pTabla
  End If
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, Ub, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerTodos = rs
 Ub.Close
 
Exit Function
errHandler:
  Err.Raise Err.Number
  
End Function

