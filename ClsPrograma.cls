VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPerm As Object ' Scripting.Dictionary (late-bound, no referencias)
Private mUsuario As Long

Private mMap As Object ' Dictionary de dictionaries


Public Function DameEstadoMesa(Estado As String, nMesa As Integer) As Boolean
  Dim sSql As String
  Dim rs As Recordset
  
  db.Open sDb
  
  Set rs = New Recordset
    
  If Estado = "VACIA" Then
     sSql = "SELECT * FROM AUXILIAR WHERE "
     sSql = sSql & "A_MESA = " & nMesa & " AND A_MENSAJE <> 'DELIVERY'"
  Else
    sSql = "SELECT * FROM AUXILIAR WHERE "
    sSql = sSql & "A_MESA = " & nMesa
    sSql = sSql & " AND A_MENSAJE = '" & Estado & "'"
  End If

  rs.Open sSql, db, adOpenStatic, adLockOptimistic
      
  If Estado = "VACIA" Then
     If rs.EOF And rs.BOF Then
        DameEstadoMesa = True
     Else
        DameEstadoMesa = False
     End If
  Else
     If rs.EOF And rs.BOF Then
        DameEstadoMesa = False
     Else
        DameEstadoMesa = True
     End If
  End If

   db.Close

End Function

Public Function AnularMesa() As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
   
  db.Open sDb
  
  Set rs = New Recordset
  
  db.BeginTrans
  
  sSql = "DELETE  FROM auxiliar WHERE a_mesa=" & nMesa & " AND a_mensaje='ABIERTA'"
  
  db.Execute sSql
    
  db.CommitTrans
  
  db.Close
  
  Exit Function
    
errHandler:
    db.RollbackTrans
    err.Raise err.Number

End Function

Public Function DameTotalMesa(nCod As Integer) As Currency
  Dim sSql As String
  Dim rs As Recordset
    
  sDb = "FILE NAME=" & App.Path & "\Conex.udl"
  db.Open sDb
  
  Set rs = New Recordset
  
  If bDelivery = False Then
     sSql = "SELECT SUM(a_total) AS Toprecio FROM Auxiliar WHERE a_mesa=" & nCod & " AND a_mensaje='ABIERTA'"
  Else
     sSql = "SELECT SUM(a_total) AS Toprecio FROM Auxiliar WHERE a_mesa=" & nCod & " AND a_mensaje='DELIVERY'"
  End If
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  DameTotalMesa = 0
  
  If rs.RecordCount <> 0 Then
     DameTotalMesa = rs![Toprecio]
  End If
  
  rs.Close
  db.Close
  
End Function

Public Function DameTotalMoto(nCod As Integer) As Currency
  Dim sSql As String
  Dim rs As Recordset
    
  sDb = "FILE NAME=" & App.Path & "\Conex.udl"
  db.Open sDb
  
  Set rs = New Recordset
  
  sSql = "SELECT SUM(a_total) AS Toprecio FROM Auxiliar WHERE  a_mensaje='DELIVERY' AND a_mozo=" & nCod
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  DameTotalMoto = 0
  
  If rs.RecordCount <> 0 Then
     DameTotalMoto = IIf(IsNull(rs![Toprecio]), 0, rs![Toprecio])
  End If
  
  rs.Close
  db.Close
  
End Function

Public Function DameMesaImp(nCodi As Integer) As Recordset
  Dim sSql As String
  Dim rs As Recordset
  
  db.Open sDb
  
  Set rs = New Recordset

  sSql = "SELECT Auxiliar.A_CODART, Auxiliar.A_DESCRI,  Sum(Auxiliar.A_CANT) AS Cantidad, Sum(Auxiliar.A_TOTAL) AS Total,Auxiliar.a_mozo,auxiliar.a_mesa,auxiliar.a_precio,Auxiliar.a_mensaje "
  sSql = sSql & "From Auxiliar "
  sSql = sSql & "GROUP BY Auxiliar.A_CODART, Auxiliar.A_DESCRI, Auxiliar.A_MESA,Auxiliar.a_mozo,auxiliar.a_precio,Auxiliar.a_mensaje "
  sSql = sSql & "Having (((Auxiliar.A_MESA) =" & nCodi & " and Auxiliar.a_mensaje='ABIERTA'  ))  "
  sSql = sSql & "ORDER BY Auxiliar.A_CODART"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameMesaImp = rs
  db.Close
    
End Function

Public Function TraerDelivery(pMoto As Integer, pVer As Boolean) As Recordset
  Dim sSql As String
  Dim rs As Recordset
  
  db.Open sDb
  
  Set rs = New Recordset
  sSql = "SELECT Auxiliar.Id,Auxiliar.A_MESA, Auxiliar.A_HORA,Auxiliar.A_Puerto3, Auxiliar.a_observ, Clientes.RazonSocial, Clientes.Domicilio, Auxiliar.A_MENSAJE,Auxiliar.A_OBSERV "
  sSql = sSql & "FROM Auxiliar LEFT OUTER JOIN "
  sSql = sSql & "Clientes ON Auxiliar.A_PUERTO2 = Clientes.Cliente "
  sSql = sSql & "WHERE (Auxiliar.A_MENSAJE = N'DELIVERY' AND A_PUERTO='D') "
  If pVer = True Then
     sSql = sSql & "AND Auxiliar.A_Mozo=" & pMoto
  End If
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerDelivery = rs
  db.Close
  
End Function

Public Function DameMozoMesa(nCodi As Integer) As Recordset

  Dim sSql As String
  Dim rs As Recordset
  
  db.Open sDb
  
  Set rs = New Recordset


  sSql = "SELECT mesa,mozo "
  sSql = sSql & "FROM mozomesa "
  sSql = sSql & "WHERE mozo=" & nCodi & " ORDER BY mesa"
  
  rs.CursorLocation = adUseClient

  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
    
  rs.ActiveConnection = Nothing
  Set DameMozoMesa = rs
  db.Close

End Function

Private Sub Class_Initialize()
    Set mPerm = CreateObject("Scripting.Dictionary")
    mPerm.CompareMode = 1 ' TextCompare (por si usás claves string)
    Set mMap = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mPerm = Nothing
End Sub

Public Property Get Usuario() As Long
    Usuario = mUsuario
End Property

Public Sub Clear()
    mPerm.RemoveAll
End Sub

' Carga permisos desde un Recordset con columnas: Opcion, Valor
Public Sub LoadFromRecordset(ByVal rs As Object, ByVal usuarioId As Long)
    Clear
    mUsuario = usuarioId

    If rs Is Nothing Then Exit Sub
    If rs.EOF Then Exit Sub

    rs.MoveFirst
    Do While Not rs.EOF
        Dim k As Long
        Dim v As Boolean

        k = CLng(Nz(rs.Fields("Opcion").Value, 0))
        v = CBool(Nz(rs.Fields("Valor").Value, 0))

        If k <> 0 Then mPerm(CStr(k)) = v

        rs.MoveNext
    Loop
End Sub

Public Function HasAny(ParamArray opciones() As Variant) As Boolean
    Dim i As Long
    For i = LBound(opciones) To UBound(opciones)
        If Has(CLng(opciones(i))) Then
            HasAny = True
            Exit Function
        End If
    Next
    HasAny = False
End Function


' Consultar permiso: si no existe, False
Public Function Has(ByVal opcion As Long) As Boolean
    Dim key As String
    key = CStr(opcion)

    If mPerm.Exists(key) Then
        Has = CBool(mPerm(key))
    Else
        Has = False
    End If
End Function


' ===== helpers =====
Private Function Nz(ByVal v As Variant, ByVal defaultValue As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

Public Sub SetComboByItemData(ByRef cbo As ComboBox, ByVal valueId As Long)
    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = valueId Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next
    cbo.ListIndex = -1
End Sub

Public Sub BuildComboIndex(ByVal mapName As String, ByVal cbo As ComboBox)
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")

    For i = 0 To cbo.ListCount - 1
        d(CStr(cbo.ItemData(i))) = CLng(i)
    Next

    If mMap.Exists(mapName) Then
        mMap.Remove mapName
    End If

    mMap.Add mapName, d
End Sub

Public Sub SetComboByIdFast(ByVal mapName As String, ByVal cbo As ComboBox, ByVal id As Variant)
    If IsNull(id) Or IsEmpty(id) Then
        cbo.ListIndex = -1
        Exit Sub
    End If

    If Not mMap.Exists(mapName) Then Exit Sub

    Dim d As Object
    Set d = mMap(mapName)

    Dim k As String
    k = CStr(CLng(id))

    If d.Exists(k) Then
        cbo.ListIndex = CLng(d(k))
    Else
        cbo.ListIndex = -1
    End If
End Sub

