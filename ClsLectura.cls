VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mConn As ADODB.Connection
Private mConnStr As String

Private Sub Class_Initialize()
    mConnStr = "FILE NAME=" & App.Path & "\Conex.udl"
    Set mConn = New ADODB.Connection
    mConn.ConnectionTimeout = 5
    mConn.CommandTimeout = 30
    mConn.Open mConnStr
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not mConn Is Nothing Then
        If mConn.State <> adStateClosed Then mConn.Close
    End If
    Set mConn = Nothing
End Sub

Public Sub EnsureConn()
    If mConn Is Nothing Then
        Set mConn = New ADODB.Connection
        sDb = "FILE NAME=" & App.Path & "\Conex.udl"
        mConn.Open sDb
    ElseIf mConn.State = adStateClosed Then
        sDb = "FILE NAME=" & App.Path & "\Conex.udl"
        mConn.Open sDb
    End If
End Sub

Private Function DetachRS(ByVal rs As ADODB.Recordset) As ADODB.Recordset
    rs.ActiveConnection = Nothing   ' <-- SIN Set
    Set DetachRS = rs
End Function

Public Sub GrabarError(pModulo As String, pError As String, pOtros As String)
  Dim intFile As Integer
  intFile = 1

  Open App.Path & "\Errores.log" For Append As #intFile
  Print #intFile, "-----------------------------------------------------------------------------------"
  Print #intFile, "Fecha :" & Date
  Print #intFile, "Hora :" & Time
  Print #intFile, "Modulo :" & pModulo
  Print #intFile, "Error :" & pError
  Print #intFile, "Mensaje :" & pOtros
  Print #intFile, "-----------------------------------------------------------------------------------"
  Print #intFile, ""
  Print #intFile, ""
  Close #intFile
End Sub

Public Function TraerRS(ByVal sp As String, Optional ByVal pCodigo As Variant, Optional ByVal pCons As Boolean = False) As ADODB.Recordset
    On Error GoTo errHandler
    EnsureConn

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = mConn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = sp

    ' Solo agrega parámetro si realmente lo mandaron y no es Empty/Null
    If Not IsMissing(pCodigo) Then
        If Not IsEmpty(pCodigo) Then
            If Not IsNull(pCodigo) Then
                If VarType(pCodigo) = vbString Then
                    cmd.Parameters.Append cmd.CreateParameter("Param", adVarChar, adParamInput, Len(CStr(pCodigo)), CStr(pCodigo))
                Else
                    cmd.Parameters.Append cmd.CreateParameter("Param", adInteger, adParamInput, , CLng(pCodigo))
                End If
            End If
        End If
    End If

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    If pCons Then
        ' CLAVE: setear cursores ANTES del Open
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.LockType = adLockReadOnly

        rs.Open cmd
        Set rs.ActiveConnection = Nothing   ' queda desconectado, pero ABIERTO
    Else
        rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    End If

    Set TraerRS = rs
    Exit Function

errHandler:
    err.Raise err.Number, "TraerRS(" & sp & ")", err.Description
End Function



Public Function TraerRSVen(pStoredProc As String, pCodigo As Variant, Optional pCons As Boolean) As ADODB.Recordset
    ' Con codigo de tipo numerico (casi todas las tablas van a tener un codigo Identity)
    ' sino hay que hacer otra rutina
    
    On Error GoTo errHandler
    Dim Com As ADODB.Command
    Dim RsTrae As ADODB.Recordset
    Set Com = New ADODB.Command
    Set RsTrae = New ADODB.Recordset
    
    db.Open sDb
    Com.ActiveConnection = db
    Com.CommandType = adCmdStoredProc
    Com.CommandText = pStoredProc
    ' arma el parametro
       ' necesita parametro
       If VarType(pCodigo) <> vbString Then
          Com.Parameters.Append Com.CreateParameter("Param", adInteger, adParamInput, , pCodigo)
       Else
          Com.Parameters.Append Com.CreateParameter("Param", adVarChar, adParamInput, Len(pCodigo), pCodigo)
       End If
    Set RsTrae = Com.Execute
    ' lo convierto en un RS desconectado
    If pCons = True Then
       RsTrae.Close
       RsTrae.CursorType = adOpenStatic
       RsTrae.CursorLocation = adUseClient
       RsTrae.LockType = adLockBatchOptimistic
       RsTrae.Open
       RsTrae.ActiveConnection = Nothing
    End If
    Set TraerRSVen = RsTrae
    
    db.Close
    Exit Function
    
errHandler:
    err.Raise err.Number
End Function

Public Function RsVacio(ByVal pTabla As String, ByVal pCampo As String, ByVal pTipo As String) As ADODB.Recordset
    On Error GoTo errHandler
    EnsureConn

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim sSql As String
    Select Case UCase$(pTipo)
        Case "N"
            sSql = "SELECT * FROM " & pTabla & " WHERE " & pCampo & "=0"
        Case "S"
            sSql = "SELECT * FROM " & pTabla & " WHERE " & pCampo & "=''"
        Case Else
            err.Raise 5, "RsVacio", "Tipo inválido: " & pTipo
    End Select
    rs.CursorLocation = adUseClient
    rs.Open sSql, mConn, adOpenStatic, adLockOptimistic
    rs.ActiveConnection = Nothing
      
   ' Set rs = DetachRSAsStatic(rs)   ' <- clave: re-abrir como navegable
   
    Set RsVacio = rs
    Exit Function

errHandler:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    err.Raise err.Number, "RsVacio(" & pTabla & ")", err.Description
End Function

Private Function DetachRSAsStatic(ByVal rsSrc As ADODB.Recordset) As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    Set rsDest = New ADODB.Recordset

    rsDest.CursorLocation = adUseClient
    rsDest.CursorType = adOpenStatic
    rsDest.LockType = adLockOptimistic

    ' Copia los datos del rsSrc al rsDest (queda desconectado y navegable)
    rsDest.Open rsSrc

    ' Cerrar el origen para liberar server cursor
    On Error Resume Next
    rsSrc.Close

    Set DetachRSAsStatic = rsDest
End Function

Option Explicit

'=========================================================
' TraerRsSQL: ejecuta un SELECT y devuelve Recordset listo
' para bindear a un DataGrid / DBGrid (client-side).
'=========================================================
Public Function TraerRsSQL(ByVal sSql As String, Optional ByVal nTimeout As Long = 30) As ADODB.Recordset
    On Error GoTo errHandler

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connStr As String

    '--- Cadena de conexión (UDL)
    '   Si ya usás sDb global, lo respetamos.
    If Len(Trim$(sDb & "")) > 0 Then
        connStr = sDb
    Else
        connStr = "FILE NAME=" & App.Path & "\Conex.udl"
    End If

    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 5
    cn.CommandTimeout = nTimeout
    cn.Open connStr

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic       'para RecordCount y navegación
    rs.LockType = adLockReadOnly

    rs.Open sSql, cn

    'Desconectado => el grid sigue funcionando y no queda la conexión colgada
    Set rs.ActiveConnection = Nothing
    cn.Close
    Set cn = Nothing

    Set TraerRsSQL = rs
    Exit Function

errHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set rs = Nothing
    Set cn = Nothing
    err.Raise err.Number, "ClsLectura.TraerRsSQL", err.Description
End Function


Public Function TraerRsSQLCmd(ByVal sSql As String, ByRef params() As Variant, Optional ByVal nTimeout As Long = 30) As ADODB.Recordset
    On Error GoTo errHandler

    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim connStr As String
    Dim i As Long

    If Len(Trim$(sDb & "")) > 0 Then
        connStr = sDb
    Else
        connStr = "FILE NAME=" & App.Path & "\Conex.udl"
    End If

    Set cn = New ADODB.Connection
    cn.ConnectionTimeout = 5
    cn.CommandTimeout = nTimeout
    cn.Open connStr

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sSql

    'params: array de arrays => Array(Array("@p1", adInteger, 123), Array("@p2", adVarChar, "hola"))
    For i = LBound(params) To UBound(params)
        cmd.Parameters.Append cmd.CreateParameter(params(i)(0), params(i)(1), adParamInput, , params(i)(2))
    Next i

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly

    rs.Open cmd

    Set rs.ActiveConnection = Nothing
    cn.Close
    Set cn = Nothing

    Set TraerRsSQLCmd = rs
    Exit Function

errHandler:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
    err.Raise err.Number, "ClsLectura.TraerRsSQLCmd", err.Description
End Function

Public Function RsVacioD(pTabla As String, pCampo As String, pTipo As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  Select Case pTipo
    Case "N"
       sSql = "SELECT * FROM " & pTabla & " WHERE " & pCampo & "=0"
    Case "S"
       sSql = "SELECT * FROM " & pTabla & " WHERE " & pCampo & "=''"
  End Select
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  Set RsVacioD = rs
  
  db.Close
Exit Function
errHandler:
  db.Close
  err.Raise err.Number
End Function

Public Sub CargaComboComprobantes(P_Combo As Object, P_Tabla As String, P_Codigo As String, P_Campo As String, P_Condicion As String)
   Dim sSql As String
   Dim iCantidad As Integer
   Dim Conexion As ADODB.Connection
   Dim Rs_Combo As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Combo = New ADODB.Recordset
   
   Conexion.Open sDb

   If P_Condicion = "" Then
      sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc FROM " & P_Tabla & " ORDER BY " & P_Campo
   Else
      sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc FROM " & P_Tabla & " WHERE " & P_Condicion & " ORDER BY " & P_Campo
   End If
 
   P_Combo.Clear
 
   Rs_Combo.Open sSql, Conexion, adOpenKeyset, adLockPessimistic
   iCantidad = Rs_Combo.RecordCount
   
   Do While Not Rs_Combo.EOF
      P_Combo.AddItem Trim(Rs_Combo!descripc)
      P_Combo.ItemData(P_Combo.NewIndex) = Rs_Combo!codigo
      Rs_Combo.MoveNext
   Loop
   Rs_Combo.Close
   Set Rs_Combo = Nothing
   Conexion.Close
   Exit Sub
   
errHandler:
   err.Raise err.Number
End Sub

Public Sub CargaCombo(P_Combo As Object, P_Tabla As String, P_Codigo As String, P_Campo As String, P_Condicion As String, Optional pOrden As String = "")
    Dim sSql As String
    Dim Rs_Combo As ADODB.Recordset
    On Error GoTo errHandler

    Set Rs_Combo = New ADODB.Recordset

    If P_Condicion = "" Then
        sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc FROM " & P_Tabla
    Else
        sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc FROM " & P_Tabla & " WHERE " & P_Condicion
    End If

    If Len(pOrden) > 0 Then
        sSql = sSql & " ORDER BY " & pOrden
    ElseIf P_Condicion = "" Then
        sSql = sSql & " ORDER BY " & P_Campo
    End If

    P_Combo.Clear

    ' Cursor liviano y solo lectura (MUCHO más rápido)
    Rs_Combo.CursorLocation = adUseServer
    Rs_Combo.Open sSql, mConn, adOpenForwardOnly, adLockReadOnly

    Do While Not Rs_Combo.EOF
        P_Combo.AddItem Trim$(CStr(Rs_Combo!descripc))
        P_Combo.ItemData(P_Combo.NewIndex) = CLng(Rs_Combo!codigo)
        Rs_Combo.MoveNext
    Loop

    Rs_Combo.Close
    Set Rs_Combo = Nothing
    Exit Sub

errHandler:
    On Error Resume Next
    If Not Rs_Combo Is Nothing Then
        If Rs_Combo.State <> adStateClosed Then Rs_Combo.Close
    End If
    Set Rs_Combo = Nothing
    MsgBox err.Description
End Sub


Public Sub CargaComboIn( _
    P_Combo As Object, _
    P_Tabla As String, _
    P_Codigo As String, _
    P_Campo As String, _
    pCampoIn As String, _
    pListaComas As String, _
    Optional pCondicionExtra As String = "", _
    Optional pOrden As String = "")

    Dim sSql As String
    Dim Rs_Combo As ADODB.Recordset
    Dim lista As String

    On Error GoTo errHandler

    Set Rs_Combo = New ADODB.Recordset

    lista = Trim$(pListaComas)

    sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc FROM " & P_Tabla & " WHERE 1=1 "

    ' IN (1,5,7)
    If Len(lista) > 0 Then
        sSql = sSql & " AND " & pCampoIn & " IN (" & lista & ") "
    End If

    ' Condición extra opcional (por ej: "Activo=1")
    If Len(Trim$(pCondicionExtra)) > 0 Then
        sSql = sSql & " AND (" & pCondicionExtra & ") "
    End If

    If Len(Trim$(pOrden)) > 0 Then
        sSql = sSql & " ORDER BY " & pOrden
    Else
        sSql = sSql & " ORDER BY " & P_Campo
    End If

    P_Combo.Clear

    Rs_Combo.CursorLocation = adUseServer
    Rs_Combo.Open sSql, mConn, adOpenForwardOnly, adLockReadOnly

    Do While Not Rs_Combo.EOF
        P_Combo.AddItem Trim$(CStr(Rs_Combo!descripc))
        P_Combo.ItemData(P_Combo.NewIndex) = CLng(Rs_Combo!codigo)
        Rs_Combo.MoveNext
    Loop

    Rs_Combo.Close
    Set Rs_Combo = Nothing
    Exit Sub

errHandler:
    On Error Resume Next
    If Not Rs_Combo Is Nothing Then
        If Rs_Combo.State <> adStateClosed Then Rs_Combo.Close
    End If
    Set Rs_Combo = Nothing
    MsgBox err.Description, vbCritical, "CargaComboIn"
End Sub

 
 Public Sub CargaComboEmpresa(P_Combo As Object, P_Tabla As String, P_Codigo As String, P_Campo As String, P_Condicion As String)
   Dim sSql As String
   Dim iCantidad As Integer
   Dim Conexion As ADODB.Connection
   Dim Rs_Combo As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Combo = New ADODB.Recordset
   
   Conexion.Open sUdl

   If P_Condicion = "" Then
      sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc FROM " & P_Tabla & " ORDER BY " & P_Campo
   Else
      sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc FROM " & P_Tabla & " WHERE " & P_Condicion & " ORDER BY " & P_Campo
   End If
 
   P_Combo.Clear
 
   Rs_Combo.Open sSql, Conexion, adOpenKeyset, adLockPessimistic
   iCantidad = Rs_Combo.RecordCount
   
   Do While Not Rs_Combo.EOF
      P_Combo.AddItem Trim(Rs_Combo!descripc)
      P_Combo.ItemData(P_Combo.NewIndex) = Rs_Combo!codigo
      Rs_Combo.MoveNext
   Loop
   Rs_Combo.Close
   Set Rs_Combo = Nothing
   Conexion.Close
   Exit Sub
   
errHandler:
   err.Raise err.Number
End Sub
 
Public Sub CargaComboLst(P_Combo As Object, P_Tabla As String, P_Codigo As String, P_Campo As String, p_Valor As String, P_Condicion As String, Optional pOrden As String = "")
   Dim sSql As String
   Dim iCantidad As Integer
   Dim Conexion As ADODB.Connection
   Dim Rs_Combo As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Combo = New ADODB.Recordset
   
   Conexion.Open sDb

   If P_Condicion = "" Then
      sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc, " & p_Valor & " as Valor  FROM " & P_Tabla & " ORDER BY " & P_Campo
   Else
      sSql = "SELECT " & P_Codigo & " as codigo, " & P_Campo & " as descripc, " & p_Valor & " as Valor  FROM " & P_Tabla & " WHERE " & P_Condicion
   End If
   If pOrden <> "" Then
      sSql = sSql & " ORDER BY " & pOrden
   Else
    sSql = sSql & " ORDER BY " & P_Campo
   End If
 
   P_Combo.Clear
 
   Rs_Combo.Open sSql, Conexion, adOpenKeyset, adLockPessimistic
   iCantidad = Rs_Combo.RecordCount
   
   Do While Not Rs_Combo.EOF
      P_Combo.AddItem Trim(Rs_Combo!descripc)
      P_Combo.ItemData(P_Combo.NewIndex) = Rs_Combo!codigo
      P_Combo.Selected(P_Combo.NewIndex) = Rs_Combo!valor
      Rs_Combo.MoveNext
   Loop
   Rs_Combo.Close
   Set Rs_Combo = Nothing
   Conexion.Close
   Exit Sub
   
errHandler:
   err.Raise err.Number
End Sub


Public Function DatoCombo(P_Tabla As String, P_Codigo As String, P_Campo As String, iCodi As Long) As String
    On Error GoTo errHandler
    EnsureConn

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim sSql As String
    sSql = "SELECT " & P_Campo & " as descripc FROM " & P_Tabla & " WHERE " & P_Codigo & "=" & iCodi

    rs.Open sSql, mConn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        DatoCombo = Trim$(CStr(rs!descripc))
    Else
        DatoCombo = vbNullString
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

errHandler:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    err.Raise err.Number, "DatoCombo(" & P_Tabla & ")", err.Description
End Function



Public Function TraerTodos(pTabla As String, pCampo As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  If pCampo <> "*" Then
     sSql = "SELECT " & pCampo & " FROM " & pTabla & " ORDER BY " & pCampo
  Else
     sSql = "SELECT " & pCampo & " FROM " & pTabla
  End If
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerTodos = rs
  
  db.Close
  Exit Function
errHandler:
  db.Close
  err.Raise err.Number
End Function

Public Function TraerRsCondi(pTabla As String, pCampo As String, pCondi As String) As ADODB.Recordset
    On Error GoTo errHandler
    EnsureConn

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim sSql As String
    sSql = "SELECT * FROM " & pTabla & " WHERE " & pCondi
    If Len(pCampo) > 0 Then sSql = sSql & " ORDER BY " & pCampo

    ' IMPORTANTE: esto va ANTES de abrir
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic

    rs.Open sSql, mConn   ' ya queda “navegable”
    rs.ActiveConnection = Nothing  ' desconectado

    Set TraerRsCondi = rs
    Exit Function

errHandler:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    err.Raise err.Number, "TraerRsCondi(" & pTabla & ")", err.Description
End Function



Public Function SumarCantidad(pTabla As String, pCampo As String, pCondi As String) As Integer
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb

  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  If pCondi <> "" Then
     sSql = "SELECT Sum(" & pCampo & ") as Total FROM " & pTabla & " WHERE " & pCondi
  Else
     sSql = "SELECT Sum(" & pCampo & ") as Total FROM " & pTabla
  End If
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  SumarCantidad = IIf(IsNull(rs!Total), 0, rs!Total)
  
  db.Close
  
  Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerCantidad(pTabla As String, pCampo As String, pCondi As String) As Integer
    On Error GoTo errHandler
    EnsureConn

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim sSql As String
    If Len(pCondi) > 0 Then
        sSql = "SELECT Count(" & pCampo & ") as Total FROM " & pTabla & " WHERE " & pCondi
    Else
        sSql = "SELECT Count(" & pCampo & ") as Total FROM " & pTabla
    End If

    rs.Open sSql, mConn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF And Not IsNull(rs!Total) Then
        TraerCantidad = CInt(rs!Total)
    Else
        TraerCantidad = 0
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

errHandler:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    err.Raise err.Number, "TraerCantidad(" & pTabla & ")", err.Description
End Function


Public Function TraerCantidadSin(pTabla As String, pCampo As String, pCondi As String) As Integer
    ' Compatibilidad: ahora hace lo mismo que TraerCantidad, pero sin tocar conexiones externas
    TraerCantidadSin = TraerCantidad(pTabla, pCampo, pCondi)
End Function


Public Function Buscar(pTabla As String, P_Descri As String, P_Codigo As String, sDato As String, pBus As Boolean) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb

  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  If pBus = True Then
     sSql = "SELECT TOP 100 " & P_Descri & " as Descripcion," & P_Codigo & " As codigo FROM " & pTabla & " WHERE " & P_Descri & " LIKE '%" & sDato & "%' order by " & P_Descri
  Else
     sSql = "SELECT TOP 100 " & P_Descri & " as Descripcion," & P_Codigo & " As codigo FROM " & pTabla & " WHERE " & P_Descri & ">= '" & sDato & "' order by " & P_Descri
  End If
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set Buscar = rs
  
  db.Close

Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function BuscarTodos(pTabla As String, P_Descri As String, P_Codigo As String) As ADODB.Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT TOP 100 " & P_Descri & " as Descripcion," & P_Codigo & " AS codigo FROM " & pTabla & " ORDER BY " & P_Descri
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set BuscarTodos = rs
  
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function
Public Function BuscarNombreProducto(sCod As String) As String
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb

  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT Descripcion FROM Productos WHERE Producto='" & sCod & "'"
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     BuscarNombreProducto = rs!Descripcion
  End If
  rs.Close
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function VerificarDeposito(nDep As Integer) As Boolean
  Dim rs As ADODB.Recordset, sSql As String
  
  db.Open sDb

  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT Sum(Venta) as Cant FROM Depositos WHERE Deposito<>" & nDep
  rs.Open sSql, db, adOpenKeyset, adLockPessimistic
  If rs.EOF = False Or rs.BOF = False Then
     If rs!cant = 1 Then VerificarDeposito = True
  End If
  
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function ConsultasDesdeHasta(pStoredProc As String, pDesde As String, pHasta As String, pCons As Boolean) As Recordset
    On Error GoTo errHandler
    Dim Com As ADODB.Command
    Dim RsTrae As ADODB.Recordset
    Set Com = New ADODB.Command
    Set RsTrae = New ADODB.Recordset
    
    db.Open sDb
  
    Com.ActiveConnection = db
    Com.CommandType = adCmdStoredProc
    Com.CommandText = pStoredProc
    ' arma el parametro
    Com.Parameters.Append Com.CreateParameter("Desde", adVarChar, adParamInput, 45, pDesde)
    Com.Parameters.Append Com.CreateParameter("Hasta", adVarChar, adParamInput, 45, pHasta)
    Set RsTrae = Com.Execute
    ' lo convierto en un RS desconectado
    If pCons = True Then
       RsTrae.Close
       RsTrae.CursorType = adOpenStatic
       RsTrae.CursorLocation = adUseClient
       RsTrae.LockType = adLockBatchOptimistic
       RsTrae.Open
       RsTrae.ActiveConnection = Nothing
    End If
    Set ConsultasDesdeHasta = RsTrae
    Exit Function
    db.Close
    
errHandler:
    err.Raise err.Number
  
End Function

Public Function TraerDefectos(pUser As Integer) As Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb

  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT * FROM Registros WHERE Usuario=" & pUser
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set TraerDefectos = rs
  
  db.Close
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function RegPorDefecto(P_Campo As String) As Integer
   Dim sSql As String
   Dim Conexion As ADODB.Connection
   Dim Rs_Text As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Text = New ADODB.Recordset
   
   If db.State = 0 Then
      db.Open sDb
   End If

   sSql = "SELECT " & P_Campo & " as descripc FROM Registros WHERE Usuario=" & nUsuario & " And Suc = " & nSucursal
 
   Rs_Text.Open sSql, db, adOpenKeyset, adLockPessimistic
   If Rs_Text.RecordCount <> 0 Then
      RegPorDefecto = Rs_Text!descripc
   End If
   Rs_Text.Close
   Set Rs_Text = Nothing
   If db.State <> 0 Then
      db.Close
   End If
   
   Exit Function
errHandler:
   err.Raise err.Number
  
End Function

Public Function SistemaPorDefecto(P_Campo As String) As String
   Dim sSql As String
   Dim Conexion As ADODB.Connection
   Dim Rs_Text As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Text = New ADODB.Recordset
   
   If db.State = 0 Then
      db.Open sDb
   End If

   sSql = "SELECT " & P_Campo & " as descripc FROM Sistema "
 
   Rs_Text.Open sSql, db, adOpenKeyset, adLockPessimistic
   If Rs_Text.RecordCount <> 0 Then
      SistemaPorDefecto = Rs_Text!descripc
   End If
   Rs_Text.Close
   Set Rs_Text = Nothing
   If db.State <> 0 Then
      db.Close
   End If
   
   Exit Function
errHandler:
   err.Raise err.Number
  
End Function

Public Function ClientePorDefecto(pCliente As Integer) As String
   Dim sSql As String
   Dim Conexion As ADODB.Connection
   Dim Rs_Text As ADODB.Recordset
 
   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Text = New ADODB.Recordset
   
   If db.State = 1 Then
      db.Open sDb
   End If
   sSql = "SELECT Cliente FROM Clientes WHERE Id=" & pCliente
 
   Rs_Text.Open sSql, db, adOpenKeyset, adLockPessimistic
   If Rs_Text.RecordCount <> 0 Then
      ClientePorDefecto = Rs_Text!Cliente
   End If
   Rs_Text.Close
   Set Rs_Text = Nothing
   
   If db.State <> 1 Then
      db.Close
   End If
   Exit Function
errHandler:
   err.Raise err.Number

End Function

Public Function ValidarUsuario(pUsuario As String, pPassw As String) As Boolean
   Dim sSql As String
   Dim Conexion As ADODB.Connection
   Dim Rs_Text As ADODB.Recordset
 
   db.Open sDb

   On Error GoTo errHandler
   Set Conexion = New ADODB.Connection
   Set Rs_Text = New ADODB.Recordset
   
   sSql = "SELECT * FROM Usuarios WHERE Usuario='" & pUsuario & "' and Contraseña='" & pPassw & "'"
 
   Rs_Text.Open sSql, db, adOpenKeyset, adLockPessimistic
   If Rs_Text.RecordCount <> 0 Then
      ValidarUsuario = True
   End If
   Rs_Text.Close
   Set Rs_Text = Nothing
   
   db.Close
   Exit Function
errHandler:
   err.Raise err.Number

End Function

Public Function DamePermisos(pUsuario As String, pPassw As String) As Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  sSql = "SELECT * FROM Usuarios WHERE Usuario='" & pUsuario & "' and Contraseña='" & pPassw & "'"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DamePermisos = rs
  
  db.Close
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function PuedoHacer(pCampo As String, pOperacion As Byte) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  db.Open sDb
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  sSql = "SELECT " & pCampo & " as Descrip FROM Permisos WHERE  Usuario=" & rUsu!ID & " and Operacion=" & pOperacion
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  PuedoHacer = rs!Descrip
  
  db.Close
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function PuedoCompro(pCampo As String, pOperacion As Byte, pComprobante As Integer) As Boolean
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  db.Open sDb

  sSql = "SELECT Comprobante FROM Permisos WHERE  Usuario=" & rUsu!ID & " and Operacion=" & pOperacion
  sSql = sSql & " and Comprobante=" & pComprobante
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  If rs.RecordCount <> 0 Then
     PuedoCompro = True
  End If
  
  db.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function


Public Function TraerValorDeUnCampo(pTabla As String, pDescri As String, pCondi As String) As Variant
    On Error GoTo errHandler
    EnsureConn

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim sSql As String
    If Len(pCondi) > 0 Then
        sSql = "SELECT " & pDescri & " as Descripcion FROM " & pTabla & " WHERE " & pCondi
    Else
        sSql = "SELECT " & pDescri & " as Descripcion FROM " & pTabla
    End If

    rs.Open sSql, mConn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        If IsNull(rs!Descripcion) Then
            TraerValorDeUnCampo = ""
        Else
            TraerValorDeUnCampo = rs!Descripcion
        End If
    Else
        TraerValorDeUnCampo = 0
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

errHandler:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    err.Raise err.Number, "TraerValorDeUnCampo(" & pTabla & ")", err.Description
End Function


Public Function TraerValorDeUnCampoSin(pTabla As String, pDescri As String, pCondi As String) As Variant
  Dim sSql As String
  Dim rs As ADODB.Recordset
  
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  If pCondi <> "" Then
     sSql = "SELECT " & pDescri & " as Descripcion FROM " & pTabla & " WHERE  " & pCondi
  Else
     sSql = "SELECT " & pDescri & " as Descripcion FROM " & pTabla
  End If
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  If rs.RecordCount <> 0 Then
     TraerValorDeUnCampoSin = IIf(IsNull(rs!Descripcion), "", Trim(rs!Descripcion))
  Else
     TraerValorDeUnCampoSin = 0
  End If
 
Exit Function
errHandler:
  err.Raise err.Number

End Function
Public Function TraerMovimiento() As Integer
   Dim Conexion As ADODB.Connection, sSql As String
   Dim rs As ADODB.Recordset
 
   Dim Com As ADODB.Command
   Set Com = New ADODB.Command
   Set Conexion = New ADODB.Connection
   
   Conexion.Open "FILE NAME=" & App.Path & "\Conex.udl"

   sSql = "SELECT Movimiento FROM Sistema "
   
   rs.Open sSql, db, adOpenStatic, adLockOptimistic
   rs.ActiveConnection = Nothing
   TraerMovimiento = rs!Movimiento
   
   Conexion.Close
End Function

Public Function TraerComproPago(pVenta As Long) As Recordset
  Dim sSql As String
  Dim rs As ADODB.Recordset

  Set rs = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT CabComprobantes.Fecha, Comprobantes.Comprobante, CabComprobantes.Sucursal, CabComprobantes.Numero "
  sSql = sSql & "FROM (Aplicaciones LEFT JOIN CabComprobantes ON Aplicaciones.Venta = CabComprobantes.Movimiento) LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE Aplicaciones.Venta=" & pVenta
  rs.CursorLocation = adUseClient
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  
  Set TraerComproPago = rs
  
  db.Close
  
End Function

Public Function TraerCantidadCuotas(pVenta As Long) As Byte
  Dim sSql As String
  Dim rs As ADODB.Recordset

  Set rs = New ADODB.Recordset

  db.Open sDb

   Set rs = New ADODB.Recordset

   sSql = "SELECT Count(CuentaCorriente.Cuota) AS CuentaDeCuota "
   sSql = sSql & "FROM CabComprobantes LEFT JOIN CuentaCorriente ON CabComprobantes.Movimiento = CuentaCorriente.Venta "
   sSql = sSql & "WHERE (((CabComprobantes.Movimiento)=" & pVenta & ")) "
   rs.CursorLocation = adUseClient
   rs.Open sSql, db, adOpenStatic, adLockOptimistic
   rs.ActiveConnection = Nothing
      
   TraerCantidadCuotas = rs!CuentaDeCuota
   
   db.Close

End Function
Public Function BorrarRsD(pRsd As Recordset) As Boolean
   If pRsd.RecordCount <> 0 Then
      pRsd.MoveLast
      Do While Not pRsd.EOF
         pRsd.Delete
         If pRsd.RecordCount <> 0 Then
            pRsd.MovePrevious
            If pRsd.BOF Then
               pRsd.MoveFirst
            End If
         Else
            Exit Do
         End If
      Loop
   End If
End Function

Public Function TraerSaldos(pCliente As Long) As Currency
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset

  db.Open sDb

  sSql = "SELECT (Sum(CuentaCorriente.Debe) - Sum(CuentaCorriente.Haber)) AS Saldo "
  sSql = sSql & "FROM Clientes LEFT JOIN CuentaCorriente ON Clientes.Cliente = CuentaCorriente.Cliente "
  sSql = sSql & "WHERE Clientes.Id=" & pCliente & " "
  sSql = sSql & "GROUP BY Clientes.Id, Clientes.RazonSocial; "
  
  Rs_Chk.Open sSql, db, adOpenStatic
  
  If Rs_Chk.RecordCount <> 0 Then
     TraerSaldos = IIf(IsNull(Rs_Chk!saldo), 0, Rs_Chk!saldo)
  End If
  Set Rs_Chk = Nothing
  db.Close
  
End Function

Public Function TraerProductoPrecio(pProducto As String, pLista As Integer) As Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Dim Rs_Chk As ADODB.Recordset
 
  Set Conexion = New ADODB.Connection
  Set Rs_Chk = New ADODB.Recordset

  db.Open sDb

  If bModoSuper = True Then
     sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.Calcular, Productos.CodCompu, Productos.Fijo, Productos.TipoInterno, Productos.ValorInterno,"
     sSql = sSql & "Productos.ActuaStock , PreciosCaja.precio, Impuestos.Porcentaje as Iva,Productos.ValorEnvase "
     sSql = sSql & "FROM Productos INNER JOIN "
     sSql = sSql & "PreciosCaja ON Productos.Producto = PreciosCaja.Producto LEFT OUTER JOIN "
     sSql = sSql & "ProductoImpuesto ON Productos.Producto = ProductoImpuesto.Producto LEFT OUTER JOIN "
     sSql = sSql & "Impuestos ON ProductoImpuesto.Impuesto = Impuestos.Impuesto "
     sSql = sSql & "WHERE Productos.Producto='" & pProducto & "'"
  Else
     sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.Calcular, Productos.CodCompu, Productos.Fijo, Productos.TipoInterno, Productos.ValorInterno,"
     sSql = sSql & "Productos.ActuaStock , Precios.Precio, Impuestos.Porcentaje as Iva,Productos.ValorEnvase "
     sSql = sSql & "FROM Productos INNER JOIN "
     sSql = sSql & "Precios ON Productos.Producto = Precios.Producto LEFT OUTER JOIN "
     sSql = sSql & "ProductoImpuesto ON Productos.Producto = ProductoImpuesto.Producto LEFT OUTER JOIN "
     sSql = sSql & "Impuestos ON ProductoImpuesto.Impuesto = Impuestos.Impuesto "
     sSql = sSql & "WHERE Productos.Producto='" & pProducto & "' AND Precios.Lista=" & pLista
  
  End If
  
  Rs_Chk.CursorLocation = adUseClient
  Rs_Chk.Open sSql, db, adOpenStatic, adLockOptimistic
  Rs_Chk.ActiveConnection = Nothing
  
  If Rs_Chk.RecordCount = 0 Then
     If bModoSuper = True Then
       sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.Calcular, Productos.CodCompu, Productos.Fijo, Productos.TipoInterno, Productos.ValorInterno,"
       sSql = sSql & "Productos.ActuaStock , PreciosCaja.precio, Impuestos.Porcentaje as Iva,Productos.ValorEnvase "
       sSql = sSql & "FROM Productos INNER JOIN "
       sSql = sSql & "PreciosCaja ON Productos.Producto = PreciosCaja.Producto LEFT OUTER JOIN "
       sSql = sSql & "ProductoImpuesto ON Productos.Producto = ProductoImpuesto.Producto LEFT OUTER JOIN "
       sSql = sSql & "Impuestos ON ProductoImpuesto.Impuesto = Impuestos.Impuesto "
       sSql = sSql & "WHERE Productos.CBarra='" & pProducto & "'"
    Else
       sSql = "SELECT Productos.Producto, Productos.Descripcion, Productos.Calcular, Productos.CodCompu, Productos.Fijo, Productos.TipoInterno, Productos.ValorInterno,"
       sSql = sSql & "Productos.ActuaStock , Precios.Precio, Impuestos.Porcentaje as Iva,Productos.ValorEnvase "
       sSql = sSql & "FROM Productos INNER JOIN "
       sSql = sSql & "Precios ON Productos.Producto = Precios.Producto LEFT OUTER JOIN "
       sSql = sSql & "ProductoImpuesto ON Productos.Producto = ProductoImpuesto.Producto LEFT OUTER JOIN "
       sSql = sSql & "Impuestos ON ProductoImpuesto.Impuesto = Impuestos.Impuesto "
       sSql = sSql & "WHERE Productos.CBarra='" & pProducto & "' AND Precios.Lista=" & pLista
    End If
    
    Rs_Chk.Close
    Rs_Chk.CursorLocation = adUseClient
    Rs_Chk.Open sSql, db, adOpenStatic, adLockOptimistic
    Rs_Chk.ActiveConnection = Nothing
     
  End If
  
  Set TraerProductoPrecio = Rs_Chk
  
  db.Close
  
End Function

Public Function ActualizarMovimientoTraer() As Long
  Dim db As Connection, rs As Recordset
  
  Set db = New Connection
  Set rs = New Recordset
  On Error GoTo errHandler
  Dim sSql As String
  
  db.Open sDb
  db.BeginTrans
  
  sSql = "UPDATE Sistema SET Movimiento=Movimiento+1"
  db.Execute sSql
  
  db.CommitTrans
  
  sSql = "SELECT Movimiento FROM Sistema"
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  ActualizarMovimientoTraer = rs!Movimiento
  db.Close
  
 Exit Function
errHandler:
    db.RollbackTrans
    db.Close
    MsgBox err.Description, vbCritical, "ActualizarMovimiento"

End Function

Public Function ActualizarDeliveryTraer() As Long
  Dim db As Connection, rs As Recordset
  
  Set db = New Connection
  Set rs = New Recordset
  On Error GoTo errHandler
  Dim sSql As String
  
  db.Open sDb
  db.BeginTrans
  
  sSql = "UPDATE Sistema SET Retencion=Retencion+1"
  db.Execute sSql
  
  db.CommitTrans
  
  sSql = "SELECT Retencion FROM Sistema"
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  ActualizarDeliveryTraer = rs!Retencion
  db.Close
  
 Exit Function
errHandler:
    db.RollbackTrans
    db.Close
    MsgBox err.Description, vbCritical, "ActualizarMovimiento"

End Function

Public Function TraerTasaProducto(pProducto As String) As Single
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Set Conexion = New ADODB.Connection

  Conexion.Open sDb
  
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  

  sSql = "SELECT ProductoImpuesto.Producto, Impuestos.Porcentaje "
  sSql = sSql & "FROM ProductoImpuesto LEFT JOIN Impuestos ON ProductoImpuesto.Impuesto = Impuestos.Impuesto "
  sSql = sSql & "WHERE ProductoImpuesto.Producto='" & pProducto & "'"
   
  rs.CursorLocation = adUseClient
  rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  If rs.RecordCount <> 0 Then
     TraerTasaProducto = rs!Porcentaje
  End If
  Conexion.Close
  
Exit Function
errHandler:
  err.Raise err.Number


End Function

Public Function TraerComproCliente(pCliente As String, pFrame As String) As String
  
  If pFrame = "P" Then
      If pCliente = "." Then
         TraerComproCliente = DatoCombo("Comprobantes", "Id", "Descripcion", RegPorDefecto("VentaConsFinal"))
         Exit Function
      End If
                 
      Select Case TraerValorDeUnCampo("Clientes", "TipoIva", "Cliente='" & pCliente & "'")
             Case 1
                  If nControlador = 2 Then
                     TraerComproCliente = DatoCombo("Comprobantes", "Id", "Descripcion", RegPorDefecto("VentaOtraCondicion"))
                  Else
                     TraerComproCliente = DatoCombo("Comprobantes", "Id", "Descripcion", RegPorDefecto("VentaConsFinal"))
                  End If
             Case 3, 5, 6
                  TraerComproCliente = DatoCombo("Comprobantes", "Id", "Descripcion", RegPorDefecto("VentaOtraCondicion"))
             Case 2, 4
                  TraerComproCliente = DatoCombo("Comprobantes", "Id", "Descripcion", RegPorDefecto("VentaResponsable"))
      End Select
  Else
      TraerComproCliente = DatoCombo("Comprobantes", "Id", "Descripcion", RegPorDefecto("VentaConsFinalN"))
  End If
      
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function DameCuentaCorriente(sCliDesde As String) As ADODB.Recordset
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Set Conexion = New ADODB.Connection

  Conexion.Open sDb
  
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  sSql = "SELECT CuentaCorriente.* "
  sSql = sSql & "FROM CuentaCorriente LEFT JOIN Comprobantes ON CuentaCorriente.Comprobante = Comprobantes.Id "
  sSql = sSql & "WHERE Comprobantes.TipoOperacion=1 "
  sSql = sSql & "AND Cliente='" & sCliDesde & "' and Anulado<> 1 and Debe<>0 or (Cliente='" & sCliDesde & "' and Debe=0 and CuentaCorriente.Numero=0 and CuentaCorriente.sucursal=0 and CuentaCorriente.Cuota=0) "
  sSql = sSql & "ORDER BY CuentaCorriente.ID"
  
  rs.CursorLocation = adUseClient
  rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
  rs.ActiveConnection = Nothing
  Set DameCuentaCorriente = rs
  Conexion.Close
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function DameCodComprobante(nCod As Integer) As String
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Set Conexion = New ADODB.Connection

  Conexion.Open sDb
  
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT Descripcion FROM Comprobantes WHERE id=" & nCod
  rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     DameCodComprobante = rs!Descripcion
  End If
  Conexion.Close
  
Exit Function
errHandler:
  err.Raise err.Number
 
End Function

Public Function DamePagos(pCuota As Integer, pVenta As Long, Optional Auxi As Boolean) As Currency
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Set Conexion = New ADODB.Connection

  Conexion.Open sDb
  
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  
  If Auxi = False Then
     sSql = "SELECT SUM(Aplicaciones.Importe) AS total "
     sSql = sSql & "FROM Aplicaciones "
     sSql = sSql & "WHERE Aplicaciones.Venta=" & pVenta & " AND Aplicaciones.Cuota=" & pCuota & " AND Aplicaciones.Anulado<>1"
  Else
     sSql = "SELECT SUM(AplicacionAuxiliar.Importe) AS total "
     sSql = sSql & "FROM AplicacionAuxiliar "
     sSql = sSql & "WHERE AplicacionAuxiliar.Venta=" & pVenta & " AND AplicacionAuxiliar.Cuota=" & pCuota & " AND AplicacionAuxiliar.Anulado<>1"
  End If
          
  rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     If Not IsNull(rs!Total) Then
        DamePagos = rs!Total
     End If
  End If
  Conexion.Close
  
Exit Function
errHandler:
  err.Raise err.Number

End Function

Public Function TieneDescuento() As Boolean
   Dim cRsl As ClsLectura, bBandera As Boolean, rAuxNew As Recordset
   
   bBandera = False
   
   Set cRsl = New ClsLectura
   Set rAuxNew = cRsl.TraerRsCondi("VentasDescuentos", "Id", "Id=1")
   
   If rAuxNew!Activo = True Then
      Select Case UCase(WeekdayName(Weekday(DateSerial(Year(Date), Month(Date), Day(Date))), False, vbSunday))
             Case "LUNES"
                   bBandera = rAuxNew!Lunes 'cRsl.TraerValorDeUnCampo("VentasDescuentos", "Lunes", "Id=1")
             Case "MARTES"
                   bBandera = rAuxNew!Martes ' cRsl.TraerValorDeUnCampo("VentasDescuentos", "Martes", "Id=1")
             Case "MIERCOLES"
                    bBandera = rAuxNew!Miercoles ' cRsl.TraerValorDeUnCampo("VentasDescuentos", "Miercoles", "Id=1")
             Case "JUEVES"
                   bBandera = rAuxNew!Jueves ' cRsl.TraerValorDeUnCampo("VentasDescuentos", "Jueves", "Id=1")
             Case "VIERNES"
                   bBandera = rAuxNew!Viernes ' cRsl.TraerValorDeUnCampo("VentasDescuentos", "Viernes", "Id=1")
             Case "SABADO"
                   bBandera = rAuxNew!Sabado 'cRsl.TraerValorDeUnCampo("VentasDescuentos", "Sabado", "Id=1")
             Case "DOMINGO"
                   bBandera = rAuxNew!Domingo  'cRsl.TraerValorDeUnCampo("VentasDescuentos", "Domingo", "Id=1")
      End Select
      
       If bBandera = True Then
         dPorcDescuento = rAuxNew!Descuento 'cRsl.TraerValorDeUnCampo("VentasDescuentos", "Descuento", "Id=1")
         sFormaPago = rAuxNew!formapago  ' cRsl.TraerValorDeUnCampo("VentasDescuentos", "FormaPago", "Id=1")
       End If
   End If
   TieneDescuento = bBandera
End Function

Public Function ObtenerTotalAcumulado(pDesde As String, pHasta As String, pProv As String) As Single
  Dim sSql As String
  Dim Conexion As ADODB.Connection
  Set Conexion = New ADODB.Connection

  Conexion.Open sDb
  
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  

  sSql = "SELECT SUM(Importe) AS Importe "
  sSql = sSql & "FROM TblGananciaAcum "
  sSql = sSql & "WHERE Fecha>='" & pDesde & "' and Fecha<= '" & pHasta & "' "
  sSql = sSql & "AND Proveedor='" & pProv & "' and Anulado=0 "
  
  rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
  If rs.EOF = False Or rs.BOF = False Then
     If Not IsNull(rs!importe) Then
        ObtenerTotalAcumulado = rs!importe
     End If
  End If
  
  Conexion.Close
  
Exit Function
errHandler:
  err.Raise err.Number
  
End Function

Public Function TraerRecargoCliente(pCliente As String) As Double
 Dim sSql As String
  Dim Conexion As ADODB.Connection
  Set Conexion = New ADODB.Connection

  Conexion.Open sDb
  
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT Recargos.Porcentaje as Porcentaje "
  sSql = sSql & "FROM Clientes LEFT OUTER JOIN "
  sSql = sSql & "Recargos ON Clientes.Recargo = Recargos.Recargo "
  sSql = sSql & "WHERE Clientes.Cliente='" & pCliente & "'"
    
  rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
  
  If rs.EOF = False Or rs.BOF = False Then
     If Not IsNull(rs!Porcentaje) Then
        TraerRecargoCliente = rs!Porcentaje
     End If
  End If
  
  Conexion.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function TraerDescuentoCliente(pCliente As String) As Double
 Dim sSql As String
  Dim Conexion As ADODB.Connection
  Set Conexion = New ADODB.Connection

  Conexion.Open sDb
  
  Dim rs As ADODB.Recordset
  
  On Error GoTo errHandler
  Set rs = New ADODB.Recordset
  sSql = "SELECT Descuentos.Porcentaje as Porcentaje "
  sSql = sSql & "FROM Clientes LEFT OUTER JOIN "
  sSql = sSql & "Descuentos ON Clientes.Descuento = Descuentos.Descuento "
  sSql = sSql & "WHERE Clientes.Cliente='" & pCliente & "'"
    
  rs.Open sSql, Conexion, adOpenStatic, adLockOptimistic
  
  If rs.EOF = False Or rs.BOF = False Then
     If Not IsNull(rs!Porcentaje) Then
        TraerDescuentoCliente = rs!Porcentaje
     End If
  End If
  
  Conexion.Close
  
Exit Function
errHandler:
  err.Raise err.Number
End Function

Public Function ActualizarCampoTraer(pCampo As String) As Long
  Dim db As Connection, rs As Recordset
  
  Set db = New Connection
  Set rs = New Recordset
  On Error GoTo errHandler
  Dim sSql As String
  
  db.Open sDb
  db.BeginTrans
  
  sSql = "UPDATE Sistema SET " & pCampo & "=" & pCampo & "+1"
  db.Execute sSql
  
  db.CommitTrans
  
  sSql = "SELECT " & pCampo & " as Valor FROM Sistema"
  rs.Open sSql, db, adOpenStatic, adLockOptimistic
  
  ActualizarCampoTraer = rs!valor
  db.Close
  
 Exit Function
errHandler:
    db.RollbackTrans
    db.Close
    MsgBox err.Description, vbCritical, "ActualizarMovimiento"

End Function

Public Function TraerCodigoProducto(pCodigo As String) As String
  Dim db As Connection, rs As Recordset
  
  Set db = New Connection
  Set rs = New Recordset
  On Error GoTo errHandler
  Dim sSql As String
  Dim sCodigo As String
  
  db.Open sDb
    
  If pCodigo <> "" Then
      sSql = "SELECT Producto FROM Productos WHERE CBarra='" & pCodigo & "' "
    
      rs.Open sSql, db, adOpenStatic, adLockOptimistic
      
      If rs.RecordCount = 0 Then
         sCodigo = pCodigo
      Else
         sCodigo = rs!Producto
      End If
  End If
  
  TraerCodigoProducto = sCodigo
  db.Close
  
 Exit Function
errHandler:
    db.Close
    MsgBox err.Description, vbCritical, "ActualizarMovimiento"

End Function
