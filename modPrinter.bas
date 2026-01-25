Attribute VB_Name = "modPrinter"
Option Explicit

Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long
Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" (ByVal pszBuffer As String, ByRef pcchBuffer As Long) As Long

Public Function GetDefaultPrinterName() As String
    Dim n As Long, s As String
    n = 0
    GetDefaultPrinter vbNullString, n
    If n <= 0 Then Exit Function
    s = String$(n, vbNullChar)
    If GetDefaultPrinter(s, n) <> 0 Then
        GetDefaultPrinterName = Left$(s, n - 1)
    End If
End Function

Public Function SetDefaultPrinterName(ByVal prn As String) As Boolean
    SetDefaultPrinterName = (SetDefaultPrinter(prn) <> 0)
End Function

Public Function CrystalSafeRecordset(ByVal Rs As ADODB.Recordset, ByVal tmpPath As String) As ADODB.Recordset
    Dim R As New ADODB.Recordset
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0
    Rs.Save tmpPath, adPersistADTG
    R.Open tmpPath
    Set CrystalSafeRecordset = R
End Function

