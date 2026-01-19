VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmGrilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grilla"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   10980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdMover 
      Caption         =   "|<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   4875
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1005
      TabIndex        =   2
      Top             =   4875
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   1935
      TabIndex        =   1
      Top             =   4875
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
      Caption         =   ">>|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   2910
      TabIndex        =   0
      Top             =   4875
      Width           =   840
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   4740
      Left            =   75
      TabIndex        =   4
      Top             =   60
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   8361
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cn As ADODB.Connection
Private Rs2 As ADODB.Recordset
Const PAGE_SIZE = 50
Dim Registro As Integer

Private WithEvents Rsd As ADODB.Recordset
Attribute Rsd.VB_VarHelpID = -1

Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
  Rsd.Sort = Rsd.Fields(ColIndex).Name & " Asc"
End Sub

Private Sub CmdMover_Click(Index As Integer)
   Dim Pagina_Actual As Long
       
    With Rsd
        Select Case Index
               
            ' Primera página del recordset
            Case 0
                Pagina_Actual = 1
                Registro = PAGE_SIZE
            ' Anterior página del recordset
            Case 1
                Pagina_Actual = .AbsolutePage - 1
                Registro = Registro - PAGE_SIZE
                If Registro < PAGE_SIZE Then
                    Registro = PAGE_SIZE
                End If
            ' Siguiente página
            Case 2
                Pagina_Actual = .AbsolutePage + 1
                Registro = Registro + PAGE_SIZE
                If Registro > Rsd.RecordCount Then
                    Registro = Rsd.RecordCount
                End If
            ' Ultima
            Case 3
                Pagina_Actual = .PageCount
                Registro = Rsd.RecordCount
        End Select
           
        ' Si estamos en la primera o última pagina sale
        If Pagina_Actual < 1 Or Pagina_Actual > .PageCount Then
            Exit Sub
        End If
           
        ' Establece la Página actual al rec
        .AbsolutePage = Pagina_Actual
           
        ' Asigna el recordset paginado al control Datagrid
        Set Grid1.DataSource = Paginar_Recordset(Rsd)
     '   CabGrid
        ' visualiza en el caption la página actual y los registros
        Me.Caption = Me.Caption & "  -  Registro actual: " & Registro & _
                    " de " & Rsd.RecordCount & " registros ."
           
    End With

End Sub

Private Sub Form_Load()
   Conectar nLlamo, fDesde, fHasta
End Sub

Private Sub Conectar(pLlamo As Byte, pDesde As Date, pHasta As Date)
  Dim sSql As String
  
  Dim Cadena As String
  Set cn = New ADODB.Connection
  
  cn.ConnectionString = sDb
  cn.Open
  
  Set Rsd = New ADODB.Recordset
  Set Rs2 = New ADODB.Recordset
  
  Select Case pLlamo
         Case 1
               sSql = "SELECT Vendedores.Descripcion, DetallesComprobantes.Producto, DetallesComprobantes.Descripcion, CabComprobantes.Fecha, Comprobantes.Descripcion, CabComprobantes.Numero, CabComprobantes.Sucursal, Clientes.RazonSocial, Vendedores.Descripcion, Cajas.Descripcion, DetallesComprobantes.Cantidad ,DetallesComprobantes.PrecioUnitario,DetallesComprobantes.PrecioTotal,DetallesComprobantes.Impuesto,Comprobantes.Comprobante "
               sSql = sSql & "FROM ((((CabComprobantes LEFT JOIN Comprobantes ON CabComprobantes.Comprobante = Comprobantes.Id) LEFT JOIN Vendedores ON CabComprobantes.Vendedor = Vendedores.Vendedor) LEFT JOIN Clientes ON CabComprobantes.Cliente = Clientes.Cliente) LEFT JOIN Cajas ON CabComprobantes.Caja = Cajas.Caja) LEFT JOIN DetallesComprobantes ON CabComprobantes.Movimiento = DetallesComprobantes.Movimiento "
               sSql = sSql & "WHERE (((CabComprobantes.Anulado)<>1) AND ((CabComprobantes.Tipo)=1)) AND CabComprobantes.Suc=" & nSucursal & " "
               sSql = sSql & " and CabComprobantes.Fecha>='" & FechaAmericana(pDesde) & "' and CabComprobantes.Fecha<='" & FechaAmericana(pHasta) & "' "
               sSql = sSql & " ORDER BY CabComprobantes.Fecha,DetallesComprobantes.Descripcion, CabComprobantes.Sucursal,CabComprobantes.Numero"
  End Select
  
  With Rsd
       .CursorLocation = adUseClient
       .PageSize = PAGE_SIZE
       .Properties("Initial Fetch Size") = PAGE_SIZE
       .Open sSql, cn, adOpenDynamic, adLockOptimistic
       .ActiveConnection = Nothing
       Set Grid1.DataSource = Paginar_Recordset(Rsd)
      ' CabGrid
  End With
  Registro = PAGE_SIZE
  cn.Close
End Sub

Private Sub CabGrid()
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.Bold = True
  Grid1.AllowUpdate = True
  With Grid1
       .Columns(0).Caption = "Produc."
       .Columns(0).Width = 1500
       .Columns(1).Caption = "Descripción"
       .Columns(1).Width = 3650
       .Columns(2).Caption = "Desc. Adicional"
       .Columns(2).Width = 1900
       .Columns(3).Caption = "P.Compr"
       .Columns(3).Width = 1000
       .Columns(3).NumberFormat = nCantDecimales
       .Columns(3).Alignment = dbgRight
       .Columns(4).Caption = "% Venta"
       .Columns(4).Width = 1000
       .Columns(4).NumberFormat = nCantDecimales
       .Columns(4).Alignment = dbgRight
       .Columns(5).Caption = "P.Venta"
       .Columns(5).Width = 1100
       .Columns(5).NumberFormat = nCantDecimales
       .Columns(5).Alignment = dbgRight
  End With
End Sub
Private Function Paginar_Recordset(recset As ADODB.Recordset) As ADODB.Recordset
  
    Dim Fila As Long
    Dim Campo As Field
    Dim P As Long
  
    P = IIf(recset.AbsolutePage > 0, recset.AbsolutePage, 1)
       
    'Recordset desconectado
    With Rs2
           
        ' Si el recordset está abierto entonces lo cierra
        If .State = adStateOpen Then
            .Close
        End If
           
        ' Define y Crea los campos al recordset con Append
        For Each Campo In recset.Fields
            .Fields.Append Campo.Name, Campo.Type, _
                           Campo.DefinedSize, Campo.Attributes
        Next Campo
           
        ' Lo Abre para agregarle los registros
           
        Rs2.Open
           
        For Fila = 1 To PAGE_SIZE
               
            ' si está en el final sale
            If recset.EOF Then
                Exit For
            End If
               
            ' va añadiendo los registros
            .AddNew
               
            ' Recorre los campos
             For Each Campo In recset.Fields
                 ' Asigna el nombre del campo
                 Rs2(Campo.Name) = Campo.Value
             Next Campo
                
             'Actualiza el recordset desconectado y el otro recordset
             .Update
             ' mueve al siguiente registro
             recset.MoveNext
  
        Next Fila
        ' Mueve al primer registro
        .MoveFirst
        ' Establece la página actual
        recset.AbsolutePage = P
  
    End With
    ' devuelve a la función el conjunto de datos, es decir el recordset paginado
    Set Paginar_Recordset = Rs2
       
    ' Muestra la página actual en el caption del formulario
    Me.Caption = " Página: " & CStr(Rsd.AbsolutePage) & _
                 " de " & CStr(Rsd.PageCount)
  
End Function



