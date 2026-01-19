VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSqlVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function VerificarVersion() As Single
  Dim Rs As ADODB.Recordset
  On Error GoTo errHandler
  
  Set Rs = New ADODB.Recordset
 
  Dim sSql As String
  
  db.Open sDb
  db.BeginTrans
  
 ' Version Sistema
  sSql = " IF not exists "
  sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
  sSql = sSql & "WHERE COLUMN_NAME = 'Version' AND TABLE_NAME = 'Sistema') "
  sSql = sSql & "BEGIN "
  sSql = sSql & "ALTER TABLE Sistema ADD Version  Decimal(5,2) not null DEFAULT 17.9 WITH VALUES "
  sSql = sSql & "END"
  db.Execute sSql
  
  
'  sSql = "IF EXISTS (SELECT * FROM sysobjects WHERE name='EmailDatos') "
'  sSql = sSql & "BEGIN "
'  sSql = sSql & "UPDATE Sistema SET Version=18.1 "
'  sSql = sSql & "END"
'  db.Execute sSql

         
   db.CommitTrans
  db.Close
    
  db.Open sDb

  sSql = "SELECT Version FROM SISTEMA "
    Rs.Open sSql, db, adOpenStatic, adLockOptimistic
   VerificarVersion = Rs!Version
   
  db.Close
 
Exit Function
    
errHandler:
   db.RollbackTrans
   db.Close
   MsgBox err.Description, vbCritical, "Atención"
    
End Function

Public Sub ActualizarVersion(pVersion As Single)
  Dim cRsl As ClsLectura, nListaVersion As Integer, nCuentaVersion As Integer
  
  On Error GoTo errHandler
    
  Set cRsl = New ClsLectura
    
  nListaVersion = cRsl.RegPorDefecto("ListaPrecio")
  nCuentaVersion = cRsl.TraerValorDeUnCampo("OperacionesContables", "Cuenta", "Id=1")
  
  Dim sSql As String
      
  db.Open sDb
  db.BeginTrans
      
  If pVersion < 18 Then

    ' Mueve Stock en Sistemas
    sSql = " IF NOT EXISTS  "
    sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
    sSql = sSql & "WHERE COLUMN_NAME = 'MueveStock' AND TABLE_NAME = 'Sistema') "
    sSql = sSql & "BEGIN "
    sSql = sSql & "ALTER TABLE Sistema ADD MueveStock  bit not null DEFAULT 1 WITH VALUES "
    sSql = sSql & "END"
    db.Execute sSql
    
        ' Mueve Caja en Sistemas
    sSql = " IF NOT EXISTS  "
    sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
    sSql = sSql & "WHERE COLUMN_NAME = 'MueveCaja' AND TABLE_NAME = 'Sistema') "
    sSql = sSql & "BEGIN "
    sSql = sSql & "ALTER TABLE Sistema ADD MueveCaja  bit not null DEFAULT 1 WITH VALUES "
    sSql = sSql & "END"
    db.Execute sSql
    
    sSql = "UPDATE Sistema SET Version=18"
    db.Execute sSql
    
    GoTo email
    
  End If
  
  If pVersion = 18 Then
email:
     sSql = "(SELECT * From INFORMATION_SCHEMA.Columns "
     sSql = sSql & "WHERE COLUMN_NAME = 'Email' AND TABLE_NAME = 'Sistema') "
     sSql = sSql & "BEGIN "
     sSql = sSql & "ALTER TABLE Sistema ADD Email  tinyint not null DEFAULT 0 WITH VALUES "
     sSql = sSql & "END"
     db.Execute sSql
    
         
     sSql = "IF NOT EXISTS  (SELECT * FROM sysobjects WHERE name='EmailDatos') "
     sSql = sSql & "BEGIN "
     sSql = sSql & "CREATE TABLE [dbo].[EmailDatos]( "
     sSql = sSql & "[Id] [tinyint]  NOT NULL,"
     sSql = sSql & "[Servidor] [varchar](50) NULL,"
     sSql = sSql & "[Envia] [varchar](50) NULL,"
     sSql = sSql & "[Puerto] [int] NULL,"
     sSql = sSql & "[Usuario] [varchar](50) NULL,"
     sSql = sSql & "[Contraseña] [varchar](50) NULL,"
     sSql = sSql & "[Autenticacion] [bit] NULL,"
     sSql = sSql & "[SSL] [bit] NULL "
     sSql = sSql & "CONSTRAINT [PK_EmailDatos] PRIMARY KEY CLUSTERED "
     sSql = sSql & "( "
     sSql = sSql & "[Id] Asc "
     sSql = sSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
     sSql = sSql & ") ON [PRIMARY] "
     sSql = sSql & "END"
     
     db.Execute sSql

     sSql = "IF NOT EXISTS  (SELECT * FROM sysobjects WHERE name='EmailDatos') "
     sSql = sSql & "BEGIN "
     sSql = sSql & "INSERT [dbo].[EmailDatos] ([Id], [Servidor], [Envia], [Puerto], [Usuario], [Contraseña], [Autenticacion], [SSL]) VALUES (1, N'mail.dagsis.com.ar', N'dagsis@dagsis.com.ar', 587, N'dagsis@dagsis.com.ar', N'contraseña', 1, 0)"
     sSql = sSql & "END"
     
     db.Execute sSql
     
     sSql = "IF NOT EXISTS  (SELECT * FROM sysobjects WHERE name='EmailDatos') "
     sSql = sSql & "BEGIN "
     sSql = sSql & "INSERT [dbo].[EmailDatos] ([Id], [Servidor], [Envia], [Puerto], [Usuario], [Contraseña], [Autenticacion], [SSL]) VALUES (2, N'smtp.live.com', N'cuenta@hotmail.com', 25, N'cuenta@hotmail.com', N'contraseña', 1, 1)"
     sSql = sSql & "END"
     
     db.Execute sSql
     
     sSql = "IF NOT EXISTS  (SELECT * FROM sysobjects WHERE name='EmailDatos') "
     sSql = sSql & "BEGIN "
     sSql = sSql & "INSERT [dbo].[EmailDatos] ([Id], [Servidor], [Envia], [Puerto], [Usuario], [Contraseña], [Autenticacion], [SSL]) VALUES (3, N'smtp.gmail.com', N'dagsis@gmail.com.ar', 465, N'dagsis@gmail.com', N'contraseña', 1, 1)"
     sSql = sSql & "END"
     db.Execute sSql
     
     sSql = "IF NOT EXISTS  (SELECT * FROM sysobjects WHERE name='EmailDatos') "
     sSql = sSql & "BEGIN "
     sSql = sSql & "INSERT [dbo].[EmailDatos] ([Id], [Servidor], [Envia], [Puerto], [Usuario], [Contraseña], [Autenticacion], [SSL]) VALUES (4, N'smtp.mail.yahoo.com', N'cuenta@yahoo.com.ar', 465, N'cuenta@yahoo.com.ar', N'contraseña', 1, 1)"
     sSql = sSql & "END"
     db.Execute sSql
     
     sSql = "IF NOT EXISTS  (SELECT * FROM sysobjects WHERE name='EmailDatos') "
     sSql = sSql & "BEGIN "
     sSql = sSql & "INSERT [dbo].[UsuariosPermisos] ([Usuario], [Tipo], [Opcion], [Orden], [Descripcion], [Valor], [Ver]) VALUES (1, N'Herrami', 96, 6, N'Datos Correo Electronico', 1, 1)"
     sSql = sSql & "END"
     db.Execute sSql
     
     db.Execute "UPDATE Sistema SET Version=18.1"
               
   End If

    If pVersion <= 18.1 Then
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'ListaRes' AND TABLE_NAME = 'Registros') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Registros ADD ListaRes  tinyint not null DEFAULT " & nListaVersion & " WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
     
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'ListaDel' AND TABLE_NAME = 'Registros') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Registros ADD ListaDel  tinyint not null DEFAULT " & nListaVersion & " WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'CuentaRes' AND TABLE_NAME = 'Registros') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Registros ADD CuentaRes  tinyint not null DEFAULT " & nCuentaVersion & " WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'CuentaDel' AND TABLE_NAME = 'Registros') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Registros ADD CuentaDel  tinyint not null DEFAULT " & nCuentaVersion & " WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'CliProvAut' AND TABLE_NAME = 'Registros') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Registros ADD CliProvAut  bit not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql

                  
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Cliente' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Cliente  int not null DEFAULT 1 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
    
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Proveedor' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Proveedor  int not null DEFAULT 1 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       db.Execute "UPDATE Sistema SET Version=19.0"

    End If
  
    If pVersion <= 19 Then
       sSql = "IF NOT EXISTS  (SELECT * FROM sysobjects WHERE name='MozoMesa') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "CREATE TABLE [dbo].[MozoMesa]( "
       sSql = sSql & "[Id] [int] IDENTITY(1,1) NOT NULL,"
       sSql = sSql & "[Mozo] [int] NULL,"
       sSql = sSql & "[Mesa] [int] NULL,"
       sSql = sSql & "CONSTRAINT [aaaaaMOZOMESA_PK] PRIMARY KEY NONCLUSTERED "
       sSql = sSql & "("
       sSql = sSql & "[Id] Asc "
       sSql = sSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
       sSql = sSql & ") ON [PRIMARY]"
       sSql = sSql & "END"

       db.Execute sSql

       db.Execute "UPDATE Sistema SET Version=19.1"

   End If
   
   
   If pVersion < 19.2 Then
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'DescCaja' AND TABLE_NAME = 'Registros') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Registros ADD DescCaja  bit not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       db.Execute "UPDATE Sistema SET Version=19.9"

   End If
   
    If pVersion < 20 Then
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Percibe' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Percibe  bit not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Retiene' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Retiene  bit not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       
       db.Execute "UPDATE Sistema SET Version=20.1"

   End If
   
   If pVersion <= 20.2 Then
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Redondea' AND TABLE_NAME = 'Registros') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Registros ADD Redondea  bit not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'AnulaPres' AND TABLE_NAME = 'Registros') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Registros ADD AnulaPres  bit not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       db.Execute "UPDATE Sistema SET Version=20.3"
       
   End If
   
    If pVersion <= 20.3 Then
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Envase' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Envase varchar(10) not null DEFAULT '2999909-7' WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'TextoEnvase' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD TextoEnvase varchar(15) not null DEFAULT 'RECARGO ENVASE' WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql

       
       db.Execute "UPDATE Sistema SET Version=20.4"
       
   End If
   
    If pVersion <= 20.4 Then
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'DBalanza' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD DBalanza tinyint not null DEFAULT 2 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
              
       db.Execute "UPDATE Sistema SET Version=20.5"
       
   End If
   
     If pVersion <= 20.5 Then
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Status' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Status int not null DEFAULT 21990764 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
        sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Status' AND TABLE_NAME = 'Empresa') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Empresa ADD CFondo int not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
              
       db.Execute "UPDATE Sistema SET Version=21.0"
       
   End If
   
    If pVersion <= 21 Then
    
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'CBarra' AND TABLE_NAME = 'Productos') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Productos ADD CBarra Varchar(25) not null DEFAULT '' WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'ValorEnvase' AND TABLE_NAME = 'Productos') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Productos ADD ValorEnvase decimal(18,4) not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Balanza' AND TABLE_NAME = 'Productos') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Productos ADD Balanza int not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'CDepartamento' AND TABLE_NAME = 'Productos') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Productos ADD CDepartamento int not null DEFAULT 1 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'CFamilia' AND TABLE_NAME = 'Productos') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Productos ADD CFamilia int not null DEFAULT 1 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       
       Dim NumArchivo As Integer
       Dim s_script As String
       Dim strLinea As String
      
      ' Obtengo un número libre de archivo
       NumArchivo = FreeFile
      
      ' Abro el archivo.
        Open App.Path & "\Scripts\Consulta.sql" For Input As NumArchivo
      
'      ' Leo el archivo completo
'        s_script = Input(LOF(NumArchivo), #NumArchivo)
         Do While Not EOF(1)
            s_script = s_script + strLinea + Chr$(13)
            Line Input #1, strLinea
            If strLinea = "FINAL" Then
               db.Execute s_script, , adCmdText + adExecuteNoRecords
               s_script = ""
               strLinea = ""
            End If
         Loop
      
      ' Cierro el archivo
        Close NumArchivo
                           
       db.Execute "UPDATE Sistema SET Version=21.5"
       
   End If
   
    If pVersion <= 21.5 Then
   
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'FAmericana' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD FAmericana  bit not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       db.Execute "UPDATE Sistema SET Version=21.6"

   End If
   
    If pVersion <= 21.6 Then
   
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Comanda' AND TABLE_NAME = 'UsuariosComprobantesVentas') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE UsuariosComprobantesVentas ADD Comanda  bit not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       db.Execute "UPDATE Sistema SET Version=21.7"

   End If
   
    If pVersion <= 21.8 Then
   
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Cbu' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Cbu  varchar(50) not null DEFAULT 'XXXX' WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       db.Execute "UPDATE Sistema SET Version=22"

   End If
   
    If pVersion <= 22 Then
   
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Qr' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Qr int not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       db.Execute "UPDATE Sistema SET Version=22.1"

   End If
   
     
'    If pVersion <= 22.1 Then
'
'       sSql = " IF NOT EXISTS  "
'       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
'       sSql = sSql & "WHERE COLUMN_NAME = 'Qr' AND TABLE_NAME = 'Turnos') "
'       sSql = sSql & "BEGIN "
'       sSql = sSql & "ALTER TABLE Turnos ADD Caja int not null DEFAULT 1 WITH VALUES "
'       sSql = sSql & "END"
'       db.Execute sSql
'
'       db.Execute "UPDATE Sistema SET Version=22.5"
'
'   End If
   
   If pVersion <= 23 Then
   
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Socket' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Socket int not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = "IF NOT EXISTS  (SELECT * FROM sysobjects WHERE name='Eventos') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "CREATE TABLE [dbo].[Eventos]( "
       sSql = sSql & "[Id] [int] IDENTITY(1,1) NOT NULL,"
       sSql = sSql & "[Evento] [text] NULL,"
       sSql = sSql & "[Procesado] [int] NULL "
       sSql = sSql & "CONSTRAINT [PK_Eventos] PRIMARY KEY CLUSTERED "
       sSql = sSql & "( "
       sSql = sSql & "[Id] Asc "
       sSql = sSql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
       sSql = sSql & ") ON [PRIMARY] "
       sSql = sSql & "END"
     
       db.Execute sSql
       
       db.Execute "UPDATE Sistema SET Version=23.1"

   End If
   
   If pVersion <= 23.1 Then
   
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Compania' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Compania int not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
       
       sSql = " IF NOT EXISTS  "
       sSql = sSql & "(SELECT * From INFORMATION_SCHEMA.Columns "
       sSql = sSql & "WHERE COLUMN_NAME = 'Sucursal' AND TABLE_NAME = 'Sistema') "
       sSql = sSql & "BEGIN "
       sSql = sSql & "ALTER TABLE Sistema ADD Sucursal int not null DEFAULT 0 WITH VALUES "
       sSql = sSql & "END"
       db.Execute sSql
              
     
       
       db.Execute "UPDATE Sistema SET Version=23.5"

   End If
    
  db.CommitTrans
  db.Close

Exit Sub
    
errHandler:
  db.RollbackTrans
  db.Close
  MsgBox err.Description, vbCritical, "Atención"
End Sub
