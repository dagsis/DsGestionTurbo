ALTER Proc [dbo].[ProductosDat_A]
	(	
	@Producto varchar(25) = null,
	@Descripcion varchar(100) = null,
	@Cantidad decimal(10,2),
	@Impuesto tinyint = null,
	@Precio decimal(12,4) = null,
	@Fecha smalldatetime,
	@Balanza int,
	@CDepartamento int,
	@CFamilia int,
	@CodCompu nvarchar(2),
	@Fijo nvarchar(2))

	AS 
	
	DECLARE @Canti int
	
    SELECT @Canti = COUNT(producto) FROM Productos WHERE Producto=@Producto
	
    IF @Canti = 0
	   BEGIN
		INSERT INTO [Productos]
			   ([Producto]
			   ,[Moneda]
			   ,[Descripcion]
			   ,[Descripcion1]
			   ,[Tipo]              
			   ,[Familia]
			   ,[FamiliaSub]
			   ,[Representante]
			   ,[Marca]
			   ,[FechaAlta]
			   ,[FechaBaja]
			   ,[Proveedor]
			   ,[UMCompra]
			   ,[CantDespacho]
			   ,[UMVenta]
			   ,[CantVenta]
			   ,[Deposito]
			   ,[Origen]
			   ,[FechaCompra]
			   ,[PrecioCompra]
			   ,[Actualiza]
			   ,[CantMinima]
			   ,[CantMaxima]
			   ,[ActuaStock]
			   ,[CodCompu]
			   ,[Fijo]
			   ,[TipoInterno]
			   ,[ValorInterno]
			   ,[Calcular]
			   ,[CBarra],
			    [ValorEnvase],
				[Balanza],[CDepartamento],[CFamilia])
		 VALUES
			   (@Producto,1,@Descripcion,'',3,1,Null,0,1,@Fecha,@Fecha,1,1,1,1,1,1,1,@Fecha,0,0,0,0,'Si',@CodCompu,@Fijo,'No',0,'No','',0,@Balanza,@CDepartamento,@CFamilia)
		       
			   INSERT INTO Precios (Lista,Producto,Porcentaje,Precio) values (1,@Producto,0,@Precio)
		       
			   INSERT INTO Cantidades (Suc,Movimiento,Estado,Producto,Deposito,CantidadDebe,CantidadHaber,Anulado) VALUES (1,0,1,@Producto,1,0,0,0)
			   INSERT INTO ProductoUnidades (Producto,CantidadA,UnidadA,CantidadB,UnidadB) VALUES (@Producto,1,1,1,1)               
			   INSERT INTO ProductoImpuesto (Producto,Impuesto) VALUES (@Producto,@Impuesto);
	    END		   
	ELSE
	   BEGIN     
	     UPDATE Productos SET Descripcion=@Descripcion,FechaAlta=@Fecha,FechaBaja=@Fecha WHERE Producto=@Producto
	     UPDATE Precios SET Precio=@Precio WHERE Producto=@Producto
	   END 
	    
FINAL

ALTER Proc [dbo].[ProductoAgregar]
	(@Activo bit = null,
	@Producto nvarchar(25),
@Moneda tinyint,	
@Descripcion varchar(100) = null,
@Descripcion1 varchar(100) = null,
@Tipo int = null,
@Familia int = null,
@FamiliaSub int = null,
@Representante varchar(15),
@Peso decimal(10,4),
@Marca int = null,
@FechaAlta datetime = null,
@FechaBaja datetime = null,
@Proveedor int = null,
@UMCompra int,
@CantDespacho real = null,
@UMVenta int,
@CantVenta real = null,
@Deposito int = null,
@Origen int = null,
@FechaCompra datetime = null,
@PrecioCompra money = null,
@Actualiza bit,
@CantMinima float(15) = null,
@CantMaxima float(15) = null,
@ActuaStock nvarchar(2) = null,
@CodCompu nvarchar(2) = null,
@Fijo nvarchar(2) = null,
@TipoInterno nvarchar(10) = null,
@ValorInterno real = null,
@Calcular nvarchar(2) = null,
@CBarra varchar(50),
@ValorEnvase decimal(18,4),
@Balanza int,@CDepartamento int,@CFamilia int)
AS 
 INSERT INTO [Productos]
 (Activo,
 Producto,
Moneda, 
Descripcion,
Descripcion1,
Tipo,
Familia,
FamiliaSub,
Representante,
Peso,
Marca,
FechaAlta,
FechaBaja,
Proveedor,
UMCompra,
CantDespacho,
UMVenta,
CantVenta,
Deposito,
Origen,
FechaCompra,
PrecioCompra,
Actualiza,
CantMinima,
CantMaxima,
ActuaStock,
CodCompu,
Fijo,
TipoInterno,
ValorInterno,
Calcular,
CBarra,
ValorEnvase,
Balanza,CDepartamento,CFamilia)
 VALUES(
    @Activo,
@Producto,
@Moneda,	
@Descripcion,
@Descripcion1,
@Tipo,
@Familia,
@FamiliaSub,
@Representante,
@Peso,
@Marca,
@FechaAlta,
@FechaBaja,
@Proveedor,
@UMCompra,
@CantDespacho,
@UMVenta,
@CantVenta,
@Deposito,
@Origen,
@FechaCompra,
@PrecioCompra,
@Actualiza,
@CantMinima,
@CantMaxima,
@ActuaStock,
@CodCompu,
@Fijo,
@TipoInterno,
@ValorInterno,
@Calcular,
@CBarra,
@ValorEnvase,
@Balanza,@CDepartamento,@CFamilia)


INSERT INTO PreciosCaja (Producto,Precio) VALUES(@Producto,0)


FINAL

ALTER Proc [dbo].[ProductoModif]
	(@Activo bit,
	@Producto nvarchar(25),
@Moneda tinyint,	
@Descripcion varchar(100) = null,
@Descripcion1 varchar(100) = null,
@Tipo int = null,
@Familia int = null,
@FamiliaSub int = null,
@Representante varchar(15),
@Peso decimal(10,4),
@Marca int = null,
@FechaAlta datetime = null,
@FechaBaja datetime = null,
@Proveedor int = null,
@UMCompra int,
@CantDespacho real = null,
@UMVenta int,
@CantVenta real = null,
@Deposito int = null,
@Origen int = null,
@FechaCompra datetime = null,
@PrecioCompra money = null,
@Actualiza bit,
@CantMinima float(15) = null,
@CantMaxima float(15) = null,
@ActuaStock nvarchar(2) = null,
@CodCompu nvarchar(2) = null,
@Fijo nvarchar(2) = null,
@TipoInterno nvarchar(10) = null,
@ValorInterno real = null,
@Calcular nvarchar(2) = null,
@CBarra varchar(25),
@ValorEnvase decimal(18,4),
@Balanza int,@CDepartamento int,@CFamilia int)
AS 
 Update [Productos]
SET 
Activo=@Activo, 
Moneda=@Moneda,
Descripcion=@Descripcion,
Descripcion1=@Descripcion1,
Tipo=@Tipo,
Familia=@Familia,
FamiliaSub=@FamiliaSub,
Representante=@Representante,
Peso=@Peso,
Marca=@Marca,
FechaAlta=@FechaAlta,
FechaBaja=@FechaBaja,
Proveedor=@Proveedor,
UMCompra=@UMCompra,
CantDespacho=@CantDespacho,
UMVenta=@UMVenta,
CantVenta=@CantVenta,
Deposito=@Deposito,
Origen=@Origen,
FechaCompra=@FechaCompra,
PrecioCompra=@PrecioCompra,
Actualiza=@Actualiza,
CantMinima=@CantMinima,
CantMaxima=@CantMaxima,
ActuaStock=@ActuaStock,
CodCompu=@CodCompu,
Fijo=@Fijo,
TipoInterno=@TipoInterno,
ValorInterno=@ValorInterno,
Calcular=@Calcular,
CBarra=@CBarra,
ValorEnvase= @ValorEnvase,
Balanza=@Balanza,CDepartamento=@CDepartamento,CFamilia=@CFamilia
 where  (Producto=@Producto)

FINAL

CREATE TABLE [dbo].[Balanzas](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Balanza] [int] NULL,
	[Relacion] [int] NULL
) ON [PRIMARY]

FINAL