BEGIN TRANSACTION;
CREATE TABLE IF NOT EXISTS "Usuarios" (
	"ID_Usuario"	INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
	"NombreUsuario"	TEXT NOT NULL UNIQUE,
	"Password"	TEXT NOT NULL,
	"Rol"	TEXT NOT NULL DEFAULT 'Usuario',
	"NombreCompleto"	TEXT,
	"Email"	TEXT,
	"Activo"	INTEGER NOT NULL DEFAULT 1,
	"UltimoAcceso"	DATETIME,
	"FechaRegistro"	DATETIME DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS "Presupuestos" (
	"NumeroPresupuesto"	TEXT UNIQUE,
	"Fecha"	DATE DEFAULT CURRENT_DATE,
	"Observaciones"	TEXT,
	"CodigoCliente"	TEXT NOT NULL,
	"ID_Vendedor"	INTEGER,
	"Estado"	TEXT DEFAULT 'Pendiente',
	"ID_FormaPago"	INTEGER,
	"BaseImponible"	REAL DEFAULT 0,
	"ImporteIVA"	REAL DEFAULT 0,
	"TotalPresupuesto"	REAL DEFAULT 0,
	PRIMARY KEY("NumeroPresupuesto"),
	FOREIGN KEY("CodigoCliente") REFERENCES "Clientes"("CodigoCliente")
);
CREATE TABLE IF NOT EXISTS "LineasFacturaVenta" (
	"ID_Linea"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"NumeroFactura"	TEXT NOT NULL,
	"NumeroOrden"	INTEGER,
	"ID_Articulo"	INTEGER,
	"Descripcion"	TEXT,
	"Cantidad"	REAL DEFAULT 1,
	"PrecioUnitario"	REAL DEFAULT 0,
	"Descuento"	REAL DEFAULT 0,
	"PorcentajeIVA"	REAL DEFAULT 21,
	"Total"	REAL,
	FOREIGN KEY("ID_Articulo") REFERENCES "Articulos"("ID_Articulo"),
	FOREIGN KEY("NumeroFactura") REFERENCES "FacturasVenta"("NumeroFactura") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "FacturasVenta" (
	"NumeroFactura"	TEXT NOT NULL UNIQUE,
	"Fecha"	DATE DEFAULT CURRENT_DATE,
	"CodigoCliente"	TEXT NOT NULL,
	"NumeroAlbaran"	TEXT,
	"BaseImponible"	REAL DEFAULT 0,
	"ImporteIVA"	REAL DEFAULT 0,
	"TotalFactura"	REAL DEFAULT 0,
	"Cobrada"	INTEGER DEFAULT 0,
	"ID_FormaPago"	INTEGER,
	"ID_Vendedor"	INTEGER,
	"FechaVencimiento"	DATE,
	"NombreFiscal"	TEXT,
	"CIF"	TEXT,
	"Direccion"	TEXT,
	"Observaciones"	TEXT,
	"ID_Ruta"	INTEGER,
	"ID_Agencia"	INTEGER,
	"NumeroBultos"	INTEGER,
	"PesoTotal"	REAL,
	"CodigoSeguimiento"	TEXT,
	"Portes"	TEXT,
	"Poblacion"	TEXT,
	"CodigoPostal"	TEXT,
	"Estado"	TEXT,
	PRIMARY KEY("NumeroFactura"),
	FOREIGN KEY("NumeroAlbaran") REFERENCES "Albaranes"("NumeroAlbaran"),
	FOREIGN KEY("CodigoCliente") REFERENCES "Clientes"("CodigoCliente")
);
CREATE TABLE IF NOT EXISTS "Rutas" (
	"ID_Ruta"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"NombreZona"	TEXT NOT NULL,
	"DiasReparto"	TEXT,
	"Codigo"	TEXT,
	"Activo"	INTEGER DEFAULT 1,
	"Observaciones"	TEXT,
	"ID_Agencia"	INTEGER
);
CREATE TABLE IF NOT EXISTS "Articulos" (
	"ID_Articulo"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"CodigoReferencia"	TEXT UNIQUE,
	"Descripcion"	TEXT NOT NULL,
	"ID_Familia"	INTEGER,
	"PrecioCoste"	REAL DEFAULT 0,
	"PrecioVenta"	REAL DEFAULT 0,
	"StockActual"	REAL DEFAULT 0,
	"ID_ProveedorHabitual"	TEXT,
	"CodigoBarras"	TEXT DEFAULT '',
	"TipoIVA"	INTEGER DEFAULT 21,
	"StockMinimo"	REAL DEFAULT 0,
	"Observaciones"	TEXT DEFAULT '',
	"Activo"	INTEGER DEFAULT 1,
	FOREIGN KEY("ID_Familia") REFERENCES "Familias"("ID_Familia"),
	FOREIGN KEY("ID_ProveedorHabitual") REFERENCES "Proveedores"("CodigoProveedor")
);
CREATE TABLE IF NOT EXISTS "Proveedores" (
	"CodigoProveedor"	TEXT,
	"NombreFiscal"	TEXT NOT NULL,
	"NombreComercial"	TEXT,
	"CIF"	TEXT,
	"Direccion"	TEXT,
	"Poblacion"	TEXT,
	"Provincia"	TEXT,
	"CodigoPostal"	TEXT,
	"Telefono"	TEXT,
	"Email"	TEXT,
	"PersonaContacto"	TEXT,
	"SitioWeb"	TEXT,
	"Observaciones"	TEXT,
	"ID_FormaPago"	INTEGER,
	"Activo"	INTEGER DEFAULT 1,
	PRIMARY KEY("CodigoProveedor"),
	FOREIGN KEY("ID_FormaPago") REFERENCES "FormasPago"("ID_FormaPago")
);
CREATE TABLE IF NOT EXISTS "Albaranes" (
	"NumeroAlbaran"	TEXT UNIQUE,
	"NumeroPedido"	TEXT,
	"Fecha"	DATE DEFAULT CURRENT_DATE,
	"ID_Agencia"	INTEGER,
	"NumeroBultos"	INTEGER DEFAULT 1,
	"PesoTotal"	NUMERIC,
	"CodigoSeguimiento"	TEXT,
	"Portes"	TEXT DEFAULT 'Pagados',
	"Estado"	TEXT DEFAULT 0,
	"Observaciones"	TEXT,
	"DireccionEnvio"	TEXT,
	"Poblacion"	TEXT,
	"CodigoPostal"	TEXT,
	"FechaEntrega"	DATE,
	"BaseImponible"	NUMERIC,
	"Total"	NUMERIC,
	"CodigoCliente"	TEXT,
	"ID_Vendedor"	INTEGER,
	"ImporteIVA"	REAL DEFAULT 0,
	"ID_FormaPago"	INTEGER,
	PRIMARY KEY("NumeroAlbaran"),
	FOREIGN KEY("NumeroPedido") REFERENCES "Pedidos"("NumeroPedido"),
	FOREIGN KEY("CodigoCliente") REFERENCES "Clientes"("CodigoCliente"),
	FOREIGN KEY("ID_Agencia") REFERENCES "Agencias"("ID_Agencia"),
	FOREIGN KEY("ID_Vendedor") REFERENCES "Vendedores"("ID_Vendedor")
);
CREATE TABLE IF NOT EXISTS "Pedidos" (
	"NumeroPedido"	TEXT UNIQUE,
	"NumeroPresupuesto"	TEXT,
	"CodigoCliente"	TEXT NOT NULL,
	"Fecha"	DATE DEFAULT CURRENT_DATE,
	"Observaciones"	TEXT,
	"ID_Vendedor"	INTEGER,
	"Estado"	TEXT DEFAULT 'Pendiente',
	"BaseImponible"	NUMERIC DEFAULT 0,
	"Total"	NUMERIC DEFAULT 0,
	"FechaEntrega"	DATE,
	"Iva"	INTEGER,
	"ID_FormaPago"	INTEGER,
	"ID_Ruta"	INTEGER,
	"ImporteIVA"	REAL DEFAULT 0,
	FOREIGN KEY("NumeroPresupuesto") REFERENCES "Presupuestos"("NumeroPresupuesto"),
	PRIMARY KEY("NumeroPedido"),
	FOREIGN KEY("CodigoCliente") REFERENCES "Clientes"("CodigoCliente"),
	FOREIGN KEY("ID_Vendedor") REFERENCES "Vendedores"("ID_Vendedor")
);
CREATE TABLE IF NOT EXISTS "Clientes" (
	"CodigoCliente"	TEXT UNIQUE,
	"NombreFiscal"	TEXT NOT NULL,
	"CIF"	TEXT,
	"Direccion"	TEXT,
	"Poblacion"	TEXT,
	"Provincia"	TEXT,
	"Telefono"	TEXT,
	"Email"	TEXT,
	"ID_FormaPago"	INTEGER,
	"ID_Vendedor"	INTEGER,
	"Activo"	INTEGER DEFAULT 1,
	"CodigoPostal"	TEXT,
	"NombreComercial"	TEXT,
	"PersonaContacto"	TEXT,
	"SitioWeb"	TEXT,
	"Observaciones"	TEXT,
	"ID_Ruta"	INTEGER,
	PRIMARY KEY("CodigoCliente"),
	FOREIGN KEY("ID_Vendedor") REFERENCES "Vendedores"("ID_Vendedor"),
	FOREIGN KEY("ID_FormaPago") REFERENCES "FormasPago"("ID_FormaPago")
);
CREATE TABLE IF NOT EXISTS "LineasAlbaran" (
	"ID_Linea"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"NumeroAlbaran"	TEXT NOT NULL,
	"NumeroOrden"	INTEGER,
	"ID_Articulo"	INTEGER,
	"Descripcion"	TEXT,
	"CantidadServida"	NUMERIC,
	"PrecioUnitario"	NUMERIC,
	"Descuento"	NUMERIC,
	"Total"	NUMERIC,
	"PorcentajeIVA"	REAL DEFAULT 21,
	FOREIGN KEY("NumeroAlbaran") REFERENCES "Albaranes"("NumeroAlbaran") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "LineasPedido" (
	"ID_Linea"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"NumeroPedido"	TEXT NOT NULL,
	"NumeroOrden"	INTEGER,
	"ID_Articulo"	TEXT,
	"Descripcion"	TEXT,
	"PrecioUnitario"	NUMERIC,
	"Cantidad"	REAL,
	"Descuento"	NUMERIC,
	"Total"	NUMERIC,
	"PorcentajeIVA"	REAL DEFAULT 21,
	FOREIGN KEY("NumeroPedido") REFERENCES "Pedidos"("NumeroPedido") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "LineasPresupuesto" (
	"ID_Linea"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"NumeroPresupuesto"	TEXT NOT NULL,
	"NumeroOrden"	INTEGER,
	"ID_Articulo"	INTEGER,
	"Descripcion"	TEXT,
	"Cantidad"	REAL DEFAULT 1,
	"PrecioUnitario"	REAL DEFAULT 0,
	"Descuento"	INTEGER,
	"Total"	NUMERIC DEFAULT 0,
	"PorcentajeIVA"	REAL DEFAULT 21,
	FOREIGN KEY("NumeroPresupuesto") REFERENCES "Presupuestos"("NumeroPresupuesto") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "TipoIva" (
	"CodigoIva"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"Descripcion"	TEXT DEFAULT 0,
	"Porcentaje"	REAL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS "LineasFacturaCompra" (
	"ID_LineaFactura"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"ID_FacturaCompra"	INTEGER NOT NULL,
	"NumeroOrden"	INTEGER,
	"ID_Articulo"	INTEGER,
	"Descripcion"	TEXT,
	"Cantidad"	REAL DEFAULT 1,
	"PrecioCosteReal"	REAL DEFAULT 0,
	"Descuento"	REAL DEFAULT 0,
	FOREIGN KEY("ID_Articulo") REFERENCES "Articulos"("ID_Articulo"),
	FOREIGN KEY("ID_FacturaCompra") REFERENCES "FacturasCompra"("ID_FacturaCompra") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "FacturasCompra" (
	"ID_FacturaCompra"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"NumeroFacturaProveedor"	TEXT NOT NULL,
	"FechaEmision"	DATE DEFAULT CURRENT_DATE,
	"FechaVencimiento"	DATE,
	"ID_Proveedor"	INTEGER NOT NULL,
	"ID_AlbaranCompra"	INTEGER,
	"BaseImponible"	REAL DEFAULT 0,
	"TotalPagar"	REAL DEFAULT 0,
	"Pagada"	INTEGER DEFAULT 0,
	FOREIGN KEY("ID_Proveedor") REFERENCES "Proveedores"("ID_Proveedor"),
	FOREIGN KEY("ID_AlbaranCompra") REFERENCES "AlbaranesCompra"("ID_AlbaranCompra")
);
CREATE TABLE IF NOT EXISTS "MovimientosAlmacen" (
	"ID_Movimiento"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"Fecha"	DATETIME DEFAULT CURRENT_TIMESTAMP,
	"ID_Articulo"	INTEGER NOT NULL,
	"TipoMovimiento"	TEXT NOT NULL,
	"Cantidad"	REAL NOT NULL,
	"StockResultante"	REAL,
	"DocumentoReferencia"	TEXT,
	"ID_Usuario"	INTEGER,
	FOREIGN KEY("ID_Articulo") REFERENCES "Articulos"("ID_Articulo")
);
CREATE TABLE IF NOT EXISTS "LineasAlbaranCompra" (
	"ID_Linea"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"ID_AlbaranCompra"	INTEGER NOT NULL,
	"NumeroOrden"	INTEGER,
	"ID_Articulo"	INTEGER,
	"CantidadRecibida"	REAL,
	FOREIGN KEY("ID_AlbaranCompra") REFERENCES "AlbaranesCompra"("ID_AlbaranCompra") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "AlbaranesCompra" (
	"ID_AlbaranCompra"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"NumeroAlbaranProveedor"	TEXT NOT NULL,
	"ID_Proveedor"	INTEGER NOT NULL,
	"Fecha"	DATE DEFAULT CURRENT_DATE,
	"BaseImponible"	REAL DEFAULT 0,
	"Total"	REAL DEFAULT 0,
	FOREIGN KEY("ID_Proveedor") REFERENCES "Proveedores"("ID_Proveedor")
);
CREATE TABLE IF NOT EXISTS "LineasPedidoCompra" (
	"ID_Linea"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"ID_PedidoCompra"	INTEGER NOT NULL,
	"NumeroOrden"	INTEGER,
	"ID_Articulo"	INTEGER,
	"CantidadSolicitada"	REAL,
	FOREIGN KEY("ID_PedidoCompra") REFERENCES "PedidosCompra"("ID_PedidoCompra") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "PedidosCompra" (
	"ID_PedidoCompra"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"NumeroSerie"	TEXT,
	"ID_Proveedor"	INTEGER NOT NULL,
	"Estado"	TEXT DEFAULT 'Pendiente',
	"BaseImponible"	REAL DEFAULT 0,
	"Total"	REAL DEFAULT 0,
	FOREIGN KEY("ID_Proveedor") REFERENCES "Proveedores"("ID_Proveedor")
);
CREATE TABLE IF NOT EXISTS "Agencias" (
	"ID_Agencia"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"Nombre"	TEXT NOT NULL,
	"Telefono"	TEXT,
	"WebSeguimiento"	TEXT,
	"Codigo"	TEXT,
	"Email"	TEXT,
	"Observaciones"	TEXT,
	"Activo"	INTEGER DEFAULT 1
);
CREATE TABLE IF NOT EXISTS "Vendedores" (
	"ID_Vendedor"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"Nombre"	TEXT NOT NULL,
	"DNI"	TEXT,
	"Telefono"	TEXT,
	"Email"	TEXT,
	"Comision"	REAL DEFAULT 0,
	"Activo"	INTEGER DEFAULT 1,
	"CodigoVendedor"	TEXT,
	"Observaciones"	TEXT,
	"Direccion"	TEXT,
	"Poblacion"	TEXT,
	"Provincia"	TEXT,
	"CP"	TEXT,
	"FechaAlta"	DATE
);
CREATE TABLE IF NOT EXISTS "FormasPago" (
	"ID_FormaPago"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"Descripcion"	TEXT NOT NULL,
	"NumeroVencimientos"	INTEGER DEFAULT 1,
	"DiasPrimerVencimiento"	INTEGER DEFAULT 0,
	"DiasEntreVencimientos"	INTEGER DEFAULT 30,
	"Codigo"	TEXT,
	"Activo"	INTEGER DEFAULT 1
);
CREATE TABLE IF NOT EXISTS "Familias" (
	"ID_Familia"	INTEGER PRIMARY KEY AUTOINCREMENT,
	"Codigo"	TEXT UNIQUE,
	"Nombre"	TEXT NOT NULL,
	"Descripcion"	TEXT,
	"Activo"	INTEGER DEFAULT 1
);
CREATE TABLE IF NOT EXISTS "Empresa" (
	"ID"	INTEGER,
	"NombreFiscal"	TEXT,
	"CIF"	TEXT,
	"Direccion"	TEXT,
	"Telefono"	TEXT,
	"Logo"	BLOB,
	"Email"	TEXT,
	"Poblacion"	TEXT,
	"CodigoPostal"	TEXT,
	PRIMARY KEY("ID")
);
COMMIT;
