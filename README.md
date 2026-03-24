# 🚀 OPTIMA - ERP Framework para PyMEs

### *Proyecto de Fin de Grado (TFG) - Desarrollo de Aplicaciones Multiplataforma*

**Autor:** Paul Andrei  
**Entorno:** Visual Studio 2022  
**Lenguaje:** Visual Basic .NET (`VB.NET`)  
**Base de Datos:** SQLite (Local & Portable)  
**Interfaz:** Windows Forms con Custom UX/UI  

---

## 📝 Descripción del Proyecto

**OPTIMA** es un sistema de planificación de recursos empresariales (ERP) diseñado para digitalizar el ciclo comercial completo de una pequeña empresa. El proyecto nace bajo la premisa de ofrecer una herramienta robusta, ligera y fácil de usar, eliminando la complejidad de los grandes ERPs pero manteniendo la **trazabilidad documental** profesional.

El software gestiona desde el stock físico en almacén hasta la generación de facturas legales, incluyendo un motor de impresión propio y un sistema de cálculo de comisiones dinámico por vendedor.

---

## 📋 Módulos del Sistema

### 💰 Ciclo de Ventas e Ingresos
* **Flujo Documental:** Conversión inteligente de documentos (`Presupuesto` ➔ `Pedido` ➔ `Albarán` ➔ `Factura`) con arrastre automático de líneas, descuentos e impuestos.
* **Facturación:** Emisión de facturas con cálculo automático de bases imponibles y cuotas de IVA.
* **Logística:** Gestión de bultos, pesos y agencias de transporte en el proceso de expedición (Albaranes).

### 👥 Gestión de Terceros
* **Fichas de Clientes y Proveedores:** Centralización de datos fiscales, contactos y condiciones de pago.
* **Fuerza de Ventas:** Gestión de vendedores con asignación de porcentajes de comisión individuales.

### 📦 Almacén e Inventario
* **Control de Stock:** Actualización automática de existencias al validar documentos de salida.
* **Clasificación:** Organización de catálogo por Familias y Artículos.
* **Kardex:** Histórico de movimientos para auditoría de entradas y salidas.

### 📊 Business Intelligence (Reporting)
* **Informes PDF:** Generación de reportes profesionales mediante `GDI+` y `PrintDocument`.
* **Ranking de Ventas:** Visualización de los mejores clientes por volumen de facturación.
* **Liquidación de Comisiones:** Informe detallado de ventas por comercial y su respectiva comisión devengada.

---

## 🛠️ Hitos Técnicos y Desafíos Superados

* **Optimización de UI (Anti-Flickering):** Implementación de llamadas a la API nativa de Windows (`user32.dll`) mediante `SendMessage` y el parámetro `WM_SETREDRAW` para eliminar el parpadeo durante la carga dinámica de formularios.
* **Arquitectura de Impresión:** Desarrollo de una clase de impresión personalizada que gestiona paginación automática, dibujo de tablas dinámicas y renderizado de logotipos corporativos almacenados como BLOB en la base de datos.
* **Persistencia Eficiente:** Uso de **SQLite** como motor de base de datos, permitiendo que la aplicación sea 100% portable y no requiera de la instalación de servidores pesados (SQL Server/MySQL).
* **Diseño Responsive Interno:** Sistema de gestión de formularios hijos (`MDI-style` manual) con anclajes (`Docking` y `Padding`) para asegurar que la interfaz se adapte al tamaño del panel contenedor.

---

## 🏗️ Estructura del Proyecto

```text
OPTIMA/
├── 📂 Clases/           # Capa de Lógica y Datos
│   ├── ConexionBD.vb    # Singleton de conexión a SQLite
│   ├── GestorDatos.vb   # Funciones CRUD y consultas SQL
│   └── ComunSesion.vb   # Persistencia de usuario en sesión
├── 📂 Formularios/      # Capa de Presentación (UI)
│   ├── 📂 Ventas        # Presupuestos, Pedidos, Albaranes, Facturas
│   ├── 📂 Terceros      # Clientes, Proveedores, Vendedores
│   ├── 📂 Almacen       # Artículos, Familias, Movimientos
│   └── 📂 Informes      # Ranking, Comisiones, Listados IVA
└── 📂 Database/         # Esquema y archivos de base de datos
