# 🚀 OPTIMA - Sistema de Gestión Empresarial (ERP)

### *Trabajo de Fin de Grado (TFG) - Desarrollo de Aplicaciones Multiplataforma / Web*

**Desarrollado por:** Paul Andrei  
**Entorno de Desarrollo:** Visual Studio 2022  
**Lenguaje:** Visual Basic .NET (`VB.NET`)  
**Tecnología de Interfaz:** Windows Forms (`WinForms`)  
**Base de Datos:** [Añade aquí tu base de datos, ej: SQL Server / MySQL]

---

## 📝 Descripción del Proyecto

OPTIMA es una solución ERP integral diseñada para centralizar y optimizar los flujos de trabajo operativos de una pyme. Este proyecto cubre el ciclo completo de negocio, desde la gestión de inventario y la relación con terceros, hasta el proceso completo de ventas y facturación.

El software ha sido desarrollado siguiendo una arquitectura modular para garantizar la escalabilidad y facilitar el mantenimiento futuro del código, cumpliendo con los requisitos académicos y técnicos de un TFG.

---

## 📋 Características Principales por Módulo

El software se divide en módulos funcionales bien definidos que cubren las áreas críticas de la empresa:

### 💰 Módulo de Ventas
Gestión integral del flujo de ingresos, permitiendo la trazabilidad completa desde el primer contacto comercial hasta el cobro definitivo.
* **📋 Presupuestos:** Creación, seguimiento y envío de ofertas comerciales a clientes potenciales.
* **🛒 Pedidos:** Formalización y registro de las solicitudes de compra confirmadas.
* **📦 Albaranes:** Control de la logística de entrega de mercancía y generación de documentos de portes.
* **🧾 Facturación:** Emisión de facturas legales, registro contable automatizado y gestión de vencimientos.

### 👥 Módulo de Terceros
Centralización de la información y gestión de las relaciones con los actores externos clave.
* **👥 Clientes:** Fichas detalladas con datos de contacto, condiciones de pago, direcciones de envío y riesgo.
* **🏭 Proveedores:** Gestión de la base de datos de proveedores para el módulo de compras (en proceso).

### 📦 Módulo de Almacén
Control y supervisión del inventario físico en tiempo real.
* **🏷️ Artículos y Familias:** Clasificación jerárquica y parametrización detallada de productos.
* **🔄 Movimientos de Almacén:** Registro histórico de todas las entradas, salidas y regularizaciones de stock.

### ⚙️ Tablas Maestras
Configuración y parametrización global del sistema para su correcto funcionamiento.
* **🕴️ Vendedores y Rutas:** Gestión del equipo comercial y organización logística de entregas.
* **💳 Formas de Pago:** Definición de los métodos y plazos de cobro/pago soportados.
* **🚚 Agencias de Transporte:** Mantenimiento de las empresas de logística colaboradoras.

---

## 🖼️ Interfaz del Proyecto (Capturas de Pantalla)

Esta sección muestra la identidad visual y la usabilidad de la aplicación WinForms.

### Acceso y Panel Principal
> Diseño limpio centrado en la usabilidad del operario.

| Pantalla de Acceso (Login) | Panel de Control Principal (Dashboard) |
| :---: | :---: |
| <img src="https://github.com/user-attachments/assets/77fb8e18-4aaf-4387-a041-4889babe8098" width="350" alt="Login OPTIMA"/> | <img src="https://github.com/user-attachments/assets/f58d9b6c-bc9d-4f72-ad34-cd482ba8b49e" width="550" alt="Dashboard OPTIMA"/> |

---

### Flujo del Módulo de Ventas
> Demostración visual del ciclo documental de una venta.

<details>
<summary>📸 Haz clic aquí para ver las capturas de Ventas</summary>

| Documento | Captura de Pantalla |
| :--- | :---: |
| **1. Presupuesto** | <img src="https://github.com/user-attachments/assets/fc079f8d-c203-4218-a03a-4ea4c9bcde26" width="100%" alt="Presupuesto OPTIMA"/> |
| **2. Pedido** | <img src="https://github.com/user-attachments/assets/8c16d0db-6ddd-4cbc-aaa0-7521fe922770" width="100%" alt="Pedido OPTIMA"/> |
| **3. Albarán** | <img src="https://github.com/user-attachments/assets/83ade9db-ea03-413b-8d2d-9773552ffa47" width="100%" alt="Albarán OPTIMA"/> |
| **4. Factura** | <img src="https://github.com/user-attachments/assets/5f9a9cfc-d09b-4a49-8291-5fe1e0c28f02" width="100%" alt="Factura OPTIMA"/> |

</details>

---

### Gestión de Entidades y Almacén

<details>
<summary>📸 Haz clic aquí para ver las capturas de Terceros y Stock</summary>

#### Terceros
| Ficha de Cliente | Ficha de Proveedor |
| :---: | :---: |
| <img src="https://github.com/user-attachments/assets/ca124c82-c7ef-40ec-8140-06868c25fe37" alt="Cliente OPTIMA"/> | <img src="https://github.com/user-attachments/assets/cea24bfd-a9e2-449c-bc52-9529e95e83bd" alt="Proveedor OPTIMA"/> |

#### Almacén
| Ficha de Artículo | Gestión de Familias | Movimientos de Stock |
| :---: | :---: | :---: |
| <img src="https://github.com/user-attachments/assets/d8865513-5720-40df-8a95-755c7c27e934" alt="Artículo OPTIMA"/> | <img src="https://github.com/user-attachments/assets/15347745-97e9-4237-ad3d-c237bf61c712" width="300" alt="Familias OPTIMA"/> | <img src="https://github.com/user-attachments/assets/14d8d8a3-f479-4f4f-91a5-ed0aed6b50ea" alt="Movimientos OPTIMA"/> |

</details>

---

### Configuración (Tablas Maestras)

<details>
<summary>📸 Haz clic aquí para ver las capturas de Configuración</summary>

| Vendedores | Formas de Pago | Rutas | Agencias |
| :---: | :---: | :---: | :---: |
| <img src="https://github.com/user-attachments/assets/84f12cad-6975-4ade-86c1-d7126d08c197" width="200" alt="Vendedor OPTIMA"/> | <img src="https://github.com/user-attachments/assets/ded243bc-6062-4a7e-89fc-c8b6b685bbd4" width="200" alt="Pagos OPTIMA"/> | <img src="https://github.com/user-attachments/assets/d5603a33-7057-412c-b674-de400733cbbf" width="200" alt="Rutas OPTIMA"/> | <img src="https://github.com/user-attachments/assets/837d5c34-9ac9-46cb-b683-effd753c9d52" width="200" alt="Agencias OPTIMA"/> |

</details>

---

## 🛠️ Stack Tecnológico

Este proyecto ha sido desarrollado utilizando tecnologías robustas del ecosistema Microsoft:

* **IDE:** Visual Studio 2022 Professional.
* **Lenguaje:** Visual Basic .NET (`VB.NET`).
* **Framework de UI:** Windows Forms (`WinForms`).
* **Persistencia de Datos:** [RELLENA AQUÍ: ej. ADO.NET / Entity Framework] con [RELLENA AQUÍ: ej. SQL Server].

---

## 🏗️ Requisitos e Instalación

### Prerrequisitos
Para ejecutar o compilar este proyecto necesitas:
1.  **.NET Framework** (versión [Añade versión, ej: 4.8] o superior).
2.  Un gestor de base de datos [Añade cuál, ej: SQL Server LocalDB] instalado y configurado.

### Instalación
1.  Clona este repositorio: `git clone https://github.com/tu-usuario/OPTIMA-ERP.git`
2.  Abre la solución `OPTIMA.sln` en Visual Studio 2022.
3.  Ejecuta el script SQL ubicado en la carpeta `/database` para crear la estructura de datos.
4.  Compila y ejecuta el proyecto (`F5`).

---

---

*Este proyecto es propiedad intelectual de Paul Andrei y ha sido presentado como Trabajo de Fin de Grado.*
