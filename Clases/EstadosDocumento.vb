' =====================================================================
'  EstadosDocumento
'  Constantes con todos los estados que pueden tener los documentos del
'  ERP. Antes estos estados estaban hardcodeados como strings sueltos en
'  cada formulario, así que un typo (p.ej. "Convertdo" en lugar de
'  "Convertido") rompía la magia de estados sin avisar.
'  ---------------------------------------------------------------------
'  Usando constantes:
'    - El compilador avisa de typos.
'    - Si mañana decides cambiar "Convertido" por "Tramitado", lo
'      tocas en un sitio.
'    - Es más fácil documentar qué estados maneja el ERP.
' =====================================================================
Public Module EstadosDocumento

    ' --- Presupuestos ---
    Public Const PRESUPUESTO_PENDIENTE As String = "Pendiente"
    Public Const PRESUPUESTO_ACEPTADO As String = "Aceptado"
    Public Const PRESUPUESTO_RECHAZADO As String = "Rechazado"
    Public Const PRESUPUESTO_CONVERTIDO As String = "Convertido" ' Cuando se ha pasado a pedido

    ' --- Pedidos de Venta ---
    Public Const PEDIDO_PENDIENTE As String = "Pendiente"
    Public Const PEDIDO_SERVIDO As String = "Servido"   ' Cuando ya hay un albarán hecho
    Public Const PEDIDO_CANCELADO As String = "Cancelado"

    ' --- Albaranes de Venta ---
    Public Const ALBARAN_PENDIENTE As String = "Pendiente"
    Public Const ALBARAN_ENVIADO As String = "Enviado"
    Public Const ALBARAN_ENTREGADO As String = "Entregado"
    Public Const ALBARAN_FACTURADO As String = "Facturado"   ' Cuando ya hay una factura hecha

    ' --- Facturas de Venta ---
    Public Const FACTURA_PENDIENTE As String = "Pendiente"
    Public Const FACTURA_COBRADA As String = "Cobrada"
    Public Const FACTURA_VENCIDA As String = "Vencida"
    Public Const FACTURA_CANCELADA As String = "Cancelada"

    ' --- Pedidos de Compra ---
    Public Const PEDIDO_COMPRA_PENDIENTE As String = "Pendiente"
    Public Const PEDIDO_COMPRA_ENVIADO As String = "Enviado al Proveedor"
    Public Const PEDIDO_COMPRA_PARCIAL As String = "Recibido Parcial"
    Public Const PEDIDO_COMPRA_RECIBIDO As String = "Recibido"

    ' --- Albaranes de Compra ---
    Public Const ALBARAN_COMPRA_PENDIENTE As String = "Pendiente"
    Public Const ALBARAN_COMPRA_RECIBIDO As String = "Recibido"
    Public Const ALBARAN_COMPRA_INCIDENCIA As String = "Recibido con incidencia"
    Public Const ALBARAN_COMPRA_FACTURADO As String = "Facturado"

    ' --- Facturas de Compra ---
    Public Const FACTURA_COMPRA_PENDIENTE As String = "Pendiente"
    Public Const FACTURA_COMPRA_PAGADA As String = "Pagada"
    Public Const FACTURA_COMPRA_VENCIDA As String = "Vencida"

    ' --- Tipos de Movimiento de Almacén ---
    Public Const MOV_ENTRADA As String = "ENTRADA"
    Public Const MOV_SALIDA As String = "SALIDA"
    Public Const MOV_AJUSTE_MAS As String = "AJUSTE+"
    Public Const MOV_AJUSTE_MENOS As String = "AJUSTE-"

End Module
