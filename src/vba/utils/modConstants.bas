Attribute VB_Name = "modConstants"

' RetailWeb / IE
Public Const IE_WINDOW_NAME As String = "Internet Explorer"

Public Const SB_ID_WAITPANE As String = "waitpane"
Public Const SB_CLASS_WAITING_PANEL As String = "panel waiting-panel"
Public Const SB_ID_SCROLL_BODY As String = "scroll_body"

Public Const SB_CLASS_PULL_LEFT As String = "pull-left"
Public Const SB_CLASS_BUTTON_DEFAULT_SM As String = "btn btn-default btn-sm "
Public Const SB_CLASS_BUTTON_DEFAULT_SM_PULL As String = "btn btn-default btn-sm pull-sm-right"
Public Const SB_CLASS_BUTTON_DEFAULT As String = "btn btn-default"
Public Const SB_CLASS_BUTTON_SUCCESS_SM_PULL As String = "btn btn-success btn-sm pull-sm-right"
Public Const SB_CLASS_INPUT_SM As String = "form-control input-sm "

Public Const SB_TEXT_BUSCAR As String = "Buscar"
Public Const SB_TEXT_PAGAR As String = "Pagar"
Public Const SB_TEXT_ACEPTAR As String = "Aceptar"
Public Const SB_TEXT_CAMBIAR_ESTADO_PAGO As String = "Cambiar Estado Pago"
Public Const SB_TEXT_IMPRIMIR_FACTURA As String = "Imprimir Factura"
Public Const SB_TEXT_CONTROL_INVENTARIOS As String = "Control de Inventarios"
Public Const SB_TEXT_CONTROL_RECEPCIONES As String = "Control de Recepciones Pagadas"
Public Const SB_TEXT_SISTEMA As String = "Sistema"
Public Const SB_TEXT_REPORTES As String = "Reportes para Descargar"

' Acciones UI
Public Const ACTION_ABRIR_RETAILWEB As String = "AbrirRetailWeb"
Public Const ACTION_IMPRIMIR_FACTURA As String = "ImprimirFactura"
Public Const ACTION_CAMBIAR_ESTADO As String = "CambiarEstado"
Public Const ACTION_PAGAR_FACTURA As String = "PagarFactura"
Public Const ACTION_CAMBIAR_PAGAR As String = "CambiarPagar"

' Estados
Public Const ESTADO_CONTABILIZADO As String = "Contabilizado"
Public Const ESTADO_COMPLETAR As String = "Completar"
Public Const ESTADO_REVISAR_DATOS As String = "Revisar datos"
Public Const ESTADO_ELIMINADO As String = "Eliminado"
Public Const ESTADO_PENDIENTE_REINGRESO As String = "Pendiente de Reingreso"
Public Const ESTADO_PENDIENTE_REVISAR As String = "Pendiente de revisar por negocio"
Public Const ESTADO_ERROR_SCAN As String = "Error de Scan"
Public Const ESTADO_VALIDAR As String = "Validar"
Public Const ESTADO_OK As String = "Ok"
Public Const ESTADO_MIGRAR_SAP As String = "Migrar SAP"
Public Const ESTADO_REMITO As String = "Remito"
Public Const ESTADO_DIF_COSTO As String = "Diferencia por Costo"
Public Const ESTADO_VALIDACION_AFIP_RECHAZADA As String = "ValidaciÃ³n AFIP rechazada"

' Mensajes
Public Const MSG_TIMEOUT_SB As String = "RetailWeb estÃƒÂ¡ tardando demasiado. Vuelva a intentarlo mÃƒÂ¡s tarde."
Public Const MSG_TIMEOUT_TITLE As String = "TimeOut"
