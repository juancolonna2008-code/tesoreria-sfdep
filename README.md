# Tesorería SFDEP

Sistema de cuotas de la Sociedad de Fomento para las actividades La Vieja y Palomitas.

- **Resumen:** lo cobrado en el mes, el reparto 70% actividad / 30% Sociedad de Fomento y lo que falta cobrar.
- **Cobranza del mes:** lista por plantel o categoría para marcar Pagó, Debe, No vino o Bonificado.
- **Jugadores:** padrón con ficha, historial de pagos, cuota especial y bajas.
- **Deudas:** quién debe y qué meses, con la lista para copiar.
- **Ajustes:** valor de la cuota por mes, porcentaje del reparto, actividades y planillas para descargar.

## Cómo funciona

- La página (`index.html`) se publica con GitHub Pages.
- Los datos viven en una planilla de Google Drive del tesorero, no en este repositorio.
- La página habla con la planilla a través de un script de Google Apps Script (`apps-script/Codigo.gs`), publicado como app web.
- Para entrar hace falta la clave de la comisión. La clave está en la pestaña **Config**, celda **B2**, de la planilla. Para cambiarla, se escribe la nueva ahí y se avisa a la comisión.

## Archivos

- `index.html`: la página completa, con el escudo incluido.
- `escudo.png`: el escudo del club.
- `apps-script/Codigo.gs`: el código del script de Google. Si se cambia, hay que pegarlo en el proyecto "Tesorería SFDEP - Conexión" y crear una nueva implementación.
