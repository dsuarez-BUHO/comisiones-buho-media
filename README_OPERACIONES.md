# 📘 Guía de Operación — Sistema Central de Comisiones (Buho Media)

Documento para el equipo de **Finanzas / Operaciones**. No requiere conocimientos
técnicos: explica cómo dejar los archivos en Google Drive para que la aplicación web
calcule las comisiones correctamente.

---

## 1. Cómo funciona (en 3 pasos)

1. Tú subes los 4 Excel del período a la **carpeta compartida de Google Drive**.
2. Entras a la aplicación web, eliges el **rango de fechas** y presionas **Procesar Comisiones**.
3. Descargas el **Resumen Global** y, si lo necesitas, el reporte de cada vendedor.

La app **siempre toma la versión más reciente** de cada archivo en la carpeta. No necesitas
borrar los anteriores, pero sí respetar las reglas de nombre de abajo.

---

## 2. 🏷️ Regla de oro: nomenclatura de los archivos

Cada archivo **debe empezar con su prefijo exacto** y terminar en **`.xlsx`**.
El texto después del prefijo es libre (ideal: el mes), y sirve para identificar el período.

| Insumo | El nombre **debe empezar con** | Ejemplos válidos |
| :--- | :--- | :--- |
| 📋 Margen Tarifario | `Margen_Tarifario` | `Margen_Tarifario.xlsx`, `Margen_Tarifario_Julio.xlsx` |
| 📦 Catálogo BWMS | `Catalogo BWMS` | `Catalogo BWMS.xlsx`, `Catalogo BWMS 2025.xlsx` |
| 💵 Concentrado de Facturas | `Concentrado_Facturas` | `Concentrado_Facturas_Pagadas_Julio.xlsx` |
| 🏦 Asiento Contable | `Asiento contable` | `Asiento contable Julio.xlsx` |

> ⚠️ Si un archivo no empieza con su prefijo exacto, la app no lo encontrará y mostrará un
> aviso indicando cuál falta.

---

## 3. 🧩 Qué contiene cada archivo y por qué importa

| Insumo | Rol en el negocio | Si falta o está mal… |
| :--- | :--- | :--- |
| **📋 Margen Tarifario** | Asocia cada cliente con su KAM (vendedor) y su tasa de comisión según el plan. | No se puede asignar ninguna factura a un vendedor: las comisiones salen vacías. |
| **📦 Catálogo BWMS** | Respaldo: recupera clientes nuevos que aún no están en el Margen Tarifario. | Se pueden omitir comisiones legítimas de cuentas nuevas o en migración. |
| **💵 Concentrado de Facturas** | Aporta los montos facturados para calcular la base sin IVA (÷1.16). | No hay importes que procesar en el período. |
| **🏦 Asiento Contable** | Aporta la fecha real de cobro para aplicar las penalizaciones por cobranza tardía. | El sistema se detiene para no pagar comisiones sobre facturas no cobradas. |

Las **columnas** de cada Excel no deben renombrarse ni eliminarse. Si cambian los
encabezados, la app avisará qué columna falta.

---

## 4. ✅ Buenas prácticas en la carpeta de Drive

- **No quites el acceso** de la cuenta de servicio
  `buho-comisiones@reporte-campana.iam.gserviceaccount.com` a la carpeta.
- **Cierra el Excel antes de subirlo.** Un archivo abierto genera un temporal `~$…`; la app
  ya los ignora, pero es mejor evitarlos.
- **Sube versiones nuevas** con el mismo prefijo (agregando el mes); no hace falta renombrar
  las viejas.
- **No muevas los archivos a subcarpetas**: la app solo mira el nivel principal de la carpeta.
- Sube **un solo archivo vigente por insumo** (o deja claro cuál es el más reciente): la app
  elige por fecha de modificación.

---

## 5. 🆘 Si la app muestra un error

| Mensaje | Qué revisar |
| :--- | :--- |
| *"No se encontró un archivo que empiece con …"* | El prefijo del nombre. Corrígelo según la tabla de la sección 2. |
| *"A '…' le faltan columnas obligatorias: […]"* | Alguna columna fue renombrada/eliminada. Restaura los encabezados originales. |
| *"El archivo '…' está vacío."* | El Excel no tiene datos; vuelve a exportarlo. |
| *"No se generaron comisiones para el período…"* | No hay facturas **pagadas** en el rango elegido, o no hacen match con el catálogo. Revisa fechas y nombres de cliente. |
| *"No se pudo conectar con Google Drive…"* | Verifica que la carpeta siga compartida con la cuenta de servicio. |

Ante cualquier duda persistente, contacta al responsable técnico del sistema.
