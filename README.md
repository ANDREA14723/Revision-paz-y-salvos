# Revisión de paz y salvos (FO-AC-101 / FO-AC-103) contra el FO-AC-83

## Instalación
1. Python 3.10+ y `pip install -r requirements.txt`
2. **Tesseract OCR** (necesario para las páginas escaneadas) con el idioma español:
   - Windows: instalador de UB Mannheim; marcar "Spanish" en idiomas. Si no queda en el PATH, escribe la
     ruta de `tesseract.exe` en la barra lateral.
   - Linux: `sudo apt install tesseract-ocr tesseract-ocr-spa`
   Sin el paquete `spa` funciona con inglés, con más errores en tildes y ñ.

## Uso
`streamlit run app.py` → sube el PDF y (opcional) el FO-AC-83 → **Revisar**.

## Qué verifica en cada documento
- **Estructura**: formato identificado y frases clave de la plantilla (laboral y continuidad). No se valida el número de contrato ni el objeto: la app sirve para varios contratos.
- **Asunto**: LABORAL, OPS riesgo I-III y OPS riesgo IV-V llevan su asunto propio en FO-AC-101; la carta FO-AC-103 vale para cualquier vinculación y riesgo.
- **Datos**: nombre, cédula (y su formato de puntos), cargo, fechas del periodo, empresa.
- **Coherencia interna**: nombre, cédula y cargo del texto contra el bloque de firma; fecha del encabezado contra
  la fecha de firma; periodo invertido; carta de continuidad con fecha de desvinculación ≠ fin del paz y salvo.
- **Cruce con FO-AC-83** (por cédula; si la persona tiene varias filas —otros cargos o periodos— se empareja cada certificado con la fila de su mismo periodo, cargo y empresa): nombre, cargo, fecha de ingreso = inicio, fecha de retiro = fin, empresa,
  sección (RETIRO), tipo de vinculación (LABORAL / OPS) y clase de riesgo ARL.
- **Global**: cédulas repetidas y personas en RETIRO del FO-AC-83 sin paz y salvo en el PDF.

Niveles: ❌ ERROR (dato que no coincide o no se pudo leer), ⚠️ ALERTA (diferencia menor o dato a confirmar).

## Límites
- No valida la firma: no distingue una firma manuscrita real de un espacio en blanco. Quedó para revisión visual
  (cada documento tiene botón "Ver página").
- En páginas escaneadas el OCR puede equivocarse en una letra o un dígito; el detalle lo avisa.
- Las reglas de vinculación y riesgo de la barra lateral son configurables; los valores por defecto son
  supuestos, no una norma del IDU.

- Las comparaciones ignoran tildes, ñ, mayúsculas, espacios dobles y guiones. Si el texto del PDF viene incompleto, la app relee la página y, si hace falta, usa OCR.
