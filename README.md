# ⚡ Tabla con Superpoderes en Excel 365 (Script Lab & Office.js)

[![Licencia MIT](https://img.shields.io/badge/Licencia-MIT-green.svg?style=for-the-badge)](./LICENSE)
[![Excel 365](https://img.shields.io/badge/Microsoft_Excel-365-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/excel)
[![Office.js](https://img.shields.io/badge/API-Office.js_1.1+-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)](https://learn.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.0+-3178C6?style=for-the-badge&logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Script Lab](https://img.shields.io/badge/Microsoft-Script_Lab-D83B01?style=for-the-badge)](https://github.com/OfficeDev/script-lab)

> **Universal Table Manager**: Complemento web moderno que dota a cualquier tabla de Microsoft Excel de superpoderes interactivos: búsqueda instantánea multicolumna, ordenamiento dinámico estilo Power BI, formularios adaptativos autogenerados con validación de tipos, protección de fórmulas calculadas y generación de reportes ejecutivos con un solo clic.

---

## 🎥 Video Tutorial Paso a Paso

Aprende a construir y utilizar este motor interactivo paso a paso en nuestro canal de YouTube:

👉 **[Ver Video en YouTube - HAZ CLIC AQUÍ](https://www.youtube.com/watch?v=XXXXXXXXXXX&sub_confirmation=1)**

> 🔔 *¡No olvides dejar tu Like, comentar tus dudas y suscribirte al canal con la campanita para no perderte futuros tutoriales y plantillas avanzadas!*

---

## 🌐 Demo Interactiva en Vivo (Sin Abrir Excel)

Prueba el complemento funcionando en tiempo real directamente en tu navegador web gracias a GitHub Pages:

👉 **[Ver Demo Interactiva en Vivo (GitHub Pages)](https://intef-dev.github.io/tabla-superpoderes-excel-scriptlab/demo/)**

*(Incluye simulación completa de la cuadrícula de Excel, selector de tablas, búsqueda y alta de registros con cálculo de fórmulas en vivo).*

---

## 📥 Descarga Rápida del Proyecto (ZIP)

Si prefieres descargar todo el código y la plantilla de Excel en un solo archivo comprimido:

[![Descargar ZIP](https://img.shields.io/badge/Descargar-Proyecto_Completo_ZIP-blue?style=for-the-badge&logo=github)](https://github.com/intef-dev/tabla-superpoderes-excel-scriptlab/archive/refs/heads/main.zip)

---

## 🚀 Los 6 Superpoderes de la Aplicación

1. **Adaptabilidad Universal a Cualquier Tabla:**  
   No requiere configurar nombres fijos de columnas. La aplicación inspecciona `context.workbook.tables`, detecta encabezados y se adapta a cualquier estructura.
2. **Buscador Multicolumna & Ordenamiento Power BI:**  
   Búsqueda instantánea en tiempo real que resalta coincidencias en amarillo (`<mark>`) y ordenamiento ascendente/descendente numérico, cronológico y alfabético.
3. **Selector Dinámico de Columnas al Vuelo:**  
   Oculta o muestra columnas en la vista tabular mediante un popover con checkboxes, manteniendo protegida la columna ID obligatoria.
4. **Formularios Dinámicos con Tipado Automático:**  
   Construye campos HTML (`type="date"`, `type="number"`, `type="email"`, `<select>`) según el contenido de cada columna y autoincrementa identificadores (`ID-101`, `EMP-101`).
5. **Protección y Propagación Automática de Fórmulas:**  
   Las columnas calculadas en Excel se muestran bloqueadas con candado (`🔒`) en el formulario para evitar alteraciones. Al insertar una nueva fila, la fórmula se hereda automáticamente de la fila anterior mediante `copyFrom(..., RangeCopyType.formulas)`.
6. **Generador de Reportes Ejecutivos:**  
   Crea una nueva pestaña en el libro, le asigna color corporativo índigo, escribe títulos ejecutivos con fecha y genera una tabla oficial con autoajuste de columnas.

---

## 📋 Requisitos

- **Microsoft Excel 365** (en Windows, Mac o Excel en la Web).
- El complemento gratuito oficial de Microsoft **Script Lab** (búscalo en Excel $\rightarrow$ Pestaña *Insertar* $\rightarrow$ *Obtener complementos* $\rightarrow$ Buscar "Script Lab").

---

## 🛠️ Cómo Utilizar este Proyecto en Script Lab

### Opción 1: Importar en 1 Clic usando `snippet.yaml` (Recomendada)
1. Abre **Excel 365** y haz clic en la pestaña **Script Lab** $\rightarrow$ **Code**.
2. En el panel lateral, haz clic en el menú superior (**☰**) y selecciona **Import** (Importar).
3. Selecciona la opción **Import from YAML** (o pega el contenido del archivo [`snippet.yaml`](./snippet.yaml)).
4. Pulsa **Run** en el panel superior y ¡listo!

### Opción 2: Copiar Pestaña por Pestaña
1. Abre **Script Lab** y haz clic en **New Snippet** (Nuevo proyecto).
2. Copia el contenido de los siguientes archivos en sus respectivas pestañas:
   - **HTML**: Copia todo el contenido de [`html.html`](./html.html) y pégalo en la pestaña **HTML**.
   - **CSS**: Copia todo el contenido de [`css.css`](./css.css) y pégalo en la pestaña **CSS**.
   - **Script (TypeScript)**: Copia todo el contenido de [`script.ts`](./script.ts) y pégalo en la pestaña **Script**.
3. Haz clic en **Run** en la barra superior.

---

## 📂 Estructura del Repositorio

```text
tabla-superpoderes-excel-scriptlab/
│
├── README.md                                          # Documentación oficial del proyecto
├── LICENSE                                            # Licencia MIT
├── .gitignore                                         # Exclusiones de control de versiones
│
├── snippet.yaml                                       # Snippet completo para importación directa en Script Lab
├── script.ts                                          # Código fuente TypeScript con Office.js y CRUD
├── html.html                                          # Markup HTML5 del panel lateral y modales
├── css.css                                            # Hoja de estilos moderna estilo SaaS
│
├── Plantilla_Superpoderes_Excel_TablaPersonal.xlsx    # Libro de Excel de muestra con datos y fórmulas
│
└── demo/                                              # Demo interactiva para GitHub Pages
    ├── index.html                                     # Interfaz de demostración interactiva
    ├── demo.css                                       # Estilos del marco de Excel y ventana
    └── demo.js                                        # Emulador en memoria de Office.js
```

---

## 📊 Plantilla de Excel de Demostración

El repositorio incluye la plantilla oficial [`Plantilla_Superpoderes_Excel_TablaPersonal.xlsx`](./Plantilla_Superpoderes_Excel_TablaPersonal.xlsx) que contiene 50 registros de muestra en la tabla **`TablaPersonal`** con:
- Columnas de datos reales: `Código`, `Empleado`, `Departamento`, `Salario`, `Fecha Ingreso`.
- Columnas con fórmulas automáticas: `Bono (10%)` (`=[@Salario]*0.10`) y `Compensación Total` (`=[@Salario]+[@[Bono (10%)]]`).

---

## 📱 Comunidad y Redes Sociales

Para más tutoriales, código abierto y automatizaciones avanzadas de Excel y Office.js:

- 🎥 **YouTube:** [Suscríbete a Inteligencia Eficiente](https://www.youtube.com/watch?v=XXXXXXXXXXX&sub_confirmation=1)
- 📸 **Instagram:** [@inteligenciaeficiente](https://www.instagram.com/inteligenciaeficiente/)
- 👤 **Facebook:** [Inteligencia Eficiente Oficial](https://www.facebook.com/InteligenciaEficiente/)
- 🌐 **Sitio Web:** [inteficiente.com](https://inteficiente.com/)

---

## ☕ Apoya este Proyecto

Si este código te ha sido de utilidad y deseas apoyar la creación de contenido educativo y herramientas gratuitas:

- 💳 **PayPal:** [Donar a través de PayPal.Me](https://www.paypal.com/paypalme/INTEFDonativo)
- ☕ **Buy Me a Coffee:** [Apoyar en Buy Me a Coffee](https://buymeacoffee.com/inteligenciaeficiente)

---

## 📄 Licencia

Distribuido bajo la licencia **MIT**. Consulta el archivo [`LICENSE`](./LICENSE) para más información. Libre para uso personal, educativo y comercial.
