# AgroControl 🚜 - Sistema de Gestión de Órdenes Agrícolas

![Next.js](https://img.shields.io/badge/Next.js-14-black?style=flat&logo=next.js)
![TypeScript](https://img.shields.io/badge/TypeScript-5.0-blue?style=flat&logo=typescript)
![Tailwind CSS](https://img.shields.io/badge/Tailwind-CSS-3.0-38bdf8?style=flat&logo=tailwind-css)
![ExcelJS](https://img.shields.io/badge/ExcelJS-Export-green)

**AgroForm** es una aplicación web profesional (PWA) diseñada para automatizar la generación de Órdenes de Trabajo agrícolas. Reemplaza el llenado manual de planillas de Excel por un sistema digital inteligente que calcula dosis, sincroniza tablas de coadyuvantes y genera archivos `.xlsx` con formato perfecto para impresión o envío.

---

## 🚀 Características Principales

* **Generación de Excel "Pixel-Perfect":** Utiliza una plantilla base (`plantilla.xlsx`) para respetar logotipos, colores, celdas combinadas y formatos originales de la empresa.
* **Sistema de Doble Tabla:**
    * **Tabla Principal:** Gestión de 5 productos principales (Columnas C a G) con cálculo de Dosis x Hectárea.
    * **Tabla de Coadyuvantes:** Gestión independiente de 3 coadyuvantes (Columnas H a J) con nombres personalizables.
* **Cálculos Inteligentes (Total Real):**
    * El sistema suma automáticamente las filas de resultados (celdas grises) para dar el total exacto de litros/kilos a comprar, ignorando las dosis unitarias.
    * Suma automática de hectáreas totales.
* **Interfaz Responsiva:** Diseñada con **Tailwind CSS** para un uso fluido tanto en PC como en dispositivos móviles en el campo.
* **Seguridad de Datos:** Fórmulas de Excel incrustadas (`IF(ISNUMBER...)`) para evitar errores como `#VALUE!` o `NaN` en el archivo final.

---

## 🛠️ Tecnologías

* **Framework:** [Next.js 14](https://nextjs.org/) (App Router)
* **Lenguaje:** [TypeScript](https://www.typescriptlang.org/)
* **Estilos:** [Tailwind CSS](https://tailwindcss.com/) + [shadcn/ui](https://ui.shadcn.com/)
* **Motor Excel:** [ExcelJS](https://github.com/exceljs/exceljs)
* **Iconos:** [Lucide React](https://lucide.dev/)

---

## 💼 ¿Necesitas este sistema para tu empresa?

Si te interesa implementar este software pero utilizas una planilla de Excel diferente, contáctame.

Puedo adaptar el código para que funcione 100% a medida con el formato, logotipos y cálculos específicos de tu campo o empresa.

📩 Contacto: genaroarcee@gmail.com

---

## 📄 Licencia

- Software desarrollado para gestión agricola profesional.
- © Genaro Arce