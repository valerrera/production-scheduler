# 🏭 Production Scheduler — Herramienta de Planeación de Producción

Herramienta desarrollada en **Visual Basic for Applications (VBA)** sobre Microsoft Excel que permite ingresar parámetros de problemas de planeación de producción, definir secuencias de programación y visualizar resultados mediante **diagramas de Gantt** e **indicadores de desempeño**.

Desarrollada como material de apoyo para la asignatura **Planeación y Programación de las Operaciones** en la Pontificia Universidad Javeriana.

---

## 📋 Descripción

La herramienta cubre **8 modelos clásicos de programación de la producción**, cada uno en su propia hoja del libro de Excel. Para cada modelo, el usuario ingresa los parámetros del problema, define su decisión de secuenciación y genera automáticamente el diagrama de Gantt con sus indicadores de desempeño.

---

## ⚙️ Modelos implementados

| Modelo | Descripción |
|--------|-------------|
| **Single Machine** | Una sola máquina procesa todos los jobs. El estudiante decide el orden. |
| **Parallel Machine** | Varias máquinas idénticas en paralelo. El estudiante asigna cada job a una máquina y define el orden. |
| **Flow Shop** | Todos los jobs siguen la misma ruta fija (M1→M2→…→Mm). El estudiante decide la permutación. |
| **Job Shop** | Cada job tiene su propia ruta (parámetro). El estudiante decide la secuencia por máquina. |
| **Open Shop** | El estudiante decide tanto la ruta como la secuencia por máquina. |
| **Flexible Flow Shop** | Como Flow Shop, pero cada etapa puede tener c ≥ 1 máquinas idénticas. Asignación FIFO. |
| **Flexible Job Shop** | Como Job Shop, pero cada workstation puede tener c ≥ 1 máquinas idénticas. Asignación FIFO. |
| **Flexible Open Shop** | Como Open Shop, pero cada workstation puede tener c ≥ 1 máquinas idénticas. Asignación FIFO. |

---

## 📥 Parámetros de entrada

Cada hoja tiene una **zona de inputs** donde el usuario ingresa directamente en celdas de Excel:

| Parámetro | Significado |
|-----------|-------------|
| `r` | *Release time*: momento más temprano en que el job puede comenzar |
| `p` | *Processing time*: tiempo de procesamiento por operación |
| `s` | *Setup time*: tiempo de preparación antes de procesar |
| `d` | *Due date*: fecha de entrega comprometida |
| `w` | *Weight*: peso o prioridad del job |
| `rmaq` | *Release de máquina*: disponibilidad inicial de cada máquina |
| `c` | Número de máquinas idénticas por workstation (solo modelos Flexible) |

---

## 🖱️ Flujo de uso

Todas las hojas comparten el mismo flujo de tres pasos:

1. **Redibujar** — Se ejecuta automáticamente al cambiar la configuración. Limpia la hoja y dibuja las tablas vacías.
2. **Cargar datos** — Valida toda la información ingresada. Si hay errores, indica qué corregir.
3. **Generar outputs** — Calcula el programa, dibuja el Gantt y muestra los indicadores.

---

## 📊 Resultados generados

### Diagrama de Gantt
- Visualización gráfica de cada operación en el tiempo
- Colores claros para *setup* y oscuros para procesamiento
- En modelos Flexible: eje Y con máquinas individuales (M1.1, M1.2, M2.1, etc.)

### Indicadores globales

| Indicador | Definición |
|-----------|------------|
| Makespan (Cmax) | Máximo Cj entre todos los jobs |
| Tiempo de flujo promedio | Promedio de (Cj − rj) |
| Máximo Lateness (Lmax) | Máximo de (Cj − dj) |
| Tardanza media | Promedio de max(Cj − dj, 0) |
| Tardanza ponderada total (ΣwT) | Suma de wj × max(Cj − dj, 0) |
| # / % Jobs tarde | Cantidad y porcentaje de jobs con tardanza > 0 |

### Indicadores por job
Tabla con inicio, Cj, flujo, lateness, tardanza y wT para cada job. En modelos Flexible, incluye la asignación de máquina específica por workstation.

---

## 🚀 Cómo usar

1. Descargar o clonar el repositorio
2. Abrir el archivo `.xlsm` en Microsoft Excel con **macros habilitadas**
3. Ir a la hoja del modelo que deseas usar
4. Ingresar los parámetros del problema en las celdas indicadas
5. Presionar **Cargar datos** para validar
6. Presionar **Generar outputs** para obtener el Gantt e indicadores

> ⚠️ Al abrir el archivo, Excel mostrará una advertencia de seguridad sobre macros. Selecciona **"Habilitar contenido"** para que funcione correctamente.

> ⚠️ Al cambiar los valores de configuración global (# jobs, # máquinas), las tablas se redibujan y se pierde la información ingresada. Define primero la configuración.

---

## 🛠️ Tecnologías

- Microsoft Excel (`.xlsm`)
- Visual Basic for Applications (VBA)

---

## 👩‍💻 Autora

**Valeria Herrera**  
Estudiante de Ingeniería Industrial e Ingeniería de Sistemas  
Pontificia Universidad Javeriana — Bogotá, Colombia

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Valeria%20Herrera-0077B5?style=flat&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/valeriaherreraarboleda/)

---

## 📄 Contexto académico

Proyecto desarrollado en el marco de la asignatura **Planeación y Programación de las Operaciones**, donde también desempeñé el rol de **monitora académica**, apoyando a estudiantes en la comprensión de estos modelos y contribuyendo al desarrollo de esta herramienta como material de apoyo para la clase.
