---
proyecto:
  - Reporte Automatizado Power Query
tipo: Automatización
semana: "01"
tech-stack: Excel, Power Query
status:
  - Completado
tags:
  - proyecto
  - excel
  - powerquery
  - automation
  - semana01
  - completado
  - dashboard
---

# Proyecto 01 - Reporte Automatizado con Power Query

## Overview

**Descripción:** Sistema automatizado en Power Query que importa, limpia y combina 3 archivos CSV desde una carpeta para generar un reporte consolidado de ventas listo para análisis y dashboard.

**Objetivo:** Reducir el tiempo de generación del reporte semanal desde un proceso manual (~2 horas) a un refresh automático de segundos.

**Semana:** [[Semana-01]] - [[Power-Query]]
**Status:** Completado


## Tech Stack

- [x] Excel
- [x] Power Query
- [ ] Python
- [ ] SQL
- [ ] Power BI
- [ ] Otros: 


## Objetivos específicos

- [x] Importar 3 archivos CSV automáticamente
- [x] Limpiar datos (duplicados, tipos, nulls, espacios)
- [x] Combinar con Merge (3 tablas relacionadas)
- [x] Calcular columnas adicionales (Valor Total)
- [x] Generar tabla final lista para dashboard
- [x] Documentar pasos y código M
- [x] Probar automatización (Refresh funciona)


## Dataset utilizado

**Fuente:** 3 archivos CSV (datos creados)

**Descripción:** 1. Ventas.csv
**Filas:** 100 transacciones
**Columnas clave:** ID_Venta, ID_Producto, ID_Cliente, Cantidad, Fecha

**Descripción:** 2. Productos.csv
**Filas:** 50 productos (con 10 duplicados para probar limpieza)
**Columnas clave:** ID_Producto, Nombre, Categoría, Precio

**Descripción:** 3. Clientes.csv
**Filas:** 30 clientes
**Columnas clave:** ID_Cliente, Nombre_Cliente, Ciudad, Región

**Problemas simulados:**
- Espacios extra en IDs 
- Duplicados en Productos 
- Valores null en algunas filas 
- Tipos de datos incorrectos (fechas como texto)

## Proceso / Pipeline

**Nota:** Todo el proceso fue implementado en **una única query principal**, organizada por etapas lógicas dentro del código M, evitando dependencias visibles entre consultas.

### **Etapa 1: Carga desde Folder**

**Qué hace:** 
- Utiliza `Folder.Files()` para detectar archivos en directorio 
- Identifica archivos por nombre: `Ventas.csv`, `Productos.csv`, `Clientes.csv` 
- Importa cada archivo con `Csv.Document` 

**Código M (concepto):**
```m
let 
	Source = Folder.Files("C:\Data\"), 
	VentasFile = Table.SelectRows(Source, each [Name] = "Ventas.csv"), 
	VentasCSV = Csv.Document(VentasFile{0}[Content]) 
in 
	VentasCSV
```

**Resultado:** 
Datos cargados de forma automática y reutilizable ante cambios en los CSV. 
**Aprendizaje:** 
Importar desde folder (en lugar de rutas fijas) facilita automatización. 
Ver concepto: [[Power-Query-Basics#Get Data]]
### **Etapa 2: Promoción de encabezados y tipos**

**Transformaciones aplicadas:** 
```m 
// Promover primera fila a headers 
Table.PromoteHeaders(Source) 

// Cambiar tipos de datos 
Table.TransformColumnTypes(Source, { 
	{"ID_Venta", type text}, 
	{"ID_Producto", type text}, 
	{"ID_Cliente", type text}, 
	{"Cantidad", Int64.Type}, 
	{"Fecha", type date}, 
	{"Precio", type number} 
}) 
``` 

**Resultado:** 
- Headers correctos 
- Tipos de datos apropiados para cada columna 
**Aprendizaje clave:** 
Verificar tipos de datos **al inicio** evita errores posteriores en cálculos y merges. 
Ver concepto: [[Power-Query-Basics#Change Type]]

### **Etapa 3: Limpieza de datos** 

#### **Tabla Ventas:** 
**Transformaciones:** 
```m 
// Eliminar filas con Cantidad null 
Table.SelectRows(Source, each [Cantidad] <> null) 

// Trim en columnas ID 
Table.TransformColumns(Source, { 
	{"ID_Producto", Text.Trim}, 
	{"ID_Cliente", Text.Trim} 
}) 

// Eliminar duplicados por ID_Venta 
Table.Distinct(Source, {"ID_Venta"}) 
``` 

**Resultado:** 
- Filas válidas (sin nulls en Cantidad) 
- IDs limpios (sin espacios) 
- Sin transacciones duplicadas

#### **Tabla Productos:** 
**Transformaciones:** 
```m 
// Trim en IDs 
Table.TransformColumns(Source, {{"ID_Producto", Text.Trim}}) 

// Normalizar nombres con Text.Proper 
Table.TransformColumns(Source, {{"Nombre", Text.Proper}}) 

// Eliminar duplicados 
Table.Distinct(Source, {"ID_Producto"}) 
``` 

**Resultado:** 
- 50 filas → 40 filas únicas (10 duplicados eliminados) 
- Nombres estandarizados 
- IDs limpios 
Ver concepto: [[Power-Query-Text-Functions#Text.Trim()]]  
#### **Tabla Clientes:** 
**Transformaciones:** 
```m 
// Similar a Productos: 

// - Trim en IDs 
// - Text.Proper en nombres 
// - Remove Duplicates por ID_Cliente 
``` 

**Resultado:** 
Tablas maestras limpias y consistentes antes de merges.

### **Etapa 4: Combinación de tablas (Merge)** 
#### **Merge 1: Ventas + Productos** 
**Configuración:** 
- **Tipo:** Inner Join 
- **Clave:** ID_Producto 
- **Columnas expandidas:** Nombre, Categoría, Precio 

**Código M:** 
```m 
// Merge 
Merged_Productos = Table.NestedJoin( 
	Ventas_Clean, 
	{"ID_Producto"}, 
	Productos_Clean, 
	{"ID_Producto"}, 
	"Productos", 
	JoinKind.Inner 
), 

// Expand 
Expanded_Productos = Table.ExpandTableColumn( 
	Merged_Productos, 
	"Productos", 
	{"Nombre", "Categoría", "Precio"}, 
	{"Nombre_Producto", "Categoría", "Precio_Unitario"} ) 
``` 

**Resultado:** 
Ventas enriquecidas con información del producto. 
Ver concepto: [[Merge-vs-Append#Inner Join]]  
#### **Merge 2: Resultado + Clientes** 
**Configuración:** 
- **Tipo:** Inner Join 
- **Clave:** ID_Cliente 
- **Columnas expandidas:** Nombre_Cliente, Ciudad, Región 

**Código M:** 
```m 
// Merge 
Merged_Clientes = Table.NestedJoin( 
	Ventas_Con_Productos, 
	{"ID_Cliente"}, 
	Clientes_Clean, 
	{"ID_Cliente"}, 
	"Clientes", 
	JoinKind.Inner 
), 

// Expand Expanded_Clientes = Table.ExpandTableColumn( 
	Merged_Clientes, 
	"Clientes", 
	{"Nombre_Cliente", "Ciudad", "Región"}, 
	{"Nombre_Cliente", "Ciudad", "Región"} 
) 
``` 

**Resultado:** 
Tabla completa con información de producto y cliente. 
**Aprendizaje clave:** 
La limpieza previa (Trim + Remove Duplicates) es **CRÍTICA** para evitar: 
- Explosión de filas en merge 
- Pérdida de matches por espacios 
- Valores null post-merge

Ver concepto: [[Merge-vs-Append#Left Outer Join]]
### **Etapa 5: Columnas calculadas** 
**Transformaciones:** 

**1. ID de venta formateado:** 
```m 
Table.AddColumn( 
	Source, 
	"ID_Venta_Formateado", 
	each "V" & Text.From([ID_Venta]), 
	type text 
) 
``` 

**2. Valor Total (métrica clave):** 
```m 
Table.AddColumn( 
	Source, 
	"Valor_Total", 
	each [Cantidad] * [Precio_Unitario], 
	type number 
) 
``` 

**Resultado:** 
- ID legible para reportes 
- Métrica calculada lista para análisis 
Ver concepto: [[Power-Query-Transformations#Add Column]]

### **Etapa 6: Tabla final** 
**Transformaciones:** 
```m 
// Seleccionar solo columnas relevantes 
Table.SelectColumns( 
	Source, 
	{ 
		"ID_Venta_Formateado", 
		"Fecha", "Nombre_Cliente", 
		"Ciudad", 
		"Región", 
		"Nombre_Producto", 
		"Categoría", 
		"Cantidad", 
		"Precio_Unitario", 
		"Valor_Total" 
	} 
) 

// Ordenar por fecha descendente 
Table.Sort(Final, {{"Fecha", Order.Descending}}) 
``` 

**Resultado final:** 
Una tabla consolidada de ventas con: 
- Información completa de producto 
- Información completa de cliente 
- Métricas calculadas 
- Estructura lista para dashboard

## Problemas encontrados y soluciones 
### **Problema 1: Duplicación de filas tras el merge** 
**Descripción:** 
- Tabla Ventas: 100 filas 
- Después de merge con Productos: 300+ filas 
- Triplicó las filas 
 
**Causa:** 
Productos tenía IDs duplicados → cada venta con ID duplicado generaba múltiples filas. 

**Ejemplo:** 
``` 
Ventas:                             Productos: 
ID_Prod        | Cant          ID_Prod      | Nombre 
1                    | 10              1                  | Laptop 
                    1                 | Laptop ← Duplicado 
Merge genera: 
ID_Prod    | Cant     | Nombre 
1                | 10         | Laptop 
1                | 10         | Laptop ← Duplicado 
``` 
**Solución:** 
```m 
// ANTES de merge: 
Productos_Clean = Table.Distinct(Productos_Raw, {"ID_Producto"}) 

// LUEGO merge 
Resultado: 100 filas → 100 filas  
``` 

**Aprendizaje:** 
Las tablas secundarias **SIEMPRE** deben limpiarse de duplicados antes de un merge. 

Ver documentado: [[Merge-vs-Append#Error 2]]  
### **Problema 2: IDs que no coincidían en merge** 
**Descripción:** 
- Merge esperaba 100 matches 
- Solo obtuve 60 matches 
- 40 ventas sin información de producto 
 
**Causa:** 
Espacios invisibles en columnas ID: 
``` 
Ventas: ID_Producto = "PROD123" 
Productos: ID_Producto = "PROD123 " (espacio al final) 

"PROD123" ≠ "PROD123 " → No match 
```

**Solución:** 
```m 
// En AMBAS tablas, ANTES de merge: 
Ventas_Trimmed = Table.TransformColumns( 
	Source, 
	{{"ID_Producto", Text.Trim}, {"ID_Cliente", Text.Trim}} 
) 

Productos_Trimmed = Table.TransformColumns( 
	Source, 
	{{"ID_Producto", Text.Trim}} 
) 

Resultado: 60 matches → 98 matches 
``` 

**Aprendizaje:** 
**SIEMPRE** aplicar `Text.Trim()` a columnas de keys antes de merge. 

Ver documentado: [[Merge-vs-Append#Error 1]]  
### **Problema 3: Tipos incorrectos (Fecha como texto)** 
**Descripción:** 
- Columna Fecha venía como Text 
- No podía extraer Año/Mes con funciones de fecha 
- `Date.Year([Fecha])` generaba error 
 
**Causa:** 
CSV importa fechas como string: `"2024-01-15"` (texto) 
**Solución:** 
```m 
// Change Type en importación de Ventas 
Ventas_Types = Table.TransformColumnTypes( 
	Source, 
	{{"Fecha", type date}} 
) 

// Ahora funciona: 
Added_Year = Table.AddColumn(Source, "Año", each Date.Year([Fecha])) 
``` 

**Aprendizaje:** 
Verificar y ajustar tipos de datos **SIEMPRE** al importar. 

Ver concepto: [[Power-Query-Basics#Change Type]]

## Conceptos de Power Query aplicados 
- [[Power-Query-Basics#Applied Steps]] - Todo organizado en steps editables 
- [[Power-Query-Basics#Get Data]] - Import desde Folder 
- [[Power-Query-Basics#Change Type]] - Tipos correctos desde inicio 
- [[Power-Query-Text-Functions#Text.Trim()]] - Limpieza de IDs pre-merge 
- [[Power-Query-Text-Functions#Text.Proper()]] - Estandarizar nombres 
- [[Power-Query-Transformations#Remove Duplicates]] - Limpieza crítica 
- [[Merge-vs-Append#Inner Join]] - Combinar tablas relacionadas 
- [[Power-Query-Transformations#Add Column]] - Calcular Valor_Total 
- [[M-Language-Basics]] - Estructura let...in, funciones
## Resultados / Entregable

## 📈 Resultados / Entregable 
**Reporte automatizado funcional:** 
- Importación automática desde folder 
- Pipeline ETL completo 
- Refresh exitoso al modificar CSV 
- Código M claro y organizado 
- Tabla final lista para análisis 
 
**Test de automatización:** 
1. Modifiqué CSVs (cambié cantidades, agregué filas) 
2. Excel → Data → Refresh All 


### Insights clave:
1. **El orden importa:** Limpiar **antes** de merge evita errores graves 
2. **Duplicados en tablas maestras son peligrosos:** Causan explosión de filas 
3. **Power Query es ETL completo:** No solo "limpieza de Excel" 
4. **Applied Steps = debugging visual:** Puedo ver cada transformación 
5. **Documentar el proceso acelera aprendizaje:** Escribir problemas/soluciones ayuda 
6. **Inner Join fue el correcto:** Mantiene solo datos válidos completos 
7. **Una única query bien organizada > múltiples queries dispersas**






## Links

**GitHub:** 
**Archivos locales:** "G:\Mi unidad\DataAnalyst-Vault\03-Proyectos\Proyecto 01 - Reporte Automatizado con Power Query.xlsx"


---
**Relacionado:** [[Semana-01]] - [[Power-Query-Basics]] - [[Merge-vs-Append]] - [[2025-01-26]] 
**Tags:** #proyecto #excel #powerquery #automation #merge #etl #semana01 #completado #portfolio