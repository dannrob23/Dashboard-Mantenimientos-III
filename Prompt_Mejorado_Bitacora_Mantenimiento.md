# PROMPT MEJORADO — Hoja de cálculo de mantenimientos y auditoría de calidad (Slate & Navy)

## Rol y objetivo
Actúa como ingeniero de procesos, analista de datos y diseñador de dashboards. Con base en el archivo
"Campos dashborad.xlsx" (registro real de mantenimientos por **sede SBAN**: ID, tipo, Nombre Oficina,
Municipio, Departamento, Jefatura Regional, Fecha Inicio/Fin, Estado de la sede, cumplimiento de
cronograma, inventarios validados, actas, técnicos y conteo de equipos), diseña una hoja de cálculo
profesional y sobria (estilo corporativo **Slate & Navy**) para el **registro diario de mantenimientos y
auditoría de calidad**, con **cálculo automático de KPIs y de la desviación de mantenimientos en días**.

## Alcance de datos (verificado contra la fuente)
- La unidad de registro es la **SEDE (oficina)**, no el equipo individual.
- Jefaturas regionales válidas: `ANTIOQUIA, ORIENTE, COSTA, SUR, OCCIDENTE, SANTANDERES, CAFETERA, BOGOTA, DIRECCION GENERAL, COA`.
- Estados válidos: `Programada, En Proceso, Finalizada, Reprogramada, Sede_Cerrada`.
- Excluir de los KPIs de desviación las filas tipo `DG` (Dirección General) y `COA` (cronogramas de 74+
  días que distorsionan la métrica); registrarlas en hoja aparte si el negocio lo requiere.

## 1. Estructura del libro (4 pestañas)
1. **Instructivo_Tecnico**: manual del técnico/calidad con columnas:
   `Campo/Columna | Tipo de Dato | ¿Qué información poner? | Regla de negocio/Validación | ¿Por qué se registra? | ¿Para qué sirve? (KPI/Impacto)`
2. **Listas_Validacion**: listas desplegables (regiones, estados, causales, estatus acta, SÍ/NO) y tabla
   Región ↔ Departamentos para validación cruzada.
3. **Bitacora_Diaria**: tabla principal de registro diario (bordes finos, cuadrícula visible, fila de
   encabezado congelada). Implementar como **Tabla de Excel** (o rango con nombre) para rangos dinámicos.
4. **Dashboard_KPIs**: tarjetas de KPIs arriba + tablas de desglose (por región, por causal, por técnico,
   evolución semanal) + bloque de parámetros/metas para semáforos.

## 2. Reglas de negocio
- **Cierre Operativo** (Estado = `Finalizada`): se diligencia `Fecha_Cierre_Operativo` cuando el técnico
  concluye la intervención física en campo.
- **Cierre Administrativo** (Estado = `Sede_Cerrada`): calidad valida evidencias (inventarios, checklist,
  actas) y firma; se diligencia `Fecha_Cierre_Administrativo` y `Estatus_Acta = Firmada Conforme`.
  No se puede cerrar administrativamente sin esa validación (regla condicional).
- **Reprogramada**: obligatorio `Fecha_Nueva_Programada` y `Causal_Desviacion ≠ "N/A - A Tiempo"`.
- **Ejecutadas** = `Finalizada + Sede_Cerrada`. **Pendientes** = `Programada + En Proceso + Reprogramada`.
- **Desviación**: fecha comprometida vigente = `Fecha_Nueva_Programada` si existe, si no `Fecha_Programada_Fin`.
  `Dias_Desviacion = MAX(0; Fecha_Cierre_Operativo − fecha comprometida vigente)`.
- **Validaciones cruzadas**: `Fecha_Programada_Fin ≥ Fecha_Programada_Inicio`;
  `Fecha_Cierre_Operativo ≥ Fecha_Inicio_Real`; `Estado = Sede_Cerrada` ⇒ `Estatus_Acta = "Firmada
  Conforme"` y `Fecha_Cierre_Administrativo` no vacía.

## 3. Especificación de campos — Bitacora_Diaria (columnas A–V)

| Col | Campo | Tipo / Validación | Notas |
|---|---|---|---|
| A | ID_Orden | Texto, único | Código SBAN, ej. 1321 |
| B | Nombre_Oficina | Texto obligatorio | Ej. CAUCASIA |
| C | Municipio | Texto opcional | Contexto geográfico |
| D | Departamento | Texto opcional | Cruza con Region (Listas_Validacion) |
| E | Region | Desplegable desde Listas_Validacion | 10 jefaturas reales |
| F | Tecnico_Operativo | Texto obligatorio | Quién interviene en campo |
| G | Tecnico_Calidad | Texto obligatorio | Quién audita y firma |
| H | Fecha_Programada_Inicio | Fecha AAAA-MM-DD, obligatoria | |
| I | Fecha_Programada_Fin | Fecha AAAA-MM-DD, ≥ H | Fin del cronograma comprometido |
| J | Fecha_Inicio_Real | Fecha | Se diligencia al llegar a la sede |
| K | Fecha_Cierre_Operativo | Fecha | Obligatoria si Estado = Finalizada/Sede_Cerrada |
| L | Estado_Mantenimiento | Desplegable | Programada, En Proceso, Finalizada, Reprogramada, Sede_Cerrada |
| M | Fecha_Nueva_Programada | Fecha | Obligatoria si L = Reprogramada (validación personalizada) |
| N | Cumpli_Ingreso | Fórmula (SÍ/NO) | Inicio real ≤ inicio programado |
| O | Cumpli_Salida | Fórmula (SÍ/NO) | Cierre operativo ≤ fecha comprometida vigente |
| P | Cumpli_Cronograma | Fórmula (SÍ/NO) | SÍ solo si N y O = SÍ |
| Q | Dias_Desviacion | Fórmula (días ≥ 0) | Retraso real en días |
| R | Causal_Desviacion | Desplegable | N/A - A Tiempo, Repuesto Faltante, Cliente No Disponible, Clima, Falla Técnica, Reprogramación Cliente, Logística/Acceso |
| S | Estatus_Acta | Desplegable | Pendiente Firma, Firmada Conforme, Rechazada por Calidad |
| T | Fecha_Cierre_Administrativo | Fecha | Vacía hasta la firma de calidad |
| U | Cantidad_Equipos | Entero ≥ 0 | Total de equipos impactados de la sede |
| V | Observaciones_Calidad | Texto | Libre |

### Fórmulas de la fila 5 (arrastrar hacia abajo; si se usa Tabla de Excel, reemplazar por referencias estructuradas `Tabla1[@Campo]`)
```
N5 =SI(O(H5="";J5="");"";SI(J5<=H5;"SÍ";"NO"))
O5 =SI(O(I5="";K5="");"";SI(K5<=SI(M5<>"";M5;I5);"SÍ";"NO"))
P5 =SI(O(N5="";O5="");"";SI(Y(N5="SÍ";O5="SÍ");"SÍ";"NO"))
Q5 =SI(O(I5="";K5="");"";MAX(0;ENTERO(K5)-ENTERO(SI(M5<>"";M5;I5))))
```

## 4. Dashboard_KPIs — fórmulas

### Tarjetas
```
Ejecutadas        =CONTAR.SI(Bitacora_Diaria!L:L;"Finalizada")+CONTAR.SI(Bitacora_Diaria!L:L;"Sede_Cerrada")
Pendientes        =CONTAR.SI(Bitacora_Diaria!L:L;"Programada")+CONTAR.SI(Bitacora_Diaria!L:L;"En Proceso")+CONTAR.SI(Bitacora_Diaria!L:L;"Reprogramada")
%Cumplimiento     =SI(Ejecutadas=0;0;CONTAR.SI(Bitacora_Diaria!P:P;"SÍ")/Ejecutadas)
%Desviación       =SI(Ejecutadas=0;0;CONTAR.SI(Bitacora_Diaria!P:P;"NO")/Ejecutadas)
Desv. promedio    =SI.ERROR(PROMEDIO.SI.CONJUNTO(Bitacora_Diaria!Q:Q;Bitacora_Diaria!P:P;"NO";Bitacora_Diaria!Q:Q;">0");0)
%Cierre admin.    =SI(Ejecutadas=0;0;CONTAR.SI(Bitacora_Diaria!L:L;"Sede_Cerrada")/Ejecutadas)
Equipos impact.   =SUMAR.SI(Bitacora_Diaria!L:L;"Finalizada";Bitacora_Diaria!U:U)+SUMAR.SI(Bitacora_Diaria!L:L;"Sede_Cerrada";Bitacora_Diaria!U:U)
```

### Tablas de desglose
- **Por región** (una fila por jefatura; ejemplo ANTIOQUIA):
  ```
  % Cumplimiento =SI.ERROR(CONTAR.SI.CONJUNTO(Bitacora_Diaria!E:E;"ANTIOQUIA";Bitacora_Diaria!P:P;"SÍ")/
                  (CONTAR.SI.CONJUNTO(Bitacora_Diaria!E:E;"ANTIOQUIA";Bitacora_Diaria!L:L;"Finalizada")+
                   CONTAR.SI.CONJUNTO(Bitacora_Diaria!E:E;"ANTIOQUIA";Bitacora_Diaria!L:L;"Sede_Cerrada"));0)
  % Desviación   =SI.ERROR(CONTAR.SI.CONJUNTO(Bitacora_Diaria!E:E;"ANTIOQUIA";Bitacora_Diaria!P:P;"NO")/
                  (CONTAR.SI.CONJUNTO(Bitacora_Diaria!E:E;"ANTIOQUIA";Bitacora_Diaria!L:L;"Finalizada")+
                   CONTAR.SI.CONJUNTO(Bitacora_Diaria!E:E;"ANTIOQUIA";Bitacora_Diaria!L:L;"Sede_Cerrada"));0)
  Desv. días     =SI.ERROR(PROMEDIO.SI.CONJUNTO(Bitacora_Diaria!Q:Q;Bitacora_Diaria!E:E;"ANTIOQUIA";Bitacora_Diaria!P:P;"NO");0)
  ```
- **Por causal (Pareto)**: `=CONTAR.SI.CONJUNTO(Bitacora_Diaria!P:P;"NO";Bitacora_Diaria!R:R;"<causal>")`
  y su % sobre el total de registros con `NO`.
- **Por técnico**: cumplimiento y desviación promedio agrupados por `Tecnico_Operativo`
  (CONTAR.SI.CONJUNTO / PROMEDIO.SI.CONJUNTO con criterio sobre F).
- **Evolución semanal**: fechas de cierre operativo (K) vs. % de cumplimiento acumulado.

## 5. Estilo visual (Slate & Navy)
- Encabezados: fondo `#1E293B`, texto blanco bold, altura de fila 60 px.
- Filas alternadas (zebra): `#F8FAFC`. Bordes sutiles: `#CBD5E1`. Texto general: `#0F172A`.
- Congelar fila de encabezados y primera columna de identificación.
- Tarjetas KPI: título 11 pt gris `#64748B`, valor 28 pt navy `#0F172A`, fondo blanco con borde slate.
- Formato condicional:
  - `Cumpli_Cronograma`: SÍ → verde `#16A34A`, NO → rojo `#DC2626`.
  - `Estado_Mantenimiento`: Programada gris, En Proceso azul `#3B82F6`, Finalizada verde,
    Reprogramada ámbar `#F59E0B`, Sede_Cerrada navy oscuro.
- Formatos de número: tasas en `0.0%`; días con 1 decimal; enteros sin decimales.
- Semáforo en tarjetas según metas parametrizables (bloque Parámetros): % Cumplimiento ≥ 95% verde,
  90–95% ámbar, < 90% rojo.

## Entregables
4 pestañas funcionales, fórmulas automáticas con rangos dinámicos (sin valores quemados "100/0" como en
la columna M del archivo actual), manual del instructivo completo y semáforos según metas.
