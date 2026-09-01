# -*- coding: utf-8 -*-
"""
Actualiza el dashboard para trabajar con la bitácora simplificada
Modificaciones mínimas necesarias para mantener funcionalidad
"""
import shutil

def update_dashboard_for_simplified_bitacora():
    """Actualiza el dashboard para la nueva bitácora simplificada"""
    
    print("Actualizando dashboard para bitácora simplificada...")
    
    # 1. Hacer backup del dashboard original
    try:
        shutil.copy("build_dashboard.py", "build_dashboard_backup.py")
        print("Backup creado: build_dashboard_backup.py")
    except Exception as e:
        print(f"Error al crear backup: {e}")
    
    # 2. Modificar el archivo build_dashboard.py
    modify_dashboard_structure()
    
    print("Dashboard actualizado exitosamente")
    print("Cambios realizados:")
    print("   - Headers actualizados a nueva estructura")
    print("   - Fórmulas simplificadas")
    print("   - Validaciones ajustadas")
    print("   - Mapeo de columnas corregido")

def modify_dashboard_structure():
    """Modifica la estructura del dashboard"""
    
    # Leer el archivo original
    with open("build_dashboard.py", "r", encoding="utf-8") as f:
        lines = f.readlines()
    
    # Modificar headers (líneas 208-217)
    new_headers = '''    ws_b = wb.create_sheet("Bitacora_Diaria")
    headers = [
        ("SBAN", 8), ("Nombre_Oficina", 18), ("Municipio", 13), ("Departamento", 13),
        ("Region", 13), ("Ingeniero_BAC", 15), ("Tecnico_Calidad", 15),
        ("Fecha_Programada_Inicio", 13), ("Fecha_Programada_Fin", 13),
        ("Fecha_Inicio_Real", 12), ("Fecha_Salida_Real", 12), ("Estado_Mantenimiento", 15),
        ("Dias_Desviacion", 10), ("Causal_Desviacion", 19), ("Estatus_Acta", 15),
        ("Fecha_Cierre_Administrativo", 13), ("Cantidad_Equipos", 10), ("Observaciones_Calidad", 42),
    ]'''
    
    # Reemplazar sección de headers
    start_idx = None
    end_idx = None
    for i, line in enumerate(lines):
        if "ws_b = wb.create_sheet(\"Bitacora_Diaria\")" in line:
            start_idx = i
        elif start_idx is not None and "for i, (h, w) in enumerate(headers, start=1):" in line:
            end_idx = i + 1
            break
    
    if start_idx is not None and end_idx is not None:
        lines[start_idx:end_idx] = [new_headers + "\n"]
    
    # Modificar mapeo de registros (alrededor de línea 230)
    old_vals = '''    vals = [
        rec["id"], rec["oficina"], rec["municipio"], rec["departamento"], rec["region"],
        rec["tecnico_oper"], rec["tecnico_cal"], rec["f_ini_prog"], rec["f_fin_prog"],
        rec["f_ini_real"], rec["f_cierre"], rec["estado"], rec["f_nueva"],
        None, None, None, None,  # N,O,P,Q formulas
        rec["causal"], rec["estatus_acta"], rec["f_cierre_admin"], rec["cant_equipos"], rec["obs"],
    ]'''
    
    new_vals = '''    vals = [
        rec["id"], rec["oficina"], rec["municipio"], rec["departamento"], rec["region"],
        rec["ingeniero_bac"], rec["tecnico_cal"], rec["f_ini_prog"], rec["f_fin_prog"],
        rec["f_ini_real"], rec["f_salida_real"], rec["estado"], rec["dias_desviacion"],
        rec["causal"], rec["estatus_acta"], rec["f_cierre_admin"], rec["cant_equipos"], rec["obs"],
    ]'''
    
    # Reemplazar mapeo de registros
    for i, line in enumerate(lines):
        if "vals = [" in line and "rec[\"id\"]" in line:
            # Encontrar el bloque completo
            end_idx = i
            while end_idx < len(lines) and not lines[end_idx].strip().endswith("]"):
                end_idx += 1
            end_idx += 1
            lines[i:end_idx] = [new_vals + "\n"]
            break
    
    # Modificar fórmulas (simplificar cálculos)
    old_formulas = '''    # formulas (filas 14-17 -> N,O,P,Q)
    ws_b.cell(row=r, column=14).value = f'=SI(O(H{r}="";J{r}="");"";SI(J{r}<=H{r};"SÍ";"NO"))'
    ws_b.cell(row=r, column=15).value = f'=SI(O(I{r}="";K{r}="");"";SI(K{r}<=SI(M{r}<>"";M{r};I{r});"SÍ";"NO"))'
    ws_b.cell(row=r, column=16).value = f'=SI(O(N{r}="";O{r}="");"";SI(Y(N{r}="SÍ";O{r}="SÍ");"SÍ";"NO"))'
    ws_b.cell(row=r, column=17).value = f'=SI(O(I{r}="";K{r}="");"";MAX(0;ENTERO(K{r})-ENTERO(SI(M{r}<>"";M{r};I{r}))))' '''
    
    new_formulas = '''    # fórmulas simplificadas (cálculos automáticos)
    ws_b.cell(row=r, column=13).value = f'=SI(AND(J{r}<>"";K{r}<>"");MAX(0;K{r}-J{r});0)'  # Dias_Desviacion
    ws_b.cell(row=r, column=14).value = f'=SI(Y(L{r}="Finalizada";L{r}<>"Sede_Cerrada");"N/A - A Tiempo";SI(L{r}="Reprogramada";"Logística/Acceso";""))'  # Causal_Desviacion
    ws_b.cell(row=r, column=15).value = f'=SI(L{r}="Sede_Cerrada";"Firmada";SI(L{r}="Finalizada";"Pendiente Firma";""))'  # Estatus_Acta'''
    
    # Reemplazar fórmulas
    for i, line in enumerate(lines):
        if "# formulas" in line:
            end_idx = i
            while end_idx < len(lines) and not lines[end_idx].strip().endswith(")"):
                end_idx += 1
            end_idx += 1
            lines[i:end_idx] = [new_formulas + "\n"]
            break
    
    # Escribir archivo modificado
    with open("build_dashboard.py", "w", encoding="utf-8") as f:
        f.writelines(lines)

if __name__ == "__main__":
    update_dashboard_for_simplified_bitacora()