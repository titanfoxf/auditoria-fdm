import http.server
import socketserver
import json
import urllib.parse
import os
import openpyxl

PORT = 8050
EXCEL_PATH = r"D:\ANTIGRAVITACION\ESP_PROYECTOS_ELABORACION\PUNTO_61_MATRIZ_OBRAS_INICIADAS_Y_TERMINADAS_2022_2025.xlsx"
STATIC_DIR = os.path.dirname(os.path.abspath(__file__))

def get_projects_from_excel():
    # Load enriched data from projects_data.json
    extra_map = {}
    cached_kpis = {}
    json_path = os.path.join(STATIC_DIR, 'projects_data.json')
    if os.path.exists(json_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                d_static = json.load(f)
                cached_kpis = d_static.get('kpis', {})
                for pr in d_static.get('projects', []):
                    extra_map[pr['id']] = pr
        except Exception as e:
            print("Error loading extra map:", e)

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
    projects = []
    
    # 1. Read Obras Civiles (with checks)
    ws_oc = wb['01_OBRAS_CIVILES']
    for r in range(6, ws_oc.max_row + 1):
        c_id = ws_oc.cell(r, 1).value
        if c_id is None or not str(c_id).isdigit():
            continue
        
        gid = int(c_id)
        conv = ws_oc.cell(r, 2).value or ''
        prov = ws_oc.cell(r, 3).value or ''
        desc = ws_oc.cell(r, 4).value or ''
        cui = str(ws_oc.cell(r, 5).value or 'N/D').strip()
        est = ws_oc.cell(r, 6).value or ''
        prio = ws_oc.cell(r, 7).value or 0.0
        ejec = ws_oc.cell(r, 8).value or 0.0
        tiene_caja = str(ws_oc.cell(r, 9).value or '')
        n_cajas = str(ws_oc.cell(r, 10).value or '')

        checks = {
            'exp': '[ ✓ ]' in str(ws_oc.cell(r, 11).value),
            'cont': '[ ✓ ]' in str(ws_oc.cell(r, 12).value),
            'acta_ent': '[ ✓ ]' in str(ws_oc.cell(r, 13).value),
            'cuad': '[ ✓ ]' in str(ws_oc.cell(r, 14).value),
            'val': '[ ✓ ]' in str(ws_oc.cell(r, 15).value),
            'adic': '[ ✓ ]' in str(ws_oc.cell(r, 16).value),
            'sup': '[ ✓ ]' in str(ws_oc.cell(r, 17).value),
            'rec': '[ ✓ ]' in str(ws_oc.cell(r, 18).value),
            'liq': '[ ✓ ]' in str(ws_oc.cell(r, 19).value),
            'res': '[ ✓ ]' in str(ws_oc.cell(r, 20).value),
            'transf': '[ ✓ ]' in str(ws_oc.cell(r, 21).value),
        }
        alerta = ws_oc.cell(r, 22).value or ''

        ext = extra_map.get(gid, {})

        projects.append({
            'id': gid,
            'convenio': conv,
            'provincia': prov,
            'desc': desc,
            'cui': cui,
            'tipologia': 'OBRAS CIVILES',
            'estatus': est,
            'prio': prio,
            'costo_ejec': ejec,
            'tiene_caja': tiene_caja,
            'n_cajas': n_cajas,
            'checks': checks,
            'alerta': alerta,
            'excel_sheet': '01_OBRAS_CIVILES',
            'excel_row': r,
            'historial': ext.get('historial', []),
            'fecha_inicio': ext.get('fecha_inicio', ''),
            'fecha_inicio_sort': ext.get('fecha_inicio_sort', '9999-99-99'),
            'tiene_fecha_inicio': ext.get('tiene_fecha_inicio', False),
            'fecha_inicio_doc': ext.get('fecha_inicio_doc', ''),
            'fecha_termino_prog': ext.get('fecha_termino_prog', ''),
            'fecha_culminacion_real': ext.get('fecha_culminacion_real', ''),
            'acta_recepcion': ext.get('acta_recepcion', ''),
            'resolucion_liquidacion': ext.get('resolucion_liquidacion', ''),
            'estado_info': ext.get('estado_info', 'PARCIAL' if ext.get('historial') else 'INCOMPLETA'),
            'estado_info_label': ext.get('estado_info_label', ''),
            'detalles_faltantes': ext.get('detalles_faltantes', []),
            'confiabilidad': ext.get('confiabilidad', 'MEDIA')
        })

    # Solo Obras Civiles (69 proyectos estrictos del Punto 61)
    kpis = cached_kpis if cached_kpis else {
        'total': len(projects),
        'obras': len([p for p in projects if p['tipologia'] == 'OBRAS CIVILES']),
        'cajas': len([p for p in projects if 'CAJA' in p['tiene_caja'].upper()]),
        'pendientes_archivador': len([p for p in projects if p['tipologia'] == 'OBRAS CIVILES' and 'CAJA' not in p['tiene_caja'].upper()]),
        'con_historial': len([p for p in projects if len(p.get('historial', [])) > 0]),
        'con_fecha_inicio': len([p for p in projects if p.get('tiene_fecha_inicio')]),
        'sin_fecha_inicio': len([p for p in projects if not p.get('tiene_fecha_inicio')]),
        'info_completa': len([p for p in projects if p.get('estado_info') == 'COMPLETA']),
        'info_parcial': len([p for p in projects if p.get('estado_info') == 'PARCIAL']),
        'info_incompleta': len([p for p in projects if p.get('estado_info') == 'INCOMPLETA']),
        'tomos_procesados': 5,
        'folios_procesados': 399
    }
    return projects, kpis

def update_project_in_excel(payload):
    p_id = payload.get('id')
    checks = payload.get('checks', {})
    ubicacion = payload.get('ubicacion', '')
    observaciones = payload.get('observaciones', '')

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_oc = wb['01_OBRAS_CIVILES']
    
    target_row = None
    for r in range(6, ws_oc.max_row + 1):
        if str(ws_oc.cell(r, 1).value) == str(p_id):
            target_row = r
            break
            
    if not target_row:
        return False
        
    def val(b): return "[ ✓ ]" if b else "[   ]"
    
    ws_oc.cell(target_row, 10, ubicacion)
    ws_oc.cell(target_row, 11, val(checks.get('exp', False)))
    ws_oc.cell(target_row, 12, val(checks.get('cont', False)))
    ws_oc.cell(target_row, 13, val(checks.get('acta_ent', False)))
    ws_oc.cell(target_row, 14, val(checks.get('cuad', False)))
    ws_oc.cell(target_row, 15, val(checks.get('val', False)))
    ws_oc.cell(target_row, 16, val(checks.get('adic', False)))
    ws_oc.cell(target_row, 17, val(checks.get('sup', False)))
    ws_oc.cell(target_row, 18, val(checks.get('rec', False)))
    ws_oc.cell(target_row, 19, val(checks.get('liq', False)))
    ws_oc.cell(target_row, 20, val(checks.get('res', False)))
    ws_oc.cell(target_row, 21, val(checks.get('transf', False)))
    ws_oc.cell(target_row, 22, observaciones)

    wb.save(EXCEL_PATH)
    return True

class AuditRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == '/api/projects':
            projects, kpis = get_projects_from_excel()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'projects': projects, 'kpis': kpis}, ensure_ascii=False).encode('utf-8'))
        elif parsed.path == '/' or parsed.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            with open(os.path.join(STATIC_DIR, 'index.html'), 'rb') as f:
                self.wfile.write(f.read())
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == '/api/save':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode('utf-8'))
            success = update_project_in_excel(payload)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': success}).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

if __name__ == '__main__':
    os.chdir(STATIC_DIR)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), AuditRequestHandler) as httpd:
        print(f"=====================================================")
        print(f"APLICATIVO FDM AUDITORÍA INICIADO EXITOSAMENTE")
        print(f"Accede desde tu navegador en: http://localhost:{PORT}")
        print(f"Sincronizando con: {EXCEL_PATH}")
        print(f"=====================================================")
        httpd.serve_forever()
