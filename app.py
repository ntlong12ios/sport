import os
import sys
import math
import sqlite3
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
from typing import List, Optional
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse
from pydantic import BaseModel

# ==============================================================================
# 1. KHỞI TẠO CƠ SỞ DỮ LIỆU & IMPORT DỮ LIỆU TỪ EXCEL
# ==============================================================================
import json
DB_FILE = "muong_thanh_sports_v2.db"
EXCEL_FILE = "Data.xlsx"
MINDMAP_JSON = "saved_mindmaps.json"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cluster_configs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sport_name TEXT,
            clusters_involved TEXT,
            format_type TEXT,
            target_slots INTEGER,
            status TEXT DEFAULT 'created'
        )
    """)
    cursor.execute("CREATE TABLE IF NOT EXISTS participants (id INTEGER PRIMARY KEY AUTOINCREMENT, config_id INTEGER, name TEXT, unit_name TEXT, FOREIGN KEY (config_id) REFERENCES cluster_configs(id))")
    cursor.execute("CREATE TABLE IF NOT EXISTS matches (id INTEGER PRIMARY KEY AUTOINCREMENT, config_id INTEGER, round_number INTEGER, match_index INTEGER, participant_a_id INTEGER, participant_b_id INTEGER, score_a INTEGER, score_b INTEGER, winner_id INTEGER, next_match_id INTEGER, next_match_slot TEXT, FOREIGN KEY (config_id) REFERENCES cluster_configs(id))")
    
    cursor.execute("CREATE TABLE IF NOT EXISTS master_sports (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS master_registrations (id INTEGER PRIMARY KEY AUTOINCREMENT, cluster TEXT, unit_name TEXT, sport_name TEXT, qty INTEGER)")
    
    conn.commit()
    conn.close()

def load_data_from_excel():
    if not os.path.exists(EXCEL_FILE):
        print(f"Không tìm thấy file {EXCEL_FILE}. Bỏ qua bước nạp dữ liệu gốc.")
        return

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    cursor.execute("DELETE FROM master_registrations")
    cursor.execute("DELETE FROM master_sports")
    conn.commit()

    print("Đang nạp dữ liệu từ file Excel (Sheet: Tong hop)...")
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name='Tong hop', header=None)
        
        sports_map = {}
        for col_idx in range(4, len(df.columns)):
            val = str(df.iloc[4, col_idx]).strip()
            if val and val != 'nan':
                sports_map[col_idx] = val
                cursor.execute("INSERT INTO master_sports (name) VALUES (?)", (val,))

        current_cluster = "CỤM 1"
        for index, row in df.iloc[7:].iterrows():
            col1 = str(row[1]).strip()
            if col1.startswith('CỤM'):
                current_cluster = col1
            
            unit_name = str(row[2]).strip()
            if unit_name and unit_name not in ['TỔNG CỘNG', 'Nhân sự Ban thể thao', 'Tên đơn vị tham gia', 'nan'] and not unit_name.startswith('CỤM'):
                for col_idx, sport_name in sports_map.items():
                    qty = row[col_idx]
                    if pd.notna(qty):
                        try:
                            qty = int(float(qty))
                            if qty > 0:
                                cursor.execute("""
                                    INSERT INTO master_registrations (cluster, unit_name, sport_name, qty)
                                    VALUES (?, ?, ?, ?)
                                """, (current_cluster, unit_name, sport_name, qty))
                        except ValueError:
                            pass
        conn.commit()
        print("Nạp dữ liệu từ Excel thành công!")
    except Exception as e:
        print(f"Lỗi khi đọc file Excel: {e}")
    finally:
        conn.close()

init_db()
load_data_from_excel()

# ==============================================================================
# CÁC PHẦN CÒN LẠI GIỮ NGUYÊN HOÀN TOÀN NHƯ BẢN V2 TRƯỚC ĐÓ
# ==============================================================================
app = FastAPI(title="Mường Thanh Tournament V2 (Excel Support)")

class MindmapSaveRequest(BaseModel):
    sport: str
    state: dict

class ConfigRequest(BaseModel):
    sport_name: str
    clusters_involved: List[str]
    format_type: str
    target_slots: int
    participants: List[str]

class ScoreUpdateRequest(BaseModel):
    match_id: int
    score_a: int
    score_b: int

def generate_knockout_bracket(config_id: int, participant_ids: List[int], target_slots: int):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    N = len(participant_ids)
    if N <= target_slots:
        conn.close()
        return
        
    k = math.ceil(math.log2(N)) if N > 1 else 1
    target_k = math.ceil(math.log2(target_slots)) if target_slots > 0 else 0
    round_count = max(1, k - target_k)
    
    match_map = {}
    
    for r in range(1, round_count + 1):
        matches_in_round = 2 ** (k - r)
        for idx in range(matches_in_round):
            cursor.execute("INSERT INTO matches (config_id, round_number, match_index) VALUES (?, ?, ?)", (config_id, r, idx))
            match_map[(r, idx)] = cursor.lastrowid

    for (r, idx), m_id in match_map.items():
        if r < round_count:
            next_r = r + 1
            next_idx = idx // 2
            next_m_id = match_map.get((next_r, next_idx))
            slot = 'a' if idx % 2 == 0 else 'b'
            if next_m_id:
                cursor.execute("UPDATE matches SET next_match_id = ?, next_match_slot = ? WHERE id = ?", (next_m_id, slot, m_id))

    p_idx = 0
    for idx in range(2 ** (k - 1)):
        m_id = match_map.get((1, idx))
        if not m_id: continue
        
        if p_idx < N:
            cursor.execute("UPDATE matches SET participant_a_id = ? WHERE id = ?", (participant_ids[p_idx], m_id))
            p_idx += 1
        if p_idx < N:
            cursor.execute("UPDATE matches SET participant_b_id = ? WHERE id = ?", (participant_ids[p_idx], m_id))
            p_idx += 1
            
        cursor.execute("SELECT participant_a_id, participant_b_id, next_match_id, next_match_slot FROM matches WHERE id = ?", (m_id,))
        pa, pb, nm_id, n_slot = cursor.fetchone()
        if pa and not pb:
            cursor.execute("UPDATE matches SET winner_id = ?, score_a = 1, score_b = 0 WHERE id = ?", (pa, m_id))
            if nm_id:
                col = "participant_a_id" if n_slot == 'a' else "participant_b_id"
                cursor.execute(f"UPDATE matches SET {col} = ? WHERE id = ?", (pa, nm_id))

    conn.commit()
    conn.close()

def generate_round_robin(config_id: int, participant_ids: List[int]):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    teams = list(participant_ids)
    if len(teams) % 2 != 0:
        teams.append(None) # BYE
        
    n = len(teams)
    rounds = n - 1
    
    for r in range(1, rounds + 1):
        for i in range(n // 2):
            p1 = teams[i]
            p2 = teams[n - 1 - i]
            
            if p1 is not None and p2 is not None:
                cursor.execute("INSERT INTO matches (config_id, round_number, match_index, participant_a_id, participant_b_id) VALUES (?, ?, ?, ?, ?)", 
                               (config_id, r, i, p1, p2))
                
        teams.insert(1, teams.pop())
        
    conn.commit()
    conn.close()

@app.get("/api/master-data")
def get_master_data():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT name FROM master_sports ORDER BY id")
    sports = [r[0] for r in cursor.fetchall()]
    cursor.execute("SELECT DISTINCT cluster FROM master_registrations ORDER BY cluster")
    clusters = [r[0] for r in cursor.fetchall()]
    conn.close()
    return {"sports": sports, "clusters": clusters}

@app.get("/api/generate-participants")
def generate_participants(sport: str, clusters: str):
    cluster_list = clusters.split(",")
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    placeholders = ",".join(["?"] * len(cluster_list))
    query = f"SELECT unit_name, qty FROM master_registrations WHERE sport_name = ? AND cluster IN ({placeholders})"
    cursor.execute(query, [sport] + cluster_list)
    results = []
    for unit, qty in cursor.fetchall():
        if qty == 1:
            results.append(f"{unit}")
        else:
            for i in range(1, qty + 1):
                results.append(f"{unit}_đội {i}")
    conn.close()
    return {"participants": results}

@app.get("/api/sport-overview")
def get_sport_overview(sport: str):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT cluster, unit_name, qty FROM master_registrations WHERE sport_name = ? ORDER BY cluster, unit_name", (sport,))
    data = cursor.fetchall()
    conn.close()
    
    overview = {}
    for cluster, unit, qty in data:
        if cluster not in overview:
            overview[cluster] = []
        if qty == 1:
            overview[cluster].append(f"{unit}")
        else:
            for i in range(1, qty + 1):
                overview[cluster].append(f"{unit} (đội {i})")
                
    results = {}
    for cluster, teams in overview.items():
        cluster_num = ''.join(filter(str.isdigit, cluster))
        if not cluster_num:
            cluster_num = '0'
        
        results[cluster] = []
        for idx, team in enumerate(teams):
            results[cluster].append(f"{cluster_num}.{idx+1} {team}")
            
    return {"sport": sport, "clusters": results}

import threading
mindmap_lock = threading.Lock()

@app.post("/api/save-mindmap")
def save_mindmap(req: MindmapSaveRequest):
    with mindmap_lock:
        try:
            with open(MINDMAP_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
        except:
            data = {}
        data[req.sport] = req.state
        with open(MINDMAP_JSON, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    return {"status": "success"}


@app.get("/api/load-mindmap")
def load_mindmap(sport: str):
    try:
        with open(MINDMAP_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {"state": data.get(sport)}
    except:
        return {"state": None}

@app.get("/api/load-all-mindmaps")
def load_all_mindmaps():
    try:
        with open(MINDMAP_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    except:
        return {}

@app.post("/api/save-tournament/{config_id}")
def save_tournament(config_id: int):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("UPDATE cluster_configs SET status = 'saved' WHERE id = ?", (config_id,))
    conn.commit()
    conn.close()
    return {"message": "Saved successfully"}

@app.get("/api/qualified-teams")
def get_qualified_teams(sport: Optional[str] = None):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    query_configs = "SELECT id, sport_name, clusters_involved, format_type, target_slots FROM cluster_configs WHERE status = 'saved'"
    params = []
    if sport:
        query_configs += " AND sport_name = ?"
        params.append(sport)
    cursor.execute(query_configs, params)
    configs = cursor.fetchall()
    
    results = []
    for cid, s_name, c_inv, f_type, t_slots in configs:
        if f_type == 'round_robin':
            cursor.execute("SELECT id, name, unit_name FROM participants WHERE config_id = ?", (cid,))
            participants = {r[0]: {'id': r[0], 'name': r[1], 'unit': r[2], 'pts': 0, 'gd': 0, 'gf': 0} for r in cursor.fetchall()}
            
            cursor.execute("SELECT participant_a_id, participant_b_id, score_a, score_b FROM matches WHERE config_id = ? AND score_a IS NOT NULL AND score_b IS NOT NULL", (cid,))
            for pa, pb, sa, sb in cursor.fetchall():
                if pa not in participants or pb not in participants: continue
                participants[pa]['gf'] += sa
                participants[pb]['gf'] += sb
                participants[pa]['gd'] += (sa - sb)
                participants[pb]['gd'] += (sb - sa)
                if sa > sb:
                    participants[pa]['pts'] += 3
                elif sa < sb:
                    participants[pb]['pts'] += 3
                else:
                    participants[pa]['pts'] += 1
                    participants[pb]['pts'] += 1
            
            sorted_teams = sorted(participants.values(), key=lambda x: (x['pts'], x['gd'], x['gf']), reverse=True)
            for i in range(min(t_slots, len(sorted_teams))):
                results.append({"sport": s_name, "clusters": c_inv, "player": sorted_teams[i]['name'], "unit": sorted_teams[i]['unit']})
        else:
            cursor.execute("""
                SELECT p.name, p.unit_name 
                FROM matches m
                JOIN participants p ON m.winner_id = p.id
                WHERE m.config_id = ? AND m.next_match_id IS NULL AND m.winner_id IS NOT NULL
            """, (cid,))
            for p_name, unit_name in cursor.fetchall():
                results.append({"sport": s_name, "clusters": c_inv, "player": p_name, "unit": unit_name})

    conn.close()
    return results

@app.get("/api/search-unit")
def search_unit(unit_name: str):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT sport_name, cluster, qty FROM master_registrations WHERE unit_name LIKE ?", (f"%{unit_name}%",))
    data = [{"sport": r[0], "cluster": r[1], "qty": r[2]} for r in cursor.fetchall()]
    conn.close()
    return data

@app.get("/api/saved-configs")
def get_saved_configs():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT id, sport_name, clusters_involved FROM cluster_configs WHERE status = 'saved' ORDER BY id DESC")
    configs = [dict(zip(['id', 'sport_name', 'clusters_involved'], row)) for row in cursor.fetchall()]
    conn.close()
    return configs

@app.delete("/api/tournament/{config_id}")
def delete_tournament(config_id: int):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM matches WHERE config_id = ?", (config_id,))
    cursor.execute("DELETE FROM participants WHERE config_id = ?", (config_id,))
    cursor.execute("DELETE FROM cluster_configs WHERE id = ?", (config_id,))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.post("/api/setup-tournament")
def setup_tournament(req: ConfigRequest):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    clusters_str = ",".join(req.clusters_involved)
    cursor.execute("""
        INSERT INTO cluster_configs (sport_name, clusters_involved, format_type, target_slots, status)
        VALUES (?, ?, ?, ?, 'draft')
    """, (req.sport_name, clusters_str, req.format_type, req.target_slots))
    config_id = cursor.lastrowid
    
    participant_ids = []
    for p_name in req.participants:
        if p_name.strip():
            cursor.execute("INSERT INTO participants (config_id, name, unit_name) VALUES (?, ?, ?)", (config_id, p_name.strip(), p_name.strip()))
            participant_ids.append(cursor.lastrowid)
    conn.commit()
    conn.close()
    
    if req.format_type == 'round_robin':
        generate_round_robin(config_id, participant_ids)
    else:
        generate_knockout_bracket(config_id, participant_ids, req.target_slots)
        
    return {"status": "success", "config_id": config_id}

@app.get("/api/tournament/{config_id}")
def get_tournament_data(config_id: int):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM cluster_configs WHERE id = ?", (config_id,))
    config = dict(cursor.fetchone())
    cursor.execute("""
        SELECT m.*, pa.name as player_a_name, pb.name as player_b_name, w.name as winner_name
        FROM matches m
        LEFT JOIN participants pa ON m.participant_a_id = pa.id
        LEFT JOIN participants pb ON m.participant_b_id = pb.id
        LEFT JOIN participants w ON m.winner_id = w.id
        WHERE m.config_id = ? ORDER BY m.round_number ASC, m.match_index ASC
    """, (config_id,))
    matches = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return {"config": config, "matches": matches}

@app.post("/api/update-score")
def update_score(req: ScoreUpdateRequest):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT participant_a_id, participant_b_id, next_match_id, next_match_slot FROM matches WHERE id = ?", (req.match_id,))
    p_a, p_b, next_match_id, next_match_slot = cursor.fetchone()
    
    if req.score_a > req.score_b:
        winner_id = p_a
    elif req.score_b > req.score_a:
        winner_id = p_b
    else:
        winner_id = None
        
    cursor.execute("UPDATE matches SET score_a = ?, score_b = ?, winner_id = ? WHERE id = ?", (req.score_a, req.score_b, winner_id, req.match_id))
    
    if next_match_id and winner_id:
        col = "participant_a_id" if next_match_slot == 'a' else "participant_b_id"
        cursor.execute(f"UPDATE matches SET {col} = ? WHERE id = ?", (winner_id, next_match_id))
        
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.get("/", response_class=HTMLResponse)
def index_page():
    return """
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <title>Đại Hội Mường Thanh V2</title>
        <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
        <script src="https://cdn.tailwindcss.com"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
        <style>
            .bracket-container { display: flex; gap: 40px; padding: 20px; overflow-x: auto; min-height: 500px; background: white; }
            .round-column { display: flex; flex-direction: column; justify-content: space-around; min-width: 240px; }
            .match-box { background: #ffffff; border: 2px solid #e2e8f0; border-radius: 8px; padding: 10px; margin: 15px 0; position: relative; }
            .match-box::after { content: ''; position: absolute; right: -22px; top: 50%; width: 22px; height: 2px; background: #cbd5e1; z-index: -1; }
            .round-column:last-child .match-box::after { display: none; }
            .slot-drag { padding: 6px; border-radius: 4px; margin: 2px 0; border: 1px solid transparent; cursor: grab; }
            .slot-drag:active { cursor: grabbing; }
            .slot-drag.drag-over { border: 2px dashed #1e3a8a; background: #eff6ff; }
            .tab-active { border-bottom: 4px solid #1e3a8a; color: #1e3a8a; font-weight: bold; }
        </style>
    </head>
    <body class="bg-gray-50">
        <header class="bg-blue-900 text-white p-4 flex justify-between items-center shadow">
            <h1 class="text-xl font-bold uppercase tracking-wider">Hệ Thống Quản Lý Giải Đấu V2</h1>
            <div class="flex space-x-4">
                <button onclick="switchTab('tab_mindmap')" id="btn_tab_mindmap" class="px-4 py-2 bg-blue-800 rounded font-bold hover:bg-blue-700">🧠 Sơ Đồ Tổng Quan</button>
                <button onclick="switchTab('tab_bracket')" id="btn_tab_bracket" class="px-4 py-2 bg-blue-800 rounded font-bold hover:bg-blue-700">🏆 Quản Lý Cặp Đấu</button>
                <button onclick="switchTab('tab_qualified')" id="btn_tab_qualified" class="px-4 py-2 bg-blue-800 rounded font-bold hover:bg-blue-700">🎟️ Danh Sách VCK</button>
                <button onclick="switchTab('tab_search')" id="btn_tab_search" class="px-4 py-2 bg-blue-800 rounded font-bold hover:bg-blue-700">🔍 Tra Cứu Đăng Ký</button>
            </div>
        </header>

        <div id="tab_bracket" class="p-6 grid grid-cols-1 lg:grid-cols-4 gap-6">
            <div class="bg-white p-5 rounded shadow border lg:col-span-1">
                <h2 class="font-bold text-blue-900 mb-4 border-b pb-2">⚙️ 1. Lọc Dữ Liệu Tự Động</h2>
                <select id="sport_name" class="w-full p-2 border rounded mb-3" onchange="fetchParticipants()"></select>
                <div id="cluster_checkboxes" class="grid grid-cols-2 gap-2 bg-gray-50 p-2 rounded border mb-3"></div>
                
                <h2 class="font-bold text-blue-900 mb-2 border-b pb-2 mt-4">2. Thiết Lập</h2>
                <label class="block text-sm mb-1 mt-3">Hình thức thi đấu:</label>
                <select id="format_type" class="w-full p-2 border rounded mb-3">
                    <option value="knockout">Loại trực tiếp (Cây nhánh)</option>
                    <option value="round_robin">Đấu vòng tròn (Tính điểm)</option>
                </select>
                <label class="block text-sm mb-1">Chỉ tiêu vào VCK:</label>
                <input type="number" id="target_slots" value="2" class="w-full p-2 border rounded mb-3">
                <label class="block text-sm mb-1">Danh sách đội (Auto load từ file):</label>
                <textarea id="participants_list" rows="6" class="w-full p-2 border rounded text-xs font-mono"></textarea>
                
                <button onclick="setupTournament()" class="w-full bg-blue-800 text-white p-2 rounded mt-3 font-bold">⚡ TẠO SƠ ĐỒ MỚI</button>
                <button id="btn_save_draft" onclick="saveTournament()" class="w-full bg-green-600 text-white p-2 rounded mt-2 font-bold hidden">💾 LƯU SƠ ĐỒ ĐANG XEM</button>
                
                <h2 class="font-bold text-blue-900 mb-2 border-b pb-2 mt-6">📂 Sơ Đồ Đã Lưu</h2>
                <select id="saved_configs" class="w-full p-2 border rounded mb-2" onchange="loadBracket(this.value)">
                    <option value="">-- Chọn sơ đồ đã lưu --</option>
                </select>
                <div class="flex gap-2">
                    <button onclick="deleteConfig()" class="w-1/2 bg-red-600 text-white p-2 rounded text-sm font-bold">🗑 Xóa/Làm Lại</button>
                    <button onclick="exportToPDF()" class="w-1/2 bg-green-600 text-white p-2 rounded text-sm font-bold">📥 Xuất PDF</button>
                </div>
            </div>

            <div class="bg-white p-5 rounded shadow border lg:col-span-3 overflow-auto">
                <h2 id="current_info" class="text-lg font-bold mb-4 text-center text-blue-900">Khu Vực Sơ Đồ Cây</h2>
                <div id="bracket_workspace" class="bracket-container"></div>
            </div>
        </div>

        <div id="tab_mindmap" class="p-6 hidden h-full">
            <div class="bg-white p-6 rounded shadow border min-h-screen">
                <div class="flex justify-between mb-4 border-b pb-2">
                    <h2 class="text-xl font-bold text-blue-900">Sơ Đồ Tổng Quan Số Đội Theo Cụm</h2>
                    <div class="flex space-x-2">
                        <select id="mindmap_sport" class="p-2 border rounded font-bold text-blue-900 bg-gray-50 min-w-[200px]">
                            <!-- Options auto loaded -->
                        </select>
                        <button onclick="loadSavedMindMap()" class="bg-cyan-600 hover:bg-cyan-500 text-white px-4 py-2 rounded font-bold">Lấy sơ đồ đã lưu</button>
                        <button onclick="loadMindMap()" class="bg-blue-800 hover:bg-blue-700 text-white px-4 py-2 rounded font-bold">Vẽ Sơ Đồ</button>
                        <button id="btn_save_mindmap" onclick="saveFinalMindMap()" class="bg-orange-600 hover:bg-orange-500 text-white px-4 py-2 rounded font-bold hidden">Lưu sơ đồ</button>
                        <button onclick="exportMindMapPDF()" class="bg-green-600 hover:bg-green-500 text-white px-4 py-2 rounded font-bold">Xuất PDF</button>
                        <button onclick="exportAllMindMapsPDF()" class="bg-red-600 hover:bg-red-500 text-white px-4 py-2 rounded font-bold">Xuất tất cả ra PDF</button>
                        <button id="btn_merge_clusters" onclick="openMergeModal()" class="bg-yellow-600 hover:bg-yellow-500 text-white px-4 py-2 rounded font-bold hidden">Ghép Cụm</button>
                        <button id="btn_add_info" onclick="openAddInfoModal()" class="bg-purple-600 hover:bg-purple-500 text-white px-4 py-2 rounded font-bold hidden">Thêm thông tin</button>
                    </div>
                </div>
                <div id="mindmap_container" class="relative w-full rounded" style="min-height: 600px; background-color: #3b5c4f;">
                    <!-- Sơ đồ sẽ render ở đây -->
                </div>
            </div>
        </div>

        <div id="tab_search" class="p-6 hidden">
            <div class="bg-white p-6 rounded shadow border">
                <h2 class="text-xl font-bold text-blue-900 mb-4 border-b pb-2">Tra Cứu Thông Tin Đăng Ký Của Đơn Vị</h2>
                <div class="flex gap-4 mb-6">
                    <input type="text" id="search_unit_input" placeholder="Nhập tên khách sạn (VD: Xa La, Grand Hà Nội...)" class="w-1/2 p-3 border rounded">
                    <button onclick="searchUnit()" class="bg-blue-800 text-white px-6 py-2 rounded font-bold">Tìm Kiếm</button>
                </div>
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-blue-900 text-white">
                            <th class="p-3 border">Tên Đơn Vị</th>
                            <th class="p-3 border">Cụm</th>
                            <th class="p-3 border">Môn Đăng Ký</th>
                            <th class="p-3 border">Số Lượng Đội/VĐV</th>
                        </tr>
                    </thead>
                    <tbody id="search_results"></tbody>
                </table>
            </div>
        </div>

        <div id="tab_qualified" class="p-6 hidden">
            <div class="bg-white p-6 rounded shadow border">
                <h2 class="text-xl font-bold text-blue-900 mb-4 border-b pb-2">Danh Sách Các Đội Lọt Vào VCK</h2>
                <div class="mb-6 flex gap-4">
                    <div class="flex items-center">
                        <label class="font-bold mr-2 whitespace-nowrap">Lọc theo môn thi:</label>
                        <select id="qualified_sport_filter" class="p-2 border rounded" onchange="loadQualifiedTeams()">
                            <option value="">-- Tất cả các môn --</option>
                        </select>
                    </div>
                </div>
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-blue-900 text-white">
                            <th class="p-3 border">Môn Thi</th>
                            <th class="p-3 border">Cụm Tương Ứng</th>
                            <th class="p-3 border">Đội / VĐV Thắng Cuộc</th>
                            <th class="p-3 border">Đơn Vị Chủ Quản</th>
                        </tr>
                    </thead>
                    <tbody id="qualified_results"></tbody>
                </table>
            </div>
        </div>

        <script>
            let currentConfigId = null;

            async function initApp() {
                // Migrate from localStorage to Backend if exists
                const oldKeys = [];
                for(let i=0; i<localStorage.length; i++){
                    let k = localStorage.key(i);
                    if(k.startsWith('saved_final_mindmap_')) {
                        oldKeys.push(k);
                    }
                }
                if (oldKeys.length > 0) {
                    oldKeys.forEach(k => {
                        let sport = k.replace('saved_final_mindmap_', '');
                        let state = JSON.parse(localStorage.getItem(k));
                        fetch('/api/save-mindmap', {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify({sport: sport, state: state})
                        }).then(() => {
                            localStorage.removeItem(k);
                        });
                    });
                }

                const res = await fetch('/api/master-data');
                const data = await res.json();
                
                const sportSelect = document.getElementById('sport_name');
                const qualifiedSportFilter = document.getElementById('qualified_sport_filter');
                data.sports.forEach(s => {
                    sportSelect.add(new Option(s, s));
                    qualifiedSportFilter.add(new Option(s, s));
                });
                
                const clusterDiv = document.getElementById('cluster_checkboxes');
                data.clusters.forEach(c => {
                    clusterDiv.innerHTML += `<label class="flex items-center text-sm"><input type="checkbox" name="clusters" value="${c}" onchange="fetchParticipants()" class="mr-1"> ${c}</label>`;
                });

                loadSavedConfigs();
            }

            async function loadSavedConfigs() {
                const res = await fetch('/api/saved-configs');
                const configs = await res.json();
                const select = document.getElementById('saved_configs');
                select.innerHTML = '<option value="">-- Chọn sơ đồ đã lưu --</option>';
                configs.forEach(c => {
                    select.add(new Option(`${c.sport_name} - ${c.clusters_involved}`, c.id));
                });
            }

            async function fetchParticipants() {
                const sport = document.getElementById('sport_name').value;
                const clusters = Array.from(document.querySelectorAll('input[name="clusters"]:checked')).map(cb => cb.value).join(',');
                if(!sport || !clusters) return;

                const res = await fetch(`/api/generate-participants?sport=${encodeURIComponent(sport)}&clusters=${encodeURIComponent(clusters)}`);
                const data = await res.json();
                document.getElementById('participants_list').value = data.participants.join('\\n');
            }

            async function setupTournament() {
                const sport_name = document.getElementById('sport_name').value;
                const format_type = document.getElementById('format_type').value;
                const target_slots = parseInt(document.getElementById('target_slots').value);
                const participants = document.getElementById('participants_list').value.split('\\n').filter(p => p.trim());
                const clusters = Array.from(document.querySelectorAll('input[name="clusters"]:checked')).map(cb => cb.value);
                
                if (clusters.length === 0 || participants.length === 0) return Swal.fire('Lỗi', 'Chưa đủ dữ liệu!', 'error');

                const res = await fetch('/api/setup-tournament', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sport_name, clusters_involved: clusters, format_type, target_slots, participants })
                });
                const data = await res.json();
                if(res.ok) {
                    Swal.fire('Thành công', 'Đã khởi tạo sơ đồ!', 'success');
                    document.getElementById('saved_configs').value = "";
                    document.getElementById('btn_save_draft').classList.remove('hidden');
                    loadBracket(data.config_id);
                }
            }

            async function saveTournament() {
                if(!currentConfigId) return;
                const res = await fetch(`/api/save-tournament/${currentConfigId}`, { method: 'POST' });
                if(res.ok) {
                    Swal.fire('Thành công', 'Đã lưu sơ đồ vào danh sách chính thức!', 'success');
                    document.getElementById('btn_save_draft').classList.add('hidden');
                    await loadSavedConfigs();
                    document.getElementById('saved_configs').value = currentConfigId;
                }
            }

            async function loadBracket(config_id) {
                if(!config_id) return;
                currentConfigId = config_id;
                if (document.getElementById('saved_configs').value == config_id) {
                    document.getElementById('btn_save_draft').classList.add('hidden');
                }
                
                const res = await fetch(`/api/tournament/${config_id}`);
                const data = await res.json();
                
                document.getElementById('current_info').innerText = `${data.config.sport_name} | ${data.config.clusters_involved}`;
                const workspace = document.getElementById('bracket_workspace');
                workspace.innerHTML = '';
                workspace.className = '';

                if (data.config.format_type === 'round_robin') {
                    workspace.classList.add('flex', 'flex-col', 'lg:flex-row', 'gap-4', 'bg-white', 'p-4');
                    
                    let ptsData = {};
                    data.matches.forEach(m => {
                        if(m.participant_a_id) { if(!ptsData[m.participant_a_id]) ptsData[m.participant_a_id] = {name: m.player_a_name, pts:0, gd:0, gf:0, p:0}; }
                        if(m.participant_b_id) { if(!ptsData[m.participant_b_id]) ptsData[m.participant_b_id] = {name: m.player_b_name, pts:0, gd:0, gf:0, p:0}; }
                        
                        if(m.score_a !== null && m.score_b !== null) {
                            ptsData[m.participant_a_id].p++;
                            ptsData[m.participant_b_id].p++;
                            ptsData[m.participant_a_id].gf += m.score_a;
                            ptsData[m.participant_b_id].gf += m.score_b;
                            ptsData[m.participant_a_id].gd += (m.score_a - m.score_b);
                            ptsData[m.participant_b_id].gd += (m.score_b - m.score_a);
                            
                            if(m.score_a > m.score_b) ptsData[m.participant_a_id].pts += 3;
                            else if(m.score_a < m.score_b) ptsData[m.participant_b_id].pts += 3;
                            else { ptsData[m.participant_a_id].pts += 1; ptsData[m.participant_b_id].pts += 1; }
                        }
                    });

                    let sortedTeams = Object.values(ptsData).sort((a,b) => b.pts - a.pts || b.gd - a.gd || b.gf - a.gf);
                    
                    let tableHtml = `
                    <div class="w-full lg:w-1/2 border-r pr-4">
                        <h3 class="font-bold mb-4 text-blue-900 text-lg">Bảng Xếp Hạng</h3>
                        <table class="w-full border text-sm text-center">
                            <tr class="bg-blue-100"><th class="p-2 border">Hạng</th><th class="p-2 border text-left">Đội</th><th class="p-2 border">Trận</th><th class="p-2 border">HS</th><th class="p-2 border">Điểm</th></tr>
                            ${sortedTeams.map((t, idx) => `
                            <tr class="${idx < data.config.target_slots ? 'bg-green-50 font-bold' : ''}">
                                <td class="p-2 border">${idx+1}</td>
                                <td class="p-2 border text-left">${t.name} ${idx < data.config.target_slots ? '🏆' : ''}</td>
                                <td class="p-2 border">${t.p}</td>
                                <td class="p-2 border">${t.gd > 0 ? '+'+t.gd : t.gd}</td>
                                <td class="p-2 border text-blue-800">${t.pts}</td>
                            </tr>
                            `).join('')}
                        </table>
                    </div>`;

                    let matchesHtml = `<div class="w-full lg:w-1/2 pl-4"><h3 class="font-bold mb-4 text-blue-900 text-lg">Lịch Thi Đấu</h3><div class="grid gap-3">`;
                    data.matches.forEach(match => {
                        const nameA = match.player_a_name || '...';
                        const nameB = match.player_b_name || '...';
                        let scoreBox = '';
                        if (match.score_a !== null && match.score_b !== null) {
                            scoreBox = `<div class="text-center font-bold px-2 py-1 bg-gray-100 rounded cursor-pointer" onclick="openScoreModal(${match.id})">${match.score_a} - ${match.score_b}</div>`;
                        } else {
                            scoreBox = `<button onclick="openScoreModal(${match.id})" class="text-xs bg-gray-200 hover:bg-amber-400 px-3 py-1 rounded">Nhập điểm</button>`;
                        }
                        
                        matchesHtml += `
                            <div class="border rounded p-2 flex justify-between items-center shadow-sm">
                                <div class="text-[10px] text-gray-500 font-mono w-12">Vòng ${match.round_number}</div>
                                <div class="flex-1 text-right pr-3 font-semibold ${match.score_a > match.score_b ? 'text-green-700' : ''}">${nameA}</div>
                                ${scoreBox}
                                <div class="flex-1 text-left pl-3 font-semibold ${match.score_b > match.score_a ? 'text-green-700' : ''}">${nameB}</div>
                            </div>
                        `;
                    });
                    matchesHtml += `</div></div>`;
                    
                    workspace.innerHTML = tableHtml + matchesHtml;
                    return;
                }

                workspace.classList.add('bracket-container');

                const rounds = {};
                data.matches.forEach(m => {
                    if (!rounds[m.round_number]) rounds[m.round_number] = [];
                    rounds[m.round_number].push(m);
                });

                const roundNumbers = Object.keys(rounds).sort((a,b) => a-b);
                const lastRoundNum = roundNumbers.length > 0 ? Math.max(...roundNumbers) : 0;

                roundNumbers.forEach(rNum => {
                    const roundCol = document.createElement('div');
                    roundCol.className = 'round-column';
                    
                    rounds[rNum].forEach(match => {
                        const mBox = document.createElement('div');
                        mBox.className = 'match-box';
                        const nameA = match.player_a_name || '<span class="text-gray-300">Trống</span>';
                        const nameB = match.player_b_name || '<span class="text-gray-300">Trống</span>';
                        
                        const isRound1 = match.round_number === 1;
                        const attrDragA = isRound1 && match.participant_a_id ? `draggable="true" ondragstart="drag(event)" data-match-id="${match.id}" data-slot="a" data-participant-id="${match.participant_a_id}"` : '';
                        const attrDragB = isRound1 && match.participant_b_id ? `draggable="true" ondragstart="drag(event)" data-match-id="${match.id}" data-slot="b" data-participant-id="${match.participant_b_id}"` : '';
                        
                        const attrDropA = isRound1 ? `ondrop="drop(event)" ondragover="allowDrop(event)" ondragenter="dragEnter(event)" ondragleave="dragLeave(event)" data-match-id="${match.id}" data-slot="a" data-participant-id="${match.participant_a_id || ''}"` : '';
                        const attrDropB = isRound1 ? `ondrop="drop(event)" ondragover="allowDrop(event)" ondragenter="dragEnter(event)" ondragleave="dragLeave(event)" data-match-id="${match.id}" data-slot="b" data-participant-id="${match.participant_b_id || ''}"` : '';

                        mBox.innerHTML = `
                            <div class="text-[10px] text-gray-400 font-mono">M-${match.id}</div>
                            <div class="slot-drag ${match.winner_id === match.participant_a_id ? 'bg-green-100 font-bold' : ''}" ${attrDragA} ${attrDropA}>
                                ${nameA} <span class="float-right font-bold pointer-events-none">${match.score_a !== null ? match.score_a : '-'}</span>
                            </div>
                            <div class="border-t my-1 border-gray-100"></div>
                            <div class="slot-drag ${match.winner_id === match.participant_b_id ? 'bg-green-100 font-bold' : ''}" ${attrDragB} ${attrDropB}>
                                ${nameB} <span class="float-right font-bold pointer-events-none">${match.score_b !== null ? match.score_b : '-'}</span>
                            </div>
                            ${!match.winner_id && match.player_a_name && match.player_b_name ? `
                                <button onclick="openScoreModal(${match.id})" class="mt-2 w-full text-center bg-gray-100 hover:bg-amber-500 text-xs py-1 rounded">Nhập điểm</button>
                            ` : ''}
                        `;
                        roundCol.appendChild(mBox);
                    });
                    workspace.appendChild(roundCol);
                });

                if (lastRoundNum > 0) {
                    const targetCol = document.createElement('div');
                    targetCol.className = 'round-column';
                    
                    rounds[lastRoundNum].forEach(match => {
                        const tBox = document.createElement('div');
                        tBox.className = 'match-box bg-blue-50 border-blue-200 opacity-90';
                        
                        const winnerName = match.winner_name || '<span class="text-gray-400 italic">Chờ kết quả...</span>';
                        
                        tBox.innerHTML = `
                            <div class="text-[10px] text-blue-600 font-bold mb-1 uppercase">🎟️ VÉ ĐI VCK</div>
                            <div class="p-2 border border-blue-300 rounded bg-white text-center font-bold text-blue-900 shadow-sm flex items-center justify-center min-h-[40px]">
                                ${winnerName}
                            </div>
                        `;
                        targetCol.appendChild(tBox);
                    });
                    workspace.appendChild(targetCol);
                }
            }

            function allowDrop(ev) { ev.preventDefault(); }
            function dragEnter(ev) { ev.preventDefault(); ev.currentTarget.classList.add('drag-over'); }
            function dragLeave(ev) { ev.currentTarget.classList.remove('drag-over'); }
            function drag(ev) {
                ev.dataTransfer.setData("matchId", ev.currentTarget.getAttribute('data-match-id'));
                ev.dataTransfer.setData("slot", ev.currentTarget.getAttribute('data-slot'));
                ev.dataTransfer.setData("participantId", ev.currentTarget.getAttribute('data-participant-id'));
            }
            async function drop(ev) {
                ev.preventDefault();
                ev.currentTarget.classList.remove('drag-over');
                const fromMatchId = ev.dataTransfer.getData("matchId");
                const fromSlot = ev.dataTransfer.getData("slot");
                const toMatchId = ev.currentTarget.getAttribute('data-match-id');
                const toSlot = ev.currentTarget.getAttribute('data-slot');
                if (fromMatchId === toMatchId && fromSlot === toSlot) return;
                const res = await fetch('/api/swap-participant', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ from_match_id: parseInt(fromMatchId), from_slot: fromSlot, to_match_id: parseInt(toMatchId), to_slot: toSlot })
                });
                if(res.ok) loadBracket(currentConfigId);
            }

            async function openScoreModal(matchId) {
                const { value: formValues } = await Swal.fire({
                    title: 'Nhập tỷ số',
                    html: '<input id="swal-score-a" type="number" class="swal2-input w-24 m-1" placeholder="Đội 1"> - <input id="swal-score-b" type="number" class="swal2-input w-24 m-1" placeholder="Đội 2">',
                    preConfirm: () => [parseInt(document.getElementById('swal-score-a').value), parseInt(document.getElementById('swal-score-b').value)]
                });
                if (formValues) {
                    await fetch('/api/update-score', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ match_id: matchId, score_a: formValues[0], score_b: formValues[1] }) });
                    loadBracket(currentConfigId);
                }
            }

            async function deleteConfig() {
                if(!currentConfigId) return;
                if(confirm("Xóa toàn bộ sơ đồ này?")) {
                    await fetch(`/api/tournament/${currentConfigId}`, { method: 'DELETE' });
                    document.getElementById('bracket_workspace').innerHTML = '';
                    currentConfigId = null;
                    loadSavedConfigs();
                }
            }

            function exportToPDF() {
                const element = document.getElementById('bracket_workspace');
                const opt = { margin: 0.5, filename: 'So_Do_Thi_Dau.pdf', image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'a3', orientation: 'landscape' } };
                html2pdf().set(opt).from(element).save();
            }

            async function searchUnit() {
                const term = document.getElementById('search_unit_input').value;
                if(!term) return;
                const res = await fetch(`/api/search-unit?unit_name=${encodeURIComponent(term)}`);
                const data = await res.json();
                const tbody = document.getElementById('search_results');
                tbody.innerHTML = '';
                data.forEach(r => {
                    tbody.innerHTML += `
                        <tr class="border-b hover:bg-blue-50">
                            <td class="p-3 font-bold">${term}</td>
                            <td class="p-3">${r.cluster}</td>
                            <td class="p-3">${r.sport}</td>
                            <td class="p-3 text-center font-bold text-blue-800">${r.qty}</td>
                        </tr>
                    `;
                });
            }

            async function loadQualifiedTeams() {
                const sport = document.getElementById('qualified_sport_filter').value;
                let url = '/api/qualified-teams';
                if (sport) url += `?sport=${encodeURIComponent(sport)}`;
                
                const res = await fetch(url);
                const data = await res.json();
                const tbody = document.getElementById('qualified_results');
                tbody.innerHTML = '';
                
                if (data.length === 0) {
                    tbody.innerHTML = `<tr><td colspan="4" class="p-4 text-center text-gray-500 italic">Chưa có đội nào đạt vé VCK cho tiêu chí này.</td></tr>`;
                    return;
                }
                
                data.forEach(r => {
                    tbody.innerHTML += `
                        <tr class="border-b hover:bg-blue-50">
                            <td class="p-3 font-bold text-blue-900">${r.sport}</td>
                            <td class="p-3">${r.clusters}</td>
                            <td class="p-3 font-bold text-green-700">${r.player}</td>
                            <td class="p-3">${r.unit}</td>
                        </tr>
                    `;
                });
            }

            function switchTab(tabId) {
                document.getElementById('tab_mindmap').classList.add('hidden');
                document.getElementById('tab_bracket').classList.add('hidden');
                document.getElementById('tab_search').classList.add('hidden');
                document.getElementById('tab_qualified').classList.add('hidden');
                document.getElementById(tabId).classList.remove('hidden');

                document.getElementById('btn_tab_mindmap').classList.remove('bg-white', 'text-blue-900');
                document.getElementById('btn_tab_bracket').classList.remove('bg-white', 'text-blue-900');
                document.getElementById('btn_tab_search').classList.remove('bg-white', 'text-blue-900');
                document.getElementById('btn_tab_qualified').classList.remove('bg-white', 'text-blue-900');
                
                document.getElementById('btn_tab_mindmap').classList.add('text-white');
                document.getElementById('btn_tab_bracket').classList.add('text-white');
                document.getElementById('btn_tab_search').classList.add('text-white');
                document.getElementById('btn_tab_qualified').classList.add('text-white');

                const activeBtn = document.getElementById('btn_' + tabId);
                activeBtn.classList.remove('text-white');
                activeBtn.classList.add('bg-white', 'text-blue-900');
                
                if(tabId === 'tab_qualified') {
                    loadQualifiedTeams();
                } else if(tabId === 'tab_mindmap') {
                    // Populate mindmap sports dropdown if empty
                    const mmSelect = document.getElementById('mindmap_sport');
                    const bracketSelect = document.getElementById('sport_name');
                    if (mmSelect.options.length === 0 && bracketSelect.options.length > 0) {
                        mmSelect.innerHTML = bracketSelect.innerHTML;
                    }
                }
            }

            // --- MIND MAP LOGIC ---
            let mindmapData = null;
            let currentLeftClusters = [];
            let currentRightClusters = [];

            async function loadMindMap() {
                const sport = document.getElementById('mindmap_sport').value;
                if(!sport) {
                    Swal.fire('Lỗi', 'Chưa có môn thi để vẽ!', 'error');
                    return;
                }
                
                const res = await fetch(`/api/sport-overview?sport=${encodeURIComponent(sport)}`);
                mindmapData = await res.json();
                
                const clusters = Object.keys(mindmapData.clusters);
                const half = Math.ceil(clusters.length / 2);
                currentLeftClusters = clusters.slice(0, half);
                currentRightClusters = clusters.slice(half);
                
                renderMindMap();
            }

            function renderMindMap() {
                if(!mindmapData) return;
                document.getElementById('btn_merge_clusters').classList.remove('hidden');
                document.getElementById('btn_add_info').classList.remove('hidden');
                document.getElementById('btn_save_mindmap').classList.remove('hidden');
                
                // Auto load cluster info from localStorage
                if (!mindmapData.clusterInfo) mindmapData.clusterInfo = {};
                Object.keys(mindmapData.clusters).forEach(key => {
                    if (!mindmapData.clusterInfo[key]) {
                        const lsKey = 'mindmap_info_' + mindmapData.sport + '_' + key;
                        const saved = localStorage.getItem(lsKey);
                        if (saved) mindmapData.clusterInfo[key] = saved;
                    }
                });
                
                if (mindmapData.generalInfo === undefined) {
                    const genLsKey = 'mindmap_general_info_' + mindmapData.sport;
                    const savedGen = localStorage.getItem(genLsKey);
                    if (savedGen) mindmapData.generalInfo = savedGen;
                }
                
                const container = document.getElementById('mindmap_container');
                container.innerHTML = `
                    <canvas id="mindmap_canvas" class="absolute top-0 left-0 w-full h-full pointer-events-none" style="z-index: 0;"></canvas>
                    <div class="relative z-10 flex justify-center items-stretch w-full min-h-full py-8 px-8 gap-8 lg:gap-16">
                        <div id="mm_left" class="flex flex-col justify-around items-end gap-2"></div>
                        <div id="mm_center" class="flex flex-col justify-center items-center shrink-0 z-20"></div>
                        <div id="mm_right" class="flex flex-col justify-around items-start gap-2"></div>
                    </div>
                `;
                
                let totalTeams = Object.values(mindmapData.clusters).reduce((sum, arr) => sum + arr.length, 0);
                const centerNode = document.createElement('div');
                centerNode.className = "bg-green-700 text-white font-bold text-xl px-6 py-3 rounded border-2 border-yellow-400 shadow-lg text-center uppercase";
                centerNode.id = "mm_center_node";
                centerNode.innerHTML = `${mindmapData.sport}<div class="text-sm text-white normal-case font-normal italic mt-1">(Tổng số: ${totalTeams} đội)</div>`;
                document.getElementById('mm_center').appendChild(centerNode);
                
                function createClusterHTML(cName, side) {
                    const teams = mindmapData.clusters[cName];
                    const clusterTotal = teams.length;
                    
                    let addInfoHTML = '';
                    if (mindmapData.clusterInfo && mindmapData.clusterInfo[cName]) {
                        addInfoHTML = `<div class="text-sm font-normal mt-2 text-left whitespace-pre-wrap w-full">${mindmapData.clusterInfo[cName]}</div>`;
                    }
                    
                    let html = '';
                    if (side === 'left') {
                        html = `
                            <div class="flex items-center justify-end gap-3 w-full group cursor-move my-1" draggable="true" ondragstart="mmDragStart(event)" ondrop="mmDrop(event)" ondragover="mmDragOver(event)" data-cluster="${cName}">
                                <div class="flex flex-col items-end">
                                    ${teams.map(t => `<div class="text-white text-xs whitespace-nowrap mb-1 font-mono border-b border-gray-400/50 pb-0.5 text-right w-full">${t}</div>`).join('')}
                                </div>
                                <div class="bg-yellow-400 text-blue-900 font-bold px-4 py-2 rounded shadow cluster-box z-10 shrink-0 flex flex-col items-center justify-center text-center" id="box_${cName.replace(/\s+/g, '')}">
                                    <span>${cName}</span>
                                    <span class="text-sm font-normal italic mt-1">(Tổng số: ${clusterTotal} đội)</span>
                                    ${addInfoHTML}
                                </div>
                            </div>
                        `;
                    } else {
                        html = `
                            <div class="flex items-center justify-start gap-3 w-full group cursor-move my-1" draggable="true" ondragstart="mmDragStart(event)" ondrop="mmDrop(event)" ondragover="mmDragOver(event)" data-cluster="${cName}">
                                <div class="bg-yellow-400 text-blue-900 font-bold px-4 py-2 rounded shadow cluster-box z-10 shrink-0 flex flex-col items-center justify-center text-center" id="box_${cName.replace(/\s+/g, '')}">
                                    <span>${cName}</span>
                                    <span class="text-sm font-normal italic mt-1">(Tổng số: ${clusterTotal} đội)</span>
                                    ${addInfoHTML}
                                </div>
                                <div class="flex flex-col items-start">
                                    ${teams.map(t => `<div class="text-white text-xs whitespace-nowrap mb-1 font-mono border-b border-gray-400/50 pb-0.5 text-left w-full">${t}</div>`).join('')}
                                </div>
                            </div>
                        `;
                    }
                    return html;
                }
                
                currentLeftClusters.forEach(c => {
                    document.getElementById('mm_left').insertAdjacentHTML('beforeend', createClusterHTML(c, 'left'));
                });
                
                currentRightClusters.forEach(c => {
                    document.getElementById('mm_right').insertAdjacentHTML('beforeend', createClusterHTML(c, 'right'));
                });
                
                if (mindmapData.generalInfo) {
                    const generalInfoDiv = document.createElement('div');
                    generalInfoDiv.className = "absolute bottom-8 right-8 border-2 border-gray-400/50 p-4 text-white text-sm font-medium whitespace-pre-wrap max-w-sm rounded shadow-lg z-10 text-left";
                    generalInfoDiv.style.backgroundColor = "transparent";
                    generalInfoDiv.innerText = mindmapData.generalInfo;
                    container.appendChild(generalInfoDiv);
                }
                
                setTimeout(drawMindMapLines, 100);
            }

            let draggedClusterNode = null;
            function mmDragStart(ev) {
                draggedClusterNode = ev.currentTarget;
                ev.dataTransfer.effectAllowed = 'move';
            }
            function mmDragOver(ev) {
                ev.preventDefault();
            }

            function mmDrop(ev) {
                ev.preventDefault();
                if(!draggedClusterNode) return;
                let targetNode = ev.currentTarget;
                
                const targetCluster = targetNode.getAttribute('data-cluster');
                const dragCluster = draggedClusterNode.getAttribute('data-cluster');
                if(!targetCluster || targetCluster === dragCluster) return;
                
                let dArr = currentLeftClusters.includes(dragCluster) ? currentLeftClusters : currentRightClusters;
                let tArr = currentLeftClusters.includes(targetCluster) ? currentLeftClusters : currentRightClusters;
                
                let dIdx = dArr.indexOf(dragCluster);
                let tIdx = tArr.indexOf(targetCluster);
                
                dArr[dIdx] = targetCluster;
                tArr[tIdx] = dragCluster;
                
                renderMindMap();
            }

            function drawMindMapLines() {
                const canvas = document.getElementById('mindmap_canvas');
                const center = document.getElementById('mm_center_node');
                const container = document.getElementById('mindmap_container');
                if(!canvas || !center || !container) return;
                
                const contRect = container.getBoundingClientRect();
                
                // Set canvas internal resolution
                canvas.width = contRect.width;
                canvas.height = contRect.height;
                
                const ctx = canvas.getContext('2d');
                ctx.clearRect(0, 0, canvas.width, canvas.height);
                
                const cRect = center.getBoundingClientRect();
                
                const cx = cRect.left + cRect.width/2 - contRect.left;
                const cy = cRect.top + cRect.height/2 - contRect.top;
                
                ctx.strokeStyle = "white";
                ctx.lineWidth = 2;
                
                document.querySelectorAll('.cluster-box').forEach(box => {
                    const bRect = box.getBoundingClientRect();
                    const isLeft = (bRect.left < cRect.left);
                    
                    let startX = isLeft ? cRect.left - contRect.left : cRect.right - contRect.left;
                    let startY = cy;
                    let endX = isLeft ? (bRect.right - contRect.left) : (bRect.left - contRect.left);
                    let endY = bRect.top + bRect.height/2 - contRect.top;
                    
                    let cp1x = isLeft ? startX - 40 : startX + 40;
                    let cp1y = startY;
                    let cp2x = isLeft ? endX + 40 : endX - 40;
                    let cp2y = endY;
                    
                    ctx.beginPath();
                    ctx.moveTo(startX, startY);
                    ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, endX, endY);
                    ctx.stroke();
                });
            }

            window.addEventListener('resize', () => {
                if(!document.getElementById('tab_mindmap').classList.contains('hidden')) {
                    drawMindMapLines();
                }
            });

            async function exportMindMapPDF() {
                const element = document.getElementById('mindmap_container');
                drawMindMapLines();
                
                Swal.fire({
                    title: 'Đang xuất PDF...',
                    text: 'Vui lòng chờ...',
                    allowOutsideClick: false,
                    didOpen: () => { Swal.showLoading(); }
                });
                
                const canvas = await html2canvas(element, { scale: 2, useCORS: true });
                const imgData = canvas.toDataURL('image/jpeg', 1.0);
                
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });
                
                const pdfWidth = 420;
                const pdfHeight = 297;
                const imgWidth = canvas.width;
                const imgHeight = canvas.height;
                const ratio = Math.min(pdfWidth / imgWidth, pdfHeight / imgHeight);
                
                const finalWidth = imgWidth * ratio;
                const finalHeight = imgHeight * ratio;
                const x = (pdfWidth - finalWidth) / 2;
                const y = (pdfHeight - finalHeight) / 2;
                
                pdf.setFillColor(59, 92, 79);
                pdf.rect(0, 0, pdfWidth, pdfHeight, 'F');
                pdf.addImage(imgData, 'JPEG', x, y, finalWidth, finalHeight);
                
                pdf.save('SoDoTongQuan.pdf');
                Swal.close();
            }

            async function exportAllMindMapsPDF() {
                const res = await fetch('/api/load-all-mindmaps');
                const allSaved = await res.json();
                const keys = Object.keys(allSaved);
                
                if(keys.length === 0) {
                    Swal.fire('Thông báo', 'Chưa có sơ đồ nào được lưu! Bạn cần "Lưu sơ đồ" trước khi xuất tất cả.', 'info');
                    return;
                }

                const mmSelect = document.getElementById('mindmap_sport');
                const orderMap = {};
                for (let i = 0; i < mmSelect.options.length; i++) {
                    orderMap[mmSelect.options[i].value] = i;
                }
                keys.sort((a, b) => (orderMap[a] ?? 999) - (orderMap[b] ?? 999));

                Swal.fire({
                    title: 'Đang xuất PDF...',
                    text: 'Vui lòng chờ, tiến trình có thể mất một lúc...',
                    allowOutsideClick: false,
                    didOpen: () => { Swal.showLoading(); }
                });

                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });
                const pdfWidth = 420;
                const pdfHeight = 297;
                
                const originalSport = document.getElementById('mindmap_sport').value;
                const element = document.getElementById('mindmap_container');
                
                let isFirst = true;
                for(let i=0; i<keys.length; i++){
                    let k = keys[i];
                    let state = allSaved[k];
                    mindmapData = state.mindmapData;
                    currentLeftClusters = state.left;
                    currentRightClusters = state.right;
                    
                    document.getElementById('mindmap_sport').value = mindmapData.sport;
                    renderMindMap();
                    await new Promise(r => setTimeout(r, 800)); // wait for render
                    drawMindMapLines();
                    await new Promise(r => setTimeout(r, 200)); 
                    
                    const canvas = await html2canvas(element, { scale: 2, useCORS: true });
                    const imgData = canvas.toDataURL('image/jpeg', 1.0);
                    
                    if (!isFirst) pdf.addPage();
                    isFirst = false;
                    
                    const imgWidth = canvas.width;
                    const imgHeight = canvas.height;
                    const ratio = Math.min(pdfWidth / imgWidth, pdfHeight / imgHeight);
                    
                    const finalWidth = imgWidth * ratio;
                    const finalHeight = imgHeight * ratio;
                    const x = (pdfWidth - finalWidth) / 2;
                    const y = (pdfHeight - finalHeight) / 2;
                    
                    pdf.setFillColor(59, 92, 79);
                    pdf.rect(0, 0, pdfWidth, pdfHeight, 'F');
                    pdf.addImage(imgData, 'JPEG', x, y, finalWidth, finalHeight);
                }
                
                pdf.save('TatCaSoDo.pdf');
                
                if(originalSport) {
                    document.getElementById('mindmap_sport').value = originalSport;
                    const currentStr = allSaved[originalSport];
                    if (currentStr) {
                        mindmapData = currentStr.mindmapData;
                        currentLeftClusters = currentStr.left;
                        currentRightClusters = currentStr.right;
                        renderMindMap();
                    }
                }
                
                Swal.close();
                Swal.fire('Thành công', 'Đã xuất thành công file PDF gồm tất cả sơ đồ!', 'success');
            }

            function openMergeModal() {
                if (!mindmapData || !mindmapData.clusters) return;
                const clusterKeys = Object.keys(mindmapData.clusters);
                if (clusterKeys.length < 2) {
                    Swal.fire('Thông báo', 'Cần ít nhất 2 cụm để có thể ghép.', 'info');
                    return;
                }
                
                let selectOptions = '<option value="">-- Không ghép --</option>';
                const maxGroups = Math.max(2, Math.floor(clusterKeys.length / 2));
                for(let i=1; i<=maxGroups; i++) {
                    selectOptions += `<option value="Nhóm ${i}">Ghép vào Nhóm ${i}</option>`;
                }

                let formHtml = '<div class="text-left mt-4" style="max-height: 400px; overflow-y: auto;">';
                clusterKeys.forEach((key, index) => {
                    formHtml += `
                        <div class="mb-3 flex items-center justify-between">
                            <span class="text-gray-700 font-medium w-1/2">${key}</span>
                            <select id="merge_group_${index}" class="form-select w-1/2 border rounded p-1 merge-select" data-key="${key}">
                                ${selectOptions}
                            </select>
                        </div>
                    `;
                });
                formHtml += '</div>';

                Swal.fire({
                    title: 'Chọn các Cụm để ghép',
                    html: formHtml,
                    showCancelButton: true,
                    confirmButtonText: 'Ghép Cụm',
                    cancelButtonText: 'Hủy',
                    preConfirm: () => {
                        const selects = document.querySelectorAll('.merge-select');
                        let groups = {};
                        selects.forEach(sel => {
                            const val = sel.value;
                            const key = sel.getAttribute('data-key');
                            if(val) {
                                if(!groups[val]) groups[val] = [];
                                groups[val].push(key);
                            }
                        });
                        
                        for(let g in groups) {
                            if(groups[g].length < 2) {
                                Swal.showValidationMessage(`Vui lòng chọn ít nhất 2 cụm cho ${g}`);
                                return false;
                            }
                        }
                        
                        if(Object.keys(groups).length === 0) {
                            Swal.showValidationMessage('Bạn chưa chọn nhóm nào để ghép');
                            return false;
                        }
                        
                        return groups;
                    }
                }).then((result) => {
                    if (result.isConfirmed) {
                        const groups = result.value;
                        
                        for(let g in groups) {
                            const selectedClusters = groups[g];
                            const newClusterName = selectedClusters.join(' & ');
                            
                            let combinedTeams = [];
                            selectedClusters.forEach(key => {
                                combinedTeams = combinedTeams.concat(mindmapData.clusters[key]);
                                delete mindmapData.clusters[key];
                            });
                            
                            mindmapData.clusters[newClusterName] = combinedTeams;
                        }
                        
                        const updatedKeys = Object.keys(mindmapData.clusters);
                        const half = Math.ceil(updatedKeys.length / 2);
                        currentLeftClusters = updatedKeys.slice(0, half);
                        currentRightClusters = updatedKeys.slice(half);
                        
                        renderMindMap();
                        Swal.fire('Thành công', 'Đã ghép cụm thành công!', 'success');
                    }
                });
            }

            function openAddInfoModal() {
                if (!mindmapData || !mindmapData.clusters) return;
                const clusterKeys = Object.keys(mindmapData.clusters);
                
                if (!mindmapData.clusterInfo) {
                    mindmapData.clusterInfo = {};
                }

                let currentGeneralInfo = mindmapData.generalInfo || '';
                if (!currentGeneralInfo) {
                    const lsKey = 'mindmap_general_info_' + mindmapData.sport;
                    currentGeneralInfo = localStorage.getItem(lsKey) || '';
                }

                let formHtml = '<div class="text-left mt-4" style="max-height: 400px; overflow-y: auto;">';
                
                formHtml += `
                    <div class="mb-6 border-b-2 border-gray-200 pb-4">
                        <label class="block text-purple-700 font-bold mb-1 uppercase">THÔNG TIN CHUNG</label>
                        <textarea id="info_general" class="w-full border-2 border-purple-300 rounded p-2 text-sm" rows="3" placeholder="Nhập thông tin chung cho môn thi...">${currentGeneralInfo}</textarea>
                    </div>
                `;
                
                clusterKeys.forEach(key => {
                    let currentInfo = mindmapData.clusterInfo[key] || '';
                    if (!currentInfo) {
                        const lsKey = 'mindmap_info_' + mindmapData.sport + '_' + key;
                        currentInfo = localStorage.getItem(lsKey) || '';
                    }
                    
                    formHtml += `
                        <div class="mb-4">
                            <label class="block text-gray-700 font-bold mb-1">${key}</label>
                            <textarea id="info_${key.replace(/\s+/g, '')}" class="w-full border rounded p-2 text-sm" rows="3" placeholder="Nhập thông tin bổ sung cho ${key}...">${currentInfo}</textarea>
                        </div>
                    `;
                });
                formHtml += '</div>';

                Swal.fire({
                    title: 'Thêm thông tin cho các Cụm',
                    html: formHtml,
                    width: 600,
                    showCancelButton: true,
                    confirmButtonText: 'Lưu thông tin',
                    cancelButtonText: 'Hủy',
                    preConfirm: () => {
                        let results = {};
                        const genEl = document.getElementById('info_general');
                        results['generalInfo'] = genEl ? genEl.value.trim() : '';
                        
                        clusterKeys.forEach(key => {
                            const el = document.getElementById('info_' + key.replace(/\s+/g, ''));
                            if (el) {
                                results[key] = el.value.trim();
                            }
                        });
                        return results;
                    }
                }).then((result) => {
                    if (result.isConfirmed) {
                        const infoData = result.value;
                        
                        mindmapData.generalInfo = infoData['generalInfo'];
                        const genLsKey = 'mindmap_general_info_' + mindmapData.sport;
                        if (mindmapData.generalInfo) {
                            localStorage.setItem(genLsKey, mindmapData.generalInfo);
                        } else {
                            localStorage.removeItem(genLsKey);
                        }
                        
                        clusterKeys.forEach(key => {
                            mindmapData.clusterInfo[key] = infoData[key];
                            const lsKey = 'mindmap_info_' + mindmapData.sport + '_' + key;
                            if (infoData[key]) {
                                localStorage.setItem(lsKey, infoData[key]);
                            } else {
                                localStorage.removeItem(lsKey);
                            }
                        });
                        renderMindMap();
                        Swal.fire('Thành công', 'Đã lưu thông tin bổ sung!', 'success');
                    }
                });
            }

            async function saveFinalMindMap() {
                if (!mindmapData) return;
                const sport = mindmapData.sport;
                const state = {
                    mindmapData: mindmapData,
                    left: currentLeftClusters,
                    right: currentRightClusters
                };
                
                await fetch('/api/save-mindmap', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({sport: sport, state: state})
                });
                
                Swal.fire('Thành công', 'Đã lưu sơ đồ hoàn thiện cho môn ' + sport, 'success');
            }

            async function loadSavedMindMap() {
                const sport = document.getElementById('mindmap_sport').value;
                if(!sport) {
                    Swal.fire('Lỗi', 'Chưa có môn thi để chọn!', 'error');
                    return;
                }
                
                try {
                    const res = await fetch(`/api/load-mindmap?sport=${encodeURIComponent(sport)}`);
                    const json = await res.json();
                    if (!json.state) {
                        Swal.fire('Thông báo', 'Chưa có sơ đồ nào được lưu cho môn này!', 'info');
                        return;
                    }
                    const state = json.state;
                    mindmapData = state.mindmapData;
                    currentLeftClusters = state.left;
                    currentRightClusters = state.right;
                    renderMindMap();
                    Swal.fire('Thành công', 'Đã tải sơ đồ được lưu!', 'success');
                } catch(e) {
                    Swal.fire('Lỗi', 'Không thể đọc dữ liệu đã lưu trên server.', 'error');
                }
            }

            initApp();
        </script>
    </body>
    </html>
    """

class SwapRequest(BaseModel):
    from_match_id: int
    from_slot: str
    to_match_id: int
    to_slot: str

@app.post("/api/swap-participant")
def swap_participant(req: SwapRequest):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # Get current participants
    cursor.execute(f"SELECT participant_{req.from_slot}_id FROM matches WHERE id = ?", (req.from_match_id,))
    from_p = cursor.fetchone()[0]
    cursor.execute(f"SELECT participant_{req.to_slot}_id FROM matches WHERE id = ?", (req.to_match_id,))
    to_p_row = cursor.fetchone()
    to_p = to_p_row[0] if to_p_row else None
    
    # Swap
    cursor.execute(f"UPDATE matches SET participant_{req.from_slot}_id = ? WHERE id = ?", (to_p, req.from_match_id))
    cursor.execute(f"UPDATE matches SET participant_{req.to_slot}_id = ? WHERE id = ?", (from_p, req.to_match_id))
    
    # Xử lý cập nhật điểm/người thắng nếu match chứa vị trí trống
    for m_id in [req.from_match_id, req.to_match_id]:
        cursor.execute("SELECT participant_a_id, participant_b_id, next_match_id, next_match_slot FROM matches WHERE id = ?", (m_id,))
        pa, pb, nm_id, n_slot = cursor.fetchone()
        if pa and not pb:
            cursor.execute("UPDATE matches SET winner_id = ?, score_a = 1, score_b = 0 WHERE id = ?", (pa, m_id))
            if nm_id:
                col = "participant_a_id" if n_slot == 'a' else "participant_b_id"
                cursor.execute(f"UPDATE matches SET {col} = ? WHERE id = ?", (pa, nm_id))
        elif pb and not pa:
            cursor.execute("UPDATE matches SET winner_id = ?, score_a = 0, score_b = 1 WHERE id = ?", (pb, m_id))
            if nm_id:
                col = "participant_a_id" if n_slot == 'a' else "participant_b_id"
                cursor.execute(f"UPDATE matches SET {col} = ? WHERE id = ?", (pb, nm_id))
        elif not pa and not pb:
            # reset winner
            cursor.execute("UPDATE matches SET winner_id = NULL, score_a = NULL, score_b = NULL WHERE id = ?", (m_id,))

    conn.commit()
    conn.close()
    return {"status": "success"}

if __name__ == "__main__":
    import uvicorn
    import multiprocessing
    import threading
    import webbrowser
    import time
    
    multiprocessing.freeze_support()
    
    def open_browser():
        time.sleep(1.5)
        webbrowser.open("http://127.0.0.1:8000")
        
    threading.Thread(target=open_browser, daemon=True).start()
    uvicorn.run(app, host="0.0.0.0", port=8000)
