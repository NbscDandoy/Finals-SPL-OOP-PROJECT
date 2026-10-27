import os
import re
import json
import uuid
import socket
import sqlite3
import datetime
import urllib.parse
from abc import ABC, abstractmethod
from typing import List, Optional, Dict, Any
from pathlib import Path

from fastapi import FastAPI, Request, File, UploadFile, HTTPException, Query
from fastapi.responses import JSONResponse, HTMLResponse, Response, FileResponse
from fastapi.staticfiles import StaticFiles
import socketio
import openpyxl
import io

# Optional qrcode dependency import with fallback safety
try:
    import qrcode
    HAS_QRCODE_LIB = True
except ImportError:
    HAS_QRCODE_LIB = False

# ==========================================
# 1. OOP DOMAIN MODELS & ABSTRACTION
# ==========================================

class Guest(ABC):
    """
    Abstract Base Class representing an event attendee.
    Demonstrates Abstraction and Encapsulation.
    """
    def __init__(self, name: str, seat_plan: str = "Unassigned", nickname: Optional[str] = None, guest_id: Optional[int] = None):
        self._id: Optional[int] = guest_id
        self._name: str = name.strip()
        self._nickname: str = nickname.strip() if nickname else ""
        self._seat_plan: str = seat_plan.strip() if seat_plan else "Unassigned"
        self._status: str = "Not Checked-In"
        self._check_in_time: Optional[str] = None
        self._qr_code: Optional[str] = None

    # Encapsulated Getters
    @property
    def guest_id(self) -> Optional[int]:
        return self._id

    @property
    def name(self) -> str:
        return self._name

    @property
    def nickname(self) -> str:
        return self._nickname

    @property
    def seat_plan(self) -> str:
        return self._seat_plan

    @property
    def status(self) -> str:
        return self._status

    @property
    def check_in_time(self) -> Optional[str]:
        return self._check_in_time

    @property
    def qr_code(self) -> Optional[str]:
        return self._qr_code

    def execute_check_in(self, time_str: str) -> bool:
        """Encapsulates check-in state modification."""
        if self._status == "Checked-In":
            return False
        self._status = "Checked-In"
        self._check_in_time = time_str
        return True

    def undo_check_in(self) -> None:
        self._status = "Not Checked-In"
        self._check_in_time = None

    @abstractmethod
    def get_category_type(self) -> str:
        """Polymorphic method overridden by subclasses."""
        pass

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self._id,
            "name": self._name,
            "nickname": self._nickname,
            "category": self.get_category_type(),
            "seat_plan": self._seat_plan,
            "status": self._status,
            "check_in_time": self._check_in_time,
            "qr_code": self._qr_code
        }


class RegularGuest(Guest):
    def get_category_type(self) -> str:
        return "Guest"


class VIPGuest(Guest):
    def get_category_type(self) -> str:
        return "VIP"


class WalkInGuest(Guest):
    def get_category_type(self) -> str:
        return "Walk-in Guest"


# ==========================================
# 2. DESIGN PATTERNS (FACTORY & REPOSITORY)
# ==========================================

class GuestFactory:
    """Factory Method Pattern for dynamic guest object creation."""
    @staticmethod
    def create_guest(category: str, name: str, seat_plan: str = "Unassigned", nickname: str = "", guest_id: Optional[int] = None) -> Guest:
        cat_upper = (category or "").upper()
        if "VIP" in cat_upper:
            return VIPGuest(name=name, seat_plan=seat_plan, nickname=nickname, guest_id=guest_id)
        elif "WALK" in cat_upper:
            return WalkInGuest(name=name, seat_plan=seat_plan, nickname=nickname, guest_id=guest_id)
        else:
            return RegularGuest(name=name, seat_plan=seat_plan, nickname=nickname, guest_id=guest_id)


class SQLiteGuestRepository:
    """Repository Pattern encapsulating SQLite persistence operations."""
    def __init__(self, db_path: str = "events.db"):
        self.db_path = db_path
        self._init_db()

    def _get_connection(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self) -> None:
        with self._get_connection() as conn:
            conn.execute("PRAGMA journal_mode = WAL;")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS guests (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT,
                    nickname TEXT,
                    category TEXT,
                    seat_plan TEXT,
                    status TEXT DEFAULT 'Not Checked-In',
                    check_in_time TEXT,
                    qr_code TEXT
                )
            """)

    def save(self, guest: Guest) -> int:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if guest.guest_id:
                cursor.execute("""
                    UPDATE guests 
                    SET name = ?, nickname = ?, category = ?, seat_plan = ?, status = ?, check_in_time = ?, qr_code = ?
                    WHERE id = ?
                """, (guest.name, guest.nickname, guest.get_category_type(), guest.seat_plan, guest.status, guest.check_in_time, guest.qr_code, guest.guest_id))
                return guest.guest_id
            else:
                cursor.execute("""
                    INSERT INTO guests (name, nickname, category, seat_plan, status, check_in_time, qr_code)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (guest.name, guest.nickname, guest.get_category_type(), guest.seat_plan, guest.status, guest.check_in_time, guest.qr_code))
                guest._id = cursor.lastrowid
                return cursor.lastrowid

    def update_qr_code(self, guest_id: int, qr_data_url: str) -> None:
        with self._get_connection() as conn:
            conn.execute("UPDATE guests SET qr_code = ? WHERE id = ?", (qr_data_url, guest_id))

    def update_status(self, guest_id: int, status: str, check_in_time: Optional[str]) -> int:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE guests SET status = ?, check_in_time = ? WHERE id = ?", (status, check_in_time, guest_id))
            return cursor.rowcount

    def find_by_id(self, guest_id: int) -> Optional[Guest]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM guests WHERE id = ?", (guest_id,))
            row = cursor.fetchone()
            if not row:
                return None
            return self._row_to_guest(row)

    def find_by_name(self, name: str) -> Optional[Guest]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM guests WHERE LOWER(name) = LOWER(?)", (name.strip(),))
            row = cursor.fetchone()
            if not row:
                return None
            return self._row_to_guest(row)

    def get_all(self) -> List[Guest]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM guests ORDER BY name ASC")
            rows = cursor.fetchall()
            return [self._row_to_guest(r) for r in rows]

    def delete_by_id(self, guest_id: int) -> None:
        with self._get_connection() as conn:
            conn.execute("DELETE FROM guests WHERE id = ?", (guest_id,))

    def clear_all(self) -> None:
        with self._get_connection() as conn:
            conn.execute("DELETE FROM guests")
            try:
                conn.execute("DELETE FROM sqlite_sequence WHERE name='guests'")
            except sqlite3.OperationalError:
                pass

    def get_dashboard_metrics(self) -> Dict[str, int]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    COUNT(*) AS total_guests,
                    SUM(CASE WHEN status = 'Checked-In' THEN 1 ELSE 0 END) AS checked_in,
                    SUM(CASE WHEN status = 'Not Checked-In' OR status IS NULL THEN 1 ELSE 0 END) AS not_checked_in,
                    SUM(CASE WHEN category = 'VIP' THEN 1 ELSE 0 END) AS vip_count
                FROM guests
            """)
            row = cursor.fetchone()
            return {
                "total_guests": row["total_guests"] or 0 if row else 0,
                "checked_in": row["checked_in"] or 0 if row else 0,
                "not_checked_in": row["not_checked_in"] or 0 if row else 0,
                "vip_count": row["vip_count"] or 0 if row else 0
            }

    def get_table_occupation(self) -> List[Dict[str, Any]]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    seat_plan AS table_name,
                    COUNT(*) AS total_guests,
                    SUM(CASE WHEN status = 'Checked-In' THEN 1 ELSE 0 END) AS checked_in_count
                FROM guests
                GROUP BY seat_plan
                ORDER BY seat_plan ASC
            """)
            rows = cursor.fetchall()
            return [{"table_name": r["table_name"] or "Unassigned", "total_guests": r["total_guests"], "checked_in_count": r["checked_in_count"]} for r in rows]

    def _row_to_guest(self, row: sqlite3.Row) -> Guest:
        guest = GuestFactory.create_guest(
            category=row["category"],
            name=row["name"],
            seat_plan=row["seat_plan"],
            nickname=row["nickname"],
            guest_id=row["id"]
        )
        guest._status = row["status"]
        guest._check_in_time = row["check_in_time"]
        guest._qr_code = row["qr_code"]
        return guest


# ==========================================
# 3. HELPER UTILITIES & GLOBAL STATE
# ==========================================

current_event_title = "Events By Anj Registration"

def get_philippine_time() -> str:
    pst_tz = datetime.timezone(datetime.timedelta(hours=8))
    now = datetime.datetime.now(pst_tz)
    return now.strftime("%Y-%m-%d %I:%M:%S %p")

def get_local_ip() -> str:
    """Finds the server's LAN IP address for local network access on mobile devices."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def generate_qr_code_base64(payload: Any) -> str:
    """Generates a base64 encoded PNG QR code with safety fallbacks."""
    import base64
    data_str = json.dumps(payload) if isinstance(payload, dict) else str(payload)

    if HAS_QRCODE_LIB:
        try:
            qr = qrcode.QRCode(version=1, box_size=10, border=2)
            qr.add_data(data_str)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode()
            return f"data:image/png;base64,{img_str}"
        except Exception as e:
            print(f"Warning: qrcode generation failed in server ({e}). Using web API fallback.")

    encoded_str = urllib.parse.quote(data_str)
    return f"https://api.qrserver.com/v1/create-qr-code/?size=220x220&data={encoded_str}"

def parse_excel_rows(sheet) -> List[Dict[str, str]]:
    """
    Dynamic Header Parser: Matches column names flexibly based on Row 1 headers.
    Allows uploading Excel files regardless of column order.
    """
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]

    def find_col_idx(keywords: List[str]) -> Optional[int]:
        for idx, h in enumerate(headers):
            if any(kw in h for kw in keywords):
                return idx
        return None

    name_idx = find_col_idx(["name", "guest", "fullname", "full name", "pangalan"])
    nickname_idx = find_col_idx(["nickname", "nick", "alias"])
    category_idx = find_col_idx(["category", "cat", "type", "group", "vip"])
    seat_idx = find_col_idx(["table", "seat", "seating", "table no", "table number", "plan"])

    if name_idx is None:
        name_idx = 0
    if nickname_idx is None:
        nickname_idx = 1 if len(headers) > 1 else None
    if category_idx is None:
        category_idx = 2 if len(headers) > 2 else None
    if seat_idx is None:
        seat_idx = 3 if len(headers) > 3 else None

    parsed_guests = []
    for row in rows[1:]:
        if not row:
            continue

        raw_name = row[name_idx] if name_idx < len(row) else None
        if raw_name is None or str(raw_name).strip() == "":
            continue

        name_val = str(raw_name).strip()
        nickname_val = str(row[nickname_idx]).strip() if nickname_idx is not None and nickname_idx < len(row) and row[nickname_idx] is not None else ""
        category_val = str(row[category_idx]).strip() if category_idx is not None and category_idx < len(row) and row[category_idx] is not None else "Guest"
        seat_val = str(row[seat_idx]).strip() if seat_idx is not None and seat_idx < len(row) and row[seat_idx] is not None else "Unassigned"

        parsed_guests.append({
            "name": name_val,
            "nickname": nickname_val,
            "category": category_val,
            "seat_plan": seat_val
        })

    return parsed_guests


# ==========================================
# 4. APP INITIALIZATION & SOCKET.IO SETUP
# ==========================================

repo = SQLiteGuestRepository("events.db")
upload_dir = Path(__file__).parent / "uploads"
upload_dir.mkdir(parents=True, exist_ok=True)

sio = socketio.AsyncServer(async_mode='asgi', cors_allowed_origins='*')
fastapi_app = FastAPI()
app = socketio.ASGIApp(sio, other_asgi_app=fastapi_app)

# Bypass Ngrok warning page for guest mobile browsers
@fastapi_app.middleware("http")
async def add_ngrok_skip_warning_header(request: Request, call_next):
    response = await call_next(request)
    response.headers["ngrok-skip-browser-warning"] = "true"
    return response

async def trigger_realtime_updates():
    await sio.emit('guestListReload')
    await sio.emit('tableOccupationUpdated', {'tables': repo.get_table_occupation()})
    metrics = repo.get_dashboard_metrics()
    await sio.emit('dashboardMetricsUpdated', metrics)


# ==========================================
# 5. SOCKET.IO EVENT HANDLERS
# ==========================================

@sio.event
async def connect(sid, environ):
    pass

@sio.event
async def checkIn(sid, data):
    guest_id = data.get("id")
    time_str = get_philippine_time()
    changes = repo.update_status(guest_id, "Checked-In", time_str)
    if changes > 0:
        await sio.emit("guestUpdated", {"id": guest_id, "status": "Checked-In", "check_in_time": time_str})
        await trigger_realtime_updates()

@sio.event
async def uncheckIn(sid, data):
    guest_id = data.get("id")
    changes = repo.update_status(guest_id, "Not Checked-In", None)
    if changes > 0:
        await sio.emit("guestUpdated", {"id": guest_id, "status": "Not Checked-In", "check_in_time": None})
        await trigger_realtime_updates()


# ==========================================
# 6. REST API ENDPOINTS
# ==========================================

@fastapi_app.get("/api/event-info")
async def get_event_info():
    """Returns current active event title for registration pages."""
    return {"success": True, "event_title": current_event_title}


@fastapi_app.post("/api/update-event-title")
async def update_event_title(request: Request):
    """Allows staff/organizer to customize event title for QR standee."""
    global current_event_title
    body = await request.json()
    new_title = body.get("title", "").strip()
    if new_title:
        current_event_title = new_title
        await trigger_realtime_updates()
        return {"success": True, "event_title": current_event_title}
    return JSONResponse(status_code=400, content={"success": False, "error": "Title is required."})


@fastapi_app.post("/api/import")
async def import_excel(file: UploadFile = File(...)):
    """Handles Excel file uploads using dynamic header parsing."""
    global current_event_title
    if not file.filename.endswith(('.xlsx', '.xls', '.csv')):
        return JSONResponse(status_code=400, content={"success": False, "error": "Invalid file type. Please upload Excel or CSV."})

    try:
        file_path = upload_dir / file.filename
        contents = await file.read()
        with open(file_path, "wb") as f:
            f.write(contents)

        clean_title = Path(file.filename).stem.replace("_", " ").replace("-", " ").title()
        current_event_title = clean_title

        repo.clear_all()
        
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
        sheet = wb.active

        guest_data_list = parse_excel_rows(sheet)
        imported_count = 0

        for item in guest_data_list:
            guest = GuestFactory.create_guest(
                category=item["category"],
                name=item["name"],
                seat_plan=item["seat_plan"],
                nickname=item["nickname"]
            )
            repo.save(guest)
            imported_count += 1

        await trigger_realtime_updates()

        return {
            "success": True, 
            "importedCount": imported_count, 
            "activeFile": file.filename,
            "event_title": current_event_title
        }
    except Exception as e:
        print(f"Excel import error: {e}")
        return JSONResponse(status_code=500, content={"success": False, "error": f"Error parsing Excel file: {str(e)}"})


@fastapi_app.post("/api/load-file")
async def load_file(request: Request):
    """Handles switching active files from the sidebar."""
    body = await request.json()
    file_name = body.get("fileName")
    if not file_name:
        raise HTTPException(status_code=400, detail="Filename required.")

    file_path = upload_dir / file_name
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found on server.")

    try:
        repo.clear_all()
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active

        guest_data_list = parse_excel_rows(sheet)
        imported_count = 0

        for item in guest_data_list:
            guest = GuestFactory.create_guest(
                category=item["category"],
                name=item["name"],
                seat_plan=item["seat_plan"],
                nickname=item["nickname"]
            )
            repo.save(guest)
            imported_count += 1

        await trigger_realtime_updates()
        return {"success": True, "importedCount": imported_count, "activeFile": file_name}
    except Exception as e:
        print(f"Error loading file {file_name}: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to process file: {str(e)}")


@fastapi_app.get("/api/recent-files")
async def get_recent_files():
    files = [f.name for f in upload_dir.glob("*") if f.is_file() and f.suffix in ['.xlsx', '.xls', '.csv']]
    return {"files": files}


@fastapi_app.delete("/api/recent-files/{file_name}")
async def delete_recent_file(file_name: str):
    file_path = upload_dir / file_name
    if file_path.exists():
        file_path.unlink()
    
    # Clears DB guests and updates dashboard in real-time when file is removed
    repo.clear_all()
    await trigger_realtime_updates()
    return {"success": True}


@fastapi_app.delete("/api/guests/clear-all")
async def clear_all_guests():
    """Clears all guest records from SQLite database."""
    repo.clear_all()
    await trigger_realtime_updates()
    return {"success": True, "message": "All guests cleared successfully."}


@fastapi_app.post("/api/register")
async def register_guest(request: Request):
    body = await request.json()
    name = body.get("name", "").strip()
    nickname = body.get("nickname", "").strip()
    category = body.get("category", "Walk-in Guest")
    seat_plan = body.get("seat_plan", "Unassigned")

    if not name:
        return JSONResponse(status_code=400, content={"success": False, "error": "Pakilagay ang iyong pangalan."})

    check_in_time = get_philippine_time()
    matched = repo.find_by_name(name)

    if matched:
        matched.execute_check_in(check_in_time)
        if nickname:
            matched._nickname = nickname
        repo.save(matched)
        await trigger_realtime_updates()
        return {
            "success": True,
            "message": f"Thank you for registering, {matched.name}!",
            "table_no": matched.seat_plan,
            "is_existing": True,
            "guest": matched.to_dict()
        }

    new_guest = GuestFactory.create_guest(category=category, name=name, seat_plan=seat_plan, nickname=nickname)
    new_guest.execute_check_in(check_in_time)
    new_id = repo.save(new_guest)

    qr_url = generate_qr_code_base64({"id": new_id, "name": name, "table": seat_plan})
    repo.update_qr_code(new_id, qr_url)
    await trigger_realtime_updates()

    return {
        "success": True,
        "message": f"Thank you for registering, {name}! Welcome to the event.",
        "table_no": seat_plan,
        "is_existing": False,
        "id": new_id
    }


@fastapi_app.get("/api/event-qrcode")
async def get_event_qrcode(request: Request):
    try:
        forwarded_host = request.headers.get("x-forwarded-host") or request.headers.get("host")
        forwarded_proto = request.headers.get("x-forwarded-proto") or request.url.scheme
        
        if forwarded_host and "ngrok" in forwarded_host.lower():
            host_url = f"{forwarded_proto}://{forwarded_host}/register.html"
        else:
            host_name = request.url.hostname
            port = request.url.port or 3000
            if host_name in ["localhost", "127.0.0.1"]:
                local_ip = get_local_ip()
                host_url = f"http://{local_ip}:{port}/register.html"
            else:
                host_url = f"{request.url.scheme}://{request.url.netloc}/register.html"

        qr_data_url = generate_qr_code_base64(host_url)
        return {
            "success": True, 
            "registrationUrl": host_url, 
            "qrCode": qr_data_url,
            "event_title": current_event_title
        }
    except Exception as e:
        print(f"Error in /api/event-qrcode: {e}")
        fallback_ip = get_local_ip()
        fallback_url = f"http://{fallback_ip}:3000/register.html"
        encoded = urllib.parse.quote(fallback_url)
        return {
            "success": True,
            "registrationUrl": fallback_url,
            "qrCode": f"https://api.qrserver.com/v1/create-qr-code/?size=220x220&data={encoded}",
            "event_title": current_event_title
        }


@fastapi_app.get("/api/dashboard-summary")
async def get_dashboard_summary():
    return repo.get_dashboard_metrics()


@fastapi_app.get("/api/guests")
async def get_guests():
    guests = repo.get_all()
    return {"guests": [g.to_dict() for g in guests]}


@fastapi_app.post("/api/guests")
async def create_guest(request: Request):
    body = await request.json()
    name = body.get("name")
    if not name:
        raise HTTPException(status_code=400, detail="Name is required.")

    guest = GuestFactory.create_guest(
        category=body.get("category", "Guest"),
        name=name,
        seat_plan=body.get("seat_plan", "Unassigned"),
        nickname=body.get("nickname", "")
    )
    new_id = repo.save(guest)
    qr_url = generate_qr_code_base64({"id": new_id, "name": name, "table": guest.seat_plan})
    repo.update_qr_code(new_id, qr_url)
    await trigger_realtime_updates()

    return {"success": True, "id": new_id, "qrCode": qr_url}


@fastapi_app.post("/api/guests/batch-checkin")
async def batch_checkin(request: Request):
    body = await request.json()
    ids = body.get("ids", [])
    status = body.get("status", "Checked-In")

    time_str = get_philippine_time() if status != "Not Checked-In" else None
    count = 0

    for gid in ids:
        count += repo.update_status(gid, status, time_str)

    await trigger_realtime_updates()
    return {"success": True, "updatedCount": count}


@fastapi_app.get("/api/table-occupation")
async def get_table_occupation():
    return {"tables": repo.get_table_occupation()}


@fastapi_app.get("/api/table-guests/{table_name}")
async def get_table_guests(table_name: str):
    guests = [g.to_dict() for g in repo.get_all() if g.seat_plan == table_name]
    return {"table_name": table_name, "guests": guests}


@fastapi_app.put("/api/guests/{guest_id}")
async def update_guest(guest_id: int, request: Request):
    body = await request.json()
    guest = repo.find_by_id(guest_id)
    if not guest:
        raise HTTPException(status_code=404, detail="Guest not found.")

    guest._name = body.get("name", guest.name)
    guest._nickname = body.get("nickname", guest.nickname)
    guest._seat_plan = body.get("seat_plan", guest.seat_plan)
    qr_url = generate_qr_code_base64({"id": guest_id, "name": guest.name, "table": guest.seat_plan})
    guest._qr_code = qr_url

    repo.save(guest)
    await trigger_realtime_updates()
    return {"success": True, "qrCode": qr_url}


@fastapi_app.delete("/api/guests/{guest_id}")
async def delete_guest(guest_id: int):
    repo.delete_by_id(guest_id)
    await trigger_realtime_updates()
    return {"success": True}


# Dynamic Static Files Delivery (Resolution for Frontend folder location)
possible_public_paths = [
    Path(__file__).parent.parent / "Frontend",
    Path(__file__).parent.parent / "Frontend" / "public",
    Path(__file__).parent / "Frontend",
    Path(__file__).parent / "public",
    Path(__file__).parent
]

static_dir = next((p for p in possible_public_paths if p.exists() and (p / "register.html").exists()), None)
if not static_dir:
    static_dir = next((p for p in possible_public_paths if p.exists()), possible_public_paths[0])

fastapi_app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="0.0.0.0", port=3000, reload=True)