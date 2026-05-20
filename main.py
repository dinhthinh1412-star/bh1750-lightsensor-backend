from fastapi import FastAPI, Request, Query
from fastapi.responses import JSONResponse, HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
import mysql.connector
from mysql.connector import pooling
from datetime import datetime, timedelta
from dotenv import load_dotenv
load_dotenv()
import io
import re
import random
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://dinhthinh1412-star.github.io"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

templates = Jinja2Templates(directory="web")

db_pool = pooling.MySQLConnectionPool(
    pool_name="mypool",
    pool_size=5,
    host=os.getenv("DB_HOST"),
    port=int(os.getenv("DB_PORT")),
    user=os.getenv("DB_USER"),
    password=os.getenv("DB_PASS"),
    database=os.getenv("DB_NAME"),
    ssl_ca="ca.pem"
)

def get_db_connection():
    conn = db_pool.get_connection()
    cursor = conn.cursor()
    cursor.execute("SET time_zone = '+07:00'")
    cursor.close()
    return conn


@app.post("/data")
async def receive_data(request: Request):
    db = None
    cursor = None
    try:
        data = await request.json()
        lux = data.get("lux")

        db = get_db_connection()
        cursor = db.cursor()

        cursor.execute(
            "INSERT INTO lux (lux_value) VALUES (%s)",
            (lux,)
        )
        db.commit()

        return JSONResponse({"status": "ok"})

    except Exception as e:
        print("DB ERROR:", e)
        return JSONResponse({"status": "error"}, status_code=500)

    finally:
        if cursor:
            cursor.close()
        if db:
            db.close()


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/api/export")
async def export_data(start_date: str, start_time: str, end_date: str, end_time: str, format: str = "txt"):
    db = get_db_connection()
    cursor = db.cursor()

    try:
        def parse_dt(d, t):
            return datetime.strptime(f"{d} {t}", "%d/%m/%Y %H:%M:%S")

        start_dt = parse_dt(start_date, start_time if start_time else "00:00:00")
        end_dt = parse_dt(end_date, end_time if end_time else "23:59:59")

        cursor.execute("""
            SELECT time, lux_value
            FROM lux
            WHERE time BETWEEN %s AND %s
            ORDER BY time ASC
        """, (start_dt, end_dt))

        rows = cursor.fetchall()

        safe_start = start_date.replace("/", "-")
        safe_end = end_date.replace("/", "-")

        if not rows:
            output = io.StringIO()
            output.write("Không có dữ liệu cảm biến trong khoảng thời gian này.\n")
            output.seek(0)
            headers = {"Content-Disposition": f'attachment; filename="lux_data_empty.txt"'}
            return StreamingResponse(output, media_type="text/plain", headers=headers)

        if format == "xlsx":
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment

            wb = Workbook()
            ws = wb.active
            ws.title = "Data"

            ws.append(["Date", "Time", "Lux"])

            for col_idx in range(1, 4):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='left')

            for r in rows:
                date_str = r[0].strftime('%d/%m/%Y')
                time_str = r[0].strftime('%H:%M:%S')
                lux_val = float(r[1])
                ws.append([date_str, time_str, lux_val])

            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 12

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            headers = {"Content-Disposition": f'attachment; filename="lux_data_{safe_start}_to_{safe_end}.xlsx"'}
            
            return StreamingResponse(output, media_type=media_type, headers=headers)

        else:
            output = io.StringIO()
            output.write(f"{'Time(YYYY/MM/DD HH:MM:SS)':<30}Lux\n")
            for r in rows:
                time_str = r[0].strftime('%Y/%m/%d %H:%M:%S')
                output.write(f"{time_str:<30}{r[1]}\n")
            
            output.seek(0)
            headers = {"Content-Disposition": f'attachment; filename="lux_data_{safe_start}_to_{safe_end}.txt"'}
            return StreamingResponse(output, media_type="text/plain", headers=headers)

    finally:
        cursor.close()
        db.close()


@app.get("/api/suggest-dates")
async def suggest_dates(
    q: str = "", 
    is_focused: bool = Query(False)
):
    if not is_focused:
        return JSONResponse({"results": [], "message": ""})

    if q and not re.match(r'^[\d/]+$', q):
        return JSONResponse({"results": [], "message": ""})

    db = get_db_connection()
    cursor = db.cursor()
    message = ""

    try:
        if not q:
            cursor.execute("SELECT DISTINCT DATE(time) FROM lux")
            rows = cursor.fetchall()
        else:
            parts = [p for p in q.split('/') if p.isdigit()]
            sql = "SELECT DISTINCT DATE(time) FROM lux WHERE 1=1"
            params = []
            
            if len(parts) >= 1:
                day = int(parts[0])
                if day < 1 or day > 31: return JSONResponse({"results": [], "message": ""})
                sql += " AND DAY(time) = %s"
                params.append(day)
            
            if len(parts) >= 2:
                month = int(parts[1])
                if month < 1 or month > 12: return JSONResponse({"results": [], "message": ""})
                sql += " AND MONTH(time) = %s"
                params.append(month)
                
            if len(parts) >= 3:
                year = int(parts[2])
                sql += " AND YEAR(time) = %s"
                params.append(year)

            cursor.execute(sql, tuple(params))
            rows = cursor.fetchall()

            if not rows:
                if len(parts) < 3:
                    cursor.execute("SELECT DISTINCT DATE(time) FROM lux")
                    rows = cursor.fetchall()
                    message = "Không có dữ liệu phù hợp, gợi ý gần đúng"
                else:
                    try:
                        target_date = datetime(year, month, day)
                        deltas = [30, 60, 90, 180, 365, 730, 1825, 3650]
                        found = False

                        for d_days in deltas:
                            start_d = target_date - timedelta(days=d_days)
                            end_d = target_date + timedelta(days=d_days)

                            cursor.execute(
                                "SELECT DISTINCT DATE(time) FROM lux WHERE DATE(time) BETWEEN %s AND %s",
                                (start_d.date(), end_d.date())
                            )
                            rows = cursor.fetchall()

                            if rows:
                                found = True
                                if d_days <= 90:
                                    message = f"Không có dữ liệu, gợi ý trong ± {d_days} ngày"
                                elif d_days <= 365:
                                    message = f"Không có dữ liệu, gợi ý trong ± {d_days // 30} tháng"
                                else:
                                    message = f"Không có dữ liệu, gợi ý trong ± {d_days // 365} năm"
                                break

                        if not found:
                            cursor.execute("SELECT DISTINCT DATE(time) FROM lux")
                            rows = cursor.fetchall()
                            message = "Không có dữ liệu, gợi ý toàn bộ"

                    except ValueError:
                        return JSONResponse({"results": [], "message": "Ngày không hợp lệ!"})

        # Lấy danh sách ngày thực tế, sắp xếp giảm dần (mới nhất lên đầu)
        raw_dates = [r[0] for r in rows if r[0]]
        raw_dates.sort(reverse=True)
        
        # Format lại thành chuỗi DD/MM/YYYY
        results = [d.strftime("%d/%m/%Y") for d in raw_dates]
        
        return JSONResponse({"results": results[:5], "message": message})

    except Exception as e:
        print("Suggest Dates Error:", e)
        return JSONResponse({"results": [], "message": ""})

    finally:
        cursor.close()
        db.close()


@app.get("/api/suggest-times")
async def suggest_times(
    date: str, 
    q: str = "", 
    is_focused: bool = Query(False)
):
    if not is_focused or not date:
        return JSONResponse({"results": [], "message": ""})

    if q and not re.match(r'^[\d:]+$', q):
        return JSONResponse({"results": [], "message": ""})

    db = get_db_connection()
    cursor = db.cursor()
    message = ""

    try:
        dt = datetime.strptime(date, "%d/%m/%Y").date()

        sql = "SELECT TIME(time) FROM lux WHERE DATE(time) = %s"
        params = [dt]

        if not q:
            cursor.execute(sql, tuple(params))
            rows = cursor.fetchall()
        else:
            parts = [p for p in q.split(':') if p.isdigit()]
            
            if len(parts) >= 1:
                hour = int(parts[0])
                if hour > 23: return JSONResponse({"results": [], "message": ""})
                sql += " AND HOUR(time) = %s"
                params.append(hour)
                
            if len(parts) >= 2:
                minute = int(parts[1])
                if minute > 59: return JSONResponse({"results": [], "message": ""})
                sql += " AND MINUTE(time) = %s"
                params.append(minute)
                
            if len(parts) >= 3:
                second = int(parts[2])
                if second > 59: return JSONResponse({"results": [], "message": ""})
                sql += " AND SECOND(time) = %s"
                params.append(second)

            cursor.execute(sql, tuple(params))
            rows = cursor.fetchall()

            if not rows:
                if len(parts) < 3:
                    cursor.execute("SELECT TIME(time) FROM lux WHERE DATE(time) = %s", (dt,))
                    rows = cursor.fetchall()
                    message = "Không có dữ liệu phù hợp, gợi ý gần đúng"
                else:
                    try:
                        target_dt = datetime(dt.year, dt.month, dt.day, hour, minute, second)
                        deltas_mins = [15, 30, 60, 180, 360, 720]
                        found = False
                        
                        for m in deltas_mins:
                            start_t = target_dt - timedelta(minutes=m)
                            end_t = target_dt + timedelta(minutes=m)
                            cursor.execute(
                                "SELECT TIME(time) FROM lux WHERE DATE(time) = %s AND time BETWEEN %s AND %s",
                                (dt, start_t, end_t)
                            )
                            rows = cursor.fetchall()
                            if rows:
                                found = True
                                if m < 60:
                                    message = f"Không có dữ liệu, gợi ý trong ± {m} phút"
                                else:
                                    message = f"Không có dữ liệu, gợi ý trong ± {m // 60} giờ"
                                break
                        
                        if not found:
                            cursor.execute("SELECT TIME(time) FROM lux WHERE DATE(time) = %s", (dt,))
                            rows = cursor.fetchall()
                            message = "Không có dữ liệu, gợi ý toàn bộ trong ngày"

                    except ValueError:
                        return JSONResponse({"results": [], "message": "Giờ không hợp lệ!"})

        def format_timedelta(td):
            total_seconds = int(td.total_seconds())
            hours, remainder = divmod(total_seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

        # Lấy danh sách giờ (timedelta), sắp xếp giảm dần (mới nhất lên đầu)
        raw_times = [r[0] for r in rows if r[0] is not None]
        raw_times.sort(reverse=True)
        
        # Format lại thành chuỗi HH:MM:SS
        times = [format_timedelta(td) for td in raw_times]
        
        return JSONResponse({"results": times[:5], "message": message})

    except Exception as e:
        print("Suggest Times Error:", e)
        return JSONResponse({"results": [], "message": ""})

    finally:
        cursor.close()
        db.close()


@app.get("/api/get-lux-by-time")
async def get_lux_by_time(date: str, time: str):
    if not date or not time:
        return JSONResponse({"lux": None})

    db = get_db_connection()
    cursor = db.cursor()

    try:
        dt = datetime.strptime(date, "%d/%m/%Y").date()

        cursor.execute("""
            SELECT lux_value
            FROM lux
            WHERE DATE(time) = %s
              AND TIME(time) = %s
            LIMIT 1
        """, (dt, time))

        row = cursor.fetchone()
        return JSONResponse({"lux": row[0] if row else None})

    finally:
        cursor.close()
        db.close()


@app.get("/get-stats")
async def get_stats(date: str = None):
    db = get_db_connection()
    cursor = db.cursor()

    try:
        where_clause = ""
        params = ()
        if date:
            where_clause = "WHERE DATE_FORMAT(time, '%d/%m/%Y') = %s"
            params = (date,)

        cursor.execute(f"SELECT lux_value, time FROM lux {where_clause} ORDER BY id DESC LIMIT 1", params)
        current = cursor.fetchone()
        current_val = current[0] if current else 0
        
        current_time_str = ""
        if current and current[1]:
            current_time_str = current[1].strftime('%H:%M:%S')

        cursor.execute(f"SELECT MAX(lux_value) FROM lux {where_clause}", params)
        max_res = cursor.fetchone()
        max_val = max_res[0] if max_res and max_res[0] is not None else 0

        cursor.execute(f"SELECT MIN(lux_value) FROM lux {where_clause}", params)
        min_res = cursor.fetchone()
        min_val = min_res[0] if min_res and min_res[0] is not None else 0

        return JSONResponse({
            "current": current_val,
            "current_time": current_time_str,
            "max": max_val,
            "min": min_val
        })

    finally:
        cursor.close()
        db.close()


@app.get("/get-last-5")
async def get_last_5(date: str = None):
    db = get_db_connection()
    cursor = db.cursor()

    try:
        where_clause = ""
        params = ()
        if date:
            where_clause = "WHERE DATE_FORMAT(time, '%d/%m/%Y') = %s"
            params = (date,)

        cursor.execute(f"""
            SELECT lux_value, time
            FROM lux
            {where_clause}
            ORDER BY id DESC
            LIMIT 5
        """, params)

        rows = cursor.fetchall()

        values = []
        for r in rows:
            values.append({
                "lux": r[0],
                "time": str(r[1])
            })

        values.reverse()

        return JSONResponse({"values": values})

    finally:
        cursor.close()
        db.close()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
