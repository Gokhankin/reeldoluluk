import os
import json
import datetime
from flask import Flask, jsonify, render_template, request
import pyodbc
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

DB_CONN_STR = os.getenv("DB_CONNECTION_STRING") or os.getenv("CONN_STR") or "DRIVER={ODBC Driver 18 for SQL Server};SERVER=192.168.0.41,1433;DATABASE=SednaAdakoy;UID=gokhan;PWD=Ad!!2025!!;TrustServerCertificate=yes;"
PORT = int(os.getenv("PORT", 5085))

# Load static Excel JSON data
EXCEL_DATA_PATH = os.path.join(os.path.dirname(__file__), 'excel_data.json')
excel_data = {}
if os.path.exists(EXCEL_DATA_PATH):
    try:
        with open(EXCEL_DATA_PATH, 'r', encoding='utf-8') as f:
            excel_data = json.load(f)
    except Exception as e:
        print("Error loading excel_data.json:", e)

def get_db_connection():
    return pyodbc.connect(DB_CONN_STR)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data')
def get_data():
    year = request.args.get('year', '2026')
    month = request.args.get('month', '06')
    
    # Format month to 2 digits
    try:
        month_int = int(month)
        month = f"{month_int:02d}"
    except ValueError:
        pass
        
    year_month = f"{year}-{month}"
    
    # 1. Fetch live SQL data for this month
    start_date = f"{year}{month}01"
    # End date calculation
    try:
        y = int(year)
        m = int(month)
        if m == 12:
            next_month = datetime.date(y + 1, 1, 1)
        else:
            next_month = datetime.date(y, m + 1, 1)
        end_date_obj = next_month - datetime.timedelta(days=1)
        end_date = end_date_obj.strftime('%Y%m%d')
    except Exception as e:
        print("Date parse error:", e)
        start_date = "20260601"
        end_date = "20260630"
        
    db_results = {}
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        query = """
            SELECT 
                CAST(dd.StayDate AS DATE) AS StayDate,
                COUNT(DISTINCT CASE WHEN a.AgencyCode != 'COMP' THEN dd.Room END) AS db_sold_room,
                SUM(CASE WHEN a.AgencyCode != 'COMP' THEN dd.Pax ELSE 0 END) AS db_adult,
                COUNT(DISTINCT CASE WHEN a.AgencyCode = 'NEILSON' AND r.Voucher = 'NOSHOW' THEN dd.Room END) AS db_unused_room,
                SUM(CASE WHEN a.AgencyCode = 'NEILSON' AND r.Voucher = 'NOSHOW' THEN dd.Pax ELSE 0 END) AS db_unused_pax,
                COUNT(DISTINCT CASE WHEN a.AgencyCode = 'COMP' THEN dd.Room END) AS db_comp_room,
                SUM(CASE WHEN a.AgencyCode = 'COMP' THEN dd.Pax ELSE 0 END) AS db_comp_pax
            FROM DailyDetail dd
            JOIN Reservation r ON r.RecId = dd.ReservationId
            JOIN Agency a ON a.RecId = r.AgencyId
            WHERE dd.StayDate >= ? AND dd.StayDate <= ?
              AND dd.Status != -1 AND r.Status != -1
            GROUP BY CAST(dd.StayDate AS DATE)
            ORDER BY StayDate
        """
        cursor.execute(query, (start_date, end_date))
        rows = cursor.fetchall()
        
        for r in rows:
            date_str = str(r.StayDate)
            db_results[date_str] = {
                'db_sold_room': r.db_sold_room,
                'db_adult': r.db_adult,
                'db_unused_room': r.db_unused_room,
                'db_unused_pax': r.db_unused_pax,
                'db_comp_room': r.db_comp_room,
                'db_comp_pax': r.db_comp_pax
            }
        conn.close()
    except Exception as e:
        print("Database query failed:", e)
        # We will continue and return Excel data only if database is offline
        
    # 2. Get Excel reference data
    ref_list = excel_data.get(year_month, [])
    
    # 3. Merge data
    merged_data = []
    # If Excel reference list is empty, construct days from database results or calendar
    if not ref_list:
        # Construct dynamic days
        try:
            num_days = (datetime.datetime.strptime(end_date, '%Y%m%d') - datetime.datetime.strptime(start_date, '%Y%m%d')).days + 1
            dates = [datetime.datetime.strptime(start_date, '%Y%m%d') + datetime.timedelta(days=i) for i in range(num_days)]
            for d in dates:
                d_str = d.strftime('%Y-%m-%d')
                db_item = db_results.get(d_str, {
                    'db_sold_room': 0, 'db_adult': 0, 'db_unused_room': 0, 'db_unused_pax': 0, 'db_comp_room': 0, 'db_comp_pax': 0
                })
                cap = 111 # default capacity
                sold = db_item['db_sold_room']
                unused = db_item['db_unused_room']
                reel_sold = max(0, sold - unused)
                reel_pax = max(0, db_item['db_adult'] - db_item['db_unused_pax'])
                
                merged_data.append({
                    'date': d_str,
                    'capacity': cap,
                    'sold_room': sold,
                    'adult': db_item['db_adult'],
                    'unused_room': unused,
                    'unused_pax': db_item['db_unused_pax'],
                    'reel_sold_room': reel_sold,
                    'reel_pax': reel_pax,
                    'occ_pct': (sold / cap) if cap > 0 else 0,
                    'reel_occ_pct': (reel_sold / cap) if cap > 0 else 0,
                    'is_live_only': True
                })
        except Exception as e:
            print("Error creating fallback days:", e)
    else:
        for ref in ref_list:
            d_str = ref['date']
            db_item = db_results.get(d_str, None)
            
            item = {
                'date': d_str,
                'capacity': ref['capacity'],
                
                # Excel references
                'ex_sold_room': ref['sold_room'],
                'ex_adult': ref['adult'],
                'ex_unused_room': ref['unused_room'],
                'ex_unused_pax': ref['unused_pax'],
                'ex_reel_sold_room': ref['reel_sold_room'],
                'ex_reel_pax': ref['reel_pax'],
                'ex_occ_pct': ref['occ_pct'],
                'ex_reel_occ_pct': ref['reel_occ_pct'],
                
                # Live SQL references (fallback to Excel if DB query failed)
                'sold_room': db_item['db_sold_room'] if db_item else ref['sold_room'],
                'adult': db_item['db_adult'] if db_item else ref['adult'],
                'unused_room': db_item['db_unused_room'] if db_item else ref['unused_room'],
                'unused_pax': db_item['db_unused_pax'] if db_item else ref['unused_pax'],
                'reel_sold_room': (db_item['db_sold_room'] - db_item['db_unused_room']) if db_item else ref['reel_sold_room'],
                'reel_pax': (db_item['db_adult'] - db_item['db_unused_pax']) if db_item else ref['reel_pax'],
                'comp_room': db_item['db_comp_room'] if db_item else 0,
                'comp_pax': db_item['db_comp_pax'] if db_item else 0,
                'is_live_only': False
            }
            # Recompute dynamic percentages
            cap = item['capacity']
            item['occ_pct'] = (item['sold_room'] / cap) if cap > 0 else 0
            item['reel_occ_pct'] = (item['reel_sold_room'] / cap) if cap > 0 else 0
            
            merged_data.append(item)
            
    return jsonify({
        'success': True,
        'year_month': year_month,
        'data': merged_data
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=PORT, debug=False)
