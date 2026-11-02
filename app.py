import os
import json
import datetime
import calendar
from flask import Flask, jsonify, render_template, request
import pyodbc
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
    return pyodbc.connect(DB_CONN_STR, timeout=10)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data')
def get_data():
    year = request.args.get('year', '2026')
    now = datetime.date.today()
    default_month = f'{now.month:02d}' if 5 <= now.month <= 10 else '08'
    month = request.args.get('month', default_month)
    
    is_all_season = (str(month).upper() == 'ALL')
    
    if is_all_season:
        year_month = f"{year}-ALL"
        start_date = f"{year}0501"
        end_date = f"{year}1031"
    else:
        try:
            month_int = int(month)
            month = f"{month_int:02d}"
        except ValueError:
            pass
            
        year_month = f"{year}-{month}"
        start_date = f"{year}{month}01"
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
            start_date = "20260801"
            end_date = "20260831"
            
    db_results = {}
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        query = """
            SELECT 
                CAST(dd.StayDate AS DATE) AS StayDate,
                COUNT(DISTINCT CASE WHEN a.AgencyCode != 'COMP' THEN dd.Room END) AS db_sold_room,
                SUM(CASE WHEN a.AgencyCode != 'COMP' THEN dd.Pax ELSE 0 END) AS db_adult,
                COUNT(DISTINCT CASE WHEN a.AgencyCode = 'NEILSON' AND (UPPER(REPLACE(r.Voucher, ' ', '')) = 'NOSHOW') THEN dd.Room END) AS db_unused_room,
                SUM(CASE WHEN a.AgencyCode = 'NEILSON' AND (UPPER(REPLACE(r.Voucher, ' ', '')) = 'NOSHOW') THEN dd.Pax ELSE 0 END) AS db_unused_pax,
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

    # Build dictionary of reference Excel data by date
    ref_map = {}
    if is_all_season:
        for m_idx in range(5, 11):
            for ref in excel_data.get(f"{year}-{m_idx:02d}", []):
                ref_map[ref['date']] = ref
    else:
        for ref in excel_data.get(year_month, []):
            ref_map[ref['date']] = ref

    merged_data = []
    num_days = (datetime.datetime.strptime(end_date, '%Y%m%d') - datetime.datetime.strptime(start_date, '%Y%m%d')).days + 1
    dates = [datetime.datetime.strptime(start_date, '%Y%m%d') + datetime.timedelta(days=i) for i in range(num_days)]
    
    for d in dates:
        d_str = d.strftime('%Y-%m-%d')
        ref = ref_map.get(d_str)
        db_item = db_results.get(d_str)
        
        cap = ref['capacity'] if ref else 111
        if db_item:
            sold = db_item['db_sold_room']
            adult = db_item['db_adult']
            unused = db_item['db_unused_room']
            unused_pax = db_item['db_unused_pax']
            reel_sold = max(0, sold - unused)
            reel_pax = max(0, adult - unused_pax)
            comp_room = db_item.get('db_comp_room', 0)
            comp_pax = db_item.get('db_comp_pax', 0)
        elif ref:
            sold = ref['sold_room']
            adult = ref['adult']
            unused = ref['unused_room']
            unused_pax = ref['unused_pax']
            reel_sold = ref['reel_sold_room']
            reel_pax = ref['reel_pax']
            comp_room = 0
            comp_pax = 0
        else:
            sold = 0
            adult = 0
            unused = 0
            unused_pax = 0
            reel_sold = 0
            reel_pax = 0
            comp_room = 0
            comp_pax = 0
            
        occ = (sold / cap) if cap > 0 else 0
        reel_occ = (reel_sold / cap) if cap > 0 else 0
        
        merged_data.append({
            'date': d_str,
            'capacity': cap,
            'sold_room': sold,
            'adult': adult,
            'unused_room': unused,
            'unused_pax': unused_pax,
            'reel_sold_room': reel_sold,
            'reel_pax': reel_pax,
            'comp_room': comp_room,
            'comp_pax': comp_pax,
            'occ_pct': occ,
            'reel_occ_pct': reel_occ,
            'ex_sold_room': ref['sold_room'] if ref else None,
            'ex_unused_room': ref['unused_room'] if ref else None,
            'ex_reel_sold_room': ref['reel_sold_room'] if ref else None,
            'is_live_only': bool(db_item and not ref)
        })
        
    return jsonify({
        'success': True,
        'year_month': year_month,
        'data': merged_data
    })

EXCEL_2025_SEASONAL = [
    {'month_num': 5, 'month_name': 'Mayıs', 'month_name_en': 'May', 'capacity': 3441, 'sold_room': 1907, 'difference': 1534, 'adult': 3426, 'total_pax': 3426, 'cin_room': 561, 'cout_room': 501, 'occ_pct': 0.5542, 'unused_room': 302, 'unused_pax': 604, 'reel_sold_room': 1605, 'reel_pax': 2822, 'reel_occ_pct': 0.4664},
    {'month_num': 6, 'month_name': 'Haziran', 'month_name_en': 'June', 'capacity': 3330, 'sold_room': 2265, 'difference': 1065, 'adult': 4113, 'total_pax': 4113, 'cin_room': 480, 'cout_room': 471, 'occ_pct': 0.6802, 'unused_room': 136, 'unused_pax': 272, 'reel_sold_room': 2129, 'reel_pax': 3841, 'reel_occ_pct': 0.6393},
    {'month_num': 7, 'month_name': 'Temmuz', 'month_name_en': 'July', 'capacity': 3438, 'sold_room': 2039, 'difference': 1399, 'adult': 3914, 'total_pax': 3914, 'cin_room': 407, 'cout_room': 411, 'occ_pct': 0.5931, 'unused_room': 91, 'unused_pax': 182, 'reel_sold_room': 1948, 'reel_pax': 3732, 'reel_occ_pct': 0.5666},
    {'month_num': 8, 'month_name': 'Ağustos', 'month_name_en': 'Agust', 'capacity': 3441, 'sold_room': 2359, 'difference': 1082, 'adult': 4459, 'total_pax': 4459, 'cin_room': 492, 'cout_room': 490, 'occ_pct': 0.6856, 'unused_room': 45, 'unused_pax': 90, 'reel_sold_room': 2314, 'reel_pax': 4369, 'reel_occ_pct': 0.6725},
    {'month_num': 9, 'month_name': 'Eylül', 'month_name_en': 'September', 'capacity': 3328, 'sold_room': 2204, 'difference': 1124, 'adult': 4028, 'total_pax': 4028, 'cin_room': 374, 'cout_room': 389, 'occ_pct': 0.6623, 'unused_room': 6, 'unused_pax': 12, 'reel_sold_room': 2198, 'reel_pax': 4016, 'reel_occ_pct': 0.6605},
    {'month_num': 10, 'month_name': 'Ekim', 'month_name_en': 'October', 'capacity': 3441, 'sold_room': 1677, 'difference': 1764, 'adult': 2944, 'total_pax': 2944, 'cin_room': 319, 'cout_room': 314, 'occ_pct': 0.4874, 'unused_room': 278, 'unused_pax': 556, 'reel_sold_room': 1399, 'reel_pax': 2388, 'reel_occ_pct': 0.4066}
]

@app.route('/api/seasonal_data')
def get_seasonal_data():
    year = request.args.get('year', '2026')
    if str(year) == '2025':
        return jsonify({'success': True, 'year': year, 'data': EXCEL_2025_SEASONAL})
        
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        q = """
            SELECT 
                MONTH(dd.StayDate) AS month_num,
                COUNT(DISTINCT CASE WHEN a.AgencyCode != 'COMP' THEN CAST(dd.StayDate AS VARCHAR) + '_' + CAST(dd.Room AS VARCHAR) END) AS sold_room,
                SUM(CASE WHEN a.AgencyCode != 'COMP' THEN dd.Pax ELSE 0 END) AS adult,
                COUNT(DISTINCT CASE WHEN a.AgencyCode = 'NEILSON' AND (UPPER(REPLACE(r.Voucher, ' ', '')) = 'NOSHOW') THEN CAST(dd.StayDate AS VARCHAR) + '_' + CAST(dd.Room AS VARCHAR) END) AS unused_room,
                SUM(CASE WHEN a.AgencyCode = 'NEILSON' AND (UPPER(REPLACE(r.Voucher, ' ', '')) = 'NOSHOW') THEN dd.Pax ELSE 0 END) AS unused_pax
            FROM DailyDetail dd
            JOIN Reservation r ON r.RecId = dd.ReservationId
            JOIN Agency a ON a.RecId = r.AgencyId
            WHERE YEAR(dd.StayDate) = ? AND MONTH(dd.StayDate) BETWEEN 5 AND 10
              AND dd.Status != -1 AND r.Status != -1
            GROUP BY MONTH(dd.StayDate)
            ORDER BY month_num
        """
        cursor.execute(q, (year,))
        rows = cursor.fetchall()
        
        q_cin_cout = """
            SELECT 
                MONTH(dd.StayDate) AS month_num,
                COUNT(DISTINCT CASE WHEN CAST(r.CheckinDate AS DATE) = CAST(dd.StayDate AS DATE) AND a.AgencyCode != 'COMP' THEN dd.Room END) AS cin_room,
                COUNT(DISTINCT CASE WHEN CAST(r.CheckOutDate AS DATE) = CAST(dd.StayDate AS DATE) AND a.AgencyCode != 'COMP' THEN dd.Room END) AS cout_room
            FROM DailyDetail dd
            JOIN Reservation r ON r.RecId = dd.ReservationId
            JOIN Agency a ON a.RecId = r.AgencyId
            WHERE YEAR(dd.StayDate) = ? AND MONTH(dd.StayDate) BETWEEN 5 AND 10
              AND dd.Status != -1 AND r.Status != -1
            GROUP BY MONTH(dd.StayDate)
        """
        cursor.execute(q_cin_cout, (year,))
        cc_rows = {r.month_num: {'cin_room': r.cin_room, 'cout_room': r.cout_room} for r in cursor.fetchall()}

        month_names_tr = {5: 'Mayıs', 6: 'Haziran', 7: 'Temmuz', 8: 'Ağustos', 9: 'Eylül', 10: 'Ekim'}
        month_names_en = {5: 'May', 6: 'June', 7: 'July', 8: 'Agust', 9: 'September', 10: 'October'}

        res = []
        for r in rows:
            m = r.month_num
            num_days = calendar.monthrange(int(year), m)[1]
            cap = num_days * 111
            sold = r.sold_room
            diff = max(0, cap - sold)
            adult = r.adult
            unused_r = r.unused_room
            unused_p = r.unused_pax
            reel_sold = max(0, sold - unused_r)
            reel_p = max(0, adult - unused_p)
            cc = cc_rows.get(m, {'cin_room': 0, 'cout_room': 0})
            
            res.append({
                'month_num': m,
                'month_name': month_names_tr[m],
                'month_name_en': month_names_en[m],
                'capacity': cap,
                'sold_room': sold,
                'difference': diff,
                'adult': adult,
                'total_pax': adult,
                'cin_room': cc['cin_room'],
                'cout_room': cc['cout_room'],
                'occ_pct': (sold / cap) if cap > 0 else 0,
                'unused_room': unused_r,
                'unused_pax': unused_p,
                'reel_sold_room': reel_sold,
                'reel_pax': reel_p,
                'reel_occ_pct': (reel_sold / cap) if cap > 0 else 0
            })
        conn.close()
        return jsonify({'success': True, 'year': year, 'data': res})
    except Exception as e:
        print("Seasonal query failed:", e)
        return jsonify({'success': True, 'year': year, 'data': EXCEL_2025_SEASONAL})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=PORT, debug=False)
