import sys
sys.stdout.reconfigure(encoding='utf-8')
import sqlite3
conn = sqlite3.connect('muong_thanh_sports_v2.db')
cursor = conn.cursor()
res = cursor.execute("SELECT sport_name, SUM(qty) FROM master_registrations GROUP BY sport_name").fetchall()
for r in res:
    print(r)
