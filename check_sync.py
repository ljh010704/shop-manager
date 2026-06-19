import sqlite3
conn = sqlite3.connect('data.db')
conn.row_factory = sqlite3.Row

print("=== Shops Table ===")
shops = conn.execute('SELECT * FROM shops LIMIT 5').fetchall()
for s in shops:
    print(dict(s))

print("\n=== Daily Snapshot Count ===")
cnt = conn.execute('SELECT COUNT(*) FROM daily_snapshot').fetchone()[0]
print("Snapshots:", cnt)

print("\n=== Orders Count ===")
cnt = conn.execute('SELECT COUNT(*) FROM orders').fetchone()[0]
print("Orders:", cnt)

print("\n=== Dashboard Shops (with snapshots) ===")
dash_shops = conn.execute("""
    SELECT DISTINCT s.id, s.name, s.group_name
    FROM shops s
    JOIN daily_snapshot d ON s.id = d.shop_id
    LIMIT 5
""").fetchall()
for s in dash_shops:
    print(dict(s))

print("\n=== Profit Shops (with orders) ===")
profit_shops = conn.execute("""
    SELECT DISTINCT s.id, s.name, s.group_name
    FROM shops s
    JOIN orders o ON s.id = o.shop_id
    LIMIT 5
""").fetchall()
for s in profit_shops:
    print(dict(s))

conn.close()
