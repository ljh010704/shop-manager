import sqlite3, random
from datetime import date, timedelta

conn = sqlite3.connect(r'F:\报表\douyin-profit-tracker\data.db')

# Add shops
shops = [('女装旗舰店','女装'), ('女装专营店','女装'), ('男装旗舰店','男装'), ('童装店','童装')]
for name, group in shops:
    try: conn.execute('INSERT INTO shops(name,group_name) VALUES(?,?)', (name, group))
    except: pass

shop_ids = [r[0] for r in conn.execute('SELECT id FROM shops').fetchall()]
products = ['法式复古连衣裙','韩版宽松T恤','高腰阔腿裤','碎花半身裙','针织开衫','牛仔外套','运动套装','雪纺衬衫','蕾丝上衣','百褶裙']
statuses = ['待发货','已发货','已发货','已发货','已发货退款','退货退款']
warehouses = ['未到仓库','已到达仓库未发货','已到仓库已发货']
today = date.today()
cnt = 0

for day in range(15):
    d = today - timedelta(days=day)
    for _ in range(random.randint(5, 15)):
        dy_no = 'DY' + d.strftime('%Y%m%d') + str(random.randint(10000, 99999))
        da = round(random.uniform(49.9, 299.9), 2)
        ta = round(da * random.uniform(0.3, 0.6), 2)
        try:
            conn.execute(
                'INSERT INTO orders(douyin_order_no,shop_id,product_name,douyin_amount,taobao_order_no,taobao_amount,refund_status,order_date,is_influencer,warehouse_status,logistics_company,logistics_no) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',
                (dy_no, random.choice(shop_ids), random.choice(products), da,
                 'TB' + str(random.randint(100000, 999999)), ta,
                 random.choice(statuses), d.isoformat(),
                 random.choice(['否','否','否','是']),
                 random.choice(warehouses),
                 '顺丰速运', 'SF' + str(random.randint(1000000000, 9999999999))))
            cnt += 1
        except:
            pass

conn.commit()
conn.close()
print(f'Generated {cnt} orders')
