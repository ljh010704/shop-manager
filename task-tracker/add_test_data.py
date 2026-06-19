import sqlite3
import os
import sys
from datetime import datetime, date

sys.path.insert(0, r'F:\报表\task-tracker')
import crypto

DB_PATH = r'F:\报表\task-tracker\data.db'
conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row

# 添加营业执照
licenses = [
    ('张三营业执照', '91110108MA01XXXXX', '张三', '2028-12-31', '主执照'),
    ('李四营业执照', '91110108MA02YYYYY', '李四', '2027-06-30', '副执照'),
]
for lic in licenses:
    try:
        conn.execute('INSERT INTO licenses(license_name, license_no, holder_name, expire_date, remark) VALUES(?,?,?,?,?)', lic)
    except:
        pass
conn.commit()

# 添加店铺
shops = [
    ('抖音小店-A01', 1, '抖音A组', 'normal'),
    ('抖音小店-A02', 1, '抖音A组', 'normal'),
    ('抖音小店-A03', 1, '抖音A组', 'new'),
    ('快手小店-B01', 1, '快手组', 'normal'),
    ('快手小店-B02', 1, '快手组', 'normal'),
    ('拼多多-C01', 2, '拼多多组', 'new'),
    ('拼多多-C02', 2, '拼多多组', 'new'),
    ('淘宝店-D01', 2, '淘宝组', 'normal'),
    ('京东店-E01', None, '京东组', 'new'),
]
for shop in shops:
    try:
        conn.execute('INSERT INTO shops(shop_name, license_id, group_name, status) VALUES(?,?,?,?)', shop)
    except:
        pass
conn.commit()

# 添加平台账号
accounts = [
    (1, '抖音', 'dy_zhangsan_01', 'pwd123456', '主账号'),
    (1, '千川', 'qc_001', 'qcpwd123', '千川广告'),
    (2, '抖音', 'dy_zhangsan_02', 'pwd234567', ''),
    (3, '抖音', 'dy_zhangsan_03', 'pwd345678', '新店待开通'),
    (4, '快手', 'ks_lisi_01', 'kspwd123', ''),
    (4, '千川', 'qc_ks_001', 'qckspwd', ''),
    (5, '快手', 'ks_lisi_02', 'kspwd234', ''),
    (6, '拼多多', 'pdd_lisi_01', 'pddpwd123', '新店'),
    (7, '拼多多', 'pdd_lisi_02', 'pddpwd234', '新店'),
    (8, '淘宝', 'tb_wang_01', 'tbpwd123', ''),
    (9, '京东', 'jd_zhang_01', 'jdpwd123', '新店'),
]
for acc in accounts:
    try:
        enc_pwd = crypto.encrypt(acc[3])
        conn.execute('INSERT INTO platform_accounts(shop_id, platform_name, account, password_enc, remark) VALUES(?,?,?,?,?)',
                     (acc[0], acc[1], acc[2], enc_pwd, acc[4]))
    except Exception as e:
        print(f'Account error: {e}')
conn.commit()

# 为新店完成部分任务
new_shop_ids = [3, 6, 7, 9]
for sid in new_shop_ids:
    tasks = conn.execute('SELECT id FROM new_shop_task_templates WHERE is_active=1').fetchall()
    for i, task in enumerate(tasks[:3]):
        try:
            conn.execute('INSERT OR IGNORE INTO new_shop_tasks(shop_id, task_id, is_completed, completed_at) VALUES(?, ?, 1, ?)',
                        (sid, task['id'], datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        except:
            pass
conn.commit()

# 为普通店铺添加今日任务完成记录
today = date.today().isoformat()
normal_shop_ids = [1, 2, 4, 5, 8]
for sid in normal_shop_ids:
    tasks = conn.execute('SELECT id FROM daily_task_templates WHERE is_active=1').fetchall()
    for i, task in enumerate(tasks):
        try:
            is_done = 1 if i < 8 else 0
            completed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S') if is_done else ''
            conn.execute('INSERT OR REPLACE INTO daily_tasks(shop_id, task_id, task_date, is_completed, completed_at) VALUES(?,?,?,?,?)',
                        (sid, task['id'], today, is_done, completed_at))
        except:
            pass
conn.commit()

print('Data added successfully!')
print('Licenses:', conn.execute('SELECT COUNT(*) FROM licenses').fetchone()[0])
print('Shops:', conn.execute('SELECT COUNT(*) FROM shops').fetchone()[0])
print('Accounts:', conn.execute('SELECT COUNT(*) FROM platform_accounts').fetchone()[0])
conn.close()
