#!/usr/bin/env python3
"""
ICode 签到系统 — MySQL 初始化 + SQLite 数据迁移脚本
用法:
    python3 migrate_to_mysql.py
运行前请确认已安装: pip3 install pymysql
"""
import sqlite3, pymysql, json, sys, os

# ══ 配置（与 app.py 保持一致）══════════════════════════════════════
SQLITE_DB  = '/var/www/check-in/icode.db'
MYSQL_HOST = 'localhost'
MYSQL_PORT = 3306
MYSQL_USER = 'root'
MYSQL_PASS = 'NewPassword123!'
MYSQL_DB   = 'icode_checkin'
APP_USER   = 'icode'         # 给应用使用的低权限账号
APP_PASS   = 'Icode2025!'
# ════════════════════════════════════════════════════════════════════

def step(msg): print(f"\n{'='*50}\n▶ {msg}")
def ok(msg):   print(f"  ✅ {msg}")
def warn(msg): print(f"  ⚠  {msg}")

def setup_mysql():
    """建库、建用户、建表"""
    step("连接 MySQL (root)")
    root = pymysql.connect(host=MYSQL_HOST, port=MYSQL_PORT,
                           user=MYSQL_USER, password=MYSQL_PASS,
                           charset='utf8mb4',
                           cursorclass=pymysql.cursors.DictCursor)
    c = root.cursor()

    # 建库
    c.execute(f"CREATE DATABASE IF NOT EXISTS `{MYSQL_DB}` "
              f"DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
    ok(f"数据库 {MYSQL_DB} 已就绪")

    # 建应用账号（限本机登录）
    try:
        c.execute(f"CREATE USER IF NOT EXISTS '{APP_USER}'@'localhost' "
                  f"IDENTIFIED BY '{APP_PASS}'")
        c.execute(f"GRANT ALL PRIVILEGES ON `{MYSQL_DB}`.* "
                  f"TO '{APP_USER}'@'localhost'")
        c.execute("FLUSH PRIVILEGES")
        ok(f"应用账号 {APP_USER} 已创建并授权")
    except Exception as e:
        warn(f"创建账号跳过（可能已存在）: {e}")

    root.select_db(MYSQL_DB)

    # 建表
    tables = [
        """CREATE TABLE IF NOT EXISTS admins (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(100) NOT NULL,
            phone VARCHAR(50) UNIQUE NOT NULL,
            password VARCHAR(255) NOT NULL,
            is_main TINYINT DEFAULT 0,
            role_name VARCHAR(100) DEFAULT '',
            is_active TINYINT DEFAULT 1,
            permissions TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",

        """CREATE TABLE IF NOT EXISTS competitions (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(255) NOT NULL,
            location VARCHAR(255) DEFAULT '',
            start_time VARCHAR(50) DEFAULT '',
            end_time VARCHAR(50) DEFAULT '',
            description TEXT,
            description_images TEXT,
            album_url TEXT,
            manager_name VARCHAR(100) DEFAULT '',
            comp_admins TEXT,
            banner_text VARCHAR(255) DEFAULT '欢迎参加ICode比赛',
            banner_color VARCHAR(20) DEFAULT '#1a6fa8',
            banner_accent VARCHAR(20) DEFAULT '#0099cc',
            `groups` TEXT,
            display_fields TEXT,
            query_field VARCHAR(100) DEFAULT 'player_no,account',
            query_hint VARCHAR(255) DEFAULT '请输入报名编号或选手账号',
            is_active TINYINT DEFAULT 1,
            created_by INT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",

        """CREATE TABLE IF NOT EXISTS players (
            id INT AUTO_INCREMENT PRIMARY KEY,
            competition_id INT NOT NULL,
            player_no VARCHAR(100) DEFAULT '',
            account VARCHAR(100) DEFAULT '',
            name VARCHAR(100) NOT NULL,
            school VARCHAR(200) DEFAULT '',
            grade VARCHAR(50) DEFAULT '',
            group_name VARCHAR(100) DEFAULT '',
            comp_date VARCHAR(50) DEFAULT '',
            session VARCHAR(50) DEFAULT '',
            seat_no VARCHAR(50) DEFAULT '',
            shirt_size VARCHAR(20) DEFAULT '',
            checked_in TINYINT DEFAULT 0,
            checkin_time VARCHAR(50) DEFAULT '',
            remark TEXT,
            INDEX idx_comp (competition_id),
            INDEX idx_player_no (player_no),
            INDEX idx_account (account)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",

        """CREATE TABLE IF NOT EXISTS checkin_logs (
            id INT AUTO_INCREMENT PRIMARY KEY,
            player_id INT,
            competition_id INT,
            operator VARCHAR(100) DEFAULT '选手自助',
            checkin_time DATETIME DEFAULT CURRENT_TIMESTAMP,
            INDEX idx_comp (competition_id)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
    ]
    for sql in tables:
        c.execute(sql)
    root.commit()
    ok("所有表已创建")
    root.close()

def migrate_data():
    """从 SQLite 迁移数据到 MySQL"""
    if not os.path.exists(SQLITE_DB):
        warn(f"SQLite 文件不存在: {SQLITE_DB}，跳过迁移")
        return

    step("读取 SQLite 数据")
    lite = sqlite3.connect(SQLITE_DB)
    lite.row_factory = sqlite3.Row

    mysql = pymysql.connect(host=MYSQL_HOST, port=MYSQL_PORT,
                            user=MYSQL_USER, password=MYSQL_PASS,
                            db=MYSQL_DB, charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    mc = mysql.cursor()

    def migrate_table(table, columns, insert_sql, transform=None):
        rows = lite.execute(f"SELECT * FROM {table}").fetchall()
        if not rows:
            warn(f"{table}: 无数据"); return
        count = 0
        for row in rows:
            d = dict(row)
            vals = transform(d) if transform else tuple(d.get(c, '') or '' for c in columns)
            try:
                mc.execute(insert_sql, vals)
                count += 1
            except Exception as e:
                warn(f"{table} 行 {d.get('id','?')} 跳过: {e}")
        mysql.commit()
        ok(f"{table}: 迁移 {count}/{len(rows)} 行")

    # admins
    migrate_table('admins',
        ['id','name','phone','password','is_main','role_name','is_active','permissions','created_at'],
        """INSERT IGNORE INTO admins
           (id,name,phone,password,is_main,role_name,is_active,permissions,created_at)
           VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        lambda d: (
            d['id'], d['name'], d['phone'], d['password'],
            d.get('is_main',0), d.get('role_name','') or '',
            d.get('is_active',1),
            d.get('permissions','{}') or '{}',
            d.get('created_at') or None
        )
    )

    # competitions
    migrate_table('competitions',
        ['id','name','location','start_time','end_time','description',
         'description_images','album_url','manager_name','comp_admins',
         'banner_text','banner_color','banner_accent','groups','display_fields',
         'query_field','query_hint','is_active','created_by','created_at'],
        """INSERT IGNORE INTO competitions
           (id,name,location,start_time,end_time,description,
            description_images,album_url,manager_name,comp_admins,
            banner_text,banner_color,banner_accent,`groups`,display_fields,
            query_field,query_hint,is_active,created_by,created_at)
           VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        lambda d: (
            d['id'], d['name'], d.get('location',''), d.get('start_time',''),
            d.get('end_time',''), d.get('description','') or '',
            d.get('description_images','[]') or '[]',
            d.get('album_url','') or '', d.get('manager_name','') or '',
            d.get('comp_admins','[]') or '[]',
            d.get('banner_text','') or '', d.get('banner_color','') or '',
            d.get('banner_accent','') or '',
            d.get('groups','[]') or '[]',
            d.get('display_fields','[]') or '[]',
            d.get('query_field','player_no,account') or 'player_no,account',
            d.get('query_hint','') or '', d.get('is_active',1),
            d.get('created_by'), d.get('created_at') or None
        )
    )

    # players
    migrate_table('players',
        ['id','competition_id','player_no','account','name','school','grade',
         'group_name','comp_date','session','seat_no','shirt_size',
         'checked_in','checkin_time','remark'],
        """INSERT IGNORE INTO players
           (id,competition_id,player_no,account,name,school,grade,
            group_name,comp_date,session,seat_no,shirt_size,
            checked_in,checkin_time,remark)
           VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        lambda d: (
            d['id'], d['competition_id'],
            d.get('player_no','') or '', d.get('account','') or '',
            d['name'], d.get('school','') or '', d.get('grade','') or '',
            d.get('group_name','') or '', d.get('comp_date','') or '',
            d.get('session','') or '', d.get('seat_no','') or '',
            d.get('shirt_size','') or '', d.get('checked_in',0),
            d.get('checkin_time','') or '', d.get('remark','') or ''
        )
    )

    # checkin_logs
    migrate_table('checkin_logs',
        ['id','player_id','competition_id','operator','checkin_time'],
        """INSERT IGNORE INTO checkin_logs
           (id,player_id,competition_id,operator,checkin_time)
           VALUES(%s,%s,%s,%s,%s)""",
        lambda d: (
            d['id'], d.get('player_id'), d.get('competition_id'),
            d.get('operator','选手自助') or '选手自助',
            d.get('checkin_time') or None
        )
    )

    lite.close(); mysql.close()

def verify():
    """验证迁移结果"""
    step("验证数据")
    mysql = pymysql.connect(host=MYSQL_HOST, port=MYSQL_PORT,
                            user=MYSQL_USER, password=MYSQL_PASS,
                            db=MYSQL_DB, charset='utf8mb4',
                            cursorclass=pymysql.cursors.DictCursor)
    c = mysql.cursor()
    for table in ['admins','competitions','players','checkin_logs']:
        c.execute(f"SELECT COUNT(*) cnt FROM {table}")
        n = c.fetchone()['cnt']
        ok(f"{table}: {n} 条记录")
    mysql.close()

if __name__ == '__main__':
    print("\n🚀 ICode MySQL 迁移工具")
    print("=" * 50)
    try:
        setup_mysql()
        migrate_data()
        verify()
        print("\n" + "=" * 50)
        print("✅ 全部完成！")
        print("\n接下来:")
        print("  1. 把新版 app.py 替换到 /var/www/check-in/app.py")
        print("  2. pip3 install pymysql")
        print("  3. systemctl restart checkin")
        print("=" * 50)
    except Exception as e:
        print(f"\n❌ 出错: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)
