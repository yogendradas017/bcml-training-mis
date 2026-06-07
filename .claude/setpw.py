import sqlite3
from werkzeug.security import generate_password_hash
con = sqlite3.connect('data/training.db')
con.execute('UPDATE users SET password=?, must_change_password=0 WHERE username=?', (generate_password_hash('TestPwd@2026Strong!'), 'balrampur'))
con.commit()
print('updated. MCP:', con.execute('SELECT must_change_password FROM users WHERE username=?', ('balrampur',)).fetchone())
con.close()
