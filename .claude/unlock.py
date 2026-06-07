import sqlite3
from werkzeug.security import generate_password_hash

db = sqlite3.connect('data/training.db')
# Unlock all users + reset failed attempts
db.execute("UPDATE users SET failed_attempts=0, locked_until=NULL")
# Reset balrampur password to default just in case
db.execute(
    "UPDATE users SET password=?, must_change_password=0 WHERE username='balrampur'",
    (generate_password_hash('bcml@1234'),)
)
db.commit()
rows = db.execute(
    "SELECT username, failed_attempts, locked_until, must_change_password FROM users WHERE username IN ('balrampur','admin','central')"
).fetchall()
for r in rows:
    print(r)
print('done')
