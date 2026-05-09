from app import app, db, User
from werkzeug.security import generate_password_hash

def create_admin():
    with app.app_context():
        # Check if admin already exists
        admin = User.query.filter_by(username='admin').first()
        if not admin:
            hashed_pw = generate_password_hash('123', method='pbkdf2:sha256')
            admin = User(username='admin', password=hashed_pw, is_admin=True)
            db.session.add(admin)
            db.session.commit()
            print("Admin account created: admin / 123")
        else:
            # Ensure existing admin has is_admin=True and reset password
            admin.is_admin = True
            admin.password = generate_password_hash('123', method='pbkdf2:sha256')
            db.session.commit()
            print("Admin account already exists. Password reset to: 123")

if __name__ == '__main__':
    create_admin()
