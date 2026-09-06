import os
import psycopg2
import bcrypt
from dotenv import load_dotenv

load_dotenv()

DB_HOST = os.getenv("DB_HOST", "aws-0-ap-south-1.pooler.supabase.com")
DB_NAME = os.getenv("DB_NAME", "postgres")
DB_USER = os.getenv("DB_USER", "postgres.ddvyyeexzamsgogjqnpy")
DB_PASS = os.getenv("DB_PASS", "Thalir@2026-")
DB_PORT = os.getenv("DB_PORT", "6543")

def get_connection():
    return psycopg2.connect(
        host=DB_HOST,
        database=DB_NAME,
        user=DB_USER,
        password=DB_PASS,
        port=DB_PORT,
        sslmode="require" if "supabase" in DB_HOST else "prefer"
    )

def hash_password(plain_text: str) -> str:
    return bcrypt.hashpw(plain_text.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

def verify_password(plain_text: str, hashed_text: str) -> bool:
    return bcrypt.checkpw(plain_text.encode("utf-8"), hashed_text.encode("utf-8"))

def init_production_db():
    conn = get_connection()
    conn.autocommit = True
    cur = conn.cursor()

    # 1. Users
    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username VARCHAR(50) UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role VARCHAR(20) NOT NULL CHECK (role IN ('Partner', 'Admin', 'Staff')),
        full_name VARCHAR(100),
        is_active BOOLEAN DEFAULT TRUE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 2. Partners
    cur.execute("""
    CREATE TABLE IF NOT EXISTS partners (
        id SERIAL PRIMARY KEY,
        name VARCHAR(100) NOT NULL,
        phone VARCHAR(20),
        profit_percentage NUMERIC(5, 2) NOT NULL CHECK (profit_percentage >= 0 AND profit_percentage <= 100),
        initial_investment NUMERIC(12, 2) DEFAULT 0.00,
        is_active BOOLEAN DEFAULT TRUE,
        joined_date DATE DEFAULT CURRENT_DATE
    );
    """)

    # 3. Vendors
    cur.execute("""
    CREATE TABLE IF NOT EXISTS vendors (
        id SERIAL PRIMARY KEY,
        name VARCHAR(120) NOT NULL,
        contact_person VARCHAR(100),
        phone VARCHAR(20),
        email VARCHAR(100),
        address TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 4. Products
    cur.execute("""
    CREATE TABLE IF NOT EXISTS products (
        id SERIAL PRIMARY KEY,
        barcode VARCHAR(100) UNIQUE NOT NULL,
        name VARCHAR(200) NOT NULL,
        vendor_id INTEGER REFERENCES vendors(id) ON DELETE SET NULL,
        wholesale_cost NUMERIC(10, 2) NOT NULL CHECK (wholesale_cost >= 0),
        selling_price NUMERIC(10, 2) NOT NULL CHECK (selling_price >= 0),
        stock_quantity INTEGER DEFAULT 0 CHECK (stock_quantity >= 0),
        min_threshold INTEGER DEFAULT 3,
        is_active BOOLEAN DEFAULT TRUE
    );
    """)

    # 5. Inventory History
    cur.execute("""
    CREATE TABLE IF NOT EXISTS inventory_history (
        id SERIAL PRIMARY KEY,
        product_id INTEGER REFERENCES products(id) ON DELETE CASCADE,
        transaction_type VARCHAR(20) NOT NULL CHECK (transaction_type IN ('STOCK_IN', 'STOCK_OUT', 'RETURN', 'ADJUSTMENT')),
        quantity INTEGER NOT NULL CHECK (quantity > 0),
        unit_price NUMERIC(10, 2) NOT NULL,
        cost_price NUMERIC(10, 2) NOT NULL,
        catalog_price NUMERIC(10, 2),
        discount_amount NUMERIC(10, 2) DEFAULT 0.00,
        customer_name VARCHAR(100),
        customer_phone VARCHAR(20),
        customer_place VARCHAR(100),
        biller_name VARCHAR(100),
        payment_mode VARCHAR(50) DEFAULT 'Cash',
        performed_by INTEGER REFERENCES users(id),
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 6. Expenses
    cur.execute("""
    CREATE TABLE IF NOT EXISTS expenses (
        id SERIAL PRIMARY KEY,
        category VARCHAR(50) NOT NULL CHECK (category IN ('Rent', 'Electricity', 'Vehicle / Fuel', 'Salary', 'Maintenance', 'Packaging', 'Miscellaneous')),
        amount NUMERIC(10, 2) NOT NULL CHECK (amount > 0),
        description TEXT,
        approved_by VARCHAR(100),
        logged_by INTEGER REFERENCES users(id),
        expense_date DATE DEFAULT CURRENT_DATE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 7. Audit Logs
    cur.execute("""
    CREATE TABLE IF NOT EXISTS backup_audit_logs (
        id SERIAL PRIMARY KEY,
        action_type VARCHAR(50) NOT NULL,
        performed_by INTEGER REFERENCES users(id),
        user_fullname VARCHAR(100),
        details TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # Safe Alters
    try:
        cur.execute("SET lock_timeout = '2s';")
        cur.execute("ALTER TABLE vendors ADD COLUMN IF NOT EXISTS address TEXT;")
        cur.execute("ALTER TABLE inventory_history ADD COLUMN IF NOT EXISTS customer_name VARCHAR(100);")
        cur.execute("ALTER TABLE inventory_history ADD COLUMN IF NOT EXISTS customer_phone VARCHAR(20);")
        cur.execute("ALTER TABLE inventory_history ADD COLUMN IF NOT EXISTS customer_place VARCHAR(100);")
        cur.execute("ALTER TABLE inventory_history ADD COLUMN IF NOT EXISTS biller_name VARCHAR(100);")
        cur.execute("ALTER TABLE inventory_history ADD COLUMN IF NOT EXISTS payment_mode VARCHAR(50) DEFAULT 'Cash';")
        cur.execute("ALTER TABLE inventory_history ADD COLUMN IF NOT EXISTS catalog_price NUMERIC(10, 2);")
        cur.execute("ALTER TABLE inventory_history ADD COLUMN IF NOT EXISTS discount_amount NUMERIC(10, 2) DEFAULT 0.00;")
        cur.execute("ALTER TABLE expenses ADD COLUMN IF NOT EXISTS approved_by VARCHAR(100);")
    except Exception:
        pass

    # Seed Admin
    cur.execute("SELECT COUNT(*) FROM users;")
    if cur.fetchone()[0] == 0:
        admin_pass = hash_password("ThalirAdmin@2026")
        cur.execute("""
            INSERT INTO users (username, password_hash, role, full_name)
            VALUES ('admin', %s, 'Partner', 'Renuka S');
        """, (admin_pass,))

    cur.close()
    conn.close()

if __name__ == "__main__":
    init_production_db()
    print("Database schema checked and ready!")