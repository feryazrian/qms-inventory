from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
from datetime import datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
import os
import subprocess
import tempfile
import psycopg2
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-only-change-me")

# DATA LOGIN
USERNAME = "admin"
PASSWORD = "112233"

# DATABASE CONFIG (LOCAL)
DB_CONFIG = {
    "host": "localhost",
    "database": "qms_inventory",
    "user": "postgres",
    "password": "12345678",
}

TARGET_MAP = {
    "1/8": Decimal("0.5"),
    "CGP": Decimal("0.166"),
    "CG1": Decimal("0.37"),
    "CG2": Decimal("0.60"),
}

LINE_MAP = {
    "CG1": 1,
    "CG2": 2,
    "1/8": None,
    "CGP": None,
}

CG_PDF_TEMPLATE_PATH = os.getenv(
    "CG_PDF_TEMPLATE_PATH",
    os.path.join(
        os.path.expanduser("~"),
        "Downloads",
        "SK-PROSBY-FM-01 ( Laporan Produksi Harian CG dan Pemakaian Bahan Penolong ).pdf",
    ),
)


def infer_per_roll_from_product(nama_produk):
    text = (nama_produk or "").strip().lower()
    if not text:
        return None
    if "cushion gum" in text:
        return Decimal("10")
    if "sidewall 1/8b" in text:
        return Decimal("7")
    if "sidewall 1/8mm" in text:
        return Decimal("10")
    if "cg potong" in text:
        return Decimal("25")
    return None


def get_db_conn():
    return get_connection()


def get_connection():
    database_url = os.getenv("DATABASE_URL")
    if database_url:
        return psycopg2.connect(database_url, sslmode="require")
    return psycopg2.connect(**DB_CONFIG)


def fetch_master_produk():
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, nama_produk
            FROM master_produk
            WHERE aktif = TRUE
            ORDER BY nama_produk
            """
        )
        return cur.fetchall()
    finally:
        if conn:
            conn.close()


def fetch_master_produk_all():
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, nama_produk, aktif
            FROM master_produk
            ORDER BY aktif DESC, nama_produk ASC
            """
        )
        return cur.fetchall()
    finally:
        if conn:
            conn.close()


def fetch_master_bahan_msc():
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT nama_bahan
            FROM master_bahan_msc
            ORDER BY id
            """
        )
        return [row[0] for row in cur.fetchall()]
    except Exception:
        # Jika tabel master belum ada, halaman MSC tetap bisa dibuka.
        return []
    finally:
        if conn:
            conn.close()


def fetch_laporan_cushion_gum():
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS nama_operator VARCHAR(150)
            """
        )
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS no_mesin VARCHAR(100)
            """
        )
        cur.execute(
            """
            SELECT
                id,
                tanggal_produksi,
                total_target,
                total_aktual,
                total_persentase,
                berat_kg_total,
                batch_uid
            FROM grand_total
            ORDER BY tanggal_produksi DESC, id DESC
            """
        )
        return cur.fetchall()
    except Exception:
        return []
    finally:
        if conn:
            conn.close()


def fetch_laporan_gum_cord():
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS nama_operator VARCHAR(150)
            """
        )
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS no_mesin VARCHAR(100)
            """
        )
        cur.execute(
            """
            SELECT
                ctid::text AS row_token,
                tanggal_produksi,
                target_total,
                aktual_kotak,
                persentase
            FROM production_gum_cord
            ORDER BY tanggal_produksi DESC
            """
        )
        return cur.fetchall()
    except Exception:
        return []
    finally:
        if conn:
            conn.close()

def fetch_laporan_msc():
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                id,
                tanggal_produksi,
                total_pakai_menit,
                total_target,
                total_aktual,
                total_persentase,
                batch_uid
            FROM grand_total_msc
            ORDER BY tanggal_produksi DESC, id DESC
            """
        )
        return cur.fetchall()
    except Exception:
        return []
    finally:
        if conn:
            conn.close()


def ensure_cushion_batch_columns(conn):
    cur = conn.cursor()
    cur.execute(
        """
        ALTER TABLE production_cushion_gum
        ADD COLUMN IF NOT EXISTS batch_uid VARCHAR(50)
        """
    )
    cur.execute(
        """
        ALTER TABLE production_cushion_gum
        ADD COLUMN IF NOT EXISTS order_roll INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE production_cushion_gum
        ADD COLUMN IF NOT EXISTS per_roll NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE production_cushion_gum
        ADD COLUMN IF NOT EXISTS berat_total NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total
        ADD COLUMN IF NOT EXISTS batch_uid VARCHAR(50)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total
        ADD COLUMN IF NOT EXISTS nama_operator VARCHAR(150)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total
        ADD COLUMN IF NOT EXISTS no_mesin VARCHAR(100)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total
        ADD COLUMN IF NOT EXISTS berat_kg_total NUMERIC(12, 2)
        """
    )


def ensure_pemakaian_plastik_table(conn):
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS pemakaian_plastik (
            id BIGSERIAL PRIMARY KEY,
            tanggal_produksi DATE,
            batch_uid VARCHAR(50),
            "230_blue" NUMERIC(12, 2),
            "210_green" NUMERIC(12, 2),
            "190_yellow" NUMERIC(12, 2),
            "630_birupolos" NUMERIC(12, 2),
            "630_red" NUMERIC(12, 2),
            "270_red" NUMERIC(12, 2),
            "240_red" NUMERIC(12, 2),
            plastik_gumcord NUMERIC(12, 2),
            total_plastik NUMERIC(12, 2),
            plastik_terbuang NUMERIC(12, 2),
            plastik_terbuang_cgpotong NUMERIC(12, 2),
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS tanggal_produksi DATE
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS batch_uid VARCHAR(50)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS "230_blue" NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS "210_green" NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS "190_yellow" NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS "630_birupolos" NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS "630_red" NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS "270_red" NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS "240_red" NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS plastik_gumcord NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS total_plastik NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS plastik_terbuang NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ADD COLUMN IF NOT EXISTS plastik_terbuang_cgpotong NUMERIC(12, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_plastik
        ALTER COLUMN "230_blue" TYPE NUMERIC(12, 2) USING "230_blue"::NUMERIC(12, 2),
        ALTER COLUMN "210_green" TYPE NUMERIC(12, 2) USING "210_green"::NUMERIC(12, 2),
        ALTER COLUMN "190_yellow" TYPE NUMERIC(12, 2) USING "190_yellow"::NUMERIC(12, 2),
        ALTER COLUMN "630_birupolos" TYPE NUMERIC(12, 2) USING "630_birupolos"::NUMERIC(12, 2),
        ALTER COLUMN "630_red" TYPE NUMERIC(12, 2) USING "630_red"::NUMERIC(12, 2),
        ALTER COLUMN "270_red" TYPE NUMERIC(12, 2) USING "270_red"::NUMERIC(12, 2),
        ALTER COLUMN "240_red" TYPE NUMERIC(12, 2) USING "240_red"::NUMERIC(12, 2),
        ALTER COLUMN plastik_gumcord TYPE NUMERIC(12, 2) USING plastik_gumcord::NUMERIC(12, 2),
        ALTER COLUMN total_plastik TYPE NUMERIC(12, 2) USING total_plastik::NUMERIC(12, 2),
        ALTER COLUMN plastik_terbuang TYPE NUMERIC(12, 2) USING plastik_terbuang::NUMERIC(12, 2),
        ALTER COLUMN plastik_terbuang_cgpotong TYPE NUMERIC(12, 2) USING plastik_terbuang_cgpotong::NUMERIC(12, 2)
        """
    )


def ensure_pemakaian_kotak_table(conn):
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS pemakaian_kotak (
            id BIGSERIAL PRIMARY KEY,
            tanggal_produksi DATE,
            batch_uid VARCHAR(50),
            box_160 INTEGER,
            box_185 INTEGER,
            box_200 INTEGER,
            box_220 INTEGER,
            box_310 INTEGER,
            box_350 INTEGER,
            box_gumcord INTEGER,
            total INTEGER,
            terbuang INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS tanggal_produksi DATE
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS batch_uid VARCHAR(50)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS box_160 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS box_185 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS box_200 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS box_220 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS box_310 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS box_350 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS box_gumcord INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS total INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_kotak
        ADD COLUMN IF NOT EXISTS terbuang INTEGER
        """
    )


def ensure_pemakaian_tungkul_table(conn):
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS pemakaian_tungkul (
            id BIGSERIAL PRIMARY KEY,
            tanggal_produksi DATE,
            batch_uid VARCHAR(50),
            tp_165 INTEGER,
            tp_195 INTEGER,
            tp_210 INTEGER,
            tp_240 INTEGER,
            total INTEGER,
            terbuang INTEGER,
            lakban INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS tanggal_produksi DATE
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS batch_uid VARCHAR(50)
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS tp_165 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS tp_195 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS tp_210 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS tp_240 INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS total INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS terbuang INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE pemakaian_tungkul
        ADD COLUMN IF NOT EXISTS lakban INTEGER
        """
    )


def generate_batch_uid(conn, tanggal_produksi, kode_utama, table_name):
    allowed_tables = {"grand_total", "grand_total_msc"}
    if table_name not in allowed_tables:
        raise ValueError("Table tidak diizinkan untuk batch_uid.")

    tanggal_obj = datetime.strptime(tanggal_produksi, "%Y-%m-%d").date()
    date_code = f"{tanggal_obj.day:02d}{tanggal_obj.month:02d}{tanggal_obj.year % 10}"
    prefix = f"{date_code}{kode_utama}"

    cur = conn.cursor()
    cur.execute(
        f"""
        SELECT batch_uid
        FROM {table_name}
        WHERE tanggal_produksi = %s
          AND batch_uid LIKE %s
        """,
        (tanggal_produksi, f"{prefix}%"),
    )
    rows = cur.fetchall()

    max_seq = 0
    for row in rows:
        value = row[0] or ""
        if not value.startswith(prefix):
            continue
        suffix = value[len(prefix) :]
        if suffix.isdigit():
            max_seq = max(max_seq, int(suffix))

    return f"{prefix}{max_seq + 1}"


def ensure_grand_total_msc_table(conn):
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS grand_total_msc (
            id BIGSERIAL PRIMARY KEY,
            tanggal_produksi DATE,
            nama_operator VARCHAR(150),
            no_mesin VARCHAR(100),
            regu VARCHAR(50),
            total_pakai_menit INTEGER,
            total_target NUMERIC(12, 3),
            total_aktual NUMERIC(12, 3),
            total_persentase NUMERIC(8, 2),
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS id BIGSERIAL
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS tanggal_produksi DATE
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS nama_operator VARCHAR(150)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS no_mesin VARCHAR(100)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS regu VARCHAR(50)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS total_pakai_menit INTEGER
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS total_target NUMERIC(12, 3)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS total_aktual NUMERIC(12, 3)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS total_persentase NUMERIC(8, 2)
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        """
    )
    cur.execute(
        """
        ALTER TABLE grand_total_msc
        ADD COLUMN IF NOT EXISTS batch_uid VARCHAR(50)
        """
    )
    cur.execute(
        """
        ALTER TABLE production_msc
        ADD COLUMN IF NOT EXISTS batch_uid VARCHAR(50)
        """
    )
    # Hapus constraint UNIQUE/PK pada tanggal_produksi agar histori per tanggal bisa lebih dari 1 baris.
    cur.execute(
        """
        DO $$
        DECLARE
            r RECORD;
        BEGIN
            FOR r IN
                SELECT c.conname
                FROM pg_constraint c
                JOIN pg_class t ON t.oid = c.conrelid
                JOIN pg_namespace n ON n.oid = t.relnamespace
                JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = ANY(c.conkey)
                WHERE n.nspname = 'public'
                  AND t.relname = 'grand_total_msc'
                  AND c.contype IN ('u', 'p')
                  AND a.attname = 'tanggal_produksi'
            LOOP
                EXECUTE format('ALTER TABLE public.grand_total_msc DROP CONSTRAINT %I', r.conname);
            END LOOP;
        END
        $$;
        """
    )
    # Hapus unique index manual pada tanggal_produksi jika ada.
    cur.execute(
        """
        DO $$
        DECLARE
            r RECORD;
        BEGIN
            FOR r IN
                SELECT i.relname AS index_name
                FROM pg_class t
                JOIN pg_index ix ON ix.indrelid = t.oid
                JOIN pg_class i ON i.oid = ix.indexrelid
                JOIN pg_namespace n ON n.oid = t.relnamespace
                JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = ANY(ix.indkey)
                WHERE n.nspname = 'public'
                  AND t.relname = 'grand_total_msc'
                  AND ix.indisunique = TRUE
                  AND a.attname = 'tanggal_produksi'
            LOOP
                EXECUTE format('DROP INDEX IF EXISTS public.%I', r.index_name);
            END LOOP;
        END
        $$;
        """
    )
    # Pastikan ada primary key di kolom id.
    cur.execute(
        """
        DO $$
        BEGIN
            IF NOT EXISTS (
                SELECT 1
                FROM pg_constraint c
                JOIN pg_class t ON t.oid = c.conrelid
                JOIN pg_namespace n ON n.oid = t.relnamespace
                WHERE n.nspname = 'public'
                  AND t.relname = 'grand_total_msc'
                  AND c.contype = 'p'
            ) THEN
                ALTER TABLE public.grand_total_msc ADD PRIMARY KEY (id);
            END IF;
        END
        $$;
        """
    )


def fetch_msc_batch(batch_uid):
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        ensure_grand_total_msc_table(conn)
        cur.execute(
            """
            SELECT
                batch_uid,
                tanggal_produksi,
                nama_operator,
                no_mesin,
                regu,
                total_pakai_menit,
                total_target,
                total_aktual,
                total_persentase
            FROM grand_total_msc
            WHERE batch_uid = %s
            ORDER BY id DESC
            LIMIT 1
            """,
            (batch_uid,),
        )
        header = cur.fetchone()
        if not header:
            return None

        cur.execute(
            """
            SELECT
                nama_bahan,
                jam_awal,
                jam_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_batch,
                persentase,
                obat_timbang,
                obat_sisa,
                keterangan
            FROM production_msc
            WHERE batch_uid = %s
            ORDER BY id
            """,
            (batch_uid,),
        )
        details = cur.fetchall()
        return {"header": header, "details": details}
    finally:
        if conn:
            conn.close()


@app.route("/laporan/msc/read/<batch_uid>", methods=["GET"])
def laporan_msc_read(batch_uid):
    if "user" not in session:
        return jsonify({"ok": False, "message": "Unauthorized"}), 401

    result = fetch_msc_batch(batch_uid)
    if not result:
        return jsonify({"ok": False, "message": "Data batch tidak ditemukan."}), 404

    header = result["header"]
    details = result["details"]
    return jsonify(
        {
            "ok": True,
            "data": {
                "batch_uid": header[0],
                "tanggal_produksi": str(header[1]) if header[1] else "",
                "nama_operator": header[2] or "",
                "no_mesin": header[3] or "",
                "regu": header[4] or "",
                "total_pakai_menit": header[5],
                "total_target": float(header[6]) if header[6] is not None else None,
                "total_aktual": float(header[7]) if header[7] is not None else None,
                "total_persentase": float(header[8]) if header[8] is not None else None,
                "rows": [
                    {
                        "nama_bahan": row[0] or "",
                        "jam_awal": row[1].strftime("%H:%M") if row[1] else "",
                        "jam_akhir": row[2].strftime("%H:%M") if row[2] else "",
                        "pakai_menit": row[3],
                        "target_per_menit": float(row[4]) if row[4] is not None else None,
                        "target_total": float(row[5]) if row[5] is not None else None,
                        "aktual_batch": float(row[6]) if row[6] is not None else None,
                        "persentase": float(row[7]) if row[7] is not None else None,
                        "obat_timbang": float(row[8]) if row[8] is not None else None,
                        "obat_sisa": float(row[9]) if row[9] is not None else None,
                        "keterangan": row[10] or "",
                    }
                    for row in details
                ],
            },
        }
    )


def fetch_cushion_batch(batch_uid):
    conn = None
    try:
        conn = get_db_conn()
        ensure_cushion_batch_columns(conn)
        ensure_pemakaian_plastik_table(conn)
        ensure_pemakaian_kotak_table(conn)
        ensure_pemakaian_tungkul_table(conn)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                batch_uid,
                tanggal_produksi,
                nama_operator,
                no_mesin,
                total_target,
                total_aktual,
                total_persentase,
                berat_kg_total
            FROM grand_total
            WHERE batch_uid = %s
            ORDER BY id DESC
            LIMIT 1
            """,
            (batch_uid,),
        )
        header = cur.fetchone()
        if not header:
            return None

        cur.execute(
            """
            SELECT
                nama_produk,
                order_roll,
                waktu_awal,
                waktu_akhir,
                line,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_roll,
                persentase,
                per_roll,
                berat_total
            FROM production_cushion_gum
            WHERE batch_uid = %s
            ORDER BY tanggal_produksi, waktu_awal, waktu_akhir, nama_produk
            """,
            (batch_uid,),
        )
        details = cur.fetchall()

        cur.execute(
            """
            SELECT
                "230_blue",
                "210_green",
                "190_yellow",
                "630_birupolos",
                "630_red",
                "270_red",
                "240_red",
                plastik_gumcord,
                total_plastik,
                plastik_terbuang,
                plastik_terbuang_cgpotong
            FROM pemakaian_plastik
            WHERE batch_uid = %s
            ORDER BY id DESC
            LIMIT 1
            """,
            (batch_uid,),
        )
        plastik = cur.fetchone()
        if not plastik and header[1]:
            cur.execute(
                """
                SELECT
                    "230_blue",
                    "210_green",
                    "190_yellow",
                    "630_birupolos",
                    "630_red",
                    "270_red",
                    "240_red",
                    plastik_gumcord,
                    total_plastik,
                    plastik_terbuang,
                    plastik_terbuang_cgpotong
                FROM pemakaian_plastik
                WHERE tanggal_produksi = %s
                ORDER BY id DESC
                LIMIT 1
                """,
                (header[1],),
            )
            plastik = cur.fetchone()
        if not plastik and header[1]:
            cur.execute(
                """
                SELECT
                    "230_blue",
                    "210_green",
                    "190_yellow",
                    "630_birupolos",
                    "630_red",
                    "270_red",
                    "240_red",
                    plastik_gumcord,
                    total_plastik,
                    plastik_terbuang,
                    plastik_terbuang_cgpotong
                FROM pemakaian_plastik
                WHERE created_at::date = %s
                ORDER BY id DESC
                LIMIT 1
                """,
                (header[1],),
            )
            plastik = cur.fetchone()

        cur.execute(
            """
            SELECT
                box_160,
                box_185,
                box_200,
                box_220,
                box_310,
                box_350,
                box_gumcord,
                total,
                terbuang
            FROM pemakaian_kotak
            WHERE batch_uid = %s
            ORDER BY id DESC
            LIMIT 1
            """,
            (batch_uid,),
        )
        kotak = cur.fetchone()
        if not kotak and header[1]:
            cur.execute(
                """
                SELECT
                    box_160,
                    box_185,
                    box_200,
                    box_220,
                    box_310,
                    box_350,
                    box_gumcord,
                    total,
                    terbuang
                FROM pemakaian_kotak
                WHERE tanggal_produksi = %s
                ORDER BY id DESC
                LIMIT 1
                """,
                (header[1],),
            )
            kotak = cur.fetchone()
        if not kotak and header[1]:
            cur.execute(
                """
                SELECT
                    box_160,
                    box_185,
                    box_200,
                    box_220,
                    box_310,
                    box_350,
                    box_gumcord,
                    total,
                    terbuang
                FROM pemakaian_kotak
                WHERE created_at::date = %s
                ORDER BY id DESC
                LIMIT 1
                """,
                (header[1],),
            )
            kotak = cur.fetchone()

        cur.execute(
            """
            SELECT
                tp_165,
                tp_195,
                tp_210,
                tp_240,
                total,
                terbuang,
                lakban
            FROM pemakaian_tungkul
            WHERE batch_uid = %s
            ORDER BY id DESC
            LIMIT 1
            """,
            (batch_uid,),
        )
        tungkul = cur.fetchone()
        if not tungkul and header[1]:
            cur.execute(
                """
                SELECT
                    tp_165,
                    tp_195,
                    tp_210,
                    tp_240,
                    total,
                    terbuang,
                    lakban
                FROM pemakaian_tungkul
                WHERE tanggal_produksi = %s
                ORDER BY id DESC
                LIMIT 1
                """,
                (header[1],),
            )
            tungkul = cur.fetchone()
        if not tungkul and header[1]:
            cur.execute(
                """
                SELECT
                    tp_165,
                    tp_195,
                    tp_210,
                    tp_240,
                    total,
                    terbuang,
                    lakban
                FROM pemakaian_tungkul
                WHERE created_at::date = %s
                ORDER BY id DESC
                LIMIT 1
                """,
                (header[1],),
            )
            tungkul = cur.fetchone()

        return {
            "header": header,
            "details": details,
            "plastik": plastik,
            "kotak": kotak,
            "tungkul": tungkul,
        }
    finally:
        if conn:
            conn.close()


@app.route("/laporan/cushion-gum/read/<batch_uid>", methods=["GET"])
def laporan_cushion_read(batch_uid):
    if "user" not in session:
        return jsonify({"ok": False, "message": "Unauthorized"}), 401

    result = fetch_cushion_batch(batch_uid)
    if not result:
        return jsonify({"ok": False, "message": "Data batch tidak ditemukan."}), 404

    header = result["header"]
    details = result["details"]
    plastik = result.get("plastik")
    kotak = result.get("kotak")
    tungkul = result.get("tungkul")
    return jsonify(
        {
            "ok": True,
            "data": {
                "batch_uid": header[0],
                "tanggal_produksi": str(header[1]) if header[1] else "",
                "nama_operator": header[2] or "",
                "no_mesin": header[3] or "",
                "total_target": float(header[4]) if header[4] is not None else None,
                "total_aktual": float(header[5]) if header[5] is not None else None,
                "total_persentase": float(header[6]) if header[6] is not None else None,
                "berat_kg_total": float(header[7]) if header[7] is not None else None,
                "rows": [
                    {
                        "nama_produk": row[0] or "",
                        "order_roll": row[1],
                        "waktu_awal": row[2].strftime("%H:%M") if row[2] else "",
                        "waktu_akhir": row[3].strftime("%H:%M") if row[3] else "",
                        "line": row[4],
                        "pakai_menit": row[5],
                        "target_per_menit": float(row[6]) if row[6] is not None else None,
                        "target_total": float(row[7]) if row[7] is not None else None,
                        "aktual_roll": row[8],
                        "persentase": float(row[9]) if row[9] is not None else None,
                        "per_roll": float(row[10]) if row[10] is not None else None,
                        "berat_total": float(row[11]) if row[11] is not None else None,
                    }
                    for row in details
                ],
                "plastik": {
                    "230_blue": float(plastik[0]) if plastik and plastik[0] is not None else None,
                    "210_green": float(plastik[1]) if plastik and plastik[1] is not None else None,
                    "190_yellow": float(plastik[2]) if plastik and plastik[2] is not None else None,
                    "630_birupolos": float(plastik[3]) if plastik and plastik[3] is not None else None,
                    "630_red": float(plastik[4]) if plastik and plastik[4] is not None else None,
                    "270_red": float(plastik[5]) if plastik and plastik[5] is not None else None,
                    "240_red": float(plastik[6]) if plastik and plastik[6] is not None else None,
                    "plastik_gumcord": float(plastik[7]) if plastik and plastik[7] is not None else None,
                    "total_plastik": float(plastik[8]) if plastik and plastik[8] is not None else None,
                    "plastik_terbuang": float(plastik[9]) if plastik and plastik[9] is not None else None,
                    "plastik_terbuang_cgpotong": float(plastik[10]) if plastik and plastik[10] is not None else None,
                },
                "kotak": {
                    "box_160": kotak[0] if kotak and kotak[0] is not None else None,
                    "box_185": kotak[1] if kotak and kotak[1] is not None else None,
                    "box_200": kotak[2] if kotak and kotak[2] is not None else None,
                    "box_220": kotak[3] if kotak and kotak[3] is not None else None,
                    "box_310": kotak[4] if kotak and kotak[4] is not None else None,
                    "box_350": kotak[5] if kotak and kotak[5] is not None else None,
                    "box_gumcord": kotak[6] if kotak and kotak[6] is not None else None,
                    "total": kotak[7] if kotak and kotak[7] is not None else None,
                    "terbuang": kotak[8] if kotak and kotak[8] is not None else None,
                },
                "tungkul": {
                    "tp_165": tungkul[0] if tungkul and tungkul[0] is not None else None,
                    "tp_195": tungkul[1] if tungkul and tungkul[1] is not None else None,
                    "tp_210": tungkul[2] if tungkul and tungkul[2] is not None else None,
                    "tp_240": tungkul[3] if tungkul and tungkul[3] is not None else None,
                    "total": tungkul[4] if tungkul and tungkul[4] is not None else None,
                    "terbuang": tungkul[5] if tungkul and tungkul[5] is not None else None,
                    "lakban": tungkul[6] if tungkul and tungkul[6] is not None else None,
                },
            },
        }
    )


@app.route("/laporan/cushion-gum/update/<batch_uid>", methods=["POST"])
def laporan_cushion_update(batch_uid):
    if "user" not in session:
        return jsonify({"ok": False, "message": "Unauthorized"}), 401

    payload = request.get_json(silent=True) or {}
    header = payload.get("header") or {}
    rows = payload.get("rows") or []
    plastik_payload = payload.get("plastik") or {}
    kotak_payload = payload.get("kotak") or {}
    tungkul_payload = payload.get("tungkul") or {}

    if not rows:
        return jsonify({"ok": False, "message": "Data detail Cushion Gum kosong."}), 400

    tanggal_produksi = (header.get("tanggal_produksi") or "").strip()
    if not tanggal_produksi:
        return jsonify({"ok": False, "message": "Tanggal produksi wajib diisi."}), 400

    nama_operator = (header.get("nama_operator") or "").strip() or None
    no_mesin = (header.get("no_mesin") or "").strip() or None
    berat_kg_total = parse_decimal(str(header.get("berat_kg_total") or ""))
    plastik_230_blue = parse_decimal(str(plastik_payload.get("230_blue") or ""))
    plastik_210_green = parse_decimal(str(plastik_payload.get("210_green") or ""))
    plastik_190_yellow = parse_decimal(str(plastik_payload.get("190_yellow") or ""))
    plastik_630_birupolos = parse_decimal(str(plastik_payload.get("630_birupolos") or ""))
    plastik_630_red = parse_decimal(str(plastik_payload.get("630_red") or ""))
    plastik_270_red = parse_decimal(str(plastik_payload.get("270_red") or ""))
    plastik_240_red = parse_decimal(str(plastik_payload.get("240_red") or ""))
    plastik_gumcord = parse_decimal(str(plastik_payload.get("plastik_gumcord") or ""))
    total_plastik = parse_decimal(str(plastik_payload.get("total_plastik") or ""))
    plastik_terbuang = parse_decimal(str(plastik_payload.get("plastik_terbuang") or ""))
    plastik_terbuang_cgpotong = parse_decimal(str(plastik_payload.get("plastik_terbuang_cgpotong") or ""))

    box_160 = parse_int(str(kotak_payload.get("box_160") or ""))
    box_185 = parse_int(str(kotak_payload.get("box_185") or ""))
    box_200 = parse_int(str(kotak_payload.get("box_200") or ""))
    box_220 = parse_int(str(kotak_payload.get("box_220") or ""))
    box_310 = parse_int(str(kotak_payload.get("box_310") or ""))
    box_350 = parse_int(str(kotak_payload.get("box_350") or ""))
    box_gumcord = parse_int(str(kotak_payload.get("box_gumcord") or ""))
    kotak_total = parse_int(str(kotak_payload.get("total") or ""))
    kotak_terbuang = parse_int(str(kotak_payload.get("terbuang") or ""))

    tp_165 = parse_int(str(tungkul_payload.get("tp_165") or ""))
    tp_195 = parse_int(str(tungkul_payload.get("tp_195") or ""))
    tp_210 = parse_int(str(tungkul_payload.get("tp_210") or ""))
    tp_240 = parse_int(str(tungkul_payload.get("tp_240") or ""))
    tungkul_total = parse_int(str(tungkul_payload.get("total") or ""))
    tungkul_terbuang = parse_int(str(tungkul_payload.get("terbuang") or ""))
    lakban = parse_int(str(tungkul_payload.get("lakban") or ""))

    plastik_items = [
        plastik_230_blue,
        plastik_210_green,
        plastik_190_yellow,
        plastik_630_birupolos,
        plastik_630_red,
        plastik_270_red,
        plastik_240_red,
        plastik_gumcord,
    ]
    if total_plastik is None and any(value is not None for value in plastik_items):
        total_plastik = sum((value or Decimal("0")) for value in plastik_items)

    kotak_items = [box_160, box_185, box_200, box_220, box_310, box_350, box_gumcord]
    if kotak_total is None and any(value is not None for value in kotak_items):
        kotak_total = sum(value or 0 for value in kotak_items)

    tungkul_items = [tp_165, tp_195, tp_210, tp_240]
    if tungkul_total is None and any(value is not None for value in tungkul_items):
        tungkul_total = sum(value or 0 for value in tungkul_items)

    parsed_rows = []
    total_target = Decimal("0")
    total_aktual = 0
    computed_berat_kg_total = Decimal("0")

    for row in rows:
        row = row or {}
        nama_produk = (row.get("nama_produk") or "").strip() or None
        order_roll = parse_int(str(row.get("order_roll") or ""))
        waktu_awal = (row.get("waktu_awal") or "").strip() or None
        waktu_akhir = (row.get("waktu_akhir") or "").strip() or None
        line_val = parse_int(str(row.get("line") or ""))
        pakai_menit = parse_int(str(row.get("pakai_menit") or ""))
        target_per_menit = parse_decimal(str(row.get("target_per_menit") or ""))
        target_total = parse_decimal(str(row.get("target_total") or ""))
        aktual_roll = parse_int(str(row.get("aktual_roll") or ""))
        persentase = parse_decimal(str(row.get("persentase") or ""))
        per_roll = parse_decimal(str(row.get("per_roll") or ""))
        berat_total = parse_decimal(str(row.get("berat_total") or ""))

        if not any(
            [
                nama_produk,
                order_roll is not None,
                waktu_awal,
                waktu_akhir,
                line_val is not None,
                pakai_menit is not None,
                target_per_menit is not None,
                target_total is not None,
                aktual_roll is not None,
                persentase is not None,
            ]
        ):
            continue

        if pakai_menit is None and waktu_awal and waktu_akhir:
            try:
                t_awal = datetime.strptime(waktu_awal, "%H:%M")
                t_akhir = datetime.strptime(waktu_akhir, "%H:%M")
                diff = int((t_akhir - t_awal).total_seconds() / 60)
                if diff >= 0:
                    pakai_menit = diff
            except ValueError:
                pakai_menit = None

        if target_total is None and pakai_menit is not None and target_per_menit is not None:
            target_total = (Decimal(pakai_menit) * target_per_menit).quantize(Decimal("0.01"))

        if persentase is None and target_total and target_total != 0 and aktual_roll is not None:
            persentase = ((Decimal(aktual_roll) / target_total) * Decimal("100")).quantize(
                Decimal("0.01")
            )
        if per_roll is None:
            per_roll = infer_per_roll_from_product(nama_produk)
        if berat_total is None and aktual_roll is not None and per_roll is not None:
            berat_total = (Decimal(aktual_roll) * per_roll).quantize(Decimal("0.01"))

        total_target += target_total or Decimal("0")
        total_aktual += aktual_roll or 0
        computed_berat_kg_total += berat_total or Decimal("0")

        parsed_rows.append(
            (
                tanggal_produksi,
                nama_produk,
                order_roll,
                waktu_awal,
                waktu_akhir,
                line_val,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_roll,
                persentase,
                per_roll,
                berat_total,
                batch_uid,
            )
        )

    if not parsed_rows:
        return jsonify({"ok": False, "message": "Tidak ada baris valid untuk disimpan."}), 400

    total_persentase = None
    if total_target != 0:
        total_persentase = ((Decimal(total_aktual) / total_target) * Decimal("100")).quantize(
            Decimal("0.01")
        )
    if berat_kg_total is None:
        berat_kg_total = computed_berat_kg_total or None

    conn = None
    try:
        conn = get_db_conn()
        ensure_cushion_batch_columns(conn)
        ensure_pemakaian_plastik_table(conn)
        ensure_pemakaian_kotak_table(conn)
        ensure_pemakaian_tungkul_table(conn)
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE grand_total
            SET
                tanggal_produksi = %s,
                nama_operator = %s,
                no_mesin = %s,
                total_target = %s,
                total_aktual = %s,
                total_persentase = %s,
                berat_kg_total = %s
            WHERE batch_uid = %s
            """,
            (
                tanggal_produksi,
                nama_operator,
                no_mesin,
                total_target,
                total_aktual,
                total_persentase,
                berat_kg_total,
                batch_uid,
            ),
        )
        cur.execute("DELETE FROM production_cushion_gum WHERE batch_uid = %s", (batch_uid,))
        cur.executemany(
            """
            INSERT INTO production_cushion_gum (
                tanggal_produksi,
                nama_produk,
                order_roll,
                waktu_awal,
                waktu_akhir,
                line,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_roll,
                persentase,
                per_roll,
                berat_total,
                batch_uid
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
            """,
            parsed_rows,
        )

        cur.execute("DELETE FROM pemakaian_plastik WHERE batch_uid = %s", (batch_uid,))
        plastik_values = [
            plastik_230_blue,
            plastik_210_green,
            plastik_190_yellow,
            plastik_630_birupolos,
            plastik_630_red,
            plastik_270_red,
            plastik_240_red,
            plastik_gumcord,
            total_plastik,
            plastik_terbuang,
            plastik_terbuang_cgpotong,
        ]
        if any(value is not None for value in plastik_values):
            cur.execute(
                """
                INSERT INTO pemakaian_plastik (
                    tanggal_produksi,
                    batch_uid,
                    "230_blue",
                    "210_green",
                    "190_yellow",
                    "630_birupolos",
                    "630_red",
                    "270_red",
                    "240_red",
                    plastik_gumcord,
                    total_plastik,
                    plastik_terbuang,
                    plastik_terbuang_cgpotong
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
                """,
                (tanggal_produksi, batch_uid, *plastik_values),
            )

        cur.execute("DELETE FROM pemakaian_kotak WHERE batch_uid = %s", (batch_uid,))
        kotak_values = [
            box_160,
            box_185,
            box_200,
            box_220,
            box_310,
            box_350,
            box_gumcord,
            kotak_total,
            kotak_terbuang,
        ]
        if any(value is not None for value in kotak_values):
            cur.execute(
                """
                INSERT INTO pemakaian_kotak (
                    tanggal_produksi,
                    batch_uid,
                    box_160,
                    box_185,
                    box_200,
                    box_220,
                    box_310,
                    box_350,
                    box_gumcord,
                    total,
                    terbuang
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
                """,
                (tanggal_produksi, batch_uid, *kotak_values),
            )

        cur.execute("DELETE FROM pemakaian_tungkul WHERE batch_uid = %s", (batch_uid,))
        tungkul_values = [
            tp_165,
            tp_195,
            tp_210,
            tp_240,
            tungkul_total,
            tungkul_terbuang,
            lakban,
        ]
        if any(value is not None for value in tungkul_values):
            cur.execute(
                """
                INSERT INTO pemakaian_tungkul (
                    tanggal_produksi,
                    batch_uid,
                    tp_165,
                    tp_195,
                    tp_210,
                    tp_240,
                    total,
                    terbuang,
                    lakban
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
                """,
                (tanggal_produksi, batch_uid, *tungkul_values),
            )
        conn.commit()
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Gagal update: {e}"}), 500
    finally:
        if conn:
            conn.close()

    return jsonify({"ok": True, "message": "Data Cushion Gum berhasil diupdate."})


@app.route("/laporan/gum-cord/read/<row_token>", methods=["GET"])
def laporan_gum_cord_read(row_token):
    if "user" not in session:
        return jsonify({"ok": False, "message": "Unauthorized"}), 401

    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS nama_operator VARCHAR(150)
            """
        )
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS no_mesin VARCHAR(100)
            """
        )
        cur.execute(
            """
            SELECT
                ctid::text AS row_token,
                tanggal_produksi,
                nama_operator,
                no_mesin,
                nama_produk,
                order_kotak,
                waktu_awal,
                waktu_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_kotak,
                persentase,
                berat_per_kotak,
                berat_total
            FROM production_gum_cord
            WHERE ctid = %s::tid
            LIMIT 1
            """,
            (row_token,),
        )
        row = cur.fetchone()
        if not row:
            return jsonify({"ok": False, "message": "Data Gum Cord tidak ditemukan."}), 404

        return jsonify(
            {
                "ok": True,
                "data": {
                    "row_token": row[0],
                    "tanggal_produksi": str(row[1]) if row[1] else "",
                    "nama_operator": row[2] or "",
                    "no_mesin": row[3] or "",
                    "nama_produk": row[4] or "",
                    "order_kotak": row[5],
                    "waktu_awal": row[6].strftime("%H:%M") if row[6] else "",
                    "waktu_akhir": row[7].strftime("%H:%M") if row[7] else "",
                    "pakai_menit": row[8],
                    "target_per_menit": float(row[9]) if row[9] is not None else None,
                    "target_total": float(row[10]) if row[10] is not None else None,
                    "aktual_kotak": row[11],
                    "persentase": float(row[12]) if row[12] is not None else None,
                    "berat_per_kotak": float(row[13]) if row[13] is not None else None,
                    "berat_total": float(row[14]) if row[14] is not None else None,
                },
            }
        )
    finally:
        if conn:
            conn.close()


@app.route("/laporan/gum-cord/update/<row_token>", methods=["POST"])
def laporan_gum_cord_update(row_token):
    if "user" not in session:
        return jsonify({"ok": False, "message": "Unauthorized"}), 401

    payload = request.get_json(silent=True) or {}
    header = payload.get("header") or {}
    row = payload.get("row") or {}

    tanggal_produksi = (header.get("tanggal_produksi") or "").strip()
    if not tanggal_produksi:
        return jsonify({"ok": False, "message": "Tanggal produksi wajib diisi."}), 400

    nama_operator = (header.get("nama_operator") or "").strip() or None
    no_mesin = (header.get("no_mesin") or "").strip() or None
    nama_produk = (row.get("nama_produk") or "").strip() or "Gum Cord"
    order_kotak = parse_int(str(row.get("order_kotak") or ""))
    waktu_awal = (row.get("waktu_awal") or "").strip() or None
    waktu_akhir = (row.get("waktu_akhir") or "").strip() or None
    pakai_menit = parse_int(str(row.get("pakai_menit") or ""))
    target_per_menit = parse_decimal(str(row.get("target_per_menit") or ""))
    target_total = parse_decimal(str(row.get("target_total") or ""))
    aktual_kotak = parse_int(str(row.get("aktual_kotak") or ""))
    persentase = parse_decimal(str(row.get("persentase") or ""))
    berat_per_kotak = parse_decimal(str(row.get("berat_per_kotak") or ""))
    berat_total = parse_decimal(str(row.get("berat_total") or ""))

    if pakai_menit is None and waktu_awal and waktu_akhir:
        try:
            t_awal = datetime.strptime(waktu_awal, "%H:%M")
            t_akhir = datetime.strptime(waktu_akhir, "%H:%M")
            diff = int((t_akhir - t_awal).total_seconds() / 60)
            if diff >= 0:
                pakai_menit = diff
        except ValueError:
            pakai_menit = None

    if target_total is None and pakai_menit is not None and target_per_menit is not None:
        target_total = (Decimal(pakai_menit) * target_per_menit).quantize(Decimal("0.01"))

    if persentase is None and target_total and target_total != 0 and aktual_kotak is not None:
        persentase = ((Decimal(aktual_kotak) / target_total) * Decimal("100")).quantize(
            Decimal("0.01")
        )

    if berat_total is None and aktual_kotak is not None and berat_per_kotak is not None:
        berat_total = (Decimal(aktual_kotak) * berat_per_kotak).quantize(Decimal("0.01"))

    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS nama_operator VARCHAR(150)
            """
        )
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS no_mesin VARCHAR(100)
            """
        )
        cur.execute(
            """
            UPDATE production_gum_cord
            SET
                tanggal_produksi = %s,
                nama_operator = %s,
                no_mesin = %s,
                nama_produk = %s,
                order_kotak = %s,
                waktu_awal = %s,
                waktu_akhir = %s,
                pakai_menit = %s,
                target_per_menit = %s,
                target_total = %s,
                aktual_kotak = %s,
                persentase = %s,
                berat_per_kotak = %s,
                berat_total = %s
            WHERE ctid = %s::tid
            """,
            (
                tanggal_produksi,
                nama_operator,
                no_mesin,
                nama_produk,
                order_kotak,
                waktu_awal,
                waktu_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_kotak,
                persentase,
                berat_per_kotak,
                berat_total,
                row_token,
            ),
        )
        if cur.rowcount == 0:
            conn.rollback()
            return jsonify({"ok": False, "message": "Data Gum Cord tidak ditemukan."}), 404
        conn.commit()
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Gagal update: {e}"}), 500
    finally:
        if conn:
            conn.close()

    return jsonify({"ok": True, "message": "Data Gum Cord berhasil diupdate."})


@app.route("/laporan/msc/update/<batch_uid>", methods=["POST"])
def laporan_msc_update(batch_uid):
    if "user" not in session:
        return jsonify({"ok": False, "message": "Unauthorized"}), 401

    payload = request.get_json(silent=True) or {}
    header = payload.get("header") or {}
    rows = payload.get("rows") or []

    if not rows:
        return jsonify({"ok": False, "message": "Data detail MSC kosong."}), 400

    tanggal_produksi = (header.get("tanggal_produksi") or "").strip()
    if not tanggal_produksi:
        return jsonify({"ok": False, "message": "Tanggal produksi wajib diisi."}), 400

    nama_operator = (header.get("nama_operator") or "").strip() or None
    no_mesin = (header.get("no_mesin") or "").strip() or None
    regu = (header.get("regu") or "").strip() or None

    parsed_rows = []
    total_pakai_menit = 0
    total_target = Decimal("0")
    total_aktual = Decimal("0")

    for row in rows:
        nama_bahan = ((row or {}).get("nama_bahan") or "").strip() or None
        jam_awal = ((row or {}).get("jam_awal") or "").strip() or None
        jam_akhir = ((row or {}).get("jam_akhir") or "").strip() or None
        pakai_menit = parse_int(str((row or {}).get("pakai_menit") or ""))
        target_per_menit = parse_decimal(str((row or {}).get("target_per_menit") or ""))
        target_total = parse_decimal(str((row or {}).get("target_total") or ""))
        aktual_batch = parse_decimal(str((row or {}).get("aktual_batch") or ""))
        persentase = parse_decimal(str((row or {}).get("persentase") or ""))
        obat_timbang = parse_decimal(str((row or {}).get("obat_timbang") or ""))
        obat_sisa = parse_decimal(str((row or {}).get("obat_sisa") or ""))
        keterangan = ((row or {}).get("keterangan") or "").strip() or None

        if not any(
            [
                nama_bahan,
                jam_awal,
                jam_akhir,
                pakai_menit is not None,
                target_per_menit is not None,
                target_total is not None,
                aktual_batch is not None,
                persentase is not None,
                obat_timbang is not None,
                obat_sisa is not None,
                keterangan,
            ]
        ):
            continue

        total_pakai_menit += pakai_menit or 0
        total_target += target_total or Decimal("0")
        total_aktual += aktual_batch or Decimal("0")

        parsed_rows.append(
            (
                tanggal_produksi,
                nama_bahan,
                jam_awal,
                jam_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_batch,
                persentase,
                obat_timbang,
                obat_sisa,
                keterangan,
                batch_uid,
            )
        )

    if not parsed_rows:
        return jsonify({"ok": False, "message": "Tidak ada baris valid untuk disimpan."}), 400

    total_persentase = None
    if total_target != 0:
        total_persentase = ((total_aktual / total_target) * Decimal("100")).quantize(
            Decimal("0.01")
        )

    conn = None
    try:
        conn = get_db_conn()
        ensure_grand_total_msc_table(conn)
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE grand_total_msc
            SET
                tanggal_produksi = %s,
                nama_operator = %s,
                no_mesin = %s,
                regu = %s,
                total_pakai_menit = %s,
                total_target = %s,
                total_aktual = %s,
                total_persentase = %s
            WHERE batch_uid = %s
            """,
            (
                tanggal_produksi,
                nama_operator,
                no_mesin,
                regu,
                total_pakai_menit,
                total_target,
                total_aktual,
                total_persentase,
                batch_uid,
            ),
        )
        cur.execute("DELETE FROM production_msc WHERE batch_uid = %s", (batch_uid,))
        cur.executemany(
            """
            INSERT INTO production_msc (
                tanggal_produksi,
                nama_bahan,
                jam_awal,
                jam_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_batch,
                persentase,
                obat_timbang,
                obat_sisa,
                keterangan,
                batch_uid
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
            """,
            parsed_rows,
        )
        conn.commit()
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Gagal update: {e}"}), 500
    finally:
        if conn:
            conn.close()

    return jsonify({"ok": True, "message": "Data MSC berhasil diupdate."})


def parse_int(value):
    if value is None:
        return None
    value = value.strip()
    if value == "":
        return None
    try:
        return int(value)
    except ValueError:
        return None


def parse_decimal(value):
    if value is None:
        return None
    value = value.strip()
    if value == "":
        return None
    try:
        return Decimal(value)
    except (InvalidOperation, ValueError):
        return None


def format_number_display(value):
    if value is None:
        return ""
    try:
        if isinstance(value, Decimal):
            text = format(value, "f")
        else:
            text = format(Decimal(str(value)), "f")
    except (InvalidOperation, ValueError, TypeError):
        text = str(value)
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def format_product_name_for_print(nama_produk):
    text = (nama_produk or "").strip()
    if not text:
        return ""

    # Khusus produk Cushion Gum: paksa baris 1 "CUSHION GUM", detail ukuran di baris 2.
    lowered = text.lower()
    if lowered.startswith("cushion gum "):
        return f"CUSHION GUM\n{text[12:].strip()}"
    if lowered == "cushion gum":
        return "CUSHION GUM"

    return text


def build_cushion_pdf_from_template(batch_data):
    try:
        from pypdf import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase import pdfmetrics
    except Exception as exc:
        raise RuntimeError(
            "Library PDF belum terpasang. Jalankan: pip install reportlab pypdf"
        ) from exc

    if not os.path.exists(CG_PDF_TEMPLATE_PATH):
        raise FileNotFoundError(
            f"Template PDF tidak ditemukan di path: {CG_PDF_TEMPLATE_PATH}"
        )

    header = batch_data.get("header") or ()
    rows = batch_data.get("details") or []

    template_reader = PdfReader(CG_PDF_TEMPLATE_PATH)
    if not template_reader.pages:
        raise RuntimeError("Template PDF kosong.")

    template_page = template_reader.pages[0]
    page_width = float(template_page.mediabox.width)
    page_height = float(template_page.mediabox.height)

    max_rows_per_page = 21
    row_chunks = [
        rows[i : i + max_rows_per_page] for i in range(0, len(rows), max_rows_per_page)
    ] or [[]]

    writer = PdfWriter()

    def fit_text(text, font_name, font_size, max_width):
        value = "" if text is None else str(text)
        if max_width <= 0:
            return ""
        if pdfmetrics.stringWidth(value, font_name, font_size) <= max_width:
            return value
        ellipsis = "..."
        while value:
            candidate = value[:-1].rstrip()
            trial = f"{candidate}{ellipsis}"
            if pdfmetrics.stringWidth(trial, font_name, font_size) <= max_width:
                return trial
            value = candidate
        return ellipsis

    for page_index, chunk in enumerate(row_chunks):
        overlay_buffer = BytesIO()
        c = canvas.Canvas(overlay_buffer, pagesize=(page_width, page_height))

        c.setFont("Helvetica", 9)
        tanggal = header[1].strftime("%d-%m-%Y") if len(header) > 1 and header[1] else ""
        c.drawString(95, page_height - 105, (header[2] or "") if len(header) > 2 else "")
        c.drawString(360, page_height - 105, (header[3] or "") if len(header) > 3 else "")
        c.drawString(505, page_height - 105, tanggal)

        # Fine-tuned against SK-PROSBY-FM-01 template grid so text sits inside row cells.
        y = page_height - 188
        row_height = 28
        font_name = "Helvetica"
        font_size = 8
        pad_x = 4
        x_cols = {
            "nama_produk": (18, 107, "left"),
            "order_roll": (107, 148, "right"),
            "waktu_awal": (148, 198, "center"),
            "waktu_akhir": (198, 249, "center"),
            "line": (249, 292, "center"),
            "pakai_menit": (292, 336, "right"),
            "target_per_menit": (336, 380, "right"),
            "target_total": (380, 424, "right"),
            "aktual_roll": (424, 468, "right"),
            "persentase": (468, 512, "right"),
            "per_roll": (512, 556, "right"),
            "berat_total": (556, 600, "right"),
        }

        c.setFont(font_name, font_size)

        def draw_cell(col_key, y_pos, text):
            x0, x1, align = x_cols[col_key]
            max_w = max((x1 - x0) - (pad_x * 2), 1)
            clean_text = fit_text(text, font_name, font_size, max_w)
            if align == "left":
                c.drawString(x0 + pad_x, y_pos, clean_text)
            elif align == "center":
                c.drawCentredString((x0 + x1) / 2, y_pos, clean_text)
            else:
                c.drawRightString(x1 - pad_x, y_pos, clean_text)

        for row in chunk:
            nama_produk = (row[0] or "") if len(row) > 0 else ""
            order_roll = row[1] if len(row) > 1 else None
            waktu_awal = row[2].strftime("%H:%M") if len(row) > 2 and row[2] else ""
            waktu_akhir = row[3].strftime("%H:%M") if len(row) > 3 and row[3] else ""
            line = row[4] if len(row) > 4 else None
            pakai_menit = row[5] if len(row) > 5 else None
            target_per_menit = row[6] if len(row) > 6 else None
            target_total = row[7] if len(row) > 7 else None
            aktual_roll = row[8] if len(row) > 8 else None
            persentase = row[9] if len(row) > 9 else None
            per_roll = row[10] if len(row) > 10 else None
            berat_total = row[11] if len(row) > 11 else None

            draw_cell("nama_produk", y, nama_produk)
            draw_cell("order_roll", y, format_number_display(order_roll))
            draw_cell("waktu_awal", y, waktu_awal)
            draw_cell("waktu_akhir", y, waktu_akhir)
            draw_cell("line", y, format_number_display(line))
            draw_cell("pakai_menit", y, format_number_display(pakai_menit))
            draw_cell("target_per_menit", y, format_number_display(target_per_menit))
            draw_cell("target_total", y, format_number_display(target_total))
            draw_cell("aktual_roll", y, format_number_display(aktual_roll))
            draw_cell("persentase", y, format_number_display(persentase))
            draw_cell("per_roll", y, format_number_display(per_roll))
            draw_cell("berat_total", y, format_number_display(berat_total))
            y -= row_height

        is_last_page = page_index == (len(row_chunks) - 1)
        if is_last_page:
            total_target = header[4] if len(header) > 4 else None
            total_aktual = header[5] if len(header) > 5 else None
            total_persen = header[6] if len(header) > 6 else None
            total_berat = header[7] if len(header) > 7 else None

            total_y = 90
            c.setFont("Helvetica-Bold", 8)
            c.drawRightString(
                x_cols["target_total"][1] - pad_x, total_y, format_number_display(total_target)
            )
            c.drawRightString(
                x_cols["aktual_roll"][1] - pad_x, total_y, format_number_display(total_aktual)
            )
            c.drawRightString(
                x_cols["persentase"][1] - pad_x, total_y, format_number_display(total_persen)
            )
            c.drawRightString(
                x_cols["berat_total"][1] - pad_x, total_y, format_number_display(total_berat)
            )

        c.save()
        overlay_buffer.seek(0)

        overlay_pdf = PdfReader(overlay_buffer)
        page_template_reader = PdfReader(CG_PDF_TEMPLATE_PATH)
        page = page_template_reader.pages[0]
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output


def build_download_pdf(title, metadata_rows, sections, page_size):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except Exception as exc:
        raise RuntimeError(
            "Library PDF belum terpasang. Jalankan: pip install reportlab pypdf"
        ) from exc

    buffer = BytesIO()
    resolved_page_size = landscape(A4) if page_size == "landscape" else A4
    doc = SimpleDocTemplate(
        buffer,
        pagesize=resolved_page_size,
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DownloadTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=18,
        spaceAfter=8,
    )
    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#b91c1c"),
        spaceAfter=6,
        spaceBefore=8,
    )
    body_style = ParagraphStyle(
        "BodyTextCompact",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
    )

    story = [Paragraph(title, title_style)]

    if metadata_rows:
        metadata_table = Table(
            [[Paragraph(f"<b>{label}</b>", body_style), Paragraph(value or "-", body_style)] for label, value in metadata_rows],
            colWidths=[45 * mm, doc.width - (45 * mm)],
        )
        metadata_table.setStyle(
            TableStyle(
                [
                    ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#111827")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d1d5db")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f3f4f6")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        story.extend([metadata_table, Spacer(1, 8)])

    for section in sections:
        story.append(Paragraph(section["title"], heading_style))
        table = Table(section["rows"], repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#111827")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#fee2e2")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), section.get("font_size", 7)),
                    ("LEADING", (0, 0), (-1, -1), section.get("leading", 9)),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.extend([table, Spacer(1, 8)])

    doc.build(story)
    buffer.seek(0)
    return buffer


def build_combined_laporan_download_pdf(cushion_result, gum_cord_row, batch_uid_label):
    header = (cushion_result or {}).get("header") or ()
    details = (cushion_result or {}).get("details") or []
    plastik = (cushion_result or {}).get("plastik")
    kotak = (cushion_result or {}).get("kotak")
    tungkul = (cushion_result or {}).get("tungkul")

    tanggal_obj = header[1] if len(header) > 1 and header[1] else (gum_cord_row or {}).get("tanggal_produksi")
    tanggal = tanggal_obj.strftime("%d-%m-%Y") if tanggal_obj else "-"
    nama_operator = (header[2] if len(header) > 2 and header[2] else (gum_cord_row or {}).get("nama_operator")) or "-"
    no_mesin = (header[3] if len(header) > 3 and header[3] else (gum_cord_row or {}).get("no_mesin")) or "-"

    cushion_rows = [[
        "Nama Produk", "Order", "Awal", "Akhir", "Line", "Pakai",
        "Target/mnt", "Target", "Aktual", "%", "Per Roll", "Berat",
    ]]
    for row in details:
        cushion_rows.append(
            [
                format_product_name_for_print((row[0] or "") if len(row) > 0 else "").replace("\n", " "),
                format_number_display(row[1] if len(row) > 1 else None),
                row[2].strftime("%H:%M") if len(row) > 2 and row[2] else "-",
                row[3].strftime("%H:%M") if len(row) > 3 and row[3] else "-",
                format_number_display(row[4] if len(row) > 4 else None),
                format_number_display(row[5] if len(row) > 5 else None),
                format_number_display(row[6] if len(row) > 6 else None),
                format_number_display(row[7] if len(row) > 7 else None),
                format_number_display(row[8] if len(row) > 8 else None),
                format_number_display(row[9] if len(row) > 9 else None),
                format_number_display(row[10] if len(row) > 10 else None),
                format_number_display(row[11] if len(row) > 11 else None),
            ]
        )
    if len(cushion_rows) == 1:
        cushion_rows.append(["-"] * 12)

    helper_rows = [
        ["Komponen", "Nilai"],
        ["Plastik Total", format_number_display(plastik[8] if plastik else None)],
        ["Plastik Terbuang", format_number_display(plastik[9] if plastik else None)],
        ["Plastik Rework CG Potong", format_number_display(plastik[10] if plastik else None)],
        ["Kotak Total", format_number_display(kotak[7] if kotak else None)],
        ["Kotak Terbuang", format_number_display(kotak[8] if kotak else None)],
        ["Tungkul Total", format_number_display(tungkul[4] if tungkul else None)],
        ["Tungkul Terbuang", format_number_display(tungkul[5] if tungkul else None)],
        ["Lakban", format_number_display(tungkul[6] if tungkul else None)],
    ]

    gum_cord_rows = [[
        "Nama Produk", "Order", "Awal", "Akhir", "Pakai", "Target/mnt",
        "Target", "Aktual", "%", "Berat/Kotak", "Berat Total",
    ]]
    gum_cord_rows.append(
        [
            (gum_cord_row or {}).get("nama_produk") or "Gum Cord",
            format_number_display((gum_cord_row or {}).get("order_kotak")),
            ((gum_cord_row or {}).get("waktu_awal").strftime("%H:%M") if (gum_cord_row or {}).get("waktu_awal") else "-"),
            ((gum_cord_row or {}).get("waktu_akhir").strftime("%H:%M") if (gum_cord_row or {}).get("waktu_akhir") else "-"),
            format_number_display((gum_cord_row or {}).get("pakai_menit")),
            format_number_display((gum_cord_row or {}).get("target_per_menit")),
            format_number_display((gum_cord_row or {}).get("target_total")),
            format_number_display((gum_cord_row or {}).get("aktual_kotak")),
            format_number_display((gum_cord_row or {}).get("persentase")),
            format_number_display((gum_cord_row or {}).get("berat_per_kotak")),
            format_number_display((gum_cord_row or {}).get("berat_total")),
        ]
    )

    metadata_rows = [
        ("Batch UID", batch_uid_label or "-"),
        ("Tanggal", tanggal),
        ("Nama Operator", nama_operator),
        ("No. Mesin", no_mesin),
        ("Total Target", format_number_display(header[4] if len(header) > 4 else None)),
        ("Total Aktual", format_number_display(header[5] if len(header) > 5 else None)),
        ("Total Persentase", format_number_display(header[6] if len(header) > 6 else None)),
        ("Total Berat", format_number_display(header[7] if len(header) > 7 else None)),
    ]

    return build_download_pdf(
        "Laporan Produksi Harian CG dan Pemakaian Bahan Penolong",
        metadata_rows,
        [
            {"title": "A. Hasil Produksi Cushion Gum", "rows": cushion_rows, "font_size": 6.5},
            {"title": "B. Ringkasan Bahan Penolong", "rows": helper_rows, "font_size": 8},
            {"title": "C. Hasil Produksi Gum Cord", "rows": gum_cord_rows, "font_size": 7},
        ],
        page_size="landscape",
    )


def build_msc_download_pdf(result, batch_uid):
    header = result.get("header") or ()
    details = result.get("details") or []

    tanggal_obj = header[1] if len(header) > 1 and header[1] else None
    tanggal = tanggal_obj.strftime("%d-%m-%Y") if tanggal_obj else "-"

    metadata_rows = [
        ("Batch UID", batch_uid or "-"),
        ("Tanggal", tanggal),
        ("Nama Operator", (header[2] or "-") if len(header) > 2 else "-"),
        ("No. Mesin", (header[3] or "-") if len(header) > 3 else "-"),
        ("Regu", (header[4] or "-") if len(header) > 4 else "-"),
        ("Total Pakai Menit", format_number_display(header[5] if len(header) > 5 else None)),
        ("Total Target", format_number_display(header[6] if len(header) > 6 else None)),
        ("Total Aktual", format_number_display(header[7] if len(header) > 7 else None)),
        ("Total Persentase", format_number_display(header[8] if len(header) > 8 else None)),
    ]

    msc_rows = [[
        "Nama Bahan", "Jam Awal", "Jam Akhir", "Pakai", "Target/mnt", "Target",
        "Aktual", "%", "Obat Timbang", "Obat Sisa", "Keterangan",
    ]]
    for row in details:
        msc_rows.append(
            [
                row[0] or "-",
                row[1].strftime("%H:%M") if len(row) > 1 and row[1] else "-",
                row[2].strftime("%H:%M") if len(row) > 2 and row[2] else "-",
                format_number_display(row[3] if len(row) > 3 else None),
                format_number_display(row[4] if len(row) > 4 else None),
                format_number_display(row[5] if len(row) > 5 else None),
                format_number_display(row[6] if len(row) > 6 else None),
                format_number_display(row[7] if len(row) > 7 else None),
                format_number_display(row[8] if len(row) > 8 else None),
                format_number_display(row[9] if len(row) > 9 else None),
                row[10] or "-",
            ]
        )
    if len(msc_rows) == 1:
        msc_rows.append(["-"] * 11)

    return build_download_pdf(
        "Laporan Produksi Harian MSC",
        metadata_rows,
        [{"title": "Hasil Produksi MSC", "rows": msc_rows, "font_size": 7}],
        page_size="landscape",
    )


def load_static_css(filename):
    css_path = os.path.join(app.root_path, "static", "css", filename)
    with open(css_path, "r", encoding="utf-8") as css_file:
        return css_file.read()


def find_pdf_browser_executable():
    candidates = [
        os.environ.get("CHROME_PATH"),
        os.environ.get("EDGE_PATH"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None


def render_template_to_pdf_bytes(template_name, context, css_filename):
    browser_path = find_pdf_browser_executable()
    if not browser_path:
        raise RuntimeError(
            "Browser untuk generate PDF tidak ditemukan. Pastikan Chrome atau Edge terpasang di lokasi default."
        )

    html_context = dict(context)
    html_context["inline_css"] = load_static_css(css_filename)
    html_context["show_toolbar"] = False
    html = render_template(template_name, **html_context)

    html_file = None
    pdf_file = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as html_handle:
            html_handle.write(html)
            html_file = html_handle.name

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as pdf_handle:
            pdf_file = pdf_handle.name

        subprocess.run(
            [
                browser_path,
                "--headless=new",
                "--disable-gpu",
                "--allow-file-access-from-files",
                "--no-pdf-header-footer",
                f"--print-to-pdf={pdf_file}",
                html_file,
            ],
            check=True,
            timeout=90,
        )

        with open(pdf_file, "rb") as result_file:
            pdf_bytes = result_file.read()
        return BytesIO(pdf_bytes)
    except subprocess.CalledProcessError as exc:
        raise RuntimeError("Gagal membuat PDF dari template print.") from exc
    finally:
        for temp_path in [html_file, pdf_file]:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass


@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        # AMANKAN INPUT
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        # LOGIN BENAR
        if username == USERNAME and password == PASSWORD:
            session["user"] = username
            return redirect(url_for("home"))

        # KEDUANYA SALAH
        if username != USERNAME and password != PASSWORD:
            error = "Username atau Password Salah"

        # USERNAME SALAH
        elif username != USERNAME:
            error = "Username Tidak Terdaftar"

        # PASSWORD SALAH
        else:
            error = "Password Salah"

        return render_template("login.html", error=error)

    return render_template("login.html")


@app.route("/home")
def home():
    if "user" not in session:
        return redirect(url_for("login"))

    return render_template("home.html", user=session["user"])


@app.route("/gum-cord")
def gum_cord():
    if "user" not in session:
        return redirect(url_for("login"))

    return render_template("gum_cord.html", user=session["user"])


@app.route("/item-code", methods=["GET", "POST"])
def item_code():
    if "user" not in session:
        return redirect(url_for("login"))

    notice = ""
    error = ""
    is_ajax = request.headers.get("X-Requested-With") == "XMLHttpRequest"

    if request.method == "POST":
        action = (request.form.get("action") or "add").strip().lower()
        conn = None
        try:
            conn = get_db_conn()
            cur = conn.cursor()

            if action == "delete":
                id_produk = parse_int(request.form.get("id_produk"))
                if id_produk is None:
                    error = "ID produk tidak valid."
                else:
                    cur.execute("DELETE FROM master_produk WHERE id = %s", (id_produk,))
                    notice = "Produk berhasil dihapus."

            elif action == "edit":
                id_asal = parse_int(request.form.get("id_asal"))
                id_baru = parse_int(request.form.get("id_produk"))
                nama_produk = (request.form.get("nama_produk") or "").strip()
                aktif = (request.form.get("aktif") or "").strip() == "1"

                if id_asal is None or id_baru is None:
                    error = "ID produk tidak valid."
                elif not nama_produk:
                    error = "Nama produk wajib diisi."
                else:
                    cur.execute(
                        """
                        SELECT id
                        FROM master_produk
                        WHERE LOWER(TRIM(nama_produk)) = LOWER(TRIM(%s))
                          AND id <> %s
                        LIMIT 1
                        """,
                        (nama_produk, id_asal),
                    )
                    duplicate_name = cur.fetchone()
                    if duplicate_name:
                        error = "Nama produk sudah dipakai produk lain."
                    else:
                        cur.execute(
                            """
                            UPDATE master_produk
                            SET id = %s, nama_produk = %s, aktif = %s
                            WHERE id = %s
                            """,
                            (id_baru, nama_produk, aktif, id_asal),
                        )
                        notice = "Produk berhasil diperbarui."
                        if is_ajax:
                            response_payload = {
                                "ok": True,
                                "notice": notice,
                                "row": {
                                    "id_asal": id_asal,
                                    "id_produk": id_baru,
                                    "nama_produk": nama_produk,
                                    "aktif": aktif,
                                    "aktif_label": "Aktif" if aktif else "Nonaktif",
                                },
                            }

            else:
                nama_produk = (request.form.get("nama_produk") or "").strip()
                aktif = (request.form.get("aktif") or "").strip() == "1"
                if not nama_produk:
                    error = "Nama produk wajib diisi."
                else:
                    cur.execute(
                        """
                        SELECT id, aktif
                        FROM master_produk
                        WHERE LOWER(TRIM(nama_produk)) = LOWER(TRIM(%s))
                        LIMIT 1
                        """,
                        (nama_produk,),
                    )
                    existing = cur.fetchone()
                    if existing:
                        cur.execute(
                            """
                            UPDATE master_produk
                            SET nama_produk = %s, aktif = %s
                            WHERE id = %s
                            """,
                            (nama_produk, aktif, existing[0]),
                        )
                        notice = "Produk sudah ada. Data diperbarui."
                    else:
                        cur.execute(
                            """
                            INSERT INTO master_produk (nama_produk, aktif)
                            VALUES (%s, %s)
                            """,
                            (nama_produk, aktif),
                        )
                        notice = "Produk baru berhasil ditambahkan."

            if not error:
                conn.commit()
                if is_ajax and action == "edit":
                    return jsonify(response_payload)
                return redirect(url_for("item_code", notice=notice))
            conn.rollback()
            if is_ajax and action == "edit":
                return jsonify({"ok": False, "message": error}), 400
        except Exception as e:
            if conn:
                conn.rollback()
            error = f"Gagal menyimpan produk: {e}"
            if is_ajax and action == "edit":
                return jsonify({"ok": False, "message": error}), 500
        finally:
            if conn:
                conn.close()

    notice_from_query = (request.args.get("notice") or "").strip()
    if notice_from_query:
        notice = notice_from_query

    products = fetch_master_produk_all()
    return render_template(
        "item_code.html",
        products=products,
        notice=notice,
        error=error,
        user=session["user"],
    )


@app.route("/laporan")
def laporan():
    if "user" not in session:
        return redirect(url_for("login"))

    active_tab = request.args.get("tab", "cushion-gum")
    if active_tab not in ["cushion-gum", "gum-cord", "msc"]:
        active_tab = "cushion-gum"

    laporan_cushion_gum = fetch_laporan_cushion_gum()
    laporan_gum_cord = fetch_laporan_gum_cord()
    laporan_msc = fetch_laporan_msc()
    return render_template(
        "laporan.html",
        user=session["user"],
        active_tab=active_tab,
        laporan_cushion_gum=laporan_cushion_gum,
        laporan_gum_cord=laporan_gum_cord,
        laporan_msc=laporan_msc,
    )


def fetch_latest_gum_cord_by_date(tanggal_produksi):
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                ctid::text AS row_token,
                tanggal_produksi,
                nama_operator,
                no_mesin,
                nama_produk,
                order_kotak,
                waktu_awal,
                waktu_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_kotak,
                persentase,
                berat_per_kotak,
                berat_total
            FROM production_gum_cord
            WHERE tanggal_produksi = %s
            ORDER BY ctid DESC
            LIMIT 1
            """,
            (tanggal_produksi,),
        )
        row = cur.fetchone()
        if not row:
            return None
        return {
            "row_token": row[0],
            "tanggal_produksi": row[1],
            "nama_operator": row[2] if len(row) > 2 else None,
            "no_mesin": row[3] if len(row) > 3 else None,
            "nama_produk": row[4] if len(row) > 4 else None,
            "order_kotak": row[5] if len(row) > 5 else None,
            "waktu_awal": row[6] if len(row) > 6 else None,
            "waktu_akhir": row[7] if len(row) > 7 else None,
            "pakai_menit": row[8] if len(row) > 8 else None,
            "target_per_menit": row[9] if len(row) > 9 else None,
            "target_total": row[10] if len(row) > 10 else None,
            "aktual_kotak": row[11] if len(row) > 11 else None,
            "persentase": row[12] if len(row) > 12 else None,
            "berat_per_kotak": row[13] if len(row) > 13 else None,
            "berat_total": row[14] if len(row) > 14 else None,
        }
    finally:
        if conn:
            conn.close()


def fetch_gum_cord_by_row_token(row_token):
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                ctid::text AS row_token,
                tanggal_produksi,
                nama_operator,
                no_mesin,
                nama_produk,
                order_kotak,
                waktu_awal,
                waktu_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_kotak,
                persentase,
                berat_per_kotak,
                berat_total
            FROM production_gum_cord
            WHERE ctid = %s::tid
            LIMIT 1
            """,
            (row_token,),
        )
        row = cur.fetchone()
        if not row:
            return None
        return {
            "row_token": row[0],
            "tanggal_produksi": row[1],
            "nama_operator": row[2] if len(row) > 2 else None,
            "no_mesin": row[3] if len(row) > 3 else None,
            "nama_produk": row[4] if len(row) > 4 else None,
            "order_kotak": row[5] if len(row) > 5 else None,
            "waktu_awal": row[6] if len(row) > 6 else None,
            "waktu_akhir": row[7] if len(row) > 7 else None,
            "pakai_menit": row[8] if len(row) > 8 else None,
            "target_per_menit": row[9] if len(row) > 9 else None,
            "target_total": row[10] if len(row) > 10 else None,
            "aktual_kotak": row[11] if len(row) > 11 else None,
            "persentase": row[12] if len(row) > 12 else None,
            "berat_per_kotak": row[13] if len(row) > 13 else None,
            "berat_total": row[14] if len(row) > 14 else None,
        }
    finally:
        if conn:
            conn.close()


def fetch_latest_cushion_batch_uid_by_date(tanggal_produksi):
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT batch_uid
            FROM grand_total
            WHERE tanggal_produksi = %s
              AND batch_uid IS NOT NULL
            ORDER BY id DESC
            LIMIT 1
            """,
            (tanggal_produksi,),
        )
        row = cur.fetchone()
        return row[0] if row and row[0] else None
    finally:
        if conn:
            conn.close()


def build_print_laporan_combined_context(cushion_result, gum_cord_row, batch_uid_label, download_url):
    header = (cushion_result or {}).get("header") or ()
    details = (cushion_result or {}).get("details") or []
    plastik = (cushion_result or {}).get("plastik")
    kotak = (cushion_result or {}).get("kotak")
    tungkul = (cushion_result or {}).get("tungkul")

    tanggal_obj = None
    if len(header) > 1 and header[1]:
        tanggal_obj = header[1]
    elif gum_cord_row and gum_cord_row.get("tanggal_produksi"):
        tanggal_obj = gum_cord_row.get("tanggal_produksi")

    tanggal = ""
    hari_tanggal = ""
    if tanggal_obj:
        tanggal = tanggal_obj.strftime("%d-%m-%Y")
        hari_map = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
        hari_tanggal = f"{hari_map[tanggal_obj.weekday()]} / {tanggal}"

    nama_operator = ""
    no_mesin = ""
    if len(header) > 2 and header[2]:
        nama_operator = header[2]
    elif gum_cord_row and gum_cord_row.get("nama_operator"):
        nama_operator = gum_cord_row.get("nama_operator") or ""

    if len(header) > 3 and header[3]:
        no_mesin = header[3]
    elif gum_cord_row and gum_cord_row.get("no_mesin"):
        no_mesin = gum_cord_row.get("no_mesin") or ""

    def empty_row():
        return {
            "nama_produk": "",
            "order_roll": "",
            "waktu_awal": "",
            "waktu_akhir": "",
            "line": "",
            "pakai_menit": "",
            "target_per_menit": "",
            "target_total": "",
            "aktual_roll": "",
            "persentase": "",
            "per_roll": "",
            "berat_total": "",
        }

    rows = []
    for row in details:
        rows.append(
            {
                "nama_produk": format_product_name_for_print((row[0] or "") if len(row) > 0 else ""),
                "order_roll": format_number_display(row[1] if len(row) > 1 else None),
                "waktu_awal": row[2].strftime("%H:%M") if len(row) > 2 and row[2] else "",
                "waktu_akhir": row[3].strftime("%H:%M") if len(row) > 3 and row[3] else "",
                "line": format_number_display(row[4] if len(row) > 4 else None),
                "pakai_menit": format_number_display(row[5] if len(row) > 5 else None),
                "target_per_menit": format_number_display(row[6] if len(row) > 6 else None),
                "target_total": format_number_display(row[7] if len(row) > 7 else None),
                "aktual_roll": format_number_display(row[8] if len(row) > 8 else None),
                "persentase": format_number_display(row[9] if len(row) > 9 else None),
                "per_roll": format_number_display(row[10] if len(row) > 10 else None),
                "berat_total": format_number_display(row[11] if len(row) > 11 else None),
            }
        )

    total_target = format_number_display(header[4] if len(header) > 4 else None)
    total_aktual = format_number_display(header[5] if len(header) > 5 else None)
    total_persen = format_number_display(header[6] if len(header) > 6 else None)
    total_berat = format_number_display(header[7] if len(header) > 7 else None)

    rows_per_page = 13
    row_chunks = [rows[i : i + rows_per_page] for i in range(0, len(rows), rows_per_page)] or [[]]

    pages = []
    total_pages = len(row_chunks)
    for index, chunk in enumerate(row_chunks):
        page_rows = list(chunk)
        if len(page_rows) < rows_per_page:
            page_rows.extend([empty_row() for _ in range(rows_per_page - len(page_rows))])
        pages.append({"rows": page_rows, "is_last": index == (total_pages - 1)})

    gum_cord_print = {
        "nama_produk": (gum_cord_row or {}).get("nama_produk") or "Gum Cord",
        "order_kotak": format_number_display((gum_cord_row or {}).get("order_kotak")),
        "waktu_awal": ((gum_cord_row or {}).get("waktu_awal").strftime("%H:%M") if (gum_cord_row or {}).get("waktu_awal") else ""),
        "waktu_akhir": ((gum_cord_row or {}).get("waktu_akhir").strftime("%H:%M") if (gum_cord_row or {}).get("waktu_akhir") else ""),
        "pakai_menit": format_number_display((gum_cord_row or {}).get("pakai_menit")),
        "target_per_menit": format_number_display((gum_cord_row or {}).get("target_per_menit")),
        "target_total": format_number_display((gum_cord_row or {}).get("target_total")),
        "aktual_kotak": format_number_display((gum_cord_row or {}).get("aktual_kotak")),
        "persentase": format_number_display((gum_cord_row or {}).get("persentase")),
        "berat_per_kotak": format_number_display((gum_cord_row or {}).get("berat_per_kotak")),
        "berat_total": format_number_display((gum_cord_row or {}).get("berat_total")),
    }

    plastik_print = {
        "230_blue": format_number_display(plastik[0] if plastik else None),
        "210_green": format_number_display(plastik[1] if plastik else None),
        "190_yellow": format_number_display(plastik[2] if plastik else None),
        "630_birupolos": format_number_display(plastik[3] if plastik else None),
        "630_red": format_number_display(plastik[4] if plastik else None),
        "270_red": format_number_display(plastik[5] if plastik else None),
        "240_red": format_number_display(plastik[6] if plastik else None),
        "plastik_gumcord": format_number_display(plastik[7] if plastik else None),
        "total_plastik": format_number_display(plastik[8] if plastik else None),
        "plastik_terbuang": format_number_display(plastik[9] if plastik else None),
        "plastik_terbuang_cgpotong": format_number_display(plastik[10] if plastik else None),
    }
    kotak_print = {
        "box_160": format_number_display(kotak[0] if kotak else None),
        "box_185": format_number_display(kotak[1] if kotak else None),
        "box_200": format_number_display(kotak[2] if kotak else None),
        "box_220": format_number_display(kotak[3] if kotak else None),
        "box_310": format_number_display(kotak[4] if kotak else None),
        "box_350": format_number_display(kotak[5] if kotak else None),
        "box_gumcord": format_number_display(kotak[6] if kotak else None),
        "total": format_number_display(kotak[7] if kotak else None),
        "terbuang": format_number_display(kotak[8] if kotak else None),
    }
    tungkul_print = {
        "tp_165": format_number_display(tungkul[0] if tungkul else None),
        "tp_195": format_number_display(tungkul[1] if tungkul else None),
        "tp_210": format_number_display(tungkul[2] if tungkul else None),
        "tp_240": format_number_display(tungkul[3] if tungkul else None),
        "total": format_number_display(tungkul[4] if tungkul else None),
        "terbuang": format_number_display(tungkul[5] if tungkul else None),
        "lakban": format_number_display(tungkul[6] if tungkul else None),
    }

    return {
        "download_url": download_url,
        "batch_uid": batch_uid_label or "-",
        "nama_operator": nama_operator,
        "no_mesin": no_mesin,
        "tanggal": tanggal,
        "hari_tanggal": hari_tanggal,
        "pages": pages,
        "total_target": total_target,
        "total_aktual": total_aktual,
        "total_persen": total_persen,
        "total_berat": total_berat,
        "plastik": plastik_print,
        "kotak": kotak_print,
        "tungkul": tungkul_print,
        "gum_cord": gum_cord_print,
    }


def render_print_laporan_combined(cushion_result, gum_cord_row, batch_uid_label, download_url):
    context = build_print_laporan_combined_context(
        cushion_result,
        gum_cord_row,
        batch_uid_label,
        download_url,
    )
    return render_template("print_laporan.html", **context)


def build_print_laporan_msc_context(result, batch_uid, download_url):
    header = result.get("header") or ()
    details = result.get("details") or []

    tanggal = ""
    hari_tanggal = ""
    if len(header) > 1 and header[1]:
        tanggal_obj = header[1]
        tanggal = tanggal_obj.strftime("%d-%m-%Y")
        hari_map = [
            "Senin",
            "Selasa",
            "Rabu",
            "Kamis",
            "Jumat",
            "Sabtu",
            "Minggu",
        ]
        hari_tanggal = f"{hari_map[tanggal_obj.weekday()]} / {tanggal}"

    rows = []
    for row in details:
        rows.append(
            {
                "nama_bahan": row[0] or "",
                "jam_awal": row[1].strftime("%H:%M") if row[1] else "",
                "jam_akhir": row[2].strftime("%H:%M") if row[2] else "",
                "pakai_menit": row[3] if row[3] is not None else "",
                "target_per_menit": format_number_display(row[4] if len(row) > 4 else None),
                "target_total": format_number_display(row[5] if len(row) > 5 else None),
                "aktual_batch": format_number_display(row[6] if len(row) > 6 else None),
                "persentase": format_number_display(row[7] if len(row) > 7 else None),
                "obat_timbang": format_number_display(row[8] if len(row) > 8 else None),
                "obat_sisa": format_number_display(row[9] if len(row) > 9 else None),
                "keterangan": row[10] or "",
            }
        )

    max_rows = 12
    if len(rows) < max_rows:
        rows.extend(
            [
                {
                    "nama_bahan": "",
                    "jam_awal": "",
                    "jam_akhir": "",
                    "pakai_menit": "",
                    "target_per_menit": "",
                    "target_total": "",
                    "aktual_batch": "",
                    "persentase": "",
                    "obat_timbang": "",
                    "obat_sisa": "",
                    "keterangan": "",
                }
                for _ in range(max_rows - len(rows))
            ]
        )
    else:
        rows = rows[:max_rows]

    return {
        "download_url": download_url,
        "batch_uid": batch_uid,
        "nama_operator": (header[2] or "") if len(header) > 2 else "",
        "no_mesin": (header[3] or "") if len(header) > 3 else "",
        "regu": (header[4] or "") if len(header) > 4 else "",
        "tanggal": tanggal,
        "hari_tanggal": hari_tanggal,
        "rows": rows,
        "total_pakai_menit": format_number_display(header[5] if len(header) > 5 else None),
        "total_target": format_number_display(header[6] if len(header) > 6 else None),
        "total_aktual": format_number_display(header[7] if len(header) > 7 else None),
        "total_persen": format_number_display(header[8] if len(header) > 8 else None),
    }


@app.route("/laporan/cushion-gum/cetak/<batch_uid>", methods=["GET"])
def laporan_cushion_cetak(batch_uid):
    if "user" not in session:
        return redirect(url_for("login"))

    result = fetch_cushion_batch(batch_uid)
    if not result:
        return "Data batch Cushion Gum tidak ditemukan.", 404
    header = result.get("header") or ()
    tanggal_produksi = header[1] if len(header) > 1 else None
    gum_cord_row = fetch_latest_gum_cord_by_date(tanggal_produksi) if tanggal_produksi else None
    return render_print_laporan_combined(
        result,
        gum_cord_row,
        batch_uid,
        url_for("laporan_cushion_download", batch_uid=batch_uid),
    )


@app.route("/laporan/gum-cord/cetak/<row_token>", methods=["GET"])
def laporan_gum_cord_cetak(row_token):
    if "user" not in session:
        return redirect(url_for("login"))

    gum_cord_row = fetch_gum_cord_by_row_token(row_token)
    if not gum_cord_row:
        return "Data Gum Cord tidak ditemukan.", 404

    tanggal_produksi = gum_cord_row.get("tanggal_produksi")
    cushion_result = None
    batch_uid = "-"
    if tanggal_produksi:
        batch_uid_by_date = fetch_latest_cushion_batch_uid_by_date(tanggal_produksi)
        if batch_uid_by_date:
            cushion_result = fetch_cushion_batch(batch_uid_by_date)
            batch_uid = batch_uid_by_date

    return render_print_laporan_combined(
        cushion_result,
        gum_cord_row,
        batch_uid,
        url_for("laporan_gum_cord_download", row_token=row_token),
    )


@app.route("/laporan/cushion-gum/download/<batch_uid>", methods=["GET"])
def laporan_cushion_download(batch_uid):
    if "user" not in session:
        return redirect(url_for("login"))

    result = fetch_cushion_batch(batch_uid)
    if not result:
        return "Data batch Cushion Gum tidak ditemukan.", 404

    header = result.get("header") or ()
    tanggal_produksi = header[1] if len(header) > 1 else None
    gum_cord_row = fetch_latest_gum_cord_by_date(tanggal_produksi) if tanggal_produksi else None
    context = build_print_laporan_combined_context(result, gum_cord_row, batch_uid, "")
    pdf_buffer = render_template_to_pdf_bytes("print_laporan.html", context, "print.css")
    return send_file(
        pdf_buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"laporan-cushion-gum-{batch_uid}.pdf",
    )


@app.route("/laporan/gum-cord/download/<row_token>", methods=["GET"])
def laporan_gum_cord_download(row_token):
    if "user" not in session:
        return redirect(url_for("login"))

    gum_cord_row = fetch_gum_cord_by_row_token(row_token)
    if not gum_cord_row:
        return "Data Gum Cord tidak ditemukan.", 404

    tanggal_produksi = gum_cord_row.get("tanggal_produksi")
    cushion_result = None
    batch_uid = "-"
    if tanggal_produksi:
        batch_uid_by_date = fetch_latest_cushion_batch_uid_by_date(tanggal_produksi)
        if batch_uid_by_date:
            cushion_result = fetch_cushion_batch(batch_uid_by_date)
            batch_uid = batch_uid_by_date

    context = build_print_laporan_combined_context(cushion_result, gum_cord_row, batch_uid, "")
    pdf_buffer = render_template_to_pdf_bytes("print_laporan.html", context, "print.css")
    safe_token = row_token.replace("(", "").replace(")", "").replace(",", "-")
    return send_file(
        pdf_buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"laporan-gum-cord-{safe_token}.pdf",
    )


@app.route("/laporan/msc/cetak/<batch_uid>", methods=["GET"])
def laporan_msc_cetak(batch_uid):
    if "user" not in session:
        return redirect(url_for("login"))

    result = fetch_msc_batch(batch_uid)
    if not result:
        return "Data batch MSC tidak ditemukan.", 404

    context = build_print_laporan_msc_context(
        result,
        batch_uid,
        url_for("laporan_msc_download", batch_uid=batch_uid),
    )
    return render_template("print_laporan_msc.html", **context)


@app.route("/laporan/msc/download/<batch_uid>", methods=["GET"])
def laporan_msc_download(batch_uid):
    if "user" not in session:
        return redirect(url_for("login"))

    result = fetch_msc_batch(batch_uid)
    if not result:
        return "Data batch MSC tidak ditemukan.", 404

    context = build_print_laporan_msc_context(result, batch_uid, "")
    pdf_buffer = render_template_to_pdf_bytes("print_laporan_msc.html", context, "print_msc.css")
    return send_file(
        pdf_buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"laporan-msc-{batch_uid}.pdf",
    )


@app.route("/laporan/delete", methods=["POST"])
def laporan_delete():
    if "user" not in session:
        return redirect(url_for("login"))

    data_key = (request.form.get("id") or "").strip()
    sumber = (request.form.get("sumber") or "").strip()

    table_map = {
        "cushion-gum": "grand_total",
        "gum-cord": "production_gum_cord",
        "msc": "grand_total_msc",
    }
    table_name = table_map.get(sumber)
    if not table_name or not data_key:
        return redirect(url_for("laporan"))

    conn = None
    try:
        conn = get_db_conn()
        ensure_cushion_batch_columns(conn)
        ensure_pemakaian_plastik_table(conn)
        ensure_pemakaian_kotak_table(conn)
        ensure_pemakaian_tungkul_table(conn)
        cur = conn.cursor()
        if sumber == "cushion-gum":
            data_id = parse_int(data_key)
            if data_id is None:
                return redirect(url_for("laporan", tab=sumber))

            cur.execute(
                """
                SELECT batch_uid
                FROM grand_total
                WHERE id = %s
                """,
                (data_id,),
            )
            row = cur.fetchone()
            if row:
                batch_uid = row[0]
                cur.execute("DELETE FROM grand_total WHERE id = %s", (data_id,))
                if batch_uid:
                    cur.execute(
                        """
                        DELETE FROM production_cushion_gum
                        WHERE batch_uid = %s
                        """,
                        (batch_uid,),
                    )
                    cur.execute(
                        """
                        DELETE FROM pemakaian_plastik
                        WHERE batch_uid = %s
                        """,
                        (batch_uid,),
                    )
                    cur.execute(
                        """
                        DELETE FROM pemakaian_kotak
                        WHERE batch_uid = %s
                        """,
                        (batch_uid,),
                    )
                    cur.execute(
                        """
                        DELETE FROM pemakaian_tungkul
                        WHERE batch_uid = %s
                        """,
                        (batch_uid,),
                    )
        elif sumber == "gum-cord":
            # production_gum_cord tidak punya kolom id, jadi pakai row token PostgreSQL (ctid).
            cur.execute(
                "DELETE FROM production_gum_cord WHERE ctid = %s::tid",
                (data_key,),
            )
        elif sumber == "msc":
            data_id = parse_int(data_key)
            if data_id is None:
                return redirect(url_for("laporan", tab=sumber))

            cur.execute(
                """
                SELECT batch_uid
                FROM grand_total_msc
                WHERE id = %s
                """,
                (data_id,),
            )
            row = cur.fetchone()
            if row:
                batch_uid = row[0]
                cur.execute("DELETE FROM grand_total_msc WHERE id = %s", (data_id,))
                if batch_uid:
                    cur.execute(
                        """
                        DELETE FROM production_msc
                        WHERE batch_uid = %s
                        """,
                        (batch_uid,),
                    )
        else:
            data_id = parse_int(data_key)
            if data_id is None:
                return redirect(url_for("laporan", tab=sumber))
            cur.execute(f"DELETE FROM {table_name} WHERE id = %s", (data_id,))
        conn.commit()
    except Exception:
        if conn:
            conn.rollback()
        raise
    finally:
        if conn:
            conn.close()

    return redirect(url_for("laporan", tab=sumber))

@app.route("/cushion-gum", methods=["GET", "POST"])
def cushion_gum():
    if "user" not in session:
        return redirect(url_for("login"))

    init_conn = None
    try:
        init_conn = get_db_conn()
        ensure_pemakaian_plastik_table(init_conn)
        ensure_pemakaian_kotak_table(init_conn)
        ensure_pemakaian_tungkul_table(init_conn)
        init_conn.commit()
    finally:
        if init_conn:
            init_conn.close()

    if request.method == "POST":
        tanggal_produksi = request.form.get("tanggal_produksi")
        if not tanggal_produksi:
            return "Tanggal produk wajib diisi.", 400
        nama_operator = (request.form.get("nama_operator") or "").strip()
        no_mesin = (request.form.get("no_mesin") or "").strip()
        print(
            "[CUSHION_GUM] submit:",
            "tanggal_produksi=", tanggal_produksi,
            "nama_operator=", repr(nama_operator),
            "no_mesin=", repr(no_mesin),
        )

        nama_produk_list = request.form.getlist("nama_produk[]")
        order_roll_list = request.form.getlist("order_roll[]")
        waktu_awal_list = request.form.getlist("waktu_awal[]")
        waktu_akhir_list = request.form.getlist("waktu_akhir[]")
        line_list = request.form.getlist("line[]")
        pakai_menit_list = request.form.getlist("pakai_menit[]")
        target_per_menit_list = request.form.getlist("target_per_menit[]")
        aktual_roll_list = request.form.getlist("aktual_roll[]")
        per_roll_list = request.form.getlist("per_roll[]")
        berat_total_list = request.form.getlist("berat_total[]")
        berat_kg_total = parse_decimal(request.form.get("berat_kg_total"))
        plastik_230_blue = parse_decimal(request.form.get("230_blue"))
        plastik_210_green = parse_decimal(request.form.get("210_green"))
        plastik_190_yellow = parse_decimal(request.form.get("190_yellow"))
        plastik_630_birupolos = parse_decimal(request.form.get("630_birupolos"))
        plastik_630_red = parse_decimal(request.form.get("630_red"))
        plastik_270_red = parse_decimal(request.form.get("270_red"))
        plastik_240_red = parse_decimal(request.form.get("240_red"))
        plastik_gumcord = parse_decimal(request.form.get("plastik_gumcord"))
        total_plastik = parse_decimal(request.form.get("total_plastik"))
        plastik_terbuang = parse_decimal(request.form.get("plastik_terbuang"))
        plastik_terbuang_cgpotong = parse_decimal(
            request.form.get("plastik_terbuang_cgpotong")
        )
        box_160 = parse_int(request.form.get("box_160"))
        box_185 = parse_int(request.form.get("box_185"))
        box_200 = parse_int(request.form.get("box_200"))
        box_220 = parse_int(request.form.get("box_220"))
        box_310 = parse_int(request.form.get("box_310"))
        box_350 = parse_int(request.form.get("box_350"))
        box_gumcord = parse_int(request.form.get("box_gumcord"))
        kotak_total = parse_int(request.form.get("total"))
        kotak_terbuang = parse_int(request.form.get("terbuang"))
        tp_165 = parse_int(request.form.get("tp_165"))
        tp_195 = parse_int(request.form.get("tp_195"))
        tp_210 = parse_int(request.form.get("tp_210"))
        tp_240 = parse_int(request.form.get("tp_240"))
        tungkul_total = parse_int(request.form.get("tungkul_total"))
        tungkul_terbuang = parse_int(request.form.get("tungkul_terbuang"))
        lakban = parse_int(request.form.get("lakban"))
        tungkul_items = [tp_165, tp_195, tp_210, tp_240]
        if any(value is not None for value in tungkul_items):
            tungkul_total = sum(value or 0 for value in tungkul_items)

        max_len = max(
            len(nama_produk_list),
            len(order_roll_list),
            len(waktu_awal_list),
            len(waktu_akhir_list),
            len(line_list),
            len(pakai_menit_list),
            len(target_per_menit_list),
            len(aktual_roll_list),
            len(per_roll_list),
            len(berat_total_list),
        )

        batch_uid = None
        rows = []
        for i in range(max_len):
            nama_produk = (
                nama_produk_list[i].strip() if i < len(nama_produk_list) else ""
            )
            order_roll = parse_int(order_roll_list[i]) if i < len(order_roll_list) else None
            waktu_awal = waktu_awal_list[i] if i < len(waktu_awal_list) else ""
            waktu_akhir = waktu_akhir_list[i] if i < len(waktu_akhir_list) else ""
            line_val = parse_int(line_list[i]) if i < len(line_list) else None
            pakai_menit = (
                parse_int(pakai_menit_list[i]) if i < len(pakai_menit_list) else None
            )
            target_code = (
                target_per_menit_list[i] if i < len(target_per_menit_list) else ""
            )
            aktual_roll = (
                parse_int(aktual_roll_list[i]) if i < len(aktual_roll_list) else None
            )
            per_roll = parse_decimal(per_roll_list[i]) if i < len(per_roll_list) else None
            berat_total = parse_decimal(berat_total_list[i]) if i < len(berat_total_list) else None

            if not any(
                [
                    nama_produk,
                    order_roll is not None,
                    waktu_awal,
                    waktu_akhir,
                    target_code,
                    aktual_roll is not None,
                    pakai_menit is not None,
                    per_roll is not None,
                    berat_total is not None,
                ]
            ):
                continue

            if pakai_menit is None and waktu_awal and waktu_akhir:
                try:
                    t_awal = datetime.strptime(waktu_awal, "%H:%M")
                    t_akhir = datetime.strptime(waktu_akhir, "%H:%M")
                    diff = int((t_akhir - t_awal).total_seconds() / 60)
                    if diff >= 0:
                        pakai_menit = diff
                except ValueError:
                    pakai_menit = None

            target_per_menit = TARGET_MAP.get(target_code)
            if line_val is None:
                line_val = LINE_MAP.get(target_code)

            target_total = None
            if pakai_menit is not None and target_per_menit is not None:
                target_total = (Decimal(pakai_menit) * target_per_menit).quantize(
                    Decimal("0.01")
                )

            persentase = None
            if target_total is not None and target_total != 0 and aktual_roll is not None:
                persentase = (
                    (Decimal(aktual_roll) / target_total) * Decimal("100")
                ).quantize(Decimal("0.01"))
            if per_roll is None:
                per_roll = infer_per_roll_from_product(nama_produk)
            if berat_total is None and aktual_roll is not None and per_roll is not None:
                berat_total = (Decimal(aktual_roll) * per_roll).quantize(Decimal("0.01"))

            rows.append(
                (
                    tanggal_produksi,
                    nama_produk or None,
                    order_roll,
                    waktu_awal or None,
                    waktu_akhir or None,
                    line_val,
                    pakai_menit,
                    target_per_menit,
                    target_total,
                    aktual_roll,
                    persentase,
                    per_roll,
                    berat_total,
                    None,
                )
            )

        conn = None
        try:
            conn = get_db_conn()
            ensure_cushion_batch_columns(conn)
            ensure_pemakaian_plastik_table(conn)
            ensure_pemakaian_kotak_table(conn)
            ensure_pemakaian_tungkul_table(conn)
            batch_uid = generate_batch_uid(conn, tanggal_produksi, "CG", "grand_total")
            rows = [row[:-1] + (batch_uid,) for row in rows]
            cur = conn.cursor()

            if rows:
                insert_sql = """
                    INSERT INTO production_cushion_gum (
                        tanggal_produksi,
                        nama_produk,
                        order_roll,
                        waktu_awal,
                        waktu_akhir,
                        line,
                        pakai_menit,
                        target_per_menit,
                        target_total,
                        aktual_roll,
                        persentase,
                        per_roll,
                        berat_total,
                        batch_uid
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                """
                cur.executemany(insert_sql, rows)

            # Grand total disamakan dengan total yang sedang tersaji di form saat simpan,
            # bukan akumulasi semua simpan di tanggal yang sama.
            total_target = sum(
                (row[8] if row[8] is not None else Decimal("0")) for row in rows
            )
            total_aktual = sum(
                (row[9] if row[9] is not None else 0) for row in rows
            )
            if berat_kg_total is None:
                total_berat = sum(
                    (row[12] if row[12] is not None else Decimal("0")) for row in rows
                )
                berat_kg_total = total_berat if total_berat != 0 else None

            total_persentase = None
            if total_target != 0:
                total_persentase = (
                    (Decimal(total_aktual) / total_target) * Decimal("100")
                ).quantize(Decimal("0.01"))

            insert_grand_sql = """
                INSERT INTO grand_total (
                    tanggal_produksi,
                    nama_operator,
                    no_mesin,
                    total_target,
                    total_aktual,
                    total_persentase,
                    berat_kg_total,
                    batch_uid
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s
                )
            """
            cur.execute(
                insert_grand_sql,
                (
                    tanggal_produksi,
                    nama_operator or None,
                    no_mesin or None,
                    total_target,
                    total_aktual,
                    total_persentase,
                    berat_kg_total,
                    batch_uid,
                ),
            )

            plastik_values = [
                plastik_230_blue,
                plastik_210_green,
                plastik_190_yellow,
                plastik_630_birupolos,
                plastik_630_red,
                plastik_270_red,
                plastik_240_red,
                plastik_gumcord,
                total_plastik,
                plastik_terbuang,
                plastik_terbuang_cgpotong,
            ]
            if any(value is not None for value in plastik_values):
                cur.execute(
                    """
                    INSERT INTO pemakaian_plastik (
                        tanggal_produksi,
                        batch_uid,
                        "230_blue",
                        "210_green",
                        "190_yellow",
                        "630_birupolos",
                        "630_red",
                        "270_red",
                        "240_red",
                        plastik_gumcord,
                        total_plastik,
                        plastik_terbuang,
                        plastik_terbuang_cgpotong
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                    """,
                    (tanggal_produksi, batch_uid, *plastik_values),
                )

            kotak_values = [
                box_160,
                box_185,
                box_200,
                box_220,
                box_310,
                box_350,
                box_gumcord,
                kotak_total,
                kotak_terbuang,
            ]
            if any(value is not None for value in kotak_values):
                cur.execute(
                    """
                    INSERT INTO pemakaian_kotak (
                        tanggal_produksi,
                        batch_uid,
                        box_160,
                        box_185,
                        box_200,
                        box_220,
                        box_310,
                        box_350,
                        box_gumcord,
                        total,
                        terbuang
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                    """,
                    (tanggal_produksi, batch_uid, *kotak_values),
                )

            tungkul_values = [
                tp_165,
                tp_195,
                tp_210,
                tp_240,
                tungkul_total,
                tungkul_terbuang,
                lakban,
            ]
            if any(value is not None for value in tungkul_values):
                cur.execute(
                    """
                    INSERT INTO pemakaian_tungkul (
                        tanggal_produksi,
                        batch_uid,
                        tp_165,
                        tp_195,
                        tp_210,
                        tp_240,
                        total,
                        terbuang,
                        lakban
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                    """,
                    (tanggal_produksi, batch_uid, *tungkul_values),
                )

            conn.commit()
        except Exception:
            if conn:
                conn.rollback()
            raise
        finally:
            if conn:
                conn.close()

        return redirect(url_for("cushion_gum", saved="1"))

    master_produk = fetch_master_produk()
    return render_template("cushion_gum.html", master_produk=master_produk)

@app.route("/cushion-gum-cord", methods=["POST"])
def cushion_gum_cord():
    if "user" not in session:
        return redirect(url_for("login"))

    tanggal_produksi = request.form.get("tanggal_produksi")
    if not tanggal_produksi:
        return "Tanggal produk wajib diisi.", 400

    nama_operator = (request.form.get("nama_operator") or "").strip() or None
    no_mesin = (request.form.get("no_mesin") or "").strip() or None
    nama_produk = request.form.get("nama_produk")
    order_kotak = parse_int(request.form.get("order_kotak"))
    waktu_awal = request.form.get("waktu_awal") or None
    waktu_akhir = request.form.get("waktu_akhir") or None
    pakai_menit = parse_int(request.form.get("pakai_menit"))
    target_per_menit = parse_decimal(request.form.get("target_per_menit"))
    target_total = parse_decimal(request.form.get("target_total"))
    aktual_kotak = parse_int(request.form.get("aktual_kotak"))
    persentase = parse_decimal(request.form.get("persentase"))
    berat_per_kotak = parse_decimal(request.form.get("berat_per_kotak"))
    berat_total = parse_decimal(request.form.get("berat_total"))

    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS nama_operator VARCHAR(150)
            """
        )
        cur.execute(
            """
            ALTER TABLE production_gum_cord
            ADD COLUMN IF NOT EXISTS no_mesin VARCHAR(100)
            """
        )
        insert_sql = """
            INSERT INTO production_gum_cord (
                tanggal_produksi,
                nama_operator,
                no_mesin,
                nama_produk,
                order_kotak,
                waktu_awal,
                waktu_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_kotak,
                persentase,
                berat_per_kotak,
                berat_total
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
        """
        cur.execute(
            insert_sql,
            (
                tanggal_produksi,
                nama_operator,
                no_mesin,
                nama_produk,
                order_kotak,
                waktu_awal,
                waktu_akhir,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_kotak,
                persentase,
                berat_per_kotak,
                berat_total,
            ),
        )
        conn.commit()
    except Exception:
        if conn:
            conn.rollback()
        raise
    finally:
        if conn:
            conn.close()

    return redirect(url_for("gum_cord", saved="1"))

@app.route("/msc", methods=["GET", "POST"])
def msc():
    if "user" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        tanggal_produksi = request.form.get("tanggal_produksi")
        if not tanggal_produksi:
            return "Tanggal produksi wajib diisi.", 400
        nama_operator = (request.form.get("nama_operator") or "").strip()
        no_mesin = (request.form.get("no_mesin") or "").strip()
        regu = (request.form.get("regu") or "").strip()
        batch_uid = None

        nama_bahan_list = request.form.getlist("nama_bahan[]")
        jam_awal_list = request.form.getlist("jam_awal[]")
        jam_akhir_list = request.form.getlist("jam_akhir[]")
        pakai_menit_list = request.form.getlist("pakai_menit[]")
        target_per_menit_list = request.form.getlist("target_per_menit[]")
        target_total_list = request.form.getlist("target_total[]")
        aktual_batch_list = request.form.getlist("aktual_batch[]")
        persentase_list = request.form.getlist("persentase[]")
        obat_timbang_list = request.form.getlist("obat_timbang[]")
        obat_sisa_list = request.form.getlist("obat_sisa[]")
        keterangan_list = request.form.getlist("keterangan[]")

        max_len = max(
            len(nama_bahan_list),
            len(jam_awal_list),
            len(jam_akhir_list),
            len(pakai_menit_list),
            len(target_per_menit_list),
            len(target_total_list),
            len(aktual_batch_list),
            len(persentase_list),
            len(obat_timbang_list),
            len(obat_sisa_list),
            len(keterangan_list),
        )

        rows = []
        for i in range(max_len):
            nama_bahan = (
                nama_bahan_list[i].strip() if i < len(nama_bahan_list) else ""
            )
            jam_awal = jam_awal_list[i] if i < len(jam_awal_list) else ""
            jam_akhir = jam_akhir_list[i] if i < len(jam_akhir_list) else ""
            pakai_menit = (
                parse_int(pakai_menit_list[i]) if i < len(pakai_menit_list) else None
            )
            target_per_menit = (
                parse_decimal(target_per_menit_list[i])
                if i < len(target_per_menit_list)
                else None
            )
            target_total = (
                parse_decimal(target_total_list[i]) if i < len(target_total_list) else None
            )
            aktual_batch = (
                parse_decimal(aktual_batch_list[i]) if i < len(aktual_batch_list) else None
            )
            persentase = (
                parse_decimal(persentase_list[i]) if i < len(persentase_list) else None
            )
            obat_timbang = (
                parse_decimal(obat_timbang_list[i]) if i < len(obat_timbang_list) else None
            )
            obat_sisa = (
                parse_decimal(obat_sisa_list[i]) if i < len(obat_sisa_list) else None
            )
            keterangan = (
                keterangan_list[i].strip() if i < len(keterangan_list) else ""
            )

            if not any(
                [
                    nama_bahan,
                    jam_awal,
                    jam_akhir,
                    pakai_menit is not None,
                    target_per_menit is not None,
                    target_total is not None,
                    aktual_batch is not None,
                    persentase is not None,
                    obat_timbang is not None,
                    obat_sisa is not None,
                    keterangan,
                ]
            ):
                continue

            rows.append(
                (
                    tanggal_produksi,
                    nama_bahan or None,
                    jam_awal or None,
                    jam_akhir or None,
                    pakai_menit,
                    target_per_menit,
                    target_total,
                    aktual_batch,
                    persentase,
                    obat_timbang,
                    obat_sisa,
                    keterangan or None,
                    None,
                )
            )

        if rows:
            conn = None
            try:
                conn = get_db_conn()
                ensure_grand_total_msc_table(conn)
                batch_uid = generate_batch_uid(
                    conn, tanggal_produksi, "MSC", "grand_total_msc"
                )
                rows = [row[:-1] + (batch_uid,) for row in rows]
                cur = conn.cursor()
                insert_sql = """
                    INSERT INTO production_msc (
                        tanggal_produksi,
                        nama_bahan,
                        jam_awal,
                        jam_akhir,
                        pakai_menit,
                        target_per_menit,
                        target_total,
                        aktual_batch,
                        persentase,
                        obat_timbang,
                        obat_sisa,
                        keterangan,
                        batch_uid
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                """
                cur.executemany(insert_sql, rows)

                total_pakai_menit = sum(
                    (row[4] if row[4] is not None else 0) for row in rows
                )
                total_target = sum(
                    (row[6] if row[6] is not None else Decimal("0")) for row in rows
                )
                total_aktual = sum(
                    (row[7] if row[7] is not None else Decimal("0")) for row in rows
                )

                total_persentase = None
                if total_target != 0:
                    total_persentase = (
                        (Decimal(total_aktual) / Decimal(total_target)) * Decimal("100")
                    ).quantize(Decimal("0.01"))

                upsert_total_sql = """
                    INSERT INTO grand_total_msc (
                        tanggal_produksi,
                        nama_operator,
                        no_mesin,
                        regu,
                        total_pakai_menit,
                        total_target,
                        total_aktual,
                        total_persentase,
                        batch_uid
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                """
                cur.execute(
                    upsert_total_sql,
                    (
                        tanggal_produksi,
                        nama_operator or None,
                        no_mesin or None,
                        regu or None,
                        total_pakai_menit,
                        total_target,
                        total_aktual,
                        total_persentase,
                        batch_uid,
                    ),
                )
                conn.commit()
            except Exception:
                if conn:
                    conn.rollback()
                raise
            finally:
                if conn:
                    conn.close()

        return redirect(url_for("msc", saved="1"))

    master_bahan_msc = fetch_master_bahan_msc()
    return render_template("msc.html", master_bahan_msc=master_bahan_msc)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


if __name__ == "__main__":
    app.run(debug=True)


