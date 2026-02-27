from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
from datetime import datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
import os
import psycopg2

app = Flask(__name__)
app.secret_key = "secret123"

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
        return {"header": header, "details": details}
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

    if not rows:
        return jsonify({"ok": False, "message": "Data detail Cushion Gum kosong."}), 400

    tanggal_produksi = (header.get("tanggal_produksi") or "").strip()
    if not tanggal_produksi:
        return jsonify({"ok": False, "message": "Tanggal produksi wajib diisi."}), 400

    nama_operator = (header.get("nama_operator") or "").strip() or None
    no_mesin = (header.get("no_mesin") or "").strip() or None
    berat_kg_total = parse_decimal(str(header.get("berat_kg_total") or ""))

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
        conn.commit()
    except Exception as e:
        if conn:
            conn.rollback()
        return jsonify({"ok": False, "message": f"Gagal update: {e}"}), 500
    finally:
        if conn:
            conn.close()

    return jsonify({"ok": True, "message": "Data Cushion Gum berhasil diupdate."})


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


@app.route("/item-code", methods=["GET", "POST"])
def item_code():
    if "user" not in session:
        return redirect(url_for("login"))

    notice = ""
    error = ""

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
                return redirect(url_for("item_code", notice=notice))
            conn.rollback()
        except Exception as e:
            if conn:
                conn.rollback()
            error = f"Gagal menyimpan produk: {e}"
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


@app.route("/laporan/cushion-gum/cetak/<batch_uid>", methods=["GET"])
def laporan_cushion_cetak(batch_uid):
    if "user" not in session:
        return redirect(url_for("login"))

    result = fetch_cushion_batch(batch_uid)
    if not result:
        return "Data batch Cushion Gum tidak ditemukan.", 404

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
    row_chunks = [
        rows[i : i + rows_per_page] for i in range(0, len(rows), rows_per_page)
    ] or [[]]

    pages = []
    total_pages = len(row_chunks)
    for index, chunk in enumerate(row_chunks):
        page_rows = list(chunk)
        if len(page_rows) < rows_per_page:
            page_rows.extend([empty_row() for _ in range(rows_per_page - len(page_rows))])
        pages.append(
            {
                "rows": page_rows,
                "is_last": index == (total_pages - 1),
            }
        )

    return render_template(
        "print_laporan.html",
        batch_uid=batch_uid,
        nama_operator=(header[2] or "") if len(header) > 2 else "",
        no_mesin=(header[3] or "") if len(header) > 3 else "",
        tanggal=tanggal,
        hari_tanggal=hari_tanggal,
        pages=pages,
        total_target=total_target,
        total_aktual=total_aktual,
        total_persen=total_persen,
        total_berat=total_berat,
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
        insert_sql = """
            INSERT INTO production_gum_cord (
                tanggal_produksi,
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
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
        """
        cur.execute(
            insert_sql,
            (
                tanggal_produksi,
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

    return redirect(url_for("cushion_gum", saved="1"))

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


