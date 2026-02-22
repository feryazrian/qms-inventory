from flask import Flask, render_template, request, redirect, url_for, session, jsonify
from datetime import datetime
from decimal import Decimal, InvalidOperation
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
    "CG1": Decimal("0.30"),
    "CG2": Decimal("0.53"),
}

LINE_MAP = {
    "CG1": 1,
    "CG2": 2,
    "1/8": None,
    "CGP": None,
}


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
                total_persentase
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
                waktu_awal,
                waktu_akhir,
                line,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_roll,
                persentase
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
                "rows": [
                    {
                        "nama_produk": row[0] or "",
                        "waktu_awal": row[1].strftime("%H:%M") if row[1] else "",
                        "waktu_akhir": row[2].strftime("%H:%M") if row[2] else "",
                        "line": row[3],
                        "pakai_menit": row[4],
                        "target_per_menit": float(row[5]) if row[5] is not None else None,
                        "target_total": float(row[6]) if row[6] is not None else None,
                        "aktual_roll": row[7],
                        "persentase": float(row[8]) if row[8] is not None else None,
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

    parsed_rows = []
    total_target = Decimal("0")
    total_aktual = 0

    for row in rows:
        row = row or {}
        nama_produk = (row.get("nama_produk") or "").strip() or None
        waktu_awal = (row.get("waktu_awal") or "").strip() or None
        waktu_akhir = (row.get("waktu_akhir") or "").strip() or None
        line_val = parse_int(str(row.get("line") or ""))
        pakai_menit = parse_int(str(row.get("pakai_menit") or ""))
        target_per_menit = parse_decimal(str(row.get("target_per_menit") or ""))
        target_total = parse_decimal(str(row.get("target_total") or ""))
        aktual_roll = parse_int(str(row.get("aktual_roll") or ""))
        persentase = parse_decimal(str(row.get("persentase") or ""))

        if not any(
            [
                nama_produk,
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

        total_target += target_total or Decimal("0")
        total_aktual += aktual_roll or 0

        parsed_rows.append(
            (
                tanggal_produksi,
                nama_produk,
                waktu_awal,
                waktu_akhir,
                line_val,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_roll,
                persentase,
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
                total_persentase = %s
            WHERE batch_uid = %s
            """,
            (
                tanggal_produksi,
                nama_operator,
                no_mesin,
                total_target,
                total_aktual,
                total_persentase,
                batch_uid,
            ),
        )
        cur.execute("DELETE FROM production_cushion_gum WHERE batch_uid = %s", (batch_uid,))
        cur.executemany(
            """
            INSERT INTO production_cushion_gum (
                tanggal_produksi,
                nama_produk,
                waktu_awal,
                waktu_akhir,
                line,
                pakai_menit,
                target_per_menit,
                target_total,
                aktual_roll,
                persentase,
                batch_uid
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
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
        waktu_awal_list = request.form.getlist("waktu_awal[]")
        waktu_akhir_list = request.form.getlist("waktu_akhir[]")
        line_list = request.form.getlist("line[]")
        pakai_menit_list = request.form.getlist("pakai_menit[]")
        target_per_menit_list = request.form.getlist("target_per_menit[]")
        aktual_roll_list = request.form.getlist("aktual_roll[]")

        max_len = max(
            len(nama_produk_list),
            len(waktu_awal_list),
            len(waktu_akhir_list),
            len(line_list),
            len(pakai_menit_list),
            len(target_per_menit_list),
            len(aktual_roll_list),
        )

        batch_uid = None
        rows = []
        for i in range(max_len):
            nama_produk = (
                nama_produk_list[i].strip() if i < len(nama_produk_list) else ""
            )
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

            if not any(
                [
                    nama_produk,
                    waktu_awal,
                    waktu_akhir,
                    target_code,
                    aktual_roll is not None,
                    pakai_menit is not None,
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

            rows.append(
                (
                    tanggal_produksi,
                    nama_produk or None,
                    waktu_awal or None,
                    waktu_akhir or None,
                    line_val,
                    pakai_menit,
                    target_per_menit,
                    target_total,
                    aktual_roll,
                    persentase,
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
                        waktu_awal,
                        waktu_akhir,
                        line,
                        pakai_menit,
                        target_per_menit,
                        target_total,
                        aktual_roll,
                        persentase,
                        batch_uid
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                """
                cur.executemany(insert_sql, rows)

            # Grand total disamakan dengan total yang sedang tersaji di form saat simpan,
            # bukan akumulasi semua simpan di tanggal yang sama.
            total_target = sum(
                (row[7] if row[7] is not None else Decimal("0")) for row in rows
            )
            total_aktual = sum(
                (row[8] if row[8] is not None else 0) for row in rows
            )

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
                    batch_uid
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s
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

        return redirect(url_for("cushion_gum"))

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

    return redirect(url_for("cushion_gum"))

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

        return redirect(url_for("msc"))

    master_bahan_msc = fetch_master_bahan_msc()
    return render_template("msc.html", master_bahan_msc=master_bahan_msc)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


if __name__ == "__main__":
    app.run(debug=True)

