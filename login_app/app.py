from flask import Flask, render_template, request, redirect, url_for, session
from datetime import datetime
from decimal import Decimal, InvalidOperation
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

@app.route("/cushion-gum", methods=["GET", "POST"])
def cushion_gum():
    if "user" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        tanggal_produksi = request.form.get("tanggal_produksi")
        if not tanggal_produksi:
            return "Tanggal produk wajib diisi.", 400

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
                )
            )

        if rows:
            conn = None
            try:
                conn = get_db_conn()
                cur = conn.cursor()

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
                        persentase
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                """
                cur.executemany(insert_sql, rows)

                cur.execute(
                    """
                    SELECT
                        COALESCE(SUM(target_total), 0),
                        COALESCE(SUM(aktual_roll), 0)
                    FROM production_cushion_gum
                    WHERE tanggal_produksi = %s
                    """,
                    (tanggal_produksi,),
                )
                total_target, total_aktual = cur.fetchone()

                total_target = Decimal(total_target) if total_target is not None else Decimal("0")
                total_aktual = int(total_aktual) if total_aktual is not None else 0

                total_persentase = None
                if total_target != 0:
                    total_persentase = (
                        (Decimal(total_aktual) / total_target) * Decimal("100")
                    ).quantize(Decimal("0.01"))

                upsert_sql = """
                    INSERT INTO grand_total (
                        tanggal_produksi,
                        total_target,
                        total_aktual,
                        total_persentase
                    ) VALUES (
                        %s, %s, %s, %s
                    )
                    ON CONFLICT (tanggal_produksi)
                    DO UPDATE SET
                        total_target = EXCLUDED.total_target,
                        total_aktual = EXCLUDED.total_aktual,
                        total_persentase = EXCLUDED.total_persentase
                """
                cur.execute(
                    upsert_sql,
                    (tanggal_produksi, total_target, total_aktual, total_persentase),
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

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


if __name__ == "__main__":
    app.run(debug=True)
