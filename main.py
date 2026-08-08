from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import openpyxl
import re, io, tempfile
from datetime import datetime

app = FastAPI()

# ── Helpers ──────────────────────────────────────────────────────────────────

def read_time_summary(xl_bytes):
    df = pd.read_excel(io.BytesIO(xl_bytes), sheet_name="Time Summary", header=None)
    df.columns = [
        "_","Pozo","Equipo","EventCode","Desde","Hasta","Horas","Fase","Tarea",
        "Actividad","Codigo","NPT","Detalle_NPT","Evidencia","CausaRaiz","Prof",
        "Prof_NPT","Operacion","Descripcion","Equipment","FailedEquip","ServiceCompany"
    ]
    df = df.iloc[1:].copy()
    df["Desde"] = pd.to_datetime(df["Desde"], errors="coerce")
    df["Hasta"] = pd.to_datetime(df["Hasta"], errors="coerce")
    df["Horas"] = pd.to_numeric(df["Horas"], errors="coerce")
    return df

def read_sbp(xl_bytes):
    xl = pd.ExcelFile(io.BytesIO(xl_bytes))
    sbp_name = next((s for s in xl.sheet_names if s.upper() in ("SBP","SPB")), None)
    if not sbp_name:
        return None
    df = pd.read_excel(io.BytesIO(xl_bytes), sheet_name=sbp_name, header=None)
    df.columns = [
        "_","Time","MD","TVD","FlowIn","CoriolisRate","ReturnFlow","SBP_psi",
        "RCD_psi","CoriolicsMW","ROP","RPM","ECD","ChokeA","ChokeB","Estado","Well"
    ]
    df = df.iloc[1:].copy()
    df["Time"] = pd.to_datetime(df["Time"], errors="coerce")
    df["RPM"] = pd.to_numeric(df["RPM"].astype(str).str.replace(",","."), errors="coerce").fillna(0)
    return df

def read_bearing_serial(xl_bytes):
    xl = pd.ExcelFile(io.BytesIO(xl_bytes))
    bs_name = next((s for s in xl.sheet_names
                    if any(w in s.lower() for w in ["bearing","serial","hoja"])), None)
    if not bs_name:
        return None
    df = pd.read_excel(io.BytesIO(xl_bytes), sheet_name=bs_name, header=None)
    df = df.iloc[1:].copy()
    df.columns = ["_","Well","InstallDate","PullDate","RunNo","Serial"]
    df["InstallDate"] = pd.to_datetime(df["InstallDate"], errors="coerce")
    df["PullDate"]    = pd.to_datetime(df["PullDate"],    errors="coerce")
    df["RunNo"]       = pd.to_numeric(df["RunNo"], errors="coerce")
    return df.dropna(subset=["InstallDate","Serial"])

def get_well_name(ts):
    vals = ts["Pozo"].dropna().unique()
    return str(vals[0]) if len(vals) > 0 else "Unknown"

# ── Find MPD rig-up end ──────────────────────────────────────────────────────

def find_service_start(ts):
    """Return (start_timestamp, rig_up_hours).
    Looks for BOPSUR task in any phase (PROD1, INT2, etc.)."""
    bopsur = ts[ts["Tarea"] == "BOPSUR"].sort_values("Desde")
    if not bopsur.empty:
        # Last BOPSUR row = end of rig-up
        end = bopsur.iloc[-1]["Hasta"]
        total = round(float(bopsur["Horas"].sum()), 2)
        return end, total

    # Fallback: look for campana de viaje placement in any phase
    mask = ts["Operacion"].str.contains("monta.*campana|coloca.*campana|campana de viaje", case=False, na=False)
    found = ts[mask].sort_values("Desde")
    if not found.empty:
        return found.iloc[0]["Hasta"], 0.0

    return ts.iloc[0]["Desde"], 0.0

# ── NPT ──────────────────────────────────────────────────────────────────────

def find_npt_mpd(ts, start_date):
    """Search NPT across ALL phases from start_date."""
    period = ts[ts["Desde"] >= start_date].copy()

    combos = {
        "WREP-RMPD":        {"hs": 0.0, "detail": ""},
        "WSER-SMPD":        {"hs": 0.0, "detail": ""},
        "ROT-SFAL-RTME-MPD":{"hs": 0.0, "detail": ""},
        "ROT-LOSS-CECD":    {"hs": 0.0, "detail": ""},
        "ROT-WCON-CECD":    {"hs": 0.0, "detail": ""},
    }

    npt_rows = period[period["NPT"].notna()].copy()
    if not npt_rows.empty:
        npt_rows["combo"] = (
            npt_rows["NPT"].astype(str).str.strip() + "-" +
            npt_rows["Detalle_NPT"].astype(str).str.strip()
        )
        for combo in combos:
            matched = npt_rows[npt_rows["combo"].str.upper() == combo.upper()]
            if not matched.empty:
                combos[combo]["hs"] = round(float(matched["Horas"].sum()), 2)
                combos[combo]["detail"] = " | ".join(
                    matched["Operacion"].dropna().astype(str).str[:200].tolist()
                )

    # Also check Evidence column for CECDMPD
    for col in ["Evidencia", "CausaRaiz"]:
        if col not in period.columns:
            continue
        ev_rows = period[period[col].astype(str).str.contains("CECDMPD", case=False, na=False)]
        if ev_rows.empty:
            continue
        npt_ev = ev_rows[ev_rows["NPT"].notna()].copy()
        if npt_ev.empty:
            continue
        npt_ev["combo"] = (
            npt_ev["NPT"].astype(str).str.strip() + "-" +
            npt_ev["Detalle_NPT"].astype(str).str.strip()
        )
        for combo in combos:
            matched = npt_ev[npt_ev["combo"].str.upper() == combo.upper()]
            if not matched.empty:
                combos[combo]["hs"] = round(float(matched["Horas"].sum()), 2)
                combos[combo]["detail"] = " | ".join(
                    matched["Operacion"].dropna().astype(str).str[:200].tolist()
                )

    return combos

# ── Bearings ─────────────────────────────────────────────────────────────────

def find_bearings(ts, sbp, serial_df, start_date):
    """Detect bearing install/remove events across ALL phases from start_date."""
    period = ts[ts["Desde"] >= start_date].copy().sort_values("Desde")

    mask = period["Operacion"].str.contains("BEARING|BERING", case=False, na=False)
    b_rows = period[mask].reset_index(drop=True)

    # Classify each row as install or remove
    installs = []
    removes  = []
    for _, row in b_rows.iterrows():
        op = str(row["Operacion"]).upper()
        is_install = any(w in op for w in ["INSTALA","COLOCA BEARING","MONTA BEARING","RETIRA CAMPANA"])
        is_remove  = any(w in op for w in ["RETIRA BEARING","DESMONTA BEARING","SACA BEARING",
                                            "CAMBIA BEARING","REEMPLAZA BEARING"])
        if is_install and not is_remove:
            installs.append(row)
        elif is_remove and not is_install:
            removes.append(row)

    # Pair installs with removes chronologically
    pairs = []
    used_removes = set()
    for inst in installs:
        t_in = inst["Desde"]
        best = None
        for i, rem in enumerate(removes):
            if i in used_removes:
                continue
            if rem["Hasta"] > t_in:
                best = (i, rem)
                break
        if best:
            used_removes.add(best[0])
            pairs.append({"in_row": inst, "out_row": best[1]})

    # Build results
    results = []
    for idx, pair in enumerate(pairs):
        t_in  = pair["in_row"]["Desde"]
        t_out = pair["out_row"]["Hasta"]

        seg = period[(period["Desde"] >= t_in) & (period["Hasta"] <= t_out)]

        # Metros perforados
        drill = seg[seg["Actividad"] == "DRL"]
        total_drill = 0.0
        for _, row in drill.iterrows():
            m = re.search(r"AVANCE:\s*(\d+(?:[.,]\d+)?)\s*m", str(row["Operacion"]), re.IGNORECASE)
            if m:
                total_drill += float(m.group(1).replace(",", "."))

        # Tiempo de servicio
        svc = round((t_out - t_in).total_seconds() / 3600, 2)

        # Horas de rotación
        rot_hs = 0
        if sbp is not None:
            sbp_seg = sbp[(sbp["Time"] >= t_in) & (sbp["Time"] <= t_out)]
            rot_hs  = int((sbp_seg["RPM"] > 0).sum())

        # Motivo de cambio
        motivo = "Campana de Viaje"
        op_out = str(pair["out_row"]["Operacion"]).upper()
        if any(w in op_out for w in ["CEMENT","CEMEN"]):
            motivo = "Cementación"
        elif any(w in op_out for w in ["DAÑ","FALLA","LEAK","FUGA","ROTURA"]):
            motivo = "Daño de Bearing"
        elif "ELASTOM" in op_out:
            motivo = "Daño de Elastomero"
        elif "HORAS" in op_out:
            motivo = "Horas Acumuladas"

        # Tiempo de cambio (exclude safety meetings)
        tcambio = 0.0
        remove_win = period[
            (period["Desde"] >= t_out - pd.Timedelta(hours=1)) &
            (period["Hasta"]  <= t_out + pd.Timedelta(hours=0.5)) &
            period["Operacion"].str.contains("BEARING|BERING|CAMPANA", case=False, na=False)
        ]
        for _, row in remove_win.iterrows():
            op = str(row["Operacion"]).upper()
            if not any(w in op for w in ["REUNI","HSE","SEGUR","CHARLA"]):
                tcambio += float(row["Horas"]) if pd.notna(row["Horas"]) else 0.0
        tcambio = round(tcambio, 2)

        # Con/Sin presión
        presion = "SIN"
        for _, row in remove_win.iterrows():
            op = str(row["Operacion"]).upper()
            if any(w in op for w in ["PSI","PRESION","BOP CERR","ANULAR CERR","STRIPPING"]):
                presion = "CON"
                break

        # Serial
        serial = ""
        if serial_df is not None:
            # Match by install date (±2 hours)
            matched = serial_df[
                (serial_df["InstallDate"] >= t_in - pd.Timedelta(hours=2)) &
                (serial_df["InstallDate"] <= t_in + pd.Timedelta(hours=2))
            ]
            if not matched.empty:
                serial = str(matched.iloc[0]["Serial"])
            else:
                serial = "No cargado en OW"

        results.append({
            "run":      idx + 1,
            "serial":   serial,
            "t_in":     t_in,
            "t_out":    t_out,
            "drill_m":  round(total_drill),
            "svc_hs":   svc,
            "motivo":   motivo,
            "tcambio":  tcambio,
            "rot_hs":   rot_hs,
            "presion":  presion,
        })

    return results

# ── Fill template ─────────────────────────────────────────────────────────────

def fill_template(template_bytes, well_name, inicio, fin, total_hs, npt_combos, bearings):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    if "Horas Operativas del Servicio" in wb.sheetnames:
        ws = wb["Horas Operativas del Servicio"]
        ws["B3"] = well_name
        ws["C3"] = None
        ws["D3"] = inicio
        ws["D3"].number_format = "DD/MM/YYYY HH:MM"
        ws["E3"] = fin
        ws["E3"].number_format = "DD/MM/YYYY HH:MM"
        ws["F3"] = total_hs

    if "NPT" in wb.sheetnames:
        ws = wb["NPT"]
        total = round(sum(v["hs"] for v in npt_combos.values()), 2)
        for i, (combo, data) in enumerate(npt_combos.items(), 3):
            ws[f"B{i}"] = well_name
            ws[f"C{i}"] = None
            ws[f"D{i}"] = combo
            ws[f"E{i}"] = data["hs"]
            ws[f"F{i}"] = total
            if data["detail"]:
                ws[f"G{i}"] = data["detail"]

    if "Bearing Resumen" in wb.sheetnames:
        ws = wb["Bearing Resumen"]
        for i, b in enumerate(bearings, 3):
            ws[f"B{i}"] = well_name
            ws[f"C{i}"] = None
            ws[f"D{i}"] = b["run"]
            ws[f"E{i}"] = b["serial"]
            ws[f"F{i}"] = b["drill_m"]
            ws[f"G{i}"] = "REVISAR"
            ws[f"H{i}"] = b["svc_hs"]
            ws[f"I{i}"] = b["motivo"]
            ws[f"J{i}"] = b["tcambio"]
            ws[f"K{i}"] = b["rot_hs"]
            ws[f"L{i}"] = b["presion"]

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index():
    with open("templates/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/process")
async def process(
    data_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    try:
        data_bytes     = await data_file.read()
        template_bytes = await template_file.read()

        ts        = read_time_summary(data_bytes)
        sbp       = read_sbp(data_bytes)
        serial_df = read_bearing_serial(data_bytes)
        well_name = get_well_name(ts)

        inicio, _ = find_service_start(ts)
        fin        = ts.iloc[-1]["Hasta"]
        total_hs   = round((fin - inicio).total_seconds() / 3600, 2)

        npt_combos = find_npt_mpd(ts, inicio)
        bearings   = find_bearings(ts, sbp, serial_df, inicio)

        result_bytes = fill_template(
            template_bytes, well_name, inicio, fin, total_hs, npt_combos, bearings
        )

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(result_bytes)
        tmp.close()

        filename = f"{well_name.replace('(','').replace(')','').replace(' ','_')}_resultado.xlsx"
        return FileResponse(
            tmp.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
