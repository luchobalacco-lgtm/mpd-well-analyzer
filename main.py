from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import re
import io
import os
import tempfile
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
    bs_name = next((s for s in xl.sheet_names if "bearing" in s.lower() or "serial" in s.lower() or "hoja" in s.lower()), None)
    if not bs_name:
        return None
    df = pd.read_excel(io.BytesIO(xl_bytes), sheet_name=bs_name, header=None)
    df = df.iloc[1:].copy()
    df.columns = ["_","Well","InstallDate","PullDate","RunNo","Serial"]
    df["InstallDate"] = pd.to_datetime(df["InstallDate"], errors="coerce")
    df["PullDate"] = pd.to_datetime(df["PullDate"], errors="coerce")
    return df

def get_well_name(ts):
    vals = ts["Pozo"].dropna().unique()
    return str(vals[0]) if len(vals) > 0 else "Unknown"

def get_prod1(ts):
    return ts[ts["Fase"] == "PROD1"].copy()

def find_mpd_test_end(prod1):
    mask = prod1["Operacion"].str.contains(
        "campana de viaje|coloca.*campana|monta.*campana", case=False, na=False
    )
    found = prod1[mask]
    if not found.empty:
        row = found.iloc[0]
        # Also check for RCD test nearby
        test_mask = prod1["Operacion"].str.contains("prueba.*RCD|RCD.*prueba|4000 PSI|4.000 PSI", case=False, na=False)
        test_rows = prod1[test_mask]
        if not test_rows.empty:
            return test_rows.iloc[-1]["Hasta"]
        return row["Hasta"]
    return prod1.iloc[0]["Desde"]

def find_rig_up_hours(prod1):
    bopsur = prod1[prod1["Tarea"] == "BOPSUR"]
    if bopsur.empty:
        return 0.0, None
    return round(float(bopsur["Horas"].sum()), 2), bopsur["Hasta"].max()

def find_npt_mpd(prod1):
    npt_combos = {
        "WREP-RMPD": 0.0,
        "WSER-SMPD": 0.0,
        "ROT-SFAL-RTME-MPD": 0.0,
        "ROT-LOSS-CECD": 0.0,
        "ROT-WCON-CECD": 0.0,
    }
    npt_details = {k: "" for k in npt_combos}

    npt_rows = prod1[prod1["NPT"].notna()].copy()
    if not npt_rows.empty:
        npt_rows["combo"] = (
            npt_rows["NPT"].astype(str).str.strip() + "-" +
            npt_rows["Detalle_NPT"].astype(str).str.strip()
        )
        for combo in npt_combos:
            matched = npt_rows[npt_rows["combo"].str.upper() == combo.upper()]
            if not matched.empty:
                npt_combos[combo] = round(float(matched["Horas"].sum()), 2)
                npt_details[combo] = " | ".join(
                    matched["Operacion"].dropna().astype(str).str[:150].tolist()
                )

    # Also check Evidence column for CECDMPD
    for col in ["Evidencia", "CausaRaiz"]:
        if col in prod1.columns:
            mpd_rows = prod1[prod1[col].astype(str).str.contains("CECDMPD", case=False, na=False)]
            if not mpd_rows.empty:
                mpd_npt = mpd_rows[mpd_rows["NPT"].notna()].copy()
                if not mpd_npt.empty:
                    mpd_npt["combo"] = (
                        mpd_npt["NPT"].astype(str).str.strip() + "-" +
                        mpd_npt["Detalle_NPT"].astype(str).str.strip()
                    )
                    for combo in npt_combos:
                        matched = mpd_npt[mpd_npt["combo"].str.upper() == combo.upper()]
                        if not matched.empty:
                            npt_combos[combo] = round(float(matched["Horas"].sum()), 2)
                            npt_details[combo] = " | ".join(
                                matched["Operacion"].dropna().astype(str).str[:150].tolist()
                            )

    return npt_combos, npt_details

def find_bearings(prod1, sbp, serial_df, start_date):
    prod1_period = prod1[prod1["Desde"] >= start_date].copy() if start_date else prod1.copy()

    mask = prod1_period["Operacion"].str.contains("BEARING|BERING", case=False, na=False)
    bearing_rows = prod1_period[mask].sort_values("Desde")

    events_install = []
    events_remove = []

    for _, row in bearing_rows.iterrows():
        op = str(row["Operacion"]).upper()
        if any(w in op for w in ["INSTALA", "COLOCA", "MONTA", "RETIRA CAMPANA"]):
            events_install.append(row["Desde"])
        elif any(w in op for w in ["RETIRA", "DESMONTA", "SACA", "CAMBIA"]):
            events_remove.append(row["Desde"])

    # Build periods pairing installs with removes
    bearings = []
    for i, t_in in enumerate(events_install):
        t_out = None
        for t_rem in events_remove:
            if t_rem > t_in:
                t_out = t_rem
                break
        if t_out is None and i < len(events_remove):
            t_out = events_remove[i]
        bearings.append({"in": t_in, "out": t_out})

    # For each bearing, calculate metrics
    results = []
    for idx, b in enumerate(bearings):
        t_in = b["in"]
        t_out = b["out"]

        if t_out is None:
            continue

        period = prod1_period[
            (prod1_period["Desde"] >= t_in) & (prod1_period["Hasta"] <= t_out)
        ]

        # Metros perforados via AVANCE
        drill = period[period["Actividad"] == "DRL"]
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
            sbp_period = sbp[(sbp["Time"] >= t_in) & (sbp["Time"] <= t_out)]
            rot_hs = int((sbp_period["RPM"] > 0).sum())

        # Motivo de cambio (from the row that triggers removal)
        motivo = "Campana de Viaje"
        removal_rows = prod1_period[
            (prod1_period["Desde"] >= t_out - pd.Timedelta(hours=2)) &
            (prod1_period["Hasta"] <= t_out + pd.Timedelta(hours=1))
        ]
        for _, row in removal_rows.iterrows():
            op = str(row["Operacion"]).upper()
            if "CEMENT" in op or "CEMEN" in op:
                motivo = "Cementación"
                break
            elif "DAÑ" in op or "FALLA" in op or "LEAK" in op or "FUGA" in op:
                motivo = "Daño de Bearing"
                break
            elif "ELASTOM" in op:
                motivo = "Daño de Elastomero"
                break

        # Tiempo de cambio (excluding safety meetings)
        tcambio = 0.0
        change_rows = prod1_period[
            (prod1_period["Desde"] >= t_out - pd.Timedelta(hours=2)) &
            (prod1_period["Hasta"] <= t_out + pd.Timedelta(hours=1)) &
            prod1_period["Operacion"].str.contains("BEARING|BERING|CAMPANA", case=False, na=False)
        ]
        for _, row in change_rows.iterrows():
            op = str(row["Operacion"]).upper()
            if "REUNI" not in op and "HSE" not in op and "SEGUR" not in op:
                tcambio += float(row["Horas"]) if pd.notna(row["Horas"]) else 0.0
        tcambio = round(tcambio, 2)

        # Con/Sin presión
        presion = "SIN"
        for _, row in removal_rows.iterrows():
            op = str(row["Operacion"]).upper()
            if any(w in op for w in ["PSI", "PRESION", "BOP CERR", "ANULAR CERR", "STRIPPING"]):
                presion = "CON"
                break

        # Serial number
        serial = ""
        if serial_df is not None:
            matched = serial_df[
                (serial_df["InstallDate"] >= t_in - pd.Timedelta(hours=2)) &
                (serial_df["InstallDate"] <= t_in + pd.Timedelta(hours=2))
            ]
            if not matched.empty:
                serial = str(matched.iloc[0]["Serial"])
            else:
                serial = "No cargado en OW"

        results.append({
            "run": idx + 1,
            "serial": serial,
            "t_in": t_in,
            "t_out": t_out,
            "drill_m": round(total_drill),
            "strip_m": 0,  # requires manual review - set 0 as placeholder
            "svc_hs": svc,
            "motivo": motivo,
            "tcambio": tcambio,
            "rot_hs": rot_hs,
            "presion": presion,
        })

    return results

def fill_template(template_bytes, well_name, inicio, fin, total_hs, npt_combos, npt_details, bearings):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    # ── Horas Operativas ──
    if "Horas Operativas del Servicio" in wb.sheetnames:
        ws = wb["Horas Operativas del Servicio"]
        ws["B3"] = well_name
        ws["C3"] = None
        ws["D3"] = inicio
        ws["D3"].number_format = "DD/MM/YYYY HH:MM"
        ws["E3"] = fin
        ws["E3"].number_format = "DD/MM/YYYY HH:MM"
        ws["F3"] = total_hs

    # ── NPT ──
    if "NPT" in wb.sheetnames:
        ws = wb["NPT"]
        combos = list(npt_combos.keys())
        total = round(sum(npt_combos.values()), 2)
        for i, combo in enumerate(combos, 3):
            ws[f"B{i}"] = well_name
            ws[f"C{i}"] = None
            ws[f"D{i}"] = combo
            ws[f"E{i}"] = npt_combos[combo]
            ws[f"F{i}"] = total
            if npt_details.get(combo):
                ws[f"G{i}"] = npt_details[combo]

    # ── Bearing Resumen ──
    if "Bearing Resumen" in wb.sheetnames:
        ws = wb["Bearing Resumen"]
        for i, b in enumerate(bearings, 3):
            ws[f"B{i}"] = well_name
            ws[f"C{i}"] = None
            ws[f"D{i}"] = b["run"]
            ws[f"E{i}"] = b["serial"]
            ws[f"F{i}"] = b["drill_m"]
            ws[f"G{i}"] = b["strip_m"] if b["strip_m"] > 0 else "REVISAR"
            ws[f"H{i}"] = b["svc_hs"]
            ws[f"I{i}"] = b["motivo"]
            ws[f"J{i}"] = b["tcambio"]
            ws[f"K{i}"] = b["rot_hs"]
            ws[f"L{i}"] = b["presion"]

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


# ── Routes ───────────────────────────────────────────────────────────────────

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
        data_bytes = await data_file.read()
        template_bytes = await template_file.read()

        ts = read_time_summary(data_bytes)
        sbp = read_sbp(data_bytes)
        serial_df = read_bearing_serial(data_bytes)
        prod1 = get_prod1(ts)
        well_name = get_well_name(ts)

        # Dates
        mpd_end = find_mpd_test_end(prod1)
        _, bopsur_end = find_rig_up_hours(prod1)
        inicio = bopsur_end or mpd_end
        fin = prod1.iloc[-1]["Hasta"]
        total_hs = round((fin - inicio).total_seconds() / 3600, 2)

        # NPT
        npt_combos, npt_details = find_npt_mpd(prod1)

        # Bearings
        bearings = find_bearings(prod1, sbp, serial_df, inicio)

        # Fill template
        result_bytes = fill_template(
            template_bytes, well_name, inicio, fin, total_hs,
            npt_combos, npt_details, bearings
        )

        # Save to temp file and return
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(result_bytes)
        tmp.close()

        filename = f"{well_name.replace('(','').replace(')','').replace(' ','_')}_resultado.xlsx"
        return FileResponse(
            tmp.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
            background=None
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
