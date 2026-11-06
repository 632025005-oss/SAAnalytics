# causal_engine.py
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd

KONSTRUK_LIST = [
    "Self-Efficacy", "Keterlibatan Ortu", "Harapan Ortu", "Dukungan Sekolah",
    "Motivasi", "Kecemasan", "Fasilitas", "Kemalasan",
]

NAMA_ASPEK_MAP = {
    "Self-Efficacy":       "Self-Efficacy Akademik",
    "Keterlibatan Ortu":   "Keterlibatan Orang Tua",
    "Harapan Ortu":        "Harapan Orang Tua",
    "Dukungan Sekolah":    "Dukungan Sekolah",
    "Motivasi":            "Motivasi Belajar",
    "Kecemasan":           "Kecemasan Akademik",
    "Fasilitas":           "Fasilitas Sekolah",
    "Kemalasan":           "Kemalasan Belajar",
}

FALLBACK_PENGARUH = {
    "Self-Efficacy Akademik": 2.1995, "Keterlibatan Orang Tua": 1.7559,
    "Harapan Orang Tua": -0.6987, "Dukungan Sekolah": -1.0351,
    "Motivasi Belajar": -1.9093, "Kecemasan Akademik": 0.7197,
    "Kemalasan Belajar": 0.7320, "Fasilitas Sekolah": 0.6182,
}
FALLBACK_KEPENTINGAN = {
    "Self-Efficacy Akademik": 0.8197, "Keterlibatan Orang Tua": 0.6717,
    "Harapan Orang Tua": 0.5766, "Dukungan Sekolah": 0.4325,
    "Motivasi Belajar": 0.3601, "Kecemasan Akademik": 0.1553,
    "Kemalasan Belajar": 0.0846, "Fasilitas Sekolah": 0.0654,
}

MIN_SAMPLES = 30

def _prepare_df(df: pd.DataFrame) -> pd.DataFrame:
    cols = KONSTRUK_LIST + ["Nilai Akademik"]
    for c in cols:
        if c not in df.columns:
            df[c] = 0.0
    d = df[cols].copy()
    mask = (d[KONSTRUK_LIST] > 0).all(axis=1) & (d["Nilai Akademik"] > 0)
    return d[mask].dropna().reset_index(drop=True)

def estimasi_ate(df: pd.DataFrame) -> dict | None:
    from sklearn.linear_model import LinearRegression
    d = _prepare_df(df)
    if len(d) < MIN_SAMPLES:
        return None
    X = d[KONSTRUK_LIST].values
    y = d["Nilai Akademik"].values
    model = LinearRegression().fit(X, y)
    return {NAMA_ASPEK_MAP[k]: float(c) for k, c in zip(KONSTRUK_LIST, model.coef_)}

def estimasi_kepentingan(df: pd.DataFrame) -> dict | None:
    from sklearn.ensemble import RandomForestRegressor
    import shap
    d = _prepare_df(df)
    if len(d) < MIN_SAMPLES:
        return None
    X = d[KONSTRUK_LIST].values
    y = d["Nilai Akademik"].values
    rf = RandomForestRegressor(n_estimators=300, max_depth=4, min_samples_leaf=3, random_state=42, n_jobs=-1).fit(X, y)
    sv = shap.TreeExplainer(rf).shap_values(X)
    mean_abs = np.abs(sv).mean(axis=0)
    total = mean_abs.sum()
    if total > 0:
        mean_abs = mean_abs / total
    return {NAMA_ASPEK_MAP[k]: float(v) for k, v in zip(KONSTRUK_LIST, mean_abs)}

def hitung_semua(df: pd.DataFrame) -> dict:
    try:
        ate = estimasi_ate(df)
        fi = estimasi_kepentingan(df)
    except Exception as e:
        return {"pengaruh": FALLBACK_PENGARUH.copy(), "kepentingan": FALLBACK_KEPENTINGAN.copy(), "sumber": "fallback", "n_sampel": 0, "error": str(e)}

    if ate is None or fi is None:
        return {"pengaruh": FALLBACK_PENGARUH.copy(), "kepentingan": FALLBACK_KEPENTINGAN.copy(), "sumber": "fallback", "n_sampel": len(_prepare_df(df))}

    return {"pengaruh": ate, "kepentingan": fi, "sumber": "data_aktual", "n_sampel": len(_prepare_df(df))}