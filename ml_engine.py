# ml_engine.py
# -*- coding: utf-8 -*-
import pickle
import hashlib
from pathlib import Path
import numpy as np
import pandas as pd

KONSTRUK_LIST = [
    "Self-Efficacy", "Keterlibatan Ortu", "Harapan Ortu", "Dukungan Sekolah",
    "Motivasi", "Kecemasan", "Fasilitas", "Kemalasan",
]
NAMA_ASPEK_MAP = {
    "Self-Efficacy":     "Self-Efficacy Akademik",
    "Keterlibatan Ortu": "Keterlibatan Orang Tua",
    "Harapan Ortu":      "Harapan Orang Tua",
    "Dukungan Sekolah":  "Dukungan Sekolah",
    "Motivasi":          "Motivasi Belajar",
    "Kecemasan":         "Kecemasan Akademik",
    "Fasilitas":         "Fasilitas Sekolah",
    "Kemalasan":         "Kemalasan Belajar",
}

MODEL_DIR   = Path("models")
MODEL_DIR.mkdir(exist_ok=True)
MODEL_PATH  = MODEL_DIR / "rf_nilai.pkl"
META_PATH   = MODEL_DIR / "rf_nilai_meta.pkl"

MIN_TRAIN_SAMPLES = 30

def _prepare(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=KONSTRUK_LIST + ["Nilai Akademik"])
    cols = KONSTRUK_LIST + ["Nilai Akademik"]
    # Pastikan semua kolom ada
    for c in cols:
        if c not in df.columns:
            df[c] = 0.0
    d = df[cols].copy()
    mask = (d[KONSTRUK_LIST] > 0).all(axis=1) & (d["Nilai Akademik"] > 0)
    return d[mask].dropna().reset_index(drop=True)

def _hash(df: pd.DataFrame) -> str:
    return hashlib.md5(pd.util.hash_pandas_object(df, index=True).values.tobytes()).hexdigest()

def train(df: pd.DataFrame):
    from sklearn.ensemble import RandomForestRegressor
    from sklearn.model_selection import cross_val_predict, KFold
    from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error

    d = _prepare(df)
    if len(d) < MIN_TRAIN_SAMPLES:
        return None

    X = d[KONSTRUK_LIST].values
    y = d["Nilai Akademik"].values

    model = RandomForestRegressor(n_estimators=400, max_depth=6, min_samples_leaf=2, random_state=42, n_jobs=-1).fit(X, y)

    n_splits = min(5, len(d))
    if n_splits >= 3:
        kf = KFold(n_splits=n_splits, shuffle=True, random_state=42)
        y_cv = cross_val_predict(model, X, y, cv=kf)
        r2   = r2_score(y, y_cv)
        mae  = mean_absolute_error(y, y_cv)
        rmse = float(np.sqrt(mean_squared_error(y, y_cv)))
    else:
        r2 = mae = rmse = None

    importance = {k: float(v) for k, v in zip(KONSTRUK_LIST, model.feature_importances_)}

    meta = {
        "n_train":    int(len(d)),
        "r2_cv":      float(r2) if r2 is not None else None,
        "mae_cv":     float(mae) if mae is not None else None,
        "rmse_cv":    rmse,
        "importance": importance,
        "data_hash":  _hash(d),
        "trained_at": pd.Timestamp.now().isoformat(timespec="seconds"),
    }

    with open(MODEL_PATH, "wb") as f: pickle.dump(model, f)
    with open(META_PATH, "wb") as f: pickle.dump(meta, f)

    return {"model": model, "meta": meta}

def load_cached(df: pd.DataFrame):
    if not (MODEL_PATH.exists() and META_PATH.exists()):
        return None
    try:
        with open(MODEL_PATH, "rb") as f: model = pickle.load(f)
        with open(META_PATH, "rb") as f: meta = pickle.load(f)
    except Exception:
        return None

    d = _prepare(df)
    if len(d) < MIN_TRAIN_SAMPLES:
        return None
    if meta.get("data_hash") != _hash(d):
        return None
    return {"model": model, "meta": meta}

def get_model(df: pd.DataFrame, force_retrain: bool = False):
    if not force_retrain:
        cached = load_cached(df)
        if cached:
            return cached
    return train(df)

def predict(model_bundle, profil: dict) -> float | None:
    if model_bundle is None:
        return None
    try:
        x = np.array([[float(profil[k]) for k in KONSTRUK_LIST]])
    except KeyError as e:
        raise ValueError(f"Profil kurang kunci: {e}")
    pred = float(model_bundle["model"].predict(x)[0])
    return float(np.clip(pred, 0, 100))

def feature_importance(model_bundle) -> dict | None:
    if model_bundle is None:
        return None
    imp = model_bundle["meta"]["importance"]
    return {NAMA_ASPEK_MAP[k]: v for k, v in imp.items()}

def mockup_predict(profil: dict) -> float:
    base = 75.0
    score = (
        base
        + profil["Self-Efficacy"]     * 1.5
        + profil["Keterlibatan Ortu"] * 1.2
        + profil["Motivasi"]          * 1.0
        - profil["Kecemasan"]         * 0.8
        - profil["Kemalasan"]         * 1.5
        + profil["Dukungan Sekolah"]  * 0.5
    )
    return float(np.clip(score, 60.0, 100.0))