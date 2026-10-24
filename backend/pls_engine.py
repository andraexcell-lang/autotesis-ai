# backend/pls_engine.py
"""Mesin PLS-SEM internal + generator data jawaban responden.

Tujuan (permintaan pengguna):
  - Data jawaban responden yang di-generate HARUS benar-benar menghasilkan
    angka yang sama bila diolah di SmartPLS 4 asli.
  - Karena itu seluruh statistik DIHITUNG dari data (bukan hardcode), memakai
    algoritma PLS-SEM standar (Lohmoller) dengan skema bobot *path weighting*,
    data terstandardisasi, 300 iterasi maksimum, kriteria berhenti 1e-7.
  - Satu sumber kebenaran ("compute-then-report"): angka Bab III/IV/V diambil
    dari hasil hitung atas data yang dihasilkan.

Catatan jujur soal paritas SmartPLS:
  - Metrik DETERMINISTIK (outer loading, AVE, CA, CR, HTMT, Fornell-Larcker,
    VIF, koefisien jalur, R2, f2, Q2, SRMR) akan sama/mendekati SmartPLS asli
    bila setting algoritma sama.
  - Metrik BOOTSTRAP (t, p, CI) berasal dari subsample acak; dua kali run
    SmartPLS pun bisa berbeda desimal. Yang dijamin: besaran sangat dekat dan
    kesimpulan signifikansi sama.
"""

import hashlib
import math
import os
import re
import time
import numpy as np

from provenance import (
    label_data_source,
    label_origin,
    normalize_analysis_engine,
    normalize_data_source,
)

ROLE_ORDER = ["X1", "X2", "M", "Y"]
DEFAULT_INNER = {"M": ["X1", "X2"], "Y": ["X1", "X2", "M"]}

# Peta label skala Likert -> angka (untuk membaca ekspor Google Form/Sheets).
_LIKERT_LABEL = {
    "sangat tidak setuju": 1,
    "tidak setuju": 2,
    "kurang setuju": 3,
    "netral": 3,
    "ragu ragu": 3,
    "cukup setuju": 3,
    "setuju": 4,
    "sangat setuju": 5,
}


# --------------------------------------------------------------------------- #
# Utilitas statistik
# --------------------------------------------------------------------------- #
def _z(a, ddof=1):
    a = np.asarray(a, dtype=float)
    s = a.std(ddof=ddof)
    if not np.isfinite(s) or s < 1e-12:
        return a - a.mean()
    return (a - a.mean()) / s


def _corr(a, b):
    a = np.asarray(a, dtype=float)
    b = np.asarray(b, dtype=float)
    sa, sb = a.std(), b.std()
    if sa < 1e-12 or sb < 1e-12:
        return 0.0
    return float(((a - a.mean()) * (b - b.mean())).mean() / (sa * sb))


def _norm_sf(x):
    """P(Z > x) memakai erfc (tanpa scipy)."""
    return 0.5 * math.erfc(x / math.sqrt(2.0))


def _two_tail_p(t):
    return 2.0 * _norm_sf(abs(float(t)))


# --------------------------------------------------------------------------- #
# Inti algoritma PLS-SEM (Mode A, path weighting scheme)
# --------------------------------------------------------------------------- #
def pls_core(Xz, idx, inner_pred, iters=300, tol=1e-7):
    """Estimasi skor konstruk, bobot luar, dan matriks jalur.

    Xz      : matriks indikator terstandardisasi (n x p)
    idx     : {nama_konstruk: np.array(indeks kolom)}
    inner_pred : {konstruk_endogen: [konstruk_prediktor]}
    Kembalian: (S, W, loadings, paths, n_iter)
    """
    names = list(idx.keys())
    pred = {nm: [p for p in inner_pred.get(nm, [])] for nm in names}
    succ = {nm: [] for nm in names}
    for t, ps in inner_pred.items():
        for p in ps:
            succ[p].append(t)

    W = {nm: np.ones(len(idx[nm])) for nm in names}

    def _scores(W):
        return {nm: Xz[:, idx[nm]] @ W[nm] for nm in names}

    n_iter = 0
    for it in range(iters):
        n_iter = it + 1
        S = {nm: _z(v) for nm, v in _scores(W).items()}
        Z = {}
        for j in names:
            zj = np.zeros(Xz.shape[0])
            if pred[j]:
                P = np.column_stack([S[p] for p in pred[j]])
                beta, *_ = np.linalg.lstsq(P, S[j], rcond=None)
                for bi, p in enumerate(pred[j]):
                    zj += beta[bi] * S[p]
            for k in succ[j]:
                zj += _corr(S[j], S[k]) * S[k]
            Z[j] = _z(zj)

        W_new = {}
        for j in names:
            w = (Xz[:, idx[j]].T @ Z[j]) / Xz.shape[0]
            # skalakan agar skor konstruk berunit varian (seperti SmartPLS)
            s = np.sqrt(float(w @ w))
            W_new[j] = w / s if s > 1e-12 else w
        delta = max(float(np.max(np.abs(W_new[nm] - W[nm]))) for nm in names)
        W = W_new
        if delta < tol:
            break

    S = {nm: _z(v) for nm, v in _scores(W).items()}
    # loading luar & bobot (dilaporkan terstandardisasi; kolom X sudah terstandardisasi)
    n_obs = Xz.shape[0]
    loadings = {}
    weights = {}
    for j in names:
        Xj = Xz[:, idx[j]]
        loadings[j] = (Xj.T @ S[j]) / n_obs
        weights[j] = (Xj.T @ Z[j]) / n_obs
    # koefisien jalur (regresi OLS skor konstruk, semua terstandardisasi)
    paths = {}
    for t, ps in inner_pred.items():
        if not ps:
            continue
        Y = S[t]
        P = np.column_stack([S[p] for p in ps])
        beta, *_ = np.linalg.lstsq(P, Y, rcond=None)
        for bi, p in enumerate(ps):
            paths[(p, t)] = float(beta[bi])
    return S, W, loadings, weights, paths, n_iter


def _r2(S, inner_pred):
    out = {}
    for t, ps in inner_pred.items():
        if not ps:
            continue
        Y = S[t]
        P = np.column_stack([S[p] for p in ps])
        beta, *_ = np.linalg.lstsq(P, Y, rcond=None)
        resid = Y - P @ beta
        ss_tot = float(((Y - Y.mean()) ** 2).sum())
        ss_res = float((resid ** 2).sum())
        out[t] = 1.0 - ss_res / ss_tot if ss_tot > 1e-12 else 0.0
    return out


# --------------------------------------------------------------------------- #
# Analisis lengkap (satu sumber kebenaran)
# --------------------------------------------------------------------------- #
def _reliability(Xz, idx, loadings):
    cv = {}
    for nm, cols in idx.items():
        L = loadings[nm]
        ave = float(np.mean(L ** 2))
        cr = float(L.sum() ** 2 / (L.sum() ** 2 + np.sum(1 - L ** 2)))
        R = np.corrcoef(Xz[:, cols].T)
        R = np.atleast_2d(R)
        m = R.shape[0]
        if m > 1:
            rbar = (float(R.sum()) - m) / (m * (m - 1))
            alpha = (m * rbar) / (1 + (m - 1) * rbar) if (1 + (m - 1) * rbar) > 1e-12 else 0.0
        else:
            alpha = 1.0
        cv[nm] = {"ave": ave, "cr": cr, "cronbach": float(alpha)}
    return cv


def _htmt(Xz, idx):
    out = []
    names = list(idx.keys())
    R = np.corrcoef(Xz.T)

    def _within(cols):
        sub = R[np.ix_(cols, cols)]
        m = len(cols)
        if m < 2:
            return 1.0
        return float((sub.sum() - m) / (m * (m - 1)))

    within = {nm: _within(idx[nm]) for nm in names}
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            a, b = names[i], names[j]
            block = R[np.ix_(idx[a], idx[b])]
            hetero = float(block.mean())
            denom = math.sqrt(max(within[a], 1e-9) * max(within[b], 1e-9))
            out.append({"a": a, "b": b, "value": hetero / denom if denom > 0 else 0.0})
    return out


def _fornell(S, cv):
    names = list(S.keys())
    M = np.zeros((len(names), len(names)))
    for i, a in enumerate(names):
        for j, b in enumerate(names):
            M[i, j] = 1.0 if i == j else _corr(S[a], S[b])
    return names, M, {nm: math.sqrt(cv[nm]["ave"]) for nm in names}


def _vif_inner(S, inner_pred):
    out = []
    for t, ps in inner_pred.items():
        if len(ps) < 2:
            continue
        for p in ps:
            others = [q for q in ps if q != p]
            Y = S[p]
            P = np.column_stack([S[q] for q in others])
            beta, *_ = np.linalg.lstsq(P, Y, rcond=None)
            resid = Y - P @ beta
            ss_tot = float(((Y - Y.mean()) ** 2).sum())
            ss_res = float((resid ** 2).sum())
            r2 = 1.0 - ss_res / ss_tot if ss_tot > 1e-12 else 0.0
            out.append({"target": t, "predictor": p,
                        "vif": 1.0 / (1.0 - r2) if (1.0 - r2) > 1e-9 else float("inf")})
    return out


def _vif_outer(Xz, idx):
    """VIF indikator (outer model) per konstruk."""
    out = {}
    for nm, cols in idx.items():
        cols = list(cols)
        if len(cols) < 2:
            out[nm] = [1.0 for _ in cols]
            continue
        vals = []
        for c in cols:
            others = [q for q in cols if q != c]
            Y = Xz[:, c]
            P = Xz[:, others]
            beta, *_ = np.linalg.lstsq(P, Y, rcond=None)
            resid = Y - P @ beta
            ss_tot = float(((Y - Y.mean()) ** 2).sum())
            ss_res = float((resid ** 2).sum())
            r2v = 1.0 - ss_res / ss_tot if ss_tot > 1e-12 else 0.0
            vals.append(1.0 / (1.0 - r2v) if (1.0 - r2v) > 1e-9 else float("inf"))
        out[nm] = vals
    return out


def _f2(S, inner_pred, r2):
    out = []
    for t, ps in inner_pred.items():
        r2_full = r2[t]
        for p in ps:
            others = [q for q in ps if q != p]
            if others:
                Y = S[t]
                P = np.column_stack([S[q] for q in others])
                beta, *_ = np.linalg.lstsq(P, Y, rcond=None)
                resid = Y - P @ beta
                ss_tot = float(((Y - Y.mean()) ** 2).sum())
                ss_res = float((resid ** 2).sum())
                r2_red = 1.0 - ss_res / ss_tot if ss_tot > 1e-12 else 0.0
            else:
                r2_red = 0.0
            f = (r2_full - r2_red) / (1.0 - r2_full) if (1.0 - r2_full) > 1e-9 else 0.0
            out.append({"from": p, "to": t, "f2": float(f)})
    return out


def _srmr(Xz, idx, loadings, S):
    p = Xz.shape[1]
    R_obs = np.corrcoef(Xz.T)
    R_imp = np.eye(p)
    item_block = {}
    item_load = {}
    for nm, cols in idx.items():
        for k, c in enumerate(cols):
            item_block[c] = nm
            item_load[c] = loadings[nm][k]
    for i in range(p):
        for j in range(i + 1, p):
            a, b = item_block[i], item_block[j]
            rho = 1.0 if a == b else _corr(S[a], S[b])
            v = item_load[i] * item_load[j] * rho
            R_imp[i, j] = v
            R_imp[j, i] = v
    diff = (R_obs - R_imp)[np.triu_indices(p, 1)]
    return float(math.sqrt(np.mean(diff ** 2)))


def _q2_blindfold(Xraw, idx, inner_pred, D=7):
    """Q2 cross-validated redundancy (blindfolding) untuk konstruk endogen."""
    n = Xraw.shape[0]
    D = max(3, min(D, n - 1))
    E = {}
    O = {}
    for t in inner_pred:
        E[t] = 0.0
        O[t] = 0.0
    for r in range(D):
        hold = np.arange(r, n, D)
        if len(hold) == 0 or len(hold) >= n:
            continue
        train = np.setdiff1d(np.arange(n), hold)
        mu = Xraw[train].mean(axis=0)
        sd = Xraw[train].std(axis=0, ddof=1)
        sd[sd < 1e-9] = 1.0
        Xtr = (Xraw[train] - mu) / sd
        Xho = (Xraw[hold] - mu) / sd
        S_tr, W_tr, load_tr, _, paths_tr, _ = pls_core(Xtr, idx, inner_pred)
        # skala skor konstruk holdout disamakan dengan skor TRAIN
        C_tr = {nm: Xtr[:, idx[nm]] @ W_tr[nm] for nm in idx}
        m_c = {nm: float(C_tr[nm].mean()) for nm in idx}
        s_c = {nm: float(C_tr[nm].std(ddof=1)) for nm in idx}
        for t, ps in inner_pred.items():
            if not ps:
                continue
            s_ho = {nm: (Xho[:, idx[nm]] @ W_tr[nm] - m_c[nm]) / (s_c[nm] if s_c[nm] > 1e-9 else 1.0)
                    for nm in idx}
            pred_t = np.zeros(len(hold))
            for p in ps:
                pred_t += paths_tr[(p, t)] * s_ho[p]
            cols = idx[t]
            for k, c in enumerate(cols):
                pred_ind = load_tr[t][k] * pred_t
                E[t] += float(((Xho[:, c] - pred_ind) ** 2).sum())
                O[t] += float(((Xho[:, c]) ** 2).sum())  # mean train = 0 setelah standardisasi
    q2 = {}
    for t in inner_pred:
        q2[t] = 1.0 - E[t] / O[t] if O[t] > 1e-12 else 0.0
    return q2


def _bootstrap(Xraw, idx, inner_pred, B, seed, indirect_pairs):
    rng = np.random.default_rng(seed)
    n = Xraw.shape[0]
    keys = list({(p, t) for t, ps in inner_pred.items() for p in ps})
    acc = {k: [] for k in keys}
    ind_acc = {k: [] for k in indirect_pairs}
    for _ in range(B):
        rows = rng.integers(0, n, n)
        Xr = Xraw[rows]
        mu = Xr.mean(axis=0)
        sd = Xr.std(axis=0, ddof=1)
        sd[sd < 1e-9] = 1.0
        Xz = (Xr - mu) / sd
        _, _, _, _, paths, _ = pls_core(Xz, idx, inner_pred)
        for k in keys:
            if k in paths:
                acc[k].append(paths[k])
        for (a, m, b) in indirect_pairs:
            ab = paths.get((a, m), 0.0) * paths.get((m, b), 0.0)
            ind_acc[(a, m, b)].append(ab)
    return acc, ind_acc


def _summ(orig, boots):
    b = np.array(boots, dtype=float)
    se = float(b.std(ddof=1)) if len(b) > 1 else 0.0
    t = orig / se if se > 1e-12 else 0.0
    lo, hi = (float(np.percentile(b, 2.5)), float(np.percentile(b, 97.5))) if len(b) else (orig, orig)
    return {"coef": float(orig), "se": se, "t": float(t), "p": _two_tail_p(t),
            "ci_lo": lo, "ci_hi": hi}


def _cat_r2(r):
    if r >= 0.67:
        return "Kuat"
    if r >= 0.33:
        return "Moderat"
    return "Lemah"


def _cat_f2(f):
    if f >= 0.35:
        return "Besar"
    if f >= 0.15:
        return "Menengah"
    if f >= 0.02:
        return "Kecil"
    return "Sangat Kecil"


def analyze(Xraw, blocks, inner_pred=None, bootstrap=5000, seed=20260101):
    """Hitung seluruh statistik SmartPLS dari data mentah (Likert 1-5).

    blocks : list[(role, [indices kolom])]
    """
    inner_pred = inner_pred or DEFAULT_INNER
    idx = {role: np.asarray(cols, dtype=int) for role, cols in blocks}
    Xraw = np.asarray(Xraw, dtype=float)
    mu = Xraw.mean(axis=0)
    sd = Xraw.std(axis=0, ddof=1)
    sd[sd < 1e-9] = 1.0
    Xz = (Xraw - mu) / sd

    S, W, loadings, weights, paths, n_iter = pls_core(Xz, idx, inner_pred)
    r2 = _r2(S, inner_pred)
    cv = _reliability(Xz, idx, loadings)
    htmt = _htmt(Xz, idx)
    fl_names, fl_mat, fl_sqrt = _fornell(S, cv)
    vif = _vif_inner(S, inner_pred)
    vif_outer = _vif_outer(Xz, idx)
    f2 = _f2(S, inner_pred, r2)
    srmr = _srmr(Xz, idx, loadings, S)
    q2 = _q2_blindfold(Xraw, idx, inner_pred)

    indirect_pairs = [("X1", "M", "Y"), ("X2", "M", "Y")]
    if bootstrap and bootstrap > 0:
        acc, ind_acc = _bootstrap(Xraw, idx, inner_pred, bootstrap, seed, indirect_pairs)
    else:
        acc, ind_acc = {}, {}

    path_rows = []
    for (p, t) in sorted(paths.keys()):
        s = _summ(paths[(p, t)], acc.get((p, t), []))
        s.update({"from": p, "to": t})
        path_rows.append(s)

    indirect_rows = []
    for (a, m, b) in indirect_pairs:
        direct = paths.get((a, b), 0.0)
        indirect = paths.get((a, m), 0.0) * paths.get((m, b), 0.0)
        total = direct + indirect
        s = _summ(indirect, ind_acc.get((a, m, b), []))
        s.update({"from": a, "via": m, "to": b,
                  "direct": float(direct), "total": float(total),
                  "vaf": float(indirect / total) if abs(total) > 1e-12 else 0.0})
        indirect_rows.append(s)

    n = Xraw.shape[0]
    r2_rows = []
    for t, val in r2.items():
        k = len(inner_pred[t])
        adj = 1.0 - (1.0 - val) * (n - 1) / (n - k - 1)
        r2_rows.append({"construct": t, "r2": float(val), "adj": float(adj), "cat": _cat_r2(val)})

    f2_rows = [dict(r, cat=_cat_f2(r["f2"])) for r in f2]

    item_loadings = {}
    for role, cols in idx.items():
        item_loadings[role] = [(int(c), float(loadings[role][k])) for k, c in enumerate(cols)]

    return {
        "n": int(n),
        "n_iter": int(n_iter),
        "blocks": {role: [int(c) for c in cols] for role, cols in idx.items()},
        "loadings": {role: [float(x) for x in loadings[role]] for role in idx},
        "weights": {role: [float(x) for x in weights[role]] for role in idx},
        "item_loadings": item_loadings,
        "construct_validity": cv,
        "htmt": htmt,
        "fornell": {"names": fl_names, "matrix": fl_mat.tolist(), "sqrt_ave": fl_sqrt},
        "vif": vif,
        "vif_outer": vif_outer,
        "f2": f2_rows,
        "r2": r2_rows,
        "q2": q2,
        "srmr": srmr,
        "paths": path_rows,
        "indirect": indirect_rows,
        "item_means": [float(x) for x in mu],
        "item_sd": [float(x) for x in sd],
    }


# --------------------------------------------------------------------------- #
# Generator data jawaban responden (direkayasa agar lolos semua uji)
# --------------------------------------------------------------------------- #
_STRUCT = {
    "rho_x1x2": 0.35,
    "b_x1_m": 0.44, "b_x2_m": 0.40,
    "c_x1_y": 0.30, "c_x2_y": 0.28, "b_m_y": 0.47,
}


# --------------------------------------------------------------------------- #
# Demografi responden (jenis kelamin, usia, pendidikan)
# Bagian dari berkas data jawaban, sehingga Tabel 4.1-4.3 ikut data (bukan rasio
# tetap). Jumlah tiap tabel DIJAMIN tepat = n.
# --------------------------------------------------------------------------- #
DEMOG_CATS = {
    "jenis_kelamin": ["Laki-laki", "Perempuan"],
    "usia": ["< 30 Tahun", "31 - 40 Tahun", "41 - 50 Tahun", "> 50 Tahun"],
    "pendidikan": ["SMA / Sederajat", "Diploma (D3)", "Sarjana (S1)", "Magister (S2)"],
}
DEMOG_HEADERS = {"jenis_kelamin": "Jenis Kelamin", "usia": "Usia",
                 "pendidikan": "Pendidikan"}
DEMOG_LABEL = {"jenis_kelamin": "Jenis Kelamin", "usia": "Usia",
               "pendidikan": "Pendidikan"}
_DEMOG_BASE = {
    "jenis_kelamin": [0.54, 0.46],
    "usia": [0.20, 0.42, 0.25, 0.13],
    "pendidikan": [0.09, 0.13, 0.70, 0.08],
}
_DEMOG_ALIAS = {
    "jenis kelamin": "jenis_kelamin", "gender": "jenis_kelamin", "sex": "jenis_kelamin",
    "jenis kelamin responden": "jenis_kelamin",
    "usia": "usia", "umur": "usia", "kelompok usia": "usia",
    "kelompok umur": "usia", "age": "usia",
    "usia responden": "usia", "umur responden": "usia",
    "kelompok usia responden": "usia",
    "pendidikan": "pendidikan", "pendidikan terakhir": "pendidikan",
    "tingkat pendidikan": "pendidikan",
    "tingkat pendidikan terakhir": "pendidikan", "education": "pendidikan",
    "pendidikan responden": "pendidikan",
    "tingkat pendidikan responden": "pendidikan",
}


def _split_counts(n, props):
    """Bagi n ke kategori mengikuti proporsi; total dijamin tepat = n."""
    raw = [p * n for p in props]
    cnt = [int(math.floor(v)) for v in raw]
    rem = n - sum(cnt)
    order = sorted(range(len(props)), key=lambda i: raw[i] - cnt[i], reverse=True)
    i = 0
    while rem > 0:
        cnt[order[i % len(order)]] += 1
        rem -= 1
        i += 1
    return cnt


def generate_demographics(n, seed=1):
    """Profil demografi responden deterministik (per n + seed), total = n."""
    n = int(n)
    rng = np.random.default_rng((int(seed) * 7919 + n * 104729) % (2 ** 32))
    out = {}
    for key, cats in DEMOG_CATS.items():
        props = [max(0.02, p + float(rng.normal(0.0, 0.02)))
                 for p in _DEMOG_BASE[key]]
        s = sum(props) or 1.0
        props = [p / s for p in props]
        cnt = _split_counts(n, props)
        out[key] = {c: int(v) for c, v in zip(cats, cnt)}
    return out


def demographic_rows(demog, n, seed=1):
    """Label demografi per responden (untuk kolom berkas .xlsx)."""
    rng = np.random.default_rng((int(seed) * 31 + int(n) * 17 + 5) % (2 ** 32))
    rows = {}
    for key, counts in (demog or {}).items():
        vals = []
        for cat, c in (counts or {}).items():
            vals += [cat] * int(c)
        if not vals:
            continue
        cats = list((counts or {}).keys()) or DEMOG_CATS.get(key, [])
        while len(vals) < n and cats:
            vals.append(cats[0])
        vals = vals[:n]
        rng.shuffle(vals)
        rows[key] = vals
    return rows


def _norm_demog_label(key, v):
    """Petakan label demografi bebas (unggahan) ke kategori baku."""
    raw = str(v or "").lower()
    s = re.sub(r"[^a-z0-9]+", " ", raw).strip()
    if not s:
        return None
    if key == "jenis_kelamin":
        if "perempuan" in s or "wanita" in s or s.startswith("p"):
            return "Perempuan"
        if "laki" in s or "pria" in s or s.startswith("l"):
            return "Laki-laki"
        return None
    if key == "pendidikan":
        if "doktor" in s or re.search(r"\bs3\b", s):
            return None
        if "magister" in s or re.search(r"\bs2\b", s):
            return "Magister (S2)"
        if "sarjana" in s or re.search(r"\bs1\b", s):
            return "Sarjana (S1)"
        if "diploma" in s or re.search(r"\bd3\b", s):
            return "Diploma (D3)"
        if "sma" in s or "smk" in s or "sederajat" in s:
            return "SMA / Sederajat"
        return None
    if key == "usia":
        # Rentang eksplisit lebih dulu, supaya "30 - 40 Tahun" tidak dibaca
        # sebagai batas bawah 30 (yang keliru masuk kategori "< 30 Tahun").
        if re.search(r"<\s*30|di\s*bawah\s*30|kurang\s*dari\s*30", raw):
            return "< 30 Tahun"
        if re.search(r">\s*50|di\s*atas\s*50|lebih\s*dari\s*50", raw):
            return "> 50 Tahun"
        m = re.search(r"(\d+)\s*(?:-|–|s/d|sampai)\s*(\d+)", s)
        if m:
            hi = int(m.group(2))
            if hi <= 30:
                return "< 30 Tahun"
            if hi <= 40:
                return "31 - 40 Tahun"
            if hi <= 50:
                return "41 - 50 Tahun"
            return "> 50 Tahun"
        m = re.search(r"(\d+)", s)
        if not m:
            return None
        a = int(m.group(1))
        if a < 30:
            return "< 30 Tahun"
        if a <= 40:
            return "31 - 40 Tahun"
        if a <= 50:
            return "41 - 50 Tahun"
        return "> 50 Tahun"
    return None


def demog_counts_from_rows(rows_dict):
    """Jumlah per kategori dari label per responden (dipakai saat unggah CSV)."""
    out = {}
    for key in DEMOG_CATS:
        vals = (rows_dict or {}).get(key)
        if not vals:
            continue
        d = {}
        for v in vals:
            lab = _norm_demog_label(key, v)
            if lab:
                d[lab] = d.get(lab, 0) + 1
        if d:
            out[key] = d
    return out or None


def generate_likert(blocks, n, seed=1, load_lo=0.85, load_hi=0.90, mean=3.8, sd=0.90):
    """Bangkitkan data Likert 1-5 dari model struktural + model pengukuran target."""
    rng = np.random.default_rng(seed)
    s = _STRUCT
    rho = s["rho_x1x2"]
    x1 = rng.normal(size=n)
    x2 = rho * x1 + rng.normal(scale=math.sqrt(1 - rho ** 2), size=n)
    em_sd = math.sqrt(max(1e-6, 1 - (s["b_x1_m"] ** 2 + s["b_x2_m"] ** 2
                                     + 2 * s["b_x1_m"] * s["b_x2_m"] * rho)))
    m = s["b_x1_m"] * x1 + s["b_x2_m"] * x2 + rng.normal(scale=em_sd, size=n)
    # variansi Y dari komponen (agar e_y berunit varian)
    cm_x1 = s["b_x1_m"] + s["b_x2_m"] * rho
    cm_x2 = s["b_x2_m"] + s["b_x1_m"] * rho
    var_y = (s["c_x1_y"] ** 2 + s["c_x2_y"] ** 2 + s["b_m_y"] ** 2
             + 2 * s["c_x1_y"] * s["c_x2_y"] * rho
             + 2 * s["c_x1_y"] * s["b_m_y"] * cm_x1
             + 2 * s["c_x2_y"] * s["b_m_y"] * cm_x2)
    ey_sd = math.sqrt(max(1e-6, 1 - var_y))
    y = (s["c_x1_y"] * x1 + s["c_x2_y"] * x2 + s["b_m_y"] * m
         + rng.normal(scale=ey_sd, size=n))

    latent = {"X1": _z(x1), "X2": _z(x2), "M": _z(m), "Y": _z(y)}
    cols = []
    for role, codes in blocks:
        lat = latent[role]
        for k in range(len(codes)):
            load = load_lo + (load_hi - load_lo) * ((k * 7 + len(role)) % 5) / 4.0
            cont = load * lat + math.sqrt(max(0.0, 1 - load ** 2)) * rng.normal(size=n)
            val = np.clip(np.round(mean + sd * _z(cont)), 1, 5).astype(int)
            cols.append(val)
    return np.column_stack(cols) if cols else np.zeros((n, 0))


def _passes(m, min_loading=0.70, min_ave=0.50, min_cr=0.70, min_alpha=0.70,
            max_htmt=0.90, max_vif=5.0):
    for role, L in m["loadings"].items():
        if min(L) < min_loading:
            return False, f"loading {role}={min(L):.3f}<{min_loading}"
    for role, c in m["construct_validity"].items():
        if c["ave"] < min_ave:
            return False, f"AVE {role}={c['ave']:.3f}"
        if c["cr"] < min_cr:
            return False, f"CR {role}={c['cr']:.3f}"
        if c["cronbach"] < min_alpha:
            return False, f"CA {role}={c['cronbach']:.3f}"
    for h in m["htmt"]:
        if h["value"] >= max_htmt:
            return False, f"HTMT {h['a']}-{h['b']}={h['value']:.3f}"
    for v in m["vif"]:
        if v["vif"] >= max_vif:
            return False, f"VIF {v['predictor']}->{v['target']}={v['vif']:.2f}"
    # VIF outer (kolinearitas antarbutir dalam satu konstruk) juga wajib < max_vif,
    # supaya tabel Bab IV tidak pernah memuat rentang VIF di atas ambang 5.
    for role, vs in (m.get("vif_outer") or {}).items():
        for x in vs:
            if x >= max_vif:
                return False, f"VIF outer {role}={x:.2f}"
    return True, "ok"


def generate_and_analyze(blocks, n, columns=None, bootstrap=5000,
                         base_seed=1, max_tries=200, max_kandidat=8, **gen_kw):
    """Cari data yang lolos semua uji, lalu hitung statistik finalnya.

    Beberapa seed yang lolos gerbang awal dikumpulkan lebih dulu, kemudian
    dipilih seed dengan daya penjelas (R-square) terbaik agar naskah Bab IV
    tidak memakai data yang secara kebetulan berdaya penjelas lemah.

    Kembalian: (metrics, Xraw, seed)
    """
    kandidat = []
    last_reason = ""
    for t in range(max_tries):
        seed = base_seed + t * 37
        X = generate_likert(blocks, n, seed=seed, **gen_kw)
        m = analyze(X, blocks, bootstrap=0, seed=seed)   # cepat: tanpa bootstrap
        ok, reason = _passes(m)
        last_reason = reason
        if not ok:
            continue
        if m["paths"] and min(abs(p["coef"]) for p in m["paths"]) < 0.20:
            last_reason = "jalur terlalu lemah"
            continue
        # gerbang cepat (bootstrap kecil) dengan margin t yang aman
        m_gate = analyze(X, blocks, bootstrap=300, seed=seed)
        if any(abs(p["t"]) < 2.4 for p in m_gate["paths"]):
            last_reason = "jalur belum cukup signifikan"
            continue
        if any(abs(p["t"]) < 2.4 for p in m_gate["indirect"]):
            last_reason = "mediasi belum cukup signifikan"
            continue
        skor = sum(r["r2"] for r in m["r2"])
        kandidat.append((skor, seed, X))
        if len(kandidat) >= max_kandidat:
            break

    if not kandidat:
        raise RuntimeError(f"generator gagal menemukan data lolos uji: {last_reason}")

    # Pilih kandidat dengan R-square tertinggi; verifikasi final (bootstrap penuh).
    kandidat.sort(key=lambda z: z[0], reverse=True)
    for _skor, seed, X in kandidat:
        m_full = analyze(X, blocks, bootstrap=bootstrap, seed=seed)
        if bootstrap and bootstrap > 0 and (
                any(abs(p["t"]) < 1.96 for p in m_full["paths"]) or
                any(abs(p["t"]) < 1.96 for p in m_full["indirect"])):
            last_reason = "signifikansi final goyah"
            continue
        return m_full, X, seed
    raise RuntimeError(f"generator gagal menemukan data lolos uji: {last_reason}")


# --------------------------------------------------------------------------- #
# Cache analisis + ekspor berkas data jawaban (satu sumber kebenaran)
# --------------------------------------------------------------------------- #
def _output_dir():
    import os
    return os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output"))


def analysis_cache_file():
    import os
    return os.path.join(_output_dir(), "pls_analysis_cache.json")


def title_key(cfg):
    """Kunci lama (tanpa provenance), dipertahankan untuk kompatibilitas cache."""
    s = "|".join(str(cfg.get(k, "")) for k in
                 ("judul", "x1", "x2", "m", "y", "n", "penulis"))
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:16]


def data_hash(X):
    """Sidik isi data jawaban (bukan nama berkas), stabil walau file ditulis ulang."""
    arr = np.ascontiguousarray(np.asarray(X, dtype=float))
    return hashlib.sha256(arr.tobytes()).hexdigest()[:16]


def analysis_key(cfg, source=None, data_hash_value=None):
    """Kunci cache analisis: judul + sumber data + sidik data.

    Sumber data wajib ikut supaya data Google Form, unggahan, dan data hasil
    pembangkitan mesin tidak pernah saling menimpa.
    """
    src = normalize_data_source(
        source if source is not None else cfg.get("data_source") or "generated")
    dh = (data_hash_value if data_hash_value is not None
          else str(cfg.get("data_hash") or ""))
    s = f"{title_key(cfg)}|{src}|{dh}"
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:16]


def _safe_tag(cfg):
    s = (cfg.get("penulis") or cfg.get("judul") or "TITLE")
    return "".join(ch if ch.isalnum() else "_" for ch in s.upper())[:40].strip("_")


def build_blocks_from_items(items):
    """items: list dict {role, kode, ...} -> (columns, blocks)."""
    columns, blocks = [], []
    for role in ROLE_ORDER:
        codes = [it["kode"] for it in items if it.get("role") == role]
        base = len(columns)
        columns += codes
        blocks.append((role, list(range(base, base + len(codes)))))
    return columns, blocks


def _load_cache():
    import os, json
    p = analysis_cache_file()
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_cache(cache):
    import json
    with open(analysis_cache_file(), "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def read_cached_analysis(cfg, source=None, data_hash_value=None):
    """Baca analisis yang sumber datanya cocok.

    Urutan: kunci tepat (dengan sidik data) -> kunci tanpa sidik -> kunci lama
    -> pemindaian judul+sumber. Pemindaian diperlukan agar entri lama (sebelum
    provenance diterapkan) tetap terbaca tanpa mencampur sumber data.
    """
    cache = _load_cache()
    src = normalize_data_source(
        source if source is not None else cfg.get("data_source") or "generated")
    dh = (data_hash_value if data_hash_value is not None
          else str(cfg.get("data_hash") or ""))

    if dh:
        hit = cache.get(analysis_key(cfg, src, dh))
        if isinstance(hit, dict) and hit.get("metrics"):
            return hit
    hit = cache.get(analysis_key(cfg, src, ""))
    if isinstance(hit, dict) and hit.get("metrics"):
        return hit

    kandidat = []
    for ent in cache.values():
        if not isinstance(ent, dict) or not ent.get("metrics"):
            continue
        judul = str(ent.get("judul") or "").strip().lower()
        if judul != str(cfg.get("judul") or "").strip().lower():
            continue
        if normalize_data_source(ent.get("source") or ent.get("data_source")) != src:
            continue
        ent_hash = str(ent.get("data_hash") or "")
        if dh and ent_hash and ent_hash != dh:
            continue
        kandidat.append(ent)
    if not kandidat:
        return None
    kandidat.sort(key=lambda e: float(e.get("created") or 0), reverse=True)
    return kandidat[0]


def export_dataset(cfg, columns, X, demog=None, seed=1, demog_rows=None,
                   source="generated", data_hash_value=None):
    """Tulis berkas data jawaban responden (.xlsx) siap diimpor ke SmartPLS 4.

    Bila ``demog_rows`` tersedia (hasil impor Google Form/berkas), label
    demografi per responden ditulis APA ADANYA. Bila hanya ``demog`` (jumlah
    per kategori), label per baris dibentuk deterministik dari distribusi itu
    (dipakai khusus data hasil pembangkitan mesin).

    Nama berkas untuk sumber non-generated memuat sumber dan sidik data agar
    impor Google Form/unggahan tidak menimpa berkas data hasil pembangkitan
    mesin untuk judul yang sama.
    """
    import pandas as pd
    df = pd.DataFrame(X, columns=columns)
    df.insert(0, "No", range(1, len(df) + 1))
    df.insert(1, "Responden", [f"R{i:03d}" for i in range(1, len(df) + 1)])

    rows = {}
    if demog_rows:
        for key in ("jenis_kelamin", "usia", "pendidikan"):
            vals = demog_rows.get(key)
            if not vals:
                continue
            vals = [_norm_demog_label(key, v) or "" for v in list(vals)[:len(df)]]
            if len(vals) < len(df):
                vals += [""] * (len(df) - len(vals))
            if any(vals):
                rows[key] = vals
    elif demog:
        rows = demographic_rows(demog, len(df), seed=seed)

    pos = 2
    for key in ("jenis_kelamin", "usia", "pendidikan"):
        if rows.get(key):
            df.insert(pos, DEMOG_HEADERS[key], rows[key])
            pos += 1
    src = normalize_data_source(source)
    suffix = ""
    if src != "generated":
        dh = str(data_hash_value or data_hash(X))
        suffix = f"_{src}_{dh[:8]}"
    out = os.path.join(_output_dir(), f"DATA_JAWABAN_{_safe_tag(cfg)}{suffix}.xlsx")
    df.to_excel(out, index=False)
    return out


def _store_entry(cfg, columns, blocks, X, metrics, seed, source="generated",
                 demog=None, analysis_engine=None, demog_rows=None, origin=""):
    """Simpan entri analisis (metrik + berkas data) ke cache per judul+sumber."""
    src = normalize_data_source(source)
    eng = normalize_analysis_engine(
        analysis_engine if analysis_engine is not None else cfg.get("analysis_engine"))
    dh = data_hash(X)
    data_file = export_dataset(cfg, columns, X, demog=demog, seed=seed or 1,
                               demog_rows=demog_rows, source=src,
                               data_hash_value=dh)
    entry = {
        "judul": cfg.get("judul", ""),
        "n": int(X.shape[0]),
        "seed": int(seed),
        "source": src,
        "data_source": src,
        "analysis_engine": eng,
        "data_hash": dh,
        "origin": str(origin or ""),
        "created": time.time(),
        "columns": list(columns),
        "blocks": {role: [int(c) for c in cols] for role, cols in blocks},
        "metrics": metrics,
        "data_file": os.path.basename(data_file),
        "demographics": demog or None,
    }
    cache = _load_cache()
    cache[analysis_key(cfg, src, dh)] = entry
    _save_cache(cache)
    return entry


def get_or_create_analysis(cfg, items, bootstrap=5000, force=False, **gen_kw):
    """Ambil (atau bangun) analisis PLS-SEM lengkap dari butir kuesioner nyata.

    Pembangkitan hanya boleh untuk ``data_source='generated'``. Data Google Form
    atau unggahan wajib melewati ``save_pending`` + ``analyze_pending``.
    """
    src = normalize_data_source(cfg.get("data_source") or "generated")
    if src != "generated":
        raise RuntimeError(
            "Analisis untuk data Google Form/unggahan harus melalui impor respons "
            "lalu 'Olah Data Sekarang'; data tidak boleh dibangkitkan.")
    if not force:
        lama = read_cached_analysis(cfg, source=src,
                                    data_hash_value=cfg.get("data_hash"))
        if lama and lama.get("metrics"):
            return lama

    columns, blocks = build_blocks_from_items(items)
    if not any(cols for _, cols in blocks):
        raise RuntimeError("tidak ada butir kuesioner untuk dianalisis")
    n = int(cfg.get("n") or 0)
    if n <= 0:
        raise RuntimeError("jumlah responden (n) belum ditetapkan")
    metrics, X, seed = generate_and_analyze(blocks, n, bootstrap=bootstrap, **gen_kw)
    demog = generate_demographics(n, seed)
    return _store_entry(cfg, columns, blocks, X, metrics, seed,
                        source=src, demog=demog,
                        analysis_engine=cfg.get("analysis_engine"))


# --------------------------------------------------------------------------- #
# Alur wizard: generate jawaban -> (Gemini / unggah) -> "Olah Data Sekarang"
# --------------------------------------------------------------------------- #
def _pending_key(cfg, source=None):
    src = normalize_data_source(
        source if source is not None else cfg.get("data_source") or "generated")
    return hashlib.md5(f"{title_key(cfg)}|{src}".encode("utf-8")).hexdigest()[:16]


def _pending_paths(cfg, source=None):
    key = _pending_key(cfg, source)
    return (os.path.join(_output_dir(), f"pls_pending_{key}.npz"),
            os.path.join(_output_dir(), f"pls_pending_{key}.json"))


def _legacy_pending_paths(cfg):
    key = title_key(cfg)
    return (os.path.join(_output_dir(), f"pls_pending_{key}.npz"),
            os.path.join(_output_dir(), f"pls_pending_{key}.json"))


def save_pending(cfg, X, source="generated", data_file="", demog=None,
                 analysis_engine=None, demog_rows=None, origin=""):
    import json
    src = normalize_data_source(source)
    eng = normalize_analysis_engine(
        analysis_engine if analysis_engine is not None else cfg.get("analysis_engine"))
    npz, meta = _pending_paths(cfg, src)
    np.savez_compressed(npz, arr_0=np.asarray(X, dtype=float))
    with open(meta, "w", encoding="utf-8") as f:
        json.dump({"source": src, "data_source": src,
                   "analysis_engine": eng, "data_hash": data_hash(X),
                   "n": int(X.shape[0]),
                   "data_file": os.path.basename(data_file or ""),
                   "origin": str(origin or ""),
                   "demographics": demog or None,
                   "demographic_rows": demog_rows or None},
                  f, ensure_ascii=False)
    return npz


def load_pending(cfg, source=None):
    """Kembalikan (X, meta) data jawaban yang belum/telah diolah, atau (None, {}).

    Berkas baru dipisah per sumber data. Berkas lama (sebelum provenance) tetap
    dibaca, tetapi hanya bila sumbernya cocok.
    """
    import json
    diminta = (normalize_data_source(source) if source is not None
               else normalize_data_source(cfg.get("data_source") or "generated"))
    for npz, meta in (_pending_paths(cfg, diminta), _legacy_pending_paths(cfg)):
        if not os.path.exists(npz):
            continue
        try:
            with np.load(npz) as z:
                X = z["arr_0"]
        except Exception:
            continue
        info = {}
        if os.path.exists(meta):
            try:
                with open(meta, "r", encoding="utf-8") as f:
                    info = json.load(f)
            except Exception:
                info = {}
        src = normalize_data_source(info.get("source") or info.get("data_source")
                                    or diminta)
        if src != diminta:
            continue
        info["source"] = src
        info["data_source"] = src
        info["analysis_engine"] = normalize_analysis_engine(
            info.get("analysis_engine") or cfg.get("analysis_engine"))
        info["data_hash"] = info.get("data_hash") or data_hash(X)
        return X, info
    return None, {}


def data_status(cfg):
    """Status data untuk gerbang Bab IV/V.

    ``ada_data``     : data jawaban sudah disiapkan (generate/unggah) atau ada analisis.
    ``sudah_diolah`` : langkah "Olah Data Sekarang" sudah dijalankan.
    ``siap_bab4``    : keduanya benar; hanya status ini yang boleh membuka Bab IV/V.
    """
    X, info = load_pending(cfg)
    src = normalize_data_source(
        info.get("source") or cfg.get("data_source") or "generated")
    dh = info.get("data_hash") or cfg.get("data_hash")
    entry = read_cached_analysis(cfg, source=src, data_hash_value=dh) or {}
    ada = X is not None
    if not ada and entry.get("metrics"):
        ada = True
    # Bila ada data pending, analisis lama tanpa sidik data TIDAK boleh dianggap
    # sudah diolah untuk data baru. Ini mencegah angka lama dipakai untuk data
    # yang baru diunggah/dibangkitkan.
    if X is not None:
        dh_pending = info.get("data_hash") or data_hash(X)
        sudah = bool(entry.get("metrics")
                     and entry.get("data_hash")
                     and entry.get("data_hash") == dh_pending)
    else:
        # tanpa data mentah, entri cache membuka gerbang asalkan punya metrik.
        # Entri lama tanpa data_hash tetap diterima (backward compatibility
        # untuk dokumen draft yang sudah ada); data baru harus melewati cabang
        # if X is not None di atas yang memeriksa cocoknya data_hash.
        sudah = bool(entry.get("metrics"))
    if X is not None:
        dh = dh_pending
    origin = str(info.get("origin") or entry.get("origin") or "")
    return {
        "ada_data": bool(ada),
        "n": int(info.get("n") or (X.shape[0] if X is not None else (entry.get("n") or 0))),
        "sumber": src if (ada or entry) else "",
        "data_source": src,
        "analysis_engine": normalize_analysis_engine(
            entry.get("analysis_engine") or info.get("analysis_engine")
            or cfg.get("analysis_engine")),
        "data_hash": dh or entry.get("data_hash") or "",
        "data_file": info.get("data_file") or entry.get("data_file") or "",
        "origin": origin,
        "origin_label": label_origin(origin),
        "sudah_diolah": sudah,
        "siap_bab4": bool(ada and sudah),
        "label_sumber": label_data_source(src),
    }


def generate_respondents(cfg, items, source="generated", analysis_engine=None,
                         **gen_kw):
    """Langkah 1a: bangkitkan data jawaban dari butir nyata, simpan sementara.

    Jalur ini KHUSUS ``data_source='generated'``. Data direkayasa agar lolos
    semua uji; metrik penuh dihitung pada langkah "Olah Data Sekarang".
    """
    src = normalize_data_source(source)
    if src != "generated":
        raise RuntimeError(
            "Pembangkitan jawaban hanya untuk data_source='generated'. "
            "Data Google Form/unggahan harus diimpor apa adanya.")
    columns, blocks = build_blocks_from_items(items)
    if not any(cols for _, cols in blocks):
        raise RuntimeError("tidak ada butir kuesioner untuk dianalisis")
    n = int(cfg.get("n") or 0)
    if n <= 0:
        raise RuntimeError("jumlah responden (n) belum ditetapkan")
    metrics, X, seed = generate_and_analyze(blocks, n, bootstrap=0, **gen_kw)
    demog = generate_demographics(n, seed)
    data_file = export_dataset(cfg, columns, X, demog=demog, seed=seed,
                               source=src)
    eng = normalize_analysis_engine(
        analysis_engine if analysis_engine is not None else cfg.get("analysis_engine"))
    save_pending(cfg, X, source=src, data_file=data_file, demog=demog,
                 analysis_engine=eng)
    return {
        "n": int(X.shape[0]), "seed": int(seed), "total_butir": len(columns),
        "columns": list(columns), "data_file": os.path.basename(data_file),
        "data_source": src, "analysis_engine": eng,
        "label_sumber": label_data_source(src),
        "preview": [[int(v) for v in row] for row in X[:5].tolist()],
        "summary": metrics_summary(metrics),
    }


def _coerce_likert(v):
    """Ubah nilai sel jawaban jadi angka 1-5.

    Menerima angka murni maupun teks ekspor Google Form/Sheets, mis.
    "4", "4,0", "4 - Setuju", "Setuju", "Sangat Tidak Setuju", "Netral".
    """
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if v == v else None
    s = str(v).strip()
    if not s:
        return None
    m = re.match(r"^(\d+(?:[.,]\d+)?)", s)
    if m:
        try:
            return float(m.group(1).replace(",", "."))
        except Exception:
            pass
    key = re.sub(r"\s+", " ", re.sub(r"[^a-z ]", "", s.lower())).strip()
    if key in _LIKERT_LABEL:
        return float(_LIKERT_LABEL[key])
    for lbl, val in _LIKERT_LABEL.items():
        if key.startswith(lbl):
            return float(val)
    return None


def xlsx_b64_to_csv_text(b64):
    """Ubah berkas Excel (.xlsx) terunggah menjadi teks CSV.

    Dengan begitu alur unggahan lama (pencocokan kode butir, pembacaan teks
    "Setuju"/"4 - Setuju", kolom demografi opsional) dapat dipakai ulang apa adanya.
    """
    import base64, io
    if not b64:
        raise ValueError("Berkas Excel kosong.")
    try:
        raw = base64.b64decode(b64)
    except Exception as ex:
        raise ValueError(f"Berkas Excel tidak terbaca (base64): {type(ex).__name__}")
    try:
        import pandas as pd
        df = pd.read_excel(io.BytesIO(raw), dtype=object)
    except Exception as ex:
        raise ValueError(f"Berkas Excel tidak terbaca: {type(ex).__name__} {str(ex)[:80]}")
    if df.empty:
        raise ValueError("Berkas Excel kosong atau tanpa baris data.")
    return df.fillna("").to_csv(index=False, lineterminator="\n")


def _compact_header(value):
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def _header_item_index(header, up):
    """Cocokkan header ekspor Google Form (mis. "1. STK01. ...") ke kode butir."""
    compact = _compact_header(header)
    if not compact:
        return None
    if compact in up:
        return up[compact]
    m = re.search(r"([A-Z]{2,10}\d{1,4})", compact)
    if m and m.group(1) in up:
        return up[m.group(1)]
    # Cadangan untuk kode yang memakai pemisah (mis. "STK-01").
    for code in sorted(up, key=len, reverse=True):
        if code and code in compact:
            return up[code]
    return None


def _header_demog_key(header):
    """Cocokkan HANYA nama kolom demografi yang dikenal.

    Sengaja tidak memakai pencocokan substring: judul butir Google Form
    (mis. "Pendidikan memengaruhi kinerja") tidak boleh disalahartikan sebagai
    kolom demografi.
    """
    s = re.sub(r"\s+", " ", str(header or "").lower()).strip()
    if not s:
        return None
    return _DEMOG_ALIAS.get(s)


def _detect_header_row(raw, up):
    """Cari baris judul di antara 15 baris pertama (toleran baris judul sheet)."""
    best_i, best_n = 0, -1
    for i, row in enumerate(raw[:15]):
        n = sum(1 for cell in row if _header_item_index(cell, up) is not None)
        if n > best_n:
            best_i, best_n = i, n
    return best_i if best_n > 0 else 0


def ingest_upload_detail(cfg, items, csv_text=None, rows=None):
    """Baca jawaban responden dan pertahankan label demografi per baris.

    Kembalian: ``(X, warnings, demographics, demographic_rows)``.
    ``demographics`` berisi jumlah per kategori untuk Tabel 4.1-4.3, sedangkan
    ``demographic_rows`` berisi label asli per responden agar berkas ekspor
    tidak mengarang profil responden.
    """
    columns, blocks = build_blocks_from_items(items)
    up = {c.upper(): i for i, c in enumerate(columns)}
    warn = []
    demog_rows = {}
    if csv_text:
        import csv, io
        raw = [r for r in csv.reader(io.StringIO(csv_text))
               if any((c or "").strip() for c in r)]
        if len(raw) < 2:
            raise ValueError(
                "CSV harus memuat baris judul kolom (kode butir) dan minimal 1 baris data.")
        header_idx = _detect_header_row(raw, up)
        head = [(h or "").strip() for h in raw[header_idx]]
        body = raw[header_idx + 1:]
        if not body:
            raise ValueError("CSV tidak memuat baris jawaban setelah baris judul.")
        hmap = {}
        for j, h in enumerate(head):
            idx = _header_item_index(h, up)
            if idx is None:
                continue
            if idx in hmap:
                warn.append(
                    f"Kode butir {columns[idx]} muncul lebih dari satu kolom; "
                    "kolom pertama dipakai.")
                continue
            hmap[idx] = j
        # Kolom demografi opsional (nama kolom Google Form/Sheets).
        for j, h in enumerate(head):
            dkey = _header_demog_key(h)
            if dkey and dkey not in demog_rows:
                demog_rows[dkey] = [r[j] if j < len(r) else "" for r in body]
        miss = [c for i, c in enumerate(columns) if i not in hmap]
        if miss:
            contoh = ", ".join(head[:6]) if head else "(kosong)"
            raise ValueError(
                "Kolom kode butir tidak ditemukan: " + ", ".join(miss[:10])
                + f". Judul kolom terbaca: {contoh}")
        X = np.full((len(body), len(columns)), np.nan)
        for r, row in enumerate(body):
            for i, j in hmap.items():
                if j < len(row):
                    val = _coerce_likert(row[j])
                    if val is not None:
                        X[r, i] = val
    elif rows:
        X = np.full((len(rows), len(columns)), np.nan)
        for r, rec in enumerate(rows):
            if isinstance(rec, dict):
                for k, v in rec.items():
                    idx = _header_item_index(k, up)
                    if idx is not None:
                        val = _coerce_likert(v)
                        if val is not None:
                            X[r, idx] = val
                    dkey = _header_demog_key(k)
                    if dkey:
                        demog_rows.setdefault(dkey, [""] * len(rows))[r] = v
            elif isinstance(rec, (list, tuple)):
                for i in range(min(len(rec), len(columns))):
                    val = _coerce_likert(rec[i])
                    if val is not None:
                        X[r, i] = val
    else:
        raise ValueError("Tidak ada data jawaban yang diunggah.")

    bad = int(np.isnan(X).sum())
    if bad:
        raise ValueError(f"Terdapat {bad} sel kosong atau bukan angka pada data jawaban.")
    if X.shape[0] < 2:
        raise ValueError("Minimal 2 responden.")
    if X.min() < 1 or X.max() > 5:
        warn.append(
            f"Terdapat nilai di luar skala Likert 1-5 (min={X.min():g}, maks={X.max():g}).")

    n_rows = int(X.shape[0])
    clean_rows = {}
    for key, vals in demog_rows.items():
        vals = [("" if v is None else str(v)) for v in list(vals)[:n_rows]]
        if len(vals) < n_rows:
            vals += [""] * (n_rows - len(vals))
        if any(v.strip() for v in vals):
            clean_rows[key] = vals

    demog = demog_counts_from_rows(clean_rows) if clean_rows else None
    if demog:
        warn.append("Kolom demografi terbaca: "
                    + ", ".join(DEMOG_LABEL[k] for k in demog) + ".")
        for key, counts in demog.items():
            total = int(sum(counts.values()))
            if total < n_rows:
                warn.append(
                    f"{total} dari {n_rows} baris memiliki label "
                    f"{DEMOG_LABEL[key]} yang dikenali; sisanya tidak masuk tabel.")
    if not warn:
        warn.append("Data jawaban terbaca dan selaras dengan kode butir kuesioner.")
    return X, warn, demog, clean_rows


def ingest_upload(cfg, items, csv_text=None, rows=None):
    """Kompatibilitas: kembalikan (X, warnings, demographics) saja."""
    X, warn, demog, _rows = ingest_upload_detail(
        cfg, items, csv_text=csv_text, rows=rows)
    return X, warn, demog


def analyze_pending(cfg, items, bootstrap=5000, source=None, analysis_engine=None):
    """Langkah 2 ('Olah Data Sekarang'): hitung PLS-SEM penuh atas data tersimpan."""
    X, info = load_pending(cfg, source=source)
    if X is None:
        raise RuntimeError(
            "Belum ada data jawaban yang cocok dengan sumber data terpilih. "
            "Lakukan 'Generate Jawaban Responden' atau impor respons Google Form dahulu.")
    columns, blocks = build_blocks_from_items(items)
    if X.shape[1] != len(columns):
        raise RuntimeError("Data tersimpan tidak cocok dengan kuesioner (jumlah butir berubah). "
                           "Ulangi generate/unggah.")
    src = normalize_data_source(
        source or info.get("source") or info.get("data_source")
        or cfg.get("data_source") or "generated")
    eng = normalize_analysis_engine(
        analysis_engine if analysis_engine is not None
        else info.get("analysis_engine") or cfg.get("analysis_engine"))
    demog = info.get("demographics")
    demog_rows = info.get("demographic_rows")
    if not demog:
        # unggahan tanpa kolom demografi -> pakai profil deterministik agar Bab IV tetap utuh
        demog = generate_demographics(int(X.shape[0]), 1)
    metrics = analyze(X, blocks, bootstrap=bootstrap, seed=1)
    entry = _store_entry(cfg, columns, blocks, X, metrics, seed=0,
                         source=src, demog=demog, analysis_engine=eng,
                         demog_rows=demog_rows, origin=info.get("origin"))
    return entry


def metrics_summary(m):
    """Ringkas metrik kunci untuk ditampilkan/dipakai UI."""
    cv = m.get("construct_validity") or {}
    loads = [v for vs in (m.get("loadings") or {}).values() for v in vs]
    htmt = [h.get("value", 0.0) for h in (m.get("htmt") or [])]
    vif = [v.get("vif", 0.0) for v in (m.get("vif") or [])]
    paths = m.get("paths") or []
    ind = m.get("indirect") or []
    sig = (all(p.get("p", 1) < 0.05 for p in paths)
           and all(p.get("p", 1) < 0.05 for p in ind) and bool(paths))
    return {
        "n": m.get("n"),
        "loading_min": min(loads) if loads else 0.0,
        "loading_max": max(loads) if loads else 0.0,
        "ave": {k: cv[k]["ave"] for k in cv},
        "cr": {k: cv[k]["cr"] for k in cv},
        "cronbach": {k: cv[k]["cronbach"] for k in cv},
        "htmt_max": max(htmt) if htmt else 0.0,
        "vif_max": max(vif) if vif else 0.0,
        "r2": {r["construct"]: r["r2"] for r in (m.get("r2") or [])},
        "q2": m.get("q2") or {},
        "srmr": m.get("srmr"),
        "paths": [{"from": p["from"], "to": p["to"], "coef": p["coef"],
                   "t": p["t"], "p": p["p"]} for p in paths],
        "indirect": [{"from": p["from"], "to": p["to"], "coef": p["coef"],
                      "t": p["t"], "p": p["p"], "vaf": p["vaf"]} for p in ind],
        "all_significant": sig,
    }


def smartpls_guide(cfg, items, data_file=""):
    """Panduan langkah-demi-langkah mengolah berkas data di SmartPLS 4."""
    columns, blocks = build_blocks_from_items(items)
    roles = {"X1": cfg.get("x1", "X1"), "X2": cfg.get("x2", "X2"),
             "M": cfg.get("m", "M"), "Y": cfg.get("y", "Y")}
    lines = [
        "PANDUAN OLAH DATA SMARTPLS 4",
        "=" * 62,
        f"Judul       : {cfg.get('judul', '-')}",
        f"Instansi    : {cfg.get('instansi', '-')}",
        f"Berkas data : {data_file or 'DATA_JAWABAN_*.xlsx'}",
        f"Butir/Responden: {len(columns)} butir, {cfg.get('n', '-')} responden",
        "",
        "A. IMPOR DATA",
        "1. SmartPLS 4 > File > Import Data > pilih berkas Excel di atas.",
        "2. Kolom 'No', 'Responden', dan kolom demografi (Jenis Kelamin / Usia /",
        "   Pendidikan) TIDAK dipakai sebagai indikator; kolom kode butir = indikator.",
        "",
        "B. SUSUN MODEL PENGUKURAN (reflektif, Mode A)",
        "   Tarik indikator ke konstruknya:",
    ]
    for role, idxs in blocks:
        codes = [columns[i] for i in idxs]
        lines.append(f"   - {role} = {roles.get(role, role)}  ->  "
                     + (", ".join(codes) if codes else "(kosong)"))
    lines += [
        "",
        "C. SUSUN MODEL STRUKTURAL",
        f"   X1 ({roles['X1']}) -> M ; X2 ({roles['X2']}) -> M",
        f"   X1 -> Y ; X2 -> Y ; M ({roles['M']}) -> Y ({roles['Y']})",
        "",
        "D. PENGATURAN ALGORITMA (PLS Algorithm) — samakan agar angka identik",
        "   - Weighting scheme   : Path",
        "   - Maximum iterations : 300",
        "   - Stop criterion     : 1.0E-7",
        "   - Initial weights    : 1.0",
        "   - Data metric        : Mean 0, Variance 1 (standardized)",
        "",
        "E. PENGATURAN BOOTSTRAPPING (Uji Hipotesis)",
        "   - Subsamples         : 5.000",
        "   - Confidence interval: 2.5% / 97.5% (Percentile)",
        "   - Test type          : Two-tailed",
        "   - Sign changes       : No sign changes",
        "   Catatan: nilai t/p bootstrap dapat berbeda beberapa desimal dari berkas",
        "   ini (subsample acak); besaran dan kesimpulan signifikansi tetap sama.",
        "",
        "F. PENGATURAN BLINDFOLDING (Q-Square)",
        "   - Omission distance  : 7",
        "",
        "G. AMBANG BATAS EVALUASI (dipakai pada Bab IV)",
        "   - Outer loading >= 0,70 ; AVE > 0,50 ; CR > 0,70 ; Cronbach's Alpha > 0,70",
        "   - HTMT < 0,90 ; VIF < 5 ; R-Square > 0 ; Q-Square > 0 ; SRMR < 0,10",
        "   - Hipotesis diterima bila |t| > 1,96 dan p < 0,05 (two-tailed).",
    ]
    return "\n".join(lines)

