"""Signal alignment helpers used by the MF4 Reader GUI."""

from __future__ import annotations

from typing import Optional, Tuple

import numpy as np


def _clean_xy(x, y) -> Tuple[np.ndarray, np.ndarray]:
    """Convert inputs to sorted finite 1-D float arrays with unique X values."""
    x_arr = np.asarray(x, dtype=float).reshape(-1)
    y_arr = np.asarray(y, dtype=float).reshape(-1)
    size = min(x_arr.size, y_arr.size)
    if size == 0:
        return np.empty(0, dtype=float), np.empty(0, dtype=float)

    x_arr = x_arr[:size]
    y_arr = y_arr[:size]
    mask = np.isfinite(x_arr) & np.isfinite(y_arr)
    x_arr = x_arr[mask]
    y_arr = y_arr[mask]
    if x_arr.size == 0:
        return x_arr, y_arr

    order = np.argsort(x_arr, kind="mergesort")
    x_arr = x_arr[order]
    y_arr = y_arr[order]

    # np.interp expects increasing X. For duplicate timestamps, keep the last
    # sample, which is generally the most recent value at that instant.
    if x_arr.size > 1:
        reverse_unique = np.unique(x_arr[::-1], return_index=True)[1]
        keep = np.sort(x_arr.size - 1 - reverse_unique)
        x_arr = x_arr[keep]
        y_arr = y_arr[keep]

    return x_arr, y_arr


def trigger_time(x, y, rule: str, level: float) -> Optional[float]:
    """Return the first interpolated threshold crossing time.

    Parameters
    ----------
    x, y:
        Time and signal samples.
    rule:
        ``"rising"`` finds a crossing from below the threshold to at/above it.
        ``"falling"`` finds a crossing from above the threshold to at/below it.
    level:
        Threshold value.

    Returns ``None`` if no crossing can be found.
    """
    x_arr, y_arr = _clean_xy(x, y)
    if x_arr.size < 2:
        return None

    mode = str(rule).strip().lower()
    if mode not in {"rising", "falling"}:
        raise ValueError("rule must be 'rising' or 'falling'")

    threshold = float(level)
    y0 = y_arr[:-1]
    y1 = y_arr[1:]

    if mode == "rising":
        indices = np.flatnonzero((y0 < threshold) & (y1 >= threshold))
    else:
        indices = np.flatnonzero((y0 > threshold) & (y1 <= threshold))

    if indices.size == 0:
        return None

    i = int(indices[0])
    xa, xb = float(x_arr[i]), float(x_arr[i + 1])
    ya, yb = float(y_arr[i]), float(y_arr[i + 1])

    # Linear interpolation gives a more useful event time than simply returning
    # the next sample timestamp. Flat segments are handled conservatively.
    dy = yb - ya
    if dy == 0.0 or not np.isfinite(dy):
        return xb

    fraction = (threshold - ya) / dy
    fraction = float(np.clip(fraction, 0.0, 1.0))
    return xa + fraction * (xb - xa)


def _robust_normalize(values: np.ndarray) -> np.ndarray:
    values = np.asarray(values, dtype=float)
    centered = values - np.nanmedian(values)
    scale = np.nanpercentile(np.abs(centered), 75)
    if not np.isfinite(scale) or scale <= 1e-12:
        scale = np.nanstd(centered)
    if not np.isfinite(scale) or scale <= 1e-12:
        return np.zeros_like(centered)
    normalized = centered / scale
    # Large spikes can dominate correlation. Clipping preserves the broad trend
    # the GUI's "Auto alignment" feature is intended to match.
    return np.clip(normalized, -8.0, 8.0)


def estimate_offset(ref_x, ref_y, x, y, max_points: int = 6000) -> Tuple[float, float]:
    """Estimate the time shift needed to align ``(x, y)`` to the reference.

    Returns ``(offset_seconds, quality)``. The returned offset follows the GUI's
    convention: add it to the target X values. A target event that occurs 2 s
    later than the reference therefore returns approximately ``-2.0``.

    ``quality`` is a bounded 0..1 normalized correlation score. If the signals
    are unsuitable for correlation, ``(0.0, 0.0)`` is returned.
    """
    rx, ry = _clean_xy(ref_x, ref_y)
    tx, ty = _clean_xy(x, y)
    if rx.size < 3 or tx.size < 3:
        return 0.0, 0.0

    # Estimate a practical sample interval from both recordings, then limit the
    # working grid to keep correlation responsive for very large measurement files.
    diffs = []
    for axis in (rx, tx):
        d = np.diff(axis)
        d = d[np.isfinite(d) & (d > 0)]
        if d.size:
            diffs.append(float(np.median(d)))
    if not diffs:
        return 0.0, 0.0

    dt = max(min(diffs), np.finfo(float).eps)
    start = min(float(rx[0]), float(tx[0]))
    stop = max(float(rx[-1]), float(tx[-1]))
    duration = stop - start
    if not np.isfinite(duration) or duration <= 0:
        return 0.0, 0.0

    estimated_points = int(np.ceil(duration / dt)) + 1
    if estimated_points > max_points:
        dt = duration / float(max_points - 1)
        estimated_points = max_points
    estimated_points = max(3, estimated_points)

    grid = start + np.arange(estimated_points, dtype=float) * dt

    ref_valid = (grid >= rx[0]) & (grid <= rx[-1])
    target_valid = (grid >= tx[0]) & (grid <= tx[-1])

    ref_series = np.zeros(grid.size, dtype=float)
    target_series = np.zeros(grid.size, dtype=float)
    ref_series[ref_valid] = _robust_normalize(np.interp(grid[ref_valid], rx, ry))
    target_series[target_valid] = _robust_normalize(np.interp(grid[target_valid], tx, ty))

    if np.allclose(ref_series, 0.0) or np.allclose(target_series, 0.0):
        return 0.0, 0.0

    # Correlation lag is directly the amount to add to target timestamps when
    # using np.correlate(reference, target, 'full').
    corr = np.correlate(ref_series, target_series, mode="full")
    lags = np.arange(-target_series.size + 1, ref_series.size, dtype=int)

    # Ignore shifts that leave too little overlapping real data. This prevents
    # an accidental match of short edge fragments from winning.
    min_overlap = max(3, int(0.10 * min(ref_valid.sum(), target_valid.sum())))
    valid_lag = np.zeros_like(corr, dtype=bool)
    overlap_counts = np.zeros_like(corr, dtype=int)
    for idx, lag in enumerate(lags):
        if lag >= 0:
            a = ref_valid[lag:]
            b = target_valid[: target_valid.size - lag]
        else:
            a = ref_valid[: ref_valid.size + lag]
            b = target_valid[-lag:]
        overlap = int(np.count_nonzero(a & b))
        overlap_counts[idx] = overlap
        valid_lag[idx] = overlap >= min_overlap

    if not np.any(valid_lag):
        return 0.0, 0.0

    score = np.full_like(corr, -np.inf, dtype=float)
    score[valid_lag] = corr[valid_lag]
    best_index = int(np.argmax(score))
    best_lag = int(lags[best_index])
    offset = float(best_lag * dt)

    # Recalculate Pearson correlation only on samples that actually overlap at
    # the winning lag; this is easier to interpret as a confidence/quality value.
    if best_lag >= 0:
        ref_slice = ref_series[best_lag:]
        trg_slice = target_series[: target_series.size - best_lag]
        mask = ref_valid[best_lag:] & target_valid[: target_valid.size - best_lag]
    else:
        ref_slice = ref_series[: ref_series.size + best_lag]
        trg_slice = target_series[-best_lag:]
        mask = ref_valid[: ref_valid.size + best_lag] & target_valid[-best_lag:]

    ref_overlap = ref_slice[mask]
    trg_overlap = trg_slice[mask]
    if ref_overlap.size < 3 or np.std(ref_overlap) <= 1e-12 or np.std(trg_overlap) <= 1e-12:
        quality = 0.0
    else:
        pearson = float(np.corrcoef(ref_overlap, trg_overlap)[0, 1])
        quality = float(np.clip(pearson, 0.0, 1.0)) if np.isfinite(pearson) else 0.0

    return offset, quality
