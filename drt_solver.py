import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import nnls
import gc

# NumPy 2.0+ 호환: trapz → trapezoid
_trapz = getattr(np, 'trapezoid', None) or np.trapz


def solve_drt_core(freq, z_re, z_im, mode, lam):
    """
    DRT 계산 — MATLAB DRTtools 이식 (메모리 최적화).

    Parameters
    ----------
    freq : array-like   주파수 (Hz)
    z_re : array-like   Re(Z)
    z_im : array-like   -Im(Z) (양수, EIS 관례)
    mode : int           1=w/o Induc, 2=with Induc, 3=Discard Inductive
    lam  : float         정규화 파라미터 λ
    """
    freq = np.asarray(freq, dtype=float).copy()
    z_re = np.asarray(z_re, dtype=float).copy()
    z_im = np.asarray(z_im, dtype=float).copy()

    # 부호 안전 검사: z_im은 -Im(Z) (양수)여야 함
    # Im(Z)(음수)가 들어오면 자동 변환
    if np.mean(z_im) < 0:
        z_im = -z_im

    if mode == 3:
        mask = z_im >= 0
        freq, z_re, z_im = freq[mask], z_re[mask], z_im[mask]

    omega = 2 * np.pi * freq
    n = len(freq)
    tau_pts = 1.0 / omega

    sort_idx = np.argsort(tau_pts)
    tau_pts = tau_pts[sort_idx]
    omega = omega[sort_idx]
    z_re = z_re[sort_idx]
    z_im = z_im[sort_idx]
    log_tau = np.log(tau_pts)

    # 1. RBF 파라미터
    coeff = 0.5
    delta_log_tau = np.abs(np.mean(np.diff(log_tau))) if n > 1 else 1.0
    epsilon = 2.0 * np.sqrt(np.log(2.0)) / (coeff * delta_log_tau)

    # 2. 적분 격자 (1500pts — 정밀도 충분, 메모리 절약)
    margin = 5.0 / epsilon
    N_int = 1500
    x_int = np.linspace(np.min(log_tau) - margin,
                        np.max(log_tau) + margin, N_int)
    dx = x_int[1] - x_int[0]

    w_exp_x = omega[:, None] * np.exp(x_int)[None, :]
    denom = 1.0 + w_exp_x ** 2
    re_kernel = 1.0 / denom
    im_kernel = w_exp_x / denom
    del w_exp_x, denom

    A_re = np.zeros((n, n))
    A_im = np.zeros((n, n))
    deriv_cache = np.empty((n, N_int))

    for k in range(n):
        diff_k = x_int - log_tau[k]
        rbf_k = np.exp(-(epsilon * diff_k) ** 2)
        deriv_cache[k] = -2.0 * epsilon ** 2 * diff_k * rbf_k
        A_re[:, k] = _trapz(re_kernel * rbf_k[None, :], dx=dx, axis=1)
        A_im[:, k] = _trapz(im_kernel * rbf_k[None, :], dx=dx, axis=1)

    del re_kernel, im_kernel

    M = np.zeros((n, n))
    for k in range(n):
        M[k, k:] = _trapz(deriv_cache[k] * deriv_cache[k:], dx=dx, axis=1)
        M[k:, k] = M[k, k:]

    del deriv_cache
    gc.collect()

    # 3. NNLS 조립
    eigvals, eigvecs = np.linalg.eigh(M)
    eigvals = np.maximum(eigvals, 0.0)
    U = np.diag(np.sqrt(eigvals)) @ eigvecs.T
    del eigvals, eigvecs, M

    if mode == 2:
        nv = n + 2
        K_re = np.zeros((n, nv)); K_re[:, 1] = 1.0; K_re[:, 2:] = A_re
        K_im = np.zeros((n, nv)); K_im[:, 0] = -omega; K_im[:, 2:] = A_im
        K_reg = np.zeros((n, nv)); K_reg[:, 2:] = np.sqrt(lam) * U
    else:
        nv = n + 1
        K_re = np.zeros((n, nv)); K_re[:, 0] = 1.0; K_re[:, 1:] = A_re
        K_im = np.zeros((n, nv)); K_im[:, 1:] = A_im
        K_reg = np.zeros((n, nv)); K_reg[:, 1:] = np.sqrt(lam) * U

    del A_re, A_im, U

    K_aug = np.vstack((K_re, K_im, K_reg))
    Z_aug = np.concatenate((z_re, z_im, np.zeros(n)))
    del K_re, K_im, K_reg
    gc.collect()

    # 4. NNLS
    x_opt, _ = nnls(K_aug, Z_aug)
    del K_aug, Z_aug

    gamma_weights = x_opt[2:] if mode == 2 else x_opt[1:]

    # 5. γ(τ) — MATLAB DRTtools 동일 fine grid
    N_fine = 500
    # 범위: log10(1/f) 기준 — MATLAB과 동일 (NOT log10(1/(2πf)))
    taumax = np.ceil(np.max(np.log10(1.0 / freq))) + 0.5
    taumin = np.floor(np.min(np.log10(1.0 / freq))) - 0.5
    # freq_fine: 내림차순 (MATLAB logspace(-taumin, -taumax, N) 동일)
    freq_fine = np.logspace(-taumin, -taumax, N_fine)

    # γ 평가: τ = 1/(2πf) — RBF 중심과 동일 convention
    # ln(tau_eval/tau_pts) = ln(f_k/f_fine) → 2π가 상쇄되어 값 동일
    tau_eval = 1.0 / (2.0 * np.pi * freq_fine)

    gamma_fine = np.zeros(N_fine)
    for m in range(n):
        if gamma_weights[m] == 0.0:
            continue
        gamma_fine += gamma_weights[m] * np.exp(
            -(epsilon * np.log(tau_eval / tau_pts[m])) ** 2
        )

    # 반환: τ = 1/f (MATLAB DRTtools export/plot convention)
    tau_plot = 1.0 / freq_fine

    gc.collect()
    return tau_plot.tolist(), gamma_fine.tolist()


def plot_15_graphs(freq, z_re, z_im):
    """5λ × 3mode = 15개 DRT 그래프."""
    freq = np.asarray(freq, dtype=float)
    z_re = np.asarray(z_re, dtype=float)
    z_im = np.asarray(z_im, dtype=float)

    lambdas = [1e-1, 1e-2, 1e-3, 1e-4, 1e-5]
    modes = [1, 2, 3]
    mode_names = ["w/o Induc", "with Induc", "Discard"]

    results = []
    fig, axes = plt.subplots(3, 5, figsize=(20, 11))
    print("⏳ DRT 계산 중...")

    idx = 1
    for m_idx, mode in enumerate(modes):
        for l_idx, lam in enumerate(lambdas):
            tau_fine, gamma_fine = solve_drt_core(freq, z_re, z_im, mode, lam)
            results.append((tau_fine, gamma_fine))

            ax = axes[m_idx, l_idx]
            ax.semilogx(tau_fine, gamma_fine, color='navy', linewidth=2.0)
            ax.set_title(f"#{idx}\n{mode_names[m_idx]}, λ={lam:.1e}", fontsize=11)
            ax.set_xlabel("τ (s)")
            ax.set_ylabel("γ(τ)")
            ax.grid(True, alpha=0.3)
            idx += 1

    plt.tight_layout()
    plt.show()
    gc.collect()
    return results
