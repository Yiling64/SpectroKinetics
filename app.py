# 依據樣品的 ch_idx 與側邊欄設定動態計算三個時間點
ch_offset = (ch_idx - 1) * t_interval
t_fmlp_cur = t_start + ch_offset
t_cb_cur = t_fmlp_cur - 180
t_drug_cur = t_fmlp_cur - 300

# 繪製動態虛線
ax.axvline(x=t_drug_cur, color="#94a3b8", linestyle=":", alpha=0.7)
ax.axvline(x=t_cb_cur, color="#94a3b8", linestyle=":", alpha=0.7)
ax.axvline(x=t_fmlp_cur, color="#94a3b8", linestyle=":", alpha=0.7)

# 動態標註（放在圖表內部頂部 y=0.98，避開與上方標題重疊）
ax.text(t_drug_cur, 0.98, f"Drug ({fmt_time(t_drug_cur)})", transform=ax.get_xaxis_transform(), fontsize=8.5, color="#475569", ha="center", va="top", fontweight="bold")
ax.text(t_cb_cur, 0.98, f"CB ({fmt_time(t_cb_cur)})", transform=ax.get_xaxis_transform(), fontsize=8.5, color="#475569", ha="center", va="top", fontweight="bold")
ax.text(t_fmlp_cur, 0.98, f"fMLF ({fmt_time(t_fmlp_cur)})", transform=ax.get_xaxis_transform(), fontsize=8.5, color="#475569", ha="center", va="top", fontweight="bold")
