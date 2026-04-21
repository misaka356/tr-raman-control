from __future__ import annotations

import json
import queue
import re
import threading
import time
import traceback
from pathlib import Path
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import pyvisa

from andor_sdk_integration import (
    AndorHardwareNotFoundError,
    AndorSDKController,
    AndorSDKError,
    DRV_TEMP_STABILIZED,
    ShamrockHardwareNotFoundError,
)
from tr_raman_integrated_controller import (
    analyze_ghost_peak,
    calculate_stress_mpa,
    ChannelConfig,
    DEFAULT_STRESS_CONSTANT_MPA_PER_CM1,
    DEFAULT_STRESS_FIT_MAX_CM1,
    DEFAULT_STRESS_FIT_MIN_CM1,
    evaluate_realtime_fit_quality,
    ExperimentPaused,
    fit_raman_peak_center,
    IntegratedExperimentConfig,
    run_continuous_preview_acquisition,
    SpectrometerConfig,
    request_pause_experiment,
    resolve_andor_sdk_root,
    run_baseline_test,
    run_integrated_experiment,
    start_andor_debug,
    stop_paused_experiment,
    test_andor_connection,
    start_generator_channel,
    start_generator_debug,
    stop_generator_channel,
    stop_generator_debug,
)


if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

ICON_PNG_PATH = APP_DIR / "tr_raman_icon.png"
ICON_ICO_PATH = APP_DIR / "tr_raman_icon.ico"


def find_path_candidates(filename: str) -> list[Path]:
    return [APP_DIR / filename, APP_DIR.parent / filename]


def find_existing_path(filename: str) -> Path | None:
    for path in find_path_candidates(filename):
        if path.exists():
            return path
    return None


def resolve_ui_sdk_root(value: str) -> str:
    text = value.strip()
    return str(resolve_andor_sdk_root(text))


class TRRamanUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("时间分辨拉曼控制")
        self.root.geometry("980x680")
        self.root.minsize(900, 620)
        self._icon_image = None
        self._apply_window_icon()

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.running = False
        self.continuous_preview_stop_event: threading.Event | None = None
        self.continuous_preview_active = False
        self.continuous_preview_button_var = tk.StringVar(value="连续采集")
        self.pause_button_var = tk.StringVar(value="暂停实验")
        self.pause_stage = "idle"
        self.channel_enabled_state = {1: False, 2: False}
        self.channel_toggle_buttons: dict[int, tk.Button] = {}
        self.channel_toggle_text: dict[int, tk.StringVar] = {
            1: tk.StringVar(value="打开 CH1"),
            2: tk.StringVar(value="打开 CH2"),
        }
        self.visa_combo: ttk.Combobox | None = None
        self.vars: dict[str, tk.StringVar] = {}
        self.offline_simulation_var = tk.BooleanVar(value=False)
        self.hidden_config: dict = {}
        self.temperature_text_var = tk.StringVar(value="--.- °C")
        self.temperature_status_var = tk.StringVar(value="未连接")
        self.temperature_poll_inflight = False
        self.temperature_monitor_controller: AndorSDKController | None = None
        self.andor_session_lock = threading.RLock()
        self.generator_status_var = tk.StringVar(value="未连接")
        self.generator_status_color = "#F53F3F"
        self.andor_status_var = tk.StringVar(value="未连接")
        self.andor_status_color = "#F53F3F"
        self.experiment_state_var = tk.StringVar(value="待机")
        self.experiment_state_color = "#1D2129"
        self.spectrum_countdown_var = tk.StringVar(value="--.- s")
        self.spectrum_peak_var = tk.StringVar(value="-- / --")
        self.spectrum_ghost_var = tk.StringVar(value="未检测")
        self.stress_baseline_path: Path | None = None
        self.stress_baseline_peak_cm1: float | None = None
        self.stress_zero_var = tk.StringVar(value="未选择")
        self.stress_current_var = tk.StringVar(value="-- MPa")
        self.stress_quality_var = tk.StringVar(value="Fit: --")
        self.experiment_stress_raw_points: list[tuple[float, int, float]] = []
        self.experiment_stress_phase_values: dict[float, list[float]] = {}
        self._configure_fonts()
        self._build_ui()
        self._ensure_config_exists()
        self.load_config()
        self.refresh_visa_resources()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(150, self._drain_log_queue)
        self.root.after(800, self._schedule_temperature_poll)

    def _apply_window_icon(self) -> None:
        try:
            if ICON_PNG_PATH.exists():
                self._icon_image = tk.PhotoImage(file=str(ICON_PNG_PATH))
                self.root.iconphoto(True, self._icon_image)
            elif ICON_ICO_PATH.exists():
                self.root.iconbitmap(str(ICON_ICO_PATH))
        except Exception:
            return

    def _configure_fonts(self) -> None:
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(size=11)
        text_font = tkfont.nametofont("TkTextFont")
        text_font.configure(size=11)
        heading_font = tkfont.nametofont("TkHeadingFont")
        heading_font.configure(size=11)
        fixed_font = tkfont.nametofont("TkFixedFont")
        fixed_font.configure(size=11)

        style = ttk.Style(self.root)
        style.configure(".", font=default_font)
        style.configure("TNotebook.Tab", padding=(10, 6), font=heading_font)
        style.configure("TButton", padding=(8, 5))
        style.configure("TLabelframe.Label", font=heading_font)
        style.configure("TEntry", padding=(4, 3))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=(10, 6, 10, 10))
        outer.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(outer)
        header.pack(fill=tk.X, pady=(0, 4))
        config_toolbar = ttk.Frame(header)
        config_toolbar.pack(side=tk.LEFT)
        ttk.Button(config_toolbar, text="读取配置", command=self.load_config).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(config_toolbar, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=6)
        ttk.Button(config_toolbar, text="另存为...", command=self.save_config_as).pack(side=tk.LEFT, padx=(6, 0))
        temp_frame = ttk.Frame(header)
        temp_frame.pack(side=tk.RIGHT)
        ttk.Label(temp_frame, textvariable=self.temperature_status_var).pack(side=tk.TOP, anchor="e")
        self.temperature_badge = tk.Label(
            temp_frame,
            textvariable=self.temperature_text_var,
            width=12,
            padx=12,
            pady=4,
            font=("TkDefaultFont", 16, "bold"),
            fg="white",
            bg="#7f8c8d",
            relief=tk.GROOVE,
            bd=2,
        )
        self.temperature_badge.pack(side=tk.TOP, anchor="e", pady=(2, 0))
        status_strip = ttk.Frame(header)
        status_strip.pack(side=tk.RIGHT, padx=(0, 12))
        self.generator_status_value_label = self._create_status_card(status_strip, "信号发生器", self.generator_status_var, self.generator_status_color)
        self.andor_status_value_label = self._create_status_card(status_strip, "Andor", self.andor_status_var, self.andor_status_color)
        self.experiment_status_value_label = self._create_status_card(status_strip, "实验状态", self.experiment_state_var, self.experiment_state_color)

        notebook = ttk.Notebook(outer)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 0))

        tab_general = ttk.Frame(notebook, padding=10)
        tab_generator = ttk.Frame(notebook, padding=10)
        tab_spectrometer = ttk.Frame(notebook, padding=10)
        tab_run = ttk.Frame(notebook, padding=10)

        notebook.add(tab_general, text="基本设置")
        notebook.add(tab_generator, text="信号发生器")
        notebook.add(tab_spectrometer, text="光谱仪")
        notebook.add(tab_run, text="实验运行")

        self._build_general_tab(tab_general)
        self._build_generator_tab(tab_generator)
        self._build_spectrometer_tab(tab_spectrometer)
        self._build_run_tab(tab_run)

    def _create_status_card(self, parent, title: str, value_var: tk.StringVar, color: str) -> tk.Label:
        card = tk.Frame(parent, bg="#56646C", bd=2, relief=tk.GROOVE, padx=12, pady=6)
        card.pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(card, text=title, font=("TkDefaultFont", 10), fg="#D7DEE3", bg="#56646C").pack(anchor="w")
        value_label = tk.Label(
            card,
            textvariable=value_var,
            font=("TkDefaultFont", 13, "bold"),
            fg=color,
            bg="#56646C",
            width=8,
            anchor="w",
        )
        value_label.pack(anchor="w", pady=(4, 0))
        return value_label

    def _set_status_card(self, kind: str, text: str, color: str) -> None:
        if kind == "generator":
            self.generator_status_var.set(text)
            self.generator_status_value_label.configure(fg=color)
        elif kind == "andor":
            self.andor_status_var.set(text)
            self.andor_status_value_label.configure(fg=color)
        elif kind == "experiment":
            self.experiment_state_var.set(text)
            self.experiment_status_value_label.configure(fg=color)

    def _read_spectrum_file(self, path: Path) -> tuple[list[float], list[float]]:
        x_values: list[float] = []
        y_values: list[float] = []
        for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
            parts = line.strip().split()
            if len(parts) < 2:
                continue
            try:
                x_values.append(float(parts[0]))
                y_values.append(float(parts[1]))
            except ValueError:
                continue
        return x_values, y_values

    def select_stress_zero_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择零点光谱文件",
            filetypes=[("ASC 光谱文件", "*.asc"), ("所有文件", "*.*")],
        )
        if not filename:
            return
        path = Path(filename)
        try:
            x_values, y_values = self._read_spectrum_file(path)
            fit = fit_raman_peak_center(
                x_values,
                y_values,
                DEFAULT_STRESS_FIT_MIN_CM1,
                DEFAULT_STRESS_FIT_MAX_CM1,
            )
            quality = evaluate_realtime_fit_quality(fit)
            if not quality["reliable"]:
                raise ValueError(f"zero-point fit is unreliable: {quality['message']}")
        except Exception as exc:
            messagebox.showerror("零点拟合失败", str(exc))
            return

        self.stress_baseline_path = path
        self.stress_baseline_peak_cm1 = float(fit["center_cm1"])
        self.stress_zero_var.set(f"{path.name} / {self.stress_baseline_peak_cm1:.4f} cm^-1")
        self.stress_current_var.set("等待采集")
        self.stress_quality_var.set(f"Zero fit: {quality['level']} ({quality['message']})")
        self._append_log(
            f"应力零点已设置: {path.name}, 峰位 {self.stress_baseline_peak_cm1:.6f} cm^-1"
        )
        self._refresh_preview_from_path(path)

    def _legacy_update_stress_from_spectrum(self, x_values: list[float], y_values: list[float], path: Path) -> None:
        if self.stress_baseline_peak_cm1 is None:
            self.stress_current_var.set("未选择零点")
            return
        try:
            fit = fit_raman_peak_center(
                x_values,
                y_values,
                DEFAULT_STRESS_FIT_MIN_CM1,
                DEFAULT_STRESS_FIT_MAX_CM1,
            )
            current_peak = float(fit["center_cm1"])
            stress_mpa = calculate_stress_mpa(
                self.stress_baseline_peak_cm1,
                current_peak,
                DEFAULT_STRESS_CONSTANT_MPA_PER_CM1,
            )
            self.stress_current_var.set(f"{stress_mpa:.2f} MPa ({current_peak:.4f} cm^-1)")
        except Exception as exc:
            self.stress_current_var.set("拟合失败")
            self._append_log(f"实时应力计算失败: {path.name}: {exc}")

    def _calculate_realtime_stress_result(
        self,
        x_values: list[float],
        y_values: list[float],
        ghost_analysis: dict | None = None,
    ) -> dict:
        if self.stress_baseline_peak_cm1 is None:
            return {"display_state": "no_baseline", "message": "未选择零点", "valid_for_trend": False}
        fit = fit_raman_peak_center(
            x_values,
            y_values,
            DEFAULT_STRESS_FIT_MIN_CM1,
            DEFAULT_STRESS_FIT_MAX_CM1,
        )
        quality = evaluate_realtime_fit_quality(fit)
        risk_level = (ghost_analysis or {}).get("risk_level", "none")
        risk_type = (ghost_analysis or {}).get("warning_type", "none")
        current_peak = float(fit["center_cm1"])
        stress_mpa = calculate_stress_mpa(
            self.stress_baseline_peak_cm1,
            current_peak,
            DEFAULT_STRESS_CONSTANT_MPA_PER_CM1,
        )
        if not quality["reliable"]:
            display_state = "hidden_fit_unreliable"
            message = f"Fit reject: {quality['message']}"
        elif risk_level == "high" and quality["level"] != "good":
            display_state = "hidden_high_ghost_plus_fit"
            message = f"High ghost risk + fit {quality['level']}: {risk_type}"
        elif risk_level == "high":
            display_state = "shown_high_ghost_reference"
            message = f"Fit {quality['level']}; ghost {risk_level}: {risk_type}"
        elif risk_level == "low" or quality["level"] == "caution":
            display_state = "shown_with_caution"
            message = (
                f"Fit {quality['level']}; ghost {risk_level}: {risk_type}"
                if risk_level == "low"
                else f"Fit {quality['level']}: {quality['message']}"
            )
        else:
            display_state = "shown"
            message = f"Fit {quality['level']}: {quality['message']}"

        return {
            "display_state": display_state,
            "valid_for_trend": display_state in ("shown", "shown_with_caution"),
            "stress_mpa": stress_mpa,
            "current_peak": current_peak,
            "fit": fit,
            "quality": quality,
            "risk_level": risk_level,
            "risk_type": risk_type,
            "message": message,
        }

    def _update_stress_from_spectrum(
        self,
        x_values: list[float],
        y_values: list[float],
        path: Path,
        ghost_analysis: dict | None = None,
    ) -> None:
        if self.stress_baseline_peak_cm1 is None:
            self.stress_current_var.set("未选择零点")
            self.stress_quality_var.set("Fit: waiting for zero")
            return
        try:
            fit = fit_raman_peak_center(
                x_values,
                y_values,
                DEFAULT_STRESS_FIT_MIN_CM1,
                DEFAULT_STRESS_FIT_MAX_CM1,
            )
            quality = evaluate_realtime_fit_quality(fit)
            risk_level = (ghost_analysis or {}).get("risk_level", "none")
            risk_type = (ghost_analysis or {}).get("warning_type", "none")

            if not quality["reliable"]:
                self.stress_current_var.set("不显示: 拟合不可靠")
                self.stress_quality_var.set(f"Fit reject: {quality['message']}")
                self._append_log(f"{path.name}: realtime stress hidden, unreliable fit: {quality['message']}")
                return
            if risk_level == "high" and quality["level"] != "good":
                self.stress_current_var.set("不显示: 高风险鬼峰")
                self.stress_quality_var.set(f"High ghost risk + fit {quality['level']}: {risk_type}")
                self._append_log(
                    f"{path.name}: realtime stress hidden, high ghost risk ({risk_type}) with fit {quality['level']}"
                )
                return

            current_peak = float(fit["center_cm1"])
            stress_mpa = calculate_stress_mpa(
                self.stress_baseline_peak_cm1,
                current_peak,
                DEFAULT_STRESS_CONSTANT_MPA_PER_CM1,
            )
            suffix = ""
            if risk_level == "high":
                suffix = " / 高风险仅供参考"
            elif risk_level == "low" or quality["level"] == "caution":
                suffix = " / 注意"
            self.stress_current_var.set(f"{stress_mpa:.2f} MPa ({current_peak:.4f} cm^-1){suffix}")
            if risk_level in ("low", "high"):
                self.stress_quality_var.set(f"Fit {quality['level']}; ghost {risk_level}: {risk_type}")
            else:
                self.stress_quality_var.set(f"Fit {quality['level']}: {quality['message']}")
        except Exception as exc:
            self.stress_current_var.set("拟合失败")
            self.stress_quality_var.set("Fit: failed")
            self._append_log(f"实时应力计算失败: {path.name}: {exc}")

    def _draw_spectrum_preview(self, x_values: list[float], y_values: list[float], title: str, analysis: dict | None = None) -> None:
        canvas = self.preview_canvas
        canvas.delete("all")
        width = int(canvas.winfo_width() or 860)
        height = int(canvas.winfo_height() or 300)
        canvas.create_rectangle(0, 0, width, height, fill="#1F2329", outline="")
        if len(x_values) < 2 or len(y_values) < 2:
            canvas.create_text(width / 2, height / 2, text="暂无可显示的采集数据", fill="#A9B4BE", font=("TkDefaultFont", 12))
            return

        left = 58
        right = 18
        top = 18
        bottom = 40
        plot_w = max(10, width - left - right)
        plot_h = max(10, height - top - bottom)

        min_x_data, max_x_data = min(x_values), max(x_values)
        min_y, max_y = min(y_values), max(y_values)
        half_span = max(abs(min_x_data - 532.0), abs(max_x_data - 532.0))
        if half_span <= 0:
            half_span = 1.0
        min_x = 532.0 - half_span
        max_x = 532.0 + half_span
        if max_y == min_y:
            max_y += 1.0

        canvas.create_line(left, top, left, height - bottom, fill="#697586", width=1)
        canvas.create_line(left, height - bottom, width - right, height - bottom, fill="#697586", width=1)

        for frac in (0.0, 0.5, 1.0):
            y = top + plot_h * (1.0 - frac)
            y_value = min_y + (max_y - min_y) * frac
            canvas.create_line(left, y, width - right, y, fill="#2C323A", width=1)
            canvas.create_text(left - 8, y, text=f"{y_value:.0f}", fill="#D5DDE6", anchor="e", font=("TkDefaultFont", 9))

        for frac in (0.0, 0.5, 1.0):
            x = left + plot_w * frac
            x_value = min_x + (max_x - min_x) * frac
            canvas.create_text(x, height - bottom + 16, text=f"{x_value:.0f}", fill="#D5DDE6", anchor="n", font=("TkDefaultFont", 9))

        points: list[float] = []
        for x, y in zip(x_values, y_values):
            if x < min_x or x > max_x:
                continue
            px = left + (x - min_x) / (max_x - min_x) * plot_w
            py = top + (1.0 - (y - min_y) / (max_y - min_y)) * plot_h
            points.extend([px, py])
        if len(points) >= 4:
            canvas.create_line(*points, fill="#F3B34C", width=2)
        if analysis:
            primary_peak = analysis.get("primary_peak")
            secondary_peak = analysis.get("secondary_peak")
            if primary_peak:
                px = left + (primary_peak["x"] - min_x) / (max_x - min_x) * plot_w
                py = top + (1.0 - (primary_peak["y"] - min_y) / (max_y - min_y)) * plot_h
                canvas.create_oval(px - 3, py - 3, px + 3, py + 3, fill="#4CD964", outline="")
            if secondary_peak:
                px = left + (secondary_peak["x"] - min_x) / (max_x - min_x) * plot_w
                py = top + (1.0 - (secondary_peak["y"] - min_y) / (max_y - min_y)) * plot_h
                canvas.create_oval(px - 3, py - 3, px + 3, py + 3, fill="#FF5A5F", outline="")
            risk_level = analysis.get("risk_level", "high" if analysis.get("ghost_detected") else "none")
            if risk_level == "high":
                canvas.create_rectangle(width - 320, 12, width - 14, 42, fill="#5A1E24", outline="#FF5A5F")
                canvas.create_text(width - 167, 27, text=analysis.get("message", "鬼峰警告"), fill="#FFD5D8", font=("TkDefaultFont", 10, "bold"))

        canvas.create_text(width / 2, height - 8, text="Raman shift (cm⁻¹)", fill="#E5EAF0", anchor="s", font=("TkDefaultFont", 10))
        canvas.create_text(18, height / 2, text="Counts", fill="#E5EAF0", angle=90, font=("TkDefaultFont", 10))

    def _refresh_preview_from_path(self, path: Path) -> None:
        try:
            x_values, y_values = self._read_spectrum_file(path)
            analysis = analyze_ghost_peak(x_values, [int(v) for v in y_values], 460.0, 580.0)
            self._draw_spectrum_preview(x_values, y_values, path.name, analysis)
            primary_peak = analysis.get("primary_peak")
            if primary_peak:
                self.spectrum_peak_var.set(f"{primary_peak['x']:.2f} cm^-1 / {primary_peak['y']:.0f}")
            else:
                self.spectrum_peak_var.set("-- / --")
            risk_level = analysis.get("risk_level", "high" if analysis.get("ghost_detected") else "none")
            if risk_level in ("low", "high"):
                self.spectrum_ghost_var.set("检测到鬼峰")
                self.spectrum_ghost_label.configure(fg="#F53F3F")
                if risk_level == "low":
                    self.spectrum_ghost_var.set(f"低风险: {analysis.get('warning_type', 'ghost')}")
                    self.spectrum_ghost_label.configure(fg="#D46B08")
                else:
                    self.spectrum_ghost_var.set(f"高风险: {analysis.get('warning_type', 'ghost')}")
                self._append_log(f"{path.name}: {analysis['message']}")
            else:
                self.spectrum_ghost_var.set("未检测到")
                self.spectrum_ghost_label.configure(fg="#00B42A")
            self._update_stress_from_spectrum(x_values, y_values, path, analysis)
        except Exception as exc:
            self._append_log(f"预览刷新失败: {exc}")

    def _refresh_preview_from_latest_file(self, config: IntegratedExperimentConfig) -> None:
        sample_dir = Path(config.output_dir) / config.sample_name
        if not sample_dir.exists():
            return
        candidates = sorted(sample_dir.glob("*.asc"), key=lambda p: p.stat().st_mtime, reverse=True)
        if candidates:
            self._refresh_preview_from_path(candidates[0])

    def _reset_experiment_stress_trend(self) -> None:
        self.experiment_stress_raw_points.clear()
        self.experiment_stress_phase_values.clear()
        self._draw_experiment_stress_trend()

    def _parse_experiment_phase_repeat(self, path: Path) -> tuple[float, int] | None:
        if path.name.startswith("-") or path.name.startswith("_") or path.name.startswith("debug-"):
            return None
        match = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)-(\d+)(?:_ghost)?(?:-*)\.asc", path.name)
        if match is None:
            return None
        return float(match.group(1)), int(match.group(2))

    def _refresh_preview_and_update_experiment_trend(self, path: Path) -> None:
        self._refresh_preview_from_path(path)
        parsed = self._parse_experiment_phase_repeat(path)
        if parsed is None:
            return
        try:
            x_values, y_values = self._read_spectrum_file(path)
            analysis = analyze_ghost_peak(x_values, [int(v) for v in y_values], 460.0, 580.0)
            result = self._calculate_realtime_stress_result(x_values, y_values, analysis)
        except Exception as exc:
            self._append_log(f"实验应力趋势更新失败: {path.name}: {exc}")
            return
        if not result.get("valid_for_trend"):
            self._append_log(f"{path.name}: 未加入正式实验应力趋势 ({result.get('message', 'invalid stress')})")
            return
        phase, repeat = parsed
        stress_mpa = float(result["stress_mpa"])
        self.experiment_stress_raw_points.append((phase, repeat, stress_mpa))
        self.experiment_stress_phase_values.setdefault(phase, []).append(stress_mpa)
        self._draw_experiment_stress_trend()

    def _draw_experiment_stress_trend(self) -> None:
        canvas = getattr(self, "experiment_stress_canvas", None)
        if canvas is None:
            return
        canvas.delete("all")
        width = int(canvas.winfo_width() or 320)
        height = int(canvas.winfo_height() or 142)
        canvas.create_rectangle(0, 0, width, height, fill="#1F2329", outline="")
        left, right, top, bottom = 42, 12, 18, 28
        plot_w = max(10, width - left - right)
        plot_h = max(10, height - top - bottom)
        canvas.create_text(left, 5, text="实时应力趋势", fill="#E5EAF0", anchor="nw", font=("TkDefaultFont", 9, "bold"))
        if not self.experiment_stress_raw_points:
            canvas.create_text(width / 2, height / 2, text="等待有效应力点", fill="#A9B4BE", font=("TkDefaultFont", 10))
            return

        mean_points = [
            (phase, sum(values) / len(values))
            for phase, values in sorted(self.experiment_stress_phase_values.items())
            if values
        ]
        phases = [p for p, _, _ in self.experiment_stress_raw_points] + [p for p, _ in mean_points]
        stresses = [s for _, _, s in self.experiment_stress_raw_points] + [s for _, s in mean_points]
        min_x, max_x = min(phases), max(phases)
        min_y, max_y = min(stresses), max(stresses)
        if max_x == min_x:
            min_x -= 1.0
            max_x += 1.0
        if max_y == min_y:
            min_y -= 1.0
            max_y += 1.0
        y_pad = max(1.0, (max_y - min_y) * 0.12)
        min_y -= y_pad
        max_y += y_pad

        def to_px(phase: float, stress: float) -> tuple[float, float]:
            px = left + (phase - min_x) / (max_x - min_x) * plot_w
            py = top + (1.0 - (stress - min_y) / (max_y - min_y)) * plot_h
            return px, py

        canvas.create_line(left, top, left, height - bottom, fill="#697586")
        canvas.create_line(left, height - bottom, width - right, height - bottom, fill="#697586")
        for frac in (0.0, 0.5, 1.0):
            y = top + plot_h * (1.0 - frac)
            value = min_y + (max_y - min_y) * frac
            canvas.create_line(left, y, width - right, y, fill="#2C323A")
            canvas.create_text(left - 5, y, text=f"{value:.0f}", fill="#D5DDE6", anchor="e", font=("TkDefaultFont", 8))
        canvas.create_text(left, height - 5, text=f"{min_x:g}°", fill="#D5DDE6", anchor="sw", font=("TkDefaultFont", 8))
        canvas.create_text(width - right, height - 5, text=f"{max_x:g}°", fill="#D5DDE6", anchor="se", font=("TkDefaultFont", 8))

        for phase, _repeat, stress in self.experiment_stress_raw_points:
            px, py = to_px(phase, stress)
            canvas.create_oval(px - 2, py - 2, px + 2, py + 2, fill="#A9B4BE", outline="")

        if len(mean_points) >= 2:
            line_points: list[float] = []
            for phase, stress in mean_points:
                px, py = to_px(phase, stress)
                line_points.extend([px, py])
            canvas.create_line(*line_points, fill="#4CD964", width=2)
        for phase, stress in mean_points:
            px, py = to_px(phase, stress)
            canvas.create_oval(px - 3, py - 3, px + 3, py + 3, fill="#4CD964", outline="")

        last_phase, last_stress = mean_points[-1]
        canvas.create_text(
            width - right,
            top + 2,
            text=f"{last_phase:g}° / {last_stress:.1f} MPa",
            fill="#E5EAF0",
            anchor="ne",
            font=("TkDefaultFont", 8),
        )

    def _new_var(self, key: str, default: str = "") -> tk.StringVar:
        var = tk.StringVar(value=default)
        self.vars[key] = var
        return var

    def _build_runtime_config_quiet(self) -> IntegratedExperimentConfig | None:
        try:
            return self._build_runtime_config()
        except Exception:
            return None

    def _set_temperature_badge(self, text: str, status_text: str, color: str) -> None:
        self.temperature_text_var.set(text)
        self.temperature_status_var.set(status_text)
        self.temperature_badge.configure(bg=color, activebackground=color)

    def _close_temperature_monitor(self) -> None:
        controller = self.temperature_monitor_controller
        self.temperature_monitor_controller = None
        if controller is not None:
            try:
                controller.close()
            except Exception:
                pass

    def _on_close(self) -> None:
        if self.continuous_preview_stop_event is not None:
            self.continuous_preview_stop_event.set()
        self._close_temperature_monitor()
        self.root.destroy()

    def _schedule_temperature_poll(self) -> None:
        if not self.root.winfo_exists():
            return
        if not self.temperature_poll_inflight:
            self._start_temperature_poll()
        self.root.after(5000, self._schedule_temperature_poll)

    def _start_temperature_poll(self) -> None:
        cfg = self._build_runtime_config_quiet()
        if cfg is None or cfg.offline_simulation or self.running:
            return
        self.temperature_poll_inflight = True

        def worker() -> None:
            try:
                with AndorSDKController(cfg.andor_sdk_root) as andor:
                    andor.enable_cooler_and_set_target(cfg.spectrometer.target_temperature_c)
                    temp_c, status_code = andor.get_temperature_c()
                    stable = temp_c <= cfg.spectrometer.required_temperature_c and status_code == DRV_TEMP_STABILIZED
                    color = "#1f3dff" if stable else "#d84b4b"
                    status = "可开始实验" if stable else "制冷中"
                    self.root.after(
                        0,
                        lambda t=temp_c, s=status, c=color: (
                            self._set_temperature_badge(f"{t:.1f} °C", s, c),
                            self._set_status_card("andor", "已连接", "#00B42A"),
                        ),
                    )
            except (AndorHardwareNotFoundError, ShamrockHardwareNotFoundError):
                self.root.after(
                    0,
                    lambda: (
                        self._set_temperature_badge("--.- °C", "未连接", "#7f8c8d"),
                        self._set_status_card("andor", "未连接", "#F53F3F"),
                    ),
                )
            except (AndorSDKError, FileNotFoundError) as exc:
                self.root.after(
                    0,
                    lambda e=str(exc): (
                        self._set_temperature_badge("--.- °C", "驱动异常", "#7f8c8d"),
                        self._set_status_card("andor", "异常", "#FF7D00"),
                    ),
                )
                self._queue_log(f"Andor 温度监视失败: {exc}")
            finally:
                self.temperature_poll_inflight = False

        threading.Thread(target=worker, daemon=True).start()

    def _add_labeled_entry(
        self,
        parent,
        row: int,
        label: str,
        key: str,
        default: str = "",
        width: int = 12,
        browse: str | None = None,
        stretch: bool = False,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=4, pady=4)
        entry = ttk.Entry(parent, textvariable=self._new_var(key, default), width=width)
        entry.grid(row=row, column=1, sticky="ew" if stretch else "w", padx=4, pady=4)
        if browse == "file":
            ttk.Button(parent, text="浏览", command=lambda k=key: self._browse_file(k)).grid(
                row=row, column=2, padx=4, pady=4
            )
        elif browse == "dir":
            ttk.Button(parent, text="浏览", command=lambda k=key: self._browse_dir(k)).grid(
                row=row, column=2, padx=4, pady=4
            )

    def _add_labeled_combo(
        self,
        parent,
        row: int,
        label: str,
        key: str,
        values: list[str],
        default: str = "",
        width: int = 12,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=4, pady=4)
        combo = ttk.Combobox(
            parent,
            textvariable=self._new_var(key, default or values[0]),
            values=values,
            width=width,
            state="readonly",
        )
        combo.grid(row=row, column=1, sticky="w", padx=4, pady=4)

    def _build_general_tab(self, parent) -> None:
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(2, weight=0)
        self._add_labeled_entry(
            parent,
            0,
            "Andor SDK 目录",
            "andor_sdk_root",
            str(resolve_andor_sdk_root()),
            browse="dir",
            width=42,
            stretch=True,
        )
        ttk.Label(parent, text="RIGOL VISA 资源").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.visa_combo = ttk.Combobox(parent, textvariable=self._new_var("rigol_visa_resource", ""), width=40)
        self.visa_combo.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(parent, text="扫描 VISA", command=self.refresh_visa_resources).grid(row=1, column=2, padx=4, pady=4)

        self._add_labeled_entry(parent, 2, "输出目录", "output_dir", r"C:\AndorOutput", browse="dir", width=42, stretch=True)
        self._add_labeled_entry(parent, 3, "样品名称", "sample_name", "sample", width=12)
        self._add_labeled_entry(parent, 4, "CH1 相对延时 (ms)", "ch1_start_delay_ms", "0.0", width=12)
        self._add_labeled_entry(parent, 5, "稳定等待时间 (ms)", "settle_time_ms", "200.0", width=12)

        ttk.Checkbutton(parent, text="启用离线模拟模式（不连接任何硬件）", variable=self.offline_simulation_var).grid(
            row=6, column=0, columnspan=3, sticky="w", padx=4, pady=8
        )

    def _build_generator_tab(self, parent) -> None:
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        left = ttk.LabelFrame(parent, text="CH1 声光调制器", padding=10)
        right = ttk.LabelFrame(parent, text="CH2 音圈激振", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        right.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)

        left.columnconfigure(1, weight=1)
        right.columnconfigure(1, weight=1)

        self._add_labeled_combo(left, 0, "波形", "ch1_waveform", ["sin", "square", "pulse", "ramp", "noise", "dc"], "square", width=12)
        self._add_labeled_entry(left, 1, "频率 (Hz)", "ch1_frequency_hz", "1000.0", width=12)
        self._add_labeled_entry(left, 2, "幅值 (Vpp)", "ch1_amplitude_vpp", "5.0", width=12)
        self._add_labeled_entry(left, 3, "偏置 (Vdc)", "ch1_offset_vdc", "0.0", width=12)
        self._add_labeled_entry(left, 4, "相位 (deg)", "ch1_phase_deg", "0.0", width=12)
        self._add_labeled_entry(left, 5, "占空比 (%)", "ch1_duty_cycle_percent", "50.0", width=12)
        ch1_toggle = tk.Button(left, textvariable=self.channel_toggle_text[1], width=12, command=lambda: self.run_generator_channel_toggle(1))
        ch1_toggle.grid(row=6, column=1, sticky="w", padx=4, pady=(8, 4))
        self.channel_toggle_buttons[1] = ch1_toggle
        self._set_channel_toggle_visual(1, False)

        self._add_labeled_combo(right, 0, "波形", "ch2_waveform", ["sin", "square", "pulse", "ramp", "noise", "dc"], "sin", width=12)
        self._add_labeled_entry(right, 1, "频率 (Hz)", "ch2_frequency_hz", "1000.0", width=12)
        self._add_labeled_entry(right, 2, "幅值 (Vpp)", "ch2_amplitude_vpp", "1.0", width=12)
        self._add_labeled_entry(right, 3, "偏置 (Vdc)", "ch2_offset_vdc", "0.0", width=12)
        self._add_labeled_entry(right, 4, "相位 (deg)", "ch2_phase_deg", "0.0", width=12)
        self._add_labeled_entry(right, 5, "占空比 (%)", "ch2_duty_cycle_percent", "50.0", width=12)
        ch2_toggle = tk.Button(right, textvariable=self.channel_toggle_text[2], width=12, command=lambda: self.run_generator_channel_toggle(2))
        ch2_toggle.grid(row=6, column=1, sticky="w", padx=4, pady=(8, 4))
        self.channel_toggle_buttons[2] = ch2_toggle
        self._set_channel_toggle_visual(2, False)

        controls = ttk.LabelFrame(parent, text="发生器调试控制", padding=10)
        controls.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=8)

        ttk.Button(controls, text="启动双通道", command=self.run_generator_debug).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(controls, text="停止发生器", command=self.run_generator_stop).grid(row=0, column=1, padx=4, pady=4)

    def _build_spectrometer_tab(self, parent) -> None:
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(2, weight=0)
        self._add_labeled_entry(parent, 0, "瑞利波长 (nm)", "rayleigh_wavelength_nm", "532.0", width=12)
        self._add_labeled_entry(parent, 1, "中心波长 (nm)", "center_wavelength_nm", "500.0", width=12)
        self._add_labeled_entry(parent, 2, "光栅编号", "grating_no", "3", width=12)
        self._add_labeled_entry(parent, 3, "曝光时间 (ms)", "exposure_ms", "200.0", width=12)
        self._add_labeled_entry(parent, 4, "Andor 触发模式", "trigger_mode", "1", width=12)

        info = ttk.LabelFrame(parent, text="实时数据", padding=10)
        info.grid(row=0, column=2, rowspan=5, sticky="ne", padx=(16, 4), pady=4)
        info.columnconfigure(2, minsize=24)
        ttk.Label(info, text="曝光倒计时").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Label(info, textvariable=self.spectrum_countdown_var).grid(row=0, column=1, sticky="w", padx=(12, 0), pady=4)
        ttk.Label(info, text="当前峰").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Label(info, textvariable=self.spectrum_peak_var).grid(row=1, column=1, sticky="w", padx=(12, 0), pady=4)
        ttk.Label(info, text="鬼峰检测").grid(row=2, column=0, sticky="w", pady=4)
        self.spectrum_ghost_label = tk.Label(info, textvariable=self.spectrum_ghost_var, anchor="w", bg=self.root.cget("bg"), fg="#1D2129")
        self.spectrum_ghost_label.grid(row=2, column=1, sticky="w", padx=(12, 0), pady=4)
        ttk.Button(info, text="选择零点", command=self.select_stress_zero_file).grid(row=0, column=3, sticky="w", padx=(8, 0), pady=4)
        ttk.Label(info, textvariable=self.stress_zero_var, width=24).grid(row=1, column=3, sticky="w", padx=(8, 0), pady=4)
        ttk.Label(info, textvariable=self.stress_current_var).grid(row=2, column=3, sticky="w", padx=(8, 0), pady=4)
        ttk.Label(info, textvariable=self.stress_quality_var, width=34).grid(row=3, column=3, sticky="w", padx=(8, 0), pady=4)

        controls = ttk.LabelFrame(parent, text="光谱仪调试", padding=10)
        controls.grid(row=5, column=0, columnspan=3, sticky="ew", padx=4, pady=10)
        ttk.Button(controls, text="光谱仪采集", command=self.run_andor_debug).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(controls, textvariable=self.continuous_preview_button_var, command=self.toggle_continuous_preview).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(controls, text="连接测试", command=self.run_andor_connection_test).grid(row=0, column=2, padx=4, pady=4)
        preview = ttk.Frame(parent)
        preview.grid(row=6, column=0, columnspan=3, sticky="nsew", padx=4, pady=(4, 8))
        parent.rowconfigure(6, weight=1)
        self.preview_canvas = tk.Canvas(preview, height=420, bg="#1F2329", highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.root.after(0, lambda: self._draw_spectrum_preview([], [], ""))
        self.spectrum_countdown_var.set("--.- s")
        self.spectrum_peak_var.set("-- / --")
        self.spectrum_ghost_var.set("未检测")
        self.stress_quality_var.set("Fit: --")

    def _build_run_tab(self, parent) -> None:
        parent.columnconfigure(0, weight=1)

        exp_params = ttk.LabelFrame(parent, text="实验扫描参数", padding=10)
        exp_params.pack(fill=tk.X, pady=(0, 8))
        exp_params.columnconfigure(1, weight=1)
        exp_params.columnconfigure(2, minsize=340)
        self._add_labeled_entry(exp_params, 0, "起始相位 (deg)", "phase_start_deg", "0.0", width=12)
        self._add_labeled_entry(exp_params, 1, "终止相位 (deg)", "phase_stop_deg", "360.0", width=12)
        self._add_labeled_entry(exp_params, 2, "相位步长 (deg)", "phase_step_deg", "10.0", width=12)
        self._add_labeled_entry(exp_params, 3, "每个相位重复次数", "repeats_per_phase", "3", width=12)
        trend_frame = ttk.Frame(exp_params)
        trend_frame.grid(row=0, column=2, rowspan=4, sticky="nsew", padx=(16, 0), pady=0)
        self.experiment_stress_canvas = tk.Canvas(trend_frame, width=320, height=142, bg="#1F2329", highlightthickness=0)
        self.experiment_stress_canvas.pack(fill=tk.BOTH, expand=False)
        self._draw_experiment_stress_trend()

        toolbar = ttk.Frame(parent)
        toolbar.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(toolbar, text="基线测试", command=self.run_baseline_test).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="开始/继续实验", command=self.run_experiment).pack(side=tk.LEFT, padx=12)
        ttk.Button(toolbar, textvariable=self.pause_button_var, command=self.pause_experiment).pack(side=tk.LEFT, padx=4)

        self.status_var = tk.StringVar(value="空闲")
        ttk.Label(toolbar, textvariable=self.status_var).pack(side=tk.RIGHT, padx=4)

        progress_row = ttk.Frame(parent)
        progress_row.pack(fill=tk.X, pady=(0, 8))
        self.progress_text_var = tk.StringVar(value="等待曝光")
        ttk.Label(progress_row, textvariable=self.progress_text_var).pack(side=tk.LEFT, padx=(0, 8))
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(
            progress_row,
            orient=tk.HORIZONTAL,
            mode="determinate",
            maximum=100.0,
            variable=self.progress_var,
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.log_box = ScrolledText(parent, wrap=tk.WORD, height=28)
        self.log_box.pack(fill=tk.BOTH, expand=True)
        self.log_box.configure(state=tk.DISABLED)

    def _browse_file(self, key: str) -> None:
        path = filedialog.askopenfilename()
        if path:
            self.vars[key].set(path)

    def _browse_dir(self, key: str) -> None:
        path = filedialog.askdirectory()
        if path:
            self.vars[key].set(path)

    def refresh_visa_resources(self) -> None:
        resources: list[str] = []
        try:
            rm = pyvisa.ResourceManager()
            resources = list(rm.list_resources())
            rm.close()
        except Exception as exc:
            self._append_log(f"VISA 扫描失败: {exc}")
        if self.visa_combo is not None:
            self.visa_combo["values"] = resources
        current = self.vars["rigol_visa_resource"].get().strip()
        if resources and current not in resources:
            self.vars["rigol_visa_resource"].set(resources[0])
        self._set_status_card("generator", "未连接" if not resources else "待连接", "#F53F3F" if not resources else "#FF7D00")
        self._append_log(f"VISA 资源: {resources if resources else '未找到'}")

    def _ensure_config_exists(self) -> None:
        cfg = find_existing_path("app_config.json")
        if cfg is not None:
            return
        template = find_existing_path("app_config.template.json")
        target = find_path_candidates("app_config.json")[0]
        if template is not None:
            target.write_text(template.read_text(encoding="utf-8-sig"), encoding="utf-8")

    def _append_log(self, text: str) -> None:
        self.log_box.configure(state=tk.NORMAL)
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.log_box.configure(state=tk.DISABLED)

    def _queue_log(self, text: str) -> None:
        self.log_queue.put(text)

    def _update_spectrum_info_from_progress(self, value: float, text: str = "") -> None:
        text = text or ""
        exposure_ms = 0.0
        try:
            exposure_ms = float(self.vars.get("exposure_ms").get()) if self.vars.get("exposure_ms") is not None else 0.0
        except Exception:
            exposure_ms = 0.0
        total_s = max(0.0, exposure_ms / 1000.0)
        if 0.0 <= float(value) <= 1.0 and total_s > 0:
            remain_s = total_s * max(0.0, 1.0 - float(value))
            self.spectrum_countdown_var.set(f"{remain_s:.1f} s")
        elif float(value) >= 1.0:
            self.spectrum_countdown_var.set("0.0 s")

    def _queue_progress(self, value: float, text: str = "") -> None:
        self.log_queue.put({"type": "progress", "value": float(value), "text": text})
        self.root.after(0, lambda v=float(value), t=text: self._update_spectrum_info_from_progress(v, t))

    def _drain_log_queue(self) -> None:
        while True:
            try:
                item = self.log_queue.get_nowait()
            except queue.Empty:
                break
            if isinstance(item, dict) and item.get("type") == "progress":
                self.progress_var.set(max(0.0, min(100.0, item.get("value", 0.0) * 100.0)))
                if item.get("text"):
                    self.progress_text_var.set(str(item["text"]))
            else:
                self._append_log(str(item))
        self.root.after(150, self._drain_log_queue)

    def load_config(self) -> None:
        self._ensure_config_exists()
        config_path = find_existing_path("app_config.json")
        if config_path is None:
            messagebox.showerror("配置错误", "未找到配置文件。")
            return
        data = json.loads(config_path.read_text(encoding="utf-8-sig"))
        self.hidden_config = {
            "spectrometer": {
                key: value
                for key, value in data.get("spectrometer", {}).items()
                if key
                not in {
                    "rayleigh_wavelength_nm",
                    "excitation_wavelength_nm",
                    "center_wavelength_nm",
                    "grating_no",
                    "exposure_ms",
                    "exposure_s",
                    "trigger_mode",
                }
            }
        }
        flat = {
            "rigol_visa_resource": data.get("rigol_visa_resource", ""),
            "andor_sdk_root": resolve_ui_sdk_root(data.get("andor_sdk_root", "")),
            "output_dir": data["output_dir"],
            "sample_name": data["sample_name"],
            "phase_start_deg": str(data["phase_start_deg"]),
            "phase_stop_deg": str(data["phase_stop_deg"]),
            "phase_step_deg": str(data["phase_step_deg"]),
            "repeats_per_phase": str(data["repeats_per_phase"]),
            "ch1_start_delay_ms": str(data.get("ch1_start_delay_ms", float(data.get("ch1_start_delay_s", 0.0)) * 1000.0)),
            "settle_time_ms": str(data.get("settle_time_ms", float(data.get("settle_time_s", 0.2)) * 1000.0)),
            "trigger_mode": str(data["spectrometer"]["trigger_mode"]),
            "ch1_waveform": data["ch1"]["waveform"],
            "ch1_frequency_hz": str(data["ch1"]["frequency_hz"]),
            "ch1_amplitude_vpp": str(data["ch1"]["amplitude_vpp"]),
            "ch1_offset_vdc": str(data["ch1"]["offset_vdc"]),
            "ch1_phase_deg": str(data["ch1"]["phase_deg"]),
            "ch1_duty_cycle_percent": str(data["ch1"].get("duty_cycle_percent", 50.0)),
            "ch2_waveform": data["ch2"]["waveform"],
            "ch2_frequency_hz": str(data["ch2"]["frequency_hz"]),
            "ch2_amplitude_vpp": str(data["ch2"]["amplitude_vpp"]),
            "ch2_offset_vdc": str(data["ch2"]["offset_vdc"]),
            "ch2_phase_deg": str(data["ch2"]["phase_deg"]),
            "ch2_duty_cycle_percent": str(data["ch2"].get("duty_cycle_percent", 50.0)),
            "rayleigh_wavelength_nm": str(
                data["spectrometer"].get("rayleigh_wavelength_nm", data["spectrometer"].get("excitation_wavelength_nm", 532.0))
            ),
            "center_wavelength_nm": str(data["spectrometer"]["center_wavelength_nm"]),
            "grating_no": str(data["spectrometer"]["grating_no"]),
            "exposure_ms": str(
                data["spectrometer"].get("exposure_ms", float(data["spectrometer"].get("exposure_s", 0.2)) * 1000.0)
            ),
        }
        for key, value in flat.items():
            if key in self.vars:
                self.vars[key].set(value)
        self.offline_simulation_var.set(bool(data.get("offline_simulation", False)))
        self.status_var.set("已加载配置")
        cfg = self._build_runtime_config_quiet()
        if cfg is not None:
            self._refresh_preview_from_latest_file(cfg)

    def _collect_data(self) -> dict:
        spectrometer_data = {
            **self.hidden_config.get("spectrometer", {}),
            "rayleigh_wavelength_nm": float(self.vars["rayleigh_wavelength_nm"].get().strip()),
            "center_wavelength_nm": float(self.vars["center_wavelength_nm"].get().strip()),
            "grating_no": int(self.vars["grating_no"].get().strip()),
            "exposure_ms": float(self.vars["exposure_ms"].get().strip()),
            "trigger_mode": int(self.vars["trigger_mode"].get().strip()),
        }
        return {
            "rigol_visa_resource": self.vars["rigol_visa_resource"].get().strip(),
            "andor_sdk_root": resolve_ui_sdk_root(self.vars["andor_sdk_root"].get()),
            "output_dir": self.vars["output_dir"].get().strip(),
            "sample_name": self.vars["sample_name"].get().strip(),
            "phase_start_deg": float(self.vars["phase_start_deg"].get().strip()),
            "phase_stop_deg": float(self.vars["phase_stop_deg"].get().strip()),
            "phase_step_deg": float(self.vars["phase_step_deg"].get().strip()),
            "repeats_per_phase": int(self.vars["repeats_per_phase"].get().strip()),
            "ch1_start_delay_ms": float(self.vars["ch1_start_delay_ms"].get().strip()),
            "settle_time_ms": float(self.vars["settle_time_ms"].get().strip()),
            "offline_simulation": bool(self.offline_simulation_var.get()),
            "ch1": {
                "waveform": self.vars["ch1_waveform"].get().strip(),
                "frequency_hz": float(self.vars["ch1_frequency_hz"].get().strip()),
                "amplitude_vpp": float(self.vars["ch1_amplitude_vpp"].get().strip()),
                "offset_vdc": float(self.vars["ch1_offset_vdc"].get().strip()),
                "phase_deg": float(self.vars["ch1_phase_deg"].get().strip()),
                "duty_cycle_percent": float(self.vars["ch1_duty_cycle_percent"].get().strip()),
            },
            "ch2": {
                "waveform": self.vars["ch2_waveform"].get().strip(),
                "frequency_hz": float(self.vars["ch2_frequency_hz"].get().strip()),
                "amplitude_vpp": float(self.vars["ch2_amplitude_vpp"].get().strip()),
                "offset_vdc": float(self.vars["ch2_offset_vdc"].get().strip()),
                "phase_deg": float(self.vars["ch2_phase_deg"].get().strip()),
                "duty_cycle_percent": float(self.vars["ch2_duty_cycle_percent"].get().strip()),
            },
            "spectrometer": spectrometer_data,
        }

    def save_config(self) -> None:
        try:
            data = self._collect_data()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return
        target = find_existing_path("app_config.json") or find_path_candidates("app_config.json")[0]
        target.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        self.status_var.set("已保存配置")
        self._append_log(f"配置已保存: {target}")

    def save_config_as(self) -> None:
        try:
            data = self._collect_data()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return
        target = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not target:
            return
        Path(target).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        self.status_var.set("已保存配置")
        self._append_log(f"配置已另存为: {target}")

    def _build_runtime_config(self) -> IntegratedExperimentConfig:
        data = self._collect_data()
        return IntegratedExperimentConfig(
            rigol_visa_resource=data["rigol_visa_resource"],
            andor_sdk_root=Path(data["andor_sdk_root"]),
            output_dir=Path(data["output_dir"]),
            sample_name=data["sample_name"],
            phase_start_deg=data["phase_start_deg"],
            phase_stop_deg=data["phase_stop_deg"],
            phase_step_deg=data["phase_step_deg"],
            repeats_per_phase=data["repeats_per_phase"],
            ch1_start_delay_s=data["ch1_start_delay_ms"] / 1000.0,
            ch1=ChannelConfig(**data["ch1"]),
            ch2=ChannelConfig(**data["ch2"]),
            spectrometer=SpectrometerConfig(
                rayleigh_wavelength_nm=data["spectrometer"]["rayleigh_wavelength_nm"],
                center_wavelength_nm=data["spectrometer"]["center_wavelength_nm"],
                grating_no=data["spectrometer"]["grating_no"],
                exposure_s=data["spectrometer"]["exposure_ms"] / 1000.0,
                trigger_mode=data["spectrometer"]["trigger_mode"],
                slit_width_um=data["spectrometer"].get("slit_width_um"),
                target_temperature_c=int(data["spectrometer"].get("target_temperature_c", -60)),
                required_temperature_c=int(data["spectrometer"].get("required_temperature_c", -60)),
                cooldown_timeout_s=float(data["spectrometer"].get("cooldown_timeout_s", 3600.0)),
                pre_amp_gain=data["spectrometer"].get("pre_amp_gain"),
                horizontal_readout_mhz=data["spectrometer"].get("horizontal_readout_mhz"),
                output_amplifier=str(data["spectrometer"].get("output_amplifier", "conventional")),
                ad_channel=int(data["spectrometer"].get("ad_channel", 0)),
                camera_shutter_mode=data["spectrometer"].get("camera_shutter_mode"),
                camera_shutter_open_ms=int(data["spectrometer"].get("camera_shutter_open_ms", 0)),
                camera_shutter_close_ms=int(data["spectrometer"].get("camera_shutter_close_ms", 0)),
                shamrock_shutter_mode=data["spectrometer"].get("shamrock_shutter_mode"),
            ),
            settle_time_s=data["settle_time_ms"] / 1000.0,
            offline_simulation=data["offline_simulation"],
        )

    def _run_background(self, label: str, action) -> None:
        if self.running:
            messagebox.showinfo("忙碌", "当前已有任务在运行。")
            return

        self.running = True
        self.pause_stage = "running"
        self.pause_button_var.set("暂停实验")
        self.status_var.set(label)
        self._set_status_card("experiment", label, "#165DFF")
        self.progress_var.set(0.0)
        self.progress_text_var.set("等待曝光")
        self._append_log(f"{label}...")

        def worker() -> None:
            try:
                import contextlib

                class QueueWriter:
                    def __init__(self, callback):
                        self.callback = callback

                    def write(self, text: str) -> None:
                        text = text.strip()
                        if text:
                            self.callback(text)

                    def flush(self) -> None:
                        return

                writer = QueueWriter(self._queue_log)
                with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                    action()
                self._queue_progress(1.0, "当前任务完成")
                self.root.after(0, lambda: self._set_pause_stage("idle"))
                self.root.after(0, lambda: self.status_var.set("完成"))
                self.root.after(0, lambda: self._set_status_card("experiment", "完成", "#00B42A"))
            except ExperimentPaused as exc:
                self._queue_log(str(exc))
                self._queue_progress(0.0, "实验已暂停")
                self.root.after(0, lambda: self._set_pause_stage("paused"))
                self.root.after(0, lambda: self.status_var.set("已暂停"))
                self.root.after(0, lambda: self._set_status_card("experiment", "已暂停", "#FF7D00"))
            except Exception:
                self._queue_log(traceback.format_exc())
                self._queue_progress(0.0, "任务失败")
                self.root.after(0, lambda: self._set_pause_stage("idle"))
                self.root.after(0, lambda: self.status_var.set("失败"))
                self.root.after(0, lambda: self._set_status_card("experiment", "失败", "#F53F3F"))
                self.root.after(0, lambda: messagebox.showerror("运行失败", "任务执行失败，请查看日志窗口。"))
            finally:
                self.running = False

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def _prepare_config(self) -> IntegratedExperimentConfig | None:
        try:
            self.save_config()
            return self._build_runtime_config()
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return None

    def run_generator_debug(self) -> None:
        config = self._prepare_config()
        if config is None:
            return
        def action() -> None:
            start_generator_debug(config, config.ch1.phase_deg)
            self.root.after(0, lambda: self._set_channel_toggle_visual(1, True))
            self.root.after(0, lambda: self._set_channel_toggle_visual(2, True))
            self.root.after(0, lambda: self._set_status_card("generator", "已连接", "#00B42A"))
        self._run_background("发生器调试", action)

    def run_generator_stop(self) -> None:
        config = self._prepare_config()
        if config is None:
            return
        def action() -> None:
            stop_generator_debug(config)
            self.root.after(0, lambda: self._set_channel_toggle_visual(1, False))
            self.root.after(0, lambda: self._set_channel_toggle_visual(2, False))
            self.root.after(0, lambda: self._set_status_card("generator", "待连接", "#FF7D00"))
        self._run_background("停止发生器", action)

    def _set_channel_toggle_visual(self, channel: int, enabled: bool) -> None:
        self.channel_enabled_state[channel] = enabled
        self.channel_toggle_text[channel].set(f"{'关闭' if enabled else '打开'} CH{channel}")
        button = self.channel_toggle_buttons.get(channel)
        if button is None:
            return
        button.configure(
            relief=tk.SUNKEN if enabled else tk.RAISED,
            bd=3 if enabled else 2,
            bg="#d9e6d2" if enabled else self.root.cget("bg"),
            activebackground="#d9e6d2" if enabled else "#e9e9e9",
        )

    def run_generator_channel(self, enabled: bool, channel: int) -> None:
        config = self._prepare_config()
        if config is None:
            return
        label = f"{'打开' if enabled else '关闭'} CH{channel}"
        if enabled:
            def action() -> None:
                start_generator_channel(config, channel)
                self.root.after(0, lambda ch=channel: self._set_channel_toggle_visual(ch, True))
                self.root.after(0, lambda: self._set_status_card("generator", "已连接", "#00B42A"))
            self._run_background(label, action)
        else:
            def action() -> None:
                stop_generator_channel(config, channel)
                self.root.after(0, lambda ch=channel: self._set_channel_toggle_visual(ch, False))
            self._run_background(label, action)

    def run_generator_channel_toggle(self, channel: int) -> None:
        enabled = not self.channel_enabled_state[channel]
        self.run_generator_channel(enabled, channel)

    def toggle_continuous_preview(self) -> None:
        if self.continuous_preview_active:
            if self.continuous_preview_stop_event is not None:
                self.continuous_preview_stop_event.set()
            self.continuous_preview_button_var.set("正在停止...")
            self.status_var.set("正在停止连续采集")
            self._append_log("已请求停止连续采集。")
            return
        if self.running:
            messagebox.showinfo("忙碌", "当前已有任务在运行。")
            return
        config = self._prepare_config()
        if config is None:
            return

        stop_event = threading.Event()
        self.continuous_preview_stop_event = stop_event
        self.continuous_preview_active = True
        self.running = True
        self.pause_stage = "running"
        self.pause_button_var.set("暂停实验")
        self.continuous_preview_button_var.set("停止连续采集")
        self.status_var.set("连续采集中")
        self._set_status_card("experiment", "连续采集中", "#165DFF")
        self.progress_var.set(0.0)
        self.progress_text_var.set("等待曝光")
        self._append_log("连续采集已启动：仅刷新固定预览文件，不进入正式实验数据。")

        def finish_ui(text: str, color: str) -> None:
            self.running = False
            self.continuous_preview_active = False
            self.continuous_preview_stop_event = None
            self.continuous_preview_button_var.set("连续采集")
            self._set_pause_stage("idle")
            self.status_var.set(text)
            self._set_status_card("experiment", text, color)

        def worker() -> None:
            try:
                import contextlib

                class QueueWriter:
                    def __init__(self, callback) -> None:
                        self.callback = callback
                        self.buffer = ""

                    def write(self, text: str) -> int:
                        self.buffer += text
                        while "\n" in self.buffer:
                            line, self.buffer = self.buffer.split("\n", 1)
                            if line.strip():
                                self.callback(line)
                        return len(text)

                    def flush(self) -> None:
                        if self.buffer.strip():
                            self.callback(self.buffer.strip())
                        self.buffer = ""

                with contextlib.redirect_stdout(QueueWriter(self._queue_log)), contextlib.redirect_stderr(
                    QueueWriter(self._queue_log)
                ):
                    output_path = run_continuous_preview_acquisition(
                        config,
                        stop_event,
                        self._queue_progress,
                        self.temperature_monitor_controller,
                        lambda p: self.root.after(0, lambda path=p: self._refresh_preview_from_path(path)),
                    )
                self.root.after(0, lambda p=output_path: self._refresh_preview_from_path(p))
                self._queue_progress(0.0, "连续采集已停止")
                self.root.after(0, lambda: finish_ui("已停止", "#00B42A"))
            except Exception:
                self._queue_log(traceback.format_exc())
                self._queue_progress(0.0, "连续采集失败")
                self.root.after(0, lambda: finish_ui("失败", "#F53F3F"))
                self.root.after(0, lambda: messagebox.showerror("连续采集失败", "连续采集失败，请查看日志窗口。"))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def run_andor_debug(self) -> None:
        config = self._prepare_config()
        if config is None:
            return
        def action() -> None:
            output_path = start_andor_debug(config, self._queue_progress, self.temperature_monitor_controller)
            self.root.after(0, lambda p=output_path: self._refresh_preview_from_path(p))
        self._run_background("光谱仪调试", action)

    def run_andor_connection_test(self) -> None:
        config = self._prepare_config()
        if config is None:
            return
        self._run_background("光谱仪连接测试", lambda: test_andor_connection(config))

    def run_experiment(self) -> None:
        config = self._prepare_config()
        if config is None:
            return
        self._reset_experiment_stress_trend()
        def action() -> None:
            run_integrated_experiment(
                config,
                self._queue_progress,
                self.temperature_monitor_controller,
                lambda p: self.root.after(0, lambda path=p: self._refresh_preview_and_update_experiment_trend(path)),
            )
            self.root.after(0, lambda c=config: self._refresh_preview_from_latest_file(c))
        self._run_background("实验运行中", action)

    def run_baseline_test(self) -> None:
        config = self._prepare_config()
        if config is None:
            return
        def action() -> None:
            run_baseline_test(
                config,
                self._queue_progress,
                self.temperature_monitor_controller,
                lambda p: self.root.after(0, lambda path=p: self._refresh_preview_from_path(path)),
            )
            self.root.after(0, lambda c=config: self._refresh_preview_from_latest_file(c))
        self._run_background("基线测试", action)

    def _set_pause_stage(self, stage: str) -> None:
        self.pause_stage = stage
        if stage == "paused":
            self.pause_button_var.set("停止实验")
        else:
            self.pause_button_var.set("暂停实验")

    def pause_experiment(self) -> None:
        if self.pause_stage == "paused" and not self.running:
            config = self._prepare_config()
            if config is None:
                return
            stop_paused_experiment(config)
            self._append_log("已停止实验，并清除续跑状态。下次开始将从头运行。")
            self.status_var.set("已停止")
            self.progress_var.set(0.0)
            self.progress_text_var.set("等待曝光")
            self._set_pause_stage("idle")
            return
        if not self.running:
            messagebox.showinfo("提示", "当前没有正在运行的实验。")
            return
        request_pause_experiment()
        self._append_log("已请求暂停。当前采集会尽快停止，之后再次点击“开始/继续实验”会从未完成位置继续。")
        self._set_pause_stage("pausing")
        self.status_var.set("正在暂停")


def main() -> None:
    try:
        root = tk.Tk()
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        TRRamanUI(root)
        root.mainloop()
    except Exception as exc:
        try:
            messagebox.showerror("启动错误", f"{exc}\n\n{traceback.format_exc()}")
        except Exception:
            print(exc)
            print(traceback.format_exc())
        raise


def _replacement_start_temperature_poll(self: TRRamanUI) -> None:
    cfg = self._build_runtime_config_quiet()
    if cfg is None or cfg.offline_simulation:
        return
    self.temperature_poll_inflight = True

    def worker() -> None:
        try:
            andor = self.temperature_monitor_controller
            if andor is None:
                andor = AndorSDKController(cfg.andor_sdk_root)
                andor.open()
                andor.enable_cooler_and_set_target(cfg.spectrometer.target_temperature_c)
                self.temperature_monitor_controller = andor
            temp_c, status_code = andor.get_temperature_c()
            stable = temp_c <= (cfg.spectrometer.required_temperature_c + 0.2) and status_code == DRV_TEMP_STABILIZED
            color = "#1f3dff" if stable else "#d84b4b"
            status = "可开始实验" if stable else "制冷中"
            self.root.after(
                0,
                lambda t=temp_c, s=status, c=color: (
                    self._set_temperature_badge(f"{t:.1f} °C", s, c),
                    self._set_status_card("andor", "已连接", "#00B42A"),
                ),
            )
        except (AndorHardwareNotFoundError, ShamrockHardwareNotFoundError):
            self._close_temperature_monitor()
            self.root.after(
                0,
                lambda: (
                    self._set_temperature_badge("--.- °C", "未连接", "#7f8c8d"),
                    self._set_status_card("andor", "未连接", "#F53F3F"),
                ),
            )
        except (AndorSDKError, FileNotFoundError) as exc:
            self._close_temperature_monitor()
            self.root.after(
                0,
                lambda: (
                    self._set_temperature_badge("--.- °C", "驱动异常", "#7f8c8d"),
                    self._set_status_card("andor", "异常", "#FF7D00"),
                ),
            )
            self._queue_log(f"Andor 温度监视失败: {exc}")
        finally:
            self.temperature_poll_inflight = False

    threading.Thread(target=worker, daemon=True).start()


def _replacement_run_background(self: TRRamanUI, label: str, action) -> None:
    if self.running:
        messagebox.showinfo("忙碌", "当前已有任务在运行。")
        return

    self.running = True
    self.pause_stage = "running"
    self.pause_button_var.set("暂停实验")
    self.status_var.set(label)
    self._set_status_card("experiment", label, "#165DFF")
    self.progress_var.set(0.0)
    self.progress_text_var.set("等待曝光")
    self._append_log(f"{label}...")

    def worker() -> None:
        try:
            import contextlib

            class QueueWriter:
                def __init__(self, callback):
                    self.callback = callback

                def write(self, text: str) -> None:
                    text = text.strip()
                    if text:
                        self.callback(text)

                def flush(self) -> None:
                    return

            writer = QueueWriter(self._queue_log)
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                action()
            self._queue_progress(1.0, "当前任务完成")
            self.root.after(0, lambda: self._set_pause_stage("idle"))
            self.root.after(0, lambda: self.status_var.set("完成"))
            self.root.after(0, lambda: self._set_status_card("experiment", "完成", "#00B42A"))
        except ExperimentPaused as exc:
            self._queue_log(str(exc))
            self._queue_progress(0.0, "实验已暂停")
            self.root.after(0, lambda: self._set_pause_stage("paused"))
            self.root.after(0, lambda: self.status_var.set("已暂停"))
            self.root.after(0, lambda: self._set_status_card("experiment", "已暂停", "#FF7D00"))
        except Exception:
            self._queue_log(traceback.format_exc())
            self._queue_progress(0.0, "任务失败")
            self.root.after(0, lambda: self._set_pause_stage("idle"))
            self.root.after(0, lambda: self.status_var.set("失败"))
            self.root.after(0, lambda: self._set_status_card("experiment", "失败", "#F53F3F"))
            self.root.after(0, lambda: messagebox.showerror("运行失败", "任务执行失败，请查看日志窗口。"))
        finally:
            self.running = False

    self.worker = threading.Thread(target=worker, daemon=True)
    self.worker.start()


TRRamanUI._start_temperature_poll = _replacement_start_temperature_poll
TRRamanUI._run_background = _replacement_run_background


def _replacement_run_andor_connection_test(self: TRRamanUI) -> None:
    config = self._prepare_config()
    if config is None:
        return
    self._run_background("光谱仪连接测试", lambda: test_andor_connection(config, self.temperature_monitor_controller))


TRRamanUI.run_andor_connection_test = _replacement_run_andor_connection_test


if __name__ == "__main__":
    main()
