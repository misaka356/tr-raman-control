from __future__ import annotations

import argparse
import dataclasses
from contextlib import nullcontext
import json
import math
import sys
import threading
import time
from pathlib import Path

import pyvisa

from andor_sdk_integration import (
    AndorHardwareNotFoundError,
    AndorSDKController,
    AndorSDKError,
    DRV_TEMP_STABILIZED,
    OUTPUT_AMPLIFIER_CONVENTIONAL,
    ShamrockHardwareNotFoundError,
)

TRIGGER_INTERNAL = 0
TRIGGER_EXTERNAL = 1
RESUME_STATE_FILENAME = "_experiment_resume_state.json"
POST_OUTPUT_SETTLE_S = 0.5
DEFAULT_STRESS_CONSTANT_MPA_PER_CM1 = 435.0
DEFAULT_STRESS_FIT_MIN_CM1 = 515.0
DEFAULT_STRESS_FIT_MAX_CM1 = 535.0
REALTIME_FIT_MIN_R2 = 0.92
REALTIME_FIT_WARN_R2 = 0.96
REALTIME_FIT_MIN_SNR = 8.0
REALTIME_FIT_WARN_SNR = 14.0
REALTIME_FIT_MAX_CENTER_STD_CM1 = 0.35
REALTIME_FIT_WARN_CENTER_STD_CM1 = 0.18
REALTIME_FIT_MAX_RAW_CENTER_DELTA_CM1 = 3.0
REALTIME_FIT_WARN_RAW_CENTER_DELTA_CM1 = 1.6
REALTIME_FIT_MAX_FWHM_CM1 = 18.0
REALTIME_FIT_WARN_FWHM_CM1 = 12.0


@dataclasses.dataclass
class ChannelConfig:
    waveform: str
    frequency_hz: float
    amplitude_vpp: float
    offset_vdc: float = 0.0
    phase_deg: float = 0.0
    duty_cycle_percent: float = 50.0


@dataclasses.dataclass
class SpectrometerConfig:
    rayleigh_wavelength_nm: float
    center_wavelength_nm: float
    grating_no: int
    exposure_s: float
    trigger_mode: int = TRIGGER_EXTERNAL
    slit_width_um: float | None = None
    target_temperature_c: int = -60
    required_temperature_c: int = -60
    cooldown_timeout_s: float = 3600.0
    pre_amp_gain: float | None = 2.0
    horizontal_readout_mhz: float | None = 1.48
    output_amplifier: str = "conventional"
    ad_channel: int = 0
    camera_shutter_mode: int | None = None
    camera_shutter_open_ms: int = 0
    camera_shutter_close_ms: int = 0
    shamrock_shutter_mode: int | None = None


@dataclasses.dataclass
class IntegratedExperimentConfig:
    rigol_visa_resource: str
    andor_sdk_root: Path
    output_dir: Path
    sample_name: str
    phase_start_deg: float
    phase_stop_deg: float
    phase_step_deg: float
    repeats_per_phase: int
    ch1_start_delay_s: float
    ch1: ChannelConfig
    ch2: ChannelConfig
    spectrometer: SpectrometerConfig
    settle_time_s: float = 0.2
    offline_simulation: bool = False


class ExperimentPaused(RuntimeError):
    pass


def noop_progress_callback(value: float, text: str = "") -> None:
    return


def noop_spectrum_saved_callback(path: Path) -> None:
    return


class VisaTransport:
    def __init__(self, resource_name: str, timeout_ms: int = 5000) -> None:
        self.resource_name = resource_name
        self.timeout_ms = timeout_ms
        self.rm: pyvisa.ResourceManager | None = None
        self.resource = None

    def open(self) -> None:
        self.rm = pyvisa.ResourceManager()
        self.resource = self.rm.open_resource(self.resource_name)
        self.resource.timeout = self.timeout_ms
        self.resource.write_termination = "\n"
        self.resource.read_termination = "\n"

    def close(self) -> None:
        if self.resource is not None:
            self.resource.close()
            self.resource = None
        if self.rm is not None:
            self.rm.close()
            self.rm = None

    def write(self, cmd: str) -> None:
        if self.resource is None:
            raise RuntimeError("VISA resource not open")
        self.resource.write(cmd)

    def query(self, cmd: str) -> str:
        if self.resource is None:
            raise RuntimeError("VISA resource not open")
        return str(self.resource.query(cmd)).strip()


class RigolSCPI:
    def __init__(self, visa_resource: str) -> None:
        if not visa_resource:
            raise ValueError("VISA 资源不能为空")
        self.transport = VisaTransport(visa_resource)

    def __enter__(self) -> "RigolSCPI":
        self.transport.open()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.transport.close()

    def write(self, cmd: str) -> None:
        self.transport.write(cmd)

    def query(self, cmd: str) -> str:
        return self.transport.query(cmd)

    def identify(self) -> str:
        return self.query("*IDN?")

    def query_burst_cycles(self, ch: int) -> str:
        return self.query(f":SOURce{ch}:BURSt:NCYCles?")

    def query_burst_mode(self, ch: int) -> str:
        return self.query(f":SOURce{ch}:BURSt:MODE?")

    def query_trigger_source(self, ch: int) -> str:
        return self.query(f":TRIGger{ch}:SOURce?")

    def query_system_error(self) -> str:
        try:
            return self.query(":SYSTem:ERRor?")
        except Exception as exc:
            return f"query_failed: {exc}"

    def wait_operation_complete(self) -> None:
        result = self.query("*OPC?")
        if result != "1":
            raise RuntimeError(f"Unexpected *OPC? response: {result}")

    def configure_channel(self, ch: int, cfg: ChannelConfig) -> None:
        waveform = cfg.waveform.lower()
        if waveform == "sin":
            self.write(
                f":SOURce{ch}:APPLy:SINusoid "
                f"{cfg.frequency_hz},{cfg.amplitude_vpp},{cfg.offset_vdc},{cfg.phase_deg}"
            )
            return
        if waveform == "square":
            self.write(
                f":SOURce{ch}:APPLy:SQUare "
                f"{cfg.frequency_hz},{cfg.amplitude_vpp},{cfg.offset_vdc},{cfg.phase_deg}"
            )
            self.write(f":SOURce{ch}:FUNCtion:SQUare:DCYCle {cfg.duty_cycle_percent}")
            return
        if waveform == "pulse":
            self.write(
                f":SOURce{ch}:APPLy:PULSe "
                f"{cfg.frequency_hz},{cfg.amplitude_vpp},{cfg.offset_vdc},{cfg.phase_deg}"
            )
            self.write(f":SOURce{ch}:FUNCtion:PULSe:DCYCle {cfg.duty_cycle_percent}")
            return
        if waveform == "ramp":
            self.write(
                f":SOURce{ch}:APPLy:RAMP "
                f"{cfg.frequency_hz},{cfg.amplitude_vpp},{cfg.offset_vdc},{cfg.phase_deg}"
            )
            return
        if waveform == "noise":
            self.write(
                f":SOURce{ch}:APPLy:NOISe "
                f"{cfg.frequency_hz},{cfg.amplitude_vpp},{cfg.offset_vdc},{cfg.phase_deg}"
            )
            return
        if waveform == "dc":
            self.write(
                f":SOURce{ch}:APPLy:DC "
                f"{cfg.frequency_hz},{cfg.amplitude_vpp},{cfg.offset_vdc},{cfg.phase_deg}"
            )
            return
        raise ValueError(f"不支持的波形: {cfg.waveform}")

    def configure_continuous_output(self, ch: int, cfg: ChannelConfig) -> None:
        self.configure_channel(ch, cfg)
        self.write(f":SOURce{ch}:BURSt:STATe OFF")

    def set_phase(self, ch: int, phase_deg: float) -> None:
        self.write(f":SOURce{ch}:PHASe {phase_deg}")

    def set_burst_phase(self, ch: int, phase_deg: float) -> None:
        self.write(f":SOURce{ch}:BURSt:PHASe {phase_deg}")

    def query_burst_phase(self, ch: int) -> float:
        return float(self.query(f":SOURce{ch}:BURSt:PHASe?"))

    def set_burst_phase_verified(
        self,
        ch: int,
        phase_deg: float,
        tolerance_deg: float = 0.05,
        retries: int = 3,
    ) -> float:
        last_phase: float | None = None
        for attempt in range(1, retries + 1):
            self.set_burst_phase(ch, phase_deg)
            self.wait_operation_complete()
            last_phase = self.query_burst_phase(ch)
            phase_error = ((last_phase - phase_deg + 180.0) % 360.0) - 180.0
            if abs(phase_error) <= tolerance_deg:
                return last_phase
            time.sleep(0.1 * attempt)
        raise RuntimeError(
            f"CH{ch} Burst phase write failed: target {phase_deg:.6f} deg, "
            f"readback {last_phase:.6f} deg"
        )

    def set_output(self, ch: int, enabled: bool) -> None:
        self.write(f":OUTPut{ch}:STATe {'ON' if enabled else 'OFF'}")

    def set_idle_level_center(self, ch: int) -> None:
        self.write(f":OUTPut{ch}:IDLE CENTer")

    def set_trigger_output(self, ch: int, enabled: bool, positive_edge: bool = True) -> None:
        self.write(f":OUTPut{ch}:TRIGger {'ON' if enabled else 'OFF'}")
        if enabled:
            self.write(f":OUTPut{ch}:TRIGger:SLOPe {'POSitive' if positive_edge else 'NEGative'}")

    def _burst_cycles_is_infinite(self, response: str) -> bool:
        text = response.strip().upper()
        if "INF" in text:
            return True
        try:
            value = float(text)
            if value == -1.0:
                return True
            return value >= 9.0e36 or value >= 1.0e6
        except Exception:
            return False

    def set_burst_cycles_infinite_compat(self, ch: int) -> str:
        attempts = [
            ("INFinity", f":SOURce{ch}:BURSt:NCYCles INFinity"),
            ("INF", f":SOURce{ch}:BURSt:NCYCles INF"),
            ("MAXimum", f":SOURce{ch}:BURSt:NCYCles MAXimum"),
            ("1000000", f":SOURce{ch}:BURSt:NCYCles 1000000"),
        ]
        last_response = ""
        for _label, cmd in attempts:
            self.write(cmd)
            response = self.query_burst_cycles(ch)
            last_response = response
            if self._burst_cycles_is_infinite(response):
                return response
        return last_response

    def configure_burst_ch1_manual_master(self, trigger_delay_s: float, burst_phase_deg: float) -> None:
        self.write(":SOURce1:BURSt:STATe ON")
        self.write(":SOURce1:BURSt:MODE TRIGgered")
        self.set_burst_cycles_infinite_compat(1)
        self.write(":TRIGger1:SOURce BUS")
        self.write(f":TRIGger1:DELay {trigger_delay_s:.9f}")
        self.set_idle_level_center(1)
        self.set_burst_phase(1, burst_phase_deg)
        self.set_trigger_output(1, True, positive_edge=True)

    def configure_burst_ch2_external_slave(self, trigger_delay_s: float) -> None:
        self.write(":SOURce2:BURSt:STATe ON")
        self.write(":SOURce2:BURSt:MODE TRIGgered")
        self.set_burst_cycles_infinite_compat(2)
        self.write(":TRIGger2:SOURce EXTernal")
        self.write(":TRIGger2:SLOPe POSitive")
        self.write(f":TRIGger2:DELay {trigger_delay_s:.9f}")
        self.set_idle_level_center(2)
        self.set_trigger_output(2, False)

    def configure_triggered_infinite_burst(self, ch: int, trigger_delay_s: float) -> None:
        self.write(f":SOURce{ch}:BURSt:STATe ON")
        self.write(f":SOURce{ch}:BURSt:MODE TRIGgered")
        self.set_burst_cycles_infinite_compat(ch)
        self.write(f":TRIGger{ch}:SOURce BUS")
        self.write(f":TRIGger{ch}:DELay {trigger_delay_s:.9f}")

    def fire_manual_trigger_ch1(self) -> None:
        self.write(":TRIGger1:IMMediate")

    def abort_all(self) -> None:
        self.write(":ABORt")
        self.set_output(1, False)
        self.set_output(2, False)


class ExperimentRuntime:
    def __init__(self) -> None:
        self.pause_event = threading.Event()
        self.current_rigol: RigolSCPI | None = None
        self.current_andor: AndorSDKController | None = None

    def reset(self) -> None:
        self.pause_event.clear()
        self.current_rigol = None
        self.current_andor = None

    def request_pause(self) -> None:
        self.pause_event.set()
        if self.current_rigol is not None:
            try:
                self.current_rigol.abort_all()
            except Exception:
                pass
        if self.current_andor is not None:
            try:
                self.current_andor.abort_acquisition()
            except Exception:
                pass

    def is_pause_requested(self) -> bool:
        return self.pause_event.is_set()


RUNTIME = ExperimentRuntime()


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_default_config_path() -> Path | None:
    base = app_base_dir()
    for path in (base / "app_config.json", base.parent / "app_config.json"):
        if path.exists():
            return path
    return None


def resolve_andor_sdk_root(user_value: str | Path | None = None) -> Path:
    if user_value:
        return Path(user_value).expanduser().resolve()
    base = app_base_dir()
    candidates = [
        base / "vendor" / "andor_sdk",
        base.parent / "vendor" / "andor_sdk",
        base / "Andor SOLIS" / "Andor SOLIS",
        base.parent / "Andor SOLIS" / "Andor SOLIS",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return (base / "vendor" / "andor_sdk").resolve()


def build_phase_list(start_deg: float, stop_deg: float, step_deg: float) -> list[float]:
    if step_deg <= 0:
        raise ValueError("phase_step_deg 必须大于 0")
    phases: list[float] = []
    current = start_deg
    while current <= stop_deg + 1e-9:
        phases.append(round(current, 9))
        current += step_deg
    return phases


def format_phase_token(phase_deg: float) -> str:
    text = f"{phase_deg:.9f}".rstrip("0").rstrip(".")
    return text if text else "0"


def resolve_ch1_output_phase_deg(user_phase_deg: float) -> float:
    return -float(user_phase_deg)


def acquisition_filename(output_dir: Path, sample_name: str, phase_deg: float, rep: int) -> Path:
    return output_dir / sample_name / f"{format_phase_token(phase_deg)}-{rep}.asc"


def debug_acquisition_filename(output_dir: Path, sample_name: str) -> Path:
    return output_dir / sample_name / f"debug-{sample_name}.asc"


def continuous_preview_filename(output_dir: Path, sample_name: str) -> Path:
    return output_dir / sample_name / "_continuous_preview.asc"


def baseline_acquisition_filename(output_dir: Path, sample_name: str, rep: int) -> Path:
    return output_dir / sample_name / f"-{rep}.asc"


def sample_output_dir(output_dir: Path, sample_name: str) -> Path:
    return output_dir / sample_name


def pause_state_path(output_dir: Path, sample_name: str) -> Path:
    return sample_output_dir(output_dir, sample_name) / RESUME_STATE_FILENAME


def build_simulated_spectrum(phase_deg: float, points: int = 1600) -> tuple[list[float], list[int]]:
    x_axis: list[float] = []
    y_axis: list[int] = []
    phase_rad = math.radians(phase_deg)
    peak_center = 520.0 + 6.0 * math.sin(phase_rad)
    peak_height = 1800.0 + 700.0 * (1.0 + math.cos(phase_rad))
    peak_width = 7.5
    for i in range(points):
        x = 100.0 + i * 0.5
        baseline = 120.0 + 0.03 * x + 35.0 * math.sin(x / 55.0)
        gaussian = peak_height * math.exp(-((x - peak_center) ** 2) / (2.0 * peak_width**2))
        ripple = 18.0 * math.sin((x / 18.0) + phase_rad * 0.7)
        intensity = max(0, int(round(baseline + gaussian + ripple)))
        x_axis.append(x)
        y_axis.append(intensity)
    return x_axis, y_axis


def wavelength_nm_to_raman_shift_cm1(rayleigh_wavelength_nm: float, detected_wavelength_nm: float) -> float:
    # Follow the SOLIS help wording:
    # rs = 10 million x [(scatter - laser) / (scatter x laser)]
    return 1.0e7 * ((detected_wavelength_nm - rayleigh_wavelength_nm) / (detected_wavelength_nm * rayleigh_wavelength_nm))


def convert_wavelength_axis_to_raman_shift(
    x_axis_wavelength_nm: list[float],
    rayleigh_wavelength_nm: float,
) -> list[float]:
    return [
        wavelength_nm_to_raman_shift_cm1(rayleigh_wavelength_nm, x)
        for x in x_axis_wavelength_nm
    ]


def fit_raman_peak_center(
    x_axis: list[float],
    y_axis: list[float] | list[int],
    fit_min_cm1: float = DEFAULT_STRESS_FIT_MIN_CM1,
    fit_max_cm1: float = DEFAULT_STRESS_FIT_MAX_CM1,
) -> dict[str, float | str | bool]:
    import numpy as np
    from lmfit.models import LorentzianModel, QuadraticModel

    x = np.asarray([float(v) for v in x_axis], dtype=float)
    y = np.asarray([float(v) for v in y_axis], dtype=float)
    mask = (x >= fit_min_cm1) & (x <= fit_max_cm1)
    x_fit = x[mask]
    y_fit_data = y[mask]
    if x_fit.size < 12:
        raise ValueError(f"Too few points in fitting window {fit_min_cm1:g}-{fit_max_cm1:g} cm^-1")

    peak_idx = int(np.argmax(y_fit_data))
    raw_peak_x = float(x_fit[peak_idx])
    raw_peak_y = float(y_fit_data[peak_idx])

    peak_model = LorentzianModel(prefix="peak_")
    background_model = QuadraticModel(prefix="bg_")
    model = peak_model + background_model

    params = model.make_params()
    params.update(peak_model.guess(y_fit_data, x=x_fit))
    params["peak_center"].set(value=raw_peak_x, min=float(np.min(x_fit)), max=float(np.max(x_fit)))
    params["peak_sigma"].set(min=1e-6, max=max(10.0, float(np.max(x_fit) - np.min(x_fit))))
    params["peak_amplitude"].set(min=0.0)
    params["bg_a"].set(value=0.0)
    params["bg_b"].set(value=0.0)
    params["bg_c"].set(value=float(np.median(y_fit_data)))

    result = model.fit(y_fit_data, params, x=x_fit)
    best = result.best_values
    y_best = result.best_fit
    residual = y_fit_data - y_best
    ss_res = float(np.sum((y_fit_data - y_best) ** 2))
    ss_tot = float(np.sum((y_fit_data - np.mean(y_fit_data)) ** 2))
    fit_r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else 1.0
    center_std = (
        float(result.params["peak_center"].stderr)
        if result.params["peak_center"].stderr is not None
        else math.nan
    )
    baseline_level = float(np.percentile(y_fit_data, 20))
    peak_height = max(0.0, raw_peak_y - baseline_level)
    residual_mad = float(np.median(np.abs(residual - np.median(residual))))
    noise_floor = max(1.0, 1.4826 * residual_mad)
    snr = peak_height / noise_floor
    center_delta = abs(float(best["peak_center"]) - raw_peak_x)

    return {
        "center_cm1": float(best["peak_center"]),
        "center_std_cm1": center_std,
        "fwhm_cm1": 2.0 * float(best["peak_sigma"]),
        "amplitude": float(best["peak_amplitude"]),
        "fit_r2": fit_r2,
        "snr": snr,
        "peak_height": peak_height,
        "residual_mad": residual_mad,
        "center_raw_delta_cm1": center_delta,
        "raw_peak_x": raw_peak_x,
        "raw_peak_y": raw_peak_y,
        "method": "lmfit.LorentzianModel + QuadraticModel",
        "success": bool(result.success),
    }


def calculate_stress_mpa(
    baseline_peak_cm1: float,
    current_peak_cm1: float,
    stress_constant_mpa_per_cm1: float = DEFAULT_STRESS_CONSTANT_MPA_PER_CM1,
) -> float:
    return float(stress_constant_mpa_per_cm1) * (float(baseline_peak_cm1) - float(current_peak_cm1))


def _finite_float(value, default: float = math.nan) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return default
    return number if math.isfinite(number) else default


def evaluate_realtime_fit_quality(fit: dict) -> dict:
    """Classify a live single-peak fit before converting it to stress."""
    issues: list[str] = []
    warnings: list[str] = []
    fit_r2 = _finite_float(fit.get("fit_r2"))
    snr = _finite_float(fit.get("snr"))
    center_std = _finite_float(fit.get("center_std_cm1"))
    center_delta = _finite_float(fit.get("center_raw_delta_cm1"))
    fwhm = _finite_float(fit.get("fwhm_cm1"))

    if not fit.get("success", False):
        issues.append("lmfit did not converge")
    if not math.isfinite(fit_r2) or fit_r2 < REALTIME_FIT_MIN_R2:
        issues.append(f"low R2 {fit_r2:.3f}" if math.isfinite(fit_r2) else "R2 unavailable")
    elif fit_r2 < REALTIME_FIT_WARN_R2:
        warnings.append(f"R2 {fit_r2:.3f}")
    if not math.isfinite(snr) or snr < REALTIME_FIT_MIN_SNR:
        issues.append(f"low SNR {snr:.1f}" if math.isfinite(snr) else "SNR unavailable")
    elif snr < REALTIME_FIT_WARN_SNR:
        warnings.append(f"SNR {snr:.1f}")
    if math.isfinite(center_std):
        if center_std > REALTIME_FIT_MAX_CENTER_STD_CM1:
            issues.append(f"center uncertainty {center_std:.3f} cm^-1")
        elif center_std > REALTIME_FIT_WARN_CENTER_STD_CM1:
            warnings.append(f"center uncertainty {center_std:.3f} cm^-1")
    if not math.isfinite(center_delta) or center_delta > REALTIME_FIT_MAX_RAW_CENTER_DELTA_CM1:
        issues.append(
            f"fit/raw peak shift {center_delta:.2f} cm^-1"
            if math.isfinite(center_delta)
            else "fit/raw peak shift unavailable"
        )
    elif center_delta > REALTIME_FIT_WARN_RAW_CENTER_DELTA_CM1:
        warnings.append(f"fit/raw peak shift {center_delta:.2f} cm^-1")
    if not math.isfinite(fwhm) or fwhm <= 0.0 or fwhm > REALTIME_FIT_MAX_FWHM_CM1:
        issues.append(f"abnormal FWHM {fwhm:.2f} cm^-1" if math.isfinite(fwhm) else "FWHM unavailable")
    elif fwhm > REALTIME_FIT_WARN_FWHM_CM1:
        warnings.append(f"wide FWHM {fwhm:.2f} cm^-1")

    if issues:
        level = "reject"
        reliable = False
        message = "; ".join(issues[:3])
    elif warnings:
        level = "caution"
        reliable = True
        message = "; ".join(warnings[:3])
    else:
        level = "good"
        reliable = True
        message = f"R2 {fit_r2:.3f}, SNR {snr:.1f}"

    return {
        "level": level,
        "reliable": reliable,
        "message": message,
        "issues": issues,
        "warnings": warnings,
        "fit_r2": fit_r2,
        "snr": snr,
        "center_std_cm1": center_std,
        "center_raw_delta_cm1": center_delta,
        "fwhm_cm1": fwhm,
    }


def _legacy_analyze_ghost_peak(
    x_axis: list[float],
    y_axis: list[int],
    window_min: float = 460.0,
    window_max: float = 580.0,
) -> dict:
    points = [(float(x), float(y)) for x, y in zip(x_axis, y_axis) if window_min <= float(x) <= window_max]
    if len(points) < 7:
        return {
            "ghost_detected": False,
            "primary_peak": None,
            "secondary_peak": None,
            "window_min": window_min,
            "window_max": window_max,
            "message": "峰检测数据不足",
        }

    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    smoothed: list[float] = []
    for i in range(len(ys)):
        left = max(0, i - 2)
        right = min(len(ys), i + 3)
        smoothed.append(sum(ys[left:right]) / (right - left))

    peak_indices: list[int] = []
    for i in range(1, len(smoothed) - 1):
        if smoothed[i] > smoothed[i - 1] and smoothed[i] >= smoothed[i + 1]:
            peak_indices.append(i)

    if not peak_indices:
        return {
            "ghost_detected": False,
            "primary_peak": None,
            "secondary_peak": None,
            "window_min": window_min,
            "window_max": window_max,
            "message": "未检测到有效峰",
        }

    ranked = sorted(peak_indices, key=lambda i: smoothed[i], reverse=True)
    baseline = sorted(ys)[len(ys) // 2]
    primary_index = ranked[0]
    primary_peak = {"x": xs[primary_index], "y": ys[primary_index], "y_smooth": smoothed[primary_index]}
    min_prominence = max(120.0, (primary_peak["y_smooth"] - baseline) * 0.10)
    min_separation = 8.0
    min_secondary_ratio = 0.20
    secondary_peak = None
    ghost_detected = False

    for idx in ranked[1:]:
        separation = abs(xs[idx] - xs[primary_index])
        prominence = smoothed[idx] - baseline
        ratio = ys[idx] / max(primary_peak["y"], 1.0)
        if separation < min_separation:
            continue
        if prominence < min_prominence:
            continue
        if ratio < min_secondary_ratio:
            continue
        secondary_peak = {"x": xs[idx], "y": ys[idx], "y_smooth": smoothed[idx]}
        ghost_detected = True
        break

    if ghost_detected and secondary_peak is not None:
        message = (
            f"鬼峰警告: 主峰 {primary_peak['x']:.2f} cm^-1 / {primary_peak['y']:.0f}, "
            f"次峰 {secondary_peak['x']:.2f} cm^-1 / {secondary_peak['y']:.0f}"
        )
    else:
        message = f"主峰 {primary_peak['x']:.2f} cm^-1 / {primary_peak['y']:.0f}"

    return {
        "ghost_detected": ghost_detected,
        "primary_peak": primary_peak,
        "secondary_peak": secondary_peak,
        "window_min": window_min,
        "window_max": window_max,
        "message": message,
    }


def analyze_ghost_peak(
    x_axis: list[float],
    y_axis: list[int],
    window_min: float = 460.0,
    window_max: float = 580.0,
) -> dict:
    """Live ghost-peak warning for split, shoulder, and broadened silicon peaks."""
    import numpy as np

    points = sorted(
        [(float(x), float(y)) for x, y in zip(x_axis, y_axis) if window_min <= float(x) <= window_max],
        key=lambda item: item[0],
    )
    if len(points) < 7:
        return {
            "ghost_detected": False,
            "warning_detected": False,
            "risk_level": "none",
            "risk_score": 0.0,
            "warning_type": "insufficient_data",
            "primary_peak": None,
            "secondary_peak": None,
            "window_min": window_min,
            "window_max": window_max,
            "message": "Ghost check has too few points",
            "details": [],
        }

    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    x_np = np.asarray(xs, dtype=float)
    y_np = np.asarray(ys, dtype=float)

    smoothed: list[float] = []
    for i in range(len(ys)):
        left = max(0, i - 2)
        right = min(len(ys), i + 3)
        smoothed.append(sum(ys[left:right]) / (right - left))
    smooth_np = np.asarray(smoothed, dtype=float)

    peak_indices: list[int] = []
    for i in range(1, len(smoothed) - 1):
        if smoothed[i] > smoothed[i - 1] and smoothed[i] >= smoothed[i + 1]:
            peak_indices.append(i)

    if not peak_indices:
        return {
            "ghost_detected": False,
            "warning_detected": False,
            "risk_level": "none",
            "risk_score": 0.0,
            "warning_type": "no_peak",
            "primary_peak": None,
            "secondary_peak": None,
            "window_min": window_min,
            "window_max": window_max,
            "message": "No valid peak in ghost-check window",
            "details": [],
        }

    ranked = sorted(peak_indices, key=lambda i: smoothed[i], reverse=True)
    baseline = float(np.percentile(y_np, 20))
    primary_index = ranked[0]
    primary_peak = {"x": xs[primary_index], "y": ys[primary_index], "y_smooth": smoothed[primary_index]}
    primary_height = max(primary_peak["y_smooth"] - baseline, 1.0)
    min_prominence = max(80.0, primary_height * 0.08)
    secondary_peak = None
    risk_score = 0.0
    warning_type = "none"
    details: list[str] = []

    for idx in ranked[1:]:
        separation = abs(xs[idx] - xs[primary_index])
        prominence = smoothed[idx] - baseline
        ratio = prominence / primary_height
        if prominence < min_prominence:
            continue
        if separation >= 7.0 and ratio >= 0.18:
            secondary_peak = {"x": xs[idx], "y": ys[idx], "y_smooth": smoothed[idx]}
            warning_type = "separated_double_peak"
            risk_score = max(risk_score, 0.9 if ratio >= 0.35 else 0.62)
            details.append(f"separated secondary peak {separation:.1f} cm^-1 away, {ratio:.0%} of main peak")
            break
        if 2.5 <= separation < 7.0 and ratio >= 0.16:
            secondary_peak = {"x": xs[idx], "y": ys[idx], "y_smooth": smoothed[idx]}
            warning_type = "shoulder_peak"
            risk_score = max(risk_score, 0.75 if ratio >= 0.28 else 0.45)
            details.append(f"shoulder candidate {separation:.1f} cm^-1 away, {ratio:.0%} of main peak")
            break

    top_threshold = baseline + primary_height * 0.88
    top_indices = np.where(smooth_np >= top_threshold)[0]
    if top_indices.size:
        top_width = float(x_np[int(top_indices[-1])] - x_np[int(top_indices[0])])
        if top_width >= 5.0:
            warning_type = "broad_or_split_peak_top" if warning_type == "none" else warning_type
            risk_score = max(risk_score, 0.72 if top_width >= 7.0 else 0.42)
            details.append(f"wide or flattened peak top {top_width:.1f} cm^-1")

    half_height_indices = np.where(smooth_np >= baseline + primary_height * 0.50)[0]
    high_core_indices = np.where(smooth_np >= baseline + primary_height * 0.75)[0]
    if half_height_indices.size and high_core_indices.size:
        half_height_width = float(x_np[int(half_height_indices[-1])] - x_np[int(half_height_indices[0])])
        high_core_width = float(x_np[int(high_core_indices[-1])] - x_np[int(high_core_indices[0])])
        if half_height_width < 5.0 and high_core_width < 3.0 and primary_height >= 250.0:
            warning_type = "narrow_peak_top_spike" if warning_type == "none" else warning_type
            risk_score = max(risk_score, 0.82)
            details.append(
                f"abnormally narrow peak top: half-height {half_height_width:.1f} cm^-1, "
                f"75% height {high_core_width:.1f} cm^-1"
            )

    if risk_score >= 0.70:
        risk_level = "high"
    elif risk_score >= 0.35:
        risk_level = "low"
    else:
        risk_level = "none"
        risk_score = 0.0
        warning_type = "none"

    if risk_level == "high":
        message = f"High ghost risk: {warning_type}"
    elif risk_level == "low":
        message = f"Low ghost risk: {warning_type}"
    else:
        message = f"Main peak {primary_peak['x']:.2f} cm^-1 / {primary_peak['y']:.0f}"

    return {
        "ghost_detected": risk_level == "high",
        "warning_detected": risk_level != "none",
        "risk_level": risk_level,
        "risk_score": risk_score,
        "warning_type": warning_type,
        "primary_peak": primary_peak,
        "secondary_peak": secondary_peak,
        "window_min": window_min,
        "window_max": window_max,
        "message": message,
        "details": details,
    }


def apply_ghost_filename_marker(output_path: Path, analysis: dict) -> Path:
    if not analysis.get("ghost_detected"):
        return output_path
    if output_path.stem.endswith("_ghost"):
        return output_path
    marked_path = output_path.with_name(f"{output_path.stem}_ghost{output_path.suffix}")
    if marked_path == output_path:
        return output_path
    if marked_path.exists():
        marked_path.unlink()
    output_path.rename(marked_path)
    return marked_path


def save_ascii_data(output_path: Path, x_axis: list[float], y_axis: list[int]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fh:
        for x, y in zip(x_axis, y_axis):
            x_text = f"{x:.6f}"
            fh.write(f"{x_text}\t{y}\n")


def split_signed_delay(ch1_relative_to_ch2_s: float) -> tuple[float, float]:
    if ch1_relative_to_ch2_s >= 0:
        return ch1_relative_to_ch2_s, 0.0
    return 0.0, abs(ch1_relative_to_ch2_s)


def resolve_output_amplifier(amplifier: str) -> int:
    if amplifier.strip().lower() == "conventional":
        return OUTPUT_AMPLIFIER_CONVENTIONAL
    raise ValueError(f"Unsupported output amplifier: {amplifier}")


def build_resume_payload(cfg: IntegratedExperimentConfig, phases: list[float], phase_index: int, rep_index: int) -> dict:
    return {
        "phase_start_deg": cfg.phase_start_deg,
        "phase_stop_deg": cfg.phase_stop_deg,
        "phase_step_deg": cfg.phase_step_deg,
        "repeats_per_phase": cfg.repeats_per_phase,
        "phases": phases,
        "next_phase_index": phase_index,
        "next_rep_index": rep_index,
    }


def save_resume_state(cfg: IntegratedExperimentConfig, phases: list[float], phase_index: int, rep_index: int) -> None:
    state_path = pause_state_path(cfg.output_dir, cfg.sample_name)
    state_path.parent.mkdir(parents=True, exist_ok=True)
    payload = build_resume_payload(cfg, phases, phase_index, rep_index)
    state_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def clear_resume_state(cfg: IntegratedExperimentConfig) -> None:
    state_path = pause_state_path(cfg.output_dir, cfg.sample_name)
    if state_path.exists():
        state_path.unlink()


def load_resume_state(cfg: IntegratedExperimentConfig, phases: list[float]) -> tuple[int, int]:
    state_path = pause_state_path(cfg.output_dir, cfg.sample_name)
    if not state_path.exists():
        return 0, 1
    try:
        payload = json.loads(state_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return 0, 1

    expected = build_resume_payload(cfg, phases, 0, 1)
    for key in ("phase_start_deg", "phase_stop_deg", "phase_step_deg", "repeats_per_phase", "phases"):
        if payload.get(key) != expected.get(key):
            return 0, 1

    phase_index = int(payload.get("next_phase_index", 0))
    rep_index = int(payload.get("next_rep_index", 1))
    phase_index = max(0, min(phase_index, len(phases) - 1))
    rep_index = max(1, min(rep_index, cfg.repeats_per_phase))
    return phase_index, rep_index


def configure_generator_for_phase(
    rigol: RigolSCPI,
    cfg: IntegratedExperimentConfig,
    phase_deg: float,
    full_config: bool = True,
) -> None:
    ch1_delay, ch2_delay = split_signed_delay(cfg.ch1_start_delay_s)
    ch1_output_phase_deg = resolve_ch1_output_phase_deg(phase_deg)
    rigol.abort_all()
    if full_config:
        rigol.configure_channel(1, cfg.ch1)
        rigol.configure_channel(2, cfg.ch2)
        # CH1 is the AOM master channel. Its burst trigger output drives CH2 external trigger input.
        rigol.configure_burst_ch1_manual_master(ch1_delay, ch1_output_phase_deg)
        rigol.configure_burst_ch2_external_slave(ch2_delay)
        rigol.wait_operation_complete()
        rigol.set_burst_phase_verified(1, ch1_output_phase_deg)
    else:
        rigol.set_burst_phase_verified(1, ch1_output_phase_deg)
    rigol.set_output(1, True)
    rigol.set_output(2, True)


def ensure_not_paused(cfg: IntegratedExperimentConfig, phases: list[float], phase_index: int, rep_index: int) -> None:
    if RUNTIME.is_pause_requested():
        save_resume_state(cfg, phases, phase_index, rep_index)
        raise ExperimentPaused(f"实验已暂停，将从相位 {phases[phase_index]:g}°, 第 {rep_index} 次继续。")


def wait_with_pause(seconds: float, cfg: IntegratedExperimentConfig, phases: list[float], phase_index: int, rep_index: int) -> None:
    remaining = max(0.0, seconds)
    while remaining > 0:
        ensure_not_paused(cfg, phases, phase_index, rep_index)
        step = min(0.05, remaining)
        time.sleep(step)
        remaining -= step


def start_generator_debug(cfg: IntegratedExperimentConfig, phase_deg: float | None = None) -> None:
    phase = cfg.ch1.phase_deg if phase_deg is None else phase_deg
    ch1_output_phase_deg = resolve_ch1_output_phase_deg(phase)
    if cfg.offline_simulation:
        print(f"[离线模拟] CH1(AOM) 将作为主触发通道，CH2(音圈) 将等待外部上升沿触发")
        print(f"[离线模拟] CH1 Burst 相位设定值: {phase:.2f} deg")
        print(f"[离线模拟] CH1 Burst 实际输出相位: {ch1_output_phase_deg:.2f} deg")
        print(f"[离线模拟] CH1 相对 CH2 延时补偿: {cfg.ch1_start_delay_s:.9f} s")
        print(f"[离线模拟] 打开 CH1/CH2 后等待 {POST_OUTPUT_SETTLE_S:.1f} s 再触发")
        return
    with RigolSCPI(cfg.rigol_visa_resource) as rigol:
        print("Instrument:", rigol.identify())
        configure_generator_for_phase(rigol, cfg, phase)
        time.sleep(POST_OUTPUT_SETTLE_S)
        rigol.fire_manual_trigger_ch1()
        print(
            f"发生器已按手动触发链路启动: CH1(AOM) 设定相位 {phase:.2f} deg, "
            f"实际输出相位 {ch1_output_phase_deg:.2f} deg -> CH2(音圈) 外部触发"
        )


def stop_generator_debug(cfg: IntegratedExperimentConfig) -> None:
    if cfg.offline_simulation:
        print("[离线模拟] 发生器已停止")
        return
    with RigolSCPI(cfg.rigol_visa_resource) as rigol:
        print("Instrument:", rigol.identify())
        rigol.abort_all()
        print("发生器已停止")


def start_generator_channel(cfg: IntegratedExperimentConfig, channel: int) -> None:
    if channel not in (1, 2):
        raise ValueError("channel 必须是 1 或 2")
    channel_cfg = cfg.ch1 if channel == 1 else cfg.ch2
    if cfg.offline_simulation:
        print(f"[离线模拟] CH{channel} 独立输出已打开")
        return
    with RigolSCPI(cfg.rigol_visa_resource) as rigol:
        print("Instrument:", rigol.identify())
        rigol.configure_continuous_output(channel, channel_cfg)
        rigol.set_output(channel, True)
        print(f"CH{channel} 已独立打开")


def stop_generator_channel(cfg: IntegratedExperimentConfig, channel: int) -> None:
    if channel not in (1, 2):
        raise ValueError("channel 必须是 1 或 2")
    if cfg.offline_simulation:
        print(f"[离线模拟] CH{channel} 独立输出已关闭")
        return
    with RigolSCPI(cfg.rigol_visa_resource) as rigol:
        print("Instrument:", rigol.identify())
        rigol.set_output(channel, False)
        print(f"CH{channel} 已独立关闭")


def test_andor_connection(cfg: IntegratedExperimentConfig, andor: AndorSDKController | None = None) -> None:
    if cfg.offline_simulation:
        print("[离线模拟] 已跳过 Andor 硬件连接测试")
        print(f"Andor SDK 根目录: {cfg.andor_sdk_root}")
        return
    try:
        ctx = nullcontext(andor) if andor is not None else AndorSDKController(cfg.andor_sdk_root)
        with ctx as andor:
            xpixels, ypixels, px_w, px_h = andor.get_detector_info()
            temperature_c, temp_code = andor.get_temperature_c()
            print("Andor 连接测试通过")
            print(f"Andor SDK 根目录: {cfg.andor_sdk_root}")
            print(f"相机数量: {andor.get_camera_count()}")
            print(f"光谱仪数量: {andor.get_shamrock_count()}")
            print(f"探测器尺寸: {xpixels} x {ypixels} 像素")
            print(f"像元尺寸: {px_w:.3f} um x {px_h:.3f} um")
            print(f"当前温度: {temperature_c:.2f} C (状态码 {temp_code})")
            try:
                current_grating = andor.get_current_grating()
                current_wavelength_nm = andor.get_current_wavelength_nm()
                coeffs = andor.get_pixel_calibration_coefficients()
                print(f"当前光栅编号: {current_grating}")
                print(f"当前中心波长: {current_wavelength_nm:.6f} nm")
                print(
                    "Pixel calibration coefficients: "
                    f"A={coeffs[0]:.8g}, B={coeffs[1]:.8g}, C={coeffs[2]:.8g}, D={coeffs[3]:.8g}"
                )
            except AndorSDKError as exc:
                print(f"警告：Shamrock 已枚举到设备，但读取当前光栅/波长失败：{exc}")
                print("这通常表示光谱仪通信仍不稳定，请检查 Shamrock USB、供电、驱动和开机顺序。")
    except (AndorHardwareNotFoundError, ShamrockHardwareNotFoundError) as exc:
        raise RuntimeError(str(exc)) from exc


def start_andor_debug(
    cfg: IntegratedExperimentConfig,
    progress_callback=noop_progress_callback,
    andor: AndorSDKController | None = None,
    spectrum_saved_callback=noop_spectrum_saved_callback,
) -> Path:
    sample_output_dir(cfg.output_dir, cfg.sample_name).mkdir(parents=True, exist_ok=True)
    output_file = debug_acquisition_filename(cfg.output_dir, cfg.sample_name)
    try:
        if cfg.offline_simulation:
            progress_callback(0.0, "离线模拟曝光中")
            x_axis, y_axis = build_simulated_spectrum(cfg.ch1.phase_deg)
            progress_callback(1.0, "离线模拟曝光完成")
            save_ascii_data(output_file, x_axis, y_axis)
            analysis = analyze_ghost_peak(x_axis, y_axis)
            output_file = apply_ghost_filename_marker(output_file, analysis)
            spectrum_saved_callback(output_file)
            if analysis.get("ghost_detected"):
                print(analysis["message"])
            analysis = analyze_ghost_peak(x_axis, y_axis)
            output_file = apply_ghost_filename_marker(output_file, analysis)
            spectrum_saved_callback(output_file)
            print(f"[离线模拟] 光谱仪调试数据已生成: {output_file}")
            return output_file
        ctx = nullcontext(andor) if andor is not None else AndorSDKController(cfg.andor_sdk_root)
        with ctx as andor:
            andor.configure_spectrograph(
                grating_no=cfg.spectrometer.grating_no,
                center_wavelength_nm=cfg.spectrometer.center_wavelength_nm,
                slit_width_um=cfg.spectrometer.slit_width_um,
                shamrock_shutter_mode=cfg.spectrometer.shamrock_shutter_mode,
            )
            andor.enable_cooler_and_set_target(cfg.spectrometer.target_temperature_c)
            wait_for_detector_temperature(andor, cfg.spectrometer)
            readout_info = andor.configure_camera_readout(
                horizontal_readout_mhz=cfg.spectrometer.horizontal_readout_mhz,
                pre_amp_gain_value=cfg.spectrometer.pre_amp_gain,
                output_amplifier=resolve_output_amplifier(cfg.spectrometer.output_amplifier),
                ad_channel=cfg.spectrometer.ad_channel,
            )
            x_axis_wavelength_nm, y_axis, actual_exp = andor.acquire_spectrum(
                exposure_s=cfg.spectrometer.exposure_s,
                trigger_mode=0,
                progress_callback=lambda value: progress_callback(value, f"光谱仪调试曝光进度 {value * 100:.0f}%"),
                camera_shutter_mode=cfg.spectrometer.camera_shutter_mode,
                camera_shutter_open_ms=cfg.spectrometer.camera_shutter_open_ms,
                camera_shutter_close_ms=cfg.spectrometer.camera_shutter_close_ms,
            )
            x_axis_raman_shift = convert_wavelength_axis_to_raman_shift(
                x_axis_wavelength_nm,
                cfg.spectrometer.rayleigh_wavelength_nm,
            )
            save_ascii_data(output_file, x_axis_raman_shift, y_axis)
            analysis = analyze_ghost_peak(x_axis_raman_shift, y_axis)
            output_file = apply_ghost_filename_marker(output_file, analysis)
            spectrum_saved_callback(output_file)
            if analysis.get("ghost_detected"):
                print(analysis["message"])
            print(
                f"光谱仪调试数据已保存: {output_file} "
                f"(实际曝光 {actual_exp:.6f}s, 读出 {readout_info['horizontal_speed_mhz']:.2f} MHz, "
                f"Pre-Amp {readout_info['preamp_gain_value']:.2f}x)"
            )
    except (AndorHardwareNotFoundError, ShamrockHardwareNotFoundError) as exc:
        raise RuntimeError(str(exc)) from exc
    return output_file


def run_continuous_preview_acquisition(
    cfg: IntegratedExperimentConfig,
    stop_event: threading.Event,
    progress_callback=noop_progress_callback,
    andor: AndorSDKController | None = None,
    spectrum_saved_callback=noop_spectrum_saved_callback,
) -> Path:
    sample_output_dir(cfg.output_dir, cfg.sample_name).mkdir(parents=True, exist_ok=True)
    output_file = continuous_preview_filename(cfg.output_dir, cfg.sample_name)
    frame_index = 0
    try:
        if cfg.offline_simulation:
            print("[离线模拟] 连续预览采集启动，固定覆盖 _continuous_preview.asc")
            while not stop_event.is_set():
                frame_index += 1
                phase = cfg.ch1.phase_deg + frame_index * 3.0
                progress_callback(0.0, f"连续预览 第 {frame_index} 帧模拟曝光")
                x_axis, y_axis = build_simulated_spectrum(phase)
                save_ascii_data(output_file, x_axis, y_axis)
                spectrum_saved_callback(output_file)
                progress_callback(1.0, f"连续预览 第 {frame_index} 帧已刷新")
                stop_event.wait(max(0.1, cfg.spectrometer.exposure_s))
            print("[离线模拟] 连续预览采集已停止")
            return output_file

        ctx = nullcontext(andor) if andor is not None else AndorSDKController(cfg.andor_sdk_root)
        with ctx as andor:
            andor.configure_spectrograph(
                grating_no=cfg.spectrometer.grating_no,
                center_wavelength_nm=cfg.spectrometer.center_wavelength_nm,
                slit_width_um=cfg.spectrometer.slit_width_um,
                shamrock_shutter_mode=cfg.spectrometer.shamrock_shutter_mode,
            )
            andor.enable_cooler_and_set_target(cfg.spectrometer.target_temperature_c)
            wait_for_detector_temperature(andor, cfg.spectrometer)
            readout_info = andor.configure_camera_readout(
                horizontal_readout_mhz=cfg.spectrometer.horizontal_readout_mhz,
                pre_amp_gain_value=cfg.spectrometer.pre_amp_gain,
                output_amplifier=resolve_output_amplifier(cfg.spectrometer.output_amplifier),
                ad_channel=cfg.spectrometer.ad_channel,
            )
            print(
                "连续预览采集启动: 固定覆盖 _continuous_preview.asc "
                f"(读出 {readout_info['horizontal_speed_mhz']:.2f} MHz, "
                f"Pre-Amp {readout_info['preamp_gain_value']:.2f}x)"
            )
            while not stop_event.is_set():
                frame_index += 1
                x_axis_wavelength_nm, y_axis, actual_exp = andor.acquire_spectrum(
                    exposure_s=cfg.spectrometer.exposure_s,
                    trigger_mode=0,
                    progress_callback=lambda value, i=frame_index: progress_callback(
                        value, f"连续预览 第 {i} 帧曝光进度 {value * 100:.0f}%"
                    ),
                    camera_shutter_mode=cfg.spectrometer.camera_shutter_mode,
                    camera_shutter_open_ms=cfg.spectrometer.camera_shutter_open_ms,
                    camera_shutter_close_ms=cfg.spectrometer.camera_shutter_close_ms,
                )
                x_axis_raman_shift = convert_wavelength_axis_to_raman_shift(
                    x_axis_wavelength_nm,
                    cfg.spectrometer.rayleigh_wavelength_nm,
                )
                save_ascii_data(output_file, x_axis_raman_shift, y_axis)
                spectrum_saved_callback(output_file)
                progress_callback(1.0, f"连续预览 第 {frame_index} 帧已刷新 ({actual_exp:.3f}s)")
            print("连续预览采集已停止")
    except (AndorHardwareNotFoundError, ShamrockHardwareNotFoundError) as exc:
        raise RuntimeError(str(exc)) from exc
    return output_file


def run_baseline_test(
    cfg: IntegratedExperimentConfig,
    progress_callback=noop_progress_callback,
    andor: AndorSDKController | None = None,
    spectrum_saved_callback=noop_spectrum_saved_callback,
) -> None:
    sample_dir = sample_output_dir(cfg.output_dir, cfg.sample_name)
    sample_dir.mkdir(parents=True, exist_ok=True)
    RUNTIME.pause_event.clear()

    if cfg.offline_simulation:
        print("[离线模拟] 基线测试仅启动 CH1(AOM)")
        print("[离线模拟] 输出目录:", sample_dir)
        print("[离线模拟] 已打开 CH1(AOM) 独立输出，等待稳定")
        wait_with_pause(cfg.settle_time_s, cfg, [0.0], 0, 1)
        for rep in range(1, cfg.repeats_per_phase + 1):
            ensure_not_paused(cfg, [0.0], 0, rep)
            output_file = baseline_acquisition_filename(cfg.output_dir, cfg.sample_name, rep)
            progress_callback(0.0, f"基线测试 第 {rep} 次曝光中")
            x_axis, y_axis = build_simulated_spectrum(rep * 0.2)
            progress_callback(1.0, f"基线测试 第 {rep} 次曝光完成")
            save_ascii_data(output_file, x_axis, y_axis)
            print(f"  Baseline {rep}: saved {output_file.name} [离线模拟]")
        print("[离线模拟] 基线测试完成，CH1 已关闭")
        return

    try:
        andor_ctx = nullcontext(andor) if andor is not None else AndorSDKController(cfg.andor_sdk_root)
        with andor_ctx as andor, RigolSCPI(cfg.rigol_visa_resource) as rigol:
            RUNTIME.current_andor = andor
            RUNTIME.current_rigol = rigol
            print("Instrument:", rigol.identify())
            print("Andor SDK root:", cfg.andor_sdk_root)
            andor.configure_spectrograph(
                grating_no=cfg.spectrometer.grating_no,
                center_wavelength_nm=cfg.spectrometer.center_wavelength_nm,
                slit_width_um=cfg.spectrometer.slit_width_um,
                shamrock_shutter_mode=cfg.spectrometer.shamrock_shutter_mode,
            )
            andor.enable_cooler_and_set_target(cfg.spectrometer.target_temperature_c)
            wait_for_detector_temperature(andor, cfg.spectrometer)
            readout_info = andor.configure_camera_readout(
                horizontal_readout_mhz=cfg.spectrometer.horizontal_readout_mhz,
                pre_amp_gain_value=cfg.spectrometer.pre_amp_gain,
                output_amplifier=resolve_output_amplifier(cfg.spectrometer.output_amplifier),
                ad_channel=cfg.spectrometer.ad_channel,
            )
            print(
                f"读出设置: {readout_info['horizontal_speed_mhz']:.2f} MHz, "
                f"Pre-Amp {readout_info['preamp_gain_value']:.2f}x, "
                f"{cfg.spectrometer.output_amplifier}"
            )
            rigol.configure_continuous_output(1, cfg.ch1)
            rigol.wait_operation_complete()
            rigol.set_output(1, True)
            print("已打开 CH1(AOM) 独立输出，等待稳定")
            wait_with_pause(cfg.settle_time_s, cfg, [0.0], 0, 1)

            for rep in range(1, cfg.repeats_per_phase + 1):
                ensure_not_paused(cfg, [0.0], 0, rep)
                output_file = baseline_acquisition_filename(cfg.output_dir, cfg.sample_name, rep)
                try:
                    x_axis_wavelength_nm, y_axis, actual_exp = andor.acquire_spectrum(
                        exposure_s=cfg.spectrometer.exposure_s,
                        trigger_mode=0,
                        progress_callback=lambda value, r=rep: progress_callback(
                            value, f"基线测试 第 {r} 次曝光进度 {value * 100:.0f}%"
                        ),
                        camera_shutter_mode=cfg.spectrometer.camera_shutter_mode,
                        camera_shutter_open_ms=cfg.spectrometer.camera_shutter_open_ms,
                        camera_shutter_close_ms=cfg.spectrometer.camera_shutter_close_ms,
                    )
                except AndorSDKError as exc:
                    if RUNTIME.is_pause_requested():
                        raise ExperimentPaused("基线测试已暂停。再次点击“基线测试”会重新开始。") from exc
                    raise
                x_axis_raman_shift = convert_wavelength_axis_to_raman_shift(
                    x_axis_wavelength_nm,
                    cfg.spectrometer.rayleigh_wavelength_nm,
                )
                save_ascii_data(output_file, x_axis_raman_shift, y_axis)
                analysis = analyze_ghost_peak(x_axis_raman_shift, y_axis)
                output_file = apply_ghost_filename_marker(output_file, analysis)
                spectrum_saved_callback(output_file)
                if analysis.get("ghost_detected"):
                    print(analysis["message"])
                print(f"  Baseline {rep}: saved {output_file.name} (实际曝光 {actual_exp:.6f}s)")

            rigol.abort_all()
            print("基线测试完成，CH1 已关闭")
    except (AndorHardwareNotFoundError, ShamrockHardwareNotFoundError) as exc:
        raise RuntimeError(str(exc)) from exc
    except ExperimentPaused:
        raise
    finally:
        if RUNTIME.current_rigol is not None:
            try:
                RUNTIME.current_rigol.abort_all()
            except Exception:
                pass
        RUNTIME.reset()


def request_pause_experiment() -> None:
    RUNTIME.request_pause()


def stop_paused_experiment(cfg: IntegratedExperimentConfig) -> None:
    clear_resume_state(cfg)
    RUNTIME.reset()


def wait_for_detector_temperature(andor: AndorSDKController, spec: SpectrometerConfig) -> None:
    deadline = time.time() + max(1.0, spec.cooldown_timeout_s)
    tolerance_c = 0.2
    while True:
        current_temp, status_code = andor.get_temperature_c()
        print(
            f"探测器温度: {current_temp:.2f} C / 目标 {spec.target_temperature_c} C "
            f"(要求不高于 {spec.required_temperature_c} C)"
        )
        temp_ready = current_temp <= (spec.required_temperature_c + tolerance_c)
        if temp_ready and status_code == DRV_TEMP_STABILIZED:
            print("探测器已冷却并稳定，可开始采集")
            return
        if temp_ready and status_code != DRV_TEMP_STABILIZED:
            print("探测器温度已达到要求，等待温度稳定")
        if time.time() >= deadline:
            raise RuntimeError(
                f"探测器未能在限定时间内冷却到 {spec.required_temperature_c} C 并稳定，当前温度 {current_temp:.2f} C"
            )
        time.sleep(5.0)


def run_integrated_experiment(
    cfg: IntegratedExperimentConfig,
    progress_callback=noop_progress_callback,
    andor: AndorSDKController | None = None,
    spectrum_saved_callback=noop_spectrum_saved_callback,
) -> None:
    sample_dir = sample_output_dir(cfg.output_dir, cfg.sample_name)
    sample_dir.mkdir(parents=True, exist_ok=True)
    phases = build_phase_list(cfg.phase_start_deg, cfg.phase_stop_deg, cfg.phase_step_deg)
    start_phase_index, start_rep_index = load_resume_state(cfg, phases)
    RUNTIME.pause_event.clear()

    if cfg.offline_simulation:
        print("[离线模拟] 不连接信号发生器和光谱仪，按实验流程生成模拟数据")
        print("[离线模拟] 输出目录:", sample_dir)
        if pause_state_path(cfg.output_dir, cfg.sample_name).exists():
            print(f"[离线模拟] 继续未完成实验: 相位 {phases[start_phase_index]:g}°, 第 {start_rep_index} 次")
        for phase_index in range(start_phase_index, len(phases)):
            phase = phases[phase_index]
            ch1_output_phase_deg = resolve_ch1_output_phase_deg(phase)
            rep_begin = start_rep_index if phase_index == start_phase_index else 1
            ensure_not_paused(cfg, phases, phase_index, rep_begin)
            print(f"Phase {phase:.2f} deg (CH1 实际输出相位 {ch1_output_phase_deg:.2f} deg)")
            print("  [离线模拟] 打开 CH1/CH2，随后由 CH1 手动触发带动 CH2 外部触发启动")
            wait_with_pause(cfg.settle_time_s, cfg, phases, phase_index, rep_begin)
            for rep in range(rep_begin, cfg.repeats_per_phase + 1):
                ensure_not_paused(cfg, phases, phase_index, rep)
                output_file = acquisition_filename(cfg.output_dir, cfg.sample_name, phase, rep)
                progress_callback(0.0, f"相位 {phase:g}°, 第 {rep} 次曝光中")
                x_axis, y_axis = build_simulated_spectrum(phase + rep * 0.2)
                progress_callback(1.0, f"相位 {phase:g}°, 第 {rep} 次曝光完成")
                save_ascii_data(output_file, x_axis, y_axis)
                analysis = analyze_ghost_peak(x_axis, y_axis)
                output_file = apply_ghost_filename_marker(output_file, analysis)
                spectrum_saved_callback(output_file)
                if analysis.get("ghost_detected"):
                    print(analysis["message"])
                print(f"  Rep {rep}: saved {output_file.name} [离线模拟]")
                next_phase_index = phase_index + 1 if rep == cfg.repeats_per_phase else phase_index
                next_rep_index = 1 if rep == cfg.repeats_per_phase else rep + 1
                if next_phase_index < len(phases):
                    save_resume_state(cfg, phases, next_phase_index, next_rep_index)
            print("  [离线模拟] 当前相位采集完成，关闭发生器")
        clear_resume_state(cfg)
        print("[离线模拟] 实验完成")
        return

    try:
        andor_ctx = nullcontext(andor) if andor is not None else AndorSDKController(cfg.andor_sdk_root)
        with andor_ctx as andor, RigolSCPI(cfg.rigol_visa_resource) as rigol:
            RUNTIME.current_andor = andor
            RUNTIME.current_rigol = rigol
            print("Instrument:", rigol.identify())
            print("Andor SDK root:", cfg.andor_sdk_root)
            if pause_state_path(cfg.output_dir, cfg.sample_name).exists():
                print(f"继续未完成实验: 相位 {phases[start_phase_index]:g}°, 第 {start_rep_index} 次")
            andor.configure_spectrograph(
                grating_no=cfg.spectrometer.grating_no,
                center_wavelength_nm=cfg.spectrometer.center_wavelength_nm,
                slit_width_um=cfg.spectrometer.slit_width_um,
                shamrock_shutter_mode=cfg.spectrometer.shamrock_shutter_mode,
            )
            andor.enable_cooler_and_set_target(cfg.spectrometer.target_temperature_c)
            wait_for_detector_temperature(andor, cfg.spectrometer)
            readout_info = andor.configure_camera_readout(
                horizontal_readout_mhz=cfg.spectrometer.horizontal_readout_mhz,
                pre_amp_gain_value=cfg.spectrometer.pre_amp_gain,
                output_amplifier=resolve_output_amplifier(cfg.spectrometer.output_amplifier),
                ad_channel=cfg.spectrometer.ad_channel,
            )
            print(
                f"读出设置: {readout_info['horizontal_speed_mhz']:.2f} MHz, "
                f"Pre-Amp {readout_info['preamp_gain_value']:.2f}x, "
                f"{cfg.spectrometer.output_amplifier}"
            )

            generator_full_config_needed = True
            for phase_index in range(start_phase_index, len(phases)):
                phase = phases[phase_index]
                rep_begin = start_rep_index if phase_index == start_phase_index else 1
                ensure_not_paused(cfg, phases, phase_index, rep_begin)
                configure_generator_for_phase(rigol, cfg, phase, full_config=generator_full_config_needed)
                generator_full_config_needed = False
                wait_with_pause(POST_OUTPUT_SETTLE_S, cfg, phases, phase_index, rep_begin)
                rigol.fire_manual_trigger_ch1()
                wait_with_pause(cfg.settle_time_s, cfg, phases, phase_index, rep_begin)

                for rep in range(rep_begin, cfg.repeats_per_phase + 1):
                    ensure_not_paused(cfg, phases, phase_index, rep)
                    output_file = acquisition_filename(cfg.output_dir, cfg.sample_name, phase, rep)
                    try:
                        x_axis_wavelength_nm, y_axis, actual_exp = andor.acquire_spectrum(
                            exposure_s=cfg.spectrometer.exposure_s,
                            trigger_mode=0,
                            progress_callback=lambda value, p=phase, r=rep: progress_callback(
                                value, f"相位 {p:g}°, 第 {r} 次曝光进度 {value * 100:.0f}%"
                            ),
                            camera_shutter_mode=cfg.spectrometer.camera_shutter_mode,
                            camera_shutter_open_ms=cfg.spectrometer.camera_shutter_open_ms,
                            camera_shutter_close_ms=cfg.spectrometer.camera_shutter_close_ms,
                        )
                    except AndorSDKError as exc:
                        if RUNTIME.is_pause_requested():
                            save_resume_state(cfg, phases, phase_index, rep)
                            raise ExperimentPaused(
                                f"实验已暂停，将从相位 {phase:g}°, 第 {rep} 次继续。"
                            ) from exc
                        raise
                    x_axis_raman_shift = convert_wavelength_axis_to_raman_shift(
                        x_axis_wavelength_nm,
                        cfg.spectrometer.rayleigh_wavelength_nm,
                    )
                    save_ascii_data(output_file, x_axis_raman_shift, y_axis)
                    analysis = analyze_ghost_peak(x_axis_raman_shift, y_axis)
                    output_file = apply_ghost_filename_marker(output_file, analysis)
                    spectrum_saved_callback(output_file)
                    if analysis.get("ghost_detected"):
                        print(analysis["message"])
                    print(f"  Rep {rep}: saved {output_file.name} (实际曝光 {actual_exp:.6f}s)")
                    next_phase_index = phase_index + 1 if rep == cfg.repeats_per_phase else phase_index
                    next_rep_index = 1 if rep == cfg.repeats_per_phase else rep + 1
                    if next_phase_index < len(phases):
                        save_resume_state(cfg, phases, next_phase_index, next_rep_index)

                rigol.abort_all()
                print("  当前相位采集完成，发生器已关闭")

        clear_resume_state(cfg)
        print("实验完成")
    except (AndorHardwareNotFoundError, ShamrockHardwareNotFoundError) as exc:
        raise RuntimeError(str(exc)) from exc
    except ExperimentPaused:
        raise
    finally:
        if RUNTIME.current_rigol is not None:
            try:
                RUNTIME.current_rigol.abort_all()
            except Exception:
                pass
        RUNTIME.reset()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Integrated TR-Raman controller for RIGOL VISA + Andor SDK")
    parser.add_argument("--config", help="Path to JSON config file")
    parser.add_argument("--visa-resource", default="")
    parser.add_argument("--andor-sdk-root", default="")
    parser.add_argument("--output-dir", default=r"C:\AndorOutput")
    parser.add_argument("--sample-name", default="sample")
    parser.add_argument("--phase-start", type=float, default=0.0)
    parser.add_argument("--phase-stop", type=float, default=360.0)
    parser.add_argument("--phase-step", type=float, default=10.0)
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--ch1-delay", type=float, default=0.0)
    parser.add_argument("--ch1-freq", type=float, default=1000.0)
    parser.add_argument("--ch1-amp", type=float, default=1.0)
    parser.add_argument("--ch2-freq", type=float, default=1000.0)
    parser.add_argument("--ch2-amp", type=float, default=5.0)
    parser.add_argument("--center-wavelength", type=float, default=500.0)
    parser.add_argument("--grating", type=int, default=1)
    parser.add_argument("--exposure", type=float, default=0.2)
    parser.add_argument("--trigger-mode", type=int, default=TRIGGER_EXTERNAL)
    return parser.parse_args()


def load_json_config(config_path: Path) -> IntegratedExperimentConfig:
    data = json.loads(config_path.read_text(encoding="utf-8-sig"))
    return IntegratedExperimentConfig(
        rigol_visa_resource=data.get("rigol_visa_resource", ""),
        andor_sdk_root=resolve_andor_sdk_root(data.get("andor_sdk_root", "")),
        output_dir=Path(data["output_dir"]),
        sample_name=data["sample_name"],
        phase_start_deg=float(data["phase_start_deg"]),
        phase_stop_deg=float(data["phase_stop_deg"]),
        phase_step_deg=float(data["phase_step_deg"]),
        repeats_per_phase=int(data["repeats_per_phase"]),
        ch1_start_delay_s=float(data.get("ch1_start_delay_ms", data.get("ch1_start_delay_s", 0.0))) / (
            1000.0 if "ch1_start_delay_ms" in data else 1.0
        ),
        ch1=ChannelConfig(**data["ch1"]),
        ch2=ChannelConfig(**data["ch2"]),
        spectrometer=SpectrometerConfig(
            rayleigh_wavelength_nm=float(
                data["spectrometer"].get(
                    "rayleigh_wavelength_nm",
                    data["spectrometer"].get("excitation_wavelength_nm", 532.0),
                )
            ),
            center_wavelength_nm=float(data["spectrometer"]["center_wavelength_nm"]),
            grating_no=int(data["spectrometer"]["grating_no"]),
            exposure_s=float(data["spectrometer"].get("exposure_ms", data["spectrometer"].get("exposure_s", 0.2)))
            / (1000.0 if "exposure_ms" in data["spectrometer"] else 1.0),
            trigger_mode=int(data["spectrometer"]["trigger_mode"]),
            slit_width_um=(
                float(data["spectrometer"]["slit_width_um"])
                if data["spectrometer"].get("slit_width_um") is not None
                else None
            ),
            target_temperature_c=int(data["spectrometer"].get("target_temperature_c", -60)),
            required_temperature_c=int(data["spectrometer"].get("required_temperature_c", -60)),
            cooldown_timeout_s=float(data["spectrometer"].get("cooldown_timeout_s", 3600.0)),
            pre_amp_gain=(
                float(data["spectrometer"]["pre_amp_gain"])
                if data["spectrometer"].get("pre_amp_gain") is not None
                else None
            ),
            horizontal_readout_mhz=(
                float(data["spectrometer"]["horizontal_readout_mhz"])
                if data["spectrometer"].get("horizontal_readout_mhz") is not None
                else None
            ),
            output_amplifier=str(data["spectrometer"].get("output_amplifier", "conventional")),
            ad_channel=int(data["spectrometer"].get("ad_channel", 0)),
            camera_shutter_mode=(
                int(data["spectrometer"]["camera_shutter_mode"])
                if data["spectrometer"].get("camera_shutter_mode") is not None
                else None
            ),
            camera_shutter_open_ms=int(data["spectrometer"].get("camera_shutter_open_ms", 0)),
            camera_shutter_close_ms=int(data["spectrometer"].get("camera_shutter_close_ms", 0)),
            shamrock_shutter_mode=(
                int(data["spectrometer"]["shamrock_shutter_mode"])
                if data["spectrometer"].get("shamrock_shutter_mode") is not None
                else None
            ),
        ),
        settle_time_s=float(data.get("settle_time_ms", data.get("settle_time_s", 0.2)))
        / (1000.0 if "settle_time_ms" in data else 1.0),
        offline_simulation=bool(data.get("offline_simulation", False)),
    )


def build_config(args: argparse.Namespace) -> IntegratedExperimentConfig:
    return IntegratedExperimentConfig(
        rigol_visa_resource=args.visa_resource,
        andor_sdk_root=resolve_andor_sdk_root(args.andor_sdk_root),
        output_dir=Path(args.output_dir),
        sample_name=args.sample_name,
        phase_start_deg=args.phase_start,
        phase_stop_deg=args.phase_stop,
        phase_step_deg=args.phase_step,
        repeats_per_phase=args.repeats,
        ch1_start_delay_s=args.ch1_delay,
        ch1=ChannelConfig(
            waveform="square",
            frequency_hz=args.ch1_freq,
            amplitude_vpp=args.ch1_amp,
        ),
        ch2=ChannelConfig(
            waveform="sin",
            frequency_hz=args.ch2_freq,
            amplitude_vpp=args.ch2_amp,
        ),
        spectrometer=SpectrometerConfig(
            rayleigh_wavelength_nm=532.0,
            center_wavelength_nm=args.center_wavelength,
            grating_no=args.grating,
            exposure_s=args.exposure,
            trigger_mode=args.trigger_mode,
            target_temperature_c=-60,
            required_temperature_c=-60,
            pre_amp_gain=2.0,
            horizontal_readout_mhz=1.48,
            output_amplifier="conventional",
        ),
        offline_simulation=False,
    )


if __name__ == "__main__":
    try:
        args = parse_args()
        if args.config:
            config = load_json_config(Path(args.config))
        else:
            default_config = find_default_config_path()
            config = load_json_config(default_config) if default_config is not None else build_config(args)
        run_integrated_experiment(config)
    except Exception as exc:
        print(f"Startup failed: {exc}")
        print("Tip: edit app_config.json or pass --config explicitly.")
        if getattr(sys, "frozen", False):
            try:
                input("Press Enter to close...")
            except EOFError:
                pass
        raise
