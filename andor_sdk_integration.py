from __future__ import annotations

import ctypes
import os
import time
from contextlib import contextmanager
from ctypes import byref, c_char_p, c_float, c_int, c_long, create_unicode_buffer
from pathlib import Path


DRV_SUCCESS = 20002
DRV_VXDNOTINSTALLED = 20003
DRV_NOT_INITIALIZED = 20075
DRV_ERROR_NOCAMERA = 20990
DRV_NOT_AVAILABLE = 20992
DRV_TEMP_OFF = 20034
DRV_TEMP_NOT_STABILIZED = 20035
DRV_TEMP_STABILIZED = 20036
DRV_TEMP_NOT_REACHED = 20037
DRV_TEMP_OUT_RANGE = 20038
DRV_TEMP_NOT_SUPPORTED = 20039
DRV_TEMP_DRIFT = 20040
SHAMROCK_SUCCESS = 20202

READ_MODE_FVB = 0
ACQ_MODE_SINGLE = 1
OUTPUT_AMPLIFIER_EMCCD = 0
OUTPUT_AMPLIFIER_CONVENTIONAL = 1

CAMERA_ERROR_NAMES = {
    DRV_SUCCESS: "DRV_SUCCESS",
    DRV_VXDNOTINSTALLED: "DRV_VXDNOTINSTALLED",
    DRV_NOT_INITIALIZED: "DRV_NOT_INITIALIZED",
    DRV_ERROR_NOCAMERA: "DRV_ERROR_NOCAMERA",
    DRV_NOT_AVAILABLE: "DRV_NOT_AVAILABLE",
}

SHAMROCK_ERROR_NAMES = {
    20201: "SHAMROCK_COMMUNICATION_ERROR",
    SHAMROCK_SUCCESS: "SHAMROCK_SUCCESS",
    20275: "SHAMROCK_NOT_INITIALIZED",
    20292: "SHAMROCK_NOT_AVAILABLE",
}


class AndorSDKError(RuntimeError):
    pass


class AndorHardwareNotFoundError(AndorSDKError):
    pass


class ShamrockHardwareNotFoundError(AndorSDKError):
    pass


class AndorSDKController:
    def __init__(self, sdk_root: Path) -> None:
        self.sdk_root = self._normalize_path(sdk_root.resolve())
        self._dll_dirs = []
        self._atmcd = None
        self._shamrock = None
        self._camera_initialized = False
        self._shamrock_initialized = False
        self._available_cameras = 0
        self._available_shamrocks = 0

    @staticmethod
    def _normalize_path(path: Path) -> Path:
        try:
            kernel32 = ctypes.windll.kernel32
            short_size = kernel32.GetShortPathNameW(str(path), None, 0)
            if short_size:
                buffer = create_unicode_buffer(short_size)
                result = kernel32.GetShortPathNameW(str(path), buffer, short_size)
                if result:
                    return Path(buffer.value)
        except Exception:
            pass
        return path

    def __enter__(self) -> "AndorSDKController":
        self.open()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    @property
    def drivers_dir(self) -> Path:
        return self._normalize_path(self.sdk_root / "Drivers")

    @property
    def shamrock_dir(self) -> Path:
        return self._normalize_path(self.drivers_dir / "Shamrock64")

    def _register_dll_dir(self, path: Path) -> None:
        path = self._normalize_path(path)
        if path.exists():
            self._dll_dirs.append(os.add_dll_directory(str(path)))

    def _camera_call(self, name: str, *args) -> int:
        if self._atmcd is None:
            raise AndorSDKError("Camera SDK not loaded")
        func = getattr(self._atmcd, name)
        code = int(func(*args))
        if code != DRV_SUCCESS:
            error_name = CAMERA_ERROR_NAMES.get(code, "UNKNOWN_CAMERA_ERROR")
            raise AndorSDKError(f"{name} failed with code {code} ({error_name})")
        return code

    def _shamrock_call(self, name: str, *args) -> int:
        if self._shamrock is None:
            raise AndorSDKError("Shamrock SDK not loaded")
        func = getattr(self._shamrock, name)
        last_code = SHAMROCK_SUCCESS
        for attempt in range(3):
            code = int(func(*args))
            last_code = code
            if code == SHAMROCK_SUCCESS:
                return code
            if code == 20201 and attempt < 2:
                time.sleep(0.2)
                continue
            break
        error_name = SHAMROCK_ERROR_NAMES.get(last_code, "UNKNOWN_SHAMROCK_ERROR")
        raise AndorSDKError(f"{name} failed with code {last_code} ({error_name})")

    @contextmanager
    def _sdk_working_directory(self):
        previous = Path.cwd()
        os.chdir(self.sdk_root)
        try:
            yield
        finally:
            os.chdir(previous)

    def _load_libraries(self) -> None:
        self._register_dll_dir(self.sdk_root)
        self._register_dll_dir(self.drivers_dir)
        self._register_dll_dir(self.shamrock_dir)
        self._register_dll_dir(self.sdk_root / "Shamrock USB Drivers")
        self._register_dll_dir(self.sdk_root / "Shamrock USB Drivers" / "amd64")
        self._register_dll_dir(self.sdk_root / "Device Driver" / "USB" / "amd64")
        self._register_dll_dir(self.sdk_root / "Device Driver" / "WinUSB" / "amd64")

        atmcd_path = self._normalize_path(self.sdk_root / "atmcd64d.dll")
        shamrock_path = self._normalize_path(self.shamrock_dir / "ShamrockCIF.dll")
        if not atmcd_path.exists():
            raise FileNotFoundError(f"Camera DLL not found: {atmcd_path}")
        if not shamrock_path.exists():
            raise FileNotFoundError(f"Shamrock DLL not found: {shamrock_path}")

        self._atmcd = ctypes.WinDLL(str(atmcd_path))
        self._shamrock = ctypes.WinDLL(str(shamrock_path))

        self._atmcd.Initialize.argtypes = [c_char_p]
        self._atmcd.Initialize.restype = c_int
        self._atmcd.CoolerON.argtypes = []
        self._atmcd.CoolerON.restype = c_int
        self._atmcd.CoolerOFF.argtypes = []
        self._atmcd.CoolerOFF.restype = c_int
        self._atmcd.ShutDown.argtypes = []
        self._atmcd.ShutDown.restype = c_int
        self._atmcd.GetAvailableCameras.argtypes = [ctypes.POINTER(c_int)]
        self._atmcd.GetAvailableCameras.restype = c_int
        self._atmcd.GetTemperatureF.argtypes = [ctypes.POINTER(c_float)]
        self._atmcd.GetTemperatureF.restype = c_int
        self._atmcd.SetTemperature.argtypes = [c_int]
        self._atmcd.SetTemperature.restype = c_int
        self._atmcd.SetADChannel.argtypes = [c_int]
        self._atmcd.SetADChannel.restype = c_int
        self._atmcd.GetNumberADChannels.argtypes = [ctypes.POINTER(c_int)]
        self._atmcd.GetNumberADChannels.restype = c_int
        self._atmcd.GetBitDepth.argtypes = [c_int, ctypes.POINTER(c_int)]
        self._atmcd.GetBitDepth.restype = c_int
        self._atmcd.SetOutputAmplifier.argtypes = [c_int]
        self._atmcd.SetOutputAmplifier.restype = c_int
        self._atmcd.GetNumberHSSpeeds.argtypes = [c_int, c_int, ctypes.POINTER(c_int)]
        self._atmcd.GetNumberHSSpeeds.restype = c_int
        self._atmcd.GetHSSpeed.argtypes = [c_int, c_int, c_int, ctypes.POINTER(c_float)]
        self._atmcd.GetHSSpeed.restype = c_int
        self._atmcd.SetHSSpeed.argtypes = [c_int, c_int]
        self._atmcd.SetHSSpeed.restype = c_int
        self._atmcd.GetNumberPreAmpGains.argtypes = [ctypes.POINTER(c_int)]
        self._atmcd.GetNumberPreAmpGains.restype = c_int
        self._atmcd.GetPreAmpGain.argtypes = [c_int, ctypes.POINTER(c_float)]
        self._atmcd.GetPreAmpGain.restype = c_int
        self._atmcd.SetPreAmpGain.argtypes = [c_int]
        self._atmcd.SetPreAmpGain.restype = c_int
        self._atmcd.SetShutter.argtypes = [c_int, c_int, c_int, c_int]
        self._atmcd.SetShutter.restype = c_int
        self._atmcd.GetDetector.argtypes = [ctypes.POINTER(c_int), ctypes.POINTER(c_int)]
        self._atmcd.GetDetector.restype = c_int
        self._atmcd.GetPixelSize.argtypes = [ctypes.POINTER(c_float), ctypes.POINTER(c_float)]
        self._atmcd.GetPixelSize.restype = c_int
        self._atmcd.SetAcquisitionMode.argtypes = [c_int]
        self._atmcd.SetAcquisitionMode.restype = c_int
        self._atmcd.SetReadMode.argtypes = [c_int]
        self._atmcd.SetReadMode.restype = c_int
        self._atmcd.SetTriggerMode.argtypes = [c_int]
        self._atmcd.SetTriggerMode.restype = c_int
        self._atmcd.SetExposureTime.argtypes = [c_float]
        self._atmcd.SetExposureTime.restype = c_int
        self._atmcd.AbortAcquisition.argtypes = []
        self._atmcd.AbortAcquisition.restype = c_int
        self._atmcd.GetAcquisitionTimings.argtypes = [
            ctypes.POINTER(c_float),
            ctypes.POINTER(c_float),
            ctypes.POINTER(c_float),
        ]
        self._atmcd.GetAcquisitionTimings.restype = c_int
        self._atmcd.StartAcquisition.argtypes = []
        self._atmcd.StartAcquisition.restype = c_int
        self._atmcd.WaitForAcquisitionTimeOut.argtypes = [c_int]
        self._atmcd.WaitForAcquisitionTimeOut.restype = c_int
        self._atmcd.GetAcquiredData.argtypes = [ctypes.POINTER(c_long), c_int]
        self._atmcd.GetAcquiredData.restype = c_int
        self._atmcd.GetStatus.argtypes = [ctypes.POINTER(c_int)]
        self._atmcd.GetStatus.restype = c_int

        self._shamrock.ShamrockInitialize.argtypes = [c_char_p]
        self._shamrock.ShamrockInitialize.restype = c_int
        self._shamrock.ShamrockClose.argtypes = []
        self._shamrock.ShamrockClose.restype = c_int
        self._shamrock.ShamrockGetNumberDevices.argtypes = [ctypes.POINTER(c_int)]
        self._shamrock.ShamrockGetNumberDevices.restype = c_int
        self._shamrock.ShamrockSetGrating.argtypes = [c_int, c_int]
        self._shamrock.ShamrockSetGrating.restype = c_int
        self._shamrock.ShamrockSetWavelength.argtypes = [c_int, c_float]
        self._shamrock.ShamrockSetWavelength.restype = c_int
        self._shamrock.ShamrockGetGrating.argtypes = [c_int, ctypes.POINTER(c_int)]
        self._shamrock.ShamrockGetGrating.restype = c_int
        self._shamrock.ShamrockGetWavelength.argtypes = [c_int, ctypes.POINTER(c_float)]
        self._shamrock.ShamrockGetWavelength.restype = c_int
        self._shamrock.ShamrockAutoSlitIsPresent.argtypes = [c_int, c_int, ctypes.POINTER(c_int)]
        self._shamrock.ShamrockAutoSlitIsPresent.restype = c_int
        self._shamrock.ShamrockSetAutoSlitWidth.argtypes = [c_int, c_int, c_float]
        self._shamrock.ShamrockSetAutoSlitWidth.restype = c_int
        self._shamrock.ShamrockShutterIsPresent.argtypes = [c_int, ctypes.POINTER(c_int)]
        self._shamrock.ShamrockShutterIsPresent.restype = c_int
        self._shamrock.ShamrockSetShutter.argtypes = [c_int, c_int]
        self._shamrock.ShamrockSetShutter.restype = c_int
        self._shamrock.ShamrockSetNumberPixels.argtypes = [c_int, c_int]
        self._shamrock.ShamrockSetNumberPixels.restype = c_int
        self._shamrock.ShamrockSetPixelWidth.argtypes = [c_int, c_float]
        self._shamrock.ShamrockSetPixelWidth.restype = c_int
        self._shamrock.ShamrockGetCalibration.argtypes = [c_int, ctypes.POINTER(c_float), c_int]
        self._shamrock.ShamrockGetCalibration.restype = c_int
        self._shamrock.ShamrockGetPixelCalibrationCoefficients.argtypes = [
            c_int,
            ctypes.POINTER(c_float),
            ctypes.POINTER(c_float),
            ctypes.POINTER(c_float),
            ctypes.POINTER(c_float),
        ]
        self._shamrock.ShamrockGetPixelCalibrationCoefficients.restype = c_int

    def _initialize_camera(self) -> None:
        init_code = int(self._atmcd.Initialize(str(self.sdk_root).encode("ascii")))
        if init_code not in (DRV_SUCCESS, DRV_VXDNOTINSTALLED):
            error_name = CAMERA_ERROR_NAMES.get(init_code, "UNKNOWN_CAMERA_ERROR")
            raise AndorSDKError(f"Initialize failed with code {init_code} ({error_name})")
        self._camera_initialized = True

        camera_count = c_int()
        count_code = int(self._atmcd.GetAvailableCameras(byref(camera_count)))
        if count_code != DRV_SUCCESS:
            error_name = CAMERA_ERROR_NAMES.get(count_code, "UNKNOWN_CAMERA_ERROR")
            raise AndorSDKError(f"GetAvailableCameras failed with code {count_code} ({error_name})")
        self._available_cameras = camera_count.value
        if self._available_cameras <= 0:
            raise AndorHardwareNotFoundError(
                "未检测到 Andor 相机。当前环境未连接相机或驱动未就绪，所以本次不会执行光谱采集。"
            )

    def _initialize_shamrock(self) -> None:
        init_code = int(self._shamrock.ShamrockInitialize(str(self.sdk_root).encode("ascii")))
        if init_code != SHAMROCK_SUCCESS:
            error_name = SHAMROCK_ERROR_NAMES.get(init_code, "UNKNOWN_SHAMROCK_ERROR")
            raise AndorSDKError(f"ShamrockInitialize failed with code {init_code} ({error_name})")
        self._shamrock_initialized = True
        time.sleep(0.3)

        device_count = c_int()
        count_code = int(self._shamrock.ShamrockGetNumberDevices(byref(device_count)))
        if count_code != SHAMROCK_SUCCESS:
            error_name = SHAMROCK_ERROR_NAMES.get(count_code, "UNKNOWN_SHAMROCK_ERROR")
            raise AndorSDKError(f"ShamrockGetNumberDevices failed with code {count_code} ({error_name})")
        self._available_shamrocks = device_count.value
        if self._available_shamrocks <= 0:
            raise ShamrockHardwareNotFoundError(
                "未检测到 Shamrock 光谱仪。当前环境未连接光谱仪或其 USB 驱动未就绪，所以本次不会执行光谱采集。"
            )

    def open(self) -> None:
        if not self.sdk_root.exists():
            raise FileNotFoundError(f"Andor SDK root not found: {self.sdk_root}")

        self._load_libraries()
        with self._sdk_working_directory():
            self._initialize_camera()
            self._initialize_shamrock()

    def close(self) -> None:
        if self._shamrock_initialized:
            try:
                self._shamrock.ShamrockClose()
            finally:
                self._shamrock_initialized = False
        if self._camera_initialized:
            try:
                self._atmcd.ShutDown()
            finally:
                self._camera_initialized = False
        for handle in reversed(self._dll_dirs):
            handle.close()
        self._dll_dirs.clear()

    def get_detector_info(self) -> tuple[int, int, float, float]:
        xpixels = c_int()
        ypixels = c_int()
        px_w = c_float()
        px_h = c_float()
        self._camera_call("GetDetector", byref(xpixels), byref(ypixels))
        self._camera_call("GetPixelSize", byref(px_w), byref(px_h))
        return xpixels.value, ypixels.value, px_w.value, px_h.value

    def get_camera_count(self) -> int:
        return int(self._available_cameras)

    def get_shamrock_count(self) -> int:
        return int(self._available_shamrocks)

    def get_current_grating(self) -> int:
        grating = c_int()
        self._shamrock_call("ShamrockGetGrating", 0, byref(grating))
        return int(grating.value)

    def get_current_wavelength_nm(self) -> float:
        wavelength = c_float()
        self._shamrock_call("ShamrockGetWavelength", 0, byref(wavelength))
        return float(wavelength.value)

    def get_pixel_calibration_coefficients(self) -> tuple[float, float, float, float]:
        a = c_float()
        b = c_float()
        c = c_float()
        d = c_float()
        self._shamrock_call(
            "ShamrockGetPixelCalibrationCoefficients",
            0,
            byref(a),
            byref(b),
            byref(c),
            byref(d),
        )
        return float(a.value), float(b.value), float(c.value), float(d.value)

    def _camera_call_allow_temp_status(self, name: str, *args) -> int:
        if self._atmcd is None:
            raise AndorSDKError("Camera SDK not loaded")
        func = getattr(self._atmcd, name)
        code = int(func(*args))
        if code in (
            DRV_SUCCESS,
            DRV_TEMP_OFF,
            DRV_TEMP_NOT_STABILIZED,
            DRV_TEMP_STABILIZED,
            DRV_TEMP_NOT_REACHED,
            DRV_TEMP_OUT_RANGE,
            DRV_TEMP_NOT_SUPPORTED,
            DRV_TEMP_DRIFT,
        ):
            return code
        error_name = CAMERA_ERROR_NAMES.get(code, "UNKNOWN_CAMERA_ERROR")
        raise AndorSDKError(f"{name} failed with code {code} ({error_name})")

    def get_temperature_c(self) -> tuple[float, int]:
        temperature = c_float()
        code = self._camera_call_allow_temp_status("GetTemperatureF", byref(temperature))
        return float(temperature.value), code

    def enable_cooler_and_set_target(self, target_temperature_c: int) -> None:
        self._camera_call("CoolerON")
        self._camera_call("SetTemperature", int(target_temperature_c))

    def configure_camera_readout(
        self,
        horizontal_readout_mhz: float | None,
        pre_amp_gain_value: float | None,
        output_amplifier: int = OUTPUT_AMPLIFIER_CONVENTIONAL,
        ad_channel: int = 0,
    ) -> dict[str, float | int]:
        self._camera_call("SetADChannel", int(ad_channel))

        requested_output_amplifier = int(output_amplifier)
        selected_output_amplifier = requested_output_amplifier
        amplifier_candidates: list[int] = []
        for candidate in (requested_output_amplifier, OUTPUT_AMPLIFIER_CONVENTIONAL, OUTPUT_AMPLIFIER_EMCCD):
            if candidate not in amplifier_candidates:
                amplifier_candidates.append(candidate)

        amplifier_error: AndorSDKError | None = None
        for candidate in amplifier_candidates:
            try:
                self._camera_call("SetOutputAmplifier", int(candidate))
                selected_output_amplifier = int(candidate)
                amplifier_error = None
                break
            except AndorSDKError as exc:
                amplifier_error = exc
                continue
        if amplifier_error is not None:
            raise amplifier_error

        speed_count = c_int()
        self._camera_call("GetNumberHSSpeeds", int(ad_channel), int(selected_output_amplifier), byref(speed_count))
        best_speed_index = 0
        best_speed_value = 0.0
        if speed_count.value > 0:
            candidates: list[tuple[int, float]] = []
            for idx in range(speed_count.value):
                speed = c_float()
                self._camera_call("GetHSSpeed", int(ad_channel), int(selected_output_amplifier), idx, byref(speed))
                candidates.append((idx, float(speed.value)))
            if horizontal_readout_mhz is None:
                best_speed_index, best_speed_value = candidates[0]
            else:
                best_speed_index, best_speed_value = min(
                    candidates,
                    key=lambda item: abs(item[1] - float(horizontal_readout_mhz)),
                )
            self._camera_call("SetHSSpeed", int(selected_output_amplifier), best_speed_index)

        gain_count = c_int()
        self._camera_call("GetNumberPreAmpGains", byref(gain_count))
        best_gain_index = 0
        best_gain_value = 0.0
        if gain_count.value > 0:
            candidates = []
            for idx in range(gain_count.value):
                gain = c_float()
                self._camera_call("GetPreAmpGain", idx, byref(gain))
                candidates.append((idx, float(gain.value)))
            if pre_amp_gain_value is None:
                best_gain_index, best_gain_value = candidates[0]
            else:
                best_gain_index, best_gain_value = min(
                    candidates,
                    key=lambda item: abs(item[1] - float(pre_amp_gain_value)),
                )
            self._camera_call("SetPreAmpGain", best_gain_index)

        return {
            "ad_channel": int(ad_channel),
            "output_amplifier": int(selected_output_amplifier),
            "horizontal_speed_index": best_speed_index,
            "horizontal_speed_mhz": best_speed_value,
            "preamp_gain_index": best_gain_index,
            "preamp_gain_value": best_gain_value,
        }

    def configure_spectrograph(
        self,
        grating_no: int,
        center_wavelength_nm: float,
        slit_width_um: float | None = None,
        shamrock_shutter_mode: int | None = None,
    ) -> None:
        self._shamrock_call("ShamrockSetGrating", 0, grating_no)
        self._shamrock_call("ShamrockSetWavelength", 0, c_float(center_wavelength_nm))
        if slit_width_um is not None:
            present = c_int()
            self._shamrock_call("ShamrockAutoSlitIsPresent", 0, 1, byref(present))
            if present.value:
                self._shamrock_call("ShamrockSetAutoSlitWidth", 0, 1, c_float(slit_width_um))
        if shamrock_shutter_mode is not None:
            present = c_int()
            self._shamrock_call("ShamrockShutterIsPresent", 0, byref(present))
            if present.value:
                self._shamrock_call("ShamrockSetShutter", 0, int(shamrock_shutter_mode))

    def configure_acquisition(
        self,
        exposure_s: float,
        trigger_mode: int,
        camera_shutter_mode: int | None = None,
        camera_shutter_open_ms: int = 0,
        camera_shutter_close_ms: int = 0,
    ) -> float:
        self._camera_call("SetAcquisitionMode", ACQ_MODE_SINGLE)
        self._camera_call("SetReadMode", READ_MODE_FVB)
        self._camera_call("SetTriggerMode", trigger_mode)
        if camera_shutter_mode is not None:
            self._camera_call(
                "SetShutter",
                1,
                int(camera_shutter_mode),
                int(camera_shutter_close_ms),
                int(camera_shutter_open_ms),
            )
        self._camera_call("SetExposureTime", c_float(exposure_s))
        actual_exp = c_float()
        accum = c_float()
        kinetic = c_float()
        self._camera_call("GetAcquisitionTimings", byref(actual_exp), byref(accum), byref(kinetic))
        return actual_exp.value

    def acquire_spectrum(
        self,
        exposure_s: float,
        trigger_mode: int,
        timeout_ms: int | None = None,
        progress_callback=None,
        camera_shutter_mode: int | None = None,
        camera_shutter_open_ms: int = 0,
        camera_shutter_close_ms: int = 0,
    ) -> tuple[list[float], list[int], float]:
        xpixels, _, pixel_width_um, _ = self.get_detector_info()
        self._shamrock_call("ShamrockSetNumberPixels", 0, xpixels)
        self._shamrock_call("ShamrockSetPixelWidth", 0, c_float(pixel_width_um))
        actual_exposure = self.configure_acquisition(
            exposure_s,
            trigger_mode,
            camera_shutter_mode=camera_shutter_mode,
            camera_shutter_open_ms=camera_shutter_open_ms,
            camera_shutter_close_ms=camera_shutter_close_ms,
        )
        timeout = timeout_ms if timeout_ms is not None else max(5000, int(actual_exposure * 1000) + 10000)
        self._camera_call("StartAcquisition")
        if progress_callback is not None:
            progress_callback(0.0)
            target = max(actual_exposure, 0.1)
            start_time = time.time()
            while True:
                elapsed = time.time() - start_time
                if elapsed >= target:
                    break
                progress_callback(min(0.98, elapsed / target))
                time.sleep(min(0.2, max(0.05, target / 100.0)))
            progress_callback(0.99)
        wait_code = int(self._atmcd.WaitForAcquisitionTimeOut(timeout))
        if wait_code != DRV_SUCCESS:
            error_name = CAMERA_ERROR_NAMES.get(wait_code, "UNKNOWN_CAMERA_ERROR")
            if trigger_mode == 1:
                raise AndorSDKError(
                    f"等待采集超时：当前为外触发模式，但在 {timeout / 1000:.1f} s 内未收到外部触发信号 "
                    f"(code {wait_code}, {error_name})"
                )
            raise AndorSDKError(
                f"采集等待超时或失败 (code {wait_code}, {error_name})"
            )
        if progress_callback is not None:
            progress_callback(1.0)
        data_buffer = (c_long * xpixels)()
        self._camera_call("GetAcquiredData", data_buffer, xpixels)
        calibration_buffer = (c_float * xpixels)()
        self._shamrock_call("ShamrockGetCalibration", 0, calibration_buffer, xpixels)
        x_axis = [float(calibration_buffer[i]) for i in range(xpixels)]
        y_axis = [int(data_buffer[i]) for i in range(xpixels)]
        return x_axis, y_axis, actual_exposure

    def abort_acquisition(self) -> None:
        if self._atmcd is None or not self._camera_initialized:
            return
        try:
            self._atmcd.AbortAcquisition()
        except Exception:
            return

    def save_ascii(self, output_path: Path, x_axis: list[float], y_axis: list[int]) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with output_path.open("w", encoding="utf-8") as fh:
            for x, y in zip(x_axis, y_axis):
                x_text = f"{x:.6f}"
                fh.write(f"{x_text}\t{y}\n")
