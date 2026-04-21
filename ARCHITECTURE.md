# ARCHITECTURE.md

## 总体结构

```text
用户
  -> 时间分辨拉曼控制软件.exe / start_tr_raman.bat
  -> tr_raman_ui.py
      -> app_config.json
      -> tr_raman_integrated_controller.py
          -> RIGOL VISA / SCPI
          -> andor_sdk_integration.py
              -> atmcd64d.dll
              -> ShamrockCIF.dll
          -> 输出 .asc
          -> 实时预览 / 当前应力
```

## 目录说明

```text
.
├─ tr_raman_ui.py                       # 主 UI
├─ tr_raman_integrated_controller.py    # 实验编排和算法
├─ andor_sdk_integration.py             # Andor SDK 封装
├─ app_config.json                      # 当前配置
├─ app_config.template.json             # 配置模板
├─ start_tr_raman.bat                   # 启动脚本
├─ 时间分辨拉曼控制软件.exe              # 轻量启动器
├─ vendor/andor_sdk                     # Andor SDK 运行文件
├─ TRRamanUI/                           # 打包版 UI 目录
├─ Launcher/                            # 启动器相关文件
└─ *.md                                 # 项目记忆、运行和打包文档
```

## 核心模块

### `tr_raman_ui.py`

负责：

- 构建 UI。
- 读取/保存配置。
- 扫描 VISA。
- 启动后台线程执行实验、基线测试、调试和连接测试。
- 显示温度、状态、曝光进度、日志、预览和当前应力。
- 选择零点文件并调用后端拟合当前应力。

### `tr_raman_integrated_controller.py`

负责：

- 定义 `ChannelConfig`、`SpectrometerConfig`、`IntegratedExperimentConfig`。
- RIGOL VISA transport 和 SCPI 封装。
- 构建相位列表。
- 将 UI 相位转换为 CH1 实际输出相位：`resolve_ch1_output_phase_deg()`。
- 输出文件命名。
- 离线模拟光谱。
- 波长轴到 Raman shift 的公式换算。
- `lmfit` 峰拟合和应力计算。
- 鬼峰检测和文件名标注。
- 暂停续跑状态保存。
- 正式实验、基线测试、Andor 调试、发生器调试。

### `andor_sdk_integration.py`

负责：

- 加载项目内 Andor SDK DLL。
- 初始化相机和 Shamrock。
- 设置光栅、中心波长、曝光、触发、读出、快门、制冷。
- 获取温度和采集数据。
- 获取光谱标定轴。

## RIGOL 同步链路

硬件连接：

```text
RIGOL CH1/AUX trigger output -> BNC -> RIGOL CH2 external trigger input
```

软件流程：

```text
配置 CH1(AOM) 波形
配置 CH2(音圈) 波形
CH1 Burst = triggered, source = BUS/manual command
CH1 trigger output = ON, positive slope
CH2 Burst = triggered, source = external rising edge
设置 CH1 Burst phase = -UI_phase
*OPC?
打开 CH1/CH2
等待 0.5 s
:TRIGger1:IMMediate
等待 settle_time_ms
当前相位连续采集 n 次
关闭 CH1/CH2
```

为降低换相位时仪器漏写概率，代码中已有 `set_burst_phase_verified()`，会写入 CH1 Burst 相位并回读校验。

## 数据流

```text
UI 参数
  -> IntegratedExperimentConfig
  -> RIGOL 配置当前相位
  -> Andor SDK 单次采集
  -> Shamrock 标定波长轴
  -> Raman shift 公式换算
  -> 保存 ASC
  -> UI 预览
  -> lmfit 拟合峰位
  -> 当前应力显示
```

## 文件命名

- 正式实验：`输出目录\样品名\相位-重复.asc`，例如 `20-1.asc`。
- 基线测试：`输出目录\样品名\-1.asc`、`-2.asc`。
- 暂停续跑状态：保存在样品输出目录中，具体文件名见 `pause_state_path()`。

## 应力算法

实时应力只显示当前值，不生成应力表。

拟合流程：

```text
515-535 cm^-1 窗口
LorentzianModel(prefix="peak_")
+ QuadraticModel(prefix="bg_")
peak_model.guess()
model.fit()
center_cm1 = result.params["peak_center"].value
stress_mpa = 435.0 * (baseline_peak_cm1 - current_peak_cm1)
```

该算法必须与 `raman-data-organizer` 保持一致。
## 2026-04-20 realtime stress quality gate

- The live stress preview still uses one `lmfit.LorentzianModel + QuadraticModel` peak fit in the 515-535 cm^-1 window.
- Each fit now carries live quality metrics: R2, SNR, center uncertainty, raw-peak/fit-center delta, and FWHM.
- Stress is hidden when the fit is unreliable instead of showing a numeric MPa value with weak evidence.
- Ghost-peak analysis now reports `risk_level` as `none`, `low`, or `high`, with warning types for separated double peaks, shoulder peaks, and broadened or flattened peak tops.
- Low-risk ghost warnings do not automatically suppress stress. High-risk ghost warnings suppress stress only when the fit quality is not clean; high-risk with a good fit is displayed as a reference value with a warning.
- The live preview does not use the previous frame peak center to seed or constrain the next frame.

## 2026-04-20 continuous spectrometer preview

- The spectrometer tab has a single toggle button for continuous acquisition.
- Continuous preview uses `run_continuous_preview_acquisition()` and repeatedly overwrites `_continuous_preview.asc` under the sample output folder.
- Each refreshed preview file flows through the same UI path as a saved spectrum, so the plot, peak readout, ghost warning, and realtime stress preview stay live.
- Continuous preview is a monitoring path only. It does not create phase/repeat filenames, baseline filenames, resume state, stress trend tables, or phase-stress plots.

## 2026-04-20 experiment stress trend preview

- The experiment run tab includes a compact `tk.Canvas` stress trend preview in the blank area to the right of the experiment scan parameters.
- Formal experiment spectrum callbacks parse only `phase-repeat.asc` files and ignore baseline, debug, and `_continuous_preview.asc` files.
- Raw valid stress points are plotted as small gray dots. Per-phase averages are plotted as the green trend line.
- A stress point enters the trend only when realtime fit quality passes and ghost risk is not high; high-risk reference values are not used for the formal trend preview.
