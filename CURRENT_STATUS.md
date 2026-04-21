# CURRENT_STATUS.md

## 当前状态

这是确认后的主维护目录，且已经初始化 git：

```text
C:\Users\adimn\Desktop\code\andor\release\TRRaman_UI_Backup_20260408\TRRaman_UI_Backup_20260410
```

## 已确认实现

- `andor_sdk_integration.py` 存在。
- UI 可以引用后端的基线、暂停、连接测试、应力拟合等接口。
- RIGOL 主路径为 VISA。
- CH1 = AOM，CH2 = 音圈。
- CH1 相位下发时自动取负号。
- RIGOL 使用 CH1 触发输出带动 CH2 外部触发。
- 配置后使用 `*OPC?`。
- 打开 CH1/CH2 后到 `:TRIGger1:IMMediate` 前有稳定等待。
- 换相位可只写 CH1 Burst 相位并回读校验。
- 输出文件进入样品子目录。
- 离线模拟模式存在。
- 基线测试存在。
- 暂停/继续/停止逻辑存在。
- 自动温度监视和制冷逻辑存在。
- Raman shift 使用公式换算。
- 实时应力拟合使用 `lmfit.LorentzianModel + QuadraticModel`。
- 启动脚本检查并安装 `lmfit` 等依赖。

## 需要注意

- 真实 RIGOL 同步仍必须用示波器验证，尤其是每次开机后的 CH1/CH2 延时。
- Andor SDK 初始化错误要区分 DLL 路径、硬件未连接、相机占用和驱动问题。
- `tr_raman_ui.py` 文件末尾存在 monkey patch 风格的替换函数，例如 `_replacement_start_temperature_poll`、`_replacement_run_background`。这是历史修补痕迹，后续重构时要小心，不要误删有效逻辑。
- `Launcher/`、`TRRamanUI/`、`vendor/` 体积可能较大；清理前确认是否影响实验室运行。
- 已增加 `requirements.txt` 作为 GitHub 源码仓库的基础 Python 依赖清单。
- 还没有系统化测试目录。

## 最近关键决策

- 不再维护 `C:\Users\adimn\Desktop\编程练习\andor` 作为主线。
- 不再使用额外横坐标校正参数，回到 Raman shift 公式换算。
- 实时应力算法禁止手写拟合和抛物线兜底。
- 换相位时优先只改 CH1 Burst 相位，避免重新清空/重配造成偶发漏写。

## 下次先做

1. 运行语法检查：

```powershell
python -m py_compile .\tr_raman_ui.py .\tr_raman_integrated_controller.py .\andor_sdk_integration.py
```

2. 运行离线模拟，检查 `0-360`、步长 `20`、每点 `3` 次的文件名和日志。
3. 有硬件时，用示波器验证 CH1 触发 CH2 的实际延时。
4. 后续依赖变化时同步维护 `requirements.txt`。
5. 将这些项目记忆文档纳入 git 提交。
## 2026-04-20 update

- Realtime stress preview now keeps the single `lmfit.LorentzianModel + QuadraticModel` fit, but adds live quality gating before stress is displayed.
- Unreliable live fits are shown as hidden/failed quality states instead of numeric MPa values.
- Ghost warnings are now graded as low or high risk and distinguish separated double peaks, shoulder peaks, and broadened or flattened peak tops.
- UI preview now shows ghost risk level and fit quality text separately from the stress value.
- Validated against `C:\Users\adimn\Desktop\实验数据\20260417-5hz-3%-30s-4\20260417-5hz-3%-30s-4`: 39 spectra, baseline `-1.asc` center `523.684972 cm^-1`, 35 normal shown, 2 low-risk cautions, 1 high-risk reference display, and 1 unreliable high-risk spectrum hidden from realtime stress.
- Revalidated against the noisier `C:\Users\adimn\Desktop\实验数据\20260409-10-0.3-3\20260409-10-0.3-3`: added `narrow_peak_top_spike` so `140-1.asc` is now high-risk and hidden; all 21 `_ghost` files are flagged, and the 20260417 cleaner dataset keeps the same 35/2/2 risk split.
- The spectrometer tab now has a single-button continuous preview mode. It repeatedly overwrites `_continuous_preview.asc`, refreshes the preview/peak/ghost/stress UI, and is kept separate from formal experiment filenames and trend outputs.
- The experiment run tab now has a compact realtime stress trend canvas fixed to the right side of the experiment scan parameter area. Only formal experiment `phase-repeat.asc` saves update this trend; baseline, debug, and continuous preview paths do not.
