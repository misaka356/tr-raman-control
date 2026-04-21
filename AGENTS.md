# AGENTS.md

本文件是本目录的协作与执行指令。后续维护以当前目录为准：

```text
C:\Users\adimn\Desktop\code\andor\release\TRRaman_UI_Backup_20260408\TRRaman_UI_Backup_20260410
```

最高优先级上下文来源：

```text
C:\Users\adimn\.codex\sessions\2026\04\07\rollout-2026-04-07T18-56-58-019d6796-a214-7182-873e-38d7e6ffc440.jsonl
```

不要再编辑其他旧备份目录，除非用户明确指定。

## 项目目标

本项目是时间分辨拉曼实验控制软件，用 Windows 图形界面统一控制 RIGOL 信号发生器和 Andor 光谱仪，实现相位扫描、多次曝光采集、基线测试、暂停续跑、离线模拟、实时预览和当前应力显示。

## 技术栈

- Python 3.13。
- UI：Tkinter / ttk。
- 信号发生器：PyVISA / pyvisa-py，主线只使用 VISA。
- 光谱仪：`andor_sdk_integration.py` 通过 Andor 官方 SDK DLL 控制相机和 Shamrock。
- 打包/启动：源码文件夹 + `start_tr_raman.bat` + 轻量启动器。
- 实时应力拟合：`lmfit.models.LorentzianModel + QuadraticModel`，必须与 `raman-data-organizer` 保持一致。

## 核心约定

- CH1 = AOM 声光调制器，通常为方波。
- CH2 = 音圈电机激振，通常为正弦。
- 双通道同步使用 CH1 作为主触发：
  - CH1 Burst 源为 BUS/命令触发。
  - CH1 触发输出为上升沿。
  - CH2 Burst 源为外部上升沿。
  - 背部 AUX BNC 线把 CH1 触发输出接到 CH2 外触发。
  - 打开 CH1/CH2 后等待 `0.5 s`，再发送 `:TRIGger1:IMMediate`。
- 相位扫描修改 CH1(AOM) 的 Burst 相位。
- UI 输入正相位，实际下发给 CH1 时自动取负号：`0, 20, 40` -> `0, -20, -40`。
- 当前相位点内连续采集 `n` 次时，不关闭信号发生器；采完该相位后再关闭并切换下一相位。
- Raman 横坐标使用公式换算：`1e7 / rayleigh_nm - 1e7 / wavelength_nm`。

## 修改原则

- 修改前先读：`CURRENT_STATUS.md`、`ARCHITECTURE.md`、`TODO.md`。
- 涉及实验流程时，同时检查 `tr_raman_ui.py`、`tr_raman_integrated_controller.py`、`app_config.template.json`。
- 涉及 Andor 时，同时检查 `andor_sdk_integration.py`。
- 涉及拟合/应力时，不要写自定义拟合算法；必须调用 `lmfit` 模型。
- 涉及配置字段时，保持 `app_config.json` 向后兼容。
- 涉及启动/打包时，同步更新 `start_tr_raman.bat`、`PACKAGING_README.md` 或 `UI_README.md`。
- 修改后至少运行 `python -m py_compile tr_raman_ui.py tr_raman_integrated_controller.py andor_sdk_integration.py`。

## 不要做

- 不要把其他备份目录的代码改动误同步到本目录。
- 不要退回顺序打开 CH1/CH2 的同步方式。
- 不要把双通道共同内部 BUS 触发当成当前主线。
- 不要重新引入 `scipy curve_fit` 手写 Lorentzian 或抛物线兜底。
- 不要隐藏硬件错误；要区分未连接、DLL 缺失、VISA 缺失、SDK 初始化失败。
- 不要硬编码密钥、仪器序列号或只能在某台电脑使用的路径。
