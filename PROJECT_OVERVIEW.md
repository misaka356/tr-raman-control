# PROJECT_OVERVIEW.md

## 项目简介

这是时间分辨拉曼实验控制软件的当前维护版本。它用于在实验室电脑上配置 RIGOL 双通道信号发生器和 Andor 光谱仪，完成 AOM 相位扫描、每相位多次光谱采集、基线测试、暂停续跑和实时应力显示。

## 当前主要功能

- 图形界面：`tr_raman_ui.py`。
- 配置读写：`app_config.json` / `app_config.template.json`。
- RIGOL VISA/SCPI 控制：`tr_raman_integrated_controller.py`。
- Andor SDK 控制：`andor_sdk_integration.py`。
- 离线模拟模式，不连接硬件也能跑完整流程。
- 自动制冷与右上角温度状态显示。
- 光谱仪连接测试、光谱仪调试采集。
- 发生器双通道调试和单通道开关。
- 正式相位扫描实验。
- 暂停、继续、停止实验。
- 基线测试，只启动 AOM 并输出 `-1.asc`、`-2.asc` 等文件。
- 输出目录按样品名建立子目录。
- 采集后预览光谱和当前应力。
- 实时应力拟合使用 `lmfit.LorentzianModel + QuadraticModel`。
- 启动脚本会检查并安装 `pyvisa`、`pyvisa-py`、`psutil`、`zeroconf`、`lmfit`。

## 实验约定

- CH1 = AOM，主触发通道。
- CH2 = 音圈，外部触发从通道。
- 相位扫描的是 CH1(AOM) 的 Burst 相位。
- UI 输入 `0 -> 360`，实际下发给 CH1 为 `0 -> -360`。
- 当前相位的 `n` 次采集期间不关闭发生器。
- 光谱仪需要冷却到 `-60 C` 并稳定后再进行正式实验。
- 第一列 Raman shift 当前由波长轴公式换算得到。

## 关键文件

- `tr_raman_ui.py`：Tkinter UI、按钮回调、日志、进度、温度、预览和当前应力显示。
- `tr_raman_integrated_controller.py`：实验流程、RIGOL 控制、文件命名、暂停续跑、基线测试、Raman 换算、拟合和应力计算。
- `andor_sdk_integration.py`：Andor 相机和 Shamrock 光谱仪 SDK 封装。
- `app_config.template.json`：配置模板。
- `start_tr_raman.bat`：用户启动入口。
- `tr_raman_icon.ico/png`：软件图标。
- `vendor/andor_sdk`：项目内 Andor SDK 运行文件。

## 运行方式

源码运行：

```powershell
python .\tr_raman_ui.py
```

常用启动：

```powershell
.\start_tr_raman.bat
```

也可以双击当前目录中的：

```text
时间分辨拉曼控制软件.exe
```

## 依赖

从代码和启动脚本确认：

- Python 3.13
- numpy
- pyvisa
- pyvisa-py
- pyserial
- pywinauto
- lmfit
- psutil
- zeroconf
- Andor SDK DLL
- RIGOL VISA 设备或可用的 VISA backend

已增加 `requirements.txt`，可用 `python -m pip install -r requirements.txt` 安装 Python 依赖。
