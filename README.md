# 时间分辨拉曼控制软件

用于时间分辨拉曼实验的 Windows 桌面控制程序。软件统一配置 RIGOL 双通道信号发生器、Andor 光谱仪和 AOM/音圈电机相位扫描流程，支持离线模拟、基线测试、正式相位扫描、暂停/继续、实时光谱预览和实时应力质量判断。

## 主要功能

- Tkinter/ttk 图形界面。
- RIGOL VISA/SCPI 双通道控制。
- Andor SDK DLL 直接采集光谱。
- CH1(AOM) 主触发、CH2(音圈) 外部上升沿触发的同步流程。
- 相位扫描时将 UI 相位转换为 CH1 Burst 负相位输出。
- 按样品名创建输出目录，保存 `phase-repeat.asc` 光谱文件。
- 基线测试输出 `-1.asc`、`-2.asc` 等参考光谱。
- 实时应力预览使用 `lmfit.models.LorentzianModel + QuadraticModel`，并带拟合质量和鬼峰风险判断。
- 离线模拟模式可在无硬件环境下检查流程和文件命名。

## 关键文件

- `tr_raman_ui.py`: 主界面、配置读写、后台线程、预览和实时状态。
- `tr_raman_integrated_controller.py`: 实验流程、RIGOL 控制、文件命名、Raman shift 转换、应力拟合。
- `andor_sdk_integration.py`: Andor 相机和 Shamrock 光谱仪 SDK 封装。
- `app_config.template.json`: 配置模板。实际运行时复制为 `app_config.json`。
- `start_tr_raman.bat`: Windows 启动脚本。
- `AGENTS.md`: 本仓库协作规则和主版本路径约定。
- `CURRENT_STATUS.md`、`ARCHITECTURE.md`、`TODO.md`: 当前状态、架构说明和后续任务。

## 安装依赖

建议在虚拟环境中运行：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -r requirements.txt
```

Andor SDK DLL、RIGOL VISA 后端和真实硬件驱动需要按实验室电脑环境单独安装或放置。厂商 SDK、打包产物和本地配置不纳入源码仓库。

## 运行

```powershell
python .\tr_raman_ui.py
```

或使用：

```powershell
.\start_tr_raman.bat
```

首次运行前可从模板复制配置：

```powershell
Copy-Item .\app_config.template.json .\app_config.json
```


## 验证

无硬件环境下至少先运行语法检查：

```powershell
python -m py_compile .\tr_raman_ui.py .\tr_raman_integrated_controller.py .\andor_sdk_integration.py
```

真实硬件流程仍需要在实验室用示波器和实际采集验证，尤其是 CH1 触发 CH2 的时序、Andor SDK 初始化和连续预览停止逻辑。
