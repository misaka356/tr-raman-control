# TODO.md

## P0 必须先做

### 1. 提交项目记忆文档

- 相关文件：`AGENTS.md`、`PROJECT_OVERVIEW.md`、`ARCHITECTURE.md`、`CURRENT_STATUS.md`、`TODO.md`。
- 任务：确认文档内容无误后提交 git。
- 验收标准：`git status --short` 只显示预期文档，提交信息清楚。
- 依赖关系：无。

### 2. 做一次主流程静态验证

- 相关文件：`tr_raman_ui.py`、`tr_raman_integrated_controller.py`、`andor_sdk_integration.py`。
- 任务：运行 `py_compile`，确认当前主版本无语法错误。
- 验收标准：三个文件编译通过。
- 依赖关系：无。

### 3. 离线模拟完整实验

- 相关文件：`tr_raman_ui.py`、`tr_raman_integrated_controller.py`、`app_config.json`。
- 任务：启用离线模拟，跑 `0-360`、步长 `20`、每点 `3` 次。
- 验收标准：生成预期文件名；日志显示 CH1 实际相位为 `0, -20, -40 ... -360`；暂停续跑逻辑可验证。
- 依赖关系：P0-2。

## P1 重要但不阻塞

### 4. 维护依赖清单

- 相关文件：`requirements.txt`。
- 任务：已记录 `numpy`、`pyvisa`、`pyvisa-py`、`pyserial`、`pywinauto`、`lmfit`、`psutil`、`zeroconf` 等依赖；后续依赖变化时同步维护。
- 验收标准：新环境可一条命令安装依赖。
- 依赖关系：无。

### 5. 系统测试 RIGOL 命令序列

- 相关文件：`tr_raman_integrated_controller.py`。
- 任务：在 mock transport 下记录正式实验、调试、基线测试的 SCPI 命令。
- 验收标准：命令序列符合 CH1 主触发、CH2 外触发、`*OPC?`、0.5s 等待、相位取负。
- 依赖关系：P0-2。

### 6. 真实硬件同步验证

- 相关文件：代码和实验记录。
- 任务：用示波器验证 CH1/CH2 输出起始时刻和延时校准流程。
- 验收标准：记录仪器开机后校准步骤、推荐延时参数、异常处理方式。
- 依赖关系：需要实验室硬件。

### 7. Andor 连接诊断优化

- 相关文件：`andor_sdk_integration.py`、`tr_raman_ui.py`。
- 任务：把 SDK 路径错误、DLL 缺失、相机未连接、Shamrock 未连接、相机被占用做成清晰中文提示。
- 验收标准：用户不需要读 Python traceback 就能判断问题层级。
- 依赖关系：需要至少一次无硬件和一次有硬件测试。

### 8. 清理历史 monkey patch

- 相关文件：`tr_raman_ui.py`。
- 任务：将文件末尾的 `_replacement_*` 逻辑整合回类定义，降低后续维护风险。
- 验收标准：行为不变，语法检查通过，UI 按钮仍能调用正确逻辑。
- 依赖关系：建议先做 git 备份。

## P2 后续优化

### 9. 清理体积

- 相关文件：`Launcher/`、`TRRamanUI/`、`vendor/`、`__pycache__/`。
- 任务：区分运行必需、打包产物、缓存和历史遗留。
- 验收标准：删除内容前有清单；删除后启动软件不受影响。
- 依赖关系：先备份或提交 git。

### 10. 增加测试目录

- 相关文件：新增 `tests/`。
- 任务：测试相位列表、相位取负、文件命名、Raman 换算、lmfit 拟合、SCPI 命令序列。
- 验收标准：无硬件环境可运行。
- 依赖关系：P1-5。

### 11. 文档持续维护

- 相关文件：全部 `.md`。
- 任务：每次改硬件流程、配置字段、算法或打包方式后更新文档。
- 验收标准：新线程打开本目录能快速恢复上下文。
- 依赖关系：持续执行。
## 2026-04-20 completed

- Realtime stress preview quality gate added.
- Ghost warning level changed from a single boolean UI message to low/high risk display.
- Initial real-data validation completed with `20260417-5hz-3%-30s-4`: output written to `analysis_output\realtime_preview_validation_20260420`.
- Extreme real-data validation completed with `20260409-10-0.3-3`: output written to `analysis_output\realtime_preview_validation_20260420_extreme`; `140-1.asc` is now caught as `narrow_peak_top_spike`.
- Spectrometer continuous preview button added. It still needs a hardware smoke test to confirm repeated SDK acquisitions stop cleanly after the current exposure.
- Experiment run stress trend preview added in the scan-parameter area. It still needs a formal offline/hardware run smoke test to confirm callbacks update the raw and mean points as expected.
- Remaining hardware task: validate warning thresholds with additional real spectra from the lab, especially borderline shoulder peaks and broadened peak tops.
