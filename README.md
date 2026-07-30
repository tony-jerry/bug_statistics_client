# 缺陷统计客户端

这是一个 Windows 桌面客户端，用于查询绩效系统中指定计划版本的缺陷，并按“缺陷引入人”进行统计。

## 使用方法

1. 优先双击 `dist\BugStatisticsClient_v17.exe`；也可以双击 `run_client.bat` 从源码启动。
2. 输入用户名、绩效系统密码和 CRM 密码，点击“登录”。
   两个密码框后的眼睛按钮可分别显示或隐藏密码。
3. 登录区下方通过浅色分段导航切换“绩效缺陷明细”和“CRM 缺陷查询”；界面采用现代浅色卡片主题。
4. “绩效查询与筛选”和统计卡片只显示在“绩效缺陷明细”页签中；登录后会加载当前年份的计划版本，并默认选中当前月份，选择后点击“查询接口”。
5. 登录成功后，“缺陷引入人”会自动填写当前登录人的姓名；点击“应用筛选”可查看匹配明细。
6. 双击缺陷明细后，CRM 表单中的全部业务字段（包括空值）会在同一个滚动区域连续显示。
7. 点击详情中的图片可打开原始尺寸查看器，并可放大、缩小、恢复 100% 或适应窗口。
8. “CRM 缺陷查询”页签支持按 Bug 编号或标题关键字搜索，每页固定显示 50 条，并显示总数和分页按钮。
9. CRM 查询结果和绩效缺陷明细中的 Bug 编号均显示为无下划线的蓝色链接，单击编号即可打开完整缺陷详情。
10. CRM 查询结果和绩效缺陷明细的 Bug 编号后都有复制图标，单击即可复制编号，并短暂显示“复制成功”提示。
11. 缺陷详情底部的“在 CRM 中打开”可在默认浏览器中打开 CRM 原始详情页。
12. 点击“导出 Excel”可导出当前筛选结果和完整的引入人汇总。
13. 点击“导出回溯 MD”会读取当前筛选缺陷的 CRM 详情，按“缺陷概况、主要问题根因、典型Bug深度分析、后续改进”生成归纳型报告，并按“月份版本bug回溯-引入人.md”命名。

登录、查询、分页和详情读取等后台操作执行期间会显示半透明全局 Loading 层，在保留页面内容可见的同时阻止重复点击。

## 安全说明

- 登录成功后，两套系统的密码会使用 Windows DPAPI 按当前 Windows 用户加密保存，并在下次启动时自动填入。
- `%APPDATA%\BugStatisticsClient\settings.json` 仅保存服务地址、上次使用的用户名、计划版本和筛选条件；加密密码单独保存在 `credentials.dat` 中，不保存明文密码或登录令牌。
- 关闭客户端后，会话随程序一并销毁。

## 运行环境

- Windows
- Python 3.10 或以上
- `requests`
- `openpyxl`
- `cryptography`
- `Pillow`

如缺少依赖，可在本目录执行：

```powershell
pip install -r requirements.txt
```

## 本地编译 EXE

### 1. 准备环境

安装 Python 3.10 或以上版本。打开 PowerShell，进入项目目录：

```powershell
Set-Location "C:\Users\user\Desktop\工作文件夹\脚本\performance_scripts\bug_statistics_client"
```

安装运行依赖和 PyInstaller：

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
```

### 2. 编译前测试

建议先检查语法并运行离线测试：

```powershell
python -m py_compile app.py analytics.py api_client.py credential_store.py crm_client.py detail_utils.py retrospective.py
python test_client.py
```

看到“核心逻辑测试通过”后再进行打包。

### 3. 使用 spec 文件编译

先关闭正在运行的同版本客户端，然后执行：

```powershell
python -m PyInstaller --noconfirm --clean --workpath build_v16 BugStatisticsClient_v16.spec
```

编译完成后的文件位于：

```text
dist\BugStatisticsClient_v16.exe
```

请优先使用项目中的 `BugStatisticsClient_v16.spec`，不要直接用简单的
`pyinstaller app.py` 命令。spec 文件已通过 `pathex` 和 `hiddenimports`
显式包含以下本地模块，可避免启动时报 `No module named 'analytics'`：

- `analytics`
- `api_client`
- `credential_store`
- `crm_client`
- `detail_utils`
- `retrospective`

### 4. 编译新版本

例如需要生成 v17：

1. 复制 `BugStatisticsClient_v16.spec`，命名为 `BugStatisticsClient_v17.spec`。
2. 将 spec 文件中的 `name='BugStatisticsClient_v16'` 改为
   `name='BugStatisticsClient_v17'`。
3. 执行：

```powershell
python -m PyInstaller --noconfirm --clean --workpath build_v17 BugStatisticsClient_v17.spec
```

新版本会生成在 `dist\BugStatisticsClient_v17.exe`。

注意：`--workpath build_v17` 只修改临时编译目录，不会修改 EXE 名称。
EXE 名称由 spec 文件中的 `name` 决定。因此下面这条命令仍会尝试生成并覆盖
`BugStatisticsClient_v16.exe`，不能用于生成 v17：

```powershell
# 错误示例：使用的仍是 v16 spec
python -m PyInstaller --noconfirm --clean --workpath build_v17 BugStatisticsClient_v16.spec
```

如果 `BugStatisticsClient_v16.exe` 正在运行，错误示例还会触发
`PermissionError: [WinError 5] 拒绝访问`。

### 5. 常见问题

- 提示 `PermissionError: [WinError 5] 拒绝访问`：输出目录中的同名 EXE
  正在运行。关闭对应版本，或者改用名称匹配的新版本 spec 后重新编译。
- 提示找不到模块：确认使用的是项目内的 spec 文件，并从项目目录执行命令。
- 编译后立即退出：先运行 `python test_client.py`，再检查
  `build_vXX\BugStatisticsClient_vXX\warn-BugStatisticsClient_vXX.txt`。
- Windows 安全软件拦截：将项目目录加入可信范围后重新编译，并确认 EXE
  来源是本机生成的文件。
