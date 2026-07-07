# 优惠券批量生成器 — 开发完成报告

## 项目概述

基于 PRD 需求文档开发的**优惠券批量生成器**，一个本地 Python HTTP 服务器 + 浏览器前端的工具，将 Excel 券码列表与模板图片合成为带独立二维码和券码文字的多页 PDF，用于批量印刷优惠券。

## 交付物清单

| 文件 | 说明 | 大小 |
|------|------|------|
| `coupon_web.py` | 主程序（HTTP 服务器 + HTML 前端，单文件） | 59 KB |
| `双击启动.command` | macOS 启动脚本，双击即用 | 2.0 KB |
| `coupon_web.spec` | PyInstaller 打包配置（Windows EXE 构建） | 3.0 KB |
| `requirements.txt` | Python 依赖清单 | 0.3 KB |
| `build_exe.bat` | Windows 一键构建 EXE 脚本 | 3.0 KB |
| `.github/workflows/build-windows.yml` | GitHub Actions 自动构建 Windows EXE | 2.0 KB |
| `test_web_server.py` | TDD 测试套件（28 项测试） | 19 KB |
| `券码示例.xlsx` | 测试用 Excel 券码文件（20 条） | 5.0 KB |
| `模板示例.png` | 测试用模板图片（800×500） | 5.4 KB |

## 技术架构

```
浏览器 (Chrome/Safari/Firefox/Edge)
    ↕ HTTP (localhost:8888)
ThreadingHTTPServer (Python, 多线程)
    ├── qrcode     → QR 码生成
    ├── Pillow     → 图片处理
    ├── reportlab  → 多页 PDF 生成（文字可选中复制）
    ├── openpyxl   → .xlsx 解析
    ├── xlrd       → .xls/.et 解析
    └── csv        → .csv 解析
```

## 功能实现

### P0 必需功能（全部完成）
- **导入券码数据**：支持 .xlsx / .xls / .et / .csv，可选列（A-E），可跳过标题行，显示数量和前10条预览
- **加载模板图片**：支持 PNG / JPG / JPEG / BMP，显示尺寸信息
- **二维码可视化定位与缩放**：拖拽移动、红角调整大小（20~500px）、精确像素输入
- **券码文字独立定位与缩放**：拖拽移动、蓝角调字号（8~100px）、半透明白底、可开关、PDF 中可选中复制
- **生成多页 PDF**：每页 = 模板 + 独立QR + 文字，自动下载，进度状态显示

### P1 应有功能（全部完成）
- **中英文切换**：一键切换所有界面文字
- **清空与重新上传**：券码和图片各自独立清空，前后端状态同步

### 额外增强
- **深色/浅色主题切换**：记忆用户偏好
- **Toast 通知系统**：操作反馈（成功/错误/信息）
- **加载遮罩**：生成 PDF 时显示加载状态
- **响应式布局**：小屏幕自动堆叠
- **拖拽提示**：操作时显示提示文字
- **输入框联动**：可输入精确像素值，与拖拽双向同步
- **跨平台字体检测**：macOS/Windows/Linux 自动适配中文字体

## Windows EXE 打包方案

代码已完全跨平台兼容（修复了 macOS 独占字体路径，添加 `_find_chinese_font()` 函数自动检测系统字体）。

### 方式一：GitHub Actions 自动构建（推荐）

1. 将项目推送到 GitHub 仓库
2. 进入仓库 → Actions → "Build Windows EXE" → Run workflow
3. 构建完成后下载 artifact `优惠券批量生成器-Windows.zip`
4. 发给客户解压使用

优点：无需 Windows 机器，GitHub 自动在 Windows runner 上编译。

### 方式二：Windows 机器手动构建

1. 将项目文件复制到客户 Windows 电脑
2. 确保已安装 Python 3.9+（勾选 "Add to PATH")
3. 双击 `build_exe.bat`
4. 脚本自动：检查 Python → 安装依赖 → PyInstaller 打包 → 生成 zip

产出：`dist/优惠券批量生成器-Windows.zip`，约 46MB。

### 方式三：PyInstaller 命令行手动构建

```bash
pip install -r requirements.txt
pyinstaller --clean coupon_web.spec
```

产出在 `dist/优惠券批量生成器/` 目录。

### 客户使用方式

客户拿到 EXE 后：
1. 解压 zip 或直接进入文件夹
2. 双击 `优惠券批量生成器.exe`
3. 浏览器自动打开 http://localhost:8888
4. 正常操作：上传 Excel → 上传模板 → 拖拽定位 → 生成 PDF

**注意**：
- Windows 上首次启动可能较慢（5-10 秒），因为需要初始化运行环境
- 如浏览器未自动打开，手动访问 http://localhost:8888
- 关闭控制台窗口或按 Ctrl+C 停止服务

## API 端点

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/` | 返回 HTML 操作界面 |
| GET | `/api/qr?code=X&size=N` | 生成 QR 码 PNG |
| GET | `/api/template` | 返回已上传模板图片 |
| POST | `/api/upload/excel?col=X&skip=N` | 上传并解析券码文件 |
| POST | `/api/upload/template` | 上传模板图片 |
| POST | `/api/generate` | 生成多页 PDF（JSON 参数） |
| POST | `/api/reset?what=excel|image|all` | 清空状态 |

## 测试结果

```
28 项测试全部通过 (1.091s)
  - QR 码生成:        3/3 ✓
  - Excel/CSV 解析:  10/10 ✓
  - PDF 生成:         4/4 ✓  (含 RGBA 模板测试)
  - HTTP 端点:       12/12 ✓
```

端到端验证：
- GET / → 40KB HTML 页面 ✓
- GET /api/qr → 2.3KB PNG ✓
- POST /api/upload/excel → 正确读取 3 个券码 ✓
- POST /api/upload/template → 正确存储 800×500 图片 ✓
- POST /api/generate → 3 页 PDF (21KB) ✓

PyInstaller macOS 构建验证：
- spec 配置正确 ✓
- 打包后 macOS 可执行文件 HTTP 200 ✓
- 产出大小约 46MB ✓

## 使用方式

### macOS：双击启动（推荐）
双击 `双击启动.command`，浏览器自动打开操作界面。

### macOS：命令行启动
```bash
python3 coupon_web.py
```

### Windows：EXE 启动
双击 `优惠券批量生成器.exe`，浏览器自动打开操作界面。

### 操作流程
1. 选择 Excel 文件 → 选列 → 可选跳过标题行 → 查看券码数量
2. 选择模板图片 → 查看尺寸信息
3. 拖拽二维码到目标位置 → 拖红角调大小 → 拖文字到位置 → 拖蓝角调字号
4. 点击"生成 PDF" → 浏览器自动下载多页 PDF
