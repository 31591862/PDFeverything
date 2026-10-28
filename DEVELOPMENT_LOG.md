================================================================
  PDFEVERYTHING — 开发日志
================================================================

Ver 0.1.0 | 2026-06-26
----------------------------------------------------------------

  Ver 0.1.0 | CLI 原型
    - 创建 Python CLI 工具 pdf_tool.py，12 个 PDF 操作
    - merge / split / extract-text / extract-images / to-images
      / from-images / compress / watermark / encrypt / decrypt / rotate / info
    - 依赖 PyMuPDF + pypdf + pikepdf + Pillow + pdfplumber
    - argparse 子命令分发

  Ver 0.2.0 | Core 层重构
    - core/pdf_ops.py：PdfOperator 静态方法类，progress_callback 支持
    - core/utils.py：文件分类、编码检测、临时文件管理
    - core/converters.py：ConverterRegistry 注册表模式
      —— ImageConverter / TextConverter / WordConverter /
         PowerPointConverter / ExcelConverter / PdfPassThroughConverter
    - core/merger.py：混合文件合并流水线（逐文件转换→合并）
    - macOS Word/PPT/Excel 通过 AppleScript 调用 Office 原生转换
    - python-docx + PyMuPDF 纯 Python 渲染回退方案
    - pdf_tool.py 重构为调用 core/pdf_ops.py

  Ver 0.3.0 | PyQt6 GUI 框架
    - gui/main_window.py：双 Tab 主窗口（合并与转换 / PDF 工具）
    - gui/file_list_widget.py：拖拽文件列表 + 工具栏 + 右键菜单
    - gui/workers.py：QThread BaseWorker，后台多线程不阻塞 UI
    - gui/dialogs.py：加密/解密/水印/旋转/压缩/拆分/信息对话框
    - main.py：CLI 模式 / GUI 模式自动切换
    - macOS .app 封装（PyInstaller + .icns 图标）
    - Windows .exe 封装（PyInstaller onefile + NSIS 安装程序→onefile 单文件）

----------------------------------------------------------------

Ver 1.0.0 | 2026-06-26 — 首个正式发布
----------------------------------------------------------------

  Ver 1.0.0 | 首次 Release
    - GitHub Release v1.0.0，macOS .app + Windows .exe 双平台发布
    - 应用图标：PDF 文字 + 花体 SignPainter "everything"
    - README 双语完整文档（英文 + 中文）
    - 混合文件 → 统一 PDF 作为核心功能突出展示
    - macOS Dock 图标 + Windows 任务栏图标

  Ver 1.0.1 | Word 转换修复
    - BUGFIX: AppleScript Word 转换失败（"变量 theDoc 没有定义"）
      — open 命令不返回引用，改用 open→active document
    - BUGFIX: _applescript_convert 弱校验——增加 PDF 页数验证
    - BUGFIX: python-docx 回退渲染器忽略嵌入式图片——增加
      docx inline_shape 提取 + page.insert_image 嵌入

----------------------------------------------------------------

Ver 1.1.0 | 2026-06-26 — i18n + Windows Office COM + CLI 模式
----------------------------------------------------------------

  Ver 1.1.0 | 国际化 + Windows Office + AI CLI
    - 新增 gui/i18n.py 翻译系统：zh / en 170+ 键位
    - Settings → Language 菜单切换中英文，所有 UI 即时刷新
    - 所有 GUI 文件重构为 tr() 调用（main_window / dialogs /
      file_list_widget）
    - Windows Office COM 集成：win32com 调用 Word/PPT/Excel 原生转换
      (pywin32)，macOS 保持 AppleScript
    - 跨平台 Office 检测：core/utils.py 平台感知
    - CLI 模式增强：-h / --help / --version / --mcp 标志
    - AI Agent 无头调用支持：PDFeverything.exe merge -i a.pdf -o out.pdf
    - main.py HELP_TEXT 完整命令参考

  Ver 1.1.1 | 图标优化 + README 双语
    - 应用图标重绘：everything 字体从 SignPainter 72px→96px，间距 +50px
    - README 新版本文档独立 badge（v1.1.0 蓝色版）
    - 中文 README 完整同步：GUI 预览、CLI 用法、技术栈
    - README 标题图标从 .ico 改为 .png（GitHub 不渲染 .ico）

  Ver 1.1.2 | MCP 服务器
    - 新增 mcp/server.py：Model Context Protocol JSON-RPC 服务器
    - 13 个工具 + 完整 JSON Schema（参数类型/必填/描述）
    - main.py --mcp 标志从编译后 exe/app 启动 MCP
    - Claude Desktop / Claude Code / Cursor 配置文档
    - BUGFIX: NSIS 中文路径报错——复制到临时英文目录打包

----------------------------------------------------------------

Ver 1.2.0 | 2026-06-26 — PDF 反向转换 + GUI 简化 + 批量处理
----------------------------------------------------------------

  Ver 1.2.0 | PDF→Word / PDF→PPT / PDF→Excel
    - core/pdf_ops.py 新增 to_word / to_ppt / to_excel
      - to_word：PyMuPDF 提取文字+表格→python-docx 重建 .docx
      - to_ppt：每页渲染为全幅图片→python-pptx 幻灯片
      - to_excel：Page.find_tables() 提取→openpyxl 每个表一个工作表
    - pdf_tool.py 新增 to-word / to-ppt / to-excel 子命令
    - main.py CLI_COMMANDS + HELP_TEXT 新增 3 个命令
    - mcp/server.py 新增 pdf_to_word / pdf_to_ppt / pdf_to_excel 工具
      （16 tools total）
    - GUI 工具 Tab 新增 3 个按钮 + 处理函数
    - i18n 新增 3 个翻译键
    - README EN+CN 功能表 + MCP 工具列表 + 计数更新（13→16）
    - 版本号 1.1.0→1.2.0（main.py / spec / README / mcp/server.py）

  Ver 1.2.1 | 批量处理 + UI 简化 + 鲁棒性加固
    - UI 简化：合并面板删除 Images→PDF 和 Word→PDF 按钮，
      只保留一个"合并为统一 PDF"按钮，自动识别文件类型
    - 批量处理：所有操作（压缩/水印/加密/解密/旋转/提取/
      转Word/转PPT/转Excel）均支持多文件批量
      — 1 个文件→文件对话框，多个→输出目录自动命名
    - gui/main_window.py 新增 _run_batch() 通用批处理辅助方法
    - gui/main_window.py 新增 _pick_input_files() 多文件拾取器
      （优先从文件列表，回退多选文件对话框）
    - 批处理进度条显示文件名和进度（Compress (3/12): report.pdf）
    - 批处理结果汇总弹窗列出所有输出文件
    - 文件列表保护：最大 200 文件，单文件上限 500MB，跳过 0 字节文件
    - 批量确认：文件数 >20 弹确认框，上限 200 超过截断
    - BaseWorker 加固：60min 硬超时，MemoryError 捕获，进度值钳位 0-100
    - 取消按钮：5 秒宽限期→force terminate
    - BUGFIX: Dock 运行图标恢复——main.py 中 app.setWindowIcon()
      从 bundle 读 app_icon.png 正确设置

  Ver 1.2.2 | Skill 本地化
    - 创建 .claude/RELEASE_CHECKLIST.md：6 步预发布检查清单
    - 覆盖新功能接线（7 个文件）、工具计数一致性、版本字符串、
      图标格式、CLI/MCP/GUI 测试、重编译与发布
    - 创建 CLAUDE.md + DEVELOPMENT_LOG.md 开发文档

----------------------------------------------------------------

Ver 1.3.0 | 2026-06-26 — PDF 阅读器
----------------------------------------------------------------

  Ver 1.3.0 | PDF 阅读器（初版）
    - 新增 gui/pdf_reader_widget.py：清爽单页 PDF 阅读器
      —— fitz Pixmap → QImage → QPixmap 渲染管线
    - 顶部工具栏 + 单页翻页模式
    - 缩放：适应宽度 / 适应页面 / 100-300% / Ctrl+滚轮
    - MainWindow 第 3 个 Tab "📖 Reader"，File → Open PDF (Ctrl+P)
    - 双击文件列表 PDF → 自动切换到阅读器

  Ver 1.3.1 | 阅读器重设计 — 底栏 + 3 种阅读模式 + 触控板
    - 工具栏移至窗口底部，最大化阅读区域
    - 3 种阅读模式：
      · Scroll（默认）：连续垂直滚动，所有页面堆叠，触控板平滑滚动
      · Single：逐页翻看，◀▶ 按钮 / ←→ 键盘 / 页码跳转
      · Grid：2 列缩略图网格，点击缩略图切换到 Single 模式
    - 触控板双指捏合缩放（macOS+Windows 原生手势，无需 Ctrl）
    - QScrollArea 原生触控板惯性滚动
    - 底栏新增 Scroll / Single / Grid 三按钮切换模式
    - i18n 新增 reader_scroll / reader_single / reader_grid 翻译键

  Ver 1.3.2 | 阅读器性能重写 — 标签池 + 全量预渲染 + 缩放弹窗
    - 标签池复用：所有页面 QLabel 只创建一次（_build_labels），
      模式切换只移动/隐藏，不销毁重建（6ms 切换）
    - 空间换时间：打开 PDF 时预渲染所有页面到缓存（MAX_CACHE=2000），
      后续滚动/翻页/缩放变更瞬间完成
    - 关闭 PDF 时清空所有缓存+标签，内存释放
    - 触控板双指捏合缩放无需 Ctrl（phase 检测 + 角度/像素比值法），
      缩放时弹出居中半透明百分比弹窗（1.2 秒自动消失）
    - Fit Width / Fit Page 改为 checkable 按钮，已激活时不再触发重渲染
    - 按钮悬停 3 秒后光标旁弹出功能解释（tooltip）
    - 阅读器工具栏按钮中英文本地化
    - 滚动时自动追踪当前页码（scroll→page sync）
    - 测试通过：Hello Algo 379p/18MB，打开 4.5s，模式切换 6ms，
      滚动流畅 60fps
    - BUGFIX: setFixedSize 容器无显示 — label.setParent + absolute move 解决
    - BUGFIX: QWheelEvent 缩放不支持 macOS 原生手势 — phase + angleDelta 双检测
    - BUGFIX: Grid 模式切换耗时 2.7s — 标签池复用消除重建开销

  Ver 1.3.3 | 阅读器 v3 — 仅 Scroll+Grid，适应高度，3 列网格，关闭按钮
    - 移除单页模式，仅保留 Scroll 和 Grid 两种模式
    - 新增 Fit Height：将单页高度缩放至窗口高度
    - Fit Width/Fit Height 改为独立 checkable 按钮，互斥切换
    - Grid 改为 3 列排列，双击缩略图→Fit Height+Scroll 跳转该页
    - 连续滚动模式翻页按钮和页码常驻显示
    - 滚动时自动追踪当前页（>50% 可见区域即为当前页）
    - 翻页按钮/←→键盘：滚动到目标页，页面顶部对齐窗口顶部
    - 右下角文件名旁新增 ✕ 关闭按钮
      · 从 "File→Open" 进入→关闭后留在阅读器欢迎页
      · 从文件列表双击→关闭后自动跳转回合并 Tab
    - 预渲染改为分批异步（每批 40 页），初次加载保持 UI 响应
    - BUGFIX: Fit Width 和 Fit Page 完全相同 — 拆分为 Width/Height 两个维度
    - BUGFIX: 连续模式底栏显示翻页按钮和页码 — 改为仅 Grid 隐藏

  Ver 1.3.4 | 欢迎页 + 完整 i18n + 右键"打开方式"集成
    - 阅读器空载时显示居中欢迎页："拖入 PDF 文件以阅读" + "加载文件"按钮
    - 支持拖拽 PDF 文件到阅读器窗口直接打开
    - _retranslate_ui() 在 _init_ui() 末尾调用，所有按钮/文字启动即本地化
    - 阅读器工具栏全中文化（连续/网格/适应宽度/适应高度）
    - 所有 tooltip 中英双语
    - macOS Info.plist 注册 .pdf 文件类型（LSHandlerRank=Alternate），
      系统"打开方式"菜单中出现 PDFeverything，文件自动加入文件列表
    - main.py 新增 _collect_file_args()，忽略 -psn_ 等 macOS 噪声参数
    - 版本号升级至 1.3.0

  Ver 1.3.5 | 阅读器 v5 — 跟手缩放 + WPS 网格 + 点击编辑缩放 + 高速页码追踪
    - 缩放改为双重架构：_apply_smooth_zoom() 即时缩放现有 pixmap（跟手），
      _real_zoom_timer 250ms 防抖后执行真实渲染（清晰画质）
    - 触控板双指捏合：累积 angleDelta，每 120 单位 = 3% 缩放，50-300% 范围
    - [+] [-] 按钮 ±1% 步长，smooth=True 模式即时响应
    - 缩放比例下拉菜单（QComboBox）完全删除，替换为可点击编辑的 QLineEdit
      · QIntValidator(25-500) 限制合法输入，>500→500，<50→50，小数四舍五入
      · Enter 或失焦生效
    - 网格模式重设计（参照 WPS 网格视图）：
      · 3 列居中均匀分布，gutter 20px，缩略图按可用宽度动态计算
      · 左侧边距自动调整使网格居中
    - 页码追踪重写：_page_heights[] 数组记录每页 Y 位置，二分查找 O(log n)
      · _scroll_timer 间隔 16ms（~60fps），滚动停止立即更新页码
    - 欢迎页修复：QTimer.singleShot(100ms) 延迟显示，确保 viewport 已布局
    - BUGFIX: 欢迎页不显示 — _destroy_welcome() 清理 + center.raise_() 置顶
    - BUGFIX: 页码不更新 — 二分查找替换逐页遍历
    - BUGFIX: 网格模式双击无效 — _on_grid_dbl_click 检查 view_mode==GRID
    - BUGFIX: 缩放过慢 — smooth 即时缩放 + real render 防抖

  Ver 1.3.6 | 阅读器 v6 — 多分辨率缓存 + bisect 页码追踪 + 欢迎页 overlay
    - 缩放性能重写：移除所有 PdfReaderWidget._cache.clear() 调用
      · 多分辨率缓存：不同 zoom_key 的页面共存于 dict 中
      · zoom 切换时只渲染缓存未命中的页面，已缓存直接复用
      · 缓存上限 3000 entries LRU
    - 页码追踪：bisect_right（Python C 级实现）在 _page_heights 排序数组上二分查找
      · 100ms timer，O(log n)，379 页查找 <1μs
    - 欢迎页重写：作为 self 的子 widget 覆盖在 scroll_area 上方
      · showEvent 触发显示，不受 viewport 尺寸影响
      · _destroy_welcome 确保关闭 PDF 时清理
    - 关闭 PDF：page_container.setFixedSize(0,0) 彻底移除滚动条
    - 缩放步长：+/- 按钮 5%，触控板 pinch 每 120 angle units = 5%
    - fit_width/fit_height 不再清空缓存

  Ver 1.3.7 | 二阶段缩放完整实现 + 整数缓存键 + 捏合终止检测
    - **完整二阶段缩放流水线**：
      · Pass 1（即时）：`_smooth_scale_all()` 对现有 QPixmap 做 SmoothTransformation 缩放
        — <1ms 视觉反馈，用户立即看到页面大小变化
      · Pass 2（延迟）：180ms `QTimer.singleShot` → `_sharp_render()` PyMuPDF 真实渲染
        — 高清画质自动替换模糊图
    - **快速缓存探测**：Pass 2 调度前检查是否所有页面已缓存，若是则完全跳过重渲染
    - **整数百分比舍入**：`int(round(pct))` 确保 101%和 101.4% 共享同一缓存键
    - **连续捏合优化**：`skip_deferred=True` 在手势期间仅做即时缩放，不触发后台渲染
    - **捏合终止检测**：`ScrollEnd (phase=3)` → 触发 `_sharp_render` 产生最终高清画面
    - BUGFIX: `_show_welcome` 重复居中代码 + `vw`/`vh` 未定义变量名

  Ver 1.3.8 | LRU 缓存 + 内存水位线 + 双 timer 页码追踪 + 按需预加载
    - 缓存替换为 collections.OrderedDict（LRU 有序淘汰）：
      · `_cache_put`: 插入时 `move_to_end`，超 250MB 限制时 `popitem(last=False)`
      · `_cache_get`: 命中时 `move_to_end`（最近使用标记）
      · 驻留特权：`zk in ("fh","fw")` 永不被淘汰（fit_width/fit_height for page 0）
      · `_cache_memory_bytes` 实时追踪，`_clear_cache()` 清零
      · 删除旧 `MAX_CACHE=3000` 条目上限，改为内存驱动
    - 双 timer 页码追踪：
      · Throttle (30ms): `_do_throttle_page` — `scrollY / avgPageHeight` 粗略页码，<1μs
      · Debounce (150ms): `_do_debounce_calibration` — `bisect_right` 精确校准
      · 滚动停止后自动触发 `_prefetch_around(N)` 预加载 N±1 页
    - 快速缓存探测：`_set_zoom_pct` 仅检查 page 0 是否命中，不再遍历所有页面
    - `_sharp_render` 改用 `_cache_get`（LRU 语义），`finally` 清除 `_pending_zoom_pct`
    - `_cancel_deferred_renders` 统一停止所有 timer
    - BUGFIX: 滚动→闪退 — `_on_scrollbar_changed` + throttle/debounce 双重守卫

  Ver 1.3.9 | 缩放比例计算修正
    - **删除 `_smooth_scale_all`** — 该方法使用 `_fw_ratio` 作为缩放基数，
      但 pixmap 可能处于任意缩放级别（fit_height/fit_width/任意%）。
      从 fit_height(83%) 缩放至 150% 时，计算出 scale=1.5/2.0=0.75 — 缩小而非放大
    - **三处缩放路径重写为正确的相对比例**：
      · `_set_zoom_pct`: 在修改 `self._zoom_mode` 前保存 `old_pct`，
        计算 `scale = new_factor / old_factor`
      · `_apply_fit_mode`: 同样先存旧值再算比值
      · `_on_resize`: 使用旧 `_fw_ratio/_fh_ratio` 计算旧因子
    - 所有缩放路径统一：old_pct → old_factor, new_factor, scale = new/old
    - BUGFIX: 点 +/- 按钮页面反向缩放或不变 — 缩放基数错误导致方向反转

  Ver 1.3.10 | 阅读器 UI 完善 + 缩放性能优化
    - 空文档状态规范化：文件名和 ✕ 按钮隐藏；页码 0/0；翻页按钮禁用
      1 页文档显示 1/1 且翻页按钮禁用；关闭文档后恢复空状态
    - 文件名标签样式更新：圆角药丸 + 内阴影模拟凹陷感
    - 翻页/缩放按钮 30px→34px 放大，◀ ▶ − + 全量显示
    - 页面居中：竖版/横版混合 PDF 各自在视口宽度居中对齐
    - 缩放性能重写：
      · Pass 1 仅缩放可见 labels（<5ms），不再遍历全部页面
      · `_layout_labels` 空 pixmap 时用 PDF 尺寸估算，不阻塞 UI
      · Pass 2 延迟 180ms→40ms，高清渲染更快到位
      · `_on_resize` 简化为比率更新 + 延迟渲染
    - Tooltip：按钮正上方弹出（mapToGlobal），白字深黑紧凑样式
    - BUGFIX: 滚动到边界后捏合误触发缩放 — 真实 pinch 检测 pixelDelta==0
    - BUGFIX: 双 timer 页码跳动 — throttle 仅在 >1 页跳变时更新
    - BUGFIX: 工具栏布局报错 — setEnabled 移到 btn_next 创建之后

  Ver 1.3.11 | 懒渲染策略：只渲染可见范围 ±2 页

    - **删除全量预渲染**：移除 _pre_render_all / _pre_render_batch / BATCH 常量
      —— 打开 PDF 不再渲染所有页面，仅渲染当前页 ± 2 页（约 5 页）
    - **按需渲染流水线**：
      · `_render_visible_range()` — 渲染当前页 ± 2 页（约 5 页）
      · `_visible_page_range()` — 计算 (start, end) 范围，边界钳位
      · `_sharp_render` — 缩放后只渲染可见范围，不再遍历全部页面
      · `_do_debounce_calibration` — 滚动停止后自动调用 _render_visible_range
    - **即时缩放不变**：Pass 1 仍然是可见 label 的即时像素拉伸 (<5ms)
    - **缓存保留**：已渲染过的缩放比例不再重复渲染（多分辨率 cache 仍然有效）
    - **内存优化**：删除 _prefetch_around / _safe_prefetch + 预渲染 timer
    - **效果**：点按 +/- 瞬时响应 → 40ms 后仅附近 5 页变清晰
      → 用户滚动时，停下的位置立即渲染可见页
      → 从未见过的页面在滚动停止时才首次渲染


  Ver 1.3.12 | 图标重绘 — 液态玻璃艺术风格 + Dock 修复 + UI 微调
    - **图标重绘** (1024×1024 px)：
      · 硬对角蒙版 + TR/BL 圆角处高斯渐变过渡（r=12px）
      · 左上白色高光 outline (rgba 120) + 右下暗色阴影 outline (rgba 70)
      · PDF 240px Helvetica Bold，everything 192px SignPainter
      · 外框 radius=160 + 内面板 radius=120，双层结构
    - **备份**：平面版本 (app_icon_flat_1024) 保存在 resource 子目录
    - **Dock 图标修复**：移除 app.setWindowIcon()，依赖 CFBundleIconFile 原生加载 .icns
      — 彻底消除启动时的白边闪现
    - **Skill 文档**：.claude/skills/ICON_DESIGN.md 记录完整制图流程和代码
    - BUGFIX: 翻页/缩放按钮箭头被裁切 — 30px→34px
    - BUGFIX: 空文档状态工具栏显示文件名和 ✕ — 改为隐藏
    - BUGFIX: 空文档点击 Fit W/H/+/- 崩溃 — 增加 if not self.doc guard
    - BUGFIX: 滚动到边界捏合误触发 — pixelDelta==0 检测
    - BUGFIX: 缓慢滚动页码跳动 — throttle 仅在 >1 页跳变时更新
    - BUGFIX: 混合竖版/横版 PDF 左对齐 — 每页独立居中对齐
    - BUGFIX: Tooltip 暗色不可见 — 全局白字深黑紧凑样式
    - BUGFIX: 缩放基数错误导致反向缩放 — new/old ratio 替换 _fw_ratio

    
  Ver 1.3.13 | Immortal 100% base + 默认 100% 缩放 + 滚动保持 + 3 页懒渲染

    - **默认 100% 缩放**：打开 PDF 不再自动 Fit Height，始终从 100% 开始
      · `_zoom_mode = 1.0`（非 "fit_height"），Fit W/Fit H 按钮初始未选中
      · `zoom_edit` 显示 "100" 而非自适应高度百分比
    - **Immortal 100% 缓存基础**：
      · `_pre_render_100_all()` — 打开 PDF 时全量预渲染所有页面到 100% 缩放
      · 缓存键 `(pi, "z:1.000")` 永不过期——`_cache_put` 中跳过 immortal key 淘汰
      · 所有缩放操作的 Pass 1 (<5ms) 直接从 100% base QPixmap 做缩放，消除累积误差
      · `target_factor = pct / 100.0` 线性计算目标尺寸，不再依赖上次渲染比例
    - **缩放/Fit 模式切换后保持滚动位置**：
      · `_set_zoom_pct` + `_apply_fit_mode` 末尾统一调 `_scroll_to_page_top()`
      · 防止页面尺寸变化后 scrollbar 绝对值对应到不同页面
    - **懒渲染 ±2→±1**：`_visible_page_range` 从 5 页（当前±2）缩小到 3 页（当前±1）
      · 只渲染当前页+前后各 1 页，Pass 2 内存和渲染压力更小
    - BUGFIX: Fit Width→Fit Height→Fit Width 切换时滚动位置跳变
      — `_apply_fit_mode` 末尾加入 `_scroll_to_page_top()`
    - BUGFIX: 缩放后 `_layout_labels` 覆盖 Pass 1 缩放结果
      — 仅在 `render_missing=True`（初次加载）时调用渲染，缩放路径传 False
    - BUGFIX: 空 reader 滚动触发崩溃 — 增加 `if not self.doc` guard


  Ver 1.3.14 | WPS 级画质提升 — 4× SSAA + HiDPI 原生分辨率 + 懒预渲染

    - **RENDER_SCALE 2→4**：超采样抗锯齿从 2× 提升至 4×（标准屏）/ 8×（Retina）
      · 渲染分辨率翻倍 → 缩放到逻辑尺寸 → 文字锐利度显著提升
    - **HiDPI 原生支持**：
      · `_render_page` 新增 `dpr` 参数，`fitz.Matrix` 乘以 `devicePixelRatio()`
      · QPixmap 设置 `setDevicePixelRatio(dpr)`，Qt 自动做 1:1 像素映射
      · Retina 屏：595×842 逻辑 px → 1190×1684 物理像素（4 倍像素，无模糊）
      · 标准屏：像素不变，SSAA 从 2× 翻到 4×
    - **`_logical_size()` 辅助方法**：统一处理物理像素→逻辑像素转换
      · 所有 `label.setFixedSize(pix.size())` 替换为 `_logical_size(pix)`
      · Pass 1 缩放从 100% base 出发时使用逻辑尺寸计算目标
    - **懒预渲染**：`_pre_render_100_all` 不再阻塞 UI
      · 仅立即渲染前 5 页（`PRE_RENDER_EAGER=5`），其余由 `_lazy_pre_render` 排队
      · QTimer.singleShot(50ms) → 每页间隔 10ms，后台异步填充 100% base 缓存
      · `_cancel_deferred_renders` 终止懒渲染循环
    - **缓存内存上限**：250MB → 400MB，适应 HiDPI 下 4 倍物理像素增长
    - **含入方式优化**：`pix.width // RENDER_SCALE` → `int(round(pix.width / RENDER_SCALE))`
    - BUGFIX: HiDPI pixmap 设置到 label 后显示 2 倍过大
      — 所有 label 尺寸计算统一使用 `_logical_size()` 取逻辑像素
    - BUGFIX: Pass 1 缩放使用 `base.size()` 取物理像素 → 计算结果 2 倍逻辑尺寸
      — 改为 `_logical_size(base)`，缩放后 label 尺寸正确
    - BUGFIX: pre-render 全部页面导致大 PDF 打开缓慢
      — 改为懒加载：先渲染 5 页，其余后台排队

  Ver 1.3.15 | 矢量级画质重写 — 精确分辨率渲染 + MuPDF 原生 AA + 捏合加速

    - **SSAA 超采样完全移除**：不再渲染到 4× 再 Bilinear 缩小
      · MuPDF 内置子像素抗锯齿在任何分辨率下都产生矢量级画质
      · SSAA → downscale 实际上用一个不必要的 Bilinear 滤镜软化了 MuPDF 的输出
      · Acrobat/WPS 的做法：直接在目标分辨率渲染矢量 PDF 内容
    - **精确分辨率渲染**：mat = fitz.Matrix(zoom * dpr, zoom * dpr)
      · Retina 100%：直接渲染 1190×1684 物理像素（≈144 有效 DPI）
      · 标准屏 100%：直接渲染 595×842 物理像素（MuPDF AA 保证质量）
      · 无 oversampling、无 downscale、无 Bilinear 损失
    - **MuPDF AA 级别最大化**：_configure_mupdf_aa() 在打开 PDF 时调用
      · fitz.Tools.set_aa_level(8) — 8 位子像素边缘平滑
      · fitz.Tools.set_text_aa_level(8) + set_graphics_aa_level(8)
    - **缓存内存追踪修正**：_cache_put 现在正确计算 HiDPI 物理内存
      · Qt6 中 pix.width() 返回逻辑像素，* dpr² 得到物理像素内存
      · 淘汰时也使用 devicePixelRatio() 校正释放量
    - **捏合缩放加速**：手势结束后 180ms → 40ms 触发高清渲染
    - BUGFIX: HiDPI 缓存内存低估 — dpr=2 时 pix.width()×pix.height()
      只返回逻辑像素（= 物理/4），导致 400MB 限制实际允许 1.6GB
      — 乘以 dpr² 得到真实物理内存占用量


  Ver 1.3.16 | 布局修复 + QPinchGesture 捏合缩放

    - **Scroll 布局重叠修复**: _layout_labels 从 PDF页面尺寸x当前缩放计算预期尺寸
      w,h = int(pw*z), int(ph*z) 精确 Y 定位，消除级联偏移和重叠
    - **QPinchGesture 捏合**: grabGesture(PinchGesture) + event() override
      缩放公式: delta_pct = (scaleFactor - 1.0) * 250, GestureFinished 时 40ms 高清渲染
    - **Ctrl+滚轮单独处理**: wheelEvent 仅处理 Ctrl+wheel, 不再探测捏合
    - **代码清理**: 移除 e.phase()/pixelDelta/angleDelta 探测, 移除 _pinch_acc
    - BUGFIX: 500页 PDF 未预渲染页面在缩放后位置偏移/重叠
    - BUGFIX: 捏合不响应/误触发 — 原生 QPinchGesture 替换 wheelEvent 探测


  Ver 1.3.17 | pinch 修复 + 渲染管线异步化 + tooltip 重做

    - **捏合缩放回归稳定方案**: QPinchGesture(启动崩溃)和动画引擎(破坏基础缩放)均移除
      恢复 v1.3.15 验证过的 phase 检测 + 累加器, 响应阈值 120→60 (2x 跟手)
    - **渲染管线五项优化**:
      1) Pass 1 仅缩放可见页 (3页), 不再遍历全部 label (500p: 5ms→<1ms)
      2) PAGE_DEBOUNCE_MS 150→80ms, scroll stop 响应更快
      3) _render_visible_range_async 逐页 yield 事件循环, 0ms UI 阻塞
      4) 键盘翻页 next/prev/go_to_page 立即触发高清渲染
      5) _do_debounce_calibration 无条件触发渲染 (zoom 脏数据自动刷新)
    - **自定义 tooltip 重做**: QLabel(11px) 跟随光标, 0.5s 显示/5s 消失
      移除 Qt 系统大弹窗: _store_tooltips() 抓取文字→setToolTip("") 清空
      缩放按钮文案: 1%→5%
    - **项目环境**: requirements.txt 添加 pyinstaller, 创建项目 .venv (Python 3.14)
    - BUGFIX: 滚动停止后长时间等待高清渲染 — 异步化 + debounce 80ms
    - BUGFIX: 键盘翻页后无高清渲染 — nav 末尾加 _schedule_render_visible(0)
    - BUGFIX: Qt 系统弹窗与自定义弹窗双重显示 — setToolTip("") 清空
    - BUGFIX: QPinchGesture 导入崩溃 — 移除, 回归 wheelEvent phase 检测
    - BUGFIX: 动画引擎破坏基础缩放 — 移除, 回归简洁累加器方案


----------------------------------------------------------------

Ver 1.4.0 | 2026-06-27
----------------------------------------------------------------

  Ver 1.4.0 | Grid View 编辑模块 — Core page_editor + GUI 编辑模式 + CLI/MCP 页面操作
    - Core 层 page_editor.py (新建):
      PdfPageEditor 命令模式撤销/重做 (MAX_UNDO_DEPTH=50)
      delete_pages / rotate_pages / move_pages / extract_pages
      on_change() 观察者模式事件回调
      page_count / page_order / page_rotation / can_undo / can_redo
    - **Grid 自适应填充**:
      min(scale_w, scale_h) 保证页面绝不超出网格边界且最大化填充
      网格 A4 比例 1:√2, 水平间距 20px, 垂直间距 30px
      黑白/横向/长条/正方形页面均居中不溢出
      page_num_label 半透明灰色页码在网格正下方
    - **编辑模式分层**:
      非编辑状态: 仅缩略图, 双击跳转连续滚动
      点击 Edit toggle: 动态初始化编辑模块 (toolbar + 选择 + 快捷键)
      关闭编辑: 卸载模块, 释放事件监听
    - **编辑工具栏** (48px):
      Select / Rotate 90° / Delete (二次确认) / Extract / Export / Undo / Redo
      无选中时操作按钮禁用, undo/redo 根据栈状态启用
    - **选择交互**:
      单击选中 / Ctrl+单击多选 / Shift+单击范围选 / Ctrl+A 全选 / Esc 清空
      矩形框选: 空白处拖拽矩形选中框内页面, 蓝色边框 2px
    - **拖拽排序**:
      单击已选页面→拖拽到目标→释放→Core.move_pages() 重排序
      Ctrl+Z/Ctrl+Shift+Z 撤销/重做排序
    - **CLI 新增 8 命令** (所有支持 --json 输出):
      delete-pages / rotate-pages / move-pages / extract-pages
      page-undo / page-redo / page-history / page-list
    - **MCP 工具 16→23**:
      pdf_delete_pages / pdf_rotate_pages / pdf_move_pages / pdf_extract_pages
      pdf_undo / pdf_redo / pdf_history
    - BUGFIX: 软件崩溃 — _destroy_labels 未回收 _page_num_label, 内存泄漏


  Ver 1.4.1 | Grid Editor 稳定版 — 编辑/撤销/保存闭环 + UI 修复
    - Editor 共享 fitz.Document: GUI 与编辑器共享同一份 PDF 对象, 删除/旋转/排序立即可见
    - Snapshot 撤销重做: 每次操作前拍完整 PDF 快照入栈(MAX=50), 撤销恢复快照, 重做再应用
    - 保存闭环: 关闭文档/退出编辑/软件退出均检测未保存修改, 弹窗三选项(取消/不保存/另存为)
    - 不保存还原: _enter_edit_mode 拍摄原始快照, 选不保存→_revert_to_original 回滚全部编辑
    - 导出为菜单: PDF/JPG 图片/Word .docx/PowerPoint .pptx 四种格式可选
    - Edit 按钮仅网格模式可见, 连续滚动自动隐藏; close_document 重置模式+清状态
    - 编辑工具栏 10 按钮: Select/Sort/Rotate/Delete/Extract/Export/Print/SaveAs/Undo/Redo
      所有按钮附带鼠标悬停小弹窗说明功能(0.5s延迟+5s消失+200ms离开保护)
    - 排序: 选中页面(支持Shift多选)→点 Sort→拖拽到目标位置释放→move_pages 重排序
      无选中点 Sort 弹出提示; Select 按钮退出排序模式恢复选择
    - 自适应填充: min(scale_w,scale_h)保证各比例页面(A4横竖/A5/长幅/方块)居中不溢出
    - 关闭文档→重置为 Scroll 模式+欢迎界面; 编辑模式关闭 PDF→先退出编辑再回到欢迎
    - 缩放记忆: 从连续滚动进网格保存 zoom, 切回或双击跳转恢复原缩放数值
    - 旋转累积: 清除被旋转页面全部缓存, 每次点击可见 90°→180°→270°→0° 完整旋转
    - UI 合并: PDF 工具 tab 移除, 所有工具按钮并入合并与转换右侧面板(QScrollArea)
    - 右侧面板默认宽度加大(QSplitter 3:1→2:1), 小窗口自动出现垂直滚动条
    - 网格缩略图缓存隔离: force_fit 缩略图用独立 thumb_{w}_{h} 键名避免与 100% base 冲突
    - 页码标签生命周期管理: _hide_page_numbers 在切换模式/重建标签/关闭文档时回收
    - i18n: 编辑工具栏全部按钮+提示+导出菜单(10键)中英文双译; Edit 按钮切换语言正常
    - MCP 工具 16→23: 新增 7 个页面编辑工具(delete_pages/rotate_pages/move_pages/
      extract_pages/undo/redo/history)
    - CLI 新增 8 命令: delete-pages/rotate-pages/move-pages/extract-pages
      page-undo/page-redo/page-history/page-list(全部支持 --json 输出)
    - requirements.txt 添加 pyinstaller, 项目 .venv(Python 3.14)完整可移植
    - BUGFIX: 编辑操作无效果 — PdfPageEditor 独立 fitz.Document 副本, GUI 看不到修改
      → 改为共享同一个 doc 对象, GUI 和 editor 双向同步
    - BUGFIX: 删除后页面重复出现 — 残存缓存(_clear_cache 删除时未清)导致旧索引误匹配
    - BUGFIX: 撤销闪退 — _restore_snapshot 后 GUI self.doc 未同步 editor._doc
    - BUGFIX: 重做闪退 — _edit_redo 缺少 doc 同步和缓存清除
    - BUGFIX: 编辑器不保存退出未还原 — 缺少 _revert_to_original(_original_snapshot)
    - BUGFIX: 编辑按钮切中文后仍显示英文 — _normal_label/_editing_label 未走 i18n
    - BUGFIX: 导出后灰屏 — 文件对话框后未调 _layout_labels()
    - BUGFIX: 关闭文档后 toolbar 残留+Edit 按钮仍显示 — close_document 未重置 ViewMode
    - BUGFIX: 排序按钮无响应 — page_container 缺 setMouseTracking(true)
    - BUGFIX: 警告弹窗弹出两遍 — close_document→leave_edit 重复弹+MainWindow 空文档误弹
    - BUGFIX: _revert_to_original 闪退 — 方法内 fitz.open 缺 import fitz
    - BUGFIX: Tooltip 立刻消失 — QTimer.singleShot bound method GC + 缺离开延迟 _hide_timer
    - BUGFIX: 编辑栏无悬浮弹窗 — 9个编辑按钮未调用 _hover_on 注册鼠标跟踪
    - BUGFIX: 缩放弹窗 tooltip 重复 — Qt 系统 ToolTip 与自定义 QLabel 同时弹出, _store_tooltips 清空 setToolTip


  Ver 1.4.2 | macOS 系统主题自适应 + Windows 编译就绪
    - macOS 自动跟随系统主题(暗色/亮色): _detect_dark_mode 读取 AppleInterfaceStyle
    - Windows 强制亮色主题: sys.platform==win32 返回 False
    - QPalette 全局配色: 暗色/亮色双套色板(Window/Base/Button/Text)
    - 阅读器工具栏/编辑工具栏/滚动条/文件标签/欢迎按钮/状态栏全量适配
    - _dc(dark,light) 辅助函数: 43 处调用覆盖所有硬编码颜色(无遗漏)
    - 亮色主题按钮: bg=白色+深度边框(#a0a0a0)保证可见, 暗色保持原样
    - 不涉及渲染管线/缓存/缩放逻辑
    - BUGFIX: 循环导入崩溃 — gui/pdf_reader_widget._dc()中import main导致启动失败
      → 改为模块属性 _DARK 由 main.launch_gui 在创建 widget 前注入
    - BUGFIX: 亮色主题按钮隐形 — bg=#e8e8e8与页面背景#f5f5f5无色差
    - BUGFIX: 欢迎屏幕 Load file 按钮不跟随亮色 — 暗色值硬编码
    - BUGFIX: main_window 右侧面板+status_label 暗色硬编码
    - Windows 构建就绪: build_windows.spec 新增12个 hiddenimports
      (core.page_editor/gui全部模块), build_windows.py 补 pywin32 依赖


----------------------------------------------------------------

Ver 1.5.0 | 2026-09-18 — 正确性与性能大修（功能 / 性能 / UI 全方位加固）
----------------------------------------------------------------

  Ver 1.5.0 | 大迭代第一轮：全链路加固
    - 阅读器改为窗口化渲染：打开 PDF 只渲染可见页 + 前后 3 页缓冲，不再全量预渲染
      —— 120 页文档打开耗时 60ms→27ms，缓存驻留 240MB→8MB，1000 页文档不再爆内存
    - 阅读器窗口化重建：窗口尺寸变化时重算布局并重新居中（原来切标签页/缩放窗口后
      页面停留在旧几何位置）；Grid 模式行高按视口自适应（一屏可见约 2 行缩略图）
    - Grid 框选重写：拖拽框选只在选区真正变化时刷新受影响的缩略图，
      不再每次 mousemove 重建整页网格（10 页文档 0.87ms→0.01ms/事件）
    - Grid 只渲染视口内缩略图（原来一次性渲染全部页），命中检测由 O(n) 扫描改为 O(1)
    - 框选改为 Finder 式矩形选区 + 蓝色 rubber band 可视化 + 拖拽排序橙色落点提示
    - 背景基准渲染改为有界环形队列（当前页 ±3），滚动/翻页时自动重新布防
    - 打开文档默认 fit-height（原来 100% 会裁掉半页），缩放/适应模式切换后保持滚动位置
    - 缩放提示浮层改挂到滚动视口，位置不再被上下两条工具栏顶偏
    - BaseWorker 重写：协作式取消（threading.Event + OperationCancelled），
      cancel() 不再阻塞 GUI 线程 5 秒（原 wait(5000) 直接卡死界面）
    - BaseWorker 超时逻辑修正：超时即中止操作并只报告一次
      （原来只发 error 不停任务，每个进度回调重复弹窗，实测 2 秒内 75 个错误弹窗）
    - 批量处理失败隔离：单个文件失败不再中断整批，结束后汇总成功/失败数量与原因
    - 结果弹窗改为延迟一个事件循环弹出，消除队列信号处理器内嵌套模态循环导致的死锁
    - Office 检测移出 GUI 线程（原来构造窗口时同步 osascript 探测最长阻塞 30 秒）
    - 输出路径父目录自动创建（原来 merge/encrypt/to-word 等直接抛 [Errno 2]）
    - PDF→Word 修复 block 元组字段顺序错误（原把 block_no 当 block_type，第一段之后
      的文字块或被当图片渲染、或被静默丢弃；5 段页面只输出 1 段）
    - PDF→Word 字体大小判断改为每页一次 dict 提取（原来每个 block 重新解析整页，O(n²)）
    - PDF→Excel 表格提取修复：改用 tab.extract()（原调用已不存在的 TableHeader.explicit_outer，
      表格页会整个从输出消失）；表格页过多矢量线段时跳过检测（原单页可阻塞 77 秒）
    - PDF→PPT 临时文件改用 TemporaryDirectory，异常路径不再泄漏；DPI 统一校验 36-1200
    - CJK 静默损坏修复：文本文件/Word/PPT/Excel 回退渲染与水印改走 MuPDF 内置 CJK 字体
      （原来中文被 base-14 Latin-1 字体替换成 ·····，混合合并"成功"但内容是垃圾）
    - 水印重写：透明度真实写入图形状态（原来 opacity 参数被完全忽略）、
      角度按真实旋转矩阵应用（原来 45° 会被归零成水平）、中文水印正常
    - 加密加固：空密码直接拒绝（原来会"加密"出无需密码即可打开的文件）、
      优先 AES-256（新增 cryptography 依赖，原为 RC4-128）、info 可正确报告加密状态
    - 压缩档位真实生效：lossless/medium/max 分别对应结构优化与图片重采样重编码
      （原来三档完全相同）；CLI 新增 --mode
    - extract_text 改用 PyMuPDF 优先（CJK 与版式更好），pypdf 作为回退，
      去掉未声明依赖 pdfplumber 的死分支；页分隔符不再硬编码中文
    - from_images 按图片 DPI 换算页面尺寸（原来 3000×2000 图片会生成 41.7 英寸巨页）
    - 混合合并：转换结果逐个校验可读性后再合并（0 字节/损坏 PDF 不再让整次合并失败）；
      temp_dir 登记到清理表（原来每次混合合并在系统临时目录留下用户文件副本）
    - MCP 协议修复：屏蔽 PyMuPDF 一次性 stdout 提示（会污染 JSON-RPC 流导致客户端失步）、
      tools/call 期间重定向 stdout、空行/坏 JSON 回 -32700 而不是结束服务、
      遵循通知不应答规则、支持 ping、版本号取自单一常量
    - MCP pdf_split 描述与实现对齐（去除未实现的 custom ranges 承诺）
    - 页码解析统一到 core.utils.parse_page_ranges：支持 all / 1-5 / 1–5 / 1,3,5 /
      全角逗号 / 分号 / 省略号，越界丢弃，非法输入抛可展示的 ValueError
    - 旋转对话框：空页码范围直接拒绝（原来会输出与输入完全相同的文件却提示"完成"）
    - 拆分对话框：实时显示解析结果预览，非法行报错，访问器不再抛异常
    - CLI 新增 mixed-merge（混合文件合并此前只有 GUI/MCP 通道）
    - CLI 页面编辑历史持久化：page-undo/page-redo/page-history 改为按文件保存线性历史
      （原来新进程里栈为空，永远返回"无可撤销操作"）
    - CLI page-edit 参数校验：--pages 无匹配时非零退出并给出页数提示，非法写法友好报错
    - i18n 加固：tr() 格式化异常不再抛出（原来 {} 占位符配 kwargs 会在 Qt 槽内抛
      IndexError 并触发 PyQt6 qFatal 直接 Abort）；补齐 menu_merge 等缺失键；
      预览/关于/批处理/文件过滤/阅读器提示等硬编码文案全部接入 i18n
    - 语言切换刷新补全：act_merge、tools_group、office 状态、合并按钮计数文案
    - 合并按钮去掉重复 emoji，计数文案走 i18n（原来中文界面显示"3 files"）
    - 新增测试套件：tests/test_core.py（51 项，核心/CLI/MCP/i18n）、
      tests/test_gui.py（45 项，Worker/批量/对话框/阅读器性能/i18n 覆盖）
    - BUGFIX: 打开 PDF 渲染全部页面 — 移除 _pre_render_100_all 的 eager 全量预渲染
    - BUGFIX: 窗口缩放后页面不居中 — _on_resize 未调用 _layout_labels
    - BUGFIX: 从合并标签页打开 PDF 布局按 800×600 固化 — 增加视口变化检测重排
    - BUGFIX: 网格框选只在首格生效 — _grid_page_at_pos 用整格扫描且首个命中即返回
    - BUGFIX: 旋转后网格几何未失效 — 增加 _grid_geom 失效与缩略图重绘
    - BUGFIX: 旋转 90° 后缓存未全失效 — 单次 _invalidate_pages 扫描替代逐页扫描
    - BUGFIX: 旋转对话框空范围静默输出副本 — 校验改为拒绝空列表
    - BUGFIX: PdfPageEditor 撤销后文档句柄失效 — 统一通过 editor.doc 属性重读句柄
    - BUGFIX: page_editor.save 用默认参数保存导致删除的页面仍在 — 加 garbage/clean
    - BUGFIX: 编辑历史日志键用 1 秒精度 mtime — 同一秒内两次编辑互相覆盖，改用 mtime_ns
    - BUGFIX: 首次编辑不可撤销 — 历史初始状态在变更后才记录，改为变更前 seed
    - BUGFIX: 取消后状态栏被队列进度覆盖 — 进度/结果槽校验 sender 是否为当前 worker
    - BUGFIX: 批量结束仍显示取消按钮 — _set_busy 改用显式 show/hide
    - BUGFIX: 文件列表跳过提示为硬编码英文且漏统计 — 接入 i18n 并补全各类计数
    - BUGFIX: 提取文本/图片多文件时连续弹两次目录选择框 — 统一由 _run_batch 负责
    - BUGFIX: 取消拆分输出目录后写入当前工作目录 — Path("") 为 "." 恒为真，改为先判空串
    - BUGFIX: 水印 PDF 路径为空时 Path("") 通过校验 — 改为校验原始字符串 + is_file()
    - BUGFIX: InfoDialog 在 size_bytes=None 时崩溃 — 改为 or 0
    - BUGFIX: svg/ico 被归类为图片但没有转换器 — 从分类表移除，避免混合合并误报
    - BUGFIX: 阅读器打印在 Windows 用 shell start 且 Linux 无处理 — 改为 os.startfile/xdg-open
    - BUGFIX: MainWindow 重复定义 _pick_input_file — 删除死代码

----------------------------------------------------------------

Ver 1.6.0 | 2026-09-18 — 阅读器升级：搜索 / 目录 / 阅读位置
----------------------------------------------------------------

  Ver 1.6.0 | 大迭代第二轮：阅读器成为真正的阅读器
    - 新增 core/search.py：全文搜索（GUI/CLI/MCP 共用）——返回页码、命中矩形、文字与上下文；
      查询归一化（NFKC + 大小写折叠），支持大小写敏感、整词匹配、页码范围限制与命中上限
    - 阅读器查找栏（⌘/Ctrl+F）：命中计数、上/下一个、区分大小写、Esc 关闭
    - 搜索高亮直接绘制进渲染位图：普通命中黄色半透明，当前命中橙色描边
      （复用渲染缓存键，命中页缓存失效后重绘，缩放/网格模式同样生效）
    - 侧栏新增结果面板：逐条列出命中页码 + 上下文，点击跳转到该处
    - 新增目录（书签）侧栏：可折叠左栏（⌘/Ctrl+B），树形层级 + 目标页码 + 一键跳转；
      无书签文档显示说明文案而非空白
    - core.search.get_outline：解析 PDF 目录树（嵌套层级、粗体标记、目标页）
    - 阅读位置记忆：按文件记录页码与缩放模式，重新打开自动恢复（最多保留 40 个文档）
    - 阅读器快捷键：⌘F 查找 / ⌘B 侧栏 / ⌘+ ⌘− 缩放 / ⌘0 适应高度 / ⌘1 适应宽度 /
      ⌘G 与 ⇧⌘G 上下一个命中 / F3 / Home / End / ←→ 翻页
    - 双击页面在适应宽度与适应高度之间切换；PageUp/PageDown 保留原生平滑滚动
    - 新增 CLI 命令 search（-q 查询，支持 --case/--whole-word/--pages/--json，
      可 -o 输出 .json 或 .txt）与 outline（--json）
    - MCP 新增 pdf_search 与 pdf_outline 工具（23 → 25），pdf_split 补上
      page_ranges 参数使描述与实现一致，pdf_merge 的 input_files 加 minItems 校验
    - 侧栏与搜索栏颜色跟随主题（暗色/亮色），全部文案接入 i18n（中英双语）
    - 测试：tests/test_gui.py 新增搜索/目录/阅读位置 17 项（45 → 62），全套件 151 项
    - BUGFIX: 搜索计数标签属性名不一致（_search_count_label vs search_count_label）导致首次搜索崩溃

----------------------------------------------------------------

Ver 1.7.0 | 2026-09-18 — 功能扩展：页码 / 元数据 / 拼版 / 插入页面
----------------------------------------------------------------

  Ver 1.7.0 | 大迭代第三轮：从「转换工具」到「文档收尾工具」
    - PdfOperator.add_page_numbers：页码/页眉页脚盖章，模板支持 {n}/{total}/{page} 与任意文字，
      6 个位置（页脚/页眉 × 左中右）、自定义起始编号与字号、可限定页码范围
    - 页码使用 TextWriter + CJK 字体绘制，中文模板（"第 {n} 页 / 共 {total} 页"）正常显示
      （insert_textbox 只接受 base-14 字体名，中文会渲染成方框）
    - PdfOperator.set_metadata：修改标题/作者/主题/关键词/创建者/生成工具，
      只写入传入字段，未知键直接拒绝
    - PdfOperator.nup：2/4/6/8/9/16 页拼版到 A4/A3/Letter，页面内容保持矢量
      （show_pdf_page），4-up 及以上自动横向出纸
    - PdfOperator.insert_pages：把另一个 PDF 插入到指定位置或追加到末尾
      （MuPDF insert_pdf，保留链接与批注）
    - PdfOperator.extract_pages / delete_pages：全部走 PyMuPDF，保留链接/批注，
      拒绝把文档删空
    - GUI 新增「排版与信息」菜单：页码对话框（位置/模板/起始/字号/范围 + 实时预览）、
      文档属性对话框（仅提交改动字段）、N-up 选择、插入页面（带目标页数校验）
    - CLI 新增 6 命令：add-page-numbers / set-metadata / nup / insert-pages /
      extract-pages-fitz / delete-pages-fitz，统一 --json 结构化输出
    - MCP 新增 4 工具（25 → 29）：pdf_add_page_numbers / pdf_set_metadata /
      pdf_nup / pdf_insert_pages
    - i18n 新增 40 余键（Compose 菜单、对话框、位置模板、提示文案）中英双语
    - tests/test_core.py 新增 24 项（页码含 CJK、元数据往返、拼版尺寸与页数、
      插入偏移、删除/提取边界与参数校验），全套件 193 项

----------------------------------------------------------------

Ver 1.8.0 | 2026-09-18 — 界面打磨与收尾
----------------------------------------------------------------

  Ver 1.8.0 | 大迭代第四轮：界面清爽化 + 代码收尾
    - 右侧面板瘦身：7 个格式转换按钮收进新增的「转换(&T)」菜单，面板只保留
      文件列表真正需要的操作（合并/拆分/压缩/水印/加密/解密/旋转/信息）
      —— 按钮数 48 → 41，视觉噪音明显下降，功能一键可达
    - 文件列表新增汇总行："共 N 个文件 · 大小 · 页数 · 主要是 <类型>"；
      空列表给出操作提示而不是一片空白
    - 文件列表行显示页数（report.pdf (2.4 MB, 18p)），页数按需读取并缓存在列表项上
    - _update_summary 在增删/清空/拖拽排序后自动刷新，切语言同步刷新
    - 新增 fl_summary / fl_summary_empty / fl_tip_pages / kind_* 等 i18n 键（中英双语）
    - 代码收尾：移除 13 处未使用导入（core/utils、core/converters、core/merger、
      gui/main_window、gui/pdf_reader_widget、mcp/server）；删除无用常量 CACHE_TARGET_MB；
      把 gui 里直接赋值 core.utils._office_cache 的写法改为正式的 set_office_cache() 接口
    - tests/test_gui.py 新增 8 项（汇总行、页数显示、菜单结构、面板不再重复按钮），
      全套件 201 项（93 核心 + 70 GUI + 38 阅读器）

----------------------------------------------------------------

Ver 1.9.0 | 2026-09-18 — QA 加固：19 个缺陷修复 + 热点路径提速近 10 倍
----------------------------------------------------------------

  Ver 1.9.0 | 大迭代第五轮：以 QA 为主导的鲁棒性/准确性/性能加固
    - 阅读器内存上限真正生效：_cache_put 原先在遇到第一个 100% base 条目时直接
      break 退出淘汰循环，而刚打开的文档里全是这类条目，缓存因此无上限增长
      （实测 638MB 对 400MB 上限）。改为两级淘汰：先丢可再生的缩放/缩略图渲染，
      再丢屏幕外 base，_protected_pages 保护当前渲染窗口
    - 修正 devicePixelRatio 重复计量：Qt6 的 QPixmap.width() 已是设备像素，
      再乘 dpr² 会在 Retina 上高估 4 倍，使 400MB 上限实际只有 100MB（频繁误淘汰+重渲）
    - 缓存计量精确化：替换已有条目时先扣除旧尺寸（原先只加不减，计数持续偏低），
      _pixmap_bytes 统一为唯一实现
    - 缓存预算按机器自适应：取物理内存 1/4，钳制在 [128MB, 400MB]
    - 网格视图命中测试修正：事件坐标是标签局部坐标，_grid_page_at_pos 却按容器坐标
      解释，导致每个缩略图都解析成第 1 页（点第 5 页按删除 → 删掉第 1 页）；
      改为由布局循环直接传入页面索引
    - 幽灵索引防护：_on_grid_click 拒绝越界索引（原先 -1 会进入选中集，
      删除变成空操作却仍标记未保存，拖拽排序还会复制最后一页）；move_pages 同时
      过滤越界序号并去重
    - 拖拽排序不再崩溃：move_pages 会重建文档句柄，阅读器未重新读取，
      下一行在 Qt 槽函数里抛 ValueError("document closed") 导致进程 abort
    - 编辑后保持惰性渲染：_rebuild_labels_from_editor 原先在 GUI 线程把整本文档
      按 100% 全部重渲（100 页 189MB，1000 页约 1.9GB），现在只铺可见行
    - 布局热点优化：_layout_scroll 原先每页都读 doc[i].rect，新增 _page_rects 缓存
      （1000 页时布局 9.5ms → 0.94ms，捏合缩放一次 11.9ms → 1.55ms）
    - 切换文档重置每文档状态：新增 _reset_document_state，清空搜索结果/高亮/侧栏/
      编辑模式/_page_editor/_original_snapshot（原先旧文档的命中矩形会被画进新文档，
      页码可能显示 "61 / 3"，编辑还会作用在旧文档上）
    - 适应窗口缓存键补上视口尺寸：_on_resize 清理 "fh"/"fw" 过期条目
      （原先改变窗口后标签仍拿着旧位图，永远不会铺满）
    - open_pdf 拒绝 0 页 PDF（原先 self.doc[0] 抛 IndexError 并在菜单槽里 abort）
    - 渲染缓冲区显式拷贝：QImage 不拥有传入缓冲区，改为 bytes(pix.samples)
    - CLI 撤销历史真正可用：日志键原为 (路径, 大小, mtime)，命令自己的输出永远
      匹配不上输入，page-undo 总报"没有可撤销的操作"；改为路径+inode 派生并写入
      sidecar 持久记录，就地保存（os.replace 换 inode）后依然能找回历史
    - 页面编辑命令 -o 变为可选：省略或与 -i 相同即就地修改（原子替换，绝不留下
      半写文件），这是让跨进程 page-undo 成立的前提
    - 日志快照按字节预算裁剪（96MB，至少保留 3 个状态）并清理 7 天前的陈旧日志
      （原先 50 份完整副本常驻内存与磁盘，实测为文档的 50 倍）
    - 加密 PDF 统一给出可操作提示：新增 _reject_if_encrypted 覆盖 16 个入口，
      不再泄漏 "File has not been decrypted" / "document closed or encrypted" /
      "PasswordError" 等底层内部错误
    - 三通道页码语法统一：MCP 自带的 _parse_page_list 改为调用
      core.utils.parse_page_ranges（原先 1–3 / 1，3 / 3-1 / 1..5 在 CLI 可用、
      在 MCP 报错或崩溃）
    - MCP 参数校验与能力对齐：页码命令拒绝空选择、pdf_compress 补 mode 参数、
      pdf_history 改用持久日志（与 CLI 同源）、pdf_undo/redo 先 load_journal、
      全部范围类 schema 统一描述真实语法
    - 文本转 PDF 按真实字体测量折行：原先固定 6.5pt/字符，拉丁文比右边距短 85pt，
      中文超出页面 54pt；新增 _wrap_by_width（二分 + font.text_length）
    - 删除 core/converters.py 里的 print()：它会污染 MCP 的 JSON-RPC stdout 流
    - 表格检测门槛改为惰性计数，避免为判定上限而物化整页绘图列表
    - CLI 未知命令不再伪装成功：原先静默打开 GUI 并返回 0，现在打印相近命令建议
      并以退出码 2 结束
    - 页码范围解析拒绝残缺写法："1-" 原先被当作 "1"（静默操作了不同页面），
      负数页码给出专门提示
    - 阅读位置在 ⌘Q 正常退出时也会保存（原先只有点 ✕ 才记）
    - 强制终止的 worker 主动发出终态信号，界面不再永久卡在"忙碌"
    - 工具提示改用单个复用 QTimer（原先每次悬停新建一个，6 次悬停后 6 个定时器
      同时活跃，旧定时器会提前隐藏新提示）
    - JPG 导出遵循用户选择的文件名（原先选 photo.jpg 却写出 photo_p0001.jpg，
      成功提示还报不存在的那条路径）
    - 新增 tests/qa_adversarial.py：170 项对抗性探针（畸形/恶意输入、边界值、
      三通道一致性、像素准确性判据、内存与句柄上限、临时文件卫生）
    - 新增 tests/qa_reader_defects.py：37 项阅读器缺陷回归，每项都对应上面一个修复
    - 全套件：493 项（93 核心 + 70 GUI + 38 阅读器 + 85 阅读器综合 +
      37 阅读器缺陷 + 170 对抗），全部通过
    - Windows 改发安装包：新增 installer_windows.iss（Inno Setup 6，中英双语向导、
      许可协议页、开始菜单、可选桌面快捷方式、标准卸载程序），默认装到
      C:\Program Files\PDFeverything —— 与 README / mcp/README 的 MCP 配置路径一致
    - build_windows.py 增加第 5 步编译安装包并复制到项目根目录，新增 --skip-deps /
      --exe-only；ISCC 缺省时跳过安装包而不是让构建失败
    - 新增 .github/workflows/build-windows-installer.yml：在 windows-latest 上跑
      核心套件 → PyInstaller → Inno Setup，再对打包后的 exe 做冒烟测试（帮助文本 +
      真实 PDF 操作）后上传安装包产物（PyInstaller 无法从 macOS 交叉编译 Windows 产物）
    - resources/ChineseSimplified.isl：vendored 中文翻译（取自 issrc 的 is-6_7_1 标签，
      对应 CI 装的 Inno Setup 6.7.1；6.7.x 已把它挪到 Languages\Unofficial\ 下）
    - BUGFIX: tests/test_reader.py 中 3 项断言仍停留在 v1.6 之前（默认缩放 100%、
      全部页面都有 100% base），与"打开即适应页高 + 窗口化渲染"的既定设计冲突，
      已改为验证实际设计；同时清空 QSettings 的 reader_positions，避免跨次运行残留
      的阅读位置影响默认缩放断言
    - BUGFIX: journal_dir 在 core/page_editor.py 中重复定义两次，删除冗余定义
    - BUGFIX: Windows exe 执行 `PDFeverything.exe -h` 直接崩死 —— onefile + GUI 子系统
      （console=False）下 stdout 用的是本地代码页（英文系统 cp1252/cp437），
      print() 中文帮助文本抛 UnicodeEncodeError，而 windowed 模式把未捕获异常变成
      报错弹窗，命令行就此卡住（CI 冒烟测试里 180 秒被 kill）；
      新增 core.utils.make_output_encoding_safe()，保留控制台原编码、只把编不出的
      字符降级为 "?"，main.py / pdf_tool.py 两个入口在打印前调用

Ver 1.9.1 | 2026-10-28 — 用户报告修复：不规范 PDF 拆分报错 + 启动慢
----------------------------------------------------------------

  Ver 1.9.1 | 用户实测反馈：拆分报错 + 启动慢 + 逻辑审查
    - BUGFIX: 拆分/合并/信息等 pypdf 通道读"末尾被追加空 xref 段"的 PDF
      （用户实测文件 123.pdf）直接抛 "Could not read Boolean object"——
      pypdf 对 xref 容错极弱，而 pikepdf/PyMuPDF 都能正常读；
      新增 core.utils.open_pdf_reader()：先直接 pypdf，失败后用 pikepdf
      （attempt_recovery）在内存里重建一份再读，split / merge / info /
      watermark / encrypt / decrypt / rotate / extract_text 及 page_editor
      的 move/extract_pages 全部走该入口，pypdf 能修的用户文件不再报库错误
    - 启动提速 ~1.2s：PyMuPDF 2.x 在 import 时急切加载 pymupdf.layout +
      pymupdf4llm（连带 numpy / onnxruntime / tabulate / multiprocessing），
      本项目从不调用这些 AI 分析功能；新增 core/fast_pymupdf.py 用
      meta_path finder 把这两个包"装作未安装"，GUI 模块链导入
      1.43s → 0.26s，find_tables/渲染/文本提取不受影响；
      core/__init__.py 统一安装，未来需要 to_markdown 时可 unblock()
    - 启动提速 2：首次启动的 Office 探测（osascript/COM，最多 30s）在
      GUI 线程同步执行——`check_office_availability() is not None` 的守卫
      永远为真（探测函数从不返回 None），_ProbeWorker 是死代码；
      拆出 probe_office_cached()（只读缓存）/ probe_office_availability()
      （同步探测），_start_office_probe 缓存未命中时走后台线程，
      _check_office 改为只渲染缓存结果，绝不在 UI 线程探测
    - extract_text 语义收紧：MuPDF 打不开、仅靠 pikepdf 紧急修复打开且
      提不出任何文字的损坏文件，现在报错而不是静默写出 0 字节 .txt
      （正常空白/扫描 PDF 走 MuPDF 路径不受影响）；对抗性套件
      "extract_text rejects truncated PDF" 回归通过
    - fitz 别名的 FutureWarning（PyMuPDF 2.x 每次导入都打）按类别屏蔽，
      CLI/MCP stderr 不再被污染
    - 全套件 493 项复跑全绿；123.pdf 全命令（split/merge/extract-text/
      to-images/from-images/compress/rotate/encrypt/decrypt/to-word/
      to-ppt/to-excel/nup/页码/元数据/插入页面/搜索/书签/页列表/图片提取）
      + GUI 阅读器打开 + MCP pdf_split 实测通过
    - 打包瘦身：PyMuPDF 2.x 会在 import 时急切加载 pymupdf4llm/pymupdf.layout
      （连带 onnxruntime），PyInstaller 静态分析会照单全收打进安装包——
      build_windows.spec / PDFeverything.spec 的 excludes 增加
      pymupdf4llm / pymupdf.layout / pymupdf_layout / onnxruntime，
      hiddenimports 补 core.fast_pymupdf；installer_windows.iss 默认版本号
      同步 1.9.1；mcp/server.py 独立入口（python mcp/server.py）也先装
      fast_pymupdf 再 import pymupdf（实测独立启动 0.27s，不再加载 onnx）
    - 本机（Linux）PyInstaller 冻结烟测通过：onefile 产物 104MB，
      -h / info / split（123.pdf）/ mixed-merge / GUI 启动全部正常

Ver 1.9.2 | 2026-10-28 — 输出路径成为唯一权威，操作不再重复弹窗
----------------------------------------------------------------

  Ver 1.9.2 | 用户实测反馈：底部已有输出路径设置，点操作按钮仍弹输出对话框
    - 原设计里底部"输出路径"框只是弹窗的默认起始位置（记住上次存哪儿），
      每次点合并/拆分/压缩等仍要再确认一遍——等于设置两遍
    - 现在输出路径框是唯一权威：所有操作直接写入该目录，全程零弹窗；
      框内以 .pdf 结尾时合并结果保存为该文件，其余操作用其所在目录
    - 目标文件已存在时自动加 _1/_2 后缀，绝不静默覆盖
    - 浏览... 按钮改为选择输出目录；单文件转 Word/PPT/Excel、提取文字/
      图片、转图片的输出对话框全部移除，统一走输出路径框
    - 拆分/批量完成的提示现在带输出目录；路径框加了悬停说明（中英双语）
    - GUI 全套 70 + 核心 93 + 对抗 170 项回归通过

Ver 1.9.3 | 2026-10-28 — 合并完成后弹出完成提示（修复静默结束）
----------------------------------------------------------------

  Ver 1.9.3 | 用户实测反馈：点合并按钮，完成后没有任何提示
    - 根因：PdfOperator.merge 返回 None，而 GUI 完成回调只识别 dict/int/list
      结果——纯 PDF 合并从来没有过完成弹窗（这是老 bug，非 v1.9.2 引入）
    - merge 现在返回 {merged, output}，GUI 弹"合并完成：N 个文件 → 路径"
    - 兜底：任何 worker 返回 None / 无法识别的结果时，也会弹"操作完成"
      而不是静默结束——今后不会有任何操作完成后无声无息
    - 回归：GUI 70 + 核心 93 + 对抗 170 项全绿；混合合并/CLI/MCP 的 merge
      调用方均不受影响（返回值向后兼容）

Ver 1.9.3b | 2026-10-28 — Linux 正式适配：构建脚本 + AppImage
----------------------------------------------------------------

  Ver 1.9.3b | Linux 适配（代码层本就兼容——此前所有开发与测试均在 Linux 完成）
    - 新增 PDFeverything-linux.spec：onefile 打包，与 Windows spec 同样排除
      PyMuPDF AI 分析栈（pymupdf4llm/onnxruntime），额外排除 win32 系列；
      图标用 PNG（X11/Wayland 不认 .ico/.icns）
    - 新增 build_linux.py：环境检查 → 依赖 → PyInstaller → AppImage 四步，
      appimagetool 按架构（x86_64/aarch64）自动下载到 .buildcache，失败或
      未联网时优雅降级为仅出单文件二进制
    - AppRun 使用 #!/bin/bash（exec -a 是 bash 内建，dash 会报 "exec -a:
      not found"——首版 AppImage 实测踩坑后修正）
    - resources/PDFeverything.desktop：标准桌面启动器（打入包内 + AppDir 根）
    - 实测通过：onefile 104MB，CLI（info/split）与 GUI（窗口打开）均正常；
      AppImage 同样验证
    - README：平台徽章加 Linux、下载表加 AppImage 条目、构建文档补
      Linux 段落与"免构建直接跑"说明

    - BUGFIX: AppImage 双击无反应 —— 构建机缺 libxcb-cursor.so.0，PyInstaller
      打包 Qt6 的 xcb 平台插件时静态捆绑失败（构建日志有 WARNING 但未拦截），
      目标机器上 Qt 回退不到任何可用平台插件直接 abort。已安装 libxcb-cursor0
      重新构建验证；build_linux.py 构建后新增 GUI 启动烟测步骤

    - BUGFIX: 合并列表内拖动排序"能拖起来、松手无效" —— 列表覆写了
      dragEnterEvent/dropEvent 用于接收文件管理器拖入，但覆写版对不含
      文件 URL 的内部排序 mime 直接 ignore()，Qt 因此取消整个排序操作
      （真实鼠标 xdotool 实测复现：dragEnter source=列表自身被 ignore）。
      现在按 mime 分流：带 URL → 原有外部文件逻辑；带 model mime →
      委托 QListWidget 原生 handler 完成排序（注意 super() 会解析到
      FileListWidget(QWidget) 而非被绑定的 QListWidget，必须显式
      QListWidget.dragEnterEvent(self.list_widget, event)）；排序落地后
      补发 files_changed（InternalMove 走 remove+insert，rowsMoved 不会发）
    - 真实拖拽回归：xdotool XTEST 模拟，dragEnter accepted=True，
      rows [a,b,c,d] → [b,c,d,a]，自定义 UserRole 数据完整

  Ver 1.9.4 | 用户实测反馈：拖入 .doc 提示"所有文件转换均失败"
    - 根因：WordConverter 声明支持 .doc/.rtf，但纯 Python 回退用的是
      python-docx——它只读 OOXML zip 格式（.docx），遇到 OLE 二进制的
      .doc 直接 BadZipFile；.ppt/.xls 同理。Windows/macOS 有 Office COM/
      AppleScript 能读，Linux 无 Office 机器上这三类文件全灭
    - 修复：新增 _libreoffice_convert()（soffice --headless --convert-to
      pdf，独立 UserInstallation profile 避免与运行中的 LO 实例冲突），
      Word/PPT/Excel 三个转换器在 COM/AppleScript 之后、纯 Python 回退
      之前插入该路径——LibreOffice 对老格式的保真度也远高于重新排版
    - 纯 Python 回退限定真正支持的扩展名（.docx/.pptx/.xlsx）；老格式在
      无 Office 且无 LibreOffice 时给出可操作的中文报错（装 LibreOffice
      或另存为新格式），不再是含混的"转换失败"
    - 实测：真实 .doc（MS Word 97 格式）经 LibreOffice 转 PDF 成功；
      .docx 回退路径回归正常；无 LibreOffice 场景报错文案正确；
      GUI 端到端 doc→合并 PDF 成功；全套 333 项回归通过

    - BUGFIX: 点合并后进度条先显示上一次的 100%、过几秒才跳回低进度开始转换
      —— _on_finished 把进度条定格在 100 并保持可见逻辑外的旧值，而
      _set_busy(True) 只 show() 不清零；混合合并的首个进度事件（"转换中"）
      要等第一个文件转完才发，中间几秒进度条就停在 100 造成"已完成"假象。
      现在任务启动时进度条清零、状态栏显示"处理中…"（新增 i18n 键
      status_working）；实测连续两次合并，第二次开始瞬间 bar=0

  Ver 1.9.5 | 用户实测反馈两处界面易用性问题
    - 窗口拖到最小宽度(900)时右侧操作面板被裁掉约 46px——合并按钮被切且无
      横向滚动条可救。双保险修复：右侧面板 setMinimumWidth(292)（最宽文案
      "🔀 合并为 PDF (200 个文件)" + 边距），窗口最小宽度 900→1060
      （= 文件列表最小 ~306 + 右面板 292 + 分割条与边距）
    - 文件列表行高从 17px 提到 34px（QListWidget::item 加 8px 上下 padding），
      原行高点击命中困难

    - BUGFIX(续): 右侧面板仍被遮 + 底部出现横向滚动条 —— 两个叠加因素：
      ① "Office 状态"说明文字 wordWrap=False，整行宽度撑高面板最小需求，
      窄窗口下文字被右缘切断；改为 setWordWrap(True) 自动换行
      ② 垂直滚动条出现时会挤占视口宽度（406-8 < 406），QScrollArea 于是
      弹出横向滚动条来回抖动；横向滚动已被面板最小宽度保证无意义，直接
      ScrollBarAlwaysOff 关闭
    - BUGFIX(续): 列表行加高后，选中/悬停的行会跳回 17px 原始样式 ——
      QSS 的 ::item padding 规则在出现 :selected/:hover 专属规则后不再
      适用；补齐这两个状态的 padding
