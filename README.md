# PPT 增强插件套件

把多个独立的 PowerPoint 宏加载项，收拢到一个统一的功能区标签 **「增强插件」** 下。

**为什么必须"改定义"而不是"改设置"**：PowerPoint 会把每个已加载 `.ppam`
自身 `customUI` 里声明的标签**原样显示**出来，Office 没有提供任何设置项可以把
多个加载项的标签归并到一起。所以"合并功能区"只能在 customUI 定义层面完成 ——
这正是本套件做的事：把各插件的 VBA 汇进一个工程，只声明**一个**标签、按插件分分组。

---

## 一、当前进度

| 阶段 | 内容 | 状态 |
| --- | --- | --- |
| 1 | 搭骨架 + 迁入「导出图片」 | ✅ 完成，回归测试全绿 |
| 2 | 迁入「选页另存」（原「选取另存.ppam」，顺带修无窗口打开的 bug） | ✅ 完成，回归测试全绿 |
| 3 | 迁入「图片替换」（原「PPT原位替换插件V3.ppam」）+ 修复透明度/样式/图层丢失 | ✅ 完成，回归测试全绿 |
| 3.5 | 「图片替换」支持**编组（组合）内的图片**（原实现在组内直接报"当前类型6"） | ✅ 完成，回归测试全绿 |
| 3.6 | 「导出图片」按钮补图标（原 `imageMso="ExportToFile"` 是不存在的 ID，图标静默不显示） | ✅ 完成，截图确认三按钮图标齐全 |
| 4 | 收尾：卸载被取代的旧加载项、消除重复入口（实现 `--disable-legacy`） | ✅ 完成，功能区只剩「增强插件」一个自定义标签 |

阶段 3 交付后的功能区形态：

```
… [插入] [增强插件] ← 本套件新增
              ├─ 分组「导出图片」
              │     └─ 大按钮「导出为图片」
              ├─ 分组「选页另存」
              │     └─ 大按钮「导出选中幻灯片」
              └─ 分组「图片替换」
                    └─ 大按钮「原位替换图片」
```

> 此时旧的 `PptImageExport.ppam` 仍在 AddIns 目录里（尚未卸载），
> 所以你会在功能区同时看到旧的「图片导出」标签。**确认新套件工作正常后**，
> 在第 4 阶段统一移除旧加载项。

---

## 二、关于「标签页能不能加图标」

**结论：PowerPoint 的自定义功能区标签（tab）不支持显示图标。**

这是 Office Fluent UI 的硬限制，不是本项目的取舍：Ribbon XML 规范里
`<tab>` 元素只有 `label` 一类属性，**没有** `image` / `imageMso`。
凡是走 XML 定制的功能区（VBA 加载项、VSTO、COM 加载项）都受此限制。

图标能放的位置只有这三处，按推荐度排列：

| 位置 | 能否带图标 | 说明 |
| --- | --- | --- |
| 按钮 button | ✅ 可以 | `imageMso="内置图标名"` 或 `image="自定义图片"`（配 `loadImage` 回调）。功能区视觉的主角。 |
| 分组 group | ✅ 可以 | 写 `imageMso`。但**只在功能区变窄、该组被折叠成下拉按钮时才显示**，正常展开状态下不显示。 |
| 标签 tab | ❌ 不能 | 规范不支持。变通办法是在标签文字里加 Unicode 符号，例如 `▣ 增强插件`、`◈ 增强插件`。 |

如果你确实要"标签上带图标"的效果，只有一条技术路线：改用 **Office Web 加载项**
（`manifest.xml` 里的 `IconUrl` 能让自定义标签显示图标）。代价是要 HTTPS 托管、
完全不同的技术栈、且**不能复用现有 VBA 代码** —— 为一个视觉效果换架构，不划算。

本套件的做法：图标集中在**大按钮**上（每个插件的主按钮都配 `imageMso`），
标签文字若要加装饰符号，改 `suite.json` 里 `tab.label` 一处即可全局生效。

> **`imageMso` 的 ID 必须逐个实测**（踩坑记录第 11 条）：
> 写一个"看起来应该存在"的 ID（如 `ExportToFile`）**不报任何错，只是图标不显示**。
> 验证方法见「阶段 3.6」验证记录 —— Excel 的 `GetImageMso` 初筛 +
> VBA `SavePicture` 导出目视确认 + 实机截图三步闭环。
> 当前定案：导出图片=`ExportToAnimatedGif`、选页另存=`FileSaveAs`、
> 图片替换=`PictureInsertFromFile`。

---

## 三、目录结构

```
ppt-addin-suite/
├─ suite.json                    套件元信息（名称 / 标签 / 产物文件名 / legacy 清单）
├─ docs/
│  ├─ 使用说明.html              面向使用者的图文手册（截图 base64 内嵌，单文件分发）
│  ├─ 使用说明.md                同内容的 Markdown 版（配图引用 screenshots/ 相对路径）
│  ├─ PPT 增强插件套件 · 使用说明.pdf   A4 打印版（无头浏览器导出）
│  ├─ _build_doc.py              手册生成脚本（改内容后重跑，一次产出 HTML + PDF + MD）
│  ├─ _build_md.py               Markdown 版转换脚本（由 _build_doc.py 自动调用）
│  └─ screenshots/               手册配图原始 PNG（真实界面截图）
├─ src/
│  ├─ _core/                     公共层（不属于任何分组）
│  │  ├─ plugin.json
│  │  └─ m_Core.bas              套件标识、统一错误出口、通用 Ribbon 回调
│  ├─ image-export/              插件① 导出图片
│  │  ├─ plugin.json
│  │  ├─ m_ImageExport.bas       导出引擎（589 行）
│  │  └─ frmImageExport.code.vba 设置窗体（783 行）
│  ├─ slide-export/              插件② 选页另存
│  │  ├─ plugin.json
│  │  └─ m_SlideExport.bas       选页另存（177 行）
│  └─ image-replace/             插件③ 图片替换
│     ├─ plugin.json
│     └─ m_ImageReplace.bas      属性快照 → 替换 → 逐项还原
│                                ＋ 组合（编组）内替换（1287 行）
├─ customUI/customUI.xml         构建时自动合成（不用手写）
├─ build/
│  ├─ build_suite.py             套件构建器（含 --disable-legacy 遗留清理）
│  └─ migrate_image_export.py    阶段 1 的一次性迁移脚本（留档）
├─ test/
│  ├─ modSelfTest.bas            自检驱动（仅 --selftest 注入）
│  ├─ selftest.vba               窗体自检函数（仅 --selftest 注入）
│  ├─ selftest_slide.vba         选页另存自检函数（仅 --selftest 注入）
│  ├─ selftest_replace.vba       图片替换自检函数（仅 --selftest 注入）
│  ├─ run_suite_test.py          回归测试：三个插件端到端 + OOXML 解包比对
│  ├─ verify_build_code.py       产物代码核对：源码 vs 包内代码逐行 diff
│  ├─ probe_oracle.py            编译二值预言机：这份代码到底能不能编译
│  ├─ probe_suite.py             整套编译体检（会点掉弹窗，会卡住，慎用）
│  ├─ probe_compile.py           单模块编译体检
│  ├─ probe_inject.py            验证「注入阶段不报错」→ 说明必须真正触发 Run
│  ├─ probe_vbe.py               早期尝试（PostMessage 点不掉模态框，已废弃留档）
│  ├─ lint_undeclared.py         静态扫描未声明标识符
│  ├─ diag_replace.py            图片替换专项诊断（抓 #32770 子控件文本）
│  ├─ verify_live.py             真机验证：加载项加载 + 宏可执行
│  └─ output/                    测试输出的图片与样张
└─ dist/
   ├─ PptEnhanceSuite.ppam       正式产物
   └─ PptEnhanceSuiteTest.ppam   自检产物（不安装，仅供测试）
```

### 排查工具链（按介入顺序）

| 工具 | 回答的问题 | 特点 |
| --- | --- | --- |
| `lint_undeclared.py` | 有没有引用到未声明的标识符？ | 纯静态，最快，但有噪声（类型名 / 枚举常量） |
| `probe_oracle.py` | **这份代码能不能编译？** | 非阻塞二值预言机。注意它的覆盖范围＝`DEFAULT_CALLS` 里真正进入过的过程，见下方说明 |
| `probe_suite.py` / `probe_compile.py` | 整套 / 单模块有没有编译错误？ | 会尝试点掉模态框，成功则能拿到具体错误文本 |
| `verify_build_code.py` | 产物里的代码和源码一致吗？ | 逐行 diff，专抓「注入时正常、保存后损坏」 |
| `run_suite_test.py` | 功能到底对不对？ | 端到端 + 独立解包 OOXML 双重证据 —— **唯一能证明"功能正确"的那一环** |

> **为什么必须有 `probe_oracle.py` 这种"非阻塞"设计**：VBA 的编译错误对话框是
> 模态的，一旦弹出就会把 `Application.Run` 和整个自动化进程一起卡死，只能
> `taskkill` 强杀。所以预言机不去"点掉"对话框，而是**抓到对话框即判定失败**，
> 随即杀进程退出。
>
> ### ⚠ VBA 的真实编译规则（2026-10-10 实测纠正，务必按这条写探针）
>
> 早期这里写的是「编译按**静态调用图**展开，只要探针调用了入口，
> 整条可达链都会被编译」。**这是错的**，而且错得很危险 ——
> 它让一批从未被真正进入的代码（例如组内替换那条路径）长期处于"体检通过"
> 的假象里。实测结论是：
>
> **过程「首次被进入」时才编译，编译单位就是这一个过程。**
> 静态可达 ≠ 会被检查。探针没真正走到的过程，里面的语法错误根本不会被报出来。
>
> 验证方法（可复现）：把 `tmp = "PIR_TMP_" & m_renameIdx` 改成
> `tmp = 1 +`（明显语法错误）放进 `IR_RenameTree`：
>
> | 探针调用方式 | 结果 |
> | --- | --- |
> | 只调 `ImageReplace_SelfTest` / `..._GroupSelfTest`（路径不存在 → 提前返回） | ❌ **漏检**：报「编译通过」 |
> | 再直接 `IR_RenameTree Nothing` 等逐条进入 | ✅ 报「编译错误: 语法错误」 |
>
> 所以 `probe_oracle.py` 的 `DEFAULT_CALLS` 现在把这些过程**逐条直接列出来**
> 调用（哪怕传 `Nothing` / 不存在的路径让它立刻返回）—— 进入动作本身就会
> 让它整体编译。**新增过程后要同步把名字加进 `DEFAULT_CALLS`**，
> 否则新代码等于没体检。
>
> 另外注意：预言机只保证"已覆盖的过程没有语法/未定义变量错误"，
> **功能正确性必须靠 `run_suite_test.py` 真跑一遍**。

---

## 四、新增一个插件（三步）

**① 建目录**：`src/my-plugin/`

**② 放源码 + 清单**：

```
src/my-plugin/
├─ plugin.json
├─ m_MyPlugin.bas            标准模块
└─ frmMyPlugin.code.vba      窗体（若有）
```

`plugin.json` 示例：

```json
{
  "id": "my-plugin",
  "order": 20,
  "title": "我的插件",
  "group": { "id": "grpMyPlugin", "label": "我的分组", "imageMso": "HappyFace" },
  "components": [
    { "type": "module", "name": "m_MyPlugin", "file": "m_MyPlugin.bas" },
    { "type": "form",   "name": "frmMyPlugin", "file": "frmMyPlugin.code.vba",
      "design_width_pt": 392, "design_height_pt": 312 }
  ],
  "buttons": [
    { "id": "btnMyPlugin", "label": "干点活", "size": "large",
      "imageMso": "HappyFace", "onAction": "Ribbon_MyPlugin_Run",
      "screentip": "简短提示", "supertip": "鼠标悬停时的详细说明" }
  ]
}
```

**③ 构建**：

```bash
python build/build_suite.py
```

构建器会自动：合成功能区 XML → 检查命名冲突 → 注入模块与窗体 → 打包 `.ppam`
→ 校验包结构 → 安装并登记注册表。**不需要改动任何构建代码。**

### 命名规范（避免跨插件冲突）

| 对象 | 规范 | 例 |
| --- | --- | --- |
| 模块 | `m_<插件>` | `m_ImageExport` |
| 窗体 | `frm<插件><用途>` | `frmImageExport` |
| Ribbon 回调 | `Ribbon_<插件>_<动作>` | `Ribbon_ImageExport_Export` |
| 全局变量 | 带插件前缀 | `gIE_Form` |

构建器会**扫描各文件的 Public 过程 / 常量 / 模块级变量，发现跨插件重名直接中止构建**，
不会等到 PowerPoint 里才报"编译错误"。

---

## 五、构建命令

依赖：Windows + PowerPoint 2016 及以上 + Python 3 + `pywin32`
（`pip install pywin32`，构建脚本通过 COM 驱动 PowerPoint 现场生成 VBA 工程）。

```bash
python build/build_suite.py                # 构建 + 安装 + 登记注册表
python build/build_suite.py --no-install   # 只构建
python build/build_suite.py --list         # 只看扫描到哪些插件（不做任何改动）
python build/build_suite.py --register-only# 不重建，只把 dist 产物重新安装登记
python build/build_suite.py --selftest     # 产出 PptEnhanceSuiteTest.ppam（含自检，不安装）
python build/build_suite.py --disable-legacy # 按 suite.json 的 legacy 清单清理旧加载项（先退出 PowerPoint）
```

> 构建前请**完全退出 PowerPoint**（脚本会通过 COM 拉起一个独立实例来生成
> VBA 工程，若检测到你正在使用会中止，避免打扰）。

## 六、回归测试

```bash
python build/build_suite.py --selftest     # 先产出自检版
python test/run_suite_test.py              # 加载加载项 → 真实导出 → 校验像素
```

测试会新建 3 页 16:9 演示文稿，以 4K / FHD-JPG / 300DPI 三种设置实际导出，
再逐字节解析 PNG / JPG 头，核对输出像素是否精确匹配。

同一轮还会另建一份 **5 页样张**，对「选页另存」做端到端验证：选中第 2/4/5 页
导出后解包核对页数，并断言打开临时副本的过程**没有新增任何窗口**
（`Application.Windows.Count` 不变）—— 这正是无窗口打开修复的判据。

第三份样张专供「图片替换」：目标图刻意带上裁剪 / 旋转 / 翻转 / 亮度 / 对比度 /
透明度 / 边框 / 超链接，并且**被另一张图压在下面**（用来验图层顺序还原）。
替换前后的 `.pptx` 会被直接解包，逐项比对 `spTree` 里那个 `<p:pic>` 节点的
**22 项 OOXML 属性** —— 这是独立于 VBA 自检的第二重证据，两者必须同时通过。

```bash
python build/build_suite.py --selftest
python test/run_suite_test.py            # 三个插件端到端
python test/verify_build_code.py         # 产物代码 vs 源码（--release 核对正式产物）
```

---

## 七、VBA 层面的硬约束（踩过的坑，改代码前必读）

这几条都是本项目实际撞出来的，且**报错信息极难定位**（加载项的 VBA 工程是
"隐藏工程"，编译失败只报一句 `隐藏的模块中的编译错误: <模块名>`，
既不指行也不指名字）：

| # | 约束 | 踩坑现象 |
| --- | --- | --- |
| 1 | **模块级声明（`Private`/`Public`/`Dim`/`Const`）必须全部位于任何过程之前** | 写在过程**之后**不报语法错，而是被**静默忽略** → 所有引用它的过程都变成「变量未定义」。这是本项目耗时最长的卡点 |
| 2 | **往已有模块末尾追加代码时，追加内容里不能有任何模块级声明** | 自检代码是追加的，一个 `Private Const` 落在过程后就重演上一条 |
| 3 | `Shape` 的翻转成员名是 **`HorizontalFlip` / `VerticalFlip`**（且**只读**） | 没有 `FlipHorizontal` / `FlipVertical`；改翻转只能用 `Flip(msoFlipHorizontal)` 方法，而它是**取反**语义，必须"先读当前值、不一致才翻转" |
| 4 | `Shape.Flip` 会**连带把 `Rotation` 取反** | 必须「先翻转、后设旋转」，否则旋转值被改掉 |
| 5 | `PictureFormat.ColorType` 赋值会**把亮度/对比度重置为默认** | 必须「先 ColorType、后 Brightness/Contrast」 |
| 6 | 代码里**不能出现 GBK 表示不了的字符**（`⚠`、emoji、部分符号） | VBE 保存工程时按本机代码页落盘，表示不了的字符被**静默换成 `?`**。构建期已加闸门（`check_encoding`）直接报错 |
| 7 | 源码里少用宿主内置函数做存在性判断 | 内置名保存后会被 VBE 规范化大小写。本项目统一改用 `Scripting.FileSystemObject`（后期绑定），彻底绕开 |
| 8 | `.ppam` **无法直存** | `SaveAs(..., ppSaveAsOpenXMLAddin)` 报 `(-2147352567, '发生意外。')`，必须回退存 `.pptm` 再改写内容类型 |
| 9 | `Presentations.Open` 的 `ReadOnly` 是**第 2 个**参数 | 写 `Open(path, 0)` 会把 `0` 落到 `ReadOnly` 上；无窗口打开要写 `Open(path, , , msoFalse)` |
| 10 | **不要对「从 COM 属性接回来的 Boolean」用 `Not`** —— 统一改成显式比较 `<var> = False` 或直接用 `Long` + `<> 0` | PowerPoint 以 `VARIANT_TRUE` 返回三态属性（如 `Selection.HasChildShapeRange`）时，本环境实测原始值是 **1**（不是 VBA 习惯的 -1），而 VBA 会把它**不做归一化**地塞进 `Boolean` 变量。此后 `Not 变量` 退化成按位取反：`Not 1 = -2` → **非 0 即真**，于是 `If Not hasChild Then` **永远成立**。详见下方「阶段 3.5」 |
| 11 | **`imageMso` 引用的 ID 必须逐个实测存在**，且存在 ≠ 有图标 | `ExportToFile`、`SaveAsPicture`、`ExportPicture` 这类"想当然"的 ID **都不存在**（图标静默不显示，不报错）。而 `CopyAsPicture`、`ExportMenu`、`Share` 虽能通过存在性探测，实际渲染出来是**橙色圆球占位图**。最终用 Excel 的 `CommandBars.GetImageMso` 初筛（错误码 `E_UNEXPECTED`=存在 / `E_INVALIDARG`=不存在；PowerPoint 的该方法不可用）+ `SavePicture` 导出图标**目视确认**三步闭环。本套件三个按钮最终定案：导出图片=`ExportToAnimatedGif`、选页另存=`FileSaveAs`、图片替换=`PictureInsertFromFile` |

> ⚠️ 第 10 条是本项目最隐蔽的一次踩坑：它**不报任何错**，只是让一个 `If Not ...`
> 分支永远走进去，把正确解析出来的结果静默覆盖掉。
> 现在 `IR_ResolveTarget` 里的 `hasChild` 刻意用 `Long` 承接并写 `<> 0` 判断，
> 相关处都留了说明注释 —— **不要改回 `Boolean` + `Not`**。

> 补充说明（曾经误判、已纠正）：`Dir$` **是**合法标识符，`X:\` 这类不存在路径
> 实测编译通过。**VBE 把标识符转成小写是无害的** —— 凡没被认作内置函数/属性的
> 标识符都会被转小写，`Dir(`→`dir(`、`Value`→`value` 全都正常。
> 早期把"小写化"当成"未声明"的证据，是走过的弯路。

---

## 八、验证记录

### 阶段 1 — 骨架 + 导出图片

| 项目 | 结果 |
| --- | --- |
| 插件扫描 | ✅ `_core` + `image-export` 两个 |
| 跨插件命名冲突检查 | ✅ 无重名 |
| 回调引用检查 | ✅ customUI 引用的回调在源码中均存在 |
| 包结构校验 | ✅ 内容类型 / customUI 部件 / 关系 / VBA 工程 / 内容一致 五项通过 |
| 注册表登记 | ✅ `HKCU\...\PowerPoint\AddIns\PptEnhanceSuite`（Path + AutoLoad=1，读回校验通过） |
| 加载项加载与执行 | ✅ 在 PowerPoint 中 `Loaded=True`，`Application.Run` 成功调用其宏 |
| DPI 换算 | ✅ 96→1280×720、150→2000×1125、300→4000×2250、600→8000×4500 |
| 范围解析 | ✅ `1-3→[1,2,3]`、`1,3→[1,3]`、`2-1→[1,2]`、`1-99→[1,2,3]`、`abc→[]` |
| 窗体事件绑定 | ✅ W_CHANGE / PRESET4K / FMT_JPG / RANGE_ALL / SPEC_ENABLED 全部 OK |
| 窗体尺寸 | ✅ InsideWidth 380.25×282.75，Outer 392.25×312 |
| **真实导出像素** | ✅ `T4K_001.png (3840,2160)`、`TFHD_002.jpg (1920,1080)`、`T300_003.png (4000,2250)` |

### 阶段 2 — 选页另存

| 项目 | 结果 |
| --- | --- |
| 插件扫描 | ✅ `_core` + `image-export` + `slide-export` 三个 |
| 跨插件命名冲突检查 | ✅ 无重名 |
| 包结构与回调校验 | ✅ 六项通过（含「包内 customUI 与源文件一致」） |
| **无窗口打开（bug 修复）** | ✅ `NOWIN=OK` —— 打开临时副本时 `Application.Windows.Count` 不变 |
| **按序号保留页面** | ✅ 5 页样张选第 2/4/5 页 → 导出 3 页，解包核对 `ppt/slides` 页面数一致 |
| 图片导出回归 | ✅ 原有用例全绿，未受本次改动影响 |

### 阶段 3 — 图片替换（含 bug 修复）

**原实现的缺陷**：老插件是「删除原图 → `AddPicture` 插一张新的」，因此只有
位置 / 尺寸 / 裁剪参数被手工抄回来，**其余全部丢失**，而且新图永远跑到最顶层。
现在改成 **「属性快照 → 删除 → 插入 → 逐项还原 → 恢复 Z 序」**。

| 项目 | 结果 |
| --- | --- |
| 插件扫描 | ✅ `_core` + `image-export` + `slide-export` + `image-replace` 四个 |
| 跨插件命名冲突检查 | ✅ 无重名 |
| 包结构与回调校验 | ✅ 六项通过 |
| 属性还原（VBA 自检，24 项） | ✅ 位置 / 尺寸 / 可见比例 / 裁剪分配 / 旋转 / 翻转 / 亮度 / 对比度 / 颜色类型 / 填充透明度 / 图层顺序 / 名称 / 替代文字 / 标题 / 超链接 / 动画 / 对象类型 **全部 OK** |
| **OOXML 逐项比对（独立证据）** | ✅ 22 项全部一致；目标图在 `spTree` 中序号 `3 → 3`（图层顺序保住） |
| **确认真的换了图** | ✅ 包内媒体 `1 → 2`（新图已入包），`spTree` 中 blip 引用随之改变 |
| 产物代码核对 | ✅ 6 个组件源码与包内代码逐行一致 |
| 图片导出 / 选页另存回归 | ✅ 前两个插件全绿，未受影响 |

> **裁剪为什么不能"照抄磅值"**：`PictureFormat.CropLeft` 虽以「磅」暴露，
> 但存进文件的其实是**比例**，分母是图片自身的原始尺寸。所以同一磅值对不同
> 像素尺寸的图折算出的比例完全不同 —— 实测 200px 的图（原始宽 150 磅）换成
> 300px 的图（原始宽 225 磅）后，12 磅从 8% 掉成 5.33%，文件里记录的
> `srcRect` 变了、可见矩形也跟着漂移。现在的做法是**保住比例**：先把源裁剪
> 磅值当种子套到新图上、实测它折算出的可见宽，再按
> `(1 - m_w/m_fullW) / (1 - vw/fw)` 把种子缩放到目标可见宽。
> 全程序靠实测，不假设磅值与原始尺寸之间的换算关系。
>
> 由此也修正了自检的**判据口径**：早期版本直接比裁剪磅值，会稳定误报 DIFF
> （因为磅值本来就该变）。现在改为比对「可见比例」与「四边裁剪分配比例」
> 这两个与实现无关、且用户真正在意的量，绝对磅值交给 OOXML 解包核对。

---

### 阶段 3.5 — 图片替换支持「编组（组合）内的图片」

**用户报的现象**：替换功能基本正常，只有当要替换的图片在编组内时无法识别，
弹窗提示 `选中的对象不是图片（当前类型 6）`。用户的选取方式确认为
**「手动点选组内指定的那张图片」**。

**真实根因（两个，叠在一起）**：

1. **`Shape.Parent` 对组内形状返回的是幻灯片，不是组合。**
   用它判断"这张图片属于哪个组合"永远得到"否"，组内图片于是被当成
   幻灯片上的普通图片处理。正确 API 是 **`Shape.ParentGroup`**
   （组内返回所属组合、顶层直接抛错，正好当判据）。
2. **`Not` 对「从 COM 接回来的 Boolean」失效**（见第七节第 10 条）。
   解析函数里 `If Not hasChild Then` 永远成立，于是刚刚从
   `ChildShapeRange` 正确取出的"组内那张图片"，立刻被
   `ShapeRange(1)`（＝整个组合）覆盖 —— 最终就是用户看到的"类型 6"。

**组内替换的实现路线**（三条候选路实测后的唯一可行解）：

| 候选 | 结论 |
| --- | --- |
| `GroupShapes.AddPicture` 在组内原地插图 | ❌ 该成员**不存在**（查 IDispatch 名称表确认） |
| `Shape.Fill.UserPicture` 就地换图 | ❌ 语义破坏：用后 `CropLeft` 读回 `-2147483648`，`Fill.Type` 变成图片填充 |
| **解组 → 替换 → 重组** | ✅ 唯一可行，且做到**完全无损** |

**关键技术点 —— 「变换归零法」**：`Shape.Ungroup` 会把组合的旋转/翻转
**烘焙进子形状**（组旋转 20° → 子形状 `Rotation` 也变 20°）；且
`Ungroup` **只解一层**（嵌套子组仍是 Type=6），`ShapeRange.Group()` 重组时
**成员 Z 序不按传入顺序**。因此流程是：

```
快照组合链（名字/位置/旋转翻转/Z 序/成员）
  → 每层变换临时归零（先 Flip 后 Rotation）
  → 逐层解组
  → 复用普通替换逻辑「属性快照 → 删除 → 插入 → 逐项还原」
  → 由内向外逐层重组 + 回放快照
  → 还原临时名
```

> 解组前必须先把组的变换归零，否则旋转/翻转会被"烤"进成员；实测
> 四场景（无变换 / 旋转 / 翻转 / 旋转+翻转）全部**完全无损**。

| 项目 | 结果 |
| --- | --- |
| 编组样张构造 | ✅ 图片在带**旋转 + 翻转**的组合内（`GRP_TARGET` / 组内图片 `TGT_IMAGE` / 干扰矩形 `DECO_RECT`），组外另有顶层覆盖图 |
| 组合结构还原（VBA 自检 14 项） | ✅ 组名 / 类型 / 位置 / 尺寸 / 旋转 / 翻转H / 翻转V / Z 序 / 成员数 / 成员构成 全部 OK |
| 组内图片属性还原（VBA 自检 16 项） | ✅ 位置 / 尺寸 / 旋转 / 翻转 / Z 序 / 亮度 / 对比度 / 颜色类型 / 填充透明度 / 替代文字 / 标题 / 动画 / 超链接 全部 OK（裁剪按下方口径比对） |
| **OOXML 逐项比对（独立证据）** | ✅ 33 项全部一致；组合在 `spTree` 中序号 `3 → 3`、旋转 `600000 → 600000`、翻转 `1 → 1` |
| **确认真的换了图** | ✅ 包内媒体 `1 → 2`（组内图片替换后占用的仍是同一个 rId 槽位，`rId3 → rId3`，所以判据必须看媒体数而不是 rId） |
| **选区解析 · 单击整个组合** | ✅ 解析为组内唯一图片 `TGT_IMAGE` |
| **选区解析 · 进入组合点选图片** | ✅ 解析为 `TGT_IMAGE`（这正是用户的实际操作方式） |
| **选区解析 · 进入组合点选非图片** | ✅ 正确**拒绝**（早前版本会错解析成图片，回归脚本当时只打印不判定，导致缺陷漏过 —— 现已改为硬断言） |
| **选区解析 · 选中含两张图片的组合** | ✅ 正确拒绝，并提示「双击进入组合再点选」 |
| 产物代码核对 | ✅ 5 个组件源码与包内代码逐行一致 |
| **编译体检的覆盖缺口（顺手修掉）** | ✅ 发现预言机原先**漏检**组内代码（VBA 是"过程首次进入才编译"，而原调用集会提前返回）→ 已改为把各过程逐条直接进入；注入 `tmp = 1 +` 复现验证：修前报"通过"、修后报"语法错误" |
| 真机加载与执行 | ✅ `Loaded=True`，`Application.Run` 实际调用 `m_ImageReplace.ImageReplace_Describe()` 返回 `(Nothing)` —— 证明该模块**被加载且能编译、能执行**，而不是只躺在包里 |
| 图片导出 / 选页另存 / 普通图片替换回归 | ✅ 全部全绿，未受影响 |

> **「点选矩形」这条用例为什么必须硬断言**：它区分的是"选区解析真的在看
> 子形状"还是"其实一直拿到整个组合"。两条路径在**样式上**都会最终命中组内
> 唯一的那张图片，所以只有"点选非图片必须被拒绝"才能把它们分开。
> 早期脚本对它只打印不判定，是本阶段缺陷能一路通过回归的直接原因。

#### 已知边界：嵌套组合（组里还有组）—— 实现了，但未能验证

代码里那条「祖先组合链」（`IR_BuildChain` / `m_chain`）本来就是为嵌套组合写的
（`Ungroup` 只解一层，所以按链逐层下钻、再逐层重组）。**但实测这套环境下的
PowerPoint 造不出嵌套组合**：

| 尝试方式 | 结果 |
| --- | --- |
| `Shapes.Range(...).Group()` | 展平（内层组合消失，成员被提升到外层） |
| 选中后 `CommandBars.ExecuteMso("ObjectsGroup")`（模拟 UI 分组） | 展平 |
| 存盘 → 重新打开 → 再组合 | 仍然展平 |

三条都试过、结果一致。既然**造不出用例**，这条分支就无法被回归覆盖 ——
所以它只作为防御性实现保留，**不宣称已验证**。

> 推论：若真的遇到嵌套组合（例如第三方工具生成的 `.pptx`，OOXML 本身支持
> `p:grpSp` 套 `p:grpSp`），重组阶段很可能把**内层组合摊平**。
> 图片自身的属性（位置/尺寸/裁剪/旋转/亮度/透明度/超链接/动画）仍会被完整
> 还原，但"内层组合"这个容器可能不再存在。目前按"尽力而为"处理，未做拦截。

### 阶段 3.6 — 「导出图片」按钮补图标

用户验收反馈：三个分组按钮里「选页另存」「图片替换」有图标，「导出图片」没有。
包内 `customUI.xml` 三个按钮的 `imageMso` 结构完全对称，唯一差别是 ID ——
即 **`ExportToFile` 这个 ID 在 Office 图标库里不存在**，图标静默不显示、不报错。

| 步骤 | 结果 |
| --- | --- |
| 存在性初筛 | PowerPoint 的 `CommandBars.GetImageMso` **不可用**（对所有 ID 都报错）；改用 **Excel** 的同名方法：错误码 `E_UNEXPECTED`=ID 存在、`E_INVALIDARG`=不存在。先校准（`FileSaveAs`/`PictureInsertFromFile` 判"存在"、乱造 ID 判"不存在"），再筛 155 个候选，26 个存在 |
| 图标目视确认 | 用 Excel VBA 的 `SavePicture Application.CommandBars.GetImageMso(id,32,32), path` 把候选图标**逐个导出成图片拼图查看** —— 发现 3 个假阳性（`CopyAsPicture`/`ExportMenu`/`Share` 渲染为**橙色圆球占位图**），证明"存在"必须再目视确认 |
| 选定 | `ExportToAnimatedGif`（文档 + GIF 标签）：PowerPoint 原生「导出图像文件」家族图标，与「选页另存」的文档+软盘结构一致；备选 `CopyPicture`（双图片）因与「图片替换」的图片主题重复而放弃 |
| 视觉终验 | 构建 → 加载 → 模拟点击激活「增强插件」标签 → `PrintWindow` 截图：**三个按钮图标全部显示**（导出为图片=文档+GIF、导出选中幻灯片=软盘+铅笔、原位替换图片=图片+显示器） |
| 回归 | ✅ 产物代码核对通过、端到端回归总体通过（33 项 OOXML 比对一致） |

### 阶段 4 — 卸载被取代的旧加载项

**清理前的实际状况**（侦查结论）：

| 注册表项 | 文件 | 功能区影响 |
| --- | --- | --- |
| `PptImageExport` | `%APPDATA%\Microsoft\AddIns\PptImageExport.ppam` | 多出一个重复的「图片导出」标签 |
| `PPT原位替换插件V3` | 自选目录下的 `PPT原位替换插件V3.ppam` | **无**（包内没有 customUI 部件，不产生标签，但每次启动都白白加载） |
| `选取另存` | （从未登记过启动项，仅桌面留有文件副本） | 无 |

**实现方式**：清理对象写成**声明式清单**（`suite.json` 的 `legacy` 数组），
`build_suite.py --disable-legacy` 按清单执行。刻意「只认清单、不做扫描」——
绝不去遍历注册表或 AddIns 目录猜哪些是旧的，避免误伤 iSlide 这类无关的
第三方加载项。作用域限死两处：`HKCU\...\PowerPoint\AddIns\<清单名>` 这一个
注册表键、`%APPDATA%\Microsoft\AddIns\<清单文件名>` 这一个文件；
文件找不到、登记不存在都会明确打印"跳过"。执行前要求完全退出 PowerPoint
（文件被占用删不掉）。

**执行结果**：注销注册表 2 项、删除文件 1 个；桌面 `PPT增强插件\` 与
F 盘 `Office图像原位替换工具\` 的原始文件**保持不动**（作为天然兜底，
由用户自行决定去留）。

**清理后验证**（全部通过）：

| 项目 | 结果 |
| --- | --- |
| 注册表 | ✅ `PowerPoint\AddIns` 下只剩 `PptEnhanceSuite` 一项（AutoLoad=1） |
| AddIns 目录 | ✅ 只剩 `PptEnhanceSuite.ppam` |
| 编译体检 / 自检构建 / 端到端回归 | ✅ 全绿（含编组内替换 33 项 OOXML 比对） |
| 产物代码核对 + 真机加载 | ✅ 5 个组件逐行一致；`loaded=True`，三个模块回调可执行 |
| 功能区截图 | ✅ 只剩「增强插件」一个自定义标签，「图片导出」重复入口消失，iSlide 未受影响，三个分组按钮图标齐全（见 `test/output/_final_ribbon.png`） |

---

## 九、卸载

### 卸载本套件

```text
1. PowerPoint → 文件 → 选项 → 加载项 → 转到（管理：PowerPoint 加载项）→ 取消勾选
2. 删除 %APPDATA%\Microsoft\AddIns\PptEnhanceSuite.ppam
3. 如需彻底清理注册表登记：
   HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PptEnhanceSuite
```

### 清理被本套件取代的旧加载项

```bash
python build/build_suite.py --disable-legacy
```

按 `suite.json` 里 `legacy` 数组的**声明式清单**执行（先退出 PowerPoint）：
逐条删除 `HKCU\...\PowerPoint\AddIns\<registry>` 登记项、删除
`%APPDATA%\Microsoft\AddIns\<files>` 旧文件；不存在的条目打印"跳过"。
**只认清单、不做扫描** —— 不会碰 iSlide 等第三方加载项；文件查找严格限定在
AddIns 目录内按名精确匹配，不递归、不通配。

阶段 4 已执行过一次：清理了 `PptImageExport`（重复的「图片导出」标签）与
`PPT原位替换插件V3`（无标签但随启动加载）两条登记。桌面 `PPT增强插件\` 与
原始备份目录下的旧文件未动，由用户自行决定去留。
将来若本套件被其它方案取代，也可以把自己写进 `legacy` 清单后用同一命令清理。

---

## 十、许可证

本项目以 [MIT License](LICENSE) 开源，可自由使用、修改、分发与商用，仅需保留版权声明。

第三方加载项的名称与路径仅作为「被本套件取代的旧加载项」在文档中提及，
其代码与产物不在本仓库内（见 `.gitignore` 中的 `_extracted/`）。
