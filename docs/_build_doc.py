# -*- coding: utf-8 -*-
"""生成《PPT 增强插件套件 · 使用说明》（自包含 HTML，截图 base64 内嵌）。

一次运行产出三份（内容由同一份 HTML 模板派生，不会漂移）：
    使用说明.html                          自包含单文件，可直接发给人
    PPT 增强插件套件 · 使用说明.pdf        Chrome/Edge 无头打印的 A4 版
    使用说明.md                            Markdown 版，配图走相对路径 screenshots/

用法：  python _build_doc.py            # HTML + PDF + MD
        python _build_md.py             # 只重新生成 Markdown
"""
import base64
import io
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
S = os.path.join(HERE, "screenshots")
OUT = os.path.join(HERE, "使用说明.html")
OUT_PDF = os.path.join(HERE, "PPT 增强插件套件 · 使用说明.pdf")

# 图片显示宽度上限：与正文栏宽（.page 的 960px）对齐，避免大截图铺满溢出。
DEFAULT_IMG_W = 960


def img(name, max_w=DEFAULT_IMG_W):
    """内嵌截图为 data URI。max_w 为显示宽度上限（px），实际再受容器宽度约束。"""
    p = os.path.join(S, name)
    b64 = base64.b64encode(io.open(p, "rb").read()).decode("ascii")
    style = ' style="max-width:min(%dpx,100%%)"' % max_w
    return '<img%s src="data:image/png;base64,%s" alt="%s">' % (style, b64, name)


HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PPT 增强插件套件 · 使用说明</title>
<style>
  :root{
    --ink:#2b3340; --soft:#6b7480; --accent:#c05a2e; --paper:#f7f5f1;
    --line:#e5e0d8; --card:#ffffff; --note-bg:#fdf6ee; --warn-bg:#fdf1f0;
  }
  *{box-sizing:border-box}
  body{
    margin:0; padding:0 20px 80px; background:var(--paper); color:var(--ink);
    font-family:"Microsoft YaHei UI","PingFang SC","Segoe UI",sans-serif;
    font-size:15px; line-height:1.75;
    -webkit-print-color-adjust:exact; print-color-adjust:exact;
  }
  .page{max-width:960px; margin:0 auto}
  /* 截图统一约束：不超过栏宽，始终居中，打印时不被切断 */
  .page img{
    display:block; margin:16px auto; max-width:100%; height:auto;
    border:1px solid var(--line); border-radius:6px; break-inside:avoid;
  }
  header{
    padding:42px 0 22px; border-bottom:3px solid var(--accent); margin-bottom:28px;
  }
  header h1{margin:0 0 8px; font-size:30px; letter-spacing:.5px}
  header .sub{color:var(--soft); font-size:14px}
  .ver{
    display:inline-block; background:var(--accent); color:#fff; border-radius:4px;
    padding:1px 10px; font-size:13px; margin-left:8px; vertical-align:3px;
  }
  nav{
    background:var(--card); border:1px solid var(--line); border-radius:10px;
    padding:16px 22px; margin:0 0 34px;
  }
  nav b{font-size:14px; color:var(--soft); letter-spacing:2px}
  nav ol{margin:8px 0 0; padding-left:22px}
  nav a{color:var(--accent); text-decoration:none}
  nav a:hover{text-decoration:underline}
  h2{
    font-size:22px; margin:52px 0 6px; padding-left:12px;
    border-left:5px solid var(--accent); line-height:1.4;
  }
  h2 .en{font-size:13px; color:var(--soft); font-weight:normal; margin-left:10px}
  h3{font-size:17px; margin:30px 0 6px}
  p{margin:10px 0}
  .lead{color:var(--soft)}
  figure{
    margin:18px 0 22px; background:var(--card); border:1px solid var(--line);
    border-radius:10px; padding:14px 14px 10px; text-align:center;
  }
  figure img{max-width:100%; height:auto; border-radius:4px; border:1px solid var(--line)}
  figure.narrow img{max-width:560px}
  figcaption{
    color:var(--soft); font-size:13px; padding:8px 4px 2px; text-align:center;
  }
  table{border-collapse:collapse; width:100%; margin:14px 0 22px; background:var(--card); font-size:14px}
  th,td{border:1px solid var(--line); padding:8px 12px; text-align:left; vertical-align:top}
  th{background:#f0ece4; white-space:nowrap}
  td.c{white-space:nowrap}
  .box{
    border-radius:8px; padding:12px 16px; margin:14px 0; font-size:14px;
    border:1px solid var(--line);
  }
  .box.tip{background:var(--note-bg); border-color:#eeddc4}
  .box.warn{background:var(--warn-bg); border-color:#f0cfc9}
  .box b:first-child{color:var(--accent)}
  ol.steps{padding-left:0; list-style:none; counter-reset:step; margin:14px 0}
  ol.steps li{
    counter-increment:step; position:relative; padding:6px 0 6px 44px; margin:4px 0;
  }
  ol.steps li::before{
    content:counter(step); position:absolute; left:0; top:7px;
    width:28px; height:28px; border-radius:50%; background:var(--accent); color:#fff;
    text-align:center; line-height:28px; font-size:14px; font-weight:600;
  }
  code{
    background:#efeae2; border-radius:4px; padding:1px 6px;
    font-family:Consolas,Menlo,monospace; font-size:13px;
  }
  .pill{
    display:inline-block; background:#efeae2; color:var(--ink); border-radius:20px;
    padding:0 10px; font-size:12.5px; margin:1px 4px 1px 0;
  }
  footer{margin-top:60px; color:var(--soft); font-size:13px; border-top:1px solid var(--line); padding-top:16px}
  @page{size:A4; margin:13mm 11mm}
  @media print{
    body{background:#fff; padding:0; font-size:13.5px}
    .page{max-width:none}
    figure,nav,table,.box,img,footer{break-inside:avoid}
    h2,h3{break-after:avoid}
    p,li{orphans:2; widows:2}
    a{color:var(--accent); text-decoration:none}
  }
</style>
</head>
<body>
<div class="page">

<header>
  <h1>PPT 增强插件套件 · 使用说明<span class="ver">v1.0.0</span></h1>
  <div class="sub">适用于 Windows 版 PowerPoint 2016 / 2019 / 2021 / Microsoft 365 · 更新日期 2026-10-10</div>
</header>

<nav>
  <b>目 录</b>
  <ol>
    <li><a href="#intro">套件简介</a></li>
    <li><a href="#install">安装与卸载</a></li>
    <li><a href="#export">插件一：导出为图片</a></li>
    <li><a href="#slides">插件二：导出选中幻灯片（选页另存）</a></li>
    <li><a href="#replace">插件三：原位替换图片</a></li>
    <li><a href="#faq">常见问题</a></li>
  </ol>
</nav>

<!-- ==================================================== 一、简介 -->
<h2 id="intro">一、套件简介</h2>
<p>「增强插件」是一套 PowerPoint 功能区加载项，把三个日常高频操作集成在同一张功能区内，
安装后随 PowerPoint 自动加载，无需手动开启。</p>
<table>
  <tr><th class="c">按钮</th><th class="c">所在分组</th><th>一句话说明</th></tr>
  <tr>
    <td class="c"><b>导出为图片</b></td><td class="c">导出图片</td>
    <td>按自定义分辨率（最高 8K / 印刷级 600 DPI）把幻灯片导出为 PNG / JPG / TIFF / BMP，
        绕过 PowerPoint 自带「另存为」固定 96 DPI 的限制。</td>
  </tr>
  <tr>
    <td class="c"><b>导出选中幻灯片</b></td><td class="c">选页另存</td>
    <td>在缩略图窗格勾选若干页，一键另存为一份新的 .pptx，
        版式、母版、主题、备注、动画、切换 100% 保留。</td>
  </tr>
  <tr>
    <td class="c"><b>原位替换图片</b></td><td class="c">图片替换</td>
    <td>选一张图片换掉，位置、尺寸、裁剪、旋转、图层、动画等所有版式属性原样保留，
        新图不会「跳」到最顶层。</td>
  </tr>
</table>

<p>安装完成后，功能区会出现 <b>「增强插件」</b> 标签（位于「插入」之后）：</p>
{{img:ribbon_full.png}}

<h3>三个按钮一览</h3>
{{img:ribbon_buttons.png}}

<!-- ==================================================== 二、安装 -->
<h2 id="install">二、安装与卸载</h2>

<h3>2.1 环境要求</h3>
<table>
  <tr><th class="c">项目</th><th>要求</th></tr>
  <tr><td class="c">操作系统</td><td>Windows 10 / 11</td></tr>
  <tr><td class="c">Office</td><td>PowerPoint 2016 及以上（32 / 64 位均可）</td></tr>
  <tr><td class="c">文件</td><td><code>PptEnhanceSuite.ppam</code>（约 124 KB，单文件即插即用）</td></tr>
</table>

<h3>2.2 安装步骤</h3>
<p class="lead">如果插件已经由安装脚本装好（文件位于
<code>%APPDATA%\Microsoft\AddIns\PptEnhanceSuite.ppam</code> 并已登记随启动自动加载），
可以跳过本节，直接看 <a href="#verify">2.3 验证安装</a>。</p>
<ol class="steps">
  <li>关闭所有 PowerPoint 窗口。</li>
  <li>把 <code>PptEnhanceSuite.ppam</code> 复制到一个<b>固定位置</b>（推荐
      <code>%APPDATA%\Microsoft\AddIns\</code>，即「文档」地址栏直接输入后回车打开的
      <code>C:\Users\你的用户名\AppData\Roaming\Microsoft\AddIns</code>）。</li>
  <li>打开 PowerPoint，依次点击 <b>文件 → 选项 → 加载项</b>，在页面底部把「管理」下拉框切换为
      <b>「PowerPoint 加载项」</b>，点「转到」。</li>
  <li>在弹出的加载项窗口中点 <b>「添加」</b>，选中刚才的 <code>PptEnhanceSuite.ppam</code>，
      确认它在列表中处于<b>勾选</b>状态，关闭窗口。</li>
  <li>重启 PowerPoint。</li>
</ol>

<div class="box warn">
  <b>宏安全提示：</b>本插件未做数字签名。若功能区没有出现「增强插件」标签，请检查：
  <b>文件 → 选项 → 信任中心 → 信任中心设置 → 加载项</b>，确认没有勾选「要求应用程序扩展受到受信任发布者签署」；
  并在 <b>信任中心 → 宏设置</b> 中选择「禁用无数字签署的所有宏（通知）」或更宽松的选项，
  然后重新加载插件。若曾弹过安全提示且选择了「禁用」，可在
  <b>文件 → 信息 → 启用内容</b> 里重新启用。
</div>

<h3 id="verify">2.3 验证安装</h3>
<p>打开任意演示文稿，功能区内出现 <b>「增强插件」</b> 标签、且能看到下图中三个带图标的按钮，即安装成功。
将鼠标悬停在按钮上会显示功能提示（screentip / supertip）。</p>

<h3>2.4 卸载</h3>
<ol class="steps">
  <li><b>文件 → 选项 → 加载项 → 管理：PowerPoint 加载项 → 转到</b>，取消勾选
      <code>PptEnhanceSuite</code>（取消勾选仅本次不加载；点「删除」则彻底移出列表）。</li>
  <li>删除文件 <code>%APPDATA%\Microsoft\AddIns\PptEnhanceSuite.ppam</code>。</li>
  <li>（可选）彻底清理登记项：删除注册表键
      <code>HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PptEnhanceSuite</code>。</li>
</ol>

<!-- ==================================================== 三、导出为图片 -->
<h2 id="export">三、插件一：导出为图片<span class="en">导出图片分组</span></h2>

<h3>3.1 解决什么问题</h3>
<p>PowerPoint 自带的「另存为图片」固定输出 <b>96 DPI</b>，投屏发虚、印刷不可用。
本插件按<b>像素 / DPI</b> 精确控制输出分辨率，并提供常用档位一键选择。</p>

<h3>3.2 操作步骤</h3>
<ol class="steps">
  <li>打开演示文稿，点击功能区 <b>「增强插件 → 导出为图片」</b>，弹出设置面板：</li>
</ol>
{{img:export_form.png}}
<ol class="steps" start="2">
  <li>选择<b>导出范围</b>、<b>尺寸与分辨率</b>、<b>图片格式</b>与<b>输出位置</b>（各项见 3.3）。</li>
  <li>点击 <b>「开始导出」</b>。导出过程中面板底部实时显示进度（第几页 / 共几页），
      可随时点「中止」停止。</li>
  <li>完成后按状态提示打开输出目录查看图片。</li>
</ol>

<h3>3.3 参数说明</h3>
<table>
  <tr><th class="c">参数</th><th>说明</th></tr>
  <tr>
    <td class="c"><b>导出范围</b></td>
    <td><span class="pill">当前页</span><span class="pill">全部页</span><span class="pill">选中页</span><span class="pill">页码范围</span><br>
      「选中页」对应左侧缩略图窗格中选中的页；「页码范围」支持混合写法，
      如 <code>1,3,5-8</code>（逗号分隔、连字符表区间；中文逗号、<code>～</code>、<code>—</code> 等全角符号均可识别）。</td>
  </tr>
  <tr>
    <td class="c"><b>预设档位</b></td>
    <td><span class="pill">幻灯片原始尺寸（96 DPI）</span><span class="pill">FHD 1080P</span><span class="pill">2K</span><span class="pill">4K</span><span class="pill">5K</span><span class="pill">8K</span><span class="pill">印刷级 150 / 300 / 600 DPI</span><span class="pill">自定义</span><br>
      档位按「长边像素」换算；印刷档按 DPI 换算。选择「自定义」后可手动填写。</td>
  </tr>
  <tr>
    <td class="c"><b>DPI / 宽 × 高</b></td>
    <td>三者联动换算：「宽」「高」任一改动即同步 DPI（勾选「锁定宽高比」时保持幻灯片比例）。
      面板中部会实时预览换算结果，例如 <code>输出 3840 × 2160 px · 比例 16:9 · 等效 288 DPI · PNG · 共 5 页</code>。</td>
  </tr>
  <tr>
    <td class="c"><b>图片格式</b></td>
    <td><span class="pill">PNG（无损，推荐）</span><span class="pill">JPG（体积小）</span><span class="pill">TIFF（印刷）</span><span class="pill">BMP</span></td>
  </tr>
  <tr>
    <td class="c"><b>输出目录</b></td>
    <td>默认为演示文稿所在目录下的 <code>导出图片</code> 子文件夹（文稿未保存过时为「图片」文件夹），
      可点「浏览…」修改。</td>
  </tr>
  <tr>
    <td class="c"><b>文件名前缀</b></td>
    <td>输出按 <code>前缀_页码序号.扩展名</code> 命名，如前缀 <code>Slide</code> 则为
      <code>Slide_001.png</code>、<code>Slide_002.png</code>…（三位编号）。</td>
  </tr>
</table>

<div class="box tip">
  <b>小提示：</b>输出目录、格式、前缀、上次使用的尺寸等设置都会<b>自动记忆</b>，下次打开面板直接就是上次的状态；
  想恢复默认可在注册表删除
  <code>HKEY_CURRENT_USER\Software\VB and VBA Program Settings\PptImageExport</code> 后重新打开面板。
</div>

<h3>3.4 导出效果参考</h3>
<p>同一页分别在 FHD 与 4K 档位下导出，放大后对比 —— 文字与线条边缘的差距非常明显：</p>
{{img:export_fhd_vs_4k.png}}

<!-- ==================================================== 四、选页另存 -->
<h2 id="slides">四、插件二：导出选中幻灯片<span class="en">选页另存分组</span></h2>

<h3>4.1 解决什么问题</h3>
<p>评审只发部分页、从大合集里抽出某几页给同事 —— 复制粘贴到新文件容易丢母版和动画。
本插件直接生成一份「只含所选页」的新 .pptx：<b>整本另存为副本后删除未选页</b>，
因此版式、母版、页面尺寸、主题、备注、动画与切换效果 <b>100% 保留</b>。</p>

<h3>4.2 操作步骤</h3>
<ol class="steps">
  <li>在左侧<b>缩略图窗格</b>中选中要导出的页：单击选一页；按住 <code>Ctrl</code> 单击可多选不连续的页；
      按住 <code>Shift</code> 单击可选连续区间。下例中选中了第 2、4 页：</li>
</ol>
{{img:slidesel.png}}
<ol class="steps" start="2">
  <li>点击功能区 <b>「增强插件 → 导出选中幻灯片」</b>，弹出「另存为」对话框，
      文件名默认为 <code>原文件名_selected.pptx</code>：</li>
</ol>
{{img:slidesave.png}}
<ol class="steps" start="3">
  <li>确认保存位置与文件名，点「保存」。完成后弹窗显示导出页数与文件路径。</li>
</ol>

<div class="box tip">
  <b>小提示：</b>导出的是一份<b>新文件</b>，当前演示文稿不受任何影响；
  新文件中的页码会重新从 1 开始编排。
</div>

<!-- ==================================================== 五、图片替换 -->
<h2 id="replace">五、插件三：原位替换图片<span class="en">图片替换分组</span></h2>

<h3>5.1 解决什么问题</h3>
<p>客户临时换图，但图片上的裁剪、亮度、动画都调好了 —— 直接删除重插等于重做一遍。
本插件把新图<b>原位换入</b>，以下属性全部保留：</p>
<p>
  <span class="pill">位置与尺寸</span><span class="pill">裁剪</span><span class="pill">旋转与翻转</span><span class="pill">亮度 / 对比度</span><span class="pill">透明度</span><span class="pill">阴影 / 边框 / 发光 / 艺术效果</span><span class="pill">图层顺序（不会跳到最顶层）</span><span class="pill">名称 / 替代文字 / 超链接</span><span class="pill">动画</span>
</p>

<h3>5.2 支持的三种选取方式</h3>
<table>
  <tr><th class="c">方式</th><th>操作</th></tr>
  <tr><td class="c">① 直接选中</td><td>在幻灯片上单击选中那张图片。</td></tr>
  <tr><td class="c">② 选中组合</td><td>选中一个<b>内含唯一图片</b>的组合，插件自动替换组合里那张图。</td></tr>
  <tr><td class="c">③ 进入组合点选</td><td><b>双击进入组合</b>后，单击组内要替换的那张图片（组合里有多张图片时必须用这种方式）。</td></tr>
</table>

<h3>5.3 操作步骤</h3>
<ol class="steps">
  <li>按上述任一方式选中图片（下例：幻灯片上的图片处于选中状态，四周出现控制点）：</li>
</ol>
{{img:selimg.png}}
<ol class="steps" start="2">
  <li>点击功能区 <b>「增强插件 → 原位替换图片」</b>，在弹出的文件对话框中选择新图片
      （支持 JPG / PNG / BMP / TIF / GIF）：</li>
</ol>
{{img:replace_pick.png}}
<ol class="steps" start="3">
  <li>点「打开」执行替换，完成后弹窗确认：</li>
</ol>
{{img:replace_done.png}}
<p>替换前后对比 —— 版面完全没动，只有图片内容变了：</p>
{{img:replace_before_after.png}}

<div class="box warn">
  <b>已知边界（多层编组）：</b>「组中还有组」的多层嵌套组合场景尚未充分验证。
  若替换多层编组内的图片出现异常，建议先<b>双击进入最内层组合后点选图片</b>（方式③），
  或临时解组后再替换。此问题已在后续完善计划中。
</div>

<!-- ==================================================== 六、FAQ -->
<h2 id="faq">六、常见问题</h2>

<h3>Q1：功能区里没有「增强插件」标签？</h3>
<p>按顺序检查：① 是否重启过 PowerPoint；② <b>文件 → 选项 → 加载项 → 管理：PowerPoint 加载项 → 转到</b>，
确认 <code>PptEnhanceSuite</code> 在列表中且已勾选；③ 参见 2.2 节的宏安全提示；
④ 若曾被「禁用文件」拦截，到 <b>文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录</b>
检查，或将 ppam 所在目录加入受信任位置。</p>

<h3>Q2：点「导出为图片」提示「请先打开一个演示文稿」？</h3>
<p>插件需要至少打开一份演示文稿才能计算页面尺寸。先打开（或新建）一份文稿再点按钮。</p>

<h3>Q3：导出的图片去哪里找？</h3>
<p>设置面板的「输出目录」栏显示了本次输出位置，默认在演示文稿同目录的 <code>导出图片</code> 子文件夹。
导出完成提示中也会显示文件所在路径。</p>

<h3>Q4：替换时提示「所选组合里有 N 张图片，无法确定要替换哪一张」？</h3>
<p>这是刻意的保护 —— 组合里有多张图片时插件不会猜。按 <b>5.2 节方式③</b>：
双击进入组合，单击要替换的那张图，再点「原位替换图片」。</p>

<h3>Q5：替换后图片跑到最顶层了？</h3>
<p>正常情况不会 —— 图层顺序属于被保留的属性。若遇到，请确认是通过本插件替换
（选中图片 → 原位替换图片），而不是删除后重新插入。</p>

<h3>Q6：选中了文本框 / 形状点「原位替换图片」没反应？</h3>
<p>插件只处理图片对象，选中非图片时会弹窗说明，这是预期行为。</p>

<h3>Q7：多台电脑使用，怎么分发？</h3>
<p>把 <code>PptEnhanceSuite.ppam</code> 拷贝到目标电脑，按 2.2 节步骤安装即可；
插件为单文件、免注册组件、无其他依赖。</p>

<footer>
  PPT 增强插件套件 v1.0.0 · 使用说明 · 2026-10-10<br>
  安装位置：<code>%APPDATA%\Microsoft\AddIns\PptEnhanceSuite.ppam</code> ·
  随 PowerPoint 启动自动加载
</footer>

</div>
</body>
</html>
"""

html = (HTML
        # 宽幅整屏截图（1920 级）：显式约束到正文栏宽 960px，避免撑破版面
        .replace("{{img:ribbon_full.png}}", img("ribbon_full.png", 960))
        .replace("{{img:slidesel.png}}", img("slidesel.png", 960))
        .replace("{{img:selimg.png}}", img("selimg.png", 960))
        .replace("{{img:replace_before_after.png}}", img("replace_before_after.png", 960))
        # 其余截图原始尺寸均在 960 以内，沿用默认上限；面板类略收窄更协调
        .replace("{{img:ribbon_buttons.png}}", img("ribbon_buttons.png"))
        .replace("{{img:export_form.png}}", img("export_form.png"))
        .replace("{{img:export_fhd_vs_4k.png}}", img("export_fhd_vs_4k.png"))
        .replace("{{img:slidesave.png}}", img("slidesave.png", 880))
        .replace("{{img:replace_pick.png}}", img("replace_pick.png", 880))
        .replace("{{img:replace_done.png}}", img("replace_done.png")))

io.open(OUT, "w", encoding="utf-8", newline="\n").write(html)
print("[html] written: %s  (%.1f KB)" % (OUT, os.path.getsize(OUT) / 1024))


# ------------------------------------------------------------------ PDF 导出
BROWSERS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]


def build_pdf():
    """调用本机 Chrome / Edge 无头打印 A4 PDF（版式由 @media print / @page 控制）。

    使用独立的临时 user-data-dir，不会读写用户浏览器配置，也不会弹窗。
    """
    exe = next((p for p in BROWSERS if os.path.isfile(p)), None)
    if exe is None:
        print("[pdf] 未找到 Chrome / Edge，跳过 PDF 生成")
        return False
    url = pathlib.Path(OUT).as_uri()
    tmp = tempfile.mkdtemp(prefix="builddoc_")
    try:
        for headless in ("--headless", "--headless=new"):
            try:
                os.remove(OUT_PDF)          # 避免旧文件被占用时误判成功
            except OSError:
                pass
            args = [
                exe, headless, "--disable-gpu", "--no-first-run",
                "--no-default-browser-check", "--disable-extensions",
                "--user-data-dir=" + tmp,
                "--no-pdf-header-footer",
                "--run-all-compositor-stages-before-draw",
                "--virtual-time-budget=30000",
                "--print-to-pdf=" + OUT_PDF,
                url,
            ]
            try:
                r = subprocess.run(args, stdout=subprocess.PIPE,
                                   stderr=subprocess.STDOUT, timeout=300)
            except subprocess.TimeoutExpired:
                print("[pdf] %s 超时" % headless)
                continue
            if os.path.isfile(OUT_PDF) and os.path.getsize(OUT_PDF) > 10240:
                print("[pdf ] written: %s  (%.2f MB)" %
                      (OUT_PDF, os.path.getsize(OUT_PDF) / 1048576.0))
                return True
            print("[pdf] %s 失败（退出码 %s）" % (headless, r.returncode))
            for line in r.stdout.decode("utf-8", "replace").strip().splitlines()[-5:]:
                print("       " + line)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("[pdf] !! PDF 生成失败")
    return False


if __name__ == "__main__":
    ok = build_pdf()
    # Markdown 版由 HTML 转换而来（见 _build_md.py），保证三份内容一致
    sys.path.insert(0, HERE)
    try:
        import _build_md
        _build_md.main()
    except Exception as exc:                      # 转换失败不影响 HTML / PDF 产物
        print("[md  ] !! 生成失败：%r" % (exc,))
    sys.exit(0 if ok else 1)
