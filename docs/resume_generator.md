# fastResume 使用与配置说明

一份 YAML 或 JSON 配置生成一份简历，支持可编辑 Word 和 PDF。内容与排版配置分开，不绑定姓名、岗位或项目，也没有固定的技能组数、教育条数或项目条数。

默认使用青蓝配色、常见图标和分组技能卡片。项目按简述、贡献和技术栈呈现。正文支持 `**粗体**`，项目贡献中简短的“职责前缀：”也会自动加粗。

## 安装依赖

需要 Python 3.10 或更新版本。在项目根目录运行：

```bash
python -m pip install -r requirements.txt
```

Word 文件由 Python 直接生成。导出 PDF 还需要以下任一种本机环境：

- Windows 上已安装的 Microsoft Word。
- 已安装的 LibreOffice，且 `soffice` 可在命令行找到，或通过 `--office-path` 指定。

只生成 DOCX 不需要安装 Office。程序不会自动安装 Word、LibreOffice 或字体，也不会上传简历内容。

## 快速开始

从匿名模板 `examples/resume.template.yaml` 复制一份 `my_resume.yaml`，填写真实经历。Windows PowerShell 使用 `Copy-Item examples/resume.template.yaml my_resume.yaml`，macOS / Linux 使用 `cp examples/resume.template.yaml my_resume.yaml`，然后运行：

```bash
python scripts/build_resume.py --source my_resume.yaml --output output/my_resume --formats docx pdf
```

`--output` 是输出文件前缀，上述命令生成：

```text
output/my_resume.docx
output/my_resume.pdf
```

文件默认拒绝覆盖。确认要更新同名文件时，增加 `--force`。也可以更换前缀，保留旧版。

只检查配置：

```bash
python scripts/build_resume.py --source my_resume.yaml --validate-only
```

只生成可编辑 Word：

```bash
python scripts/build_resume.py --source my_resume.yaml --output output/my_resume --formats docx
```

对 PDF 设置最多一页，并允许覆盖本次指定的输出：

```bash
python scripts/build_resume.py --source my_resume.yaml --output output/my_resume --formats docx pdf --max-pages 1 --force
```

页数超限会报错；程序不会删改内容或自动缩小字号来凑页数。减少重复文字、调整内容顺序，或明确修改字体和页边距后，再生成并检查。

## 本地网页编辑与快速导出

安装依赖后，在项目根目录运行：

```bash
python scripts/serve.py
```

打开 [http://127.0.0.1:8765](http://127.0.0.1:8765)。编辑器提供基础信息表单与完整 YAML 编辑区，可切换四份预设、追加教育/技能/实习/项目/奖项条目，并在右侧查看预览。YAML 通过校验后才能导出；空栏目会隐藏。正文支持 `**粗体**`。

点击“导入配置”读取本地 YAML / JSON，点击“保存 YAML”下载当前配置。网页不会自动保存草稿，刷新或关闭前请保存。照片会作为图片 Data URI 一起写入 YAML，便于换机器后继续编辑；配置文件最大 6 MB。

“导出 Word”直接下载可编辑文档；“导出 PDF”调用本机 Word 或 LibreOffice。导出在临时目录中完成，下载响应准备好后清理服务端临时文件。浏览器预览用于检查内容与样式，字体、换行及最终页数仍以导出的文件为准。

服务只绑定 `127.0.0.1`，不调用外部简历服务或模型 API。端口被占用时，运行 `python scripts/serve.py --port 8766` 后打开对应地址；在终端按 Ctrl+C 停止。

## 随附示例

所有示例使用同一个生成器与相同字段结构，没有为角色另写排版程序。

| 配置 | 用途 | 主色 |
| --- | --- | --- |
| `examples/resume.template.yaml` | 从零填写的最小匿名模板 | 青蓝 |
| `examples/anon.yaml` | 千早爱音主题，产品与全栈开发演示 | 粉色 |
| `examples/tomori.yaml` | 高松灯主题，AI 应用与检索演示 | 蓝色 |
| `examples/sakiko.yaml` | 丰川祥子主题，后端与平台开发演示 | 紫色 |

这三份角色主题模板用于展示排版与文字层次。学校、年级、生日及乐队担当取自官方资料，来源链接列在 README；开发经历融入原创乐队技术梗。年龄按高一年级统一推算为 16 岁。“活动与成就”使用官方人物页确认的原作乐队活动和担当，未采用声优真人演出经历。

生成其中一份：

```bash
python scripts/build_resume.py --source examples/tomori.yaml --output output/tomori --formats docx pdf --max-pages 1
```

README 中的三张预览对应上述角色配置；替换自己的内容时，无需保留角色姓名、玩笑或主题色。

## PDF 转换选项

自动选择本机可用的转换方式：

```bash
python scripts/build_resume.py --source my_resume.yaml --output output/my_resume --formats pdf --pdf-backend auto
```

在 Windows 上明确使用 Word：

```powershell
python scripts/build_resume.py --source my_resume.yaml --output output/my_resume --formats docx pdf --pdf-backend word --timeout 120
```

明确使用 LibreOffice：

```bash
python scripts/build_resume.py --source my_resume.yaml --output output/my_resume --formats docx pdf --pdf-backend libreoffice --office-path /path/to/soffice
```

Windows 的路径示例：

```powershell
python scripts/build_resume.py --source my_resume.yaml --output output/my_resume --formats pdf --pdf-backend libreoffice --office-path "C:\Program Files\LibreOffice\program\soffice.exe"
```

`--timeout 120` 将转换等待时间设为 120 秒。`--max-pages` 需要 PDF 模式才能取得页数。只有导出 PDF 时才会调用转换程序；PDF 单独导出所需的中间 DOCX 不作为正式文件保留。

## 配置结构

YAML 与 JSON 使用相同字段。配置文件名可自行决定，程序按 `--source` 指定的文件读取。JSON 文件不能包含 YAML 注释。

| 字段 | 内容 |
| --- | --- |
| `basics` | 姓名、岗位、联系方式、地点与到岗状态；`details` 为可选页眉补充文本，`photo` 为可选证件照 |
| `code` | 作品集、代码仓库和其他链接 |
| `profile` | 求职意向、相关能力与希望承担的工作 |
| `education` | 教育经历列表 |
| `skills` | 分组技能，或摘要加技能列表 |
| `internship_experiences` | 可选的实习经历，结构同项目经历；没有则省略或留空 |
| `work_experiences` | 可选的真实工作经历，结构同项目经历 |
| `project_experiences` | 项目经历列表 |
| `certificates` | 奖项和证书列表 |
| `extra_sections` | 可选的自定义栏目 |
| `section_titles` | 自定义标准栏目名称 |
| `theme` | 纸张、颜色、字体、字号、行距和页边距 |

### 基本信息与链接

```yaml
basics:
  name: 匿名键盘手
  position: 软件开发实习生
  subtitle: SOFTWARE DEVELOPMENT INTERN
  eyebrow: PERSONAL RESUME
  tagline: 可以迷路，调用链不能迷路
  phone: ""
  email: member@example.com
  location: ""
  status: 到岗时间可协商
  details: ""  # 可填补充资料；支持多行文本，不会自动计算年龄
  photo: ""    # 可选；网页上传后自动填入图片 Data URI

code:
  - label: 作品集
    url: https://example.com
    icon: globe
```

不用的可选信息可留空。电话号码建议加引号，避免 YAML 把它当成数字；链接填写完整 URL。

### 可选证件照

网页中点击“上传照片”，选择不超过 2 MB 的 PNG / JPEG / WebP 图片；程序还会检查图片是否有效及像素总量是否超过 2000 万。照片按 3:4 居中裁剪显示。未添加照片时保持原双栏页眉；有照片时显示左侧照片、中间姓名与岗位、右侧联系信息。点击“移除”即可恢复原版式。

网页上传的照片以 `data:image/png;base64,...`、`data:image/jpeg;base64,...` 或 `data:image/webp;base64,...` 形式保存在 `basics.photo` 中。网页接口只接受图片 Data URI；命令行读取配置时也支持相对配置文件目录的本地图片路径，不支持网络 URL。通过网页上传和“保存 YAML”即可生成可迁移的配置，无需手工编辑 Base64 文本。

多人物照片可设置 `basics.photo_focus: [0.25, 0.5]`：横向、纵向坐标均为 0–1，表示原图中的主体焦点，默认 `[0.5, 0.5]` 居中；预览和导出会围绕焦点进行 3:4 裁剪。新上传的照片会恢复默认居中。

命令行也可从本地图片为本次导出添加或替换照片：

```bash
python scripts/build_resume.py --source my_resume.yaml --photo portrait.jpg --output output/my_resume --formats docx pdf
```

`--photo PATH` 不改写来源配置文件。照片仍可省略：保持 `photo: ""`、设为 `null` 或删除该字段即可。

### 教育、技能与经历

```yaml
education:
  - date: 20XX.09 - 至今
    school: 某大学
    major: 计算机相关专业｜本科
    detail: 毕业年月与课程信息按实际情况填写

skills:
  groups:
    - label: 后端开发
      value: Python、FastAPI、SQLAlchemy
    - label: 工程实践
      value: Git、Docker、Pytest

internship_experiences:
  - name: 实际公司或团队｜实习岗位
    date: 20XX.XX - 20XX.XX
    role: 实际承担的职责
    url: ""
    description: 团队业务与实习工作范围；没有真实实习经历时不填。
    highlights:
      - "实际贡献：填写自己完成的工作、具体方案和可核实的结果。"
    tech_stack: Python · FastAPI

project_experiences:
  - name: 项目名称
    date: 20XX.XX - 至今
    role: 本人负责的实际职责
    url: ""
    description: 说明服务对象、解决的问题和主要功能。
    highlights:
      - "**可靠性设计：**写清实现机制、验证条件和实际结果。"
      - "性能优化：说明优化对象、测量方式和数据变化。"
    tech_stack: Python · FastAPI · SQLite
```

项目简述和个人成果分别填写：`description` 介绍项目用途，`highlights` 说明本人做了什么。`role` 写职责，`tech_stack` 写实际使用的技术。项目 URL 可以为空。

`internship_experiences` 与 `work_experiences` 使用同样的条目结构，分别渲染为“实习经历”和“工作经历”。实习经历默认位于技能之后、项目之前；没有对应经历时可省略字段或设为 `[]`。

如果不使用技能分组，也可以写：

```yaml
skills:
  summary: 一句能力概括
  items:
    - 第一项技能
    - 第二项技能
```

### 奖项与自定义栏目

```yaml
certificates:
  - name: 证书或奖项名称
    result: 实际等级或成绩
  - 也可以直接填写一条完整文本

extra_sections:
  - title: 技术分享
    items:
      - 分享主题与本人贡献

section_titles:
  profile: 求职摘要
  education: 教育背景
  skills: 专业技能
  work_experiences: 工作经历
  internship_experiences: 实习经历
  project_experiences: 项目经历
  certificates: 竞赛奖项与证书  # 角色主题可改为“活动与成就”
```

没有内容的栏目可删除或留空，不需要为了填满版面虚构经历。

### 空内容自动隐藏

只有 `basics.name` 必填。其余可选字段可以省略、设为 `null`，文本也可以写成空字符串或纯空格；列表可以写成 `[]`。例如：

```yaml
basics:
  name: 你的姓名
  phone: ""
profile: null
education: []
internship_experiences: null
work_experiences: []
project_experiences:
  - name: ""
    description: "  "
    highlights: []
certificates: []
```

上述空栏目及空项目卡片会被隐藏。有效条目中的空字段也不占位：项目未填简述或技术栈时，相关标签一并隐藏；技能只填内容时不生成空分类名；教育只填补充说明时直接显示说明。奖项只填名称或结果时合并为一栏，避免留下空单元格。

### 排版主题

```yaml
theme:
  paper: A4                    # A4 或 Letter
  accent: "#0D7187"
  ink: "#13242D"
  muted: "#52636D"
  background: "#EFF8FA"
  border: "#D3E2E7"
  font: Calibri
  east_asia_font: Microsoft YaHei
  body_size: 8.5               # 磅
  line_height: 12.6            # 磅
  margin_mm:
    top: 10
    right: 14
    bottom: 9
    left: 14
```

颜色值需要加引号，防止 `#` 被 YAML 解释成注释。正文行距 `line_height` 至少为字号 `body_size` 的 1.1 倍。字体需要在导出 PDF 的机器上可用；Linux/macOS 可改为本机安装的中文字体，例如 `Noto Sans CJK SC`。

## 保存、清理与检查

程序先生成并检查本次输出，再写出正式文件；默认不覆盖已有文件。转换使用的临时目录在成功或失败后都会清理，已有简历和输入配置不作为临时文件删除。

不同字体、Word/LibreOffice 版本可能影响换行和页数。`--validate-only` 检查数据结构，`--max-pages` 检查 PDF 页数，二者都不能代替视觉检查。投递前请打开导出的 PDF，确认中文显示、链接、项目边界和分页正常。

## 换成自己的内容

从 `examples/resume.template.yaml` 填写自己的信息，或选择喜欢的角色主题后替换人物资料和经历。程序按配置原文导出，不会自动补全或改写内容。
