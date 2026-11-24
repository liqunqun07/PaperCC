# PaperCC · 论文审阅与返修工作台

PaperCC 是一个面向研究者的本地 Word 审阅工具：**把「读懂审阅意见 → 修改对应原文 → 整理逐条回复」放进同一个界面。** 它适合导师反馈、合作者改稿和期刊返修。无需搭建服务器，也无需接入现有科研网关。

| 场景 | PaperCC 能做什么 |
| --- | --- |
| 查看导师或合作者的意见 | 打开 `.docx`，从批注跳到正文，切换显示修订与清洁视图。 |
| 处理返修稿 | 接受或拒绝修订；写入批注、回复和带修订痕迹的修改；支持备份与撤销。 |
| 集中处理多条意见 | 自配 AI 后，可排队解读批注、起草修改建议或回复。建议由使用者审阅后再采用，不会由队列直接写回文档。 |
| 整理逐条回复 | 将生成的回复草稿汇总到 Word 和 Markdown，便于定稿。 |
| 继续上次审阅 | 浏览器记住最近打开的文件。macOS 可选菜单栏入口。 |

## 安装与运行

推荐 macOS、Python 3.11 或 3.12；无需 Word/WPS 即可在浏览器审阅。选择文件弹窗、调用本地编辑器及菜单栏入口使用 macOS 功能。

```sh
git clone https://github.com/liqunqun07/PaperCC.git
cd PaperCC
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m server.main
```

浏览器访问 <http://127.0.0.1:8765>。在页面输入 `.docx` 或所在文件夹的路径，也可点「选择…」。**建议先复制一份稿件再审阅：保存操作会写入打开的文件**，应用会在写入前保留本地备份。不要同时在其他编辑器中修改同一文件。仅打开你信任的文档。

已有应用占用端口时，用 `PAPERCC_PORT=8875 python -m server.main`，再打开对应端口。服务仅供本机使用，不要经反向代理或端口转发对外开放。

### 可选：macOS 菜单栏入口

服务启动后，在另一个终端进入仓库运行 `swift macos/Menu.swift`（需要 Xcode 命令行工具）。菜单栏点击 **PaperCC → 打开论文审阅**。如果改过端口，运行时设置同一个 `PAPERCC_PORT`。退出菜单栏入口不会停止服务。

### 可选：自己配置 AI

不配置 AI 也可以查看批注、编辑和处理 Word 修订。要启用 AI 解读、翻译、改写或生成回复，配置一个兼容 OpenAI Chat Completions 的服务：

```sh
mkdir -p -m 700 ~/.papercc
cp -n config/config.example.json ~/.papercc/config.json
```

编辑 `~/.papercc/config.json`：将 `base_url` 改为服务的 API 基础地址（通常以 `/v1` 结尾，不要附加 `/chat/completions`），将 `model` 改为供应商提供的模型 ID。在启动服务的终端设置 `PAPERCC_API_KEY` 后重启：

```sh
read -rs PAPERCC_API_KEY
export PAPERCC_API_KEY
python -m server.main
```

执行第一行后粘贴密钥并按回车；密钥不会显示。不要把它写进仓库。AI 操作会将相关批注和稿件片段发送到你配置的模型服务。敏感或未发表稿件请先确认是否允许外发。回复稿提交前须人工核对。

历史、缓存、备份与导出默认保存在 `~/.papercc/`，可用 `PAPERCC_DATA_DIR` 更换目录。此仓库不包含用户稿件、真实密钥、私有网关或录音转写功能。

## 本地测试

安装依赖后运行 `python -m unittest discover -s tests -v`。测试只用现场生成的示例文档和本地模拟模型服务。
