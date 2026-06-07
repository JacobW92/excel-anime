# 🎬 Excel Anime Player

> **全程由 [Claude Code](https://docs.anthropic.com/en/docs/claude-code) 制作 — 从代码到 README，完全由 AI 驱动开发。**
>
> *Built entirely with [Claude Code](https://docs.anthropic.com/en/docs/claude-code) — from code to README, fully AI-driven development.*

将视频（比如动画 OP）转换成 Excel 像素动画。每个 Excel 格子 = 一个像素，通过宏自动滚动播放，录屏后看起来就像在用 Excel 播放动画。

![](https://img.shields.io/badge/platform-macOS%20%7C%20Windows-blue) ![](https://img.shields.io/badge/python-3.9+-green) ![](https://img.shields.io/badge/license-MIT-orange)

---

## ✨ 效果

1. 📹 输入一个视频文件
2. 🎨 每个像素变成 Excel 格子的背景色（正方形格子）
3. ▶️ 自动播放：逐帧滚动，就像在 Excel 里播动画
4. 🎥 录屏 + 配上原版音频 → 完整的 "Excel 版动画 OP"

## 🚀 快速开始

### 安装依赖

```bash
# Python 依赖
pip install -r requirements.txt

# macOS 需要 ffmpeg 和 Homebrew
brew install ffmpeg

# 播放功能需要 xlwings（已包含在 requirements.txt 中）
```

### 生成 Excel 动画

```bash
# 基础用法（默认 8fps, 200px 宽, 16:9 自动高度）
python video_to_excel.py your_video.mp4

# 自定义参数
python video_to_excel.py your_video.mp4 --fps 8 --width 200

# 限制帧数（适合先测试效果）
python video_to_excel.py your_video.mp4 --max-frames 60

# 同时保存竖排拼接长图
python video_to_excel.py your_video.mp4 --save-stacked

# 调整色彩深度（默认 32 levels/channel，最大 32,768 色）
python video_to_excel.py your_video.mp4 --colors 16
```

### 播放动画

**macOS（推荐）：**

```bash
python play.py your_video_excel.xlsx
```

自动打开 Excel、设置缩放、逐帧滚动播放。

**Windows / 手动播放：**

1. 用 Excel 打开生成的 `.xlsx`
2. 另存为 `.xlsm` 格式（启用宏）
3. `Alt + F11` → 文件 → 导入 → `anime_player_macro.bas`（自动生成）
4. `F5` 运行 `PlayAnimation`

然后开录屏！配上原版音频就是完整的 Excel 动画了。

## 📋 参数说明

### video_to_excel.py

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `video` | （必填） | 输入视频文件 |
| `--fps` | 8 | 目标帧率 |
| `--width` | 200 | 帧宽度（像素） |
| `--height` | 0 | 帧高度（0 = 自动 16:9） |
| `--output`, `-o` | 自动命名 | 输出 Excel 文件名 |
| `--max-frames` | 0 | 最大帧数（0 = 全部） |
| `--save-stacked` | 关闭 | 同时保存竖排拼接 PNG |
| `--colors` | 32 | 色彩深度（levels/channel），Excel 限制约 64,000 种唯一格式 |

### play.py

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `excel_file` | （必填） | Excel 文件路径 |
| `width` | 200 | 帧宽度 |
| `height` | 112 | 帧高度 |
| `num_frames` | 720 | 总帧数 |
| `fps` | 8 | 播放帧率 |

## 🛠️ 技术细节

- **视频抽帧**：ffmpeg 按目标帧率提取帧
- **图像处理**：Pillow 缩放 + 竖排拼接
- **色彩量化**：将 RGB 各通道量化到 32 级（32,768 色），保持在 Excel ~64,000 唯一格式限制内
- **Excel 生成**：xlsxwriter 高效写入，每个格子设置正方形尺寸 + 背景色
- **播放控制**：macOS 通过 AppleScript `Application.Goto Scroll:=True` 精确对齐帧边界

## 📁 项目结构

```
excel-anime/
├── video_to_excel.py    # 主工具：视频 → Excel
├── play.py              # macOS 播放器
├── requirements.txt     # Python 依赖
└── README.md            # 本文件
```

生成的文件（不需要提交）：

```
*_excel.xlsx             # 生成的 Excel 像素动画
*_stacked.png            # 竖排拼接长图
anime_player_macro.bas   # 自动生成的 VBA 宏
```

## ⚠️ 注意事项

- 200px 宽、90 秒视频约生成 **50MB** Excel 文件，生成时间约 45 秒
- Excel 打开大文件需要一定时间，`play.py` 默认等待 15 秒加载
- macOS Excel 的 AppleScript 接口有限，`play.py` 使用 `Goto Scroll:=True` 实现精确滚动
- 建议先用 `--max-frames 30` 测试效果，满意后再生成完整版

## 📄 License

MIT
