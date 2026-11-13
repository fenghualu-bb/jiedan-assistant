# 可复用脚本

> 全部依赖 **Windows + Microsoft PowerPoint（COM）**，Mac 上只能阅读学习。连接时一律 `DispatchEx` + 断言 `app.Path` 含 `Office16`。

| 文件夹 | 用途 | 入口 / 说明 |
|---|---|---|
| `英语课件美化/` | 美化现有课件全流程：COM 构建、XML 后处理、QA、动画比对、导出、拼对比图 | 用法见 `../附录_英语课件美化.md` 第 8 节；`example_build_unit4.py` 是完整示例 |
| `固废批量/` | 多份同模板课件批量 4:3→16:9 + 统一排版 | `engine.py → fit.py → strip_media.py → topdf.py`；检查 `render.py / grid.py / check.py`；各章配置在 `decks/`；说明见 `../交接文档/固废批量.md` |

演讲稿（原生 PPTX + 翻书视频 + 背景音乐）的构建脚本没有拷过来，只有交接文档 `../交接文档/演讲稿_算盘接力.md`。
