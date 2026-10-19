# 脚本用法

Python 3.10+，仅标准库。以脚本自身位置定位资源，无本机绝对路径依赖。

```sh
python scripts/reuse.py catalog --library validated
python scripts/reuse.py build --spec brief.json --output output.pptx
python scripts/reuse.py validate output.pptx
```

`brief.json`：

```json
{"slides":[
  {"slide":1,"texts":{
    "标题_可编辑":"护理教学汇报",
    "正文_一级":"教学目标",
    "正文_说明":"在此填写本次汇报内容。",
    "正文_长文本":"在此填写教学流程和训练安排。"
  }},
  {"slide":2,"texts":{
    "标题_可编辑":"团队分工",
    "图标标签_0":"医师","图标标签_1":"护理","图标标签_2":"记录",
    "正文_说明":"在此填写团队协作要求。"
  },"transforms":{"复用_医师":{"x":190,"y":255,"height":160}}}
]}
```

页码从 1 开始。重复选择同一源页生成独立副本。替换继承首个文本运行和段落格式，显式指定微软雅黑；换行用 `\n`。替换不保留原文本框内部混合字符格式，不自动缩字号。未指定文本保留，应自行检查清理。

默认库 `validated` 为四页样例；`--library source` 为完整范本。顶层形状名称或 ID 可从 catalog 查询。

提取源第 34 页两个图标到独立一页的 `icons.json`：

```json
{"slides":[{"slide":34,"keep_shapes":[99,249],"transforms":{
  "99":{"x":200,"y":230,"height":160},
  "249":{"x":600,"y":230,"height":160}
}}]}
```

```sh
python scripts/reuse.py build --library source --spec icons.json --output icons.pptx
```

`keep_shapes` 仅保留指定顶层对象，仍保留母版背景。坐标尺寸使用 CSS px；只给 width 或 height 保持原比例，同时给两项则严格使用用户数值。默认拒绝覆盖已有输出。找不到对象或内部关系缺失时停止。

脚本用于同库页面复制与对象编辑；不承担新绘图、跨库对象合成、SmartArt 重建、图表数据修改或 WPS 界面操作。图标跨文件插入请使用演示工具并核对主题和依赖。
