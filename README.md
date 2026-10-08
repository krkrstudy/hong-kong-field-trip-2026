# 香江翱翔 香港访学电子路书

第五届北京大学“香江翱翔”访港研学团，2026年10月11日至17日。正式时间安排唯一依据为香港校友会2026年10月7日V4行程，核查日期2026年10月8日。

## 本地运行

安装 Node.js 22+。网站不需要安装 npm 依赖、不需要私密密钥。

```bash
python3 -m http.server 8765
```

打开 http://localhost:8765。不要直接双击 index.html：浏览器需要通过 HTTP 读取 data.json。

```bash
npm run check
npm run build
```

构建结果位于 `dist/`。页面与资源均采用相对路径，可放在 GitHub Pages 子目录下。

## 文件与更新

- `data.json`：单一行程数据源。`days` 管理时间、活动、地点引用、交通和待确认状态；`places` 管理地址、坐标证据、介绍、照片及署名。
- `assets/images/`：经优化的实景照片。照片作者、许可与来源在 `sources.json` 和各地点详情中。
- `assets/maps/`：Word 使用的 OpenStreetMap 静态参考地图；地点编号不表示道路路线。
- `downloads/`：与网站配套的 Word 和 PDF。
- `scripts/build_maps.py`：按data.json中的已核实坐标生成真实底图参考点图，需要curl、Pillow；只请求所需瓦片并本地缓存。改变地点坐标后先运行此脚本。
- `scripts/build_docx.py`：根据同一 data.json 生成可编辑 Word。需 Python、python-docx、Pillow，以及STSong、Arial Unicode MS或等价的完整中文字体。
- `vendor/`：本地 Leaflet 1.9.4，BSD-2-Clause 许可见 Leaflet-LICENSE。
- `CHANGELOG.md`：相较旧版的更新摘要。
- `VERIFICATION.md`：事实边界、交通风险和验收结果。

修改时间后，同时检查 `time`、`end`、`status`、交通段落和提醒文字；不要只改显示标题。修改访问地点后必须核对地址、坐标及实际入口，更新地图与署名，重新生成和渲染 Word/PDF，再发布网站。网页的完成态不是预约确认。

Word 生成示例：

```bash
python3 scripts/build_docx.py
```

用 Word、LibreOffice 或文档渲染工具导出 PDF，并逐页检查。目录使用真正的 TOC、PAGEREF 字段；如有编辑，请在 Word 中更新全部域和目录。当前交付的页码已经过渲染核对。

## 部署

目标仓库 `krkrstudy/hong-kong-field-trip-2026`。当前使用 GitHub Pages 的 **Deploy from a branch**：`main` 分支根目录。GitHub 自动运行内置 Pages 构建与部署流程；推送后自动更新。

网址：https://krkrstudy.github.io/hong-kong-field-trip-2026/

当前 OAuth 授权缺少 `workflow` scope，GitHub 拒绝创建自定义工作流文件。因此 `deployment/pages.yml.example` 保留了完整的自定义 Actions 配置，未声称该配置已启用。

日常更新：

1. 修改 `data.json`，重新生成并渲染 Word/PDF。
2. 运行 `npm run check` 和 `npm run build`，本地检查后推送 `main`。
3. 在仓库 Actions 查看内置 Pages 部署结果，并检查实际网站。

若以后具备工作流写入权限：将 `deployment/pages.yml.example` 复制到 `.github/workflows/pages.yml`，将 Settings → Pages 的 Source 改为 **GitHub Actions**，然后推送。该自定义流程会自动执行数据检查、静态构建和部署。

若更换仓库，需要同步修改 README、index.html 的分享网址以及 Word 生成器中的网站链接；网页资源使用相对路径。

## 设计与技术取舍

采用纯 HTML/CSS/JavaScript + 本地 Leaflet，而不是 React：本项目是只读的七日行程，轻量静态结构无需维护构建依赖或服务器，适合手机及 GitHub Pages。日期切换、地图筛选、地点弹窗、导航和分享链接均由结构化数据驱动。原生 dialog 提供焦点约束和 Escape 关闭；动画尊重系统减少动态效果偏好。

参考了 Leaflet 官方 quick start、HandsOnDataViz 的 Leaflet storymaps 与 skeate/Leaflet.timeline 的地图叙事及时间组织思路。没有复制其项目代码，网页代码为本项目编写。

- https://leafletjs.com/examples/quick-start/
- https://github.com/HandsOnDataViz/leaflet-storymaps-with-google-sheets
- https://github.com/skeate/Leaflet.timeline

照片存放本地、懒加载；地图底图仍需联网。弱网或离线时使用已下载PDF。没有 Service Worker，以避免出行中看到被长期缓存的旧行程。

## 公开范围与许可

未上传原始邀请函、原始V4 PDF、私人联系人、电话、邮箱或成员名单。网站只保留所需的行程和公开机构资料。

项目自编前端代码为 MIT 许可，见 LICENSE。行程内容由活动组织方提供；照片逐张适用原许可（包括 CC BY-SA），不得据代码许可推断照片权利。照片已缩放及显示裁切，署名、原文件链接和许可均保留。地图数据 © OpenStreetMap contributors，ODbL；底图服务使用须遵守 OpenStreetMap tile usage policy。

## 尚需确认

爱诗科技未提供地址，无法核实本次接待场地，故无假配图、无地图点；欢迎晚宴和校友座谈场地未明。港大时间重叠、中诚信房号差异、接待入口及部分交通衔接详见页面与 VERIFICATION.md。不可将未核实的办公室地址或普通校园图片当作本次指定房间。
