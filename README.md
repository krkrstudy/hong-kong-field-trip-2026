# 燕园翱翔香港实践团路书2026

2026年10月11日至17日香港实践活动随身路书。网页文案按最新版正式 Word 整理，采用北大红白配色、仿宋／宋体正文、黑体标题及1.5倍行距。保留行程时间，删除草稿中的待确认状态和相关说明。没有地址的活动不补写地址，也不提供推测的定位或照片。

网站：https://krkrstudy.github.io/hong-kong-field-trip-2026/

## 手机使用

1. 用 Safari 或 Chrome 联网打开网站，等待“已保存，可离线查阅”。初次保存约23 MB，包括行程、地点实景照片、每日地图、Word 和 PDF。
2. 点击“添加到主屏幕”，按提示操作。iPhone 用 Safari 的分享菜单，Android 可用浏览器安装菜单。
3. 首次从主屏幕打开时保持联网，再次确认保存完成。之后可以断网打开、切换日期及阅读地点详情。
4. 实时导航、互动地图及外部网站需要联网。默认参考地图可离线查阅。
5. 网站更新时，新版内容完整下载后会出现“更新路书”。点击后统一切换版本。保存中断时可点击“重新保存”。清理浏览器数据后需要重新保存，建议另将 PDF 存到手机本地。

## 本地运行与构建

无需安装项目依赖或配置私密密钥。Node.js 22+ 用于构建。

```sh
python3 -m http.server 8765
node scripts/build.mjs
```

通过 http://localhost:8765 打开；不能直接双击 HTML 使用离线功能。构建生成 `dist/`，资源使用相对路径，支持 GitHub Pages 仓库子目录。`node scripts/prepare-pwa.mjs` 根据所有发布资源及 Service Worker 模板生成内容哈希版本，`node scripts/check.mjs` 检查正式文案、行程和资源完整性。

## 更新

- `data.json`：每日行程、地点介绍、学习要点、交通安排及团员须知。
- `downloads/燕园翱翔香港实践团路书2026.docx`：用户最新版正式 Word 的原文件副本。
- `downloads/燕园翱翔香港实践团路书2026.pdf`：从该 Word 重新导出并逐页核对的 PDF。
- `assets/images/`、`assets/maps/`：本地照片及已生成的参考地图，地图不表示行车路线。
- `manifest.webmanifest`、`pwa.js`：安装信息及手机端保存、更新提示。
- `scripts/sw-template.js`、`scripts/prepare-pwa.mjs`：离线缓存逻辑和版本生成器。
- `sw.js`：生成的发布文件，须在每次资源修改后重新生成并提交。

正式版文字以用户当前 Word 为准。更新 Word 时同步修改网页数据并导出新 PDF；不要运行旧版 `scripts/build_docx.py` 覆盖正式下载文件。旧生成器仅作为历史工具保留。

发布前运行 `node scripts/build.mjs`，检查手机布局、全部日期、地点弹窗及真实断网后的重新打开。随后推送 `main`。仓库使用 GitHub Pages 的 Deploy from a branch（main 根目录）；内置 Pages 流程自动部署。无需自定义 Actions 工作流。

## 离线实现

首次安装完整保存本次资源；任一资源失败即不宣称已保存。所有本地读取使用同一缓存版本。更新在后台保存完整资源，等待用户点击后激活，随后删除旧版本缓存并刷新页面，避免行程与下载文件混用版本。Word 和 PDF 均缓存完整文件；支持 PDF 的 Range 请求。导航缓存限定当前仓库路径，不影响同一 GitHub Pages 域名下的其他项目。

互动地图使用本地 Leaflet 和联网 OpenStreetMap 瓦片；不批量预存地图瓦片。离线地图采用已有静态参考图。网络中断或底图加载失败时返回参考图。

## 公开内容与署名

未上传原始邀请函、私人联络表、电话、邮箱和成员名单。图片保留作者及许可，照片为往年实景，已缩放及显示裁切。代码 MIT 许可不适用于照片；图片按各自许可使用，见 `sources.json`。地图 © OpenStreetMap contributors，ODbL；Leaflet BSD-2-Clause 许可见 vendor/Leaflet-LICENSE。

技术依据：[Service Worker API](https://developer.mozilla.org/en-US/docs/Web/API/Service_Worker_API)、[PWA installation](https://web.dev/learn/pwa/installation)。
