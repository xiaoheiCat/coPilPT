# coPilPT
coPilPT 是由 AI 驱动的 PPT 构思与创建秘书

---

## coPilPT 和其他 PPT 生成器有什么不同？
据我所知，目前主流的 PPT 生成器都是采用了 AI 生成内容 + 预设模板的格式进行制作。于是我自研了 coPilPT，这是一个 100% AI 驱动的多 Agents PPT 构思与创建秘书。与其他需要模板的 PPT 生成器不同，coPilPT 无需任何的 PPT 模板，从始至终，除了您给 coPilPT 提供大致的方向以外，几乎所有的一切都是 coPilPT 自己完成的！

## coPilPT 的原理是什么？
我们有“构思”、“大纲分页”、“单页生成”三个 Agents，“构思”Agent 将会帮助您创建一个完善的 PPT 大纲，“大纲分页”Agent 将会将大纲进行分页，然后再由“单页生成”Agent 使用 HTML 编写页面。当一切结束后，Selenium 将会使用 Chromium 浏览器渲染这些网页并输出为图片，接下来多张图片会被合并为一个 PDF，再将 PDF 转换为 PPTX 文件。于是，您就可以获得最终的 PPTX 文件了！

## 作为一个普通用户，如何使用它？
> ⚠️ 这是我预期的效果，但是有几个步骤似乎仍然很难正常运行。
1. 前往“构思”选项卡，回答 AI 提出的所有问题
2. AI 将会自动生成大纲并跳转到“创建”选项卡
3. 点击“从大纲自动生成”并等待创建完成
4. 从新弹出的文件窗口中点击“下载”图标。

## 作为一个开发者，该如何部署它？
> ⚠️ 这是一个非常早期的项目，您可能会遇到非常多的问题，包括但不限于生成的 PPT 始终过于简单，效果达不到您的预期，部分自动化操作仍未通过测试等。当然，如果您还是非常感兴趣，您也可以尝试继续部署：
1. clone 当前仓库到本地
```bash
git clone https://github.com/xiaoheiCat/coPilPT.git
```
2. 进入本仓库
```bash
cd coPilPT
```
3. 使用 Docker Compose 启动它
```bash
docker compose up -d
```
> 如果你没有安装 Docker，请先安装 Docker。以下是在 Ubuntu / Debian 中的示例：
```
sudo apt-get install -y docker.io docker-compose-v2
```

## 我想对本项目做出贡献！
当然！欢迎 PR！提前感谢您对 coPilPT 的支持！

## Star 历史
[![Star History Chart](https://api.star-history.com/svg?repos=xiaoheiCat/coPilPT&type=Date)](https://www.star-history.com/#xiaoheiCat/coPilPT&Date)
