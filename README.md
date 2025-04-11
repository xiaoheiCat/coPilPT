# coPilPT
coPilPT is an AI powered secretary for slide ideation and creation.

---

## What makes coPilPT different from other slide generators?
As far as I know, the current mainstream slide generators all adopt the format of AI-generated content combined with preset templates. So, I independently developed coPilPT. It is a 100% AI-driven multi-Agents secretary for slide ideation and creation. Unlike other slide generators that rely on templates, coPilPT doesn't require any slide templates at all. From start to finish, apart from you giving coPilPT a general direction, almost everything is accomplished by coPilPT itself!

## What is the principle of coPilPT?
We have three Agents: "ideation", "outline pagination", and "single-page generation". The "ideation" Agent will assist you in creating a comprehensive slide outline. The "outline pagination" Agent will divide the outline into different pages, and then the "single-page generation" Agent will use HTML to compose the page content. Once everything is completed, Selenium will utilize the Chromium browser to render these web pages and export them as images. Subsequently, multiple images will be merged into a PDF, and this PDF will be converted into a PPTX file. In this way, you can obtain the final PPTX file for your slides!

## As an ordinary user, how can I use it?
> ⚠️ This is the effect I expected, but there are still several steps that seem to be difficult to operate properly.
1. Navigate to the "ideation" tab and answer all the questions posed by the AI.
2. The AI will automatically generate an outline and switch to the "creation" tab.
3. Click "Automatically generate from outline" and wait until the creation process is finished.
4. Click the "download" icon in the newly popped-up file window.

## As a developer, how can I deploy it?
> ⚠️ This is a very early-stage project, and you may encounter numerous issues, including but not limited to the generated slides always being too simplistic, the results not meeting your expectations, and some automated operations still not having passed the testing phase, etc. Of course, if you are still highly interested, you can attempt the deployment as follows:
1. Clone the current repository to your local device.
```bash
git clone https://github.com/xiaoheiCat/coPilPT.git
```
2. Enter the repository directory.
```bash
cd coPilPT
```
3. Launch it using Docker Compose.
```bash
docker compose up -d --build
```
> If you haven't installed Docker yet, please install it first. Here is an example for Ubuntu / Debian systems:
```
sudo apt-get install -y docker.io docker-compose-v2
```

## I want to contribute to this project!
Certainly! Pull requests (PRs) are warmly welcome! Thank you in advance for your support of coPilPT!

## Star History
[![Star History Chart](https://api.star-history.com/svg?repos=xiaoheiCat/coPilPT&type=Date)](https://www.star-history.com/#xiaoheiCat/coPilPT&Date) 
