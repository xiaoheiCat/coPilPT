import os
import json
import time
import urllib
import gradio as gr
import requests
import img2pdf
import subprocess
from dotenv import load_dotenv
from selenium import webdriver
from pdf2pptx import convert_pdf2pptx
from langchain_openai import ChatOpenAI
from langchain.schema import HumanMessage, SystemMessage, AIMessage
from langchain.prompts import PromptTemplate
from langchain.callbacks.base import BaseCallbackHandler

# 加载环境变量
if os.getenv("LOAD_DOTENV") == "1":
    load_dotenv()
OPENAI_API_BASE = os.getenv("OPENAI_API_BASE")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL")

# 初始化 LangChain 模型
llm = ChatOpenAI(
    model=OPENAI_MODEL,
    openai_api_key=OPENAI_API_KEY,
    openai_api_base=OPENAI_API_BASE,
    temperature=0.3
)

class FunctionCallHandler(BaseCallbackHandler):
    def __init__(self):
        self.function_called = False
        self.outline = ""
        self.function_call_content = ""
        self.full_response = ""
    
    def on_llm_new_token(self, token, **kwargs):
        self.full_response += token
        
        # 增强标记检测逻辑
        begin_marker = "<|FunctionCallBegin|>"
        end_marker = "<|FunctionCallEnd|>"
        
        if begin_marker in self.full_response and not self.function_called:
            self.function_called = True
            self.function_call_content = ""
            begin_pos = self.full_response.find(begin_marker) + len(begin_marker)
            self.function_call_content = self.full_response[begin_pos:]
        
        if self.function_called and end_marker in self.full_response:
            end_pos = self.full_response.find(end_marker)
            if end_pos > 0:
                self.function_call_content = self.full_response[:end_pos].split(begin_marker)[-1]
                self.parse_function_call()
                # 立即重置状态
                self.full_response = ""
                self.function_called = False
    
    def parse_function_call(self):
        try:
            # 处理可能的JSON格式错误
            content = self.function_call_content.strip()
            if content.startswith("["):  # 处理数组格式
                function_data = json.loads(content)[0]
            else:  # 处理对象格式
                function_data = json.loads(content)
            
            if function_data.get("name") == "topic2ppt":
                self.outline = function_data.get("parameters", {}).get("content", "")
                print(f"[DEBUG] 成功提取大纲内容: {self.outline[:50]}...")  # 添加调试日志
        except json.JSONDecodeError as e:
            print(f"[ERROR] JSON解析失败: {str(e)}")
            print(f"[DEBUG] 原始内容: {self.function_call_content}")
        except Exception as e:
            print(f"[ERROR] 函数调用处理异常: {str(e)}")

    def on_llm_end(self, response, **kwargs):
        # 清理状态
        self.full_response = ""

# 定义网页渲染代码
def render_html(body_code):
    header = """
<html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="https://cdn.tailwindcss.com?plugins=forms,typography"></script>
		<script src="https://unpkg.com/unlazy@0.11.3/dist/unlazy.with-hashing.iife.js" defer init></script>
		<script type="text/javascript">
			window.tailwind.config = {
				darkMode: ['class'],
				theme: {
					extend: {
						colors: {
							border: 'hsl(var(--border))',
							input: 'hsl(var(--input))',
							ring: 'hsl(var(--ring))',
							background: 'hsl(var(--background))',
							foreground: 'hsl(var(--foreground))',
							primary: {
								DEFAULT: 'hsl(var(--primary))',
								foreground: 'hsl(var(--primary-foreground))'
							},
							secondary: {
								DEFAULT: 'hsl(var(--secondary))',
								foreground: 'hsl(var(--secondary-foreground))'
							},
							destructive: {
								DEFAULT: 'hsl(var(--destructive))',
								foreground: 'hsl(var(--destructive-foreground))'
							},
							muted: {
								DEFAULT: 'hsl(var(--muted))',
								foreground: 'hsl(var(--muted-foreground))'
							},
							accent: {
								DEFAULT: 'hsl(var(--accent))',
								foreground: 'hsl(var(--accent-foreground))'
							},
							popover: {
								DEFAULT: 'hsl(var(--popover))',
								foreground: 'hsl(var(--popover-foreground))'
							},
							card: {
								DEFAULT: 'hsl(var(--card))',
								foreground: 'hsl(var(--card-foreground))'
							},
						},
					}
				}
			}
		</script>
		<style type="text/tailwindcss">
			@layer base {
				:root {
					--background: 0 0% 100%;
--foreground: 240 10% 3.9%;
--card: 0 0% 100%;
--card-foreground: 240 10% 3.9%;
--popover: 0 0% 100%;
--popover-foreground: 240 10% 3.9%;
--primary: 240 5.9% 10%;
--primary-foreground: 0 0% 98%;
--secondary: 240 4.8% 95.9%;
--secondary-foreground: 240 5.9% 10%;
--muted: 240 4.8% 95.9%;
--muted-foreground: 240 3.8% 46.1%;
--accent: 240 4.8% 95.9%;
--accent-foreground: 240 5.9% 10%;
--destructive: 0 84.2% 60.2%;
--destructive-foreground: 0 0% 98%;
--border: 240 5.9% 90%;
--input: 240 5.9% 90%;
--ring: 240 5.9% 10%;
--radius: 0.5rem;
				}
				.dark {
					--background: 240 10% 3.9%;
--foreground: 0 0% 98%;
--card: 240 10% 3.9%;
--card-foreground: 0 0% 98%;
--popover: 240 10% 3.9%;
--popover-foreground: 0 0% 98%;
--primary: 0 0% 98%;
--primary-foreground: 240 5.9% 10%;
--secondary: 240 3.7% 15.9%;
--secondary-foreground: 0 0% 98%;
--muted: 240 3.7% 15.9%;
--muted-foreground: 240 5% 64.9%;
--accent: 240 3.7% 15.9%;
--accent-foreground: 0 0% 98%;
--destructive: 0 62.8% 30.6%;
--destructive-foreground: 0 0% 98%;
--border: 240 3.7% 15.9%;
--input: 240 3.7% 15.9%;
--ring: 240 4.9% 83.9%;
				}
			}
		</style>
  </head>
  <body>
"""

    footer = """
  </body>
</html>
    """

    return header + body_code + footer

# 定义系统消息
system_message = """
# 角色
你叫柴特·斯莱德，是一位严谨的 PPT 制作秘书。你的说话风格俏皮，生成的 PPT 严谨且适用于正式场合，且不包含任何动画。

## 技能
### 技能 1: 信息收集
通过与用户聊天一问一答的形式，向用户提出问题（不多于 8 个），收集制作 PPT 所需信息。

### 技能 2: 生成 PPT 大纲
当认为信息收集完毕时，帮用户头脑风暴一个 PPT 大纲，大纲必须遵循以下格式 (替换括号中的内容)：
```
# (主题)

## 目录
1. (子标题1)
2. (子标题2)
...

## (子标题1)
- **(小主题1)**: (正文)
- **(小主题2)**: (正文)
...

## (子标题2)
- **(小主题1)**: (正文)
- **(小主题2)**: (正文)
...
```
整个聊天始终保持简体中文，正文字数要足够，需包含正常 PPT 应有的所有内容。

### 技能 3: 大纲修改
生成大纲后，询问用户有没有需要修改的地方，并根据用户的需求对大纲进行修改。

### 技能 4: 生成 PPT
当用户明确要求生成 PPT 时，告诉用户你正在生成，请保留该页面在前台并稍作等候...再调用 topic2ppt 函数并在 content 参数传入 PPT 的完整大纲进行生成。

## 限制:
- 整个对话围绕 PPT 制作展开，拒绝回答与 PPT 制作无关的话题。
- 所输出的 PPT 大纲必须按照给定的格式进行组织，不能偏离框架要求。
- 生成的正文内容要符合 PPT 内容要求，不能过于简略。 
"""

pagination_system_message = """
# (主题)

## 目录
1. (子标题1)
2. (子标题2)
...

## (子标题1)
- **(小主题1)**: (正文)
- **(小主题2)**: (正文)
...

## (子标题2)
- **(小主题1)**: (正文)
- **(小主题2)**: (正文)
...

---

用户将传入一个 PPT 大纲，并且符合以上格式。你的任务是传入的大纲进行分片，但是不调整任何格式与文本。例如，以上的大纲可以分片为：
```
# (主题)

---

## 目录
1. (子标题1)
2. (子标题2)
...

---

## (子标题1)

---

### (子标题1)
- **(小主题1)**: (正文)
- **(小主题2)**: (正文)
...

---

## (子标题2)

---

### (子标题2)
- **(小主题1)**: (正文)
- **(小主题2)**: (正文)
...
```

按照此规则，对用户传入的 PPT 大纲进行分片。
"""
generator_system_message = """
🎉 Greetings, TailwindCSS Virtuoso! 🌟

You've mastered the art of frontend design and TailwindCSS! Your mission is to transform detailed descriptions or compelling images into stunning HTML using the versatility of TailwindCSS. Ensure your creations are seamless in both dark and light modes! Your designs should be responsive and adaptable across all devices - be it desktop, tablet, or mobile. Except for the code, do not output anything else.

*Design Guidelines:*
- Utilize placehold.co for placeholder images and descriptive alt text.
- For interactive elements, leverage modern ES6 JavaScript and native browser APIs for enhanced functionality.
- Inspired by shadcn, we provide the following colors which handle both light and dark mode:

```css
  --background
  --foreground
  --primary
	--border
  --input
  --ring
  --primary-foreground
  --secondary
  --secondary-foreground
  --accent
  --accent-foreground
  --destructive
  --destructive-foreground
  --muted
  --muted-foreground
  --card
  --card-foreground
  --popover
  --popover-foreground
```

Prefer using these colors when appropriate, for example:

```html
<button class="bg-secondary text-secondary-foreground hover:bg-secondary/80">Click me</button>
<span class="text-muted-foreground">This is muted text</span>
```

*Implementation Rules:*
- Only implement elements within the `<body>` tag, don't bother with `<html>` or `<head>` tags.
- Avoid using SVGs directly. Instead, use the `<img>` tag with a descriptive title as the alt attribute and add .svg to the placehold.co url, for example:

```html
<img aria-hidden="true" alt="magic-wand" src="https://openui.fly.dev/icons/24x24.svg?text=🪄" />
```

"""

generator_user_message = "根据以下 PPT 大纲，制作一张单独的精美无比的 PPT 页面。要能够在 4:3 比例的显示器上能一次性全部显示，不可以出现需要滚动才能查看全部内容的情况。保持所有文本内容不变，只是生成 HTML 样式。除了黑灰白外一定要加点其他的颜色，且最好具有高级感。不要生成动画效果。文字与背景的对比度明显，使观众可以很容易的阅读。主题样式应根据内容主题所决定，并且统一规定，封面页与子标题页的标题都要居中显示，但是正文页内容要根据主题样式的排版而定。例如，可以在正文页创建几个图形，文字排布在图形旁。发挥你的想象力，完成 PPT 的创建。"

# 定义函数
functions = [
    {
        "name": "topic2ppt",
        "description": "根据用户的 PPT 大纲，为用户生成 PPT。",
        "parameters": {
            "type": "object",
            "properties": {
                "content": {  # 修改参数名
                    "type": "string",
                    "description": "PPT 的完整大纲内容"
                }
            },
            "required": ["content"]  # 修改必填参数
        }
    }
]

def predict(query, history, tab_state):
    if not query:
        gr.Warning("不能发送空白消息。")
        return query, history, tab_state
    
    # 准备消息历史
    messages = [SystemMessage(content=system_message)]
    for h in history:
        if h[0]:  # 用户消息
            messages.append(HumanMessage(content=h[0]))
        if h[1]:  # AI消息
            messages.append(AIMessage(content=h[1]))
    
    # 添加当前查询
    messages.append(HumanMessage(content=query))
    
    # 处理函数调用
    function_handler = FunctionCallHandler()
    
    # 先将用户消息添加到历史中
    history.append((query, ""))
    
    # 流式响应文本
    response_text = ""
    complete_response = ""
    
    # 设置 callbacks 初始化 llm 对象
    llm_with_callbacks = ChatOpenAI(
        model=OPENAI_MODEL,
        openai_api_key=OPENAI_API_KEY,
        openai_api_base=OPENAI_API_BASE,
        temperature=0.3,
        callbacks=[function_handler],
        streaming=True
    )
    
    # 使用流式API
    for chunk in llm_with_callbacks.stream(
        messages,
        functions=functions
    ):
        if hasattr(chunk, 'content') and chunk.content:
            complete_response += chunk.content
            
            if "<|FunctionCallBegin|>" in complete_response:
                # 获取函数调用前的文本
                begin_pos = complete_response.find("<|FunctionCallBegin|>")
                visible_text = complete_response[:begin_pos]
                
                # 只更新可见部分
                if visible_text != response_text:
                    response_text = visible_text
                    history[-1] = (query, response_text)
                    # 添加完整的输出参数
                    yield "", history, tab_state, gr.Button(interactive=False), gr.Textbox(interactive=False), gr.Button(interactive=False)
            
            elif "<|FunctionCallEnd|>" in complete_response:
                # 获取函数调用后的文本
                end_pos = complete_response.find("<|FunctionCallEnd|>") + len("<|FunctionCallEnd|>")
                if end_pos < len(complete_response):
                    visible_text = complete_response[:complete_response.find("<|FunctionCallBegin|>")] + complete_response[end_pos:]
                    response_text = visible_text
                    history[-1] = (query, response_text)
                    # 添加完整的输出参数
                    yield "", history, tab_state, gr.Button(interactive=False), gr.Textbox(interactive=False), gr.Button(interactive=False)
            
            else:
                # 没有函数调用时也返回完整参数
                response_text = complete_response
                history[-1] = (query, response_text)
                yield "", history, tab_state, gr.Button(interactive=False), gr.Textbox(interactive=False), gr.Button(interactive=False)

    # 流式传输结束后手动解析部分需要修复
    if "<|FunctionCallBegin|>" in complete_response and "<|FunctionCallEnd|>" in complete_response:
        try:
            begin_pos = complete_response.find("<|FunctionCallBegin|>") + len("<|FunctionCallBegin|>")
            end_pos = complete_response.find("<|FunctionCallEnd|>")
            function_call_content = complete_response[begin_pos:end_pos]
            
            # 解析函数调用内容
            function_data = json.loads(function_call_content.strip())  # 添加 strip() 处理空白字符
            if isinstance(function_data, list) and len(function_data) > 0:
                function_data = function_data[0]
            
            if function_data.get("name") == "topic2ppt":
                outline = function_data.get("parameters", {}).get("content", "")
                if outline:
                    print(f"手动解析成功，大纲长度：{len(outline)}")
                    # 同时更新 function_handler 的状态
                    function_handler.outline = outline  # 新增此行
                    # 更新tab状态
                    tab_state = {
                        "active_tab": 1,
                        "outline": outline
                    }
                    gr.Info("coPilPT 已生成 PPT 大纲，可以开始根据大纲自动创建 PPT 了。")
        except Exception as e:
            print(f"手动解析函数调用时出错: {e}")

    # 统一处理大纲状态
    if function_handler.outline:  # 修改判断条件
        # 清理最终显示的文本
        final_text = complete_response
        # 移除所有函数调用标记
        final_text = final_text.replace("<|FunctionCallBegin|>", "").replace("<|FunctionCallEnd|>", "")
        
        # 更新最终显示
        history[-1] = (query, final_text.strip())
        
        # 确保状态更新
        tab_state = {
            "active_tab": 1,
            "outline": function_handler.outline
        }
        gr.Info("coPilPT 已生成 PPT 大纲，可以开始根据大纲自动创建 PPT 了。")
    
    # 返回最终结果并启用控件
    yield "", history, tab_state, gr.Button(interactive=True), gr.Textbox(interactive=True), gr.Button(interactive=True)

def reset_conversation():
    return None, [
        (None, "### 你好，我是你的专属 PPT 小秘书 coPilPT！\n让我们一起头脑风暴吧！今天想做些什么主题的 PPT 呢？")
    ], {"active_tab": 0, "outline": ""}

def update_outline(tab_state):
    if tab_state and "outline" in tab_state:
        return tab_state["outline"]
    return ""

def update_tab(tab_state):
    if tab_state and "active_tab" in tab_state:
        return tab_state["active_tab"]
    return 0

def generate_preview(outline):
    if not outline:
        gr.Warning("需要先提供大纲才可以开始创建 PPT。")
        yield None, None, gr.Button(interactive=True)  # 保持按钮可点击
        return

    try:
        os.remove("coPilPT.pptx")
    except FileNotFoundError:
        pass
    
    gr.Info("正在自动分页...")
    yield None, None, gr.Button(interactive=False)  # 立即禁用按钮
    
    outline = llm.predict_messages(
        [SystemMessage(content=pagination_system_message), HumanMessage(content=outline)]
    ).content
    opt = outline.split("\n---\n")
    gr.Info(f"分页完成。根据大纲内容，我们将制作 {len(opt)} 页 PPT。")
    
    latest_preview = None
    for i, j in enumerate(opt):
        page_intro_prompt = "【提示：该页为正文页】"
        if i == 0:
            page_intro_prompt = "【提示：该页为封面页】"
        elif i == 1:
            page_intro_prompt = "【提示：该页为目录页】"
        elif j.startswith("## "):
            page_intro_prompt = "【提示：该页为子标题页】"
        gr.Info("正在生成第 " + str(i+1) + " 页 PPT...")
        page_html = llm.predict_messages(
            [
                SystemMessage(content=generator_system_message),
                HumanMessage(content=generator_user_message + page_intro_prompt + "\n\n---\n\n标题: " + opt[0] + "\n\n该分页内容:\n\n"+ j),
            ]
        ).content
        image_filename = f'ppt_page_{i}.png'
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1440")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.binary_location = subprocess.check_output(['which', 'chromium']).decode('utf-8').strip()
        driver = webdriver.Chrome(service=webdriver.ChromeService(executable_path=subprocess.check_output(['which', 'chromedriver']).decode('utf-8').strip()), options=chrome_options)
        driver.get("data:text/html;charset=utf-8," + urllib.parse.quote(render_html(page_html)))
        time.sleep(8)
        driver.save_screenshot(image_filename)
        driver.quit()
        latest_preview = image_filename
        yield latest_preview, gr.File(visible=False), gr.Button(interactive=False)  # 保持禁用状态
    
    gr.Info("即将完成...")
    
    # 转换图片为 PDF，再转换为 PPTX
    with open("ppt.pdf", "wb") as f:
        # 修复: 使用img2pdf.convert的正确调用方式
        image_files = [f'ppt_page_{i}.png' for i in range(len(opt))]
        # 将图片路径列表作为参数传递给img2pdf.convert
        pdf_bytes = img2pdf.convert(image_files)
        f.write(pdf_bytes)

    # 删除临时图片
    for i in range(len(opt)):
        if os.path.exists(f'ppt_page_{i}.png'):
            os.remove(f'ppt_page_{i}.png')

    # 转换 PDF 为 PPTX
    convert_pdf2pptx("ppt.pdf", "coPilPT.pptx", 300, 0, None, False)
    os.remove("ppt.pdf")

    # 直接传递文件路径
    ppt_file_path = os.path.abspath("coPilPT.pptx")
    
    gr.Info("任务已完成!")
    yield None, gr.File(value=ppt_file_path, visible=True), gr.Button(interactive=True)  # 最终启用按钮


def create_ui():
    with gr.Blocks(analytics_enabled=False, title="coPilPT") as demo:
        gr.HTML(value="<center><h1>coPilPT 是由 AI 驱动的 PPT 构思与创建秘书</h1></center>")

        tab_state = gr.State({"active_tab": 0, "outline": ""})
        
        with gr.Tabs() as tabs:
            with gr.TabItem("构思", id=0) as tab1:
                chatbot = gr.Chatbot(
                    value=[
                        (None, "### 你好，我是你的专属 PPT 小秘书 coPilPT！\n让我们一起头脑风暴吧！今天想做些什么主题的 PPT 呢？")
                    ],
                    show_label=False,
                    height=500
                )
                with gr.Row():
                    clear_btn = gr.Button(value="清除", variant="stop")
                    with gr.Column(scale=12):
                        textbox = gr.Textbox(
                            placeholder="有问题尽管问我... (Shift + Enter = 换行)",
                            show_label=False,
                            autofocus=True,
                            container=False
                        )
                    submit_btn = gr.Button(value="发送", variant="primary")
            
            with gr.TabItem("创建", id=1) as tab2:
                with gr.Row():
                    with gr.Column(scale=1):
                        outline_textbox = gr.Textbox(
                            label="大纲",
                            placeholder="没有主意吗？前往「构思」选项卡，让 coPilPT 和你一起头脑风暴下吧~\n如果你有已经成型的大纲，也可以直接粘贴在这里哦...",
                            lines=19
                        )
                    with gr.Column(scale=1):
                        preview = gr.Image(label="当前生成页预览", height=450, width=600)
                        download_file = gr.File(label="下载 PPT", visible=False)
                
                with gr.Row():
                    generate_btn = gr.Button(value="根据大纲自动生成", variant="primary", size="lg")
                
                generate_btn.click(
                    generate_preview,
                    inputs=[outline_textbox],
                    outputs=[preview, download_file, generate_btn],  # 添加按钮状态输出
                    show_progress="full"
                )
        
        submit_btn.click(
            predict, 
            inputs=[textbox, chatbot, tab_state], 
            outputs=[textbox, chatbot, tab_state, clear_btn, textbox, submit_btn],
            api_name=False,
            queue=True
        )
        textbox.submit(
            predict, 
            inputs=[textbox, chatbot, tab_state], 
            outputs=[textbox, chatbot, tab_state, clear_btn, textbox, submit_btn],
            api_name=False,
            queue=True
        )
        clear_btn.click(
            reset_conversation, 
            inputs=None, 
            outputs=[textbox, chatbot, tab_state]
        )
        
        tab_state.change(
            update_outline,
            inputs=[tab_state],
            outputs=[outline_textbox]
        )
        tab_state.change(
            update_tab,
            inputs=[tab_state],
            outputs=[tabs]
        )
        
        # Remove the incorrect return statement
        return demo

if __name__ == '__main__':
    ui = create_ui()
    ui.queue().launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=True,
        inbrowser=True,
        show_api=False,
        allowed_paths=["/app"]
    )