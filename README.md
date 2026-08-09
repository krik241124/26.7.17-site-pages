Outreach Email Generator (外链开发邮件自动化生成器)
这是一个为独立站外链开发量身定制的 Python 自动化脚本。它可以根据你挖掘到的域名和邮箱列表，结合自定义邮件模板，批量生成排版友好的定制化 Outreach 邮件，方便你逐一复制发送。
📁 目录结构
在使用前，请确保你的项目文件夹包含以下基本结构（运行脚本会自动创建缺失的文件夹）：
code
Text
outreach/
│
├── input/
│   └── targets.xlsx           # 你的原始数据源
│
├── output/
│   └── outreach_emails.xlsx   # 脚本生成的邮件成品
│
├── templates/
│   └── template_01.txt        # 你的邮件模板文案
│
└── generate_emails.py         # Python 运行脚本
🚀 使用指南
1. 安装环境依赖
在首次使用前，请打开终端（CMD/Terminal），在 outreach 目录下运行以下命令安装必备的库：
code
Bash
pip install pandas openpyxl xlsxwriter
2. 准备原始数据 (Input)
将你用 GPT 或其他工具扒下来的数据直接粘贴并保存为 input/targets.xlsx。
💡 重点说明：
不需要手动加表头！ 你只需直接把记录复制进去即可。脚本如果检测到没有表头，会自动为你插入标准的 29 列字段。
脚本会自动寻找包含“域名”和“邮箱”数据的列。如果没有提取到邮箱，输出时会自动留空。
自动匹配的 29 个标准字段依次为：
目标域名 (Domain)
分层与优先级 (Tier)
🤖机器评分 (Auto Score)
内容类型 (Link Type)
建议策略 (Action Suggested)
竞品重合度 (Alpha)
🔗该域名下总外链 (Total Links)
DR权重
页面流量
Dofollow
Spam
对标链接样例
🧠AI：业务相关性评分 (0-100)
🧠AI：配套内容 (如 Pitch 邮件|大纲)
🧠AI：难度评级 (Easy|Hard)
🧑‍💻联系邮箱 (Email)
🧑‍💻获取入口URL
🏆总分 (Total Score)（自动计算）
🧑‍💻人工：审阅状态
🧑‍💻当前状态 (Status)
Last Action Date
Link Result
🤝合作模式 (Cooperation Model)
💰合作报价 (Cost/Price)
✍️内容提供方 (Content Provider)
📍发布板块与位置 (Placement)
⏳链接留存期 (Link Duration)
🔗约定链接属性 (Agreed Attribute)
身份
3. 设置邮件模板 (Templates)
编辑 templates/template_01.txt 文件：
第一行：脚本会默认将其提取为邮件主题 (Subject)。
剩余内容：将被视为邮件正文 (Body)。
动态替换：在正文中写入 [Name]，脚本会自动根据对方域名（比如将 uscreen.tv 转换成 Uscreen）进行替换，让邮件看起来更自然、更定制化。
4. 运行脚本
在终端中执行以下命令：
code
Bash
python generate_emails.py
5. 提取发送结果 (Output)
脚本运行完毕后，打开 output/outreach_emails.xlsx。
为了方便手动发送（如 Outlook），生成的数据结构极简且排版友好，仅包含 4 列：
Domain (域名)
To (收件邮箱) - 如果未找到邮箱，此列为空
Subject (邮件主题)
Body (邮件正文)
✨ 排版优化： 正文列（Body）已默认加宽并开启了“自动换行”。你只需要双击该单元格，直接 Ctrl+A 全选然后 Ctrl+C 复制，粘贴到邮件客户端即可，格式不会乱套！