from openai import OpenAI
import tkinter as tk
from tkinter import scrolledtext, ttk

API_KEY = "sk-4491f6b0f5a7477da6134f1bc909dd38"
BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
MODEL_NAME = "qwen-turbo"

client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

class ChatDialog:
    def __init__(self, root):
        self.root = root
        self.root.title("阿里云百炼免费对话窗口")
        self.root.geometry("700x550")

        self.chat_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("微软雅黑", 11), state=tk.DISABLED)
        self.chat_area.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        input_frame = ttk.Frame(root)
        input_frame.pack(pady=5, padx=10, fill=tk.X)

        self.input_entry = ttk.Entry(input_frame, font=("微软雅黑", 12))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.input_entry.bind("<Return>", self.send_message)

        send_btn = ttk.Button(input_frame, text="发送", command=self.send_message)
        send_btn.pack(side=tk.RIGHT)

        self.add_message("系统", "已连接阿里云百炼免费模型，输入代码或问题即可对话~")

    def add_message(self, sender, text):
        self.chat_area.config(state=tk.NORMAL)
        self.chat_area.insert(tk.END, f"【{sender}】\n{text}\n\n")
        self.chat_area.config(state=tk.DISABLED)
        self.chat_area.yview(tk.END)

    def send_message(self, event=None):
        user_text = self.input_entry.get().strip()
        if not user_text:
            return
        self.add_message("我", user_text)
        self.input_entry.delete(0, tk.END)

        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[{"role": "user", "content": user_text}]
            )
            reply = response.choices[0].message.content
            self.add_message("百炼助手", reply)
        except Exception as e:
            self.add_message("错误", f"调用失败：{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ChatDialog(root)
    root.mainloop()