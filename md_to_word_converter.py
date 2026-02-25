import subprocess
import os

def convert_markdown_to_word(input_md_file, output_docx_file, reference_doc_template):
    """
    使用 pandoc 将 Markdown 文件转换为指定 Word 模板的 DOCX 文件。

    参数:
    input_md_file (str): 输入的 Markdown 文件路径。
    output_docx_file (str): 输出的 Word 文档文件路径。
    reference_doc_template (str): 作为参考的 Word 模板文件路径。
    """
    
    # 检查文件是否存在
    if not os.path.exists(input_md_file):
        print(f"错误: 输入的 Markdown 文件不存在: {input_md_file}")
        return
    if not os.path.exists(reference_doc_template):
        print(f"错误: 参考模板文件不存在: {reference_doc_template}")
        return

    # 构建 pandoc 命令。
    # 将命令及其参数作为列表传递是推荐的做法，
    # 这样可以避免 shell 注入风险，并让系统正确处理参数中的空格。
    command = [
        "pandoc",
        "--reference-doc", reference_doc_template,
        "-s", # --standalone 的缩写，表示生成一个独立的文档
        input_md_file,
        "-o", output_docx_file
    ]

    print(f"正在执行命令: {' '.join(command)}")

    try:
        # 执行命令
        # check=True: 如果命令返回非零退出代码（表示错误），将抛出 CalledProcessError。
        # text=True (或 encoding='utf-8'): 将 stdout 和 stderr 解码为文本。
        # capture_output=True: 捕获 stdout 和 stderr。
        result = subprocess.run(command, check=True, text=True, capture_output=True, encoding='utf-8')

        print("\n转换成功！")
        print(f"输入文件: {input_md_file}")
        print(f"输出文件: {output_docx_file}")
        print(f"使用模板: {reference_doc_template}")
        
        if result.stdout:
            print("\nPandoc 输出 (stdout):\n", result.stdout)
        if result.stderr:
            print("\nPandoc 错误/警告 (stderr):\n", result.stderr)

    except FileNotFoundError:
        print(f"错误: 'pandoc' 命令未找到。请确保 'pandoc' 已安装并添加到系统的 PATH 环境变量中。")
    except subprocess.CalledProcessError as e:
        print(f"错误: Pandoc 转换失败，返回代码: {e.returncode}")
        print(f"命令: {' '.join(e.cmd)}")
        if e.stdout:
            print(f"Pandoc 输出 (stdout):\n{e.stdout}")
        if e.stderr:
            print(f"Pandoc 错误/警告 (stderr):\n{e.stderr}")
    except Exception as e:
        print(f"发生未知错误: {e}")

# --- 配置您的文件路径 ---
# 获取当前脚本所在的目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 假设 template.docx 和 input.md 在脚本的同级目录下
# 您可以根据实际情况修改这些路径
INPUT_MD = os.path.join(SCRIPT_DIR, "input.md")
OUTPUT_DOCX = os.path.join(SCRIPT_DIR, "output.docx")
REFERENCE_TEMPLATE = os.path.join(SCRIPT_DIR, "template.docx")

# --- 调用转换函数 ---
if __name__ == "__main__":
    # 示例：创建一些虚拟文件以便测试
    # 请确保您有真实的 input.md 和 template.docx
    if not os.path.exists(INPUT_MD):
        with open(INPUT_MD, "w", encoding="utf-8") as f:
            f.write("# 这是一个测试标题\n\n- 列表项1\n- 列表项2\n\n**加粗文本** 和 *斜体文本*。")
        print(f"已创建示例输入文件: {INPUT_MD}")
    
    # 对于 template.docx，您需要一个实际的 Word 模板文件。
    # 如果没有，pandoc 会使用其默认样式转换，但会给出模板不存在的警告。
    # 这里我们只检查它是否存在。
    if not os.path.exists(REFERENCE_TEMPLATE):
        print(f"警告: 模板文件 '{REFERENCE_TEMPLATE}' 不存在。Pandoc 将使用默认样式。")
        print("建议您放置一个 'template.docx' 文件在脚本同目录下以获得更佳效果。")

    convert_markdown_to_word(INPUT_MD, OUTPUT_DOCX, REFERENCE_TEMPLATE)

    # 也可以在不传入路径的情况下执行，如果文件在当前工作目录
    # convert_markdown_to_word("input.md", "output.docx", "template.docx")
