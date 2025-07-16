import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import subprocess
import shutil

# 创建主窗口
root = tk.Tk()
root.title("文件夹处理工具")  # 设置窗口标题
root.geometry("500x650")  # 增加窗口高度以容纳新功能

# 创建StringVar存储路径
source_path = tk.StringVar()
output_path = tk.StringVar()
old_extension = tk.StringVar(value=".GIF")
new_extension = tk.StringVar(value=".7z")
delete_after_extract = tk.BooleanVar(value=False)  # 解压后删除压缩文件选项
extract_password = tk.StringVar(value="www.gezi8.top")  # 解压密码存储变量
process_second_extraction = tk.BooleanVar(value=True)  # 新增：是否进行二次解压选项

# 添加一个文本标签
label = tk.Label(root, text="文件夹后缀重命名与解压工具", font=("微软雅黑", 16))
label.pack(pady=10)


# 获取Windows启动配置（隐藏命令行窗口）
def get_startup_info():
    if os.name == 'nt':
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0  # 隐藏窗口
        return startupinfo
    return None


# 源文件夹选择功能
def choose_source():
    path = filedialog.askdirectory(title="选择源文件夹")
    if path:
        source_path.set(path)
        label_source.config(text=f"源文件夹: {path}")


# 输出文件夹选择功能
def choose_output():
    path = filedialog.askdirectory(title="选择输出文件夹")
    if path:
        output_path.set(path)
        label_output.config(text=f"输出文件夹: {path}")


# 递归查找并重命名分卷文件
def find_and_rename_volume_files(directory):
    """
    递归查找目录中所有以"删除"结尾的分卷压缩文件，并重命名去掉"删除"后缀
    返回重命名后的文件列表
    """
    renamed_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            # 查找文件名中包含".7z."且以"删除"结尾的文件[1,2](@ref)
            if file.endswith('删除') and '.7z.' in file:
                old_path = os.path.join(root, file)
                # 去掉文件名末尾的"删除"字符
                new_name = file.rstrip('删除')
                new_path = os.path.join(root, new_name)

                # 重命名文件[6](@ref)
                shutil.move(old_path, new_path)
                renamed_files.append(new_path)
    return renamed_files


# 解压分卷压缩文件
def extract_volume_files(volume_files, password, delete_after=False):
    """
    解压分卷压缩文件，使用相同的密码
    返回成功和失败计数
    """
    success_count = 0
    failures = []

    for volume_file in volume_files:
        try:
            # 创建解压目录（基于分卷文件名）
            base_name = os.path.basename(volume_file).replace('.001', '')
            extract_dir = os.path.join(os.path.dirname(volume_file), base_name)
            if not os.path.exists(extract_dir):
                os.makedirs(extract_dir)

            # 调用7z解压分卷文件（自动识别后续卷）[1,2](@ref)
            cmd = ['7z', 'x', f'-o{extract_dir}', '-y']
            if password:
                cmd.extend(['-p' + password])
            cmd.append(volume_file)

            subprocess.run(cmd, check=True, startupinfo=get_startup_info())
            success_count += 1

            # 解压后删除分卷文件（如果启用）
            if delete_after:
                # 删除同一目录下的所有分卷文件（.001, .002等）
                dir_path = os.path.dirname(volume_file)
                base_pattern = os.path.basename(volume_file).split('.')[0]
                for f in os.listdir(dir_path):
                    if f.startswith(base_pattern) and f.endswith(('.001', '.002', '.003')):
                        os.remove(os.path.join(dir_path, f))

        except subprocess.CalledProcessError as e:
            failures.append(f"{os.path.basename(volume_file)}: 解压失败 (错误代码 {e.returncode})")
        except Exception as e:
            failures.append(f"{os.path.basename(volume_file)}: {str(e)}")

    return success_count, failures


# 处理功能（带密码支持和二次解压）
def process_files():
    source = source_path.get()
    output = output_path.get()
    old_ext = old_extension.get().strip()
    new_ext = new_extension.get().strip()
    delete_after = delete_after_extract.get()
    password = extract_password.get().strip()
    do_second_extraction = process_second_extraction.get()  # 是否进行二次解压

    if not source or not output:
        messagebox.showerror("错误", "请先选择源文件夹和输出文件夹")
        return

    if not old_ext or not new_ext:
        messagebox.showerror("错误", "请填写有效的文件后缀")
        return

    # 确保扩展名以点开头
    if not old_ext.startswith('.'):
        old_ext = '.' + old_ext
    if not new_ext.startswith('.'):
        new_ext = '.' + new_ext

    # 检查7z是否可用
    try:
        subprocess.run(['7z', '-h'], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                       startupinfo=get_startup_info(), check=True)
    except (FileNotFoundError, subprocess.CalledProcessError):
        messagebox.showerror("错误", "未找到7z程序，请确保已安装7-Zip并添加到系统PATH")
        return

    try:
        # 获取所有匹配的文件
        files = [f for f in os.listdir(source)
                 if f.lower().endswith(old_ext.lower()) and os.path.isfile(os.path.join(source, f))]

        if not files:
            messagebox.showinfo("完成", f"没有找到 {old_ext} 格式的文件")
            return

        success_count = 0
        skipped_count = 0
        extract_success = 0
        extract_failures = []
        moved_files = []  # 存储已移动的文件路径

        for filename in files:
            # 创建新文件名
            new_filename = re.sub(r'\.gif$', new_ext, filename, flags=re.IGNORECASE)

            # 构建完整路径
            old_file = os.path.join(source, filename)
            new_file = os.path.join(output, new_filename)

            # 如果目标文件已存在则跳过
            if os.path.exists(new_file):
                skipped_count += 1
                continue

            # 重命名（移动）文件
            shutil.move(old_file, new_file)
            success_count += 1
            moved_files.append(new_file)  # 记录已移动的文件

        # 解压文件（带密码支持）
        for compressed_file in moved_files:
            # 更新状态
            status_bar.config(text=f"正在解压: {os.path.basename(compressed_file)}")
            root.update()

            try:
                # 创建解压目录
                extract_dir = os.path.join(output, os.path.splitext(os.path.basename(compressed_file))[0])
                if not os.path.exists(extract_dir):
                    os.makedirs(extract_dir)

                # 调用7z解压（带密码支持）[1,2](@ref)
                cmd = ['7z', 'x', f'-o{extract_dir}', '-y']
                if password:
                    cmd.extend(['-p' + password])
                cmd.append(compressed_file)

                subprocess.run(cmd, check=True, startupinfo=get_startup_info())

                # 如果需要，解压后删除压缩文件
                if delete_after:
                    os.remove(compressed_file)

                extract_success += 1
            except subprocess.CalledProcessError as e:
                # 处理密码错误情况
                extract_failures.append(f"{os.path.basename(compressed_file)}: 解压失败 (错误代码 {e.returncode})")
            except Exception as e:
                extract_failures.append(f"{os.path.basename(compressed_file)}: {str(e)}")

        # 二次处理：查找并解压分卷压缩文件
        second_extract_success = 0
        second_extract_failures = []
        volume_files = []

        if do_second_extraction and extract_success > 0:
            # 递归查找并重命名分卷文件
            status_bar.config(text="正在查找分卷压缩文件...")
            root.update()
            renamed_files = find_and_rename_volume_files(output)

            if renamed_files:
                status_bar.config(text=f"找到 {len(renamed_files)} 个分卷文件，准备解压...")
                root.update()

                # 筛选出.001分卷文件（7z会自动处理后续卷）[1,2](@ref)
                volume_files = [f for f in renamed_files if f.endswith('.001')]

                if volume_files:
                    # 解压分卷文件
                    second_extract_success, second_extract_failures = extract_volume_files(
                        volume_files, password, delete_after
                    )

        # 显示结果
        message = (f"处理完成！\n\n"
                   f"成功重命名: {success_count} 个文件\n"
                   f"跳过已存在: {skipped_count} 个文件\n"
                   f"第一次解压成功: {extract_success} 个文件")

        if do_second_extraction:
            message += f"\n第二次解压成功: {second_extract_success} 个分卷文件"
            message += f"\n找到的分卷文件: {len(volume_files)} 个"

        if extract_failures or second_extract_failures:
            all_failures = extract_failures + second_extract_failures
            message += f"\n\n解压失败 ({len(all_failures)} 个文件):\n"
            message += "\n".join(all_failures[:3])  # 最多显示3个错误
            if len(all_failures) > 3:
                message += f"\n...及其他 {len(all_failures) - 3} 个错误"

        messagebox.showinfo("完成", message)
        status_bar.config(
            text=f"完成: {success_count}重命名 | {extract_success}解压 | {second_extract_success}二次解压")

    except Exception as e:
        messagebox.showerror("错误", f"处理过程中发生错误:\n{str(e)}")
        status_bar.config(text="处理失败")


# 添加后缀设置区域
def create_extension_frame():
    frame = tk.LabelFrame(root, text="文件后缀设置", padx=10, pady=10)
    frame.pack(fill=tk.X, padx=20, pady=10)

    # 旧后缀设置
    tk.Label(frame, text="原后缀:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    old_entry = tk.Entry(frame, textvariable=old_extension, width=10)
    old_entry.grid(row=0, column=1, padx=5, pady=5)

    # 新后缀设置
    tk.Label(frame, text="新后缀:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
    new_entry = tk.Entry(frame, textvariable=new_extension, width=10)
    new_entry.grid(row=0, column=3, padx=5, pady=5)

    # 提示标签
    tk.Label(frame, text="示例: .GIF → .7z", fg="gray").grid(row=0, column=4, padx=10, pady=5)

    return frame


# 添加解压选项区域（带密码输入）
def create_extract_options():
    frame = tk.LabelFrame(root, text="解压选项", padx=10, pady=10)
    frame.pack(fill=tk.X, padx=20, pady=10)

    # 解压后删除压缩文件选项
    delete_check = tk.Checkbutton(frame, text="解压后删除压缩文件",
                                  variable=delete_after_extract)
    delete_check.grid(row=0, column=0, padx=5, pady=5, sticky="w")

    # 新增：二次解压选项
    second_extract_check = tk.Checkbutton(frame, text="启用二次解压（处理分卷文件）",
                                          variable=process_second_extraction)
    second_extract_check.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    # 添加密码输入区域
    tk.Label(frame, text="解压密码:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    password_entry = tk.Entry(frame, textvariable=extract_password, width=20)
    password_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    # 提示标签
    tk.Label(frame, text="无密码请留空", fg="gray").grid(row=1, column=2, padx=5, pady=5)

    # 显示/隐藏密码复选框
    show_password = tk.BooleanVar(value=False)

    def toggle_password():
        if show_password.get():
            password_entry.config(show="")
        else:
            password_entry.config(show="*")

    show_check = tk.Checkbutton(frame, text="显示密码", variable=show_password,
                                command=toggle_password)
    show_check.grid(row=1, column=3, padx=5, pady=5)

    return frame


# 源文件夹选择区域
frame_source = tk.Frame(root)
frame_source.pack(fill=tk.X, padx=20, pady=10)

btn_source = tk.Button(frame_source, text="选择源文件夹", command=choose_source, width=15)
btn_source.pack(side=tk.LEFT)

label_source = tk.Label(frame_source, text="未选择源文件夹", anchor="w")
label_source.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

# 输出文件夹选择区域
frame_output = tk.Frame(root)
frame_output.pack(fill=tk.X, padx=20, pady=10)

btn_output = tk.Button(frame_output, text="选择输出文件夹", command=choose_output, width=15)
btn_output.pack(side=tk.LEFT)

label_output = tk.Label(frame_output, text="未选择输出文件夹", anchor="w")
label_output.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

# 添加后缀设置区域
create_extension_frame()

# 添加解压选项区域（带密码输入）
create_extract_options()

# 开始处理按钮
btn_process = tk.Button(root, text="开始处理", command=process_files,
                        bg="#4CAF50", fg="white", font=("微软雅黑", 10),
                        height=2, width=15)
btn_process.pack(pady=20)

# 状态栏
status_bar = tk.Label(root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# 启动主事件循环
root.mainloop()