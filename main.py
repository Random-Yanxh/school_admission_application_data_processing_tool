import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from tkcalendar import DateEntry
import re

class DataProcessorApp:
    def __init__(self, root):
        """
        Initialize the application.
        Args:
            root: The root Tkinter window.
        """
        self.root = root
        self.root.title("进校申请数据处理工具 v1.0")
        self.root.geometry("800x650")

        # --- Data Storage ---
        self.data = []
        self.current_index = 0
        self.file_path = None

        # --- UI Widgets ---
        self.create_widgets()
        self.update_ui_state("disabled") # Disable widgets until file is loaded

    def create_widgets(self):
        """Create all the UI widgets for the application."""
        # --- Main frame ---
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- File Operations Frame ---
        file_frame = ttk.LabelFrame(main_frame, text="文件操作", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        ttk.Button(file_frame, text="导入文件", command=self.import_file).pack(side=tk.LEFT, padx=5)
        self.export_button = ttk.Button(file_frame, text="导出文件", command=self.export_file)
        self.export_button.pack(side=tk.LEFT, padx=5)
        self.file_label = ttk.Label(file_frame, text="尚未导入文件")
        self.file_label.pack(side=tk.LEFT, padx=10)

        # --- Navigation Frame ---
        nav_frame = ttk.Frame(main_frame)
        nav_frame.pack(pady=10)

        self.prev_button = ttk.Button(nav_frame, text="< 上一条", command=self.prev_record)
        self.prev_button.pack(side=tk.LEFT, padx=10)
        self.progress_label = ttk.Label(nav_frame, text="进度: - / -", font=("Helvetica", 12))
        self.progress_label.pack(side=tk.LEFT, padx=10)
        self.next_button = ttk.Button(nav_frame, text="下一条 >", command=self.next_record)
        self.next_button.pack(side=tk.LEFT, padx=10)
        
        # --- Jump Functionality ---
        ttk.Label(nav_frame, text="跳转至:").pack(side=tk.LEFT, padx=(20, 5))
        self.jump_entry = ttk.Entry(nav_frame, width=5)
        self.jump_entry.pack(side=tk.LEFT)
        self.jump_button = ttk.Button(nav_frame, text="跳转", command=self.jump_to_record)
        self.jump_button.pack(side=tk.LEFT, padx=5)

        # --- Data Form Frame ---
        self.form_frame = ttk.LabelFrame(main_frame, text="数据编辑", padding="15")
        self.form_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- Create form fields ---
        self.create_form_fields()

    def create_form_fields(self):
        """Create labels and entry widgets for the data form."""
        self.ui_fields = [
            "访问形式*", "访客姓名*", "手机号*", "证件类型*", "证件号码*", "车辆号码",
            "审批人学工号*", "审批人姓名*", "场所名称*", "访问开始时间*", "访问结束时间*", "拜访人及事由*"
        ]
        self.entries = {}
        
        # Use a grid layout for alignment
        for i, field_text in enumerate(self.ui_fields):
            label = ttk.Label(self.form_frame, text=field_text)
            label.grid(row=i, column=0, sticky=tk.W, padx=5, pady=8)
            
            # Use the text without asterisk as the key
            field_key = field_text.replace('*', '')

            if field_text == "访问形式*":
                self.entries[field_key] = ttk.Combobox(self.form_frame, values=["公务拜访", "入校参观"], width=40)
            elif field_text == "证件类型*":
                self.entries[field_key] = ttk.Combobox(self.form_frame, values=["身份证", "护照"], width=40)
            elif field_text == "场所名称*":
                location_frame = ttk.Frame(self.form_frame)
                location_frame.grid(row=i, column=1, columnspan=2, sticky=tk.W, padx=5)
                
                # 复选框部分
                checkbox_frame = ttk.Frame(location_frame)
                checkbox_frame.pack(anchor=tk.W, pady=(0, 5))
                
                self.location_vars = {
                    "东区": tk.BooleanVar(), "西区": tk.BooleanVar(),
                    "北区": tk.BooleanVar(), "梅山校区": tk.BooleanVar()
                }
                for name, var in self.location_vars.items():
                    ttk.Checkbutton(checkbox_frame, text=name, variable=var).pack(side=tk.LEFT, padx=5)
                
                # 添加文本框
                text_frame = ttk.Frame(location_frame)
                text_frame.pack(anchor=tk.W, fill=tk.X)
                ttk.Label(text_frame, text="其他场所:").pack(side=tk.LEFT)
                self.location_text_entry = ttk.Entry(text_frame, width=30)
                self.location_text_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
                
                self.entries[field_key] = location_frame
                continue  # 跳过后面的grid设置，因为已经在上面设置了
            elif "时间" in field_text:
                date_entry = DateEntry(self.form_frame, width=18, date_pattern='yyyy-mm-dd',
                                       background='darkblue', foreground='white', borderwidth=2)
                time_frame = ttk.Frame(self.form_frame)
                hour_spin = ttk.Spinbox(time_frame, from_=0, to=23, wrap=True, width=3, format="%02.0f")
                minute_spin = ttk.Spinbox(time_frame, from_=0, to=59, wrap=True, width=3, format="%02.0f")
                hour_spin.pack(side=tk.LEFT)
                ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
                minute_spin.pack(side=tk.LEFT)
                
                self.entries[field_key] = (date_entry, hour_spin, minute_spin)
                date_entry.grid(row=i, column=1, sticky=tk.W, padx=5)
                time_frame.grid(row=i, column=2, sticky=tk.W)
                continue
            elif field_text == "拜访人及事由*":
                self.entries[field_key] = tk.Text(self.form_frame, height=4, width=40)
            else:
                self.entries[field_key] = ttk.Entry(self.form_frame, width=43)
            
            # 为非特殊字段设置grid
            if field_key not in ["场所名称"]:
                self.entries[field_key].grid(row=i, column=1, columnspan=2, sticky=tk.W, padx=5)

    def import_file(self):
        """
        Handles file import, performs robust column mapping, and pre-processes all data.
        This ensures data is clean and standardized from the moment it enters the application.
        """
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls;*.xlsx"), ("CSV files", "*.csv")]
        )
        if not path:
            return

        try:
            if path.endswith('.csv'):
                df = pd.read_csv(path, dtype=str, header=0)
            else:
                df = pd.read_excel(path, dtype=str, header=0)

            df = df.fillna('')
            
            # 添加调试信息
            print(f"DataFrame shape: {df.shape}")
            print(f"DataFrame columns: {df.columns.tolist()}")
            print(f"First few rows:\n{df.head()}")
            
            # 确保读取所有数据行，包括第一条数据
            if df.empty:
                messagebox.showwarning("警告", "文件为空，没有数据可处理。")
                return
            
            # 创建列名映射
            column_mapping = {}
            for col in df.columns:
                clean_col = col.strip().replace('*', '')
                column_mapping[clean_col] = col
            
            print(f"Column mapping: {column_mapping}")
            
            # 处理数据
            processed_records = []
            for idx, row in df.iterrows():
                new_record = {}
                
                # 映射已知字段
                field_mappings = {
                    "访问形式": ["访问形式"],
                    "访客姓名": ["访客姓名"],
                    "手机号": ["手机号"],
                    "证件类型": ["证件类型"],
                    "证件号码": ["证件号码"],
                    "车辆号码": ["车辆号码"],
                    "审批人学工号": ["审批人学工号"],
                    "审批人姓名": ["审批人姓名"],
                    "场所名称": ["场所名称"],
                    "访问开始时间": ["访问开始时间"],
                    "访问结束时间": ["访问结束时间"],
                    "拜访人及事由": ["拜访人及事由"]
                }
                
                for field_key, possible_names in field_mappings.items():
                    value = ""
                    for name in possible_names:
                        if name in column_mapping:
                            original_col = column_mapping[name]
                            if original_col in row:
                                value = str(row[original_col]).strip()
                                break
                    
                    # 数据清理
                    if field_key == "车辆号码":
                        value = value.replace(" ", "").upper()
                    
                    new_record[field_key] = value
                
                processed_records.append(new_record)
                print(f"Record {idx + 1}: {new_record}")

            self.data = processed_records
            self.file_path = path
            self.current_index = 0
            
            if not self.data:
                messagebox.showwarning("警告", "文件为空，没有数据可处理。")
                return

            self.file_label.config(text=f"已加载: {path.split('/')[-1]} ({len(self.data)}条记录)")
            
            # 重要：先更新UI状态，再加载记录
            self.update_ui_state("normal")
            self.load_record(self.current_index)
            self.update_progress()
            
            print(f"Loading first record: {self.data[0]}")

        except Exception as e:
            messagebox.showerror("导入错误", f"无法读取文件: {e}")
            print(f"详细错误信息: {e}")
            import traceback
            traceback.print_exc()

    def load_record(self, index):
        """Load data from a specific record index into the UI form."""
        if not self.data or index >= len(self.data):
            return
        
        record = self.data[index]
        print(f"Loading record {index + 1}: {record}")
        
        def get_val(key, default=''):
            value = record.get(key, default)
            print(f"Getting {key}: '{value}'")
            return value

        # 清空并填充表单字段
        try:
            # 访问形式
            self.entries["访问形式"].set(get_val("访问形式"))
            
            # 访客姓名
            self.entries["访客姓名"].delete(0, tk.END)
            self.entries["访客姓名"].insert(0, get_val("访客姓名"))
            
            # 手机号
            self.entries["手机号"].delete(0, tk.END)
            self.entries["手机号"].insert(0, get_val("手机号"))
            
            # 证件类型
            self.entries["证件类型"].set(get_val("证件类型"))
            
            # 证件号码
            self.entries["证件号码"].delete(0, tk.END)
            self.entries["证件号码"].insert(0, get_val("证件号码"))
            
            # 车辆号码
            self.entries["车辆号码"].delete(0, tk.END)
            self.entries["车辆号码"].insert(0, get_val("车辆号码"))
            
            # 审批人学工号
            self.entries["审批人学工号"].delete(0, tk.END)
            self.entries["审批人学工号"].insert(0, get_val("审批人学工号"))
            
            # 审批人姓名
            self.entries["审批人姓名"].delete(0, tk.END)
            self.entries["审批人姓名"].insert(0, get_val("审批人姓名"))
            
            # 场所名称（复选框 + 文本框）
            locations_str = get_val("场所名称")
            locations = locations_str.split('@') if locations_str else []
            
            # 重置复选框
            for name, var in self.location_vars.items():
                var.set(name in locations)
            
            # 处理文本框内容（非预设选项的内容）
            predefined_locations = set(self.location_vars.keys())
            other_locations = [loc for loc in locations if loc and loc not in predefined_locations]
            self.location_text_entry.delete(0, tk.END)
            if other_locations:
                self.location_text_entry.insert(0, "@".join(other_locations))
                
            # 拜访人及事由
            self.entries["拜访人及事由"].delete("1.0", tk.END)
            self.entries["拜访人及事由"].insert("1.0", get_val("拜访人及事由"))
            
            # 处理时间字段
            for key, widgets in [("访问开始时间", self.entries["访问开始时间"]), ("访问结束时间", self.entries["访问结束时间"])]:
                try:
                    full_time_str = get_val(key, "2025-07-12 00:00")
                    if " " in full_time_str:
                        date_str, time_str = full_time_str.split(" ", 1)
                    else:
                        date_str = full_time_str if full_time_str else "2025-07-12"
                        time_str = "00:00"
                    
                    if ":" in time_str:
                        hour, minute = time_str.split(":", 1)
                    else:
                        hour, minute = "00", "00"
                    
                    date_entry, hour_spin, minute_spin = widgets
                    if date_str:
                        try:
                            date_entry.set_date(date_str)
                        except:
                            date_entry.set_date("2025-07-12")
                    hour_spin.set(hour.zfill(2))
                    minute_spin.set(minute.zfill(2))
                except Exception as e:
                    print(f"Error setting time for {key}: {e}")
                    # 设置默认值
                    date_entry, hour_spin, minute_spin = widgets
                    date_entry.set_date("2025-07-12")
                    hour_spin.set("00")
                    minute_spin.set("00")
                    
        except Exception as e:
            print(f"Error loading record: {e}")
            import traceback
            traceback.print_exc()
        
    def save_current_record(self):
        """Save the data from the UI form back to the current record."""
        if not self.data:
            return
            
        record = self.data[self.current_index]
        
        # Save values using standard keys
        record["访问形式"] = self.entries["访问形式"].get()
        record["访客姓名"] = self.entries["访客姓名"].get()
        record["手机号"] = self.entries["手机号"].get()
        record["证件类型"] = self.entries["证件类型"].get()
        record["证件号码"] = self.entries["证件号码"].get()
        record["车辆号码"] = self.entries["车辆号码"].get().replace(" ", "").upper()
        record["审批人学工号"] = self.entries["审批人学工号"].get()
        record["审批人姓名"] = self.entries["审批人姓名"].get()
        
        # 保存场所名称（复选框 + 文本框）
        selected_locations = [name for name, var in self.location_vars.items() if var.get()]
        other_location = self.location_text_entry.get().strip()
        if other_location:
            selected_locations.append(other_location)
        record["场所名称"] = "@".join(selected_locations)
        
        # Corrected the text widget get method
        record["拜访人及事由"] = self.entries["拜访人及事由"].get("1.0", tk.END).strip()

        for key, widgets in [("访问开始时间", self.entries["访问开始时间"]), ("访问结束时间", self.entries["访问结束时间"])]:
            date_entry, hour_spin, minute_spin = widgets
            date_str = date_entry.get_date().strftime('%Y-%m-%d')
            record[key] = f"{date_str} {hour_spin.get()}:{minute_spin.get()}"

    def validate_record(self, index):
        """Validate a specific record based on the rules."""
        errors = []
        record = self.data[index]

        if not record.get("访问形式"): errors.append("访问形式不能为空")
        if not record.get("访客姓名"): errors.append("访客姓名不能为空")
        if not record.get("证件类型"): errors.append("证件类型不能为空")
        if not record.get("证件号码"): errors.append("证件号码不能为空")
        if not record.get("审批人学工号"): errors.append("审批人学工号不能为空")
        if not record.get("审批人姓名"): errors.append("审批人姓名不能为空")
        if not record.get("场所名称"): errors.append("场所名称不能为空")
        if not record.get("拜访人及事由"): errors.append("拜访人及事由不能为空")

        phone = record.get("手机号", "")
        if not phone:
            errors.append("手机号不能为空")
        elif not (phone.isdigit() and len(phone) == 11):
            errors.append("手机号必须是11位数字")

        if errors:
            messagebox.showerror(f"记录 {index + 1} 验证错误", "\n".join(errors))
            return False
        return True

    def next_record(self):
        if not self.data: return
        self.save_current_record()
        if self.current_index < len(self.data) - 1:
            self.current_index += 1
            self.load_record(self.current_index)
            self.update_progress()

    def prev_record(self):
        if not self.data: return
        self.save_current_record()
        if self.current_index > 0:
            self.current_index -= 1
            self.load_record(self.current_index)
            self.update_progress()

    def jump_to_record(self):
        if not self.data: return
        try:
            target_num = int(self.jump_entry.get())
            if 1 <= target_num <= len(self.data):
                self.save_current_record()
                self.current_index = target_num - 1
                self.load_record(self.current_index)
                self.update_progress()
            else:
                messagebox.showerror("无效输入", f"请输入一个介于 1 和 {len(self.data)} 之间的数字。")
        except ValueError:
            messagebox.showerror("无效输入", "请输入一个有效的数字。")
        finally:
            self.jump_entry.delete(0, tk.END)

    def export_file(self):
        self.save_current_record()
        # for i in range(len(self.data)):
        #     if not self.validate_record(i):
        #         messagebox.showwarning("无法导出", f"记录 {i+1} 未通过验证，请修正后再尝试导出。")
        #         self.current_index = i
        #         self.load_record(self.current_index)
        #         self.update_progress()
        #         return
            
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV file", "*.csv"), ("Excel file", "*.xlsx"), ("JSON file", "*.json")]
        )
        if not path: return

        processed_data = []
        for record in self.data:
            new_rec = record.copy()
            new_rec["手机号"] = str(new_rec.get("手机号", "")) + "#"
            new_rec["证件号码"] = str(new_rec.get("证件号码", "")) + "#"
            new_rec["审批人学工号"] = str(new_rec.get("审批人学工号", "")) + "#"
            new_rec["访问开始时间"] = str(new_rec.get("访问开始时间", "")) + "#"
            new_rec["访问结束时间"] = str(new_rec.get("访问结束时间", "")) + "#"
            processed_data.append(new_rec)
            
        df = pd.DataFrame(processed_data)
        
        output_columns = [key.replace('*', '') for key in self.ui_fields]
        df = df[[col for col in output_columns if col in df.columns]]

        try:
            if path.endswith('.csv'):
                df.to_csv(path, index=False, encoding='utf-8-sig')
            elif path.endswith('.xlsx'):
                df.to_excel(path, index=False)
            elif path.endswith('.json'):
                df.to_json(path, orient='records', indent=4, force_ascii=False)
            
            messagebox.showinfo("成功", f"文件已成功导出到:\n{path}")
        except Exception as e:
            messagebox.showerror("导出错误", f"无法保存文件: {e}")

    def update_progress(self):
        if self.data:
            self.progress_label.config(text=f"进度: {self.current_index + 1} / {len(self.data)}")
        else:
            self.progress_label.config(text="进度: - / -")

    def update_ui_state(self, state):
        for widget in [self.export_button, self.prev_button, self.next_button, self.jump_button, self.jump_entry]:
            widget.config(state=state)

        for field_key, widget_or_group in self.entries.items():
            # Handle the Frame containing Checkbuttons and text entry
            if field_key == "场所名称":
                for child in widget_or_group.winfo_children():
                    if isinstance(child, ttk.Frame):
                        for grandchild in child.winfo_children():
                            try:
                                grandchild.config(state=state)
                            except tk.TclError:
                                pass
                    else:
                        try:
                            child.config(state=state)
                        except tk.TclError:
                            pass
            # Handle the tuple of time-related widgets
            elif isinstance(widget_or_group, tuple):
                for w in widget_or_group:
                    w.config(state=state)
            # Handle regular widgets (Entry, Combobox, Text)
            else:
                try:
                    widget_or_group.config(state=state)
                except tk.TclError:
                    # This can happen with tk.Text, it's safe to ignore
                    pass

if __name__ == "__main__":
    root = tk.Tk()
    app = DataProcessorApp(root)
    root.mainloop()
