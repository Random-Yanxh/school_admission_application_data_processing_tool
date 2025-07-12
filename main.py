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
                self.location_vars = {
                    "东区": tk.BooleanVar(), "西区": tk.BooleanVar(),
                    "北区": tk.BooleanVar(), "梅山校区": tk.BooleanVar()
                }
                for name, var in self.location_vars.items():
                    ttk.Checkbutton(location_frame, text=name, variable=var).pack(side=tk.LEFT, padx=5)
                self.entries[field_key] = location_frame
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
            
            if field_key != "场所名称":
                self.entries[field_key].grid(row=i, column=1, columnspan=2, sticky=tk.W, padx=5)


    def import_file(self):
        """
        Handles file import, standardizes column names, and pre-processes data.
        """
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls;*.xlsx"), ("CSV files", "*.csv")]
        )
        if not path:
            return

        try:
            if path.endswith('.csv'):
                df = pd.read_csv(path, dtype=str)
            else:
                df = pd.read_excel(path, dtype=str)
            
            df = df.fillna('')

            # --- Data Standardization ---
            standard_keys = [key.replace('*', '') for key in self.ui_fields]
            
            # Create a mapping from cleaned original column names to their original form
            col_mapping = {col.strip().replace('*', ''): col for col in df.columns}
            
            # Create a new DataFrame with standardized columns
            new_df = pd.DataFrame()
            for key in standard_keys:
                original_col = col_mapping.get(key)
                if original_col:
                    new_df[key] = df[original_col]
                else:
                    new_df[key] = '' # Add missing columns as empty

            # --- Pre-processing ---
            # 1. Uppercase vehicle numbers
            if "车辆号码" in new_df.columns:
                new_df["车辆号码"] = new_df["车辆号码"].str.replace(" ", "").str.upper()

            self.data = new_df.to_dict('records')
            self.file_path = path
            self.current_index = 0
            
            if not self.data:
                messagebox.showwarning("警告", "文件为空，没有数据可处理。")
                return

            self.file_label.config(text=f"已加载: {path.split('/')[-1]}")
            self.load_record(self.current_index)
            self.update_ui_state("normal")
            self.update_progress()
            self.root.update_idletasks()

        except Exception as e:
            messagebox.showerror("导入错误", f"无法读取文件: {e}")

    def load_record(self, index):
        """Load data from a specific record index into the UI form."""
        if not self.data:
            return
        
        record = self.data[index]
        
        def get_val(key, default=''):
            """Gets a value from the record, handling potential missing keys."""
            return str(record.get(key, default))

        # Populate form fields using the standard keys
        self.entries["访问形式"].set(get_val("访问形式"))
        self.entries["访客姓名"].delete(0, tk.END)
        self.entries["访客姓名"].insert(0, get_val("访客姓名"))
        self.entries["手机号"].delete(0, tk.END)
        self.entries["手机号"].insert(0, get_val("手机号"))
        self.entries["证件类型"].set(get_val("证件类型"))
        self.entries["证件号码"].delete(0, tk.END)
        self.entries["证件号码"].insert(0, get_val("证件号码"))
        self.entries["车辆号码"].delete(0, tk.END)
        # The .upper() is no longer needed here as it's done during import
        self.entries["车辆号码"].insert(0, get_val("车辆号码"))
        
        self.entries["审批人学工号"].delete(0, tk.END)
        self.entries["审批人学工号"].insert(0, get_val("审批人学工号"))
        self.entries["审批人姓名"].delete(0, tk.END)
        self.entries["审批人姓名"].insert(0, get_val("审批人姓名"))
        
        locations = get_val("场所名称").split('@')
        for name, var in self.location_vars.items():
            var.set(name in locations)
            
        self.entries["拜访人及事由"].delete("1.0", tk.END)
        self.entries["拜访人及事由"].insert("1.0", get_val("拜访人及事由"))
        
        # Handle time fields
        for key, widgets in [("访问开始时间", self.entries["访问开始时间"]), ("访问结束时间", self.entries["访问结束时间"])]:
            try:
                full_time_str = get_val(key, " ")
                date_str, time_str = (full_time_str.split(" ") + ["00:00"])[:2]
                hour, minute = (time_str.split(":") + ["00", "00"])[:2]
                
                date_entry, hour_spin, minute_spin = widgets
                if date_str: date_entry.set_date(date_str)
                hour_spin.set(hour)
                minute_spin.set(minute)
            except (ValueError, IndexError):
                continue # Ignore if date/time is malformed
        
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
        
        selected_locations = [name for name, var in self.location_vars.items() if var.get()]
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

        # Since data is standardized, we can directly access keys
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
        for i in range(len(self.data)):
            if not self.validate_record(i):
                messagebox.showwarning("无法导出", f"记录 {i+1} 未通过验证，请修正后再尝试导出。")
                self.current_index = i
                self.load_record(self.current_index)
                self.update_progress()
                return
            
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV file", "*.csv"), ("Excel file", "*.xlsx"), ("JSON file", "*.json")]
        )
        if not path: return
        
        # Create a fresh DataFrame from the processed self.data
        export_df = pd.DataFrame(self.data)

        # Apply export-time formatting
        export_df["手机号"] = export_df["手机号"].astype(str) + "#"
        export_df["证件号码"] = export_df["证件号码"].astype(str) + "#"
        export_df["审批人学工号"] = export_df["审批人学工号"].astype(str) + "#"
        export_df["访问开始时间"] = export_df["访问开始时间"].astype(str) + "#"
        export_df["访问结束时间"] = export_df["访问结束时间"].astype(str) + "#"
        
        # Ensure column order matches the UI
        output_columns = [key.replace('*', '') for key in self.ui_fields]
        export_df = export_df[output_columns]


        try:
            if path.endswith('.csv'):
                export_df.to_csv(path, index=False, encoding='utf-8-sig')
            elif path.endswith('.xlsx'):
                export_df.to_excel(path, index=False)
            elif path.endswith('.json'):
                export_df.to_json(path, orient='records', indent=4, force_ascii=False)
            
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
            # Handle the Frame containing Checkbuttons
            if field_key == "场所名称":
                for child in widget_or_group.winfo_children():
                    child.config(state=state)
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
