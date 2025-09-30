"""
Main GUI Application for Excel Automation
Provides a user-friendly interface for creating and executing Excel workflows.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from typing import Optional, Dict, List
from workflow_engine import WorkflowEngine
from config import Config


class ExcelAutoGUI:
    """Main GUI application class"""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.config = Config()
        self.workflow_engine = WorkflowEngine()
        self.current_workflow = None
        self.current_file = None
        
        self.setup_window()
        self.create_menu()
        self.create_main_layout()
        self.load_default_workflows()
    
    def setup_window(self):
        """Setup main window properties"""
        self.root.title(self.config.get('app_name', 'Excel Auto'))
        window_size = self.config.get('window_size', '1000x700')
        self.root.geometry(window_size)
        
        # Center window on screen
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_menu(self):
        """Create menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="엑셀 열기", command=self.open_excel_file)
        file_menu.add_command(label="새 엑셀 파일", command=self.new_excel_file)
        file_menu.add_command(label="저장", command=self.save_excel_file)
        file_menu.add_command(label="다른 이름으로 저장", command=self.save_excel_file_as)
        file_menu.add_separator()
        
        # Recent files submenu
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="최근 파일", menu=self.recent_menu)
        self.update_recent_files_menu()
        
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.quit_app)
        
        # Workflow menu
        workflow_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="워크플로우", menu=workflow_menu)
        workflow_menu.add_command(label="새 워크플로우", command=self.new_workflow)
        workflow_menu.add_command(label="워크플로우 열기", command=self.load_workflow)
        workflow_menu.add_command(label="워크플로우 저장", command=self.save_workflow)
        workflow_menu.add_separator()
        workflow_menu.add_command(label="워크플로우 실행", command=self.execute_workflow)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="사용법", command=self.show_help)
        help_menu.add_command(label="정보", command=self.show_about)
    
    def create_main_layout(self):
        """Create main application layout"""
        # Create paned window for split view
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left panel - Workflow list and scenarios
        left_frame = ttk.Frame(main_paned)
        main_paned.add(left_frame, weight=1)
        
        # Workflow list
        ttk.Label(left_frame, text="워크플로우 목록", font=('Arial', 12, 'bold')).pack(pady=5)
        
        workflow_frame = ttk.Frame(left_frame)
        workflow_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(workflow_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.workflow_listbox = tk.Listbox(workflow_frame, yscrollcommand=scrollbar.set)
        self.workflow_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.workflow_listbox.yview)
        
        self.workflow_listbox.bind('<<ListboxSelect>>', self.on_workflow_select)
        
        # Workflow control buttons
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(btn_frame, text="실행", command=self.execute_workflow).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="편집", command=self.edit_workflow).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="삭제", command=self.delete_workflow).pack(side=tk.LEFT, padx=2)
        
        # Right panel - Workflow editor and output
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=2)
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab 1: Workflow editor
        editor_frame = ttk.Frame(self.notebook)
        self.notebook.add(editor_frame, text="워크플로우 편집기")
        
        # Workflow info
        info_frame = ttk.LabelFrame(editor_frame, text="워크플로우 정보")
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(info_frame, text="이름:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.workflow_name_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.workflow_name_var, width=40).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(info_frame, text="설명:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.workflow_desc_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.workflow_desc_var, width=40).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Steps list
        steps_frame = ttk.LabelFrame(editor_frame, text="작업 단계")
        steps_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Steps treeview
        columns = ('순서', '작업', '매개변수')
        self.steps_tree = ttk.Treeview(steps_frame, columns=columns, show='headings', height=10)
        
        for col in columns:
            self.steps_tree.heading(col, text=col)
            self.steps_tree.column(col, width=100)
        
        self.steps_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Step control buttons
        step_btn_frame = ttk.Frame(steps_frame)
        step_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(step_btn_frame, text="단계 추가", command=self.add_step_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(step_btn_frame, text="단계 제거", command=self.remove_step).pack(side=tk.LEFT, padx=2)
        ttk.Button(step_btn_frame, text="위로", command=self.move_step_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(step_btn_frame, text="아래로", command=self.move_step_down).pack(side=tk.LEFT, padx=2)
        
        # Tab 2: Scenarios (predefined workflows)
        scenario_frame = ttk.Frame(self.notebook)
        self.notebook.add(scenario_frame, text="시나리오")
        
        ttk.Label(scenario_frame, text="사전 정의된 시나리오", font=('Arial', 12, 'bold')).pack(pady=10)
        
        # Scenario buttons
        scenarios = [
            ("데이터 병합", "여러 엑셀 파일을 하나로 병합", self.scenario_merge_files),
            ("데이터 정렬", "특정 열 기준으로 데이터 정렬", self.scenario_sort_data),
            ("데이터 필터링", "조건에 맞는 데이터만 추출", self.scenario_filter_data),
            ("합계 계산", "열의 합계 계산 및 추가", self.scenario_calculate_sum),
            ("서식 적용", "헤더에 서식 적용", self.scenario_apply_formatting),
            ("시트 복사", "시트를 다른 파일로 복사", self.scenario_copy_sheet),
        ]
        
        for name, desc, command in scenarios:
            frame = ttk.Frame(scenario_frame)
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            ttk.Button(frame, text=name, command=command, width=20).pack(side=tk.LEFT, padx=5)
            ttk.Label(frame, text=desc).pack(side=tk.LEFT, padx=5)
        
        # Tab 3: Output/Log
        output_frame = ttk.Frame(self.notebook)
        self.notebook.add(output_frame, text="출력/로그")
        
        ttk.Label(output_frame, text="실행 결과", font=('Arial', 12, 'bold')).pack(pady=5)
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=20, width=80)
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Clear button
        ttk.Button(output_frame, text="로그 지우기", command=self.clear_output).pack(pady=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("준비")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def update_recent_files_menu(self):
        """Update recent files menu"""
        self.recent_menu.delete(0, tk.END)
        recent_files = self.config.get_recent_files()
        
        if not recent_files:
            self.recent_menu.add_command(label="(없음)", state=tk.DISABLED)
        else:
            for file_path in recent_files:
                self.recent_menu.add_command(
                    label=os.path.basename(file_path),
                    command=lambda f=file_path: self.open_excel_file(f)
                )
    
    def log_output(self, message: str):
        """Add message to output log"""
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
    
    def clear_output(self):
        """Clear output log"""
        self.output_text.delete(1.0, tk.END)
    
    def update_status(self, message: str):
        """Update status bar"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    # File operations
    def open_excel_file(self, file_path: Optional[str] = None):
        """Open an Excel file"""
        if not file_path:
            file_path = filedialog.askopenfilename(
                title="엑셀 파일 열기",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
        
        if file_path:
            if self.workflow_engine.excel_ops.open_file(file_path):
                self.current_file = file_path
                self.config.add_recent_file(file_path)
                self.update_recent_files_menu()
                self.update_status(f"파일 열림: {os.path.basename(file_path)}")
                self.log_output(f"파일을 열었습니다: {file_path}")
                messagebox.showinfo("성공", "파일을 성공적으로 열었습니다.")
            else:
                messagebox.showerror("오류", "파일을 열 수 없습니다.")
    
    def new_excel_file(self):
        """Create a new Excel file"""
        if self.workflow_engine.excel_ops.create_new_file():
            self.current_file = None
            self.update_status("새 파일 생성됨")
            self.log_output("새 엑셀 파일을 생성했습니다.")
            messagebox.showinfo("성공", "새 엑셀 파일을 생성했습니다.")
        else:
            messagebox.showerror("오류", "파일을 생성할 수 없습니다.")
    
    def save_excel_file(self):
        """Save current Excel file"""
        if not self.current_file:
            self.save_excel_file_as()
            return
        
        if self.workflow_engine.excel_ops.save_file(self.current_file):
            self.update_status(f"파일 저장됨: {os.path.basename(self.current_file)}")
            self.log_output(f"파일을 저장했습니다: {self.current_file}")
            messagebox.showinfo("성공", "파일을 저장했습니다.")
        else:
            messagebox.showerror("오류", "파일을 저장할 수 없습니다.")
    
    def save_excel_file_as(self):
        """Save Excel file with new name"""
        file_path = filedialog.asksaveasfilename(
            title="다른 이름으로 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            if self.workflow_engine.excel_ops.save_file(file_path):
                self.current_file = file_path
                self.config.add_recent_file(file_path)
                self.update_recent_files_menu()
                self.update_status(f"파일 저장됨: {os.path.basename(file_path)}")
                self.log_output(f"파일을 저장했습니다: {file_path}")
                messagebox.showinfo("성공", "파일을 저장했습니다.")
            else:
                messagebox.showerror("오류", "파일을 저장할 수 없습니다.")
    
    def quit_app(self):
        """Quit application"""
        if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
            self.root.quit()
    
    # Workflow operations
    def load_default_workflows(self):
        """Load default workflows"""
        # Create default workflows directory if not exists
        workflow_dir = self.config.get('default_workflow_dir', './workflows')
        if not os.path.exists(workflow_dir):
            os.makedirs(workflow_dir)
        
        # Load workflows from directory
        for file_name in os.listdir(workflow_dir):
            if file_name.endswith('.json'):
                file_path = os.path.join(workflow_dir, file_name)
                self.workflow_engine.load_workflow(file_path)
        
        self.update_workflow_list()
    
    def update_workflow_list(self):
        """Update workflow listbox"""
        self.workflow_listbox.delete(0, tk.END)
        for workflow_name in self.workflow_engine.get_workflow_list():
            self.workflow_listbox.insert(tk.END, workflow_name)
    
    def on_workflow_select(self, event):
        """Handle workflow selection"""
        selection = self.workflow_listbox.curselection()
        if selection:
            workflow_name = self.workflow_listbox.get(selection[0])
            self.current_workflow = workflow_name
            self.load_workflow_to_editor(workflow_name)
    
    def load_workflow_to_editor(self, workflow_name: str):
        """Load workflow details to editor"""
        workflow_info = self.workflow_engine.get_workflow_info(workflow_name)
        if workflow_info:
            self.workflow_name_var.set(workflow_info['name'])
            self.workflow_desc_var.set(workflow_info.get('description', ''))
            
            # Clear and populate steps tree
            for item in self.steps_tree.get_children():
                self.steps_tree.delete(item)
            
            for i, step in enumerate(workflow_info['steps'], 1):
                params_str = ', '.join([f"{k}={v}" for k, v in step['params'].items()])
                self.steps_tree.insert('', tk.END, values=(i, step['operation'], params_str))
    
    def new_workflow(self):
        """Create a new workflow"""
        dialog = tk.Toplevel(self.root)
        dialog.title("새 워크플로우")
        dialog.geometry("400x200")
        
        ttk.Label(dialog, text="워크플로우 이름:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=30).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="설명:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        desc_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=desc_var, width=30).grid(row=1, column=1, padx=10, pady=10)
        
        def create():
            name = name_var.get().strip()
            desc = desc_var.get().strip()
            
            if not name:
                messagebox.showerror("오류", "워크플로우 이름을 입력하세요.")
                return
            
            if self.workflow_engine.create_workflow(name, desc):
                self.update_workflow_list()
                self.log_output(f"새 워크플로우 생성: {name}")
                dialog.destroy()
                messagebox.showinfo("성공", "워크플로우를 생성했습니다.")
            else:
                messagebox.showerror("오류", "워크플로우를 생성할 수 없습니다.")
        
        ttk.Button(dialog, text="생성", command=create).grid(row=2, column=0, columnspan=2, pady=20)
    
    def load_workflow(self):
        """Load workflow from file"""
        file_path = filedialog.askopenfilename(
            title="워크플로우 열기",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            if self.workflow_engine.load_workflow(file_path):
                self.update_workflow_list()
                self.log_output(f"워크플로우를 불러왔습니다: {file_path}")
                messagebox.showinfo("성공", "워크플로우를 불러왔습니다.")
            else:
                messagebox.showerror("오류", "워크플로우를 불러올 수 없습니다.")
    
    def save_workflow(self):
        """Save current workflow to file"""
        if not self.current_workflow:
            messagebox.showwarning("경고", "저장할 워크플로우를 선택하세요.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="워크플로우 저장",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            if self.workflow_engine.save_workflow(self.current_workflow, file_path):
                self.log_output(f"워크플로우를 저장했습니다: {file_path}")
                messagebox.showinfo("성공", "워크플로우를 저장했습니다.")
            else:
                messagebox.showerror("오류", "워크플로우를 저장할 수 없습니다.")
    
    def execute_workflow(self):
        """Execute selected workflow"""
        if not self.current_workflow:
            messagebox.showwarning("경고", "실행할 워크플로우를 선택하세요.")
            return
        
        self.clear_output()
        self.log_output(f"워크플로우 실행 시작: {self.current_workflow}")
        self.update_status("워크플로우 실행 중...")
        
        def progress_callback(step, total, result):
            self.log_output(f"단계 {step}/{total}: {result.get('operation', '')} - {'성공' if result.get('success') else '실패'}")
            if not result.get('success'):
                self.log_output(f"  오류: {result.get('message', '')}")
        
        results = self.workflow_engine.execute_workflow(self.current_workflow, progress_callback)
        
        success_count = sum(1 for r in results if r.get('success'))
        self.log_output(f"\n워크플로우 실행 완료: {success_count}/{len(results)} 성공")
        self.update_status("워크플로우 실행 완료")
        
        messagebox.showinfo("완료", f"워크플로우 실행이 완료되었습니다.\n성공: {success_count}/{len(results)}")
    
    def edit_workflow(self):
        """Edit selected workflow"""
        if not self.current_workflow:
            messagebox.showwarning("경고", "편집할 워크플로우를 선택하세요.")
            return
        
        self.notebook.select(0)  # Switch to editor tab
    
    def delete_workflow(self):
        """Delete selected workflow"""
        if not self.current_workflow:
            messagebox.showwarning("경고", "삭제할 워크플로우를 선택하세요.")
            return
        
        if messagebox.askyesno("확인", f"'{self.current_workflow}' 워크플로우를 삭제하시겠습니까?"):
            if self.workflow_engine.delete_workflow(self.current_workflow):
                self.current_workflow = None
                self.update_workflow_list()
                self.log_output("워크플로우를 삭제했습니다.")
                messagebox.showinfo("성공", "워크플로우를 삭제했습니다.")
            else:
                messagebox.showerror("오류", "워크플로우를 삭제할 수 없습니다.")
    
    def add_step_dialog(self):
        """Show dialog to add a workflow step"""
        if not self.current_workflow:
            messagebox.showwarning("경고", "워크플로우를 먼저 선택하거나 생성하세요.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("단계 추가")
        dialog.geometry("500x400")
        
        ttk.Label(dialog, text="작업 선택:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        
        operations = list(self.workflow_engine.operations.keys())
        operation_var = tk.StringVar()
        operation_combo = ttk.Combobox(dialog, textvariable=operation_var, values=operations, state='readonly', width=30)
        operation_combo.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="매개변수 (JSON):").grid(row=1, column=0, padx=10, pady=10, sticky=tk.NW)
        params_text = tk.Text(dialog, height=15, width=40)
        params_text.grid(row=1, column=1, padx=10, pady=10)
        params_text.insert(1.0, "{}")
        
        def add():
            operation = operation_var.get()
            if not operation:
                messagebox.showerror("오류", "작업을 선택하세요.")
                return
            
            try:
                import json
                params = json.loads(params_text.get(1.0, tk.END))
                
                if self.workflow_engine.add_step(self.current_workflow, operation, params):
                    self.load_workflow_to_editor(self.current_workflow)
                    self.log_output(f"단계 추가: {operation}")
                    dialog.destroy()
                    messagebox.showinfo("성공", "단계를 추가했습니다.")
                else:
                    messagebox.showerror("오류", "단계를 추가할 수 없습니다.")
            except json.JSONDecodeError:
                messagebox.showerror("오류", "올바른 JSON 형식이 아닙니다.")
        
        ttk.Button(dialog, text="추가", command=add).grid(row=2, column=0, columnspan=2, pady=10)
    
    def remove_step(self):
        """Remove selected step from workflow"""
        selection = self.steps_tree.selection()
        if not selection:
            messagebox.showwarning("경고", "제거할 단계를 선택하세요.")
            return
        
        # This is a simplified version - in a full implementation, 
        # you'd modify the workflow engine to support step removal
        messagebox.showinfo("정보", "이 기능은 개발 중입니다.")
    
    def move_step_up(self):
        """Move selected step up"""
        messagebox.showinfo("정보", "이 기능은 개발 중입니다.")
    
    def move_step_down(self):
        """Move selected step down"""
        messagebox.showinfo("정보", "이 기능은 개발 중입니다.")
    
    # Scenario methods
    def scenario_merge_files(self):
        """Scenario: Merge multiple Excel files"""
        file_paths = filedialog.askopenfilenames(
            title="병합할 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if not file_paths:
            return
        
        # Create workflow for merging
        workflow_name = "파일_병합_" + str(len(self.workflow_engine.workflows) + 1)
        self.workflow_engine.create_workflow(workflow_name, "여러 파일 병합")
        
        self.workflow_engine.add_step(workflow_name, 'open_file', {'file_path': file_paths[0]})
        self.workflow_engine.add_step(workflow_name, 'merge_files', {'file_paths': list(file_paths)})
        
        save_path = filedialog.asksaveasfilename(
            title="병합된 파일 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if save_path:
            self.workflow_engine.add_step(workflow_name, 'save_file', {'file_path': save_path})
            self.current_workflow = workflow_name
            self.update_workflow_list()
            self.execute_workflow()
    
    def scenario_sort_data(self):
        """Scenario: Sort data by column"""
        if not self.current_file:
            messagebox.showwarning("경고", "먼저 엑셀 파일을 여세요.")
            return
        
        # Simple dialog for column name
        column = tk.simpledialog.askstring("정렬", "정렬할 열 이름을 입력하세요:")
        if column:
            workflow_name = "데이터_정렬_" + str(len(self.workflow_engine.workflows) + 1)
            self.workflow_engine.create_workflow(workflow_name, "데이터 정렬")
            self.workflow_engine.add_step(workflow_name, 'sort_data', {'column': column, 'ascending': True})
            
            self.current_workflow = workflow_name
            self.update_workflow_list()
            self.execute_workflow()
    
    def scenario_filter_data(self):
        """Scenario: Filter data"""
        messagebox.showinfo("정보", "이 시나리오는 개발 중입니다.")
    
    def scenario_calculate_sum(self):
        """Scenario: Calculate sum of column"""
        messagebox.showinfo("정보", "이 시나리오는 개발 중입니다.")
    
    def scenario_apply_formatting(self):
        """Scenario: Apply formatting to header"""
        if not self.current_file:
            messagebox.showwarning("경고", "먼저 엑셀 파일을 여세요.")
            return
        
        workflow_name = "서식_적용_" + str(len(self.workflow_engine.workflows) + 1)
        self.workflow_engine.create_workflow(workflow_name, "헤더 서식 적용")
        
        # Apply formatting to first row (header)
        for col in range(1, 11):  # First 10 columns
            self.workflow_engine.add_step(
                workflow_name, 
                'format_cell', 
                {'row': 1, 'col': col, 'bold': True, 'bg_color': 'CCCCCC', 'align': 'center'}
            )
        
        self.current_workflow = workflow_name
        self.update_workflow_list()
        self.execute_workflow()
    
    def scenario_copy_sheet(self):
        """Scenario: Copy sheet to another file"""
        messagebox.showinfo("정보", "이 시나리오는 개발 중입니다.")
    
    # Help and about
    def show_help(self):
        """Show help dialog"""
        help_text = """
Excel Auto - 사용법

1. 파일 작업:
   - 엑셀 파일 열기: 파일 > 엑셀 열기
   - 새 파일 생성: 파일 > 새 엑셀 파일
   - 저장: 파일 > 저장

2. 워크플로우:
   - 새 워크플로우 생성: 워크플로우 > 새 워크플로우
   - 워크플로우 편집: 워크플로우 목록에서 선택 후 편집 탭 사용
   - 워크플로우 실행: 워크플로우 선택 후 실행 버튼

3. 시나리오:
   - 시나리오 탭에서 사전 정의된 작업 선택
   - 버튼 클릭으로 빠른 실행

4. 작업 단계:
   - 워크플로우에 여러 작업을 순차적으로 추가
   - 각 작업은 JSON 형식의 매개변수 사용
        """
        
        dialog = tk.Toplevel(self.root)
        dialog.title("사용법")
        dialog.geometry("600x400")
        
        text = scrolledtext.ScrolledText(dialog, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text.insert(1.0, help_text)
        text.config(state=tk.DISABLED)
        
        ttk.Button(dialog, text="닫기", command=dialog.destroy).pack(pady=10)
    
    def show_about(self):
        """Show about dialog"""
        about_text = f"""
{self.config.get('app_name', 'Excel Auto')}
버전: {self.config.get('version', '1.0.0')}

엑셀 자동화 워크플로우 GUI 애플리케이션

기능:
- 엑셀 파일 읽기/쓰기
- 사용자 정의 워크플로우
- 다양한 시나리오 지원
- 데이터 처리 및 분석

개발: Python + tkinter + openpyxl
        """
        
        messagebox.showinfo("정보", about_text)


def main():
    """Main entry point"""
    root = tk.Tk()
    app = ExcelAutoGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
