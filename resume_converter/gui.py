"""
简历标准化工具的Windows桌面GUI入口。

窗口只负责路径、模板和进度交互；实际业务继续复用ResumeConverter。
"""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
import ctypes
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from typing import Any

from gui_support import GuiBatchWorker, WorkerEvent
from main import TEMPLATE_OPTIONS, TemplateOption, _find_duplicate_stems
from utils.restore_markdown import SUPPORTED_SUFFIXES
from dotenv import load_dotenv

from utils.runtime_paths import application_environment_path


APP_TITLE = "简历标准化处理工具"
PREFERRED_SIZE = (900, 700)
MINIMUM_SIZE = (740, 560)
PRIMARY_COLOR = "#2878D0"
PRIMARY_DARK = "#1F63AD"
STOP_COLOR = "#E58A2B"
STOP_DARK = "#C97316"
DISABLED_COLOR = "#90A4AE"
BACKGROUND_COLOR = "#F5F7FA"
TEXT_COLOR = "#263238"
MUTED_COLOR = "#607D8B"
REQUIRED_ENVIRONMENT = (
    "BAIDU_API_KEY",
    "BAIDU_SECRET_KEY",
    "LLM_BASE_URL",
    "LLM_MODEL",
    "LLM_API_KEY",
)


def _selected_files_summary(files: tuple[Path, ...]) -> str:
    """为只读输入框生成简洁、可辨识的文件选择摘要。"""

    if not files:
        return "尚未选择文件"
    if len(files) == 1:
        return str(files[0])
    preview = "；".join(path.name for path in files[:3])
    if len(files) > 3:
        preview += "；……"
    return f"已选择 {len(files)} 份：{preview}"


def _validate_selected_files(files: tuple[Path, ...]) -> list[Path]:
    """校验GUI明确选中的文件，并阻止格式或同stem冲突。"""

    if not files:
        raise ValueError("请选择至少一份原始简历。")

    resolved_files = [path.resolve() for path in files]
    for path in resolved_files:
        if not path.exists() or not path.is_file():
            raise ValueError(f"找不到已选择的文件：{path}")
        if path.name.startswith("~$"):
            raise ValueError(f"不能处理Word/WPS临时文件：{path.name}")
        if path.suffix.lower() not in SUPPORTED_SUFFIXES:
            raise ValueError(
                f"不支持的文件格式：{path.name}；"
                "支持PDF、DOCX、JPG、JPEG和PNG。"
            )

    duplicates = _find_duplicate_stems(resolved_files)
    if duplicates:
        duplicate_names = "\n".join(
            "、".join(path.name for path in paths)
            for paths in duplicates.values()
        )
        raise ValueError(
            "检测到文件名主体相同的简历，会覆盖同名结果：\n"
            f"{duplicate_names}\n请重命名后再处理。"
        )
    return resolved_files


class ResumeConverterApp:
    """管理简洁办公风窗口及后台批处理生命周期。"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.event_queue: queue.Queue[WorkerEvent] = queue.Queue()
        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.running = False
        self.close_after_stop = False
        self.output_directory: Path | None = None
        self.access_dialog: tk.Toplevel | None = None
        self.access_response: queue.Queue[str] | None = None
        self.selected_files: tuple[Path, ...] = ()

        self.input_var = tk.StringVar(value="尚未选择文件")
        self.output_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.current_file_var = tk.StringVar(value="尚未开始")
        self.stage_var = tk.StringVar(value="等待用户选择任务")
        self.progress_text_var = tk.StringVar(value="0 / 0")
        self.status_var = tk.StringVar(value="准备就绪")

        self.template_by_label = {
            option.label: option for option in TEMPLATE_OPTIONS.values()
        }

        self._configure_window()
        self._configure_styles()
        self._build_interface()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_events)

    def _configure_window(self) -> None:
        self.root.title(APP_TITLE)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = max(MINIMUM_SIZE[0], min(PREFERRED_SIZE[0], screen_width - 80))
        height = max(
            MINIMUM_SIZE[1],
            min(PREFERRED_SIZE[1], screen_height - 100),
        )
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2 - 10)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(
            min(MINIMUM_SIZE[0], screen_width - 40),
            min(MINIMUM_SIZE[1], screen_height - 80),
        )
        self.root.configure(background=BACKGROUND_COLOR)
        if screen_width <= 1366 or screen_height <= 768:
            # 低分辨率或高DPI办公设备上优先保证全部操作控件可见。
            self.root.state("zoomed")

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")

        style.configure("App.TFrame", background=BACKGROUND_COLOR)
        style.configure("Card.TFrame", background="#FFFFFF")
        style.configure(
            "Title.TLabel",
            background=BACKGROUND_COLOR,
            foreground=TEXT_COLOR,
            font=("Microsoft YaHei UI", 20, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=BACKGROUND_COLOR,
            foreground=MUTED_COLOR,
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "CardTitle.TLabel",
            background="#FFFFFF",
            foreground=TEXT_COLOR,
            font=("Microsoft YaHei UI", 11, "bold"),
        )
        style.configure(
            "Field.TLabel",
            background="#FFFFFF",
            foreground=TEXT_COLOR,
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Info.TLabel",
            background="#FFFFFF",
            foreground=MUTED_COLOR,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "Primary.TButton",
            font=("Microsoft YaHei UI", 10, "bold"),
            padding=(18, 8),
            foreground="#FFFFFF",
            background=PRIMARY_COLOR,
        )
        style.map(
            "Primary.TButton",
            background=[("active", PRIMARY_DARK), ("disabled", "#AAB8C2")],
        )
        style.configure(
            "Secondary.TButton",
            font=("Microsoft YaHei UI", 9),
            padding=(12, 6),
        )
        style.configure(
            "Horizontal.TProgressbar",
            background=PRIMARY_COLOR,
            troughcolor="#E7EDF3",
        )

    def _build_interface(self) -> None:
        shell = ttk.Frame(self.root, style="App.TFrame")
        shell.pack(fill="both", expand=True)
        canvas = tk.Canvas(
            shell,
            background=BACKGROUND_COLOR,
            highlightthickness=0,
            borderwidth=0,
        )
        scrollbar = ttk.Scrollbar(
            shell,
            orient="vertical",
            command=canvas.yview,
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        container = ttk.Frame(canvas, style="App.TFrame", padding=16)
        window_id = canvas.create_window(
            (0, 0),
            window=container,
            anchor="nw",
        )

        def update_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def fit_content_width(event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        def scroll_content(event) -> None:
            canvas.yview_scroll(int(-event.delta / 120), "units")

        container.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", fit_content_width)
        canvas.bind_all("<MouseWheel>", scroll_content)

        ttk.Label(
            container,
            text=APP_TITLE,
            style="Title.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            container,
            text=(
                "批量完成文档解析、信息提取和标准 Word 生成。"
                "支持 PDF、DOCX、JPG、JPEG、PNG。"
            ),
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(3, 10))

        settings = ttk.Frame(container, style="Card.TFrame", padding=14)
        settings.pack(fill="x")
        settings.columnconfigure(1, weight=1)
        ttk.Label(settings, text="任务设置", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 7)
        )
        self._add_file_row(
            settings,
            row=1,
            label="输入文件",
            command=self._select_input_files,
        )
        self._add_path_row(
            settings,
            row=2,
            label="输出文件夹",
            variable=self.output_var,
            command=self._select_output_directory,
        )

        ttk.Label(settings, text="模板模式", style="Field.TLabel").grid(
            row=3, column=0, sticky="w", padx=(0, 12), pady=7
        )
        self.template_box = ttk.Combobox(
            settings,
            textvariable=self.template_var,
            values=tuple(self.template_by_label),
            state="readonly",
            font=("Microsoft YaHei UI", 10),
        )
        self.template_box.grid(row=3, column=1, sticky="ew", pady=7)
        if self.template_by_label:
            self.template_box.current(0)

        ttk.Label(
            settings,
            text="可选择单个文件，或按住 Ctrl / Shift 同时选择多个文件。",
            style="Info.TLabel",
        ).grid(row=4, column=1, sticky="w", pady=(3, 0))

        actions = ttk.Frame(settings, style="Card.TFrame")
        actions.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        actions.columnconfigure(0, weight=1)
        self.open_button = ttk.Button(
            actions,
            text="打开输出文件夹",
            style="Secondary.TButton",
            command=self._open_output_directory,
            state="disabled",
        )
        self.open_button.grid(row=0, column=1, padx=(0, 10))
        # Windows的vista主题会覆盖ttk.Button背景色，因此主操作按钮使用
        # 原生Tk按钮，确保源码版和打包版始终呈现清晰的蓝底白字。
        self.start_button = tk.Button(
            actions,
            text="开始处理",
            command=self._start_processing,
            font=("Microsoft YaHei UI", 10, "bold"),
            foreground="#FFFFFF",
            background=PRIMARY_COLOR,
            activeforeground="#FFFFFF",
            activebackground=PRIMARY_DARK,
            disabledforeground="#F5F7FA",
            relief="flat",
            borderwidth=0,
            padx=20,
            pady=8,
            cursor="hand2",
        )
        self.start_button.grid(row=0, column=2)

        progress_card = ttk.Frame(container, style="Card.TFrame", padding=14)
        progress_card.pack(fill="x", pady=(10, 0))
        progress_card.columnconfigure(1, weight=1)
        ttk.Label(
            progress_card, text="处理进度", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            progress_card, textvariable=self.progress_text_var, style="Info.TLabel"
        ).grid(row=0, column=2, sticky="e")
        self.progress = ttk.Progressbar(
            progress_card,
            mode="determinate",
            maximum=1,
        )
        self.progress.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 12))
        ttk.Label(progress_card, text="当前文件", style="Field.TLabel").grid(
            row=2, column=0, sticky="nw", padx=(0, 12)
        )
        ttk.Label(
            progress_card,
            textvariable=self.current_file_var,
            style="Info.TLabel",
            wraplength=620,
        ).grid(row=2, column=1, columnspan=2, sticky="w")
        ttk.Label(progress_card, text="当前阶段", style="Field.TLabel").grid(
            row=3, column=0, sticky="nw", padx=(0, 12), pady=(8, 0)
        )
        ttk.Label(
            progress_card,
            textvariable=self.stage_var,
            style="Info.TLabel",
            wraplength=620,
        ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(8, 0))

        log_card = ttk.Frame(container, style="Card.TFrame", padding=14)
        log_card.pack(fill="both", expand=True, pady=(10, 0))
        ttk.Label(log_card, text="运行日志", style="CardTitle.TLabel").pack(
            anchor="w", pady=(0, 10)
        )
        self.log_text = ScrolledText(
            log_card,
            height=10,
            wrap="word",
            state="disabled",
            borderwidth=1,
            relief="solid",
            font=("Microsoft YaHei UI", 9),
            foreground=TEXT_COLOR,
            background="#FBFCFD",
            padx=10,
            pady=8,
        )
        self.log_text.pack(fill="both", expand=True)

        ttk.Label(
            container,
            textvariable=self.status_var,
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(9, 0))

    def _add_path_row(
        self,
        parent,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
    ) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=7
        )
        entry = ttk.Entry(
            parent,
            textvariable=variable,
            font=("Microsoft YaHei UI", 10),
        )
        entry.grid(row=row, column=1, sticky="ew", pady=7)
        ttk.Button(
            parent,
            text="选择文件夹",
            style="Secondary.TButton",
            command=command,
        ).grid(row=row, column=2, padx=(10, 0), pady=7)

    def _add_file_row(self, parent, row: int, label: str, command) -> None:
        """创建只读的多文件选择行。"""

        ttk.Label(parent, text=label, style="Field.TLabel").grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=7
        )
        ttk.Entry(
            parent,
            textvariable=self.input_var,
            state="readonly",
            font=("Microsoft YaHei UI", 10),
        ).grid(row=row, column=1, sticky="ew", pady=7)
        ttk.Button(
            parent,
            text="选择文件",
            style="Secondary.TButton",
            command=command,
        ).grid(row=row, column=2, padx=(10, 0), pady=7)

    def _select_input_files(self) -> None:
        """让用户通过Windows文件对话框选择一份或多份简历。"""

        initial_directory = (
            self.selected_files[0].parent
            if self.selected_files
            else Path.cwd()
        )
        selected = filedialog.askopenfilenames(
            title="选择一份或多份原始简历",
            initialdir=str(initial_directory),
            filetypes=(
                ("支持的简历文件", "*.pdf *.docx *.jpg *.jpeg *.png"),
                ("PDF 文件", "*.pdf"),
                ("Word 文档", "*.docx"),
                ("图片文件", "*.jpg *.jpeg *.png"),
                ("所有文件", "*.*"),
            ),
        )
        if not selected:
            return

        unique_files = {
            str(Path(value).resolve()).casefold(): Path(value).resolve()
            for value in selected
        }
        self.selected_files = tuple(
            sorted(unique_files.values(), key=lambda path: path.name.casefold())
        )
        self.input_var.set(_selected_files_summary(self.selected_files))
        self.output_var.set(
            str(self.selected_files[0].parent / "标准简历输出")
        )

    def _select_output_directory(self) -> None:
        initial = self.output_var.get().strip()
        if not initial and self.selected_files:
            initial = str(self.selected_files[0].parent)
        selected = filedialog.askdirectory(
            title="选择结果输出文件夹",
            initialdir=initial if initial else None,
            mustexist=False,
        )
        if selected:
            self.output_var.set(str(Path(selected).resolve()))

    def _start_processing(self) -> None:
        try:
            files, output_directory, template = self._validate_task()
        except (ValueError, OSError) as error:
            messagebox.showerror("无法开始处理", str(error), parent=self.root)
            return

        confirmation = (
            f"待处理简历：{len(files)} 份\n"
            f"模板模式：{template.label}\n\n"
            "确认开始处理吗？"
        )
        if not messagebox.askyesno("确认任务", confirmation, parent=self.root):
            return

        self.output_directory = output_directory
        self.stop_event.clear()
        self.running = True
        self.close_after_stop = False
        self.progress.configure(maximum=max(1, len(files)), value=0)
        self.progress_text_var.set(f"0 / {len(files)}")
        self.current_file_var.set("准备处理第一份简历")
        self.stage_var.set("正在启动后台任务")
        self.status_var.set("处理中，请勿关闭网络连接……")
        self._clear_log()
        self._append_log(f"已选择 {len(files)} 份待处理简历。")
        self._set_controls_running(True)

        worker = GuiBatchWorker(
            files=files,
            output_directory=output_directory,
            template=template,
            emit=self.event_queue.put,
            resolve_access=self._request_file_access,
            stop_event=self.stop_event,
        )
        self.worker_thread = threading.Thread(
            target=worker.run,
            name="resume-converter-worker",
            daemon=True,
        )
        self.worker_thread.start()

    def _validate_task(
        self,
    ) -> tuple[list[Path], Path, TemplateOption]:
        output_value = self.output_var.get().strip()
        if not self.selected_files:
            raise ValueError("请选择至少一份原始简历。")
        if not output_value:
            raise ValueError("请选择结果输出文件夹。")

        output_directory = Path(output_value).expanduser()
        if output_directory.exists() and not output_directory.is_dir():
            raise ValueError(f"输出路径不是文件夹：{output_directory}")

        files = _validate_selected_files(self.selected_files)

        template = self.template_by_label.get(self.template_var.get())
        if template is None:
            raise ValueError("请选择一个有效的模板模式。")

        env_path = application_environment_path()
        load_dotenv(dotenv_path=env_path)
        missing = [name for name in REQUIRED_ENVIRONMENT if not os.getenv(name)]
        if missing:
            raise ValueError(
                f"程序目录缺少服务配置：{env_path}\n"
                f"请确认以下配置项已填写：{', '.join(missing)}"
            )

        output_directory.mkdir(parents=True, exist_ok=True)
        probe = output_directory / ".resume_converter_write_test"
        try:
            probe.write_text("ok", encoding="utf-8")
        finally:
            if probe.exists():
                probe.unlink()
        return files, output_directory.resolve(), template

    def _request_file_access(self, path: Path, error_message: str) -> str:
        response: queue.Queue[str] = queue.Queue(maxsize=1)
        self.event_queue.put(
            WorkerEvent(
                "access_required",
                {
                    "path": path,
                    "error": error_message,
                    "response": response,
                },
            )
        )
        return response.get()

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        if self.root.winfo_exists():
            self.root.after(100, self._poll_events)

    def _handle_event(self, event: WorkerEvent) -> None:
        payload = event.payload
        if event.kind == "file_started":
            self.current_file_var.set(
                f"[{payload['index']}/{payload['total']}] {payload['path'].name}"
            )
            self.stage_var.set("正在处理当前简历")
            self._append_log(
                f"\n开始处理 [{payload['index']}/{payload['total']}]："
                f"{payload['path'].name}"
            )
        elif event.kind == "log":
            message = payload["message"]
            self._append_log(message)
            if "[阶段" in message:
                self.stage_var.set(message.strip())
        elif event.kind == "file_finished":
            self.progress.configure(value=payload["index"])
            self.progress_text_var.set(
                f"{payload['index']} / {payload['total']}"
            )
            labels = {
                "succeeded": "完成",
                "failed": "失败",
                "skipped": "跳过",
            }
            self._append_log(
                f"[{labels[payload['status']]}] {payload['path'].name}"
            )
        elif event.kind == "access_required":
            self._show_access_dialog(payload)
        elif event.kind == "completed":
            self._finish_processing(payload["result"])

    def _show_access_dialog(self, payload: dict[str, Any]) -> None:
        dialog = tk.Toplevel(self.root)
        self.access_dialog = dialog
        self.access_response = payload["response"]
        dialog.title("文件暂时无法读取")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text=f"文件：{payload['path'].name}",
            font=("Microsoft YaHei UI", 11, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            frame,
            text=(
                "该文件可能正被其他程序独占。请关闭占用软件后重试，"
                "也可以跳过当前文件或安全结束整批任务。"
            ),
            wraplength=440,
            justify="left",
        ).pack(anchor="w", pady=(10, 5))
        if payload.get("error"):
            ttk.Label(
                frame,
                text=f"系统提示：{payload['error']}",
                foreground=MUTED_COLOR,
                wraplength=440,
            ).pack(anchor="w", pady=(0, 14))

        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e")

        def answer(value: str) -> None:
            payload["response"].put(value)
            self.access_dialog = None
            self.access_response = None
            dialog.destroy()

        ttk.Button(buttons, text="结束任务", command=lambda: answer("stop")).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(buttons, text="跳过", command=lambda: answer("skip")).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(
            buttons,
            text="重试",
            style="Primary.TButton",
            command=lambda: answer("retry"),
        ).pack(side="left")

        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _finish_processing(self, result) -> None:
        self.running = False
        self._set_controls_running(False)
        self.open_button.configure(state="normal")
        self.stage_var.set("任务已结束")
        self.status_var.set("处理完成，可以查看最终简历和待补充信息。")
        summary = (
            f"总计：{result.total} 份\n"
            f"成功：{len(result.succeeded)} 份\n"
            f"失败：{len(result.failed)} 份\n"
            f"跳过：{len(result.skipped)} 份"
        )
        if result.stopped_early:
            summary += "\n\n任务已按要求安全停止。"
        self._append_log("\n" + summary.replace("\n", "；"))

        if self.close_after_stop:
            self.root.destroy()
            return
        messagebox.showinfo("批量处理结束", summary, parent=self.root)

    def _set_controls_running(self, running: bool) -> None:
        self.start_button.configure(
            state="normal",
            text="当前文件完成后停止" if running else "开始处理",
            command=self._request_stop if running else self._start_processing,
            background=STOP_COLOR if running else PRIMARY_COLOR,
            activebackground=STOP_DARK if running else PRIMARY_DARK,
        )
        self.template_box.configure(state="disabled" if running else "readonly")

    def _request_stop(self) -> None:
        if not self.running or self.stop_event.is_set():
            return
        if messagebox.askyesno(
            "安全停止任务",
            "当前文件会继续处理，完成后不再处理剩余文件。确定停止吗？",
            parent=self.root,
        ):
            self.stop_event.set()
            self.start_button.configure(
                state="disabled",
                text="正在安全停止……",
                background=DISABLED_COLOR,
            )
            self.status_var.set("已请求停止，等待当前文件处理完成……")

    def _on_close(self) -> None:
        if not self.running:
            self.root.destroy()
            return
        if messagebox.askyesno(
            "任务仍在处理中",
            "将在当前文件处理完成后安全停止并关闭窗口，确定退出吗？",
            parent=self.root,
        ):
            self.close_after_stop = True
            self.stop_event.set()
            if self.access_response is not None:
                self.access_response.put("stop")
                self.access_response = None
            if self.access_dialog is not None:
                self.access_dialog.destroy()
                self.access_dialog = None
            self.root.withdraw()

    def _open_output_directory(self) -> None:
        if self.output_directory and self.output_directory.exists():
            os.startfile(self.output_directory)
        else:
            messagebox.showwarning(
                "输出目录不存在",
                "尚未生成输出目录。",
                parent=self.root,
            )

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message.rstrip() + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")


def _enable_windows_dpi_awareness() -> None:
    """避免Windows再次缩放Tk窗口，导致高DPI屏幕裁切控件。"""

    if os.name != "nt":
        return
    try:
        # Per Monitor V2；旧版Windows不支持时回退到系统DPI感知。
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
    except (AttributeError, OSError):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass


def main() -> None:
    """启动Windows桌面GUI。"""

    _enable_windows_dpi_awareness()
    root = tk.Tk()
    ResumeConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
