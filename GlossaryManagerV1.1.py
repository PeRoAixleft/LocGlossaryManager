import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass
from typing import Dict, List


@dataclass
@dataclass
class Term:
    """术语类，用于存储单个术语的所有信息"""
    term: str  # 原文
    translation: str  # 译文
    category: str = ""  # 分类
    context: str = ""  # 上下文
    notes: str = ""  # 备注
    created_at: str = ""  # 创建时间
    flags: set = None  # 术语标记
    history: list = None  # 历史记录

    def __post_init__(self):
        """初始化后自动设置创建时间和其他字段"""
        if not self.created_at:
            self.created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if self.flags is None:
            self.flags = set()
        if self.history is None:
            self.history = []


class GlossaryManager:
    """术语管理器主类"""

    def __init__(self):
        """初始化术语管理器"""
        self.terms: Dict[str, Term] = {}  # 存储所有术语
        self.load_terms()  # 加载已保存的术语

    def add_term(self, term: Term) -> bool:
        """添加新术语
        Args:
            term: Term对象
        Returns:
            bool: 是否添加成功
        """
        try:
            if not term.term.strip() or not term.translation.strip():
                raise ValueError("术语和翻译不能为空")
            self.terms[term.term] = term
            self.save_terms()  # 保存到文件
            return True
        except Exception as e:
            messagebox.showerror("错误", f"添加术语失败：{str(e)}")
            return False

    def remove_term(self, term_key: str) -> bool:
        """删除术语
        Args:
            term_key: 术语原文
        Returns:
            bool: 是否删除成功
        """
        try:
            if term_key in self.terms:
                del self.terms[term_key]
                self.save_terms()
                return True
            return False
        except Exception as e:
            messagebox.showerror("错误", f"删除术语失败：{str(e)}")
            return False

    def save_terms(self):
        """保存术语到文件"""
        try:
            # 确保data目录存在
            Path("data").mkdir(exist_ok=True)

            # 将术语数据转换为可序列化的格式
            data = {
                term_key: {
                    "term": term.term,
                    "translation": term.translation,
                    "category": term.category,
                    "context": term.context,
                    "notes": term.notes,
                    "created_at": term.created_at
                }
                for term_key, term in self.terms.items()
            }

            # 保存到JSON文件
            with open("data/terms.json", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        except Exception as e:
            messagebox.showerror("错误", f"保存术语失败：{str(e)}")

    def load_terms(self):
        """从文件加载术语"""
        try:
            if Path("data/terms.json").exists():
                with open("data/terms.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.terms = {
                        key: Term(**value)
                        for key, value in data.items()
                    }
        except Exception as e:
            messagebox.showerror("错误", f"加载术语失败：{str(e)}")

    def get_statistics(self) -> dict:
        """获取术语统计信息"""
        total = len(self.terms)
        translated = sum(1 for t in self.terms.values() if t.translation)
        categories = {}
        for term in self.terms.values():
            if term.category:
                categories[term.category] = categories.get(term.category, 0) + 1

        return {
            "总数": total,
            "已翻译": translated,
            "翻译进度": f"{(translated / total * 100):.1f}%" if total else "0%",
            "分类统计": categories
        }

    def check_duplicates(self) -> dict:
        """检查重复术语和翻译"""
        duplicates = {
            "terms": [],  # 重复术语
            "translations": []  # 重复翻译
        }

        translations = {}
        for term in self.terms.values():
            # 检查重复翻译
            if term.translation in translations:
                duplicates["translations"].append((
                    translations[term.translation].term,
                    term.term
                ))
            else:
                translations[term.translation] = term

        return duplicates

    def check_consistency(self) -> list:
        """检查术语翻译一致性"""
        issues = []
        translations = {}

        for term in self.terms.values():
            # 检查同一个原文是否有不同翻译
            if term.term in translations:
                if translations[term.term] != term.translation:
                    issues.append(f"术语'{term.term}'有不同的翻译: "
                                  f"{translations[term.term]} vs {term.translation}")
            else:
                translations[term.term] = term.translation

        return issues


class GlossaryGUI:
    """术语管理器的图形界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("术语管理器")
        self.root.geometry("800x600")

        # 初始化术语管理器
        self.manager = GlossaryManager()

        # 设置界面
        self.setup_ui()
        self.setup_context_menu()

    def setup_ui(self):
        """设置主界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="5")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # 创建搜索框
        search_frame = ttk.Frame(main_frame)
        search_frame.grid(row=0, column=0, pady=5, sticky="ew")
        ttk.Label(search_frame, text="搜索：").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search)
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 创建按钮框
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, pady=5)
        ttk.Button(button_frame, text="添加术语", command=self.add_term).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="编辑术语", command=self.edit_term).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除术语", command=self.delete_term).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导入", command=self.import_terms).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导出", command=self.export_terms).pack(side=tk.LEFT, padx=5)

        # 创建术语表格
        self.create_table(main_frame)

        # 配置主框架的网格权重
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # 在button_frame中添加新按钮
        ttk.Button(button_frame, text="术语统计", command=self.show_statistics).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="查重检查", command=self.show_duplicates).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="一致性检查", command=self.show_consistency).pack(side=tk.LEFT, padx=5)

        self.setup_status_bar()

    def create_table(self, parent):
        """创建术语表格"""
        # 创建表格框架
        table_frame = ttk.Frame(parent)
        table_frame.grid(row=2, column=0, sticky="nsew")

        # 定义列
        columns = ("term", "translation", "category", "context", "notes")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        # 设置列标题
        column_names = {
            "term": "原文",
            "translation": "译文",
            "category": "分类",
            "context": "上下文",
            "notes": "备注"
        }

        for col in columns:
            self.tree.heading(col, text=column_names[col])
            self.tree.column(col, width=100)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # 放置表格和滚动条
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # 绑定双击事件
        self.tree.bind("<Double-1>", lambda e: self.edit_term())

        # 更新表格数据
        self.update_table()

        # 配置表格框架的网格权重
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

    def update_table(self, search_text=""):
        """更新表格内容"""
        # 清空现有内容
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 根据搜索文本筛选并添加术语
        search_text = search_text.lower()
        for term in self.manager.terms.values():
            if (not search_text or
                    search_text in term.term.lower() or
                    search_text in term.translation.lower()):
                self.tree.insert("", tk.END, values=(
                    term.term,
                    term.translation,
                    term.category,
                    term.context,
                    term.notes
                ))

    def on_search(self, *args):
        """搜索功能"""
        self.update_table(self.search_var.get())

    def add_term(self):
        """添加术语"""
        dialog = TermDialog(self.root)
        if result := dialog.show():
            if self.manager.add_term(result):
                self.update_table()

    def edit_term(self):
        """编辑术语"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要编辑的术语")
            return

        item = selection[0]
        term_key = self.tree.item(item)["values"][0]
        term = self.manager.terms.get(term_key)

        if term:
            dialog = TermDialog(self.root, term)
            if result := dialog.show():
                self.manager.remove_term(term_key)
                self.manager.add_term(result)
                self.update_table()

    def delete_term(self):
        """删除术语"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的术语")
            return

        if messagebox.askyesno("确认", "确定要删除选中的术语吗？"):
            for item in selection:
                term_key = self.tree.item(item)["values"][0]
                if self.manager.remove_term(term_key):
                    self.update_table()

    def import_terms(self):
        """导入术语"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel文件", "*.xlsx"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )

        if not file_path:
            return

        try:
            # 读取文件
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            # 获取列映射
            dialog = ColumnMappingDialog(self.root, list(df.columns))
            if not (mapping := dialog.show()):
                return

            # 应用列映射
            df = df.rename(columns={v: k for k, v in mapping.items() if v})

            # 导入术语
            success_count = 0
            for _, row in df.iterrows():
                try:
                    term = Term(
                        term=row['Term'],
                        translation=row['Translation'],
                        category=row.get('Category', ''),
                        context=row.get('Context', ''),
                        notes=row.get('Notes', '')
                    )
                    if self.manager.add_term(term):
                        success_count += 1
                except Exception as e:
                    continue

            self.update_table()
            messagebox.showinfo("导入完成", f"成功导入 {success_count} 条术语")

        except Exception as e:
            messagebox.showerror("错误", f"导入失败：{str(e)}")

    def export_terms(self):
        """导出术语"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[
                ("CSV文件", "*.csv"),
                ("Excel文件", "*.xlsx")
            ]
        )

        if not file_path:
            return

        try:
            # 转换为DataFrame
            data = [vars(term) for term in self.manager.terms.values()]
            df = pd.DataFrame(data)

            # 导出文件
            if file_path.lower().endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                df.to_excel(file_path, index=False)

            messagebox.showinfo("导出完成", "术语表已成功导出")

        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def show_statistics(self):
        """显示统计信息"""
        stats = self.manager.get_statistics()

        # 创建统计信息窗口
        stats_window = tk.Toplevel(self.root)
        stats_window.title("术语统计")
        stats_window.geometry("400x300")

        # 显示基本统计信息
        text = tk.Text(stats_window, wrap=tk.WORD, width=40, height=15)
        text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        text.insert(tk.END, f"术语总数: {stats['总数']}\n")
        text.insert(tk.END, f"已翻译: {stats['已翻译']}\n")
        text.insert(tk.END, f"翻译进度: {stats['翻译进度']}\n\n")

        text.insert(tk.END, "分类统计:\n")
        for category, count in stats['分类统计'].items():
            text.insert(tk.END, f"{category}: {count}条\n")

        text.configure(state='disabled')

    def show_duplicates(self):
        """显示重复检查结果"""
        duplicates = self.manager.check_duplicates()

        # 创建查重窗口
        dup_window = tk.Toplevel(self.root)
        dup_window.title("重复检查")
        dup_window.geometry("500x400")

        # 创建文本框显示结果
        text = tk.Text(dup_window, wrap=tk.WORD, width=50, height=20)
        text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        if duplicates["translations"]:
            text.insert(tk.END, "发现重复翻译:\n")
            for term1, term2 in duplicates["translations"]:
                text.insert(tk.END, f"- '{term1}' 和 '{term2}' 使用了相同的翻译\n")
        else:
            text.insert(tk.END, "未发现重复翻译\n")

        text.configure(state='disabled')

    def show_consistency(self):
        """显示一致性检查结果"""
        issues = self.manager.check_consistency()

        # 创建一致性检查窗口
        cons_window = tk.Toplevel(self.root)
        cons_window.title("一致性检查")
        cons_window.geometry("500x400")

        # 创建文本框显示结果
        text = tk.Text(cons_window, wrap=tk.WORD, width=50, height=20)
        text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        if issues:
            text.insert(tk.END, "发现以下一致性问题:\n")
            for issue in issues:
                text.insert(tk.END, f"- {issue}\n")
        else:
            text.insert(tk.END, "未发现一致性问题\n")

        text.configure(state='disabled')

    def setup_status_bar(self):
        """添加状态栏"""
        self.status_bar = ttk.Label(self.root, text="", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=999, column=0, sticky="ew")
        self.update_status_bar()

    def update_status_bar(self):
        """更新状态栏信息"""
        total = len(self.manager.terms)
        translated = sum(1 for term in self.manager.terms.values() if term.translation)
        self.status_bar.config(
            text=f" 总计: {total} 条术语 | 已翻译: {translated} 条 | "
                 f"完成度: {(translated / total * 100):.1f}%" if total else "0%"
        )

    def setup_context_menu(self):
        """设置右键菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="编辑选中项", command=self.edit_term)
        self.context_menu.add_command(label="删除选中项", command=self.delete_term)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="复制原文", command=lambda: self.copy_term_field("term"))
        self.context_menu.add_command(label="复制译文", command=lambda: self.copy_term_field("translation"))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="复制整行", command=self.copy_full_term)

        # 在树形视图中绑定右键事件
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        """显示右键菜单"""
        # 先获取点击位置的项
        item = self.tree.identify_row(event.y)
        if item:
            # 选中点击的项
            self.tree.selection_set(item)
            # 显示菜单
            self.context_menu.post(event.x_root, event.y_root)

    def copy_term_field(self, field):
        """复制指定字段到剪贴板"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            fields = ["term", "translation", "category", "context", "notes"]
            value = self.tree.item(item)["values"][fields.index(field)]
            self.root.clipboard_clear()
            self.root.clipboard_append(str(value))
            messagebox.showinfo("提示", f"已复制{field}到剪贴板")

    def copy_full_term(self):
        """复制整行信息到剪贴板"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            values = self.tree.item(item)["values"]
            text = "\t".join(str(v) for v in values)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("提示", "已复制整行信息到剪贴板")


class TermDialog:
    """术语编辑对话框"""

    def __init__(self, parent, term=None):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("编辑术语" if term else "添加术语")
        self.dialog.geometry("400x300")
        self.term = term
        self.result = None
        self.setup_ui()
        # 添加快捷键
        self.setup_shortcuts()

    def setup_shortcuts(self):
        """设置快捷键"""
        self.root.bind('<Control-n>', lambda e: self.add_term())  # 新建
        self.root.bind('<Control-f>', lambda e: self.search_var.set(''))  # 搜索
        self.root.bind('<Delete>', lambda e: self.delete_term())  # 删除
        self.root.bind('<Control-s>', lambda e: self.export_terms())  # 导出

    def setup_ui(self):
        """设置对话框界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # 创建输入字段
        fields = [
            ("term", "术语"),
            ("translation", "翻译"),
            ("category", "分类"),
            ("context", "上下文"),
            ("notes", "备注")
        ]

        self.entries = {}
        for i, (field, label) in enumerate(fields):
            ttk.Label(main_frame, text=label).grid(row=i, column=0, padx=5, pady=5)
            entry = ttk.Entry(main_frame, width=40)
            entry.grid(row=i, column=1, padx=5, pady=5)
            if self.term:
                entry.insert(0, getattr(self.term, field))
            self.entries[field] = entry

        # 添加保存按钮
        ttk.Button(main_frame, text="保存", command=self.save).grid(
            row=len(fields), column=0, columnspan=2, pady=20
        )

        self.setup_status_bar()

    def save(self):
        """保存术语"""
        try:
            term_data = {
                field: entry.get()
                for field, entry in self.entries.items()
            }
            self.result = Term(**term_data)
            self.dialog.destroy()
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def show(self):
        """显示对话框并返回结果"""
        self.dialog.grab_set()
        self.dialog.wait_window()
        return self.result


class ColumnMappingDialog:
    """列映射对话框，用于导入文件时映射列名"""

    def __init__(self, parent, columns):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("列映射")
        self.dialog.geometry("400x300")
        self.columns = columns
        self.result = None
        self.setup_ui()

    def setup_ui(self):
        """设置对话框界面"""
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # 添加说明标签
        ttk.Label(main_frame, text="请选择对应的列：").grid(
            row=0, column=0, columnspan=2, pady=10
        )

        # 创建下拉框
        self.mappings = {}
        required_fields = [
            ("Term", "原文列"),
            ("Translation", "译文列")
        ]
        optional_fields = [
            ("Category", "分类列"),
            ("Context", "上下文列"),
            ("Notes", "备注列")
        ]

        row = 1
        # 必填字段
        for field, label in required_fields:
            ttk.Label(main_frame, text=f"{label}*").grid(row=row, column=0, padx=5, pady=5)
            combo = ttk.Combobox(main_frame, values=self.columns)
            combo.grid(row=row, column=1, padx=5, pady=5)
            self.mappings[field] = combo
            row += 1

        # 选填字段
        for field, label in optional_fields:
            ttk.Label(main_frame, text=label).grid(row=row, column=0, padx=5, pady=5)
            combo = ttk.Combobox(main_frame, values=[''] + self.columns)
            combo.grid(row=row, column=1, padx=5, pady=5)
            self.mappings[field] = combo
            row += 1

        # 添加确认按钮
        ttk.Button(main_frame, text="确认", command=self.confirm).grid(
            row=row, column=0, columnspan=2, pady=20
        )

    def confirm(self):
        """确认选择的映射"""
        # 验证必填字段
        required_fields = ['Term', 'Translation']
        for field in required_fields:
            if not self.mappings[field].get():
                messagebox.showerror("错误", f"必须选择{field}列")
                return

        self.result = {
            field: combo.get()
            for field, combo in self.mappings.items()
        }
        self.dialog.destroy()

    def show(self):
        """显示对话框并返回结果"""
        self.dialog.grab_set()
        self.dialog.wait_window()
        return self.result


def main():
    """主程序入口"""
    root = tk.Tk()
    app = GlossaryGUI(root)

    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap("icon.ico")
    except:
        pass

    # 启动主循环
    root.mainloop()


if __name__ == "__main__":
    main()