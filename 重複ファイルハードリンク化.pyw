import os
import hashlib
import tkinter as tk
from tkinter import messagebox, ttk
import ctypes

def calculate_hash(file_path, buffer_size=65536):
    """ファイルのハッシュ値を計算します"""
    sha256 = hashlib.sha256()
    try:
        with open(file_path, 'rb') as f:
            while chunk := f.read(buffer_size):
                sha256.update(chunk)
        return sha256.hexdigest()
    except PermissionError:
        print(f"アクセス権限のないファイルをスキップしました: {file_path}")
        return None

def find_duplicate_files(drive):
    """指定ドライブ内で同一ハッシュを持つファイルを検索します"""
    hashes = {}
    for root, _, files in os.walk(drive):
        for file in files:
            if file.startswith("~$"):  # ~から始まるファイルを除外
                continue
            file_path = os.path.join(root, file)
            if not os.path.islink(file_path):  # ハードリンクはスキップ
                file_hash = calculate_hash(file_path)
                if file_hash:  # Noneでない場合のみ処理
                    if file_hash in hashes:
                        hashes[file_hash].append(file_path)
                    else:
                        hashes[file_hash] = [file_path]
    return {hash: paths for hash, paths in hashes.items() if len(paths) > 1}

class DuplicateFileFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("重複ファイル検索ツール")
        self.root.geometry("700x500")
        self.drives = self.get_available_drives()
        self.duplicate_files = {}

        # ドライブ選択部分
        tk.Label(root, text="検索対象ドライブを選択してください:").pack(anchor="w")
        self.drive_vars = {}
        for drive in self.drives:
            var = tk.BooleanVar(value=False)
            self.drive_vars[drive] = var
            tk.Checkbutton(root, text=drive, variable=var).pack(anchor="w")

        # 検索ボタン
        self.search_button = tk.Button(root, text="重複ファイルを検索", command=self.search_duplicates)
        self.search_button.pack(pady=10)

        # 重複ファイル一覧表示用ツリービュー
        columns = ("Hash", "File", "Path")
        self.tree = ttk.Treeview(root, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("Hash", text="ハッシュ値")
        self.tree.heading("File", text="ファイル名")
        self.tree.heading("Path", text="パス")
        self.tree.pack(fill="both", expand=True)
        
        # 行をクリックしたときのイベント
        self.tree.bind('<ButtonRelease-1>', self.on_item_click)
        
        # ハードリンク作成ボタン
        self.link_button = tk.Button(root, text="選択したファイルをハードリンク化", command=self.link_duplicates)
        self.link_button.pack(pady=10)

    def get_available_drives(self):
        """Cドライブとネットワークドライブを除外し、ローカルドライブとUSBメモリのみをリスト化"""
        available_drives = []
        drive_bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            if drive_bitmask & 1:
                drive_path = f"{letter}:\\"
                drive_type = ctypes.windll.kernel32.GetDriveTypeW(drive_path)
                # ローカルドライブ（固定ディスク）またはリムーバブルディスク（USBメモリ）、かつCドライブ以外
                if (drive_type == 3 or drive_type == 2) and letter != 'C':
                    available_drives.append(drive_path)
            drive_bitmask >>= 1
        return available_drives

    def search_duplicates(self):
        """選択されたドライブで重複ファイルを検索"""
        selected_drives = [drive for drive, var in self.drive_vars.items() if var.get()]
        if not selected_drives:
            messagebox.showwarning("警告", "検索対象のドライブを選択してください。")
            return
        self.duplicate_files = {}
        for drive in selected_drives:
            self.duplicate_files.update(find_duplicate_files(drive))
        self.display_duplicates()

    def display_duplicates(self):
        """重複ファイルをツリービューに表示"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for file_hash, files in self.duplicate_files.items():
            for file in files:
                self.tree.insert("", "end", values=(file_hash, os.path.basename(file), file))

    def on_item_click(self, event):
        """アイテムをクリックしたときに同じハッシュ値のファイルを全選択"""
        selected_item = self.tree.focus()
        if selected_item:
            item_values = self.tree.item(selected_item, "values")
            if item_values:
                file_hash = item_values[0]
                for item in self.tree.get_children():
                    item_values = self.tree.item(item, "values")
                    if item_values and item_values[0] == file_hash:
                        self.tree.selection_add(item)

    def link_duplicates(self):
        """ツリービューで選択されたファイルにハードリンクを作成"""
        selected_items = self.tree.selection()
        selected_paths = [self.tree.item(item, "values")[2] for item in selected_items]
        
        duplicates_to_link = {}
        for file_hash, files in self.duplicate_files.items():
            selected_hash_files = [f for f in files if f in selected_paths]
            if len(selected_hash_files) > 1:
                duplicates_to_link[file_hash] = selected_hash_files
        if duplicates_to_link:
            create_hardlinks(duplicates_to_link)
            messagebox.showinfo("情報", "ハードリンクが正常に作成されました。")
        else:
            messagebox.showwarning("警告", "リンク作成のためにファイルを選択してください。")

def create_hardlinks(duplicate_files):
    """選択された重複ファイルをハードリンクで統合"""
    for file_hash, files in duplicate_files.items():
        master_file = files[0]
        for duplicate_file in files[1:]:
            try:
                os.remove(duplicate_file)
                os.link(master_file, duplicate_file)
                print(f"ハードリンク作成: {duplicate_file} -> {master_file}")
            except Exception as e:
                print(f"エラー: {duplicate_file} のハードリンク作成に失敗しました。{e}")

# メインループ
root = tk.Tk()
app = DuplicateFileFinderApp(root)
root.mainloop()
