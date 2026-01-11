import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import copy
from utils import pdf_ops
from PIL import Image

# --- 【修正1】PyInstaller用のパス解決関数を追加 ---
def resource_path(relative_path):
    """ PyInstallerでリソースの絶対パスを取得する """
    try:
        # PyInstallerで作成された一時フォルダのパス
        base_path = sys._MEIPASS
    except Exception:
        # 通常実行時のパス
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
# ------------------------------------------------

# --- UI設定 ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PreviewWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("PDF Preview")
        self.geometry("700x850")
        
        self.current_page = 0
        self.total_pages = 0
        self.pdf_stream = None

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Left>", lambda e: self.prev_page())
        self.bind("<Right>", lambda e: self.next_page())
        self.after(100, self.focus_force)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- ナビゲーション (Row 0) ---
        self.control_frame = ctk.CTkFrame(self)
        self.control_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        self.control_frame.grid_columnconfigure(1, weight=1)

        self.prev_btn = ctk.CTkButton(self.control_frame, text="< 前へ", width=100, command=self.prev_page)
        self.prev_btn.grid(row=0, column=0, padx=10, pady=10)
        
        self.page_label = ctk.CTkLabel(self.control_frame, text="Page: 0 / 0", font=("Arial", 16, "bold"))
        self.page_label.grid(row=0, column=1, padx=10)
        
        self.next_btn = ctk.CTkButton(self.control_frame, text="次へ >", width=100, command=self.next_page)
        self.next_btn.grid(row=0, column=2, padx=10, pady=10)

        # --- アクションボタン (Row 1) ---
        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.action_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # 削除
        self.del_btn = ctk.CTkButton(self.action_frame, text="✕ 削除", width=120, fg_color="#cc0000", hover_color="#990000",
                                     command=self.request_delete)
        self.del_btn.grid(row=0, column=0, padx=5)

        # 回転
        self.rot_btn = ctk.CTkButton(self.action_frame, text="↻ 回転", width=120, fg_color="#0099cc", hover_color="#006699",
                                     command=self.request_rotate)
        self.rot_btn.grid(row=0, column=1, padx=5)

        # 縦横変換
        self.ori_btn = ctk.CTkButton(self.action_frame, text="▭ 縦横変換", width=120, fg_color="#ff9900", hover_color="#cc7a00",
                                     command=self.request_orientation)
        self.ori_btn.grid(row=0, column=2, padx=5)

        # --- 画像表示 (Row 2) ---
        self.image_label = ctk.CTkLabel(self, text="Loading...")
        self.image_label.grid(row=2, column=0, padx=20, pady=10)

    def on_close(self):
        self.parent.preview_window = None
        self.destroy()

    def request_delete(self):
        if not self.parent.pages or self.current_page >= len(self.parent.pages): return
        item = self.parent.pages[self.current_page]
        self.parent.delete_item(self.current_page, False, item)

    def request_rotate(self):
        if not self.parent.pages or self.current_page >= len(self.parent.pages): return
        item = self.parent.pages[self.current_page]
        self.parent.rotate_item(self.current_page, False, item)

    def request_orientation(self):
        if not self.parent.pages or self.current_page >= len(self.parent.pages): return
        item = self.parent.pages[self.current_page]
        self.parent.toggle_orientation(self.current_page, False, item)

    def update_preview(self, page_list, start_page=None):
        if not page_list:
            self.image_label.configure(text="ファイルがありません", image=None)
            self.page_label.configure(text="0 / 0")
            return
        
        if start_page is not None:
            self.current_page = start_page

        self.pdf_stream = pdf_ops.merge_pdfs_securely(page_list, output_path=None)
        
        if self.pdf_stream:
            import fitz
            try:
                with fitz.open(stream=self.pdf_stream, filetype="pdf") as doc:
                    self.total_pages = len(doc)
            except:
                self.total_pages = 0
            
            if self.current_page >= self.total_pages:
                self.current_page = max(0, self.total_pages - 1)
            self.show_page()

    def show_page(self):
        if not self.pdf_stream or self.total_pages == 0:
            self.image_label.configure(text="No Pages", image=None)
            return

        pil_image = pdf_ops.get_preview_image(self.pdf_stream, self.current_page)
        if pil_image:
            w, h = pil_image.size
            if w > h:
                self.geometry("1000x750")
            else:
                self.geometry("700x850")

            win_w = 950 if w > h else 650
            win_h = 700 if w > h else 800
            ratio = min(win_w/w, win_h/h)
            new_size = (int(w*ratio), int(h*ratio))
            pil_image = pil_image.resize(new_size, Image.Resampling.LANCZOS)
            ctk_image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=new_size)
            self.image_label.configure(image=ctk_image, text="")
            self.page_label.configure(text=f"Page: {self.current_page + 1} / {self.total_pages}")
        else:
            self.image_label.configure(text="Preview Error", image=None)

    def next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.show_page()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.show_page()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- 設定値 ---
        self.mouse_wheel_speed = 200.0   
        self.base_thumb_size = 200     
        self.auto_scroll_speed = 20    # 自動スクロールのチェック間隔(ms)
        self.auto_scroll_margin = 50   # 反応領域(px)
        self.auto_scroll_amount = 20   # ★自動スクロールの移動量 (数値を増やすと速くなる)
        # ------------

        self.title("Secure PDF Merger")
        self.geometry("1150x750")

        self.pages = []
        self.history = []
        self.redo_stack = []
        self.view_mode = "file"
        
        self.zoom_level = 1.0 

        self.preview_window = None
        self.dragging_index = None 
        self.auto_scroll_job = None 
        
        self.thumbnail_cache = {}

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # サイドバー
        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(9, weight=1)

        ctk.CTkLabel(self.sidebar_frame, text="Secure PDF\nMerger", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, padx=20, pady=(20, 10))

        # 履歴
        self.history_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.history_frame.grid(row=1, column=0, padx=10, pady=5)
        self.undo_btn = ctk.CTkButton(self.history_frame, text="←", width=40, command=self.undo_action, state="disabled")
        self.undo_btn.pack(side="left", padx=5)
        self.redo_btn = ctk.CTkButton(self.history_frame, text="→", width=40, command=self.redo_action, state="disabled")
        self.redo_btn.pack(side="left", padx=5)

        # 表示切替
        self.view_mode_btn = ctk.CTkButton(self.sidebar_frame, text="表示: ファイル単位", fg_color="#555", command=self.toggle_view_mode)
        self.view_mode_btn.grid(row=2, column=0, padx=20, pady=10)

        # ズームボタン
        self.zoom_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.zoom_in_btn = ctk.CTkButton(self.zoom_frame, text="＋", width=40, command=self.zoom_in)
        self.zoom_in_btn.pack(side="left", padx=5)
        ctk.CTkLabel(self.zoom_frame, text="Zoom", font=("Arial", 12)).pack(side="left", padx=5)
        self.zoom_out_btn = ctk.CTkButton(self.zoom_frame, text="－", width=40, command=self.zoom_out)
        self.zoom_out_btn.pack(side="left", padx=5)

        ctk.CTkButton(self.sidebar_frame, text="+ ファイルを追加", command=self.add_files_event).grid(row=4, column=0, padx=20, pady=10)
        ctk.CTkButton(self.sidebar_frame, text="↻ すべて回転", fg_color="gray", hover_color="gray30", command=self.rotate_all_event).grid(row=5, column=0, padx=20, pady=10)
        ctk.CTkButton(self.sidebar_frame, text="👁 プレビュー", fg_color="#d97706", hover_color="#b45309", command=self.open_preview).grid(row=6, column=0, padx=20, pady=10)
        ctk.CTkButton(self.sidebar_frame, text="PDFを結合して保存", fg_color="green", hover_color="darkgreen", font=ctk.CTkFont(weight="bold"), command=self.merge_event).grid(row=7, column=0, padx=20, pady=10)

        # メインエリア
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        self.list_label = ctk.CTkLabel(self.main_frame, text="結合リスト", font=ctk.CTkFont(size=16))
        self.list_label.pack(anchor="w", pady=(0, 10))

        self.scrollable_list = ctk.CTkScrollableFrame(self.main_frame, label_text="ファイル一覧")
        self.scrollable_list.pack(fill="both", expand=True)
        
        self.scrollable_list.bind("<Configure>", self.on_resize)
        
        self.scrollable_list._parent_canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        
        self.last_width = 0

    def _on_mouse_wheel(self, event):
        try:
            canvas = self.scrollable_list._parent_canvas
            units = int(-1 * (event.delta / 120) * self.mouse_wheel_speed)
            canvas.yview_scroll(units, "units")
        except: pass

    def zoom_in(self):
        self.zoom_level = min(3.0, self.zoom_level + 0.2)
        self.full_refresh_list_ui()

    def zoom_out(self):
        self.zoom_level = max(0.4, self.zoom_level - 0.2)
        self.full_refresh_list_ui()

    def on_resize(self, event):
        if self.view_mode == "page" and abs(event.width - self.last_width) > 20:
            self.last_width = event.width
            self.full_refresh_list_ui()

    def save_state(self):
        self.history.append(copy.deepcopy(self.pages))
        self.redo_stack.clear()
        self.update_history_buttons()

    def undo_action(self):
        if not self.history: return
        self.redo_stack.append(copy.deepcopy(self.pages))
        self.pages = self.history.pop()
        self.full_refresh_list_ui()
        self.update_history_buttons()

    def redo_action(self):
        if not self.redo_stack: return
        self.history.append(copy.deepcopy(self.pages))
        self.pages = self.redo_stack.pop()
        self.full_refresh_list_ui()
        self.update_history_buttons()

    def update_history_buttons(self):
        self.undo_btn.configure(state="normal" if self.history else "disabled")
        self.redo_btn.configure(state="normal" if self.redo_stack else "disabled")

    def toggle_view_mode(self):
        self.view_mode = "page" if self.view_mode == "file" else "file"
        self.view_mode_btn.configure(text=f"表示: {'ページ単位' if self.view_mode == 'page' else 'ファイル単位'}")
        
        if self.view_mode == "page":
            self.zoom_frame.grid(row=3, column=0, padx=20, pady=5)
        else:
            self.zoom_frame.grid_forget()
            
        self.full_refresh_list_ui()

    def add_files_event(self):
        filetypes = [("All Files", "*.*")]
        file_paths = filedialog.askopenfilenames(title="ファイルを選択", filetypes=filetypes)
        
        if file_paths:
            self.save_state()
            for path in file_paths:
                final_path = path
                filename = os.path.basename(path)
                
                if not path.lower().endswith('.pdf'):
                    print(f"Converting: {filename}")
                    try:
                        converted = pdf_ops.convert_to_pdf(path, is_landscape=False)
                        if converted and os.path.exists(converted):
                            final_path = converted
                        else:
                            print(f"Skipping: {filename}")
                            continue
                    except: continue

                info = pdf_ops.get_pdf_info(final_path)
                if info:
                    for i in range(info['pages']):
                        is_port = True
                        if 'page_details' in info and i < len(info['page_details']):
                            is_port = info['page_details'][i]['is_portrait']

                        self.pages.append({
                            'path': final_path,
                            'original_source_path': path,
                            'is_generated': (final_path != path),
                            'filename': filename,
                            'page_index': i,
                            'rotation': 0,
                            'is_portrait_original': is_port,
                            'is_landscape_generated': False,
                            'id': f"{filename}_{i}_{os.urandom(4).hex()}"
                        })
            self.full_refresh_list_ui()

    # --- UI更新 (軽量化対応) ---
    def full_refresh_list_ui(self, update_scroll_region=True):
        for widget in self.scrollable_list.winfo_children():
            widget.destroy()
        
        mode_text = "ファイル単位"
        if self.view_mode == "page":
            mode_text = "ページ単位 (プレビュー)"
            self.render_page_mode()
        else:
            self.render_file_mode()
        
        self.list_label.configure(text=f"結合リスト ({len(self.pages)} ページ) - {mode_text}")
        
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.update_preview(self.pages)

        # ★カクつき防止: ドラッグ中(update_scroll_region=False)は重い処理をスキップ
        if update_scroll_region:
            self.scrollable_list.update_idletasks()
            try:
                canvas = self.scrollable_list._parent_canvas
                canvas.configure(scrollregion=canvas.bbox("all"))
            except:
                pass

    # --- ページモード (グリッド表示) ---
    def render_page_mode(self):
        if not self.pages: return
        
        width = self.scrollable_list.winfo_width()
        if width < 100: width = 800 
        
        item_w = int(self.base_thumb_size * self.zoom_level)
        item_h = int(item_w * 1.414) 
        gap = 15 
        
        safe_width_percent = 0.52
        usable_width = width * safe_width_percent
        
        columns = max(1, int(usable_width // (item_w + gap)))
        
        for index, item in enumerate(self.pages):
            r = index // columns
            c = index % columns
            self.create_page_grid_item(index, item, r, c, item_w, item_h)

    def create_page_grid_item(self, index, item_data, r, c, w, h):
        fg_color = None
        if index == self.dragging_index: fg_color = "#444444"
        
        frame = ctk.CTkFrame(self.scrollable_list, width=w, height=h, fg_color=fg_color, border_width=0)
        frame.grid(row=r, column=c, padx=7, pady=7)
        frame.grid_propagate(False)

        cache_key = f"{item_data['id']}_{item_data['rotation']}_{w}"
        if cache_key in self.thumbnail_cache:
            img = self.thumbnail_cache[cache_key]
        else:
            pil_img = pdf_ops.get_page_thumbnail(item_data['path'], item_data['page_index'], item_data['rotation'])
            if pil_img:
                img_w, img_h = pil_img.size
                ratio = min((w-10)/img_w, (h-10)/img_h)
                new_size = (int(img_w*ratio), int(img_h*ratio))
                pil_img = pil_img.resize(new_size, Image.Resampling.LANCZOS)
                img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=new_size)
                self.thumbnail_cache[cache_key] = img
            else:
                img = None

        if img:
            img_label = ctk.CTkLabel(frame, text="", image=img)
            img_label.place(relx=0.5, rely=0.5, anchor="center")
            img_label.bind("<Button-1>", lambda e, idx=index: self.start_drag(e, idx))
            # 個別バインドは残すが、メインはGlobal Bindで制御
            img_label.bind("<B1-Motion>", self.on_drag)
            img_label.bind("<ButtonRelease-1>", self.stop_drag)

        # --- ボタン類の追加 (左上) ---
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent", width=90, height=30)
        btn_frame.place(relx=0.03, rely=0.03, anchor="nw")

        # 1. 削除ボタン (✕)
        del_btn = ctk.CTkButton(btn_frame, text="✕", width=24, height=24, fg_color="#cc0000", hover_color="#990000",
                                font=("Arial", 14, "bold"),
                                command=lambda: self.delete_item(index, False, item_data))
        del_btn.pack(side="left", padx=1)

        # 2. 回転ボタン (↻)
        rot_btn = ctk.CTkButton(btn_frame, text="↻", width=24, height=24, fg_color="#0099cc", hover_color="#006699",
                                font=("Arial", 14, "bold"),
                                command=lambda: self.rotate_item(index, False, item_data))
        rot_btn.pack(side="left", padx=1)

        # 3. 拡大プレビューボタン (🔍)
        zoom_btn = ctk.CTkButton(btn_frame, text="🔍", width=24, height=24, fg_color="#ff9900", hover_color="#cc7a00",
                                font=("Arial", 14, "bold"),
                                command=lambda: self.open_preview(start_page=index))
        zoom_btn.pack(side="left", padx=1)
        # -----------------------------

        num_label = ctk.CTkLabel(frame, text=str(index+1), width=20, height=20, fg_color="gray", text_color="white", corner_radius=10)
        num_label.place(relx=0.95, rely=0.05, anchor="ne")

        frame.bind("<Button-1>", lambda e, idx=index: self.start_drag(e, idx))
        frame.bind("<B1-Motion>", self.on_drag)
        frame.bind("<ButtonRelease-1>", self.stop_drag)

    # --- ファイルモード (リスト表示) ---
    def render_file_mode(self):
        if not self.pages: return
        groups = []
        curr = [self.pages[0]]
        for p in self.pages[1:]:
            if p['path'] == curr[-1]['path']: curr.append(p)
            else: groups.append(curr); curr = [p]
        groups.append(curr)

        for i, group in enumerate(groups):
            first = group[0]
            group_item = {
                'filename': first['filename'],
                'path': first['path'],
                'original_source_path': first.get('original_source_path'),
                'is_generated': first.get('is_generated', False),
                'pages_count': len(group),
                'rotation_display': first['rotation'],
                'is_portrait_original': first.get('is_portrait_original', True),
                'is_landscape_generated': first.get('is_landscape_generated', False),
                'data': group
            }
            self.create_file_list_item(i, group_item)

    def create_file_list_item(self, index, item_data):
        fg_color = None
        if index == self.dragging_index: fg_color = "#444444"
        
        item_frame = ctk.CTkFrame(self.scrollable_list, fg_color=fg_color, border_color=None, border_width=0)
        item_frame.pack(fill="x", pady=5, padx=5)

        drag = ctk.CTkLabel(item_frame, text="≡", font=("Arial", 20), cursor="hand2", width=30)
        drag.pack(side="left", padx=(5, 0))

        is_landscape_gen = item_data.get('is_landscape_generated', False)
        if is_landscape_gen: orientation_icon = "▭"
        else:
            rot = item_data.get('rotation_display', 0)
            is_orig_port = item_data.get('is_portrait_original', True)
            is_curr_port = (is_orig_port and rot%180==0) or (not is_orig_port and rot%180!=0)
            orientation_icon = "▯" if is_curr_port else "▭"

        rot_text = f" [+{item_data['rotation_display']}°]" if item_data['rotation_display'] > 0 else ""
        title = f"{index+1}. {item_data['filename']}{rot_text}"
        sub = f"({item_data['pages_count']} pages)"

        name = ctk.CTkLabel(item_frame, text=title, font=ctk.CTkFont(weight="bold"))
        name.pack(side="left", padx=10, pady=10)
        ctk.CTkLabel(item_frame, text=sub, text_color="gray").pack(side="left", padx=5)

        for w in [item_frame, drag, name]:
            w.bind("<Button-1>", lambda e, idx=index: self.start_drag(e, idx))
            w.bind("<B1-Motion>", self.on_drag)
            w.bind("<ButtonRelease-1>", self.stop_drag)

        ctk.CTkButton(item_frame, text="✕", width=30, fg_color="transparent", text_color="red", hover_color=("#fee", "#400"), 
                      command=lambda: self.delete_item(index, True, item_data)).pack(side="right", padx=10)
        ctk.CTkButton(item_frame, text="↻", width=30, fg_color="transparent", text_color="#00BFFF", hover_color=("lightblue", "navy"), 
                      command=lambda: self.rotate_item(index, True, item_data)).pack(side="right", padx=2)
        ctk.CTkButton(item_frame, text=orientation_icon, width=30, fg_color="transparent", 
                      text_color="white", font=("Arial", 20), hover_color=("gray70", "gray30"),
                      command=lambda: self.toggle_orientation(index, True, item_data)).pack(side="right", padx=2)

    # --- 共通操作 ---
    def toggle_orientation(self, index, is_group, item_data):
        self.save_state()
        if not is_group:
            target_list = [item_data]
            target_item = item_data
        else:
            target_list = item_data['data']
            target_item = item_data

        is_gen = target_item.get('is_generated', False)
        source = target_item.get('original_source_path')
        is_land = target_item.get('is_landscape_generated', False)
        
        if is_gen and source and os.path.exists(source):
            new_is_landscape = not is_land
            print(f"Re-converting to {'Landscape' if new_is_landscape else 'Portrait'}...")
            try:
                new_pdf_path = pdf_ops.convert_to_pdf(source, is_landscape=new_is_landscape)
                if new_pdf_path:
                    for item in target_list:
                        item['path'] = new_pdf_path
                        item['is_landscape_generated'] = new_is_landscape
                        item['rotation'] = 0
                    self.full_refresh_list_ui()
                    return
            except: pass
        
        self.rotate_item(index, is_group, item_data)

    def delete_item(self, index, is_group, item_data):
        self.save_state()
        if is_group:
            ids = [x['id'] for x in item_data['data']]
            self.pages = [p for p in self.pages if p['id'] not in ids]
        else:
            del self.pages[index]
        self.full_refresh_list_ui()

    def rotate_item(self, index, is_group, item_data):
        self.save_state()
        if is_group:
            for p in item_data['data']: p['rotation'] = (p['rotation'] + 90) % 360
        else:
            self.pages[index]['rotation'] = (self.pages[index]['rotation'] + 90) % 360
        self.full_refresh_list_ui()

    def rotate_all_event(self):
        if not self.pages: return
        self.save_state()
        for p in self.pages: p['rotation'] = (p['rotation'] + 90) % 360
        self.full_refresh_list_ui()

    def start_drag(self, event, index):
        self.dragging_index = index
        self.configure(cursor="fleur")
        # ドラッグ開始時はスクロール領域の計算を行っておく
        self.full_refresh_list_ui(update_scroll_region=True)
        # ★重要修正: マウスがウィンドウ外に出てもドラッグを継続させるために
        # アプリ全体(self)にイベントをバインドします
        self.bind("<B1-Motion>", self.on_drag, add="+")
        self.bind("<ButtonRelease-1>", self.stop_drag, add="+")

    def on_drag(self, event):
        if self.dragging_index is None: return
        
        # --- 自動スクロール判定 ---
        self.check_auto_scroll()
        # -----------------------

        mouse_y = self.scrollable_list._parent_canvas.winfo_pointery()
        mouse_x = self.scrollable_list._parent_canvas.winfo_pointerx()
        children = self.scrollable_list.winfo_children()
        
        target_index = -1
        
        # ヒットテスト
        for i, child in enumerate(children):
            if not child.winfo_exists(): continue
            c_left = child.winfo_rootx()
            c_top = child.winfo_rooty()
            c_width = child.winfo_width()
            c_height = child.winfo_height()
            
            # マウスがそのアイテムの上にあるか判定
            if (c_left < mouse_x < c_left + c_width) and (c_top < mouse_y < c_top + c_height):
                target_index = i
                break
        
        # マウスが外にある場合、target_indexは-1のまま（＝並び替えを実行しない）
        # これにより、UI外にカーソルが出てもエラーにならず、戻ってきたら再開できる
        if target_index != -1 and target_index != self.dragging_index:
            if self.view_mode == "page":
                self.pages.insert(target_index, self.pages.pop(self.dragging_index))
            else:
                groups = []
                curr = [self.pages[0]]
                for p in self.pages[1:]:
                    if p['path'] == curr[-1]['path']: curr.append(p)
                    else: groups.append(curr); curr = [p]
                groups.append(curr)
                groups.insert(target_index, groups.pop(self.dragging_index))
                self.pages = [p for g in groups for p in g]
            
            self.dragging_index = target_index
            self.full_refresh_list_ui(update_scroll_region=False)

    def check_auto_scroll(self):
        if self.auto_scroll_job: return

        try:
            canvas = self.scrollable_list._parent_canvas
            # キャンバスのスクリーン上の座標
            c_top = canvas.winfo_rooty()
            c_height = canvas.winfo_height()
            c_bottom = c_top + c_height
            
            mouse_y = canvas.winfo_pointery()
            
            scroll_dir = 0
            if mouse_y < c_top + self.auto_scroll_margin:
                scroll_dir = -1
            elif mouse_y > c_bottom - self.auto_scroll_margin:
                scroll_dir = 1
            
            if scroll_dir != 0:
                canvas.yview_scroll(int(scroll_dir * self.auto_scroll_amount), "units")
                self.auto_scroll_job = self.after(self.auto_scroll_speed, self.do_auto_scroll_loop)
        except: pass

    def do_auto_scroll_loop(self):
        self.auto_scroll_job = None
        if self.dragging_index is not None:
            self.check_auto_scroll()

    def stop_drag(self, event):
        # ★重要修正: バインドを解除
        self.unbind("<B1-Motion>")
        self.unbind("<ButtonRelease-1>")

        if self.auto_scroll_job:
            self.after_cancel(self.auto_scroll_job)
            self.auto_scroll_job = None

        if self.dragging_index is not None: self.save_state()
        self.dragging_index = None
        self.configure(cursor="")
        self.full_refresh_list_ui(update_scroll_region=True)

    def open_preview(self, start_page=0):
        if not self.pages: return
        if self.preview_window is None or not self.preview_window.winfo_exists():
            self.preview_window = PreviewWindow(self)
        self.preview_window.update_preview(self.pages, start_page=start_page)
        self.preview_window.focus()

    def notify_preview(self):
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.update_preview(self.pages)

    def merge_event(self):
        if not self.pages: return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_path:
            self.configure(cursor="watch"); self.update()
            success = pdf_ops.merge_pdfs_securely(self.pages, output_path)
            self.configure(cursor="")
            if success: messagebox.showinfo("成功", "完了しました")
            else: messagebox.showerror("エラー", "失敗しました")

if __name__ == "__main__":
    app = App()
    app.mainloop()