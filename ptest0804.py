import base64
import io
import json
import os
import re
import tkinter as tk
import urllib.error
import urllib.request
from collections import defaultdict
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import cv2
import numpy as np
import pandas as pd
from PIL import Image, ImageEnhance, ImageFilter, ImageOps, ImageTk
from pdf2image import convert_from_path
import pytesseract
from pytesseract import Output
from img2table.document import Image as Img2TableImage
from img2table.ocr import TesseractOCR

def _ensure_tesseract_runtime():
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]

    for candidate in candidates:
        if os.path.exists(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            tesseract_dir = os.path.dirname(candidate)
            current_path = os.environ.get("PATH", "")
            if tesseract_dir not in current_path.split(os.pathsep):
                os.environ["PATH"] = tesseract_dir + os.pathsep + current_path
            return candidate
    return None


TESSERACT_PATH = _ensure_tesseract_runtime()


class ExcelOcrApp:
    def __init__(self, window):
        self.window = window
        self.window.title("圖片/Excel OCR 工具")
        self.window.geometry("980x700")

        self.image_path = ""
        self.excel_path = ""
        self.current_df = None
        self.image_label = None
        self.text_box = None
        self.tree = None
        self.editing_row = None
        self.editing_col = None
        self.row_map = {}
        self.inline_editor = None
        self.editor_target = None
        self.sort_state = {}
        self.use_chatgpt = tk.BooleanVar(value=False)
        self.chatgpt_model = tk.StringVar(value="gpt-4o-mini")
        self.keyword_text = tk.StringVar(value="Vc, fz, ap")
        self.keyword_rows = []
        self.keyword_entry_frame = None
        self.openai_api_key = tk.StringVar(value=os.getenv("OPENAI_API_KEY", "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"))

        self.build_ui()

    def build_ui(self):
        frame = ttk.Frame(self.window, padding=10)
        frame.pack(fill="both", expand=True)

        top = ttk.Frame(frame)
        top.pack(fill="x", pady=(0, 10))

        buttons_row = ttk.Frame(top)
        buttons_row.pack(fill="x", pady=(0, 6))
        ttk.Button(buttons_row, text="A. 開啟圖片/PDF", command=self.open_source_file).pack(side="left", padx=(0, 10))
        ttk.Button(buttons_row, text="B. 開啟 Excel 檔", command=self.open_excel_file).pack(side="left", padx=(0, 10))
        ttk.Button(buttons_row, text="C. 儲存 Excel", command=self.save_excel_file).pack(side="left", padx=(0, 10))
        ttk.Button(buttons_row, text="新增列", command=self.add_row).pack(side="left", padx=(0, 10))
        ttk.Button(buttons_row, text="刪除列", command=self.delete_selected_row).pack(side="left")

        opt = ttk.Frame(top)
        opt.pack(fill="x", pady=(0, 6))
        ttk.Checkbutton(opt, text="使用 ChatGPT 辨識圖片", variable=self.use_chatgpt).pack(side="left")
        ttk.Label(opt, text="OpenAI API Key:").pack(side="left", padx=(10, 4))
        ttk.Entry(opt, textvariable=self.openai_api_key, width=50, show="*").pack(side="left")

        kw = ttk.LabelFrame(top, text="關鍵字優先 OCR")
        kw.pack(fill="x", pady=(0, 6))
        self.keyword_entry_frame = ttk.Frame(kw)
        self.keyword_entry_frame.pack(fill="x", padx=(8, 8), pady=(8, 8))
        ttk.Button(kw, text="＋新增", command=self.add_keyword_row).pack(anchor="w", padx=(8, 8), pady=(0, 8))
        self.add_keyword_row("Vc", checked=True)
        self.add_keyword_row("fz", checked=True)
        self.add_keyword_row("ap", checked=True)

        content = ttk.Panedwindow(frame, orient="horizontal")
        content.pack(fill="both", expand=True)

        left = ttk.Frame(content, padding=5)
        right = ttk.Frame(content, padding=5)
        content.add(left, weight=3)
        content.add(right, weight=4)

        ttk.Label(left, text="圖片預覽").pack(anchor="w")
        self.image_label = ttk.Label(left, text="尚未選擇圖片", background="#f0f0f0", relief="solid", padding=10)
        self.image_label.pack(fill="both", expand=True, pady=(5, 10))

        ttk.Label(right, text="OCR / Excel 內容").pack(anchor="w")
        self.text_box = ScrolledText(right, width=80, height=24)
        self.text_box.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(frame, show="headings")
        self.tree.pack(fill="both", expand=True, pady=(10, 0))
        self.tree.bind("<Double-1>", self.start_inline_edit)

        self.inline_editor = ttk.Entry(frame)
        self.inline_editor.bind("<Return>", self.commit_inline_edit)
        self.inline_editor.bind("<Escape>", self.cancel_inline_edit)
        self.inline_editor.place_forget()

        self.status_var = tk.StringVar(value="就緒")
        ttk.Label(frame, textvariable=self.status_var).pack(anchor="w", pady=(8, 0))

    def open_source_file(self):
        file_path = filedialog.askopenfilename(
            title="選擇圖片或 PDF",
            filetypes=[("Image/PDF", "*.jpg *.jpeg *.png *.pdf")]
        )
        if not file_path:
            return

        self.image_path = file_path
        ext = os.path.splitext(file_path)[1].lower()
        self.text_box.delete(1.0, tk.END)

        if ext in {".jpg", ".jpeg", ".png"}:
            self.display_image(file_path)
            df, text = self.ocr_image_to_table(file_path)
            self.text_box.insert(tk.END, text)
            self.show_excel_table(df)
        elif ext == ".pdf":
            df, text = self.ocr_pdf_to_table(file_path)
            self.text_box.insert(tk.END, text)
            self.show_preview_placeholder("PDF 已解析完成")
            self.show_excel_table(df)
        else:
            messagebox.showerror("錯誤", "請選擇 jpg / jpeg / png / pdf")
            return

        csv_path = os.path.join(os.getcwd(), "recognized_output.csv")
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        self.status_var.set(f"已辨識並匯出 CSV：{csv_path}")

    def add_keyword_row(self, default_value="", checked=False):
        if self.keyword_entry_frame is None:
            return
        row = ttk.Frame(self.keyword_entry_frame)
        row.pack(fill="x", pady=(0, 4))

        selected_var = tk.BooleanVar(value=checked)
        text_var = tk.StringVar(value=default_value)
        ttk.Checkbutton(row, variable=selected_var).pack(side="left")
        ttk.Entry(row, textvariable=text_var, width=20).pack(side="left", padx=(4, 0))

        self.keyword_rows.append({
            "selected": selected_var,
            "text": text_var,
            "row": row,
        })

    def _use_chatgpt_ocr(self):
        return bool(self.use_chatgpt.get())

    def _parse_keywords(self):
        keywords = []
        for item in self.keyword_rows:
            raw = item["text"].get().strip()
            if not raw:
                continue
            for token in re.split(r"[,;\n]+", raw):
                candidate = token.strip()
                if candidate:
                    keywords.append(candidate)
        return keywords

    def _required_keywords(self):
        required = []
        for item in self.keyword_rows:
            if item["selected"].get():
                raw = item["text"].get().strip()
                if raw:
                    required.append(raw)
        return required

    def _get_openai_key(self):
        key = self.openai_api_key.get().strip()
        if not key:
            key = os.getenv("OPENAI_API_KEY", "")
        return key

    def _call_chatgpt_ocr(self, file_path):
        api_key = self._get_openai_key()
        if not api_key:
            raise RuntimeError("未設定 OpenAI API Key，請先填入右上方欄位或設定 OPENAI_API_KEY")

        with open(file_path, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode("utf-8")

        mime_type = "image/jpeg"
        if file_path.lower().endswith(".png"):
            mime_type = "image/png"

        data_url = f"data:{mime_type};base64,{encoded}"
        payload = {
            "model": self.chatgpt_model.get(),
            "messages": [
                {
                    "role": "system",
                    "content": "You are a document OCR assistant. Extract table data from the image and return only a CSV string with a header row. If there is no clear table, return one-column CSV with the text contents. Preserve Chinese text and numeric values."
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "請從此圖片中辨識出表格內容，若是表格請保留欄位結構；若是文字段落，請整理為一列內容。請只輸出 CSV。"
                        },
                        {
                            "type": "image_url",
                            "image_url": {"url": data_url}
                        }
                    ]
                }
            ],
            "temperature": 0.1
        }

        request = urllib.request.Request(
            url="https://api.openai.com/v1/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            },
            method="POST"
        )

        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                response_body = response.read().decode("utf-8")
                result = json.loads(response_body)
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="ignore")
            try:
                error_json = json.loads(body)
                error_message = error_json.get("error", {}).get("message", body)
                error_type = error_json.get("error", {}).get("type", "")
                error_code = error_json.get("error", {}).get("code", "")
            except Exception:
                error_message = body
                error_type = ""
                error_code = ""

            if error_code == "credit_balance_exhausted" or error_type == "insufficient_quota":
                raise RuntimeError("OpenAI 餘額不足，請先補足額度；或改為不勾選外掛辨識，改用本地 OCR）") from exc

            raise RuntimeError(f"ChatGPT API 呼叫失敗：{error_message}") from exc

        content = result["choices"][0]["message"]["content"]
        if isinstance(content, list):
            content = "".join(part.get("text", "") for part in content)

        return self._parse_chatgpt_csv_output(content)

    def _parse_chatgpt_csv_output(self, content):
        candidate = content.strip()
        if "```" in candidate:
            candidate = re.sub(r"```(?:csv)?", "", candidate).strip()

        text_lines = [line.strip() for line in candidate.splitlines() if line.strip()]
        if not text_lines:
            raise RuntimeError("ChatGPT 未回傳可解析的表格內容")

        if "," not in text_lines[0] and "\t" not in text_lines[0]:
            return pd.DataFrame({"單列內容": [" ".join(text_lines)]}), "\n".join(text_lines)

        try:
            csv_io = io.StringIO("\n".join(text_lines))
            df = pd.read_csv(csv_io, encoding="utf-8-sig")
            return df, "\n".join(text_lines)
        except Exception:
            return pd.DataFrame({"單列內容": [" ".join(text_lines)]}), "\n".join(text_lines)

    def _normalize_table_dataframe(self, df):
        if df is None or df.empty:
            return None

        normalized = df.copy()
        normalized = normalized.dropna(how="all")
        normalized = normalized.loc[:, ~normalized.columns.isna()]
        normalized = normalized.fillna("")

        if normalized.shape[1] <= 1:
            return None

        cleaned_cols = []
        for col in normalized.columns:
            col_name = str(col).strip().replace("\n", " ")
            col_name = re.sub(r"\s+", " ", col_name)
            cleaned_cols.append(col_name)
        normalized.columns = cleaned_cols

        for col in normalized.columns:
            if any(keyword in col.lower() for keyword in ["working", "material", "vc", "fz", "ap"]):
                pass

        if len(normalized.columns) >= 4:
            normalized.columns = ["Working Material", "Vc", "fz", "ap"][: len(normalized.columns)]

        for idx, row in normalized.iterrows():
            for col in normalized.columns:
                value = row[col]
                if isinstance(value, str):
                    value = value.strip().replace("\u2013", "-").replace("~", "-")
                    value = re.sub(r"\s+", " ", value)
                    if value and value not in {"None", "nan"}:
                        normalized.at[idx, col] = value

        return normalized.reset_index(drop=True)

    def _run_img2table_on_image(self, file_path):
        try:
            ocr_engine = TesseractOCR(lang="chi_sim+eng", psm=11)
            if TESSERACT_PATH:
                ocr_engine = TesseractOCR(lang="chi_sim+eng", psm=11)
            img_doc = Img2TableImage(file_path)
            extracted_tables = img_doc.extract_tables(
                ocr=ocr_engine,
                implicit_rows=True,
                implicit_columns=True,
                borderless_tables=True,
                min_confidence=40,
                max_workers=1
            )

            if not extracted_tables:
                return None

            table_dfs = []
            for table in extracted_tables:
                try:
                    table_df = table.df
                    if table_df is not None and not table_df.empty:
                        normalized = self._normalize_table_dataframe(table_df)
                        if normalized is not None and not normalized.empty:
                            table_dfs.append(normalized)
                except Exception:
                    continue

            if not table_dfs:
                return None

            merged = pd.concat(table_dfs, ignore_index=True)
            return merged
        except Exception:
            return None

    def display_image(self, file_path):
        try:
            img = Image.open(file_path)
            img = img.resize((320, 240))
            photo = ImageTk.PhotoImage(img)
            self.image_label.configure(image=photo, text="")
            self.image_label.image = photo
        except Exception as exc:
            messagebox.showerror("圖片錯誤", str(exc))

    def show_preview_placeholder(self, text):
        self.image_label.configure(text=text, image="")

    def ocr_image_to_table(self, file_path):
        try:
            if self._use_chatgpt_ocr():
                self.status_var.set("正在使用 ChatGPT 辨識圖片...")
                try:
                    df, text = self._call_chatgpt_ocr(file_path)
                    self.status_var.set("已使用 ChatGPT 取得圖片表格內容")
                    return df, text
                except RuntimeError as exc:
                    if "餘額不足" in str(exc):
                        messagebox.showwarning("外掛辨識失敗", f"{exc}\n\n將自動改用本地 OCR 辨識。")
                        self.status_var.set("外掛辨識額度不足，改用本地 OCR")
                    else:
                        messagebox.showerror("OCR 錯誤", str(exc))
                        return pd.DataFrame({"辨識結果": ["OCR 失敗"]}), "OCR 失敗"

            keywords = self._parse_keywords()
            required_keywords = self._required_keywords()
            img = Image.open(file_path)
            enhanced_img = self._preprocess_for_ocr(img)
            image_data = self._run_ocr_passes(enhanced_img, keywords, required_keywords)
            local_df = self._build_dataframe_from_ocr(image_data, keywords, required_keywords)
            local_df = self._ensure_required_keywords_in_dataframe(local_df, required_keywords)
            local_text = self._text_from_dict(image_data)

            table_df = self._run_img2table_on_image(file_path)
            if table_df is not None and not table_df.empty:
                table_df = self._normalize_table_dataframe(table_df)
                if table_df is not None and not table_df.empty:
                    self.status_var.set("已搭配 img2table 做表格結構化補強")
                    return table_df, local_text

            return local_df, local_text
        except Exception as exc:
            messagebox.showerror("OCR 錯誤", f"請確認 Tesseract 已安裝，或檢查 OpenAI API 設定：{exc}")
            return pd.DataFrame({"辨識結果": ["OCR 失敗"]}), "OCR 失敗"

    def ocr_pdf_to_table(self, file_path):
        try:
            pages = convert_from_path(file_path)
            page_texts = []
            page_dataframes = []
            for page in pages:
                enhanced_page = self._preprocess_for_ocr(page)
                image_data = self._run_ocr_passes(enhanced_page)
                page_texts.append(self._text_from_dict(image_data))
                page_dataframes.append(self._build_dataframe_from_ocr(image_data))

            combined_df = pd.concat(page_dataframes, ignore_index=True) if page_dataframes else pd.DataFrame()
            return combined_df, "\n\n---Page Split---\n\n".join(page_texts)
        except Exception as exc:
            messagebox.showerror("PDF OCR 錯誤", f"請確認 Tesseract 已安裝：{exc}")
            return pd.DataFrame({"辨識結果": ["PDF OCR 失敗"]}), "PDF OCR 失敗"

    def _enhance_table_lines_with_hough(self, img):
        try:
            gray = np.array(img.convert("L"))
            if gray.size == 0:
                return img

            blur = cv2.GaussianBlur(gray, (3, 3), 0)
            edges = cv2.Canny(blur, 50, 150, apertureSize=3)
            lines = cv2.HoughLinesP(
                edges,
                rho=1,
                theta=np.pi / 180,
                threshold=80,
                minLineLength=40,
                maxLineGap=12
            )

            if lines is None or len(lines) == 0:
                return img

            line_mask = np.zeros_like(gray)
            for line in lines:
                x1, y1, x2, y2 = line[0]
                if abs(y2 - y1) <= 4 or abs(x2 - x1) <= 4:
                    cv2.line(line_mask, (x1, y1), (x2, y2), 255, 2)

            if np.count_nonzero(line_mask) == 0:
                return img

            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
            line_mask = cv2.dilate(line_mask, kernel, iterations=1)
            combined = cv2.addWeighted(gray, 0.88, line_mask, 0.55, 0)
            combined = cv2.medianBlur(combined, 3)
            return Image.fromarray(combined.astype("uint8"), mode="L")
        except Exception:
            return img

    def _run_ocr_passes(self, enhanced_img, keywords=None, required_keywords=None):
        configs = [
            "--oem 3 --psm 6 -c preserve_interword_spaces=1",
            "--oem 3 --psm 11 -c preserve_interword_spaces=1",
            "--oem 3 --psm 4 -c preserve_interword_spaces=1"
        ]

        best_data = None
        best_score = -1
        for config in configs:
            data = pytesseract.image_to_data(
                enhanced_img,
                lang="chi_sim+eng",
                output_type=Output.DICT,
                config=config
            )
            score = self._score_ocr_data(data, keywords, required_keywords)
            if score > best_score:
                best_score = score
                best_data = data

        if best_data is None:
            return pytesseract.image_to_data(
                enhanced_img,
                lang="chi_sim+eng",
                output_type=Output.DICT,
                config="--oem 3 --psm 6"
            )

        return best_data

    def _matches_keyword(self, keyword_list, text):
        if not keyword_list:
            return False
        lower_text = text.lower()
        for keyword in keyword_list:
            if keyword.lower() in lower_text:
                return True
        return False

    def _score_ocr_data(self, data, keywords=None, required_keywords=None):
        score = 0.0
        confidences = []
        text_count = 0
        keyword_list = [kw.lower() for kw in (keywords or []) if kw.strip()]
        required_list = [kw.lower() for kw in (required_keywords or []) if kw.strip()]
        for idx, text in enumerate(data["text"]):
            cell_text = text.strip()
            if not cell_text:
                continue
            try:
                conf = float(data["conf"][idx])
            except Exception:
                conf = 0

            keyword_hit = self._matches_keyword(keyword_list, cell_text)
            required_hit = self._matches_keyword(required_list, cell_text)
            if conf > 0:
                confidences.append(conf)
                score += conf
            text_count += 1

            if required_hit:
                score += 200
            elif keyword_hit:
                score += 75
            if any(ch.isdigit() for ch in cell_text):
                score += 12
            if re.search(r"[\u4e00-\u9fff]", cell_text):
                score += 6
            if len(cell_text) >= 2:
                score += 1.5
            if not keyword_hit and len(cell_text) <= 1 and not any(ch.isdigit() for ch in cell_text):
                score -= 10

        if confidences:
            score += sum(confidences) / len(confidences)
        score += text_count * 2.0
        return score

    def _preprocess_for_ocr(self, img):
        img = img.convert("RGB")
        img = ImageOps.autocontrast(img)
        img = ImageEnhance.Contrast(img).enhance(2.5)
        img = ImageEnhance.Sharpness(img).enhance(2.4)
        img = img.filter(ImageFilter.MedianFilter(size=3))

        gray = img.convert("L")
        gray = ImageOps.autocontrast(gray)
        gray = gray.resize((gray.width * 2, gray.height * 2))
        gray = self._enhance_table_lines_with_hough(gray)

        threshold = 170
        binary = gray.point(lambda p: 255 if p > threshold else 0)
        binary = binary.filter(ImageFilter.MaxFilter(3))
        binary = binary.filter(ImageFilter.MinFilter(3))
        return binary

    def _ensure_required_keywords_in_dataframe(self, df, required_keywords=None):
        if df is None or df.empty:
            if required_keywords:
                row = {"關鍵字": required_keywords[0]}
                return pd.DataFrame([row])
            return df

        required_keywords = required_keywords or []
        if not required_keywords:
            return df

        text_values = []
        for _, row in df.iterrows():
            for value in row.values:
                if isinstance(value, str):
                    text_values.append(value)

        missing = []
        for keyword in required_keywords:
            if not any(keyword.lower() in str(value).lower() for value in text_values):
                missing.append(keyword)

        if not missing:
            return df

        for keyword in missing:
            new_row = {df.columns[0]: keyword}
            for col in df.columns[1:]:
                new_row[col] = ""
            df = pd.concat([pd.DataFrame([new_row]), df], ignore_index=True)

        return df

    def _build_dataframe_from_ocr(self, image_data, keywords=None, required_keywords=None):
        tokens = []
        keyword_list = [kw.lower() for kw in (keywords or []) if kw.strip()]
        required_list = [kw.lower() for kw in (required_keywords or []) if kw.strip()]
        for idx, text in enumerate(image_data["text"]):
            cell_text = text.strip()
            if not cell_text:
                continue
            try:
                conf = float(image_data["conf"][idx])
            except Exception:
                conf = 0

            keyword_hit = self._matches_keyword(keyword_list, cell_text)
            required_hit = self._matches_keyword(required_list, cell_text)
            if not keyword_hit and not required_hit and conf < 20:
                continue
            if not keyword_hit and not required_hit and len(cell_text) <= 1 and not any(ch.isdigit() for ch in cell_text):
                continue

            tokens.append({
                "x": int(image_data["left"][idx]),
                "y": int(image_data["top"][idx]),
                "text": cell_text,
                "keyword_hit": keyword_hit or required_hit,
                "conf": conf,
            })

        if not tokens:
            return pd.DataFrame({"辨識結果": ["未能解析表格內容"]})

        rows = defaultdict(list)
        for token in tokens:
            row_key = round(token["y"] / 18) * 18
            rows[row_key].append(token)

        ordered_rows = []
        for key in sorted(rows):
            ordered_rows.append(sorted(rows[key], key=lambda token: token["x"]))

        if self._looks_like_table(ordered_rows):
            groups = self._derive_x_groups(ordered_rows)
            row_data = []
            for row in ordered_rows:
                values = ["" for _ in groups]
                for token in row:
                    group_index = self._assign_to_group(token["x"], groups)
                    values[group_index] = values[group_index] + (" " if values[group_index] else "") + token["text"]
                row_data.append(values)

            columns = [f"欄位{i + 1}" for i in range(len(groups))]
            return pd.DataFrame(row_data, columns=columns)

        single_line_rows = []
        for row in ordered_rows:
            text_line = " ".join(token["text"] for token in row)
            single_line_rows.append([text_line])

        return pd.DataFrame({"單列內容": single_line_rows})

    def _looks_like_table(self, ordered_rows):
        if len(ordered_rows) < 2:
            return False

        distinct_x_groups = set()
        for row in ordered_rows:
            if len(row) < 2:
                continue
            distinct_x_groups.update(round(token["x"] / 90) * 90 for token in row)

        if len(distinct_x_groups) < 2:
            return False

        multi_column_rows = sum(1 for row in ordered_rows if len(row) >= 2)
        return multi_column_rows >= 2

    def _derive_x_groups(self, ordered_rows):
        all_x = sorted({round(token["x"] / 90) * 90 for row in ordered_rows for token in row})
        if len(all_x) <= 1:
            return [all_x[0]] if all_x else [0]

        groups = []
        current_group = [all_x[0]]
        for value in all_x[1:]:
            if value - current_group[-1] <= 90:
                current_group.append(value)
            else:
                groups.append(sum(current_group) / len(current_group))
                current_group = [value]
        groups.append(sum(current_group) / len(current_group))
        return groups

    def _assign_to_group(self, x_value, groups):
        if len(groups) == 1:
            return 0
        return min(range(len(groups)), key=lambda idx: abs(groups[idx] - x_value))

    def _text_from_dict(self, image_data):
        lines = []
        for idx, text in enumerate(image_data["text"]):
            cell_text = text.strip()
            if not cell_text:
                continue
            try:
                conf = float(image_data["conf"][idx])
            except Exception:
                conf = 0
            if conf < 20:
                continue
            lines.append(cell_text)
        return "\n".join(lines)

    def open_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔",
            filetypes=[("Excel 檔", "*.xlsx *.xls *.csv")]
        )
        if not file_path:
            return

        self.excel_path = file_path
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext == ".csv":
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            self.current_df = df
            self.show_excel_table(df)
            self.status_var.set(f"已開啟 Excel：{file_path}")
        except Exception as exc:
            messagebox.showerror("開啟 Excel 錯誤", str(exc))

    def show_excel_table(self, df):
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.row_map = {}
        columns = list(df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=120, anchor="center")

        for row_index, row in df.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row]
            item_id = f"row_{row_index}"
            self.row_map[item_id] = row_index
            self.tree.insert("", tk.END, iid=item_id, values=values)

    def sort_by_column(self, column_name):
        if self.current_df is None:
            return

        ascending = self.sort_state.get(column_name, True)
        self.current_df = self.current_df.sort_values(
            by=column_name,
            ascending=ascending,
            na_position="last"
        ).reset_index(drop=True)
        self.sort_state[column_name] = not ascending
        self.show_excel_table(self.current_df)
        self.status_var.set(f"已依 {column_name} 排序")

    def start_inline_edit(self, event):
        if self.current_df is None:
            return

        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id or not column_id:
            return

        column_index = int(column_id.replace("#", "")) - 1
        col_name = self.current_df.columns[column_index]
        item_text = self.tree.item(row_id, "values")[column_index]

        bbox = self.tree.bbox(row_id, column_id)
        if not bbox:
            return

        x = bbox[0] + 4
        y = bbox[1] + 2
        width = bbox[2] - 6
        height = bbox[3] - 4

        self.inline_editor.place(x=x, y=y, width=width, height=height)
        self.inline_editor.delete(0, tk.END)
        self.inline_editor.insert(0, item_text)
        self.inline_editor.focus_set()
        self.editor_target = (row_id, col_name)

    def commit_inline_edit(self, event):
        if not self.editor_target:
            return "break"

        row_id, col_name = self.editor_target
        row_index = self.row_map.get(row_id)
        if row_index is None:
            self.inline_editor.place_forget()
            return "break"

        new_value = self.inline_editor.get()
        self.current_df.at[int(row_index), col_name] = new_value
        self.tree.item(row_id, values=["" if pd.isna(v) else str(v) for v in self.current_df.loc[int(row_index)]])
        self.inline_editor.place_forget()
        self.editor_target = None
        self.status_var.set("已更新暫存表格內容，按 C 可覆寫到 Excel")
        return "break"

    def cancel_inline_edit(self, event):
        self.inline_editor.place_forget()
        self.editor_target = None
        return "break"

    def add_row(self):
        if self.current_df is None:
            messagebox.showwarning("提醒", "請先開啟一個 Excel 檔")
            return

        blank_row = {col: "" for col in self.current_df.columns}
        self.current_df.loc[len(self.current_df)] = blank_row
        self.show_excel_table(self.current_df)
        self.status_var.set("已新增空白列")

    def delete_selected_row(self):
        if self.current_df is None:
            messagebox.showwarning("提醒", "請先開啟一個 Excel 檔")
            return

        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提醒", "請先選擇要刪除的列")
            return

        row_id = selected[0]
        row_index = self.row_map.get(row_id)
        if row_index is None:
            return

        self.current_df = self.current_df.drop(index=int(row_index)).reset_index(drop=True)
        self.show_excel_table(self.current_df)
        self.status_var.set("已刪除選取列")

    def on_tree_double_click(self, event):
        if self.current_df is None:
            return

        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id or not column_id:
            return

        column_index = int(column_id.replace("#", "")) - 1
        col_name = self.current_df.columns[column_index]

        item_text = self.tree.item(row_id, "values")[column_index]
        row_index = self.row_map.get(row_id, row_id)
        self.editing_row = row_id
        self.editing_col = col_name

        popup = tk.Toplevel(self.window)
        popup.title("編輯儲存格")
        popup.geometry("280x120")

        ttk.Label(popup, text=f"欄位：{col_name}").pack(pady=(10, 5))
        entry = ttk.Entry(popup, width=30)
        entry.insert(0, item_text)
        entry.pack(pady=5)

        def save_value():
            new_value = entry.get()
            self.current_df.at[int(row_index), col_name] = new_value
            updated_values = ["" if pd.isna(v) else str(v) for v in self.current_df.loc[int(row_index)]]
            self.tree.item(row_id, values=updated_values)
            popup.destroy()
            self.status_var.set("已更新暫存表格內容，按 C 可覆寫到 Excel")

        ttk.Button(popup, text="確定", command=save_value).pack(pady=10)

    def save_excel_file(self):
        if self.current_df is None or not self.excel_path:
            messagebox.showwarning("提醒", "請先開啟一個 Excel 檔")
            return

        try:
            ext = os.path.splitext(self.excel_path)[1].lower()
            if ext == ".csv":
                self.current_df.to_csv(self.excel_path, index=False, encoding="utf-8-sig")
            else:
                self.current_df.to_excel(self.excel_path, index=False)
            self.status_var.set("Excel 已成功儲存")
            messagebox.showinfo("儲存成功", "Excel 已成功儲存")
        except Exception as exc:
            messagebox.showerror("儲存錯誤", str(exc))


if __name__ == "__main__":
    # 執行前請確認 Tesseract 已安裝，或將 tesseract 可執行檔路徑加入環境變數
    root = tk.Tk()
    app = ExcelOcrApp(root)
    root.mainloop()