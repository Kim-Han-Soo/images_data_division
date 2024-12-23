import os
import cv2
import numpy as np
from pyzbar.pyzbar import decode
from PIL import Image
import shutil
import re
import json
import pandas as pd  # 엑셀 파일 처리를 위한 라이브러리
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# === 경로 정리 ===
def clean_path(path):
    """경로에 포함된 특수문자 및 공백 제거"""
    return re.sub(r'[\\/:*?"<>|]', '_', path).strip()

# === 바코드 추출 ===
def extract_barcode(image_path):
    """이미지 파일에서 바코드를 읽어와 10자리로 반환"""
    try:
        pil_image = Image.open(image_path).convert('L')
        image = np.array(pil_image)

        # Otsu Thresholding
        _, thresh = cv2.threshold(image, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

        barcodes = decode(thresh)
        for barcode in barcodes:
            barcode_data = barcode.data.decode("utf-8").strip()
            if len(barcode_data) >= 10:
                return barcode_data[:10]
    except Exception as e:
        print(f"[ERROR] 바코드 추출 오류: {e}")
    return None

# === 파일 이동 ===
def move_files_to_folder(files, destination_root, folder_name):
    """파일을 지정된 폴더로 이동"""
    folder_name = clean_path(folder_name)
    destination_folder = os.path.join(destination_root, folder_name)

    os.makedirs(destination_folder, exist_ok=True)
    for file in files:
        try:
            shutil.move(file, os.path.join(destination_folder, os.path.basename(file)))
        except Exception as e:
            print(f"[ERROR] 파일 이동 오류: {e}")

# === 작업 이력 로드 및 저장 ===
def load_log(log_file):
    """작업 이력을 로드"""
    if os.path.exists(log_file):
        with open(log_file, "r") as f:
            return json.load(f)
    return {}

def save_log(log_file, log_data):
    """작업 이력을 저장"""
    with open(log_file, "w") as f:
        json.dump(log_data, f, indent=4)

# === 폴더 작업 ===
def process_folder(source_folder, destination_root, log_file):
    """폴더 내 파일 분류"""
    log_data = load_log(log_file)
    files = sorted(os.listdir(source_folder))
    current_barcode = None
    files_to_move = []

    for file in files:
        file_path = os.path.join(source_folder, file)

        # 작업 이력 확인
        if file_path in log_data:
            print(f"[INFO] 이미 처리된 파일: {file_path}")
            continue

        barcode = extract_barcode(file_path)
        if barcode:
            if current_barcode and files_to_move:
                move_files_to_folder(files_to_move, destination_root, current_barcode)
                files_to_move = []

            current_barcode = barcode
            files_to_move.append(file_path)
        else:
            if current_barcode:
                files_to_move.append(file_path)

        # 작업 로그 업데이트
        log_data[file_path] = barcode or "Uncategorized"
        save_log(log_file, log_data)

    if current_barcode and files_to_move:
        move_files_to_folder(files_to_move, destination_root, current_barcode)

# === 엑셀 경로 처리 ===
def process_excel_folders(excel_file, log_file):
    """엑셀에서 경로를 읽어와 처리"""
    if not os.path.exists(excel_file):
        print(f"[ERROR] 엑셀 파일이 존재하지 않습니다: {excel_file}")
        return

    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"[ERROR] 엑셀 파일을 읽는 중 오류 발생: {e}")
        return

    if "SourceFolder" not in df.columns or "DestinationFolder" not in df.columns:
        print(f"[ERROR] 엑셀 파일에 'SourceFolder'와 'DestinationFolder' 열이 필요합니다.")
        return

    for _, row in df.iterrows():
        source_folder = row["SourceFolder"]
        destination_root = row["DestinationFolder"]

        if not os.path.exists(source_folder) or not os.path.exists(destination_root):
            print(f"[WARN] 유효하지 않은 경로: {source_folder} -> {destination_root}")
            continue

        print(f"[INFO] 작업 시작: {source_folder} -> {destination_root}")
        process_folder(source_folder, destination_root, log_file)
        print(f"[INFO] 작업 완료: {source_folder}")

# === GUI ===
class ManualModeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("수동 모드: 바코드 기반 파일 분류")

        # 상태 변수
        self.source_folder = None
        self.destination_root = None
        self.log_file = "processing_log.json"

        # GUI 구성
        self.setup_gui()

    def setup_gui(self):
        """GUI 레이아웃 구성"""
        tk.Label(self.root, text="수동 모드", font=("Arial", 16)).pack(pady=10)

        # 원본 폴더 선택
        tk.Button(self.root, text="원본 폴더 선택", command=self.select_source_folder).pack(pady=5)
        self.source_label = tk.Label(self.root, text="원본 폴더: 미설정", fg="blue")
        self.source_label.pack()

        # 저장 폴더 선택
        tk.Button(self.root, text="저장 폴더 선택", command=self.select_destination_folder).pack(pady=5)
        self.destination_label = tk.Label(self.root, text="저장 폴더: 미설정", fg="blue")
        self.destination_label.pack()

        # 작업 실행 버튼
        tk.Button(self.root, text="작업 실행", command=self.run_process).pack(pady=5)

        # 종료 버튼
        tk.Button(self.root, text="종료", command=self.root.quit).pack(pady=10)

    def select_source_folder(self):
        """원본 폴더 선택"""
        folder = filedialog.askdirectory(title="원본 폴더를 선택하세요")
        if folder:
            self.source_folder = folder
            self.source_label.config(text=f"원본 폴더: {folder}")

    def select_destination_folder(self):
        """저장 폴더 선택"""
        folder = filedialog.askdirectory(title="저장 폴더를 선택하세요")
        if folder:
            self.destination_root = folder
            self.destination_label.config(text=f"저장 폴더: {folder}")

    def run_process(self):
        """작업 실행"""
        if not self.source_folder or not self.destination_root:
            messagebox.showwarning("경고", "원본 폴더와 저장 폴더를 모두 선택하세요.")
            return

        try:
            process_folder(self.source_folder, self.destination_root, self.log_file)
            messagebox.showinfo("완료", "작업이 완료되었습니다!")
        except Exception as e:
            messagebox.showerror("오류", f"작업 중 오류 발생: {e}")

# === 실행 ===
if __name__ == "__main__":
    mode = input("작업 모드를 선택하세요 (1: 엑셀 모드, 2: 수동 모드): ").strip()
    if mode == "1":
        excel_file = "folders.xlsx"
        log_file = "processing_log.json"
        process_excel_folders(excel_file, log_file)
    elif mode == "2":
        root = tk.Tk()
        app = ManualModeApp(root)
        root.mainloop()
    else:
        print("[INFO] 잘못된 입력입니다. 프로그램을 종료합니다.")
