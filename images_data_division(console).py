import os
import cv2
import numpy as np
from pyzbar.pyzbar import decode
from PIL import Image
from tkinter import Tk, filedialog
import shutil
import re  # 정규 표현식 라이브러리 추가

# === 폴더 선택 ===
def select_folder(prompt):
    root = Tk()
    root.withdraw()  # GUI 창 숨김
    folder = filedialog.askdirectory(title=prompt)
    return folder

# === 경로 정리 ===
def clean_path(path):
    """불필요한 특수문자를 제거한 폴더 이름 반환"""
    return re.sub(r'[\\/:*?"<>|]', '_', path).strip()

# === 바코드 추출 (10자리로 변환) ===
def extract_barcode(image_path):
    """이미지 파일에서 바코드를 읽어와 10자리로 반환합니다."""
    if not os.path.exists(image_path):
        print(f"[WARN] 유효하지 않은 경로: {image_path}")
        return None

    try:
        pil_image = Image.open(image_path).convert('L')  # 흑백 변환
        image = np.array(pil_image)

        # 이미지 전처리
        _, thresh = cv2.threshold(image, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

        # 바코드 감지
        barcodes = decode(thresh)
        if barcodes:
            barcode_data = barcodes[0].data.decode("utf-8")[:10]
            print(f"[INFO] 바코드 추출 성공: {barcode_data}")
            return barcode_data
    except Exception as e:
        print(f"[ERROR] 바코드 추출 중 오류 발생: {e}")

    print(f"[INFO] 바코드 감지 실패: {image_path}")
    return None

# === 파일 이동 ===
def move_files_to_folder(files, destination_root, folder_name):
    """파일을 지정된 폴더로 이동"""
    # 폴더 이름만 정리
    folder_name = clean_path(folder_name)
    destination_folder = os.path.join(destination_root, folder_name)

    os.makedirs(destination_folder, exist_ok=True)  # 폴더 생성
    for file in files:
        try:
            target_path = os.path.join(destination_folder, os.path.basename(file))
            shutil.move(file, target_path)
            print(f"[INFO] 파일 이동 성공: {file} -> {target_path}")
        except Exception as e:
            print(f"[ERROR] 파일 이동 중 오류 발생: {e}")

# === 실행 ===
if __name__ == "__main__":
    source_folder = select_folder("원본 폴더를 선택하세요")
    destination_root = select_folder("파일을 저장할 폴더를 선택하세요")

    files = sorted(os.listdir(source_folder))
    current_barcode = None
    files_to_move = []

    for file in files:
        file_path = os.path.join(source_folder, file)

        barcode = extract_barcode(file_path)
        if barcode:
            # 이전 상품 이미지 이동
            if current_barcode and files_to_move:
                print(f"[INFO] {current_barcode} 폴더로 파일 이동 중...")
                move_files_to_folder(files_to_move, destination_root, current_barcode)
                files_to_move = []

            current_barcode = barcode
            files_to_move.append(file_path)
        else:
            if current_barcode:
                files_to_move.append(file_path)

    # 마지막 상품 처리
    if current_barcode and files_to_move:
        print(f"[INFO] {current_barcode} 폴더로 파일 이동 중...")
        move_files_to_folder(files_to_move, destination_root, current_barcode)

    print("[INFO] 작업 완료")
