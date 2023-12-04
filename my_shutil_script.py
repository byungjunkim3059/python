import os
import shutil

def create_folder(path):
    try:
        # 폴더 생성
        os.makedirs(path)
        print(f"폴더가 성공적으로 생성되었습니다. 경로: {path}")
    except FileExistsError:
        print(f"폴더가 이미 존재합니다. 경로: {path}")
    except Exception as e:
        print(f"폴더 생성 중 오류가 발생했습니다. 오류 내용: {e}")

# 사용 예제
folder_path = r'C:\Users\bnj30\Desktop\출고서류\10월/새로운폴더'

create_folder(folder_path)


if __name__ == '__main__':
    shutil.copytree(
        r'C:\Users\bnj30\Desktop\10월 출고서류\10월 출고서류\[국내출고] (CNF_원영)ZEROGRAM_ 10_18입고 출고지시서 첨부드립니다 (1)',
        folder_path,
        dirs_exist_ok=True
    )
