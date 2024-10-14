import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QVBoxLayout, QWidget, QPushButton, QTextEdit, QMessageBox


class ExcelLoaderApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.initUI()

    def initUI(self):
        # 창의 타이틀 설정
        self.setWindowTitle('Stock')

        self.resize(800, 400)

        # 레이아웃 설정
        layout = QVBoxLayout()

        # 파일 로드 버튼 생성
        self.load_button = QPushButton('재고 불러오기', self)
        self.load_button.clicked.connect(self.load_excel_file)
        layout.addWidget(self.load_button)

        # 파일 삭제 버튼 생성
        self.delete_button = QPushButton('파일 삭제하기', self)
        self.delete_button.clicked.connect(self.delete_excel_file)
        layout.addWidget(self.delete_button)

        # 결과 출력용 텍스트박스 생성
        self.result_text = QTextEdit(self)
        self.result_text.setReadOnly(True)  # 읽기 전용으로 설정
        layout.addWidget(self.result_text)

        # 레이아웃을 메인 윈도우에 설정
        self.setLayout(layout)

    def load_excel_file(self):
        # 디렉토리 설정 (현재 디렉토리)
        directory = './'
        folder_path = './상품'
        self.result_text.clear()

        # 디렉토리에서 파일을 찾는 부분
        for filename in os.listdir(directory):
            if filename.startswith("현재창고관리") and filename.endswith(".xls"):
                self.file_path = os.path.join(directory, filename)
                
                # 결과 출력
                # self.result_text.append(f"불러오는 파일: {file_path}")
                
                # Excel 파일을 Pandas DataFrame으로 불러오기
                df = pd.read_excel(self.file_path)

                # 데이터프레임의 내용을 표시
                # self.result_text.append(f"DataFrame 정보:\n{df.head()}")
                break
        

        try:
            new_df = df[['상품분류', '상품명','현재창고재고']]

            for filename in os.listdir(folder_path):
                if filename.endswith('.txt'):  # 텍스트 파일만 선택
                    cupname = os.path.splitext(filename)[0]

                    product_names = []  # 각 파일에 대한 상품명 리스트 초기화
                    with open(os.path.join(folder_path, filename), 'r', encoding='utf-8') as file:
                        # 각 파일의 상품명을 읽어서 리스트에 추가
                        for line in file:
                            product_names.append(line.strip())

                    # 4. 재고 값 합산
                    total_stock = 0

                    for product in product_names:
                        stock_value = new_df.loc[new_df['상품명'] == product, '현재창고재고']
                        if not stock_value.empty:
                            if cupname == "1L컵":
                                total_stock += stock_value.values[0]
                            else:
                                total_stock += (stock_value.values[0] - 1000)
                        else:
                            self.result_text.append(f'{product}는 재고에 없습니다!')

                    # 결과 출력
                    self.result_text.append(f'{os.path.splitext(filename)[0]}: {total_stock}')
        except:
            self.result_text.clear()
            self.result_text.append("재고 파일이 없습니다!")

    def delete_excel_file(self):
            if self.file_path:
                # 사용자에게 삭제 여부를 확인하는 메시지 박스
                reply = QMessageBox.question(self, '파일 삭제 확인', 
                                            f'{self.file_path} 파일을 삭제하시겠습니까?', 
                                            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

                if reply == QMessageBox.Yes:
                    try:
                        # 파일 삭제
                        os.remove(self.file_path)
                        # self.result_text.append(f"{self.file_path} 파일이 삭제되었습니다.")
                        self.file_path = None  # 파일 경로를 None으로 초기화
                    except Exception as e:
                        self.result_text.clear()
                        self.result_text.append(f"파일 삭제 실패: {str(e)}")
                else:
                    self.result_text.clear()
                    self.result_text.append("파일 삭제가 취소되었습니다.")
            else:
                self.result_text.clear()
                self.result_text.append("재고 파일이 없습니다!")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelLoaderApp()
    ex.show()
    sys.exit(app.exec_())
