import sys
import os
import re
import platform
import subprocess
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout,
    QMessageBox, QProgressBar, QLabel
)
from PyQt5.QtCore import Qt

class ExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Загрузить и фильтровать Excel')

        self.button = QPushButton('Загрузить Excel файл', self)
        self.button.clicked.connect(self.load_file)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        self.status_label = QLabel('Ожидание файла...', self)

        layout = QVBoxLayout()
        layout.addWidget(self.button)
        layout.addWidget(self.status_label)
        layout.addWidget(self.progress_bar)

        self.setLayout(layout)

    def load_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Excel",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*)",
            options=options
        )

        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)

            if df.shape[1] < 4:
                self.show_message("Ошибка", "Файл должен содержать хотя бы 4 столбца.")
                return

            # ✅ Исправлено предупреждение FutureWarning
            df.loc[:, df.columns[3]] = df.iloc[:, 3].fillna('').astype(str)

            # ✅ Удаляем строки с пустыми значениями во 2 и 4 столбцах
            df = df.dropna(subset=[df.columns[1], df.columns[3]])

            file_dir = os.path.dirname(file_path)
            base_name = os.path.basename(file_path).split('.')[0]
            new_dir = os.path.join(file_dir, f"выборка_{base_name}")

            if not os.path.exists(new_dir):
                os.makedirs(new_dir)

            self.process_data(df, new_dir)

        except Exception as e:
            self.show_message("Ошибка", f"Произошла ошибка при обработке файла: {e}")

    def process_data(self, df, new_dir):
        unique_combinations = df[[df.columns[1], df.columns[3]]].drop_duplicates()

        total_combinations = len(unique_combinations)
        self.progress_bar.setMaximum(total_combinations)
        self.progress_bar.setValue(0)

        saved_files = []

        for i, (_, row) in enumerate(unique_combinations.iterrows()):
            value_2 = row.iloc[0]
            value_4 = row.iloc[1]

            filtered_df = df[(df.iloc[:, 1] == value_2) & (df.iloc[:, 3] == value_4)]

            val2 = "" if pd.isna(value_2) else self.sanitize_filename(value_2)
            val4 = "" if pd.isna(value_4) else self.sanitize_filename(value_4)
            filter_name = f"{val2}_{val4}".strip("_") or "пустой_фильтр"

            if not filtered_df.empty:
                filename = self.save_file(filtered_df, new_dir, filter_name)
                saved_files.append(filename)

            self.progress_bar.setValue(i + 1)
            percent = int(((i + 1) / total_combinations) * 100)
            self.status_label.setText(f"Обрабатывается: {filter_name}. Прогресс: {percent}% ({i + 1} из {total_combinations})")
            QApplication.processEvents()

        log_file_path = os.path.join(new_dir, "log.txt")
        with open(log_file_path, "w", encoding="utf-8") as log_file:
            for name in saved_files:
                log_file.write(name + "\n")

        self.status_label.setText("Готово.")
        self.show_message("Готово", f"Данные успешно отфильтрованы и сохранены!\nСоздан лог: log.txt")

        self.open_folder(new_dir)

    def save_file(self, filtered_df, new_dir, filter_name):
        output_file_path = os.path.join(new_dir, f"{filter_name}.xlsx")
        filtered_df.to_excel(output_file_path, index=False)
        return f"{filter_name}.xlsx"

    def sanitize_filename(self, name):
        name = str(name)
        name = re.sub(r'[\\/*?:"<>|]', "_", name)
        return name.strip()

    def show_message(self, title, message):
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec_()

    def open_folder(self, path):
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", path])
            else:  # Linux
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            print(f"Не удалось открыть папку: {e}")

def main():
    app = QApplication(sys.argv)
    ex = ExcelApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
