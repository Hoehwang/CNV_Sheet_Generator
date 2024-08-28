from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QThread, pyqtSignal
import re, sys, os, subprocess, psutil
import time, requests
import xml.etree.ElementTree as ET
import pandas as pd
import disease_info_hardcoded
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def file_start(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])


def force_kill_process_using_file(file_path):
    try:
        # 파일 삭제 시도
        os.remove(file_path)
        print(f"{file_path} 파일이 성공적으로 삭제되었습니다.")
    except PermissionError as e:
        print(f"파일 삭제 실패: {e}")
        print("파일이 사용 중입니다. 파일을 사용하는 프로세스를 찾습니다...")

        # 파일을 사용하는 프로세스 찾기
        for proc in psutil.process_iter(['pid', 'name', 'open_files']):
            try:
                if proc.info['open_files']:
                    for open_file in proc.info['open_files']:
                        if open_file.path == file_path:
                            print(f"PID {proc.info['pid']}의 {proc.info['name']} 프로세스가 파일을 사용 중입니다.")
                            # 프로세스 강제 종료
                            proc.kill()
                            print(f"PID {proc.info['pid']} 프로세스를 강제 종료했습니다.")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

        # 프로세스 종료 후 다시 파일 삭제 시도
        time.sleep(1)  # 프로세스 종료가 완료되기를 기다림
        try:
            os.remove(file_path)
            print(f"{file_path} 파일이 성공적으로 삭제되었습니다.")
        except Exception as e:
            print(f"파일 삭제 실패: {e}")

def find_and_kill_process_using_file(file_path):
    """Find the process that is using the given file and kill it."""
    for proc in psutil.process_iter(['pid', 'name', 'open_files']):
        try:
            if proc.info['open_files']:
                for file in proc.info['open_files']:
                    if file.path == file_path:
                        print(f"Killing process {proc.info['name']} (PID: {proc.info['pid']}) using the file.")
                        proc.kill()
                        return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue
    return False


form = resource_path(r'.\CNV.ui')
form_class = uic.loadUiType(form)[0]

class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, df, download_dir):
        super().__init__()
        self.df = df
        self.download_dir = download_dir

    def run(self):
        options = Options()
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36')
        options.add_experimental_option("prefs", {
            "download.default_directory": self.download_dir,  # 다운로드 위치 설정
            "download.prompt_for_download": False,  # 다운로드 대화 상자 비활성화
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

        driver.get('https://www.orphadata.com/phenotypes/')
        time.sleep(4)

        driver.execute_script("document.body.style.zoom='80%'")
        driver.execute_script("window.scrollBy(0, 480);")

        e = driver.find_element(By.CSS_SELECTOR,'#custom-2 > div > div > div > div > div > div > div > div.h-accordion-item.h-accordion-item-first-child.style-5355.style-local-1744-c15.position-relative.h-element > a')
        driver.execute_script("arguments[0].click();", e)
        currentNum = int(driver.find_element(by=By.CSS_SELECTOR, value='#tablepress-16 > tbody > tr.row-2.even > td.column-3').text)

        if currentNum != len(self.df['Name'].unique()):
            xml_button = driver.find_element(by=By.CSS_SELECTOR,value='#tablepress-16 > tbody > tr.row-2.even > td.column-2 > a')
            driver.execute_script("arguments[0].click();", xml_button)
            while True:
                files = os.listdir(self.download_dir)
                if any(file.startswith('en_product') for file in files):
                    driver.quit()
                    break
                else:
                    time.sleep(1)

            tree = ET.parse(rf"{self.download_dir}\en_product4.xml")
            root = tree.getroot()

            data = {
                "Name": [],
                "OrphaCode": [],
                "HPOTerm": [],
                "HPOFrequency": [],
                "HPOKor": []  # HPOKor 열 추가
            }

            for disorder_set in root.findall('.//HPODisorderSetStatus'):
                disorder = disorder_set.find('Disorder')
                disorder_name = disorder.find('Name').text
                orpha_code = disorder.find('OrphaCode').text

                for association in disorder.findall('.//HPODisorderAssociation'):
                    hpo_term = association.find('.//HPOTerm').text
                    hpo_frequency = association.find('.//HPOFrequency/Name').text

                    data['Name'].append(disorder_name)
                    data['OrphaCode'].append(orpha_code)
                    data['HPOTerm'].append(hpo_term)
                    data['HPOFrequency'].append(hpo_frequency)
                    data['HPOKor'].append(None)  # 초기에는 None으로 설정

            new_df = pd.DataFrame(data)

            csv_path = './CNV_info.csv'
            xml_path = './en_product4.xml'
            hpo_translation_df = pd.read_csv(csv_path)

            unique_hpoterms = new_df['HPOTerm'].unique()

            hpo_kor_dict = {}
            total_terms = len(unique_hpoterms)
            for i, hpo_term in enumerate(unique_hpoterms):
                matching_row = hpo_translation_df[hpo_translation_df['HPOTerm'] == hpo_term]
                if not matching_row.empty:
                    hpo_kor_dict[hpo_term] = matching_row.iloc[0]['HPOKor']
                # 진행 상황 업데이트
                progress_value = int((i + 1) / total_terms * 100)
                self.progress.emit(progress_value)

            new_df['HPOKor'] = new_df['HPOTerm'].map(hpo_kor_dict)
            if os.path.exists(csv_path):
                force_kill_process_using_file(csv_path)
            os.remove(xml_path)

            new_df.to_csv(csv_path)

            self.df = pd.read_csv(csv_path)
            self.df['Name_lower'] = self.df['Name'].str.lower()

            self.hpo_dict = self.df[['HPOTerm', 'HPOKor']].drop_duplicates()
            self.hpo_dict = self.hpo_dict.set_index('HPOTerm')['HPOKor'].to_dict()
            self.term_list = list(self.hpo_dict.keys())

        self.finished.emit()  # 작업 완료 신호 송출

def git_csv_downloader(file_name, url):
    if not os.path.isfile(file_name):
        print(f"{file_name} not found. Downloading from GitHub...")

        # GitHub에서 파일을 다운로드
        response = requests.get(url)

        if response.status_code == 200:
            # 파일 저장
            with open(file_name, 'wb') as file:
                file.write(response.content)
            print(f"File downloaded successfully and saved as {file_name}")
        else:
            print(f"Failed to download file. Status code: {response.status_code}")
    else:
        print(f"{file_name} already exists. No download needed.")


def handle_link_clicked(url):
    print(f"Link clicked: {url.toString()}")

class CNV_TestSheet(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.CNVinfo_URL = "https://raw.githubusercontent.com/Hoehwang/searching_report/master/CNV_info.csv"
        self.CNVinfo_file = 'CNV_info.csv'
        self.cvSummary_URL = "https://raw.githubusercontent.com/Hoehwang/searching_report/master/variant_summary.csv"
        self.cvSummary_file = 'variant_summary.csv'
        self.previous_data_URL = "https://raw.githubusercontent.com/Hoehwang/searching_report/master/previous_result_sheet.csv"
        self.previous_data = pd.read_csv(self.previous_data_URL)

        git_csv_downloader(self.CNVinfo_file, self.CNVinfo_URL)
        git_csv_downloader(self.cvSummary_file, self.cvSummary_URL)

        self.df = pd.read_csv(f'./{self.CNVinfo_file}')
        self.df = pd.concat([self.df, self.previous_data], ignore_index=True)
        self.df['Name_lower'] = self.df['Name'].str.lower()

        self.hpo_dict = self.df[['HPOTerm', 'HPOKor']].drop_duplicates()
        self.hpo_dict = self.hpo_dict.set_index('HPOTerm')['HPOKor'].to_dict()
        self.term_list = list(self.hpo_dict.keys())

        self.frequency_order = {'Excluded (0%)': 1, 'Very rare (<4-1%)': 2, 'Occasional (29-5%)': 3,
                                'Frequent (79-30%)': 4, 'Very frequent (99-80%)': 5, 'Obligate (100%)': 6}

        self.synonyms = {''}

        self.recommendBrowser.setOpenExternalLinks(True)
        self.recommendBrowser.anchorClicked.connect(handle_link_clicked)

        self.input_field.textChanged.connect(self.on_text_changed)

        self.caseComboBox.currentIndexChanged.connect(self.update_recommend_browser)


        self.fileBrowse.released.connect(self.TargetfileOpen)
        self.fileOpenButton.released.connect(self.fileOpenStart)
        # self.startButton.released.connect(self.CNV_Start)
        self.actionUpdate.triggered.connect(self.updateCheck)

        self.generateButton.clicked.connect(self.CNV_Save)

        self.center()

        # Initialize progress bar
        self.progressBar.setValue(0)

    def update_recommend_browser(self):
        # 현재 선택된 콤보박스 항목 확인
        selected_case = self.caseComboBox.currentText()

        # 선택된 항목에 해당하는 ClinVar 링크를 recommendBrowser에 표시
        if selected_case in self.clinvar_cases:
            links = self.clinvar_cases[selected_case]
            self.recommendBrowser.clear()
            self.recommendBrowser.append("<p>".join(links))

    def on_text_changed(self):
        input_text = self.input_field.text().strip().lower()

        if not input_text:
            self.dictBrowser.clear()
            return

        # 부분 일치 및 완전 일치 검색 - 제너레이터 사용
        matching_terms = ((term, self.hpo_dict[term]) for term in self.term_list if input_text in term.lower())

        self.dictBrowser.clear()
        result_count = 0
        for term, kor in matching_terms:
            self.dictBrowser.append(f"{kor}")
            result_count += 1

        if result_count == 0:
            self.dictBrowser.append("No matching terms found.")

    def TargetfileOpen(self):
        fname = QFileDialog.getOpenFileName(None, 'Open Txt file', '', "TXT File(*.txt)")
        fname = (str(fname)).split("', '")[0][2:]

        return self.fileName.setText(fname)

    def fileOpenStart(self):
        self.caseComboBox.clear()

        self.clinvar_cases = {}

        self.result_text = ''
        self.links = []
        self.disease_lst = []
        self.counter = 0
        try:
            self.input_data_df = pd.read_csv(self.fileName.text())
            self.total = len(self.input_data_df)
        except KeyError:
            return QMessageBox.warning(self, 'Warning', '인풋 파일을 확인해주세요.')
        for i in range(len(self.input_data_df)):
            self.counter += 1
            self.input_data = self.input_data_df.loc[i]
            self.chromosome = self.input_data['Chr']
            self.disease = self.input_data['disease']
            self.disease_lst.append(self.disease.replace('_dup','dup').replace('_del','del'))
            self.length = self.input_data['Length']
            self.Start = self.input_data['Start']
            self.End = self.input_data['End']
            self.range_length = self.End - self.Start
            self.cytoBand = self.input_data['cytoBand']
            self.copyRatio = self.input_data['CopyRatio']

            if self.copyRatio > 1:
                self.typ = 'Duplication'
            else:
                self.typ = 'Deletion'
            self.summary_df = pd.read_csv(rf'./{self.cvSummary_file}')

            if self.typ == 'Duplication':
                target_df = self.summary_df[self.summary_df['Type'] == 'Duplication']
            else:
                target_df = self.summary_df[self.summary_df['Type'] == 'Deletion']

            # 질병명 Trisomy로 오는 경우 처리

            if target_df[target_df['File'].str.startswith(self.disease.split('_')[0])].empty:
                if self.disease.startswith('Trisomy '):
                    target_df = target_df[target_df['File'].str.startswith(self.disease.replace('Trisomy ','').split('_')[0])]
                else:
                    pass
            else:
                target_df = target_df[target_df['File'].str.startswith(self.disease.split('_')[0])]

            # 겹침 계산 및 퍼센테이지로 변환
            target_df['Overlap'] = target_df.apply(
                lambda row: self.calculate_overlap(self.Start, self.End, row['Start'], row['Stop']), axis=1)
            target_df['OverlapPercentage'] = (target_df['Overlap'] / self.range_length) * 100

            # 각 행의 전체 범위 길이 계산
            target_df['RangeLength'] = target_df['Stop'] - target_df['Start']

            # 주어진 범위 바깥에 있는 부분의 길이 계산
            target_df['OutsideRange'] = target_df['RangeLength'] - target_df['Overlap']

            # 주어진 범위 바깥에 있는 부분의 퍼센트 계산
            target_df['OutsidePercentage'] = (target_df['OutsideRange'] / target_df['RangeLength']) * 100

            # 오버랩 퍼센트 - (아웃사이드 퍼센트/2) = 스코어
            target_df['Score'] = target_df['OverlapPercentage'] - (target_df['OutsidePercentage'] / 2)

            # target_df = target_df[target_df['Score'] >= 80]

            # 'OverlapPercentage'와 'Score' 기준으로 내림차순 정렬
            target_df = target_df.sort_values(by=['OverlapPercentage', 'Score'], ascending=[False, False])

            # VariationID를 제외한 정보가 동일한 행들 찾기
            file_info_df = target_df.groupby(['Type', 'CytogeneticLocation', 'Start', 'Stop']).filter(
                lambda x: len(x) > 1)
            file_info_df = file_info_df.sort_values(by='MostRecentSubmission', ascending=False)

            self.clinvar_cases[f'Case {self.counter}'] = []

            for idx, id in enumerate([i.replace('.xml','').split('_')[-1] for i in list(file_info_df['File'])]):
                self.clinvar_cases[f'Case {self.counter}'].append(f'<a href="https://www.ncbi.nlm.nih.gov/clinvar/variation/{id}/">Case_{self.counter}_ClinVar Link{idx+1}: {id}</a>')

                #### 콤보박스로 변경 후 폐기된 부분 ####
                # self.links.append(f'<a href="https://www.ncbi.nlm.nih.gov/clinvar/variation/{id}/">Case_{self.counter}_ClinVar Link{idx+1}: {id}</a>')
            # if self.counter == self.total:
            #     self.recommendBrowser.setHtml('<p>'.join(self.links))
            # else:
            #     self.links.append('■■■■■■■■■■■■■■■■■■■■■■■■■■■')
            # self.recommendBrowser.setText()

            # 기준치 오버랩 검사
            if self.typ == 'Duplication':
                self.justify_df = disease_info_hardcoded.data_duplication
            else:
                self.justify_df = disease_info_hardcoded.data_deletion
            justify_df = self.justify_df[self.justify_df['disease'] == self.disease]
            # if justify_df.empty:
            #     justify_df
            # print(justify_df)
            justify_start = int(justify_df['start'].iloc[0])
            justify_end = int(justify_df['end'].iloc[0])

            # Start~End와 justify_start~justify_end 범위의 겹치는 부분 계산
            overlap_start = max(self.Start, justify_start)
            overlap_end = min(self.End, justify_end)

            # 겹치는 길이 계산
            self.overlap_length = max(0, overlap_end - overlap_start)
            # 겹치는 비율 계산 (퍼센트로)
            self.justify_overlap = (self.overlap_length / self.range_length) * 100

            self.CNV_Start()

        # print(self.clinvar_cases)
        # 콤보박스 키 추가
        self.caseComboBox.addItems(list(self.clinvar_cases.keys()))

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def set_widgets_enabled(self, enabled: bool):
        """Helper function to enable/disable UI elements during update"""
        self.gName.setEnabled(enabled)
        self.gMega.setEnabled(enabled)
        self.freqCombo.setEnabled(enabled)
        self.unknownBox.setEnabled(enabled)
        self.addition_area.setEnabled(enabled)
        self.startButton.setEnabled(enabled)

    def updateCheck(self):
        self.set_widgets_enabled(False)  # Disable widgets during update

        download_dir = os.path.abspath(os.getcwd())
        self.worker = Worker(self.df, download_dir)
        self.worker.progress.connect(self.progressBar.setValue)
        self.worker.finished.connect(self.update_complete)
        self.worker.start()

    def update_complete(self):
        # self.statusBar().showMessage('Update Complete')
        self.set_widgets_enabled(True)  # Re-enable widgets after update

    # 겹침 계산 함수
    def calculate_overlap(self, range1_start, range1_end, range2_start, range2_end):
        start = max(range1_start, range2_start)
        end = min(range1_end, range2_end)
        return max(0, end - start)

    def CNV_Save(self):
        try:
            text = self.output_area.toPlainText()
            sname = QFileDialog.getSaveFileName(None,
                                                'Save CNV Report Location',
                                                f'{'_'.join(self.disease_lst)}_{'_'.join(self.fileName.text().split('/')[-1].split('_')[:2])}',
                                                'Text File (*.txt)')

            if text.strip():  # 텍스트가 비어 있지 않은 경우에만 저장
                with open(sname[0], 'w', encoding='utf-8') as file:
                        file.write(text)
                        file.close()

            return file_start(sname[0])
        except:
            pass

    def CNV_Start(self):
        try:
            self.gname = self.disease.lower().split('_')[0]
            self.gfreq = self.freqCombo.currentText()
            if self.gfreq in self.frequency_order.keys():
                freq_que = self.frequency_order[self.gfreq]

                freq_condition = []

                for k, v in self.frequency_order.items():
                    if v >= freq_que:
                        freq_condition.append(k)
                    else:
                        continue
            else:
                freq_condition = ['Previous']
            filtered_df = self.df[(self.df['Name_lower'] == self.gname)]

            if filtered_df.empty:
                filtered_df = self.df[
                    (self.df['Name_lower'].str.startswith(self.gname)) &
                    (self.df['Name_lower'].str.contains(self.typ.lower()))
                    ]
                if filtered_df.empty:
                    self.gname = re.search(r'\(([^()]+)\)', self.gname).group(1)
                    filtered_df = self.df[(self.df['Name_lower'] == self.gname)]
                    if filtered_df.empty:
                        filtered_df = self.df[(self.df['Name_lower'] == self.disease.lower())]
                        if filtered_df.empty:
                            return QMessageBox.warning(self, 'Warning', '지정된 유전병명이 없거나 올바르지 않습니다.')
                        else:
                            original_gname = filtered_df['Name'].iloc[0]
                    else:
                        original_gname = filtered_df['Name'].iloc[0]
                else:
                    original_gname = filtered_df['Name'].iloc[0]
            else:
                original_gname = filtered_df['Name'].iloc[0]
            # print(filtered_df)
            # Previous일 때 조건 검색
            if self.gfreq == 'Previous':
                temp_df = filtered_df[filtered_df['HPOFrequency'] == 'Previous']
                # print(temp_df)
                if temp_df.empty:
                    QMessageBox.warning(self, 'Warning', '유발증상에 대한 이전 데이터가 없습니다.\nFrequency를 Frequent (79-30%) 이상으로 설정하여 유발증상을 기록합니다.')
                    freq_condition = ['Very frequent (99-80%)','Obligate (100%)']

            filtered_df = filtered_df[(filtered_df['HPOFrequency'].isin(freq_condition))]

            if filtered_df.empty:
                return QMessageBox.warning(self, 'Warning', '지정 범위 내에 유발증상이 없습니다.')

            else:
                self.orpha_code = filtered_df['OrphaCode'].iloc[0] if not filtered_df.empty else None
                if freq_condition != ['Previous']:
                    self.input_disease = ', '.join([f"{row['HPOKor']}({row['HPOTerm']})" for index, row in filtered_df.iterrows()])
                else:
                    self.input_disease = filtered_df['HPOKor'].iloc[0].replace('/',', ')

                sheetText = f"상기 검사에 따른 태아의 결과는 {original_gname}({round(self.overlap_length / 1000000, 1)}Mb) 에 대하여 고위험군입니다.\n\n"
                sheetText += f"{original_gname}으로서의 공통증상으로 {self.input_disease} 등이 나타날 수 있습니다.\n\n"

                # 범위가 20~70% 범위일 때
                if self.justify_overlap >= 30:
                    condition = f'{self.typ} 범위에서의 알려진 증상으로 *** ClinVar 내용 입력 필요 *** 등이 나타날 수 있습니다.\n\n'
                if self.justify_overlap >= 70:
                    condition += f'{self.typ} 범위가 {original_gname}의 알려진 범위의 대부분을 차지하기 때문에 위에 나열한 증상이 나타날 가능성이 높습니다.\n\n'
                elif self.justify_overlap == 100:
                    condition += f'{self.typ} 범위가 {original_gname}의 알려진 범위를 모두 포함하기 때문에 위에 나열한 증상이 나타날 가능성이 높습니다.\n\n'
                elif self.justify_overlap < 30:
                    condition = f'그렇지만, 해당 태아의 경우 {self.disease} 범위 중 {self.cytoBand}에만 해당되며 총 범위의 {self.justify_overlap}% 정도를 차지하고 이 범위 수준에서의 증상은 알려지지 않았습니다.\n\n'

                sheetText += f'{condition}'
                sheetText += f"*해당 syndrome의 증상은 사람마다 발현이나 정도가 다르며, 해당되는 증상이 나타나지 않거나 위에서 언급하지 않은 증상이 나타날 수도 있습니다.\n\n상세 내용은 아래에서 찾아보실 수 있습니다.\nOrphanet (http://www.orpha.net) ORPHA code : {self.orpha_code}\n\n"
                sheetText += f"-----------------------------------------------------------------------------------------------\n참고용\n\nchr{self.chromosome}:{self.Start}-{self.End}\n\n"
                if self.counter < len(self.input_data_df):
                    sheetText += f"-----------------------------------------------------------------------------------------------\n\n"
                self.result_text += sheetText

                if self.counter == self.total:
                    return self.output_area.setText(self.result_text)
                # return self.CNV_Save(sheetText)

        except Exception as e:
            return QMessageBox.warning(self, 'Warning', f'알려지지 않은 오류:\n{e}')

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainui = CNV_TestSheet()
    mainui.show()
    sys.exit(app.exec_())
