import os
import psutil
import re
import requests
import subprocess
import sys
import time
import xml.etree.ElementTree as ET

import pandas as pd
from PyQt5 import uic, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import *

import disease_info_hardcoded


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

def fetch_xml_from_url(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.content
    else:
        return None


form = resource_path(r'.\CNV.ui')
form_class = uic.loadUiType(form)[0]

class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, df, download_dir):
        super().__init__()
        self.orpha_df = df
        self.download_dir = download_dir

    def run(self):
        orphaXml_Loc = 'https://raw.githubusercontent.com/Orphanet/Orphadata_aggregated/master/Rare%20diseases%20with%20associated%20phenotypes/en_product4.xml'

        xml_data = fetch_xml_from_url(orphaXml_Loc)
        root = ET.fromstring(xml_data)

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

        orpha_df = pd.DataFrame(data)

        csv_path = './CNV_info.csv'
        hpo_translation_df = pd.read_csv(csv_path)

        unique_hpoterms = orpha_df['HPOTerm'].unique()

        hpo_kor_dict = {}
        total_terms = len(unique_hpoterms)
        for i, hpo_term in enumerate(unique_hpoterms):
            matching_row = hpo_translation_df[hpo_translation_df['HPOTerm'] == hpo_term]
            if not matching_row.empty:
                hpo_kor_dict[hpo_term] = matching_row.iloc[0]['HPOKor']
            # 진행 상황 업데이트
            progress_value = int((i + 1) / total_terms * 100)
            self.progress.emit(progress_value)

        orpha_df['HPOKor'] = orpha_df['HPOTerm'].map(hpo_kor_dict)
        if os.path.exists(csv_path):
            force_kill_process_using_file(csv_path)

        orpha_df.to_csv(csv_path)

        # self.orpha_df = pd.read_csv(csv_path)
        # self.orpha_df['Name_lower'] = self.orpha_df['Name'].str.lower()
        #
        # self.hpo_dict = self.orpha_df[['HPOTerm', 'HPOKor']].drop_duplicates()
        # self.hpo_dict = self.hpo_dict.set_index('HPOTerm')['HPOKor'].to_dict()
        # self.term_list = list(self.hpo_dict.keys())

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
        self.synonyms_data_URL = "https://raw.githubusercontent.com/Hoehwang/searching_report/master/synonym_info.csv"
        self.synonyms_data = pd.read_csv(self.synonyms_data_URL)
        self.synonyms_data['Synonym_lower'] = self.synonyms_data['Synonym'].str.lower()

        git_csv_downloader(self.CNVinfo_file, self.CNVinfo_URL)
        git_csv_downloader(self.cvSummary_file, self.cvSummary_URL)

        self.orpha_df = pd.read_csv(f'./{self.CNVinfo_file}')
        self.orpha_df = pd.concat([self.orpha_df, self.previous_data], ignore_index=True)
        self.orpha_df['Name_lower'] = self.orpha_df['Name'].str.lower()

        self.hpo_dict = self.orpha_df[['HPOTerm', 'HPOKor']].drop_duplicates()
        self.hpo_dict = self.hpo_dict.set_index('HPOTerm')['HPOKor'].to_dict()
        self.term_list = list(self.hpo_dict.keys())

        self.disease_all = pd.read_excel(r'disease_info_all.xlsx')

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
        self.recommendBrowser.verticalScrollBar().setValue(0)

    def on_text_changed(self):
        input_text = self.input_field.text().strip().lower()

        if not input_text:
            self.dictBrowser.clear()
            return
        elif len(input_text) > 3:
            # 부분 일치 및 완전 일치 검색 - 제너레이터 사용
            matching_terms = ((term, self.hpo_dict[term]) for term in self.term_list if input_text in str(term).lower())

            self.dictBrowser.clear()
            result_count = 0
            for term, kor in matching_terms:
                self.dictBrowser.append(f"{kor}")
                result_count += 1

            if result_count == 0:
                self.dictBrowser.append("No matching terms found.")
            return self.dictBrowser.verticalScrollBar().setValue(0)
        else:
            self.dictBrowser.clear()

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

        # 테스트중 - for문 주석 제거 및 인덴테이션 복원 필요
        for i in range(len(self.input_data_df)):
            # i = 1 # 테스트 후 나중에 제거 필요
            self.counter += 1
            self.input_data = self.input_data_df.loc[i]
            self.chromosome = self.input_data['Chr']
            self.disease = self.input_data['disease']
            self.disease_lst.append(self.disease)

            self.input_Orpha = self.disease_all[self.disease_all['disease'] == self.disease]['orphacode'].values[0]
            self.Start = self.input_data['Start']
            self.End = self.input_data['End']
            self.range_length = self.End - self.Start
            self.cytoBand = self.input_data['cytoBand']
            self.copyRatio = self.input_data['CopyRatio']

            cytoband_df = disease_info_hardcoded.hg19_cytoband[disease_info_hardcoded.hg19_cytoband['chr'] == str(self.chromosome)]
            cytoDf_start_idx = cytoband_df[(cytoband_df['start'] <= self.Start) & (cytoband_df['end'] >= self.Start)].index[0]
            cytoDf_end_idx = cytoband_df[(cytoband_df['start'] <= self.End) & (cytoband_df['end'] >= self.End)].index[0]
            cytoband_df = cytoband_df.loc[cytoDf_start_idx:cytoDf_end_idx]
            clinvar_query = [f'{self.chromosome}{pos}' for pos in cytoband_df['pos_name'].values]

            if self.copyRatio > 1:
                self.typ = 'Duplication'
            else:
                self.typ = 'Deletion'
            self.summary_df = pd.read_csv(rf'./{self.cvSummary_file}')

            if self.typ == 'Duplication':
                target_df = self.summary_df[self.summary_df['Type'] == 'Duplication']
            else:
                target_df = self.summary_df[self.summary_df['Type'] == 'Deletion']

            target_df = target_df[target_df['File'].apply(lambda x: any(x.startswith(query) for query in clinvar_query))]

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
            file_info_df = target_df.groupby(['Type', 'CytogeneticLocation', 'Start', 'Stop']).filter(lambda x: len(x) > 1)
            file_info_df = file_info_df.sort_values(by='MostRecentSubmission', ascending=False)

            self.clinvar_cases[f'Case {self.counter}'] = []

            for idx, id in enumerate([i.replace('.xml','').split('_')[-1] for i in list(file_info_df['File'])]):
                self.clinvar_cases[f'Case {self.counter}'].append(f'<a href="https://www.ncbi.nlm.nih.gov/clinvar/variation/{id}/">Case_{self.counter}_ClinVar Link{idx+1}: {id}</a>')

            # 기준치 오버랩 검사
            # hardcoded에서 엑셀 데이터 기반으로 변경
            justify_df = self.disease_all[self.disease_all['disease'] == self.disease]

            justify_start = int(justify_df['start'].iloc[0])
            justify_end = int(justify_df['end'].iloc[0])

            justify_length = justify_df['range'].iloc[0]

            # Start~End와 justify_start~justify_end 범위의 겹치는 부분 계산
            overlap_start = max(self.Start, justify_start)
            overlap_end = min(self.End, justify_end)

            # 겹치는 길이 계산
            self.overlap_length = max(0, overlap_end - overlap_start)
            # 겹치는 비율 계산 (퍼센트로)
            self.justify_overlap = (self.overlap_length / justify_length) * 100

            self.CNV_Start(self.counter, self.input_Orpha)

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
        self.fileName.setEnabled(enabled)
        self.fileBrowse.setEnabled(enabled)
        self.freqCombo.setEnabled(enabled)
        self.fileOpenButton.setEnabled(enabled)
        self.output_area.setEnabled(enabled)
        self.caseComboBox.setEnabled(enabled)
        self.recommendBrowser.setEnabled(enabled)
        self.input_field.setEnabled(enabled)
        self.fileOpenButton.setEnabled(enabled)
        self.dictBrowser.setEnabled(enabled)
        self.generateButton.setEnabled(enabled)

    def updateCheck(self):
        self.set_widgets_enabled(False)  # Disable widgets during update

        download_dir = os.path.abspath(os.getcwd())
        self.worker = Worker(self.orpha_df, download_dir)
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
                                                f'{self.disease_lst}_{self.fileName.text().split('/')[-1].split('_')[:2]}',
                                                'Text File (*.txt)')

            if text.strip():  # 텍스트가 비어 있지 않은 경우에만 저장
                with open(sname[0], 'w', encoding='utf-8') as file:
                        file.write(text)
                        file.close()

            return file_start(sname[0])
        except:
            pass

    def CNV_Start(self, case_num, input_Orpha):
        no_info_format = f'''======= Case {case_num} =======\n
상기 검사에 따른 태아의 결과는 {self.disease}({round(self.overlap_length / 1000000, 1)}Mb) 에 대하여 고위험군입니다.\n\n
{self.disease}으로서의 공통증상으로 ** No Information ** 등이 나타날 수 있습니다.\n\n
{self.typ} 범위에서의 알려진 증상으로 *** ClinVar 내용 입력 필요 *** 등이 나타날 수 있습니다.\n\n
*해당 syndrome의 증상은 사람마다 발현이나 정도가 다르며, 해당되는 증상이 나타나지 않거나 위에서 언급하지 않은 증상이 나타날 수도 있습니다.\n\n상세 내용은 아래에서 찾아보실 수 있습니다.\nOrphanet (http://www.orpha.net) ORPHA code : None\n\n
-----------------------------------------------------------------------------------------------\n참고용\n\nchr{self.chromosome}:{self.Start}-{self.End}\n\n'''
        if self.counter < len(self.input_data_df):
            no_info_format += f"-----------------------------------------------------------------------------------------------\n\n"
        try:
            self.freq = self.freqCombo.currentText()
            if self.freq in self.frequency_order.keys():
                freq_que = self.frequency_order[self.freq]

                freq_condition = []

                for k, v in self.frequency_order.items():
                    if v >= freq_que:
                        freq_condition.append(k)
                    else:
                        continue
            else:
                freq_condition = ['Previous']

            filtered_df = self.orpha_df[(self.orpha_df['OrphaCode'] == input_Orpha)]

            if self.freq == 'Previous':
                temp_df = filtered_df[filtered_df['HPOFrequency'] == 'Previous']
                # print(temp_df)
                if temp_df.empty:
                    QMessageBox.warning(self, 'Warning', f'{self.disease}: 유발증상에 대한 이전 데이터가 없습니다.\nFrequency를 전체로 설정하여 유발증상을 기록합니다.')
                    freq_condition = ['Excluded (0%)', 'Very rare (<4-1%)', 'Occasional (29-5%)',
                                      'Frequent (79-30%)', 'Very frequent (99-80%)', 'Obligate (100%)']

            symptom_df = filtered_df[(filtered_df['HPOFrequency'].isin(freq_condition))]

            if symptom_df.empty:
                QMessageBox.warning(self, 'Warning', f'{self.disease}: 지정 범위 내에 유발증상이 없습니다. \nFrequency를 전체로 설정하여 유발증상을 기록합니다.')
                freq_condition = ['Previous', 'Excluded (0%)', 'Very rare (<4-1%)', 'Occasional (29-5%)',
                                  'Frequent (79-30%)', 'Very frequent (99-80%)', 'Obligate (100%)']
                symptom_df = filtered_df[(filtered_df['HPOFrequency'].isin(freq_condition))]
                # print(symptom_df)

                if symptom_df.empty:
                    return QMessageBox.warning(self, 'Warning', f'{self.disease}: 결과 시트를 출력할 수 없습니다.')


            self.orpha_code = symptom_df['OrphaCode'].iloc[0]
            if freq_condition != ['Previous']:
                self.input_disease = ', '.join([f"{row['HPOKor']}({row['HPOTerm']})" for index, row in symptom_df.iterrows()])
            else:
                self.input_disease = symptom_df['HPOKor'].iloc[0].replace('/',', ')

            sheetText = f'======= Case {case_num} =======\n'
            sheetText += f"상기 검사에 따른 태아의 결과는 {self.disease}({round(self.overlap_length / 1000000, 1)}Mb) 에 대하여 고위험군입니다.\n\n"
            sheetText += f"{self.disease}으로서의 공통증상으로 {self.input_disease} 등이 나타날 수 있습니다.\n\n"

            # TODO 단일 Cytoband는 "총 범위의 {round(self.justify_overlap,1)}% 정도를 차지하고" 지울 것

            # 범위가 20~70% 범위일 때
            if self.justify_overlap >= 30:
                condition = f'{self.typ} 범위에서의 알려진 증상으로 *** ClinVar 내용 입력 필요 *** 등이 나타날 수 있습니다.\n\n'
            if self.justify_overlap >= 70:
                condition += f'{self.typ} 범위가 {self.disease}의 알려진 범위의 대부분을 차지하기 때문에 위에 나열한 증상이 나타날 가능성이 높습니다.\n\n'
            elif self.justify_overlap == 100:
                condition += f'{self.typ} 범위가 {self.disease}의 알려진 범위를 모두 포함하기 때문에 위에 나열한 증상이 나타날 가능성이 높습니다.\n\n'
            elif self.justify_overlap < 30:
                condition = f'그렇지만, 해당 태아의 경우 {self.disease} 범위 중 {self.cytoBand}에만 해당되며 총 범위의 {round(self.justify_overlap)}% 정도를 차지하고 이 범위 수준에서의 증상은 알려지지 않았습니다.\n\n'

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
