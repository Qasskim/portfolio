import unittest
from datetime import datetime
from appium import webdriver
from appium.options.android import UiAutomator2Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from appium.webdriver.common.appiumby import AppiumBy
import openpyxl
import os
import time
import psutil

# 설정과 앱 패키지에 대한 Desired Capabilities
capabilities = dict(
    platformName='Android',
    platformVersion='14',
    automationName='uiautomator2',
    deviceName='867400022047199',
    appPackage='com.android.settings',
    appActivity='.Settings',
    language='en'
)

appium_server_url = 'http://localhost:4723'

# 엑셀 파일 설정
excel_file = 'test_results.xlsx'

# 파일이 열려있을 경우 강제로 닫는 함수
def close_excel_file(file_name):
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if 'EXCEL' in proc.info['name'].upper():  # 프로세스 이름에 'EXCEL'이 포함된 프로세스 찾기
                for file in proc.open_files():
                    if file_name in file.path:
                        proc.terminate()  # 엑셀 프로세스 종료
                        proc.wait()  # 프로세스가 종료될 때까지 기다림
                        break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

# 엑셀 파일이 이미 열려 있다면 강제로 닫음
close_excel_file(excel_file)

# 엑셀 파일이 없으면 생성하고, 있으면 기존 파일에 추가 기록
if not os.path.exists(excel_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Test Results'
    sheet.append(['Device Name', capabilities['deviceName']])  # 테스트 디바이스 기재
    sheet.append(['test date/time', datetime.now().strftime('%x %X')])  # 테스트 일시 기재
    sheet.append(['Test Number', 'Test Case', 'Result', 'Note', 'Error Log'])  # 테스트 케이스 항목 및 결과 비고란 생성
    workbook.save(excel_file)
else:
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook['Test Results']
    sheet.append(['Device Name', capabilities['deviceName']])
    sheet.append(['test date/time', datetime.now().strftime('%x %X')])
    sheet.append(['Test Number', 'Test Case', 'Result', 'Note', 'Error Log'])

class TestAppium(unittest.TestCase):
    def setUp(self) -> None:
        self.driver = webdriver.Remote(appium_server_url, options=UiAutomator2Options().load_capabilities(capabilities))

    def test_case(self) -> None:
        case_num: int = 0  # 테스트 케이스 번호

        # Network & internet 페이지 진입
        try:
            case_num += 1
            network_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().text("Network & internet")'))
            )
            network_button.click()

            # Network & internet 페이지 진입 확인
            network_page_name = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (AppiumBy.ACCESSIBILITY_ID, 'Network & internet'))
            )
            network_page_check = network_page_name.get_attribute('content-desc')

            if network_page_check == 'Network & internet':  # Network & internet 페이지 상단 텍스트 확인
                result = 'PASS'  # 텍스트 확인 시 페이지 정상 진입 판정
            else:
                result = 'FAIL'
            sheet.append([case_num, 'Network & internet 페이지에 진입했는가?', result]) # 결과 저장
            workbook.save(excel_file)
        except Exception as e:
            result = 'FAIL'
            screenshot_file = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_network_button_error.png'
            self.driver.save_screenshot(screenshot_file)
            sheet.append([case_num, 'Network & internet 페이지 진입 오류', result, screenshot_file, str(e)]) # 에러 결과 저장
            workbook.save(excel_file)
            print(str(e))

        # Internet 페이지 진입
        try:
            case_num += 1
            internet_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().text("Internet")'))
            )
            internet_button.click()

            # Internet 페이지 진입 확인
            internet_page_name = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((AppiumBy.ACCESSIBILITY_ID, 'Internet'))
            )
            internet_page_check = internet_page_name.get_attribute('content-desc')

            if internet_page_check == 'Internet':  # Internet 페이지 상단 텍스트 확인
                result = 'PASS'
            else:
                result = 'FAIL'
            sheet.append([case_num, 'Internet 페이지에 진입했는가?', result]) # 결과 저장
            workbook.save(excel_file)
        except Exception as e:
            result = 'FAIL'
            screenshot_file = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_internet_button_error.png'
            self.driver.save_screenshot(screenshot_file)
            sheet.append([case_num, 'Internet 페이지 진입 오류', result, screenshot_file, str(e)]) # 에러 결과 저장
            workbook.save(excel_file)
            print(str(e))

        # T-Mobile 페이지 진입
        try:
            case_num += 1
            tmoblie_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().text("T-Mobile")'))
            )
            tmobile_button_text = tmoblie_button.get_attribute('text')

            if tmobile_button_text == 'T-Mobile':  # T-Mobile 버튼 확인 후 설정 버튼 터치
                tmoblie_setting_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((AppiumBy.ACCESSIBILITY_ID, 'Settings'))
                )
                tmoblie_setting_button.click()

                # T-Mobile 페이지 진입 확인
                tmobile_page_name = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((AppiumBy.ACCESSIBILITY_ID, 'T-Mobile'))
                )
                tmobile_page_check = tmobile_page_name.get_attribute('content-desc')

                if tmobile_page_check == 'T-Mobile':  # T-Mobile 페이지 상단 텍스트 확인
                    result = 'PASS'
                else:
                    result = 'FAIL'
                sheet.append([case_num, 'T-Mobile 페이지에 진입했는가?', result]) # 결과 저장
                workbook.save(excel_file)
            else:
                result = 'FAIL'
                sheet.append([case_num, 'T-Mobile 버튼 텍스트 확인 실패', result])
                workbook.save(excel_file)

        except Exception as e:
            result = 'FAIL'
            screenshot_file = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_tmobile_button_error.png'
            self.driver.save_screenshot(screenshot_file)
            sheet.append([case_num, 'T-Mobile 페이지 진입 오류', result, screenshot_file, str(e)]) # 에러 결과 저장
            workbook.save(excel_file)
            print(str(e))

        # 인터넷 사용량 확인
        try:
            case_num += 1
            usage_data = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((AppiumBy.ID, 'com.android.settings:id/data_usage_view'))
            )
            usage_data_check = usage_data.get_attribute('text')

            sheet.append([case_num, '인터넷 사용량', usage_data_check]) # 결과 저장
            workbook.save(excel_file)
        except Exception as e:
            result = 'FAIL'
            screenshot_file = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_usage_data_error.png'
            self.driver.save_screenshot(screenshot_file)
            sheet.append([case_num, '인터넷 사용량 확인 오류', result, screenshot_file, str(e)]) # 에러 결과 저장
            workbook.save(excel_file)
            print(str(e))

        # Roaming 사용 설정
        try:
            case_num += 1
            roaming_use = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((AppiumBy.XPATH, '//*[@text="Roaming"]'))
            )
            roaming_use_check = roaming_use.get_attribute('text')
            toggle_switch = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((AppiumBy.XPATH, '(//*[@resource-id="android:id/switch_widget"])[3]'))
            )
            toggle_status = toggle_switch.get_attribute("checked")

            if toggle_status == 'true':
                roaming_toggle_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable(
                        (AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().resourceId("android:id/switch_widget").instance(2)'))
                )
                roaming_toggle_button.click()  # 토글 버튼 터치

                toggle_status = toggle_switch.get_attribute("checked")  # 갱신된 toggle_status 확인

                if toggle_status == 'false':
                    result = 'PASS'
                else:
                    result = 'FAIL'
                sheet.append([case_num, '로밍 상태가 On에서 Off로 변경됐는가?', result]) # 결과 저장
                workbook.save(excel_file)

            else:
                roaming_toggle_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable(
                        (AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().resourceId("android:id/switch_widget").instance(2)'))
                )
                roaming_toggle_button.click()  # 토글 버튼 터치

                # 로밍 얼럿 팝업 출력 관련
                roaming_on_alert = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((AppiumBy.ID, 'android:id/alertTitle'))
                )
                roaming_on_alert_check = roaming_on_alert.get_attribute('text')

                roaming_on_alert_ok = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((AppiumBy.ID, 'android:id/button1'))
                )
                roaming_on_alert_ok.click()

                toggle_status = toggle_switch.get_attribute("checked")  # 갱신된 toggle_status 확인

                if toggle_status == 'true':
                    result = 'PASS'
                else:
                    result = 'FAIL'
                sheet.append([case_num, '로밍 상태가 Off에서 On으로 변경됐는가?', result]) # 결과 저장
                workbook.save(excel_file)
        except Exception as e:
            result = 'FAIL'
            screenshot_file = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_roaming_toggle_error.png'
            self.driver.save_screenshot(screenshot_file)
            sheet.append([case_num, '로밍 사용 설정 오류', result, screenshot_file, str(e)]) # 에러 결과 저장
            workbook.save(excel_file)
            print(str(e))

    def tearDown(self) -> None:
        if self.driver:
            self.driver.quit()

if __name__ == '__main__':
    unittest.main()
