import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import random
import time


def login_select_date_and_input_names_from_excel():
    # ChromeDriver 경로 설정
    chrome_service = Service('C:/WebDriver/chromedriver.exe')
    driver = webdriver.Chrome(service=chrome_service)
    
    driver.maximize_window()  # 브라우저를 최대화

    # 웹사이트 열기
    driver.get("https://koup.kccworld.net/")
    time.sleep(1.5)  # 사이트 로딩 대기

    # 로그인 절차 (아이디, 비밀번호 입력)
    username_field = driver.find_element(By.NAME, 'username')
    username_field.send_keys('22505625')

    password_field = driver.find_element(By.NAME, 'password')
    password_field.send_keys('a123456789!')

    # 로그인 버튼 클릭
    login_button = driver.find_element(By.CLASS_NAME, 'btn_popup_project_list')
    login_button.click()
    time.sleep(2)  # 로그인 처리 대기

    print("로그인 완료")

    # 로그인 후 페이지 이동
    driver.get("https://koup.kccworld.net/construction/laborattendant/main")
    time.sleep(2.5)

    # 엑셀 파일에서 데이터 읽기 (출면일보 및 고정인원 시트)
    excel_path = r'C:\Users\Administrator\Documents\03_노임 및 중기 ERP\KOUP 출면일보 등록\KOUP 입력.xlsx'  # 엑셀 경로 수정
    df_main = pd.read_excel(excel_path, sheet_name='출면일보', engine='openpyxl')
    df_fixed = pd.read_excel(excel_path, sheet_name='고정인원', engine='openpyxl')
    df_normal_labor = pd.read_excel(excel_path, sheet_name='보통인부 작업내용', engine='openpyxl')
    df_task_fallback = pd.read_excel(excel_path, sheet_name='작업내용 부족시', engine='openpyxl')
    df_inner_labor = pd.read_excel(excel_path, sheet_name='내장공 작업내용', engine='openpyxl')
    df_inner_fallback = pd.read_excel(excel_path, sheet_name='내장공 작업내용 부족시', engine='openpyxl')

    # 엑셀의 첫 번째 행은 날짜 헤더 (B열부터)
    date_columns = df_main.columns[1:]
    names_column = df_main['이름']

    # 최신 날짜부터 시작
    for date in date_columns:
        print(f"현재 처리 중인 날짜: {date}")

        # 근무한 근로자 목록과 공수 추출
        workers = df_main[df_main[date].notna()][['이름', date]].values
        print(f"근무한 근로자 목록: {workers}")

        # 날짜 선택기 조작
        date_parsed = pd.to_datetime(date)
        year = date_parsed.year
        month = date_parsed.month - 1  # 월은 0부터 시작
        day = date_parsed.day

        script = f"""
        var datepicker = $('#dailySafty_schDt').data('kendoDatePicker');
        datepicker.value(new Date({year}, {month}, {day}));
        datepicker.trigger('change');
        """
        driver.execute_script(script)
        print(f"날짜가 {date}로 설정되었습니다.")
        time.sleep(0.4)

        # 이름 필드 대기 및 이름 입력
        try:
            name_field = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, 'labor_tmpSchName'))
            )

            # 근로자 이름 입력
            for worker, work_amount in workers:
                for attempt in range(4):  # 최대 4번 재시도
                    name_field.clear()
                    time.sleep(0.03)  # 필드 초기화 후 대기

                    print(f"'{worker}' 이름 입력 중...")
                    for char in worker:
                        # 완성형 한글 한 글자씩 입력
                        name_field.send_keys(char)
                        time.sleep(0.1)
                        print(f"입력된 글자: {char}")
                    # 입력 확인
                    current_value = name_field.get_attribute('value')
                    if current_value == worker:
                        print(f"'{worker}' 이름이 정상적으로 입력되었습니다.")
                        name_field.send_keys(Keys.ENTER)
                        break
                    else:
                        print(f"'{worker}' 이름 입력 오류. 재시도 중... (시도 횟수: {attempt + 1})")

                else:
                    # 최대 시도 횟수를 초과한 경우
                    print(f"'{worker}' 이름 입력 실패. 해당 작업자 건너뜁니다.")
                    continue

        except TimeoutException:
            print(f"이름 입력 필드를 찾을 수 없습니다. 날짜: {date}")
            continue

        # 전체 근무자 리스트 확인 및 누락된 근무자 추가 등록 (무한 재검증)
        while True:
            try:
                registered_worker_elements = driver.find_elements(By.XPATH, "//tr/td[@class='text_center' and @role='gridcell'][10]")
                registered_workers_in_system = [element.text for element in registered_worker_elements]
                expected_workers = df_main[df_main[date].notna()]['이름'].tolist()

                missing_workers = list(set(expected_workers) - set(registered_workers_in_system))
                if missing_workers:
                    print(f"누락된 근무자: {missing_workers}")
                    for worker in missing_workers:
                        # 입력 검증을 위한 재시도 횟수 설정
                        max_retries = 4
                        for attempt in range(max_retries):
                            name_field.clear()
                            time.sleep(0.06)

                            # 이름 입력
                            for char in worker:
                                name_field.send_keys(char)
                                time.sleep(0.1)
                            print(f"'{worker}' 누락된 이름 입력 시도 중...")

                            # 입력된 이름 검증 (입력한 값과 실제 필드 값 비교)
                            time.sleep(0.01)
                            current_value = name_field.get_attribute('value')

                            if current_value == worker:
                                print(f"'{worker}' 누락된 이름이 정상적으로 입력되었습니다.")
                                # Enter 키 입력
                                time.sleep(1)
                                name_field.send_keys(Keys.ENTER)
                                time.sleep(0.1)
                                break
                            else:
                                print(f"'{worker}' 누락된 이름 입력 오류. 재시도 중... (시도 횟수: {attempt + 1})")

                        else:
                            print(f"'{worker}' 누락된 이름 입력 실패. 해당 작업자 건너뜁니다.")
                            continue
                else:
                    print("모든 근무자가 정상적으로 등록되었습니다.")
                    break  # 누락된 작업자가 없으면 루프 종료

            except Exception as e:
                print(f"근무자 등록 검증 중 오류 발생: {str(e)}")

            # 누락된 근무자 등록 후 대기 시간 추가
            time.sleep(0.5)

        # 이름 입력 후 공수 값이 1이 아닌 사람들 공수 수정
        for worker, work_amount in workers:
            if work_amount == 1:
                print(f"'{worker}'의 공수가 1이므로 공수 입력을 건너뜁니다.")
                continue  # 공수가 1인 경우 공수 수정 건너뛰기

            # 근로자의 이름을 포함한 tr 태그 찾기
            try:
                row = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"//td[text()='{worker}']/parent::tr")
                    )
                )

                # 공수 입력 필드를 클릭하여 활성화
                try:
                    work_cell = row.find_element(By.XPATH, f".//td[@class='text_right' and @style='background-color: LightYellow']")
                    actions = ActionChains(driver)
                    actions.move_to_element(work_cell).click().perform()  # 클릭하여 수정 가능 상태로 만듦
                    time.sleep(0.15)

                    # 텍스트 필드 내 값을 전체 선택 후 삭제 및 입력
                    for attempt in range(3):  # 최대 3번 재시도
                        try:
                            # 활성화된 입력 필드로 포커스 이동
                            work_cell_input = driver.switch_to.active_element  
                            
                            # 텍스트 필드 내 값을 설정
                            work_cell_input.send_keys(Keys.CONTROL, 'a')  # 전체 선택
                            time.sleep(0.05)
                            work_cell_input.send_keys(Keys.BACKSPACE)  # 전체 삭제
                            time.sleep(0.05)
                            work_cell_input.send_keys(str(work_amount))  # 엑셀에서 불러온 공수 값 입력
                            time.sleep(0.1)

                            # 외부 요소 클릭으로 값 확정
                            empty_element = driver.find_element(By.TAG_NAME, "header")
                            actions.move_to_element(empty_element).click().perform()
                            print(f"'{worker}'의 공수가 {work_amount}로 설정되었습니다.")
                            break  # 성공하면 루프 탈출
                        except StaleElementReferenceException:
                            print("요소 참조 오류 발생. 다시 시도합니다...")
                            time.sleep(0.5)  # 잠시 대기 후 재시도

                except (NoSuchElementException, TimeoutException) as e:
                    print(f"공수 입력 중 오류 발생: {str(e)}. 해당 작업자 건너뜁니다.")

            except TimeoutException:
                print(f"'{worker}'에 해당하는 행을 찾을 수 없습니다. 다음 작업자로 넘어갑니다.")
                continue



        # 특정 날짜에 등록된 근로자들 필터링
        specific_day_workers = df_main[df_main[date].notna()]['이름']  # 해당 날짜에 근무한 근로자들 필터링
        fixed_workers = df_fixed[df_fixed['이름'].isin(specific_day_workers)]  # 고정인원 중 해당 날짜에 등록된 사람들 필터링
        fixed_worker_names = fixed_workers['이름'].tolist()  # 고정인원 이름 리스트
        registered_workers = specific_day_workers.tolist()  # 특정 날짜에 등록된 근로자 이름 리스트

        # 고정인원 중 등록된 사람이 없는 경우 보통인부로 바로 넘어가도록 처리
        if not fixed_worker_names:
            print(f"{date}에 고정인원이 없습니다. 보통인부 작업으로 넘어갑니다.")
            continue

        # 고정인원이 있을 경우 고정인원 작업 수행
        for worker, task in fixed_workers.values:
            if worker not in registered_workers:
                print(f"'{worker}' 등록 안 됨. 건너뜁니다.")
                continue  # 등록되지 않은 근로자면 스킵

            try:
                row = WebDriverWait(driver, 1).until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"//td[text()='{worker}']/parent::tr")
                    )
                )
                task_cell = row.find_element(By.XPATH, ".//td[@class='text_left' and @style='background-color: LightYellow']")
                actions = ActionChains(driver)
                actions.move_to_element(task_cell).click().perform()  # 클릭하여 수정 가능 상태로 만듦
                time.sleep(0.2)

                # 작업 내용을 입력
                task_input = driver.switch_to.active_element
                task_input.send_keys(Keys.CONTROL, 'a')  # 기존 내용을 모두 선택
                task_input.send_keys(Keys.BACKSPACE)  # 기존 내용을 지움
                task_input.send_keys(task)  # 새로운 작업 내용을 입력
                time.sleep(0.2)

                # header 태그를 클릭
                empty_element = driver.find_element(By.TAG_NAME, "header")
                actions.move_to_element(empty_element).click().perform()
                print(f"'{worker}'에게 '{task}' 작업이 할당되었습니다.")

            except (TimeoutException, NoSuchElementException) as e:
                print(f"'{worker}' 없음. 다음으로 넘어갑니다. 오류: {str(e)}")
                continue

        # 고정인원 작업이 끝나면 보통인부 작업으로 넘어감
        normal_laborers = get_normal_laborers(driver, fixed_worker_names)  # 고정인원을 제외한 보통인부 리스트 가져오기
        assign_tasks_to_normal_laborers(driver, df_normal_labor, normal_laborers, df_task_fallback, date_parsed)  # 보통인부에 작업내용 할당
    
        # 내장공 작업내용 할당 (내장공, 도장공, 방수공, 형틀목공)
        specialized_workers = get_specialized_laborers(driver, fixed_worker_names)  # 이 위치에서 driver와 fixed_worker_names를 사용
        assign_tasks_to_specialized_laborers(driver, df_inner_labor, specialized_workers, df_inner_fallback, date_parsed)

        # 모든 작업자에 대한 작업 내용 입력이 완료된 후에만 저장 버튼 클릭
        save_after_all_tasks(driver)

    time.sleep(5)
    driver.quit()


# 저장 버튼을 클릭하는 함수 (모든 작업 내용 입력 후 호출)
def save_after_all_tasks(driver):
    """모든 작업이 완료된 후 저장 버튼을 누름."""
    try:
        # 저장 버튼 활성화 확인 및 클릭
        save_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'laborAttendantSummary_saveBtn'))
        )
        save_button.click()  # 클릭 동작 수행
        print("저장 버튼 클릭 성공")

        # 저장 완료 메시지 확인 (예: '저장 완료' 텍스트가 나타나는지 확인)
        success_message = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), '저장 완료')]"))
        )
        print(f"저장이 완료되었습니다: {success_message.text}")

    except TimeoutException:
        print("저장 버튼을 찾거나 저장 상태를 확인하는 중 오류 발생")



# 보통인부 인원 불러오기 함수 (고정인원 제외)
def get_normal_laborers(driver, fixed_workers):
    """해당 날짜에 '보통인부'로 등록된 근로자를 리스트로 반환하되, 고정인원은 제외."""
    normal_workers = []
    
    # '보통인부'가 포함된 행을 찾음
    rows = driver.find_elements(By.XPATH, "//tr[td[contains(text(), '보통인부')]]")
    
    # 각 행에서 이름을 추출 (이름이 있는 td는 10번째 칸)
    for row in rows:
        name = row.find_element(By.XPATH, ".//td[10]").text  # 근로자 이름 추출
        
        # 고정인원에 해당하지 않는 사람만 추가
        if name not in fixed_workers:
            normal_workers.append(name)
    
    print(f"보통인부 리스트 (고정인원 제외): {normal_workers}")
    return normal_workers

def assign_tasks_to_normal_laborers(driver, df_normal_labor, normal_laborers, df_task_fallback, date_parsed):
    """
    보통인부에게 작업내용을 할당하는 함수.
    - 특정 날짜의 작업내용 시트를 필터링하여 작업 분배.
    - 작업내용이 부족하면 '작업내용 부족시' 시트에서 랜덤으로 작업 내용을 가져옴.
    """
    try:
        # 현재 날짜에 해당하는 작업 내용 필터링 (date_parsed 사용)
        tasks_for_date = df_normal_labor[df_normal_labor['월/일'] == date_parsed.strftime('%Y-%m-%d')]

        current_worker_index = 0

        # 날짜별 작업 분배를 진행
        for _, row in tasks_for_date.iterrows():
            task_content = row['작업내용']  # 작업내용 열에서 작업 가져오기
            num_people = int(row['분배인원'])  # 분배인원 열에서 인원 수 가져오기

            for _ in range(num_people):
                if current_worker_index >= len(normal_laborers):
                    print("보통인부가 부족합니다.")
                    break
                worker_name = normal_laborers[current_worker_index]
                assign_task_to_worker(worker_name, task_content, driver)  # 작업 할당
                current_worker_index += 1

        # 남은 보통인부에 작업 내용 부족 시 대체 작업 할당
        remaining_workers = normal_laborers[current_worker_index:]
        if remaining_workers:
            print(f"작업이 부족한 보통인부들: {remaining_workers}")
            fallback_tasks = df_task_fallback['작업내용'].dropna().tolist()  # 대체 작업 내용 목록 가져오기
            if not fallback_tasks:
                print("작업내용 부족시 시트에 작업 내용이 없습니다. 대체 작업 할당 불가.")
                return

            for worker_name in remaining_workers:
                random_task = random.choice(fallback_tasks)  # 랜덤 작업 내용 선택
                assign_task_to_worker(worker_name, random_task, driver)  # 작업 할당
                print(f"'{worker_name}'에게 '{random_task}' 작업이 랜덤으로 할당되었습니다.")

    except KeyError as e:
        print(f"KeyError: {e}. '월/일' 열이 존재하는지 확인하세요.")

# 작업 내용을 근로자에게 할당하는 함수
def assign_task_to_worker(worker_name, task_content, driver):
    """
    작업 내용을 특정 근로자에게 할당하는 함수.
    
    Args:
        worker_name: 작업을 할당할 근로자의 이름
        task_content: 할당할 작업 내용
        driver: Selenium WebDriver 객체
    """
    try:
        # 근로자 이름 행을 찾기
        row = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//td[text()='{worker_name}']/parent::tr"))
        )

        try:
            # 작업 내용을 입력할 셀 찾기 (style 조건 제외)
            task_cell = row.find_element(By.XPATH, ".//td[@class='text_left']")
            actions = ActionChains(driver)
            actions.move_to_element(task_cell).click().perform()
            time.sleep(0.2)

            # 작업 내용을 입력
            task_input = driver.switch_to.active_element
            task_input.send_keys(Keys.CONTROL, 'a')  # 기존 내용 선택
            task_input.send_keys(Keys.BACKSPACE)  # 기존 내용 삭제
            task_input.send_keys(task_content)  # 작업 내용 입력
            time.sleep(0.2)


            # JavaScript를 사용해 특정 요소의 위치를 클릭
            try:
                # 화면 스크롤 조정 (클릭 좌표가 화면에 보이도록)
                driver.execute_script("window.scrollTo(0, 0);")  # 화면을 최상단으로 스크롤
                driver.execute_script("window.scrollBy(414, 72);")  # 클릭 위치로 스크롤 조정

                # 드래그 방지 스타일 추가
                driver.execute_script(
                    """
                    document.body.style.userSelect = 'none';
                    document.body.style.webkitUserSelect = 'none';
                    document.body.style.msUserSelect = 'none';
                    """
                )

                # 지정된 좌표에서 클릭 수행
                driver.execute_script("document.elementFromPoint(414, 72).click();")
                print(f"'{worker_name}'에게 '{task_content}' 작업이 특정 좌표 (414, 72)에서 확정되었습니다.")
            except Exception as e:
                print("디버그 정보:")
                print("현재 화면 크기:", driver.get_window_size())
                print("현재 스크롤 위치:", driver.execute_script("return window.scrollY;"))
                print(f"좌표 클릭 오류: {e}")


        except NoSuchElementException:  # 작업 셀이 없을 경우 처리
            print(f"'{worker_name}'의 작업 셀을 찾을 수 없습니다. 작업을 건너뜁니다.")
    except TimeoutException:
        print(f"'{worker_name}'에 해당하는 행을 찾을 수 없습니다.")



# 내장공 및 특수공종 인원 불러오기 함수
def get_specialized_laborers(driver, fixed_workers):
    """해당 날짜에 '내장공', '도장공', '방수공', '형틀목공' 으로 등록된 근로자를 리스트로 반환하되, 고정인원은 제외."""
    specialized_workers = []
    job_types = ['내장공', '도장공', '방수공', '형틀목공']
    
    # 각각의 직종에 해당하는 근로자를 찾음
    for job in job_types:
        try:
            rows = driver.find_elements(By.XPATH, f"//tr[td[contains(text(), '{job}')]]")
            if rows:
                print(f"'{job}'에 해당하는 작업자가 {len(rows)}명 있습니다.")
            else:
                print(f"'{job}'에 해당하는 작업자가 없습니다.")
            
            # 각 행에서 이름을 추출 (이름이 있는 td는 10번째 칸)
            for row in rows:
                try:
                    name = row.find_element(By.XPATH, ".//td[10]").text  # 근로자 이름 추출
                    # 고정인원에 해당하지 않는 사람만 추가
                    if name not in fixed_workers:
                        specialized_workers.append(name)
                except Exception as e:
                    print(f"{job} 작업자의 이름을 찾는 중 오류 발생: {e}")
                    continue  # 찾지 못한 작업자 건너뛰기
        except Exception as e:
            print(f"'{job}' 직종 작업자를 찾는 중 오류 발생: {e}")
            continue  # 찾지 못한 직종 건너뛰기
    
    print(f"특수 공종 리스트 (고정인원 제외): {specialized_workers}")
    return specialized_workers

# 특수 공종 작업 할당 함수
def assign_tasks_to_specialized_laborers(driver, df_inner_labor, specialized_workers, df_inner_fallback, date):
    """특수 공종(내장공, 도장공, 방수공, 형틀목공)에게 작업내용 할당."""
    current_worker_index = 0

    # 해당 날짜의 작업 내용 필터링
    tasks_for_day = df_inner_labor[df_inner_labor['월/일'] == date.strftime('%Y-%m-%d')]

    for _, row in tasks_for_day.iterrows():
        task_content = row['작업내용']

        for _ in range(1):  # 각 작업은 한 명에게만 할당
            if current_worker_index >= len(specialized_workers):
                print("작업이 부족합니다.")
                break
            worker_name = specialized_workers[current_worker_index]
            assign_task_to_worker(worker_name, task_content, driver)
            current_worker_index += 1

    # 남은 사람들에게 랜덤 작업 할당
    remaining_workers = specialized_workers[current_worker_index:]
    if remaining_workers:
        print(f"작업이 부족한 특수 공종들: {remaining_workers}")
        random_tasks = df_inner_fallback['작업내용'].dropna().tolist()
        for worker_name in remaining_workers:
            random_task = random.choice(random_tasks)
            assign_task_to_worker(worker_name, random_task, driver)


# 자음과 모음을 나눠 입력하는 함수
def type_hangul(field, text):
    for char in text:
        if char in "ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎㅏㅐㅑㅒㅓㅔㅕㅖㅗㅛㅜㅠㅡㅣ":
            # 자음/모음을 그대로 입력
            field.send_keys(char)
        else:
            # 완성형 한글은 직접 입력
            for sub_char in decompose_hangul(char):
                field.send_keys(sub_char)
        time.sleep(0.1)  # 입력 간격 조정
    time.sleep(0.3)  # 모든 입력 완료 후 추가 대기

# 한글 분해 함수 (자음, 모음으로 나누기)
def decompose_hangul(char):
    """
    한글 완성형 글자를 자모(초성, 중성, 종성)로 분리합니다.
    한글이 아니거나 빈 문자열인 경우 빈 문자열 세트를 반환해 안전하게 처리합니다.
    """
    # 유효성 검사: char가 한글 완성형 범위 내에 있는지 확인
    if not isinstance(char, str) or len(char) != 1 or not ('가' <= char <= '힣'):
        return '', '', ''

    UNICODE_START = 0xAC00  # 한글 시작점 ('가')
    CHOSUNG_BASE = 588
    JUNGSUNG_BASE = 28

    code = ord(char) - UNICODE_START
    chosung = code // CHOSUNG_BASE
    jungsung = (code % CHOSUNG_BASE) // JUNGSUNG_BASE
    jongsung = code % JUNGSUNG_BASE

    CHOSUNG_LIST = "ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎ"
    JUNGSUNG_LIST = [
        "ㅏ", "ㅐ", "ㅑ", "ㅒ", "ㅓ", "ㅔ", "ㅕ", "ㅖ",
        "ㅗ", "ㅘ", "ㅙ", "ㅚ", "ㅛ",
        "ㅜ", "ㅝ", "ㅞ", "ㅟ", "ㅠ",
        "ㅡ", "ㅢ", "ㅣ"
    ]
    JONGSUNG_LIST = [""] + list("ㄱㄲㄳㄴㄵㄶㄷㄹㄺㄻㄼㄽㄾㄿㅀㅁㅂㅄㅅㅆㅇㅈㅊㅋㅌㅍㅎ")

    return (
        CHOSUNG_LIST[chosung],
        JUNGSUNG_LIST[jungsung],
        JONGSUNG_LIST[jongsung],
    )



# 실행
login_select_date_and_input_names_from_excel()


