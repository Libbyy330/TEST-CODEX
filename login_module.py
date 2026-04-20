import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import os
import time
import traceback
from attendance_module import run_attendance_auto

def login_and_get_driver():
    login_file = 'login_info.txt'
    saved_id = ''
    saved_pw = ''
    if os.path.exists(login_file):
        try:
            with open(login_file, 'r', encoding='utf-8') as f:
                lines = f.read().splitlines()
                if len(lines) >= 2:
                    saved_id, saved_pw = lines[0], lines[1]
        except Exception:
            pass

    root = tk.Tk()
    root.withdraw()
    login_window = tk.Toplevel()
    login_window.title("KOUP 로그인")
    login_window.geometry("350x250")
    login_window.resizable(False, False)

    tk.Label(login_window, text="로그인", font=("맑은 고딕", 14, "bold")).pack(pady=(15, 5))
    tk.Label(login_window, text="사번(아이디)을 입력하세요:", font=("맑은 고딕", 10)).pack()
    id_entry = tk.Entry(login_window, width=25)
    id_entry.pack(pady=3)
    id_entry.insert(0, saved_id)
    tk.Label(login_window, text="비밀번호를 입력하세요:", font=("맑은 고딕", 10)).pack()
    pw_entry = tk.Entry(login_window, show="*", width=25)
    pw_entry.pack(pady=3)
    pw_entry.insert(0, saved_pw)

    save_var = tk.BooleanVar(value=True if os.path.exists('login_info.txt') else False)
    save_check = tk.Checkbutton(
        login_window, text="아이디/비밀번호 저장", variable=save_var, font=("맑은 고딕", 10)
    )
    save_check.pack(pady=(0, 5))

    result = {}
    def on_submit(event=None):
        user_id = id_entry.get()
        password = pw_entry.get()
        if not user_id:
            messagebox.showwarning("입력 필요", "사번(아이디)을 입력하세요.", parent=login_window)
            return
        if not password:
            messagebox.showwarning("입력 필요", "비밀번호를 입력하세요.", parent=login_window)
            return
        result['user_id'] = user_id
        result['password'] = password
        result['save'] = save_var.get()
        login_window.destroy()

    submit_btn = tk.Button(
        login_window, text="로그인", command=on_submit,
        width=18, height=2, font=("맑은 고딕", 10, "bold")
    )
    submit_btn.pack(pady=15)

    id_entry.bind('<Return>', on_submit)
    pw_entry.bind('<Return>', on_submit)
    id_entry.focus_set()
    login_window.grab_set()
    root.wait_window(login_window)

    user_id = result.get('user_id')
    password = result.get('password')
    save_login = result.get('save', False)
    if not user_id or not password:
        return None

    chrome_service = Service('C:/WebDriver/chromedriver.exe')
    driver = webdriver.Chrome(service=chrome_service)
    driver.maximize_window()
    driver.get("https://koup.kccworld.net/")
    time.sleep(1.5)

    username_field = driver.find_element(By.NAME, 'username')
    username_field.send_keys(user_id)
    password_field = driver.find_element(By.NAME, 'password')
    password_field.send_keys(password)
    login_button = driver.find_element(By.CLASS_NAME, 'btn_popup_project_list')
    login_button.click()
    time.sleep(2)

    if driver.current_url == "https://koup.kccworld.net/":
        messagebox.showerror("로그인 실패", "로그인에 실패했습니다. 프로그램을 종료합니다.")
        driver.quit()
        return None

    if save_login:
        try:
            with open(login_file, 'w', encoding='utf-8') as f:
                f.write(user_id + '\n' + password)
        except Exception:
            pass
    elif os.path.exists(login_file):
        try:
            os.remove(login_file)
        except Exception:
            pass
    print("로그인 완료")
    return driver

def login_and_run_attendance():
    """로그인 후 출면일보 자동입력까지 전체 실행 제어"""
    driver = None
    try:
        driver = login_and_get_driver()
        if driver is None:
            print("프로그램을 종료합니다.")
            return

        # 출면일보 자동입력 실행
        run_attendance_auto(driver)
        print("\n출면일보 자동입력이 완료되었습니다.")

    except Exception as e:
        print("\n[오류 발생] 자동화 실행 중 예외가 발생했습니다.")
        print(str(e))
        traceback.print_exc()
    finally:
        try:
            input("\n브라우저를 종료하려면 엔터를 누르세요.")
        except (EOFError, RuntimeError):
            pass
        if driver:
            driver.quit()
