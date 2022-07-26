
'''
# ---------------------------------------------------- #

* Release History
    Apr 6, 2021 (1.0.0 beta)
        - The first beta version is released
        
    Apr 8, 2021 (1.0.1 beta)
        - 'self_pay' analysis_function added
        
    Apr 12, 2021 (1.0.2 beta)
        - 'use in N-minutes' analysis_function added
        - 'OTA update' system_function added
        - bug fixed
        - changed column names
        
    Apr 15, 2021 (1.1 beta)
        - 'user certification' system_function added
        - bug fixed
        
    Apr 21, 2021 (1.1.1 beta)
        - 'user certification' system_function deleted
        
    Sep 24, 2021 (1.2 beta)
        - '4-digits similarity' analysis_function added (1/2)
    
    Oct 5, 2021 (1.2.1 beta) 
        - '4-digits similarity' analysis_function added (2/2)
        
    Oct 5, 2021 (1.2.2 beta)
        - bug hotfix
        
    Oct 12, 2021 (1.2.3 beta)
        - seller information was appended to '4-digits similarity'
        
    Oct 13, 2021 (1.2.4 beta)
        - code refactoring

    Oct 26, 2021 (1.3)
        - 'version check' system_function added, it checks version when starting program automatically 
        - bug fixed
        
    July 26, 2022 (1.3.1)
        - 'memory optimization' system_function added
        
    July 26, 2022 (1.3.2)
        - bug fixed
    
# ---------------------------------------------------- #
'''

__version__ = '1.3.2'

# -------------------- .modules. -------------------- #

import datetime, time, sys, os, urllib.request, bs4, hashlib, shutil, zipfile
from urllib.request import urlretrieve

import pandas as pd
from pandas.api.types import is_string_dtype

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, simpledialog
from tkinter.filedialog import askopenfilenames

# ------------------- .variables. ------------------- #

title = '지역상품권 부정사용 분석'     # software title
username = ''                   # username if state is logged in
workdir = ''                    # location where this software is
resultdir = ''                  # location where output file saved
filename = ''                   # file name for output file
filename_dateflag = ''          # date appended file name for output file
version = ''                    # version
mu = ''                         # open data as pandas
flag_login = 0                  # flag for login status
flag_prepro = 0                 # flag for preprocessing

# -------------------- .sysfunc. -------------------- #

def make_menu():
    try:
        menubar = tk.Menu(mainwindow)

        menu_1 = tk.Menu(menubar, tearoff = 0)
        menu_1.add_command(label = '분석파일 열기', command = load_file)
        menu_1.add_command(label = '분석내용 저장', command = save_file)
        menu_1.add_separator()
        menu_1.add_command(label = '종료', command = main_quit)
        menubar.add_cascade(label = '파일', menu = menu_1)

        menu_2 = tk.Menu(menubar, tearoff = 0)
        menu_2.add_command(label = '로그인', command = lambda:login(pleasebequiet = 0))
        menu_2.add_separator()
        menu_2.add_command(label = '로그아웃', command = logout)
        menubar.add_cascade(label = '계정', menu = menu_2)

        menu_3 = tk.Menu(menubar, tearoff = 0) 
        menu_3.add_command(label = '정보', command = about)
        menu_3.add_command(label = '업데이트', command = lambda:check_version(1))
        menu_3.add_separator()
        menu_3.add_command(label = '결과폴더 열기', command = show_result_dir)
        menubar.add_cascade(label = '정보', menu = menu_3)

        mainwindow.config(menu = menubar)
        
    except Exception as e:
        messagebox.showerror('오류', str(e) + '\n로그 시스템에 오류가 발생했습니다.')

def load_file():
    try:
        global filename, workdir, resultdir, filename_dateflag, flag_prepro
        flag_prepro = 0
        file = askopenfilenames(initialdir = workdir, filetypes = (('csv File', '*.csv'), ('All Files', '*.*')), title = '파일 선택')
        if len(file) == 0: return
        filename = file[0].split('/')[-1]
        workdir = file[0].split('/')[:-1] # 작업경로(파일 경로) 설정
        workdir = '/'.join(workdir) + '/'
        now = datetime.datetime.now()
        filename_dateflag = '[' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '] '
        
        mainwindow.title(title + '(' + username + ') ' + '- 파일 불러오는 중 (' + filename + ')')
        
        load_pandas()
        
        make_preview()
        
        mainwindow.title(title + '(' + username + ') ' + '- 작업중 (' + filename + ')')

        # analysis frame
        frame_analysis = ttk.Notebook(mainwindow, width = 1268, height = 230)
        frame_analysis.place(x = 5, y = 510)
       
        frame_prepro = tk.Frame(mainwindow)
        frame_code1 = tk.Frame(mainwindow)
        frame_code2 = tk.Frame(mainwindow)
        frame_code3 = tk.Frame(mainwindow)
        frame_code4 = tk.Frame(mainwindow)
        frame_code5 = tk.Frame(mainwindow)
        frame_code6 = tk.Frame(mainwindow)
        frame_code7 = tk.Frame(mainwindow)
        frame_code8 = tk.Frame(mainwindow)
        frame_code9 = tk.Frame(mainwindow)
        
        frame_analysis.add(frame_prepro, text = '0) 전처리')
        frame_analysis.add(frame_code1, text = '1) 건당 구매액 + 사용횟수')
        frame_analysis.add(frame_code2, text = '2) 사용횟수 + 총결제금액')
        frame_analysis.add(frame_code3, text = '3) 결제금액 N배수')
        frame_analysis.add(frame_code4, text = '4) 업종평균매출 N배')
        frame_analysis.add(frame_code5, text = '5) N분 내 결제')
        frame_analysis.add(frame_code6, text = '6) 자전거래')
        frame_analysis.add(frame_code7, text = '7) 4자리일치(점주-결제자)')
        frame_analysis.add(frame_code8, text = '8) 4자리일치(결제자 간)')
        frame_analysis.add(frame_code9, text = '9) 종합분석')
        
        bt_undo = tk.Button(mainwindow, text = '데이터 초기화', overrelief = 'solid', command = lambda:undo())
        bt_undo.place(x = 1188, y = 505)
        
         # 전처리 (Code 0)
        prepro_label_1 = ttk.Label(frame_prepro, text = '# 결제취소건 삭제, 개인식별ID 생성 등 분석을 위한 필수과정')
        prepro_label_1.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        var_optimization = tk.IntVar()
        checkbox_optimization = tk.Checkbutton(frame_prepro, text = '메모리 최적화', variable = var_optimization)
        checkbox_optimization.place(x = 20, y = 35)
        
        bt_prepro = tk.Button(frame_prepro, text = '데이터 전처리', overrelief = 'solid', command = lambda:prepro(var_optimization.get()))
        bt_prepro.place(x = 1176, y = 200)
        
         # [price]만원 이상 & [buyfreq]회 이상 (Code 1)
        a1_price = tk.StringVar()
        a1_freq = tk.StringVar()
        
        a1_label_1 = ttk.Label(frame_code1, text = '# 한 사용자의 건당 구매액이')
        a1_label_1.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        entry_a1_price = ttk.Entry(frame_code1, width = 10, textvariable = a1_price)
        entry_a1_price.grid(column = 1, row = 0)
        
        a1_label_2 = ttk.Label(frame_code1, text = ' 원 이상이고, 구매 횟수가 ')
        a1_label_2.grid(column = 2, row = 0)
        
        entry_a1_freq = ttk.Entry(frame_code1, width = 10, textvariable = a1_freq)
        entry_a1_freq.grid(column = 3, row = 0)
        
        a1_label_3 = ttk.Label(frame_code1, text = ' 회 이상인 목록')
        a1_label_3.grid(column = 4, row = 0)
        
        bt_code1 = tk.Button(frame_code1, text = '분석하기', overrelief = 'solid', command = lambda:price_and_freq(a1_price.get(), a1_freq.get()))
        bt_code1.place(x = 1204, y = 200)
        
         # 한 가맹점에서 한 사람이 결제한 총 금액이 [price] 이상 & [buyfreq]회 (Code 2)
        a2_price = tk.StringVar()
        a2_freq = tk.StringVar()
        
        a2_label_1 = ttk.Label(frame_code2, text = '# 한 사용자가 특정 가맹점에서 결제한 횟수가 ')
        a2_label_1.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        entry_a2_freq = ttk.Entry(frame_code2, width = 10, textvariable = a2_freq)
        entry_a2_freq.grid(column = 1, row = 0)
        
        a2_label_2 = ttk.Label(frame_code2, text = ' 회 이상이고 총 결제금액이 ')
        a2_label_2.grid(column = 2, row = 0)
        
        entry_a2_price = ttk.Entry(frame_code2, width = 10, textvariable = a2_price)
        entry_a2_price.grid(column = 3, row = 0)
        
        a2_label_3 = ttk.Label(frame_code2, text = ' 원 이상인 목록')
        a2_label_3.grid(column = 4, row = 0)
        
        bt_code2 = tk.Button(frame_code2, text = '분석하기', overrelief = 'solid', command = lambda:use_price_and_month_and_seller(a2_freq.get(), a2_price.get()))
        bt_code2.place(x = 1204, y = 200)
        
         # 한 가맹점에서 한 사람이 결제한 총 금액이 [price] 이상 & [unit] 배수 (Code 3)
        a3_price = tk.StringVar()
        a3_unit = tk.StringVar()
        
        a3_label_1 = ttk.Label(frame_code3, text = '# 한 사용자가 특정 가맹점에서 결제한 총 금액이 ')
        a3_label_1.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        entry_a3_price = ttk.Entry(frame_code3, width = 10, textvariable = a3_price)
        entry_a3_price.grid(column = 1, row = 0)
        
        a3_label_2 = ttk.Label(frame_code3, text = ' 원 이상이고 ')
        a3_label_2.grid(column = 2, row = 0)
        
        entry_a3_unit = ttk.Entry(frame_code3, width = 10, textvariable = a3_unit)
        entry_a3_unit.grid(column = 3, row = 0)
        
        a3_label_3 = ttk.Label(frame_code3, text = ' 원 단위인 목록')
        a3_label_3.grid(column = 4, row = 0)
        
        bt_code3 = tk.Button(frame_code3, text = '분석하기', overrelief = 'solid', command = lambda:use_price_by_unit(a3_price.get(), a3_unit.get()))
        bt_code3.place(x = 1204, y = 200)
        
         # 업종별 평균매출의 [unit]배 이상인 가맹점 (Code 4)
        a4_unit = tk.StringVar()
        
        a4_label_1 = ttk.Label(frame_code4, text = '# 업종별 평균매출의 ')
        a4_label_1.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        entry_a4_unit = ttk.Entry(frame_code4, width = 10, textvariable = a4_unit)
        entry_a4_unit.grid(column = 1, row = 0)
        
        a4_label_2 = ttk.Label(frame_code4, text = ' 배 이상인 가맹점 목록')
        a4_label_2.grid(column = 2, row = 0)
        
        bt_code4 = tk.Button(frame_code4, text = '분석하기', overrelief = 'solid', command = lambda:average_seller_type_over_unit(a4_unit.get()))
        bt_code4.place(x = 1204, y = 200)
        
         # 한 사람이 [min]분 내 [count]건 이상 결제 (Code 5)
        a5_time = tk.StringVar()
        a5_count = tk.StringVar()
        
        a5_label_1 = ttk.Label(frame_code5, text = '# 한 사용자가 ')
        a5_label_1.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        entry_a5_time = ttk.Entry(frame_code5, width = 10, textvariable = a5_time)
        entry_a5_time.grid(column = 1, row = 0)
        
        a5_label_2 = ttk.Label(frame_code5, text = ' 분 내에 결제한 횟수가 ')
        a5_label_2.grid(column = 2, row = 0)
        
        entry_a5_count = ttk.Entry(frame_code5, width = 10, textvariable = a5_count)
        entry_a5_count.grid(column = 3, row = 0)
        
        a5_label_3 = ttk.Label(frame_code5, text = ' 회 이상인 목록')
        a5_label_3.grid(column = 4, row = 0)
        
        bt_code5 = tk.Button(frame_code5, text = '분석하기', overrelief = 'solid', command = lambda:use_in_Nmin(a5_time.get(), a5_count.get()))
        bt_code5.place(x = 1204, y = 200)
        
         # 자전거래 (Code 6)
        prepro_label_6 = ttk.Label(frame_code6, text = '# 자전거래(가맹점주가 본인의 가맹점에서 결제) 목록')
        prepro_label_6.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        bt_code6 = tk.Button(frame_code6, text = '분석하기', overrelief = 'solid', command = lambda:self_pay())
        bt_code6.place(x = 1204, y = 200)
        
         # 점주-결제자 간 휴대전화 번호 끝 4자리 일치 (Code 7)
        prepro_label_7 = ttk.Label(frame_code7, text = '# 가맹점주-결제자 간 휴대전화번호 끝 4자리 일치 목록')
        prepro_label_7.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        bt_code7 = tk.Button(frame_code7, text = '분석하기', overrelief = 'solid', command = lambda:phone_seller_buyer())
        bt_code7.place(x = 1204, y = 200)
        
         # 결제자 간 휴대전화 번호 끝 4자리 일치 (Code 8)
        prepro_label_8 = ttk.Label(frame_code8, text = '# 가맹점별 결제자 간 휴대전화번호 끝 4자리 일치 목록')
        prepro_label_8.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        bt_code8 = tk.Button(frame_code8, text = '분석하기', overrelief = 'solid', command = lambda:phone_buyer_buyer())
        # bt_code8 = tk.Button(frame_code8, text = '분석하기', overrelief = 'solid', command = lambda:phone_buyer_buyer(), state = 'disable')
        bt_code8.place(x = 1204, y = 200)
        
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n알 수 없는 오류가 발생했습니다.')

def load_pandas():
    global mu
    mu = pd.read_csv(workdir + filename, encoding = 'cp949')
    #print(mu.info(memory_usage = 'deep'))
    
def make_preview(limit = 500):
    try:
        global mu
    
         # preview (https://youtu.be/PgLjwl6Br0k)
        data_preview_frame = tk.LabelFrame(mainwindow, text = '데이터 미리보기 ({}행, {}열)'.format(mu.shape[0], mu.shape[1]))
        data_preview_frame.place(height = 500, width = 1280, y = 5)
        
        data_preview = ttk.Treeview(data_preview_frame)
        data_preview.place(relheight = 1, relwidth = 1)
        
        scrolly = tk.Scrollbar(data_preview_frame, orient = 'vertical', command = data_preview.yview)
        scrollx = tk.Scrollbar(data_preview_frame, orient = 'horizontal', command = data_preview.xview)
        data_preview.configure(xscrollcommand = scrollx.set, yscrollcommand = scrolly.set)
        scrollx.pack(side = 'bottom', fill = 'x')
        scrolly.pack(side = 'right', fill = 'y')
        
        data_preview['column'] = list(mu.columns)
        data_preview['show'] = 'headings'
        for column in data_preview['columns']:
            data_preview.heading(column, text = column)
            
        cnt = 0
        mu['rows'] = mu.to_numpy().tolist()
        for row in mu['rows']:
            data_preview.insert('', 'end', values = row)
            cnt += 1
            if cnt >= int(limit):
                break
        mu = mu.drop('rows', axis = 'columns')
        
    except Exception as e:
        messagebox.showerror('오류', str(e) + '\n로그 시스템에 오류가 발생했습니다.')
        
def show_result_dir():
    os.startfile(os.getcwd() + '/result')
    
def save_file():
    global mu
    mu.to_csv('result/' + filename_dateflag + str(mu.shape[0]) + 'rows x ' + str(mu.shape[1]) + 'cols.csv', encoding = 'cp949', index = False)
    messagebox.showinfo('결과', 'result 폴더에 저장이 완료되었습니다.')
    
def undo():
    if check_prepro():
        global mu, origin_mu
        mu = origin_mu
        
        make_preview()
    else:
        return
    
def check_prepro():
    if flag_prepro == 1:
        return 1
    else:
        messagebox.showwarning('주의', '데이터 전처리 후 실행해주세요.')
        return 0

def about():
    global version
    messagebox.showinfo('정보', f'지역상품권 부정사용 분석 {__version__}\n\n경상남도 디지털정책담당관실\n김승원, 내선 2667')
    
def check_version(flag_silent = 1): # flag_silent = 1 : silent / 0 : popup
    global version
    version = __version__
    server_version = 'https://github.com/SoongFish/GSND_zero/blob/master/version'
    
    try:
        mainwindow.title(title + ' @ 업데이트 서버 연결중...')
        db_version = urllib.request.urlopen(server_version)
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n업데이트 서버에 연결할 수 없습니다.')
        mainwindow.title(title)
        return
        
    db_version_obj = bs4.BeautifulSoup(db_version, 'html.parser')
    sep = db_version_obj.find('td', {'class' : 'blob-code'})
    
    mainwindow.title(title + ' @ 업데이트 확인중...')
    
    if sep.text == version:
        if flag_silent == 1: messagebox.showinfo('버전', '지역상품권 부정사용 분석 ' + version + '\n\n최신버전입니다.')
    else:
        update = messagebox.askquestion('버전', '업데이트가 존재합니다.\n\n현재버전 : ' + version + '\n최신버전 : ' + sep.text + '\n\n업데이트 하시겠습니까?')
        if update == 'yes':
            OTA()
            mainwindow.title(title)
        else:
            mainwindow.title(title)
            return
            
    mainwindow.title(title)

def OTA(): # Over The Air Update function
    try:
        update_flag = 0
        wherenewcode = 'https://github.com/SoongFish/GSND_zero/archive/refs/heads/master.zip'
        
        urlretrieve(wherenewcode, 'master.zip')
        
        with zipfile.ZipFile('master.zip', 'r') as target_zip:
            target_zip.extractall()
        
        os.remove('master.zip')
        
        os.makedirs('bak', mode = 777, exist_ok = True)
        os.makedirs('bak/' + version, mode = 777, exist_ok = True)
        shutil.move(os.getcwd() + '\zero.exe', os.getcwd() + '/bak/' + version)
        
        shutil.move(os.getcwd() + '\GSND_zero-master\zero.exe', os.getcwd())
        shutil.rmtree('GSND_zero-master')
        update_flag = 1
        
        if update_flag == 0: # return -> 1 : success, 0 : fail
            shutil.move(os.getcwd() + 'bak_' + version + '/zero.exe', os.getcwd())
            log(desc = '업데이트 실패(prever:' + version + ')')
            messagebox.showerrer('오류', '업데이트를 실패했습니다.\n프로그램을 다시 실행해주세요.')
            mainwindow.quit()
        else:
            log(desc = '업데이트 완료 (prever : ' + version + ')')
            messagebox.showinfo('알림', '업데이트가 완료되었습니다.\n프로그램을 다시 실행해주세요.')
            mainwindow.quit()
        
    except Exception as e:
        messagebox.showerror('오류', str(e) + '\n업데이트 중 오류가 발생했습니다.')
        
def login(pleasebequiet = 1):
    server_login = 'https://github.com/SoongFish/GSND_zero/blob/master/LDB.acc'
    list_acc = list()
    
    global flag_login, username
    if flag_login == 0 or flag_login == None or flag_login == '':
        flag_login = simpledialog.askstring('인증', '사용자 코드를 입력하세요.', parent = mainwindow)
        
        if flag_login == 0 or flag_login == None or flag_login == '' or len(flag_login) > 20:
            flag_login = 0
            return
        else:
            try:
                mainwindow.title(title + ' @ 로그인 서버 연결중...')
                db_acc = urllib.request.urlopen(server_login)
            except Exception as e:
                log(desc = e)
                messagebox.showerror('계정', str(e) + '\n로그인 서버에 연결할 수 없습니다.')
                mainwindow.title(title)
                flag_login = 0
                return
            
            db_acc_obj = bs4.BeautifulSoup(db_acc, 'html.parser')
            
            sep = db_acc_obj.findAll('td', {'class' : 'blob-code'})
            mainwindow.title(title + ' @ 로그인 중...')
            
            for index in range(len(sep)):
                list_acc.append(sep[index].text)
            
            if sha256(flag_login) in list_acc:
                username = flag_login
                mainwindow.title(title + '(' + username + ') ')
            else:
                messagebox.showwarning('계정', '사용자 정보가 존재하지 않습니다!')
                flag_login = 0
                mainwindow.title(title)
                return
    else:
        if pleasebequiet != 1:
            messagebox.showwarning('계정', '이미 로그인 되어있습니다!') 
            return
    
def logout():
    try:
        global flag_login, flag_prepro, username
        if flag_login == 0 or flag_login == None or flag_login == '':
            messagebox.showwarning('계정', '먼저 로그인하세요!')
        else:
            messagebox.showinfo('계정', '로그아웃이 완료되었습니다.')
            for widget in mainwindow.winfo_children(): # 화면 클리어
                widget.destroy()
            make_menu()
            flag_login = flag_prepro = 0
            username = ''
            mainwindow.title(title)
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n로그아웃 중 오류가 발생했습니다.')
        
def sha256(acc):
    try:
        return hashlib.sha256(acc.encode()).hexdigest()
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n로그인 중 오류가 발생했습니다.')
        
def main_quit():
    mainwindow.quit()
        
def log(desc = ''):
    try:
        os.makedirs('log', mode = 777, exist_ok = True)
        with open('log/' + 'log.log', 'a') as logdata:
            logfuncname = sys._getframe().f_back.f_code.co_name
            logtime = time.strftime('%c', time.localtime(time.time()))
            logdata.write('[{} - {}]\n  # workdir : {}\n  # filename : {}\n  # funcname : {}\n  # desc : {}\n\n'.format(logtime, username if len(username) != 0 else 'anonymous', workdir, filename, logfuncname, desc))
    except Exception as e:
        messagebox.showerror('오류', str(e) + '\n로그 시스템에 오류가 발생했습니다.')

# ------------------- .analysisfunc. ------------------- #

 # 전처리
def prepro(optimization = 0):
    try:
        global mu, origin_mu, flag_prepro

        '''
        global mb, mu
         #buy
        mb['전화번호'] = mb['전화번호'].astype('string')
        mb['ID'] = mb['이용자명'] + mb['전화번호'] # 사용자ID 생성
        #mb = mb.drop('이용자명', axis = 'columns') # 원본유지
        #mb = mb.drop('전화번호', axis = 'columns') # 원본유지
        mb['총금액(A+B)'] = mb['총금액(A+B)'].str.replace(',','').astype('int32')
        mb['고객구매금(A)'] = mb['고객구매금(A)'].str.replace(',','').astype('int32')
        mb['지원금(B)'] = mb['지원금(B)'].str.replace(',','').astype('int32')
        '''

         #use
        mu['폰번호'] = mu['폰번호'].astype('string')
        mu['ID'] = mu['이용자'] + mu['폰번호'] # 사용자ID 생성
        #mu = mu.drop('이용자', axis = 'columns') # 원본유지
        #mu = mu.drop('폰번호', axis = 'columns') # 원본유지
        
        if(is_string_dtype(mu['거래금액'])):
            mu['거래금액'] = mu['거래금액'].str.replace(',','').astype('int32')
        
         # 결제취소 삭제
        cancel = list()
        cancel = mu[mu['거래구분'] == 'QR 결제취소']['원거래번호'].tolist()
        mu = mu[~mu['거래번호'].isin(cancel)]
        mu = mu[mu['거래구분'] == 'QR 결제']
        
         #dtype 메모리 최적화 (var optimization is 1 -> True / 0 -> False)
        if optimization == 1:
            mu = mu.astype({'결제채널':'category', '업종코드':'category', '업종':'category', '거래일':'category', '가맹점ID':'string', '우편번호':'int32', '폰번호':'int32', '거래금액':'int32'})
        
        #print(mu.info(memory_usage = 'deep'))
        
        origin_mu = mu
        
        make_preview()
        
        flag_prepro = 1
        
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n전처리 중 오류가 발생했습니다. (Code 0)')
    
 # [price]만원 이상 & [buyfreq]회 이상 (Code 1)
def price_and_freq(price, buyfreq):
    try:
        if check_prepro():
            global mu
            
             # n만원 이상 사용 필터
            mu = mu[mu['거래금액'] >= int(price)]

             # 구매횟수 count
            mu['구매횟수'] = mu.groupby('ID')['ID'].transform('count')
            mu = mu[mu['구매횟수'] >= int(buyfreq)]
                
            mu = mu.sort_values(by = 'ID', ascending = True)
                
            make_preview()
        else: return
            
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n분석 중 오류가 발생했습니다. (Code 1)')

 # 한 가맹점에서 한 사람이 결제한 총 금액이 [price] 이상 & [buyfreq]회 (Code 2)
def use_price_and_month_and_seller(buyfreq, price):
    try:
        if check_prepro():
            global mu
            
             # 정렬 + 30만원이상 + ID/sID 카운트
            mu = mu.sort_values(['가맹점ID', 'ID'])
            #mu = mu[mu['거래금액'] >= int(price)]
            mu['총결제금액'] = mu.groupby(['ID', '가맹점ID'])['거래금액'].transform('sum')
            mu['구매횟수'] = mu.groupby(['ID', '가맹점ID'])['ID'].transform('count')
            mu = mu[mu['총결제금액'] >= int(price)]
            mu = mu[mu['구매횟수'] >= int(buyfreq)]
            
            make_preview()
        else: return
    
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n분석 중 오류가 발생했습니다.(Code 2)')

 # 한 가맹점에서 한 사람이 결제한 총 금액이 [price] 이상 & [unit] 배수 (Code 3)
def use_price_by_unit(price, unit):
    try:
        if check_prepro():
            global mu
            
             # 정렬 + ID/sID, count/sum
            mu = mu.sort_values(['가맹점ID', 'ID'])
            mu['구매횟수'] = mu.groupby(['ID', '가맹점ID'])['ID'].transform('count')
            mu['총결제금액'] = mu.groupby(['ID', '가맹점ID'])['거래금액'].transform('sum')
            mu = mu[mu['총결제금액'] % int(unit) == 0]
            mu = mu[mu['총결제금액'] >= int(price)]
            
            make_preview()
        else: return
        
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n분석 중 오류가 발생했습니다.(Code 3)')

 # 업종별 평균매출의 [unit]배 이상인 가맹점 (Code 4)
def average_seller_type_over_unit(unit):
    try:
        if check_prepro():
            global mu, temp_mu
            temp_mu = mu
            
            temp_mu['업종별 평균매출'] = round(temp_mu.groupby(['업종코드'])['거래금액'].transform('mean'), 2)
            temp_mu['가맹점 평균매출'] = round(temp_mu.groupby(['가맹점ID'])['거래금액'].transform('mean'), 2)
            
            temp_mu['배수'] = round(temp_mu['가맹점 평균매출'] / temp_mu['업종별 평균매출'], 2)
            temp_mu = temp_mu.drop('업종별 평균매출', axis = 'columns')
            temp_mu = temp_mu.drop('가맹점 평균매출', axis = 'columns')
            temp_mu = temp_mu[temp_mu['배수'] >= float(unit)]
            temp_mu = temp_mu.drop_duplicates('가맹점ID', keep = 'first')
            
            list_seller = list()
            list_seller = temp_mu['가맹점ID'].tolist()
            mu = mu[mu['가맹점ID'].isin(list_seller)]
            
            mu = mu.sort_values(by = '가맹점ID', ascending = True)
            
            make_preview()
        else: return
        
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n분석 중 오류가 발생했습니다. (Code 4)')
        
 # [time_min]분 내 [count]회 결제 (Code 5)
def use_in_Nmin(time_min, count):
    try:
        if check_prepro():
            global mu
            unit_time = str(time_min) + 'Min'
            
            mu = mu.astype({'거래일':'object'})
            mu['datify'] = mu['거래일'] + ' ' + mu['거래시']
            mu['결제일시'] = pd.to_datetime(mu['datify'], format = '%Y-%m-%d %H:%M:%S')
            mu = mu.drop('datify', axis = 'columns')
            mu = mu.sort_values(['ID', '결제일시'])
            
            mu['결제횟수'] = mu.groupby(['ID', mu.결제일시.dt.floor(unit_time)])['결제일시'].transform('count')
            mu = mu[mu['결제횟수'] >= int(count)]
            
            make_preview()
        else: return
            
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n분석 중 오류가 발생했습니다. (Code 5)')

 # 자전거래 (Code 6)
def self_pay():
    try:
        if check_prepro():
            global mu
            
            list_seller = list()
            info_seller = pd.read_csv('sellerinfo.csv', encoding = 'cp949', low_memory = False)
            info_seller = info_seller.fillna(0)
            
            info_seller = info_seller.astype({'대표자명':'string', '대표자휴대전화':'int32', '가맹점관리번호':'string'})
            info_seller['대표자휴대전화'] = info_seller['대표자휴대전화'].astype('string')
            
            list_seller = (info_seller['대표자명'] + info_seller['대표자휴대전화'] + info_seller['가맹점관리번호']).tolist()
            
            mu = mu.astype({'가맹점ID':'string'})
            mu['자전거래ID'] = mu['ID'] + mu['가맹점ID']
            
            mu = mu[mu['자전거래ID'].isin(list_seller)]
            
            make_preview()
        else: return
            
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n분석 중 오류가 발생했습니다. (Code 6)')
    
 # 점주-결제자 간 휴대전화 끝 4자리 일치 (Code 7)    
def phone_seller_buyer():
    try:
        if check_prepro():
            global mu
            
            list_seller = list()
            info_seller = pd.read_csv('sellerinfo.csv', encoding = 'cp949', low_memory = False)
            info_seller = info_seller.fillna(0)
            
            info_seller = info_seller.astype({'대표자명':'string', '대표자휴대전화':'int32', '가맹점관리번호':'string'})
            info_seller['대표자휴대전화'] = info_seller['대표자휴대전화'].astype('string')
            
            list_seller = (info_seller['가맹점관리번호'] + [x[-4:] for x in info_seller['대표자휴대전화']]).tolist()
            
            mu['폰번호'] = mu['폰번호'].astype('string')
            mu['점주-결제자4번호일치'] = mu['가맹점ID'] + [x[-4:] for x in mu['폰번호']]
            
            mu = mu[mu['점주-결제자4번호일치'].isin(list_seller)]
            mu = mu.sort_values(by = '가맹점ID')
            
            mu.rename(columns = {'가맹점ID':'가맹점관리번호'}, inplace = True)
            mu = mu.join(info_seller.set_index('가맹점관리번호')[['대표자명'] + ['대표자휴대전화']], on = '가맹점관리번호')
            
            make_preview()
        else: return
    
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n분석 중 오류가 발생했습니다. (Code 7)')
        
 # 결제자 간 휴대전화 끝 4자리 일치 (Code 8)
def phone_buyer_buyer():
    try:
        if check_prepro():
            global mu

            mu['폰번호'] = mu['폰번호'].astype('string')
            mu['4자리'] = [x[-4:] for x in mu['폰번호']]
            mu['4자리'] = mu['4자리'].astype('string')
            mu = mu.sort_values(['가맹점ID', '폰번호'])
            mu.drop_duplicates(subset = ['가맹점ID', '폰번호'], keep = 'first', inplace = True)
            mu['그룹별4자리일치cnt'] = mu.groupby(['가맹점ID', '4자리'])['4자리'].transform('count')
            mu = mu[mu['그룹별4자리일치cnt'] >= 2]
            mu = mu.sort_values(['가맹점ID', '4자리'])
            
            make_preview()
        else: return
    
    except Exception as e:
        log(desc = e)
        messagebox.showerror('오류', str(e) + '\n분석 중 오류가 발생했습니다. (Code 8)')

# ------------------- .main. ------------------- #

now = datetime.datetime.now()
filename_dateflag = '[' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '] '

os.makedirs('result', mode = 777, exist_ok = True)

mainwindow = tk.Tk()
mainwindow.title(title)
mainwindow.geometry('1280x768')
mainwindow.resizable(False, False)
#mainwindow.tk.call('wm', 'iconphoto', mainwindow._w, tk.PhotoImage(file = 'icon.png'))
make_menu()
check_version(0)
mainwindow.mainloop()