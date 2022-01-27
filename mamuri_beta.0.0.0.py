

import os
from time import sleep
from tkinter import *
import tkinter
from tkinter.filedialog import askopenfilename
import re
import win32com.client as win32
import pyperclip as cb
from tkinter import ttk  ## 콤보박스 사용 위해 임포트.
import ctypes  # An included library with Python install.


BASE_DIR = 'C:/'

pattern1= r"\d.\D"
pattern2 = re.escape("□") #고정폭빈칸은 놓치고 감...주의. 예외사항 추가필요
pattern3= re.escape("ㅇ") #두번째 서식.
pattern4 = re.escape("*")
patternarray= [pattern1, pattern2, pattern3, pattern4]

def hwp_check_if_blank_exists_above(hwp):
    current_position = hwp.GetPos()  # 현위치 저장(간혹 다음 검색위치로 튀는문제 조치)
    hwp.HAction.Run("MoveLineBegin")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("Copy")
    hwp.SetPos(*current_position)  # 방금위치 복원
    if cb.paste() == "\r\n\r\n":
        return True
    else:
        return False


def hwp_check_if_blank_exists_below_origin(hwp):
    current_position = hwp.GetPos()  # 현위치 저장(간혹 다음 검색위치로 튀는문제 조치)
    hwp.HAction.Run("MoveLineEnd")
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("Copy")
    hwp.SetPos(*current_position)  # 방금위치 복원
    if cb.paste() == "\r\n\r\n":
        return True
    else:
        return False


def hwp_check_if_blank_exists_below(hwp):
    current_position = hwp.GetPos()  # 현위치 저장(간혹 다음 검색위치로 튀는문제 조치)
    hwp.HAction.Run("MoveLineEnd")
    hwp.HAction.Run("MoveSelDown") # 한줄전체 셀영역 만들기.
    hwp.HAction.Run("Copy")
    hwp.SetPos(*current_position)  # 방금위치 복원 
    
    del_black =cb.paste().strip()       ##문자열 공백 제거하기.
    
    
    if del_black == "": 
    ##문자열 공백 제거하기.
     
     ### https://wikidocs.net/33017
     ### https://stackoverflow.com/questions/61625307/paste-text-from-clipboard-without-writing-empty-lines
     ### https://stackoverflow.com/questions/9573244/how-to-check-if-the-string-is-empty
        return True
    else:
        return False






     

class InsertBlankLine:
    # 속성 생성
    def __init__(self, hwp, height, listform):
        self.hwp = hwp
        self.height = height
        self.listform = listform
        
     # 메소드 생성
    def insert_blankline(self):

        dAct = self.hwp.CreateAction("CharShape")
        dSet = dAct.CreateSet()
        dAct.GetDefault(dSet)
        dSet.SetItem("FaceNameUser", "휴먼명조")
        dSet.SetItem("FontTypeUser", 1)
        dSet.SetItem("FaceNameSymbol", "휴먼명조")
        dSet.SetItem("FontTypeSymbol", 1)
        dSet.SetItem("FaceNameOther", "휴먼명조")
        dSet.SetItem("FontTypeOther", 1)
        dSet.SetItem("FaceNameJapanese", "휴먼명조")
        dSet.SetItem("FontTypeJapanese", 1)
        dSet.SetItem("FaceNameHanja", "휴먼명조")
        dSet.SetItem("FontTypeHanja", 1)
        dSet.SetItem("FaceNameLatin", "휴먼명조")
        dSet.SetItem("FontTypeLatin", 1)
        dSet.SetItem("FaceNameHangul", "휴먼명조")
        dSet.SetItem("FontTypeHangul", 1)
        
        dSet.SetItem("Height", self.height)
        dAct.Execute(dSet)
        
        self.hwp.HAction.Run("Cancel")
        return self.hwp   

 

    def hwp_find_and_go(self):
        self.hwp.InitScan()
        listformpattern=re.escape(self.listform) 


        while True :
            text = self.hwp.GetText()
            if text[0] ==101 or text[0]==1:
                print(text[0])
                break
            else:
                if re.match(listformpattern, text[1].strip().replace(" ", "")): ##앞단의 빈칸을 정리해줘야 함.
                    print(text[1])
                    self.hwp.MovePos(201)  # moveScanPos : GetText() 실행 후 위치로 이동한다.
                    getText_point=self.hwp.GetPos()
                    sleep(0.2)
                    self.hwp.MovePos(7) #문단의 끝.
                    self.hwp.MovePos(201) #moveScanPos : GetText()실행 후 위치로 이동한다 =>변경필요.2021.09.14
                    self.hwp.SetPos(*getText_point)



                    if hwp_check_if_blank_exists_above(self.hwp) : # 윗 빈줄 있는지 체크! True or False
                        self.hwp.HAction.Run("MoveLineBegin")
                        dAct = self.hwp.CreateAction("InsertText")
                        # dSet = dAct.CreateSet()
                        # dAct.GetDefault(dSet)
                        # dSet.SetItem("Text", "빈줄있음")
                        # dAct.Execute(dSet)
                    
                    #빈줄이 있는 경우, 위로 올라가 폰트를 조정한다.
                    #hwp.HAction.Run("MoveLineBegin")
                        self.hwp.HAction.Run("MoveUp")
                        self.hwp.HAction.Run("MoveSelRight")
                        InsertBlankLine.insert_blankline(self) #빈줄 삽입
                    else:
                        self.hwp.HAction.Run("BreakPara")
                        self.hwp.HAction.Run("MoveUp")
                        self.hwp.HAction.Run("MoveSelRight")
                        InsertBlankLine.insert_blankline(self) #빈줄 삽입
                    self.hwp.SetPos(*getText_point) 


#기본 빈줄삽입은 윗줄을 기준으로 함/ 아랫줄은 주석
                # if hwp_check_if_blank_exists_below(hwp) : # 아래 빈줄이 있는지 체크!! True or False
                #     hwp.HAction.Run("MoveLineEnd")
                #     hwp.HAction.Run("MoveDown")
                #     hwp.HAction.Run("MoveSelRight")
                #     insert_blankline(hwp, "아래") #빈줄 삽입
                # else:
                #     hwp.HAction.Run("MoveLineEnd")
                #     dAct = hwp.CreateAction("InsertText")
                    
                #     hwp.HAction.Run("BreakPara")
                #     hwp.HAction.Run("MoveLineEnd")
                #     hwp.HAction.Run("MoveSelRight")
                #     insert_blankline(hwp,  "아래") #빈줄 삽입

                    sleep(0.2)
                    self.hwp.MovePos(20) ## 한줄 아래로 이동. 2021.08.26
                    self.hwp.MovePos(23) ## 한줄 아래로 이동. 2021.08.26


                else:
                    pass
        self.hwp.ReleaseScan()
        self.hwp.MovePos(2)

class CreateEmptyLine:
    # 속성 생성
    def __init__(self, hwp):
        self.hwp = hwp
        #self.list_form = list_form
        
     # 메소드 생성
    

# tkinter 사용법
# https://www.youtube.com/watch?v=ITaDE9LLEDY



if __name__ == '__main__': 
    defaultCharSize =1500.0 ##기본글자크기 정하자.
    root=tkinter.Tk()
    root.title("마무리")
    root.geometry("350x100")
    
    


   
    def button_command():
       print("test")
       text = entry1.get()
       print(text)
       return None
   
    
    
    def open_hwp():
     #표에 글자사이즈 입력값 전달 예시(2021.09.25)
        tkinter.Tk().withdraw()
        file_name = askopenfilename(initialdir=BASE_DIR)
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule('FilePathCheckDLL','FileAuto') # 보안 승인창 뜨지 않도록 모듈 등록
        hwp.SetMessageBoxMode(0x00020000) #ㄴ 메세지 창 뜨지 않도록 설정 ##[출처] 파이썬으로 한글(hwp)내에 미주 개수 세기|작성자 코딩헤윰
        hwp.Open(os.path.join(BASE_DIR, file_name)) #파일열기
        current_window=hwp.XHwpWindows.Item(0)
        current_window.Visible=True
        hwp.HAction.Run("MoveDocBegin") #문서 시작으로 커서 이동
        
       
        
            
            
        
        height = int(fontsize.get())*100
        listform= combolistform.get()

        insertBlankLine = InsertBlankLine(hwp, height, listform)
        insertBlankLine.hwp_find_and_go()
        
        
        
        
        #blank_line_height_set(height_pattern1)
        #insertline = CreateEmptyLine(hwp)
        #insertline.insertline(hwp)
        return hwp
    
    
    combolistform = ttk.Combobox(root, 
                            values=[
                                    "□", 
                                    "ㅇ",
                                    "-",
                                    "*", 
                                    "※"], width=4)
 
    combolistform.pack()
    combolistform.place(x=50, y=30)
    combolistform.current(0)
    
    fontsize = Entry(root, width=4)
    fontsize.insert(END, '15')    
    fontsize.pack()
    fontsize.place(x=130, y=30)
    
    title = Label(root, text="목차별 빈줄 삽입하기")
    title.pack()
    title.place(x=50, y=5)
   
    
    
    label_fontsize = Label(root, text="pt")
    label_fontsize.pack()
    label_fontsize.place(x=160, y=30)

    label_explain = Label(root, text="빈줄 기본 15pt, 윗줄에 삽입됩니다.")
    label_explain.pack()
    label_explain.place(x=50, y=60)

    
    
    button_run= Button(root, text ="실행", command=open_hwp)
    button_run.pack()
    button_run.place(x=200, y=30)
    
    
    button_close = Button(text = "닫기", command = root.destroy)
    button_close.pack()
    button_close.place(x=250, y=30)
    
    
       
    root.mainloop()
    # def get_text():
    #     print(entry1.get())
        
    # def get_height():
    #     height = int(entry1.get())*100
    #     print(height)
    #     return height

    
    # entry1 = Entry(root, bg="light green", show="15")
    # entry1.place(x=250, y=150)
    # entry1.pack()
    # button = Button(text = 'Get Text', command = get_text)
    # button.pack()    
 
    # btn1 = Button(root)
    # btn1.config(text="실행")
    # btn1.config(command=open_hwp)
    # btn1.place(x=30, y=100)
    # btn1.pack()


    # btn3 = Button(text = "닫기", command = root.destroy)
    # btn3.pack()

    # root.mainloop()    #맨 마지막에 mainloop()가 나와야 함.





