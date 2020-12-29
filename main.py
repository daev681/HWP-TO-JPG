import os  # .path.join(), .listdir(), .chdir(), .getcwd() 등 사용
import tkinter

import win32com.client as win32  # 한/글 열 수 있는 모듈

import tkinter

from pdf2image import convert_from_path
import win32gui  # 창 숨기기 위한 모듈
from win32api import Sleep

window = tkinter.Tk()
window.title("DaeHeon_image_change")
window.geometry("320x320+100+100")
window.resizable(False, False)

import subprocess
import shutil


def search_delete_jpg():
    if __name__ == "__main__":
        root_dir = path.get()
        jpg_Delete(path.get())
        for (root, dirs, files) in os.walk(root_dir):
            print("# root : " + root)
            if len(dirs) > 0:
                for dir_name in dirs:
                    print("테스트중" + root + "\\" + dir_name)
                    jpg_Delete(root + "\\" + dir_name)
                    print("dir: " + dir_name)
            if len(files) > 0:
                for file_name in files:
                    print("file: " + file_name)


def search_delete_pdf():
    if __name__ == "__main__":
        pdf_Delete(path.get())
        root_dir = path.get()
        for (root, dirs, files) in os.walk(root_dir):
            print("# root : " + root)
            if len(dirs) > 0:
                for dir_name in dirs:
                    print("테스트중" + root + "\\" + dir_name)
                    pdf_Delete(root + "\\" + dir_name)
                    print("dir: " + dir_name)
            if len(files) > 0:
                for file_name in files:
                    print("file: " + file_name)


def search_delete_hwp():
    if __name__ == "__main__":
        root_dir = path.get()
        hwp_Delete(path.get())
        for (root, dirs, files) in os.walk(root_dir):
            print("# root : " + root)
            if len(dirs) > 0:
                for dir_name in dirs:
                    print("테스트중" + root + "\\" + dir_name)
                    hwp_Delete(root_dir + "\\" + dir_name)
                    print("dir: " + dir_name)

            if len(files) > 0:
                for file_name in files:
                    print("file: " + file_name)


def search_pdf_change_jpg():
    if __name__ == "__main__":
        root_dir = path.get()
        JPG_CHANGE(path.get())
        for (root, dirs, files) in os.walk(root_dir):
            print("# root : " + root)
            if len(dirs) > 0:
                for dir_name in dirs:
                    print("테스트중" + root + "\\" + dir_name)
                    JPG_CHANGE(root + "\\" + dir_name)
                    print("dir: " + dir_name)

            if len(files) > 0:
                for file_name in files:
                    print("file: " + file_name)
    print("파일 복사 완료")

def search_hwp_change_jpg():
    if __name__ == "__main__":
        root_dir = path.get()
        PDF_CHANGE(path.get())
        JPG_CHANGE(path.get())
        pdf_Delete(path.get())
        hwp_Delete(path.get())
        for (root, dirs, files) in os.walk(root_dir):
            print("# root : " + root)
            if len(dirs) > 0:
                for dir_name in dirs:
                    print("테스트중" + root + "\\" + dir_name)
                    PDF_CHANGE(root + "\\" + dir_name)
                    JPG_CHANGE(root + "\\" + dir_name)
                    pdf_Delete(root + "\\" + dir_name)
                    hwp_Delete(root + "\\" + dir_name)
                    print("dir: " + dir_name)

            if len(files) > 0:
                for file_name in files:
                    print("file: " + file_name)


def search_hwp_change_jpg_upgrade():
    if __name__ == "__main__":
        root_dir = path.get()
        HWP_JPG_CHANGE_UPGRADE(path.get())
        #HWP_JPG_CHNAGE_except_add(path.get())
        for (root, dirs, files) in os.walk(root_dir):
            print("# root : " + root)
            if len(dirs) > 0:
                for dir_name in dirs:
                    print("테스트중" + root + "\\" + dir_name)
                    HWP_JPG_CHANGE_UPGRADE(root + "\\" + dir_name)
                    #HWP_JPG_CHNAGE_except_add(root + "\\" + dir_name)
                    print("dir: " + dir_name)

            if len(files) > 0:
                for file_name in files:
                    print("file: " + file_name)


def search_hwp_change_jpg_except_add(full_filename):
    if __name__ == "__main__":
        root_dir = full_filename
        HWP_JPG_CHANGE_UPGRADE(full_filename)
        for (root, dirs, files) in os.walk(root_dir):
            print("# root : " + root)
            if len(dirs) > 0:
                for dir_name in dirs:
                    print("테스트중" + root + "\\" + dir_name)
                    HWP_JPG_CHANGE_UPGRADE(root + "\\" + dir_name)
                    print("dir: " + dir_name)

            if len(files) > 0:
                for file_name in files:
                    print("file: " + file_name)


def main():
    p = subprocess.Popen(["pdftoppm", "-h"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = p.communicate()
    print(out, err)


def JPG_TEST():
    global images
    global img
    images = convert_from_path('C:/Users/마-03/Desktop/01-20150920202610882/01-20150920202610882.pdf', 500)
    for idx, img in enumerate(images):
        img.save('pdf_' + str(idx).zfill(len(str(len(images)))) + '.jpg', 'JPEG')  # pdf_넘버링.jpg 이런 방식으로 네이밍을 합니다.


def HPW_JPG_CHANGE():
    PDF_CHANGE()
    JPG_CHANGE()
    pdf_Delete()
    hwp_Delete()


def JPG_CHANGE(full_filename):
    global pdf
    global images
    global image
    os.chdir(full_filename)
    for i in os.listdir():
        print(full_filename + "\\" + i)
        pdf = os.path.splitext(i)

        if pdf[-1] in ".pdf" or pdf[-1] in ".PDF":
            if pdf[-1] != "":
                images = convert_from_path(full_filename + "\\" + i, 500)
                print(full_filename + "\\" + pdf[0] + ".jpg")
                for idx, image in enumerate(images):
                    image.save(full_filename + "\\" + pdf[0] + "_" + str(idx) + ".jpg", "JPEG")

def HWP_JPG_CHNAGE_except_add(full_filename):
 try:
    global hwp
    global count
    global object_Hwp
    count = 0
    print("파일 복사중 위치 : " + full_filename)
    os.chdir(full_filename)

    # 디렉토리에 있는 모든 파일들을 불러와
    file_list = os.listdir(full_filename)
    # hwp,HWP 확장자로 구성되어있는 내용들만 리스트로 얻어옴
    file_list_hwp = [file for file in file_list if file.endswith(".hwp") or file.endswith(".HWP")]

    print(len(file_list_hwp))

    for i in file_list_hwp:
        print(count)
        hwp_name = os.path.splitext(i)  # 파일 확장자를 분리 시켜줌 ex) 1234.hwp -> .hwp

        # 첫번째 문서가 시작된다면 ,
        if count == 0:
            hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # 한/글 열기
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")  # 보안창 닫기
            # hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
            hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 해당 윈도우의 핸들값 찾기
            win32gui.ShowWindow(hwnd, 0)  # 0은 숨기기, 5는 보이기, 3은 풀스크린 등

        # 이중 방어막 의미 없는 조건문 입니다
        if hwp_name[-1] in ".hwp" or hwp_name[-1] in ".HWP":
            if hwp_name[-1] != "":  # 디렉토리 인지 아닌지 체크하기 위해

                BASE_DIR = full_filename

                print(full_filename + "\\" + i + "작업중 ")
                hwp.Open(os.path.join(BASE_DIR, i))  # 한/글로 열어서
                hwp.HAction.GetDefault('FileSaveAs_S',
                                       hwp.HParameterSet.HFileOpenSave.HSet)
                hwp.HParameterSet.HFileOpenSave.filename = os.path.join(BASE_DIR,
                                                                        i.replace('.hwp', '.JPG'))

                hwp.HParameterSet.HFileOpenSave.Format = 'JPG'
                object_Hwp = i
                hwp.HAction.Execute('FileSaveAs_S', hwp.HParameterSet.HFileOpenSave.HSet)
                print(full_filename + "\\" + i + "변환완료")


                count = count + 1

        # 마지막 문서를 변환을 하고 한글 종료 시킴   , 이렇게 하지 않는다면 폴더 접근시에 한글이 열렸다 닫혔다 하기 때문에 안에서 처리 했습니다
        if count == len(file_list_hwp):
            print("한글 종료")
            win32gui.ShowWindow(hwnd, 5)
            hwp.XHwpDocuments.Close(isDirty=False)  # 열려있는 문서가 있다면 닫아줘(저장할지 물어보지 말고)
            hwp.Quit()  # 한/글 종료
 except:
     print("예외처리 발생")
     hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # 한/글 열기
     hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")  # 보안창 닫기
     # hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
     hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 해당 윈도우의 핸들값 찾기
     win32gui.ShowWindow(hwnd, 0)  # 0은 숨기기, 5는 보이기, 3은 풀스크린 등

     win32gui.ShowWindow(hwnd, 5)
     hwp.XHwpDocuments.Close(isDirty=False)  # 열려있는 문서가 있다면 닫아줘(저장할지 물어보지 말고)
     hwp.Quit()  # 한/
     print(full_filename + "\\" + object_Hwp)
     os.access(full_filename+ "\\" + object_Hwp, os.F_OK);
     os.chmod(full_filename+ "\\" + object_Hwp, 0o007)
     shutil.move(full_filename + "\\" + object_Hwp,"C:\\Users\\Administrator\\Desktop\\버그파일" + "\\" + object_Hwp)

     search_hwp_change_jpg_except_add(full_filename)







#현재 사용중인 변환 기능
def HWP_JPG_CHANGE_UPGRADE(full_filename):
 global hwp
 global count
 count = 0
 print("파일 복사중 위치 : " + full_filename)
 os.chdir(full_filename)


 # 디렉토리에 있는 모든 파일들을 불러와
 file_list = os.listdir(full_filename)
 # hwp,HWP 확장자로 구성되어있는 내용들만 리스트로 얻어옴
 file_list_hwp = [file for file in file_list if file.endswith(".hwp") or file.endswith(".HWP")]

 print(len(file_list_hwp))


 for i in file_list_hwp:
        print(count)
        hwp_name = os.path.splitext(i) #파일 확장자를 분리 시켜줌 ex) 1234.hwp -> .hwp

        # 첫번째 문서가 시작된다면 ,
        if count == 0:
           hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # 한/글 열기
           hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample") # 보안창 닫기
           # hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
           hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 해당 윈도우의 핸들값 찾기
           win32gui.ShowWindow(hwnd, 0)  # 0은 숨기기, 5는 보이기, 3은 풀스크린 등


        # 이중 방어막 의미 없는 조건문 입니다
        if hwp_name[-1] in ".hwp" or hwp_name[-1] in ".HWP":
            if hwp_name[-1] != "": # 디렉토리 인지 아닌지 체크하기 위해

                BASE_DIR = full_filename

                print(full_filename + "\\" + i+ "작업중 ")
                hwp.Open(os.path.join(BASE_DIR, i))  # 한/글로 열어서
                hwp.HAction.GetDefault('FileSaveAs_S',
                                       hwp.HParameterSet.HFileOpenSave.HSet)
                hwp.HParameterSet.HFileOpenSave.filename = os.path.join(BASE_DIR,
                                                                        i.replace('.hwp', '.JPG'))

                hwp.HParameterSet.HFileOpenSave.Format = 'JPG'

                hwp.HAction.Execute('FileSaveAs_S', hwp.HParameterSet.HFileOpenSave.HSet)
                print(full_filename + "\\" + i + "변환완료")
                count = count + 1
                Sleep(100)
                
        # 마지막 문서를 변환을 하고 한글 종료 시킴   , 이렇게 하지 않는다면 폴더 접근시에 한글이 열렸다 닫혔다 하기 때문에 안에서 처리 했습니다
        if count == len (file_list_hwp):
           print("한글 종료")
           win32gui.ShowWindow(hwnd, 5)
           hwp.XHwpDocuments.Close(isDirty=False)  # 열려있는 문서가 있다면 닫아줘(저장할지 물어보지 말고)
           hwp.Quit()  # 한/글 종료








def PDF_CHANGE(full_filename):
    global hwp

    print("파일 복사중 위치 : " + full_filename)
    os.chdir(full_filename)  # hwp 파일이 있는 폴더로 이동
    # print(os.listdir())  # 파일목록 출력해보기. 없어도 무관


    hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # 한/글 열기

    #hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
    hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 해당 윈도우의 핸들값 찾기
    win32gui.ShowWindow(hwnd, 0)  # 한/글 창을 숨겨줘. 0은 숨기기, 5는 보이기, 3은 풀스크린 등

    for i in os.listdir():  # 현재폴더 안에 있는 모든 파일을
        hwp_name = os.path.splitext(i)


        if hwp_name[-1] in ".hwp" or hwp_name[-1] in ".HWP":
            if hwp_name[-1] != "":
                # print(hwnd)  # 핸들값 출력해보기. 없어도 무관
                BASE_DIR = full_filename  # 한/글은 파일 열거나 저장할 때 전체경로를 입력해야 하므로, os.path.join(BASE_DIR, i) 식으로 사용할 것

                print(i + "작업중")
                hwp.Open(os.path.join(BASE_DIR, i))  # 한/글로 열어서

                hwp.HAction.GetDefault('FileSaveAsPdf',
                                       hwp.HParameterSet.HFileOpenSave.HSet)

                hwp.HParameterSet.HFileOpenSave.filename = os.path.join(BASE_DIR,
                                                                        i.replace('.hwp', '.PNG'))
                hwp.HParameterSet.HFileOpenSave.Format = 'PNG'
                hwp.HAction.Execute('FileSaveAsPdf', hwp.HParameterSet.HFileOpenSave.HSet)
                #딜레이

    win32gui.ShowWindow(hwnd, 5)
    hwp.XHwpDocuments.Close(isDirty=False)
    hwp.Quit()  # 한/글 종료


#  del hwp
#  del win32


#
# def PDF_CHANGE():
#
#     global hwp
#     global win32
#     global hwp_name
#     os.chdir(path.get())  # hwp 파일이 있는 폴더로 이동
#
#     # print(os.listdir())  # 파일목록 출력해보기. 없어도 무관
#
#     for i in os.listdir():  # 현재 폴더 안에 있는 모든 파일명에서
#        # print(i + "\n")
#         os.rename(i, i.replace(' - 복사본 ', ''))  # ' - 복사본 ' 부분을 지워줘.
#
#     hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # 한/글 열기
#     hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 해당 윈도우의 핸들값 찾기
#
#     print(hwnd)  # 핸들값 출력해보기. 없어도 무관
#
#     win32gui.ShowWindow(hwnd, 0)  # 한/글 창을 숨겨줘. 0은 숨기기, 5는 보이기, 3은 풀스크린 등
#     hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')  # 보안모듈 적용
#
#     BASE_DIR = path.get()  # 한/글은 파일 열거나 저장할 때 전체경로를 입력해야 하므로, os.path.join(BASE_DIR, i) 식으로 사용할 것
#     print(os.getcwd())  # 현재경로 출력. 없어도 무관
#     print(len(os.listdir()))  # 현재폴더 안에 있는 파일 갯수 출력
#
#     for i in os.listdir():  # 현재폴더 안에 있는 모든 파일을
#         hwp_name = os.path.splitext(i)
#         if hwp_name[-1] in ".hwp" or hwp_name[-1] in ".HWP":
#
#           hwp.Open(os.path.join(BASE_DIR, i))  # 한/글로 열어서
#           hwp.HAction.GetDefault('FileSaveAsPdf', hwp.HParameterSet.HFileOpenSave.HSet)  # PDF로 저장할 건데, 설정값은 아래와 같이.
#           hwp.HParameterSet.HFileOpenSave.filename = os.path.join(BASE_DIR, i.replace('.hwp', '.pdf'))  # 확장자는 .pdf로,
#           hwp.HParameterSet.HFileOpenSave.Format = 'PDF'  # 포맷은 PDF로,
#           hwp.HAction.Execute('FileSaveAsPdf', hwp.HParameterSet.HFileOpenSave.HSet)  # 위 설정값으로 실행해줘.
#
#
#     win32gui.ShowWindow(hwnd, 5)  # 다시 숨겼던 한/글 창을 보여주고,
#     hwp.XHwpDocuments.Close(isDirty=False)  # 열려있는 문서가 있다면 닫아줘(저장할지 물어보지 말고)
#     hwp.Quit()  # 한/글 종료
#   #  del hwp
#   #  del win32

def pdf_Delete(full_filename):
    global pdf
    global filename
    global filesu
    os.chdir(full_filename)
    for i in os.listdir():
        pdf = os.path.splitext(i)
        print(pdf[-1])
        if pdf[-1] in ".pdf" or pdf[-1] in ".PDF":
            if pdf[-1] != "":
                print("삭제:" + full_filename + "\\" + i)
                os.remove(full_filename + "\\" + i)


def hwp_Delete(full_filename):
    global hwp
    global filename
    global filesu
    os.chdir(full_filename)
    for i in os.listdir():
        hwp = os.path.splitext(i)
        print(hwp[-1])
        if hwp[-1] in ".hwp" or hwp[-1] in ".HWP":
            if hwp[-1] != "":
                print("삭제:" + full_filename + "\\" + i)
                os.remove(full_filename + "\\" + i)


def jpg_Delete(full_filename):
    global jpg
    global filename
    global filesu
    os.chdir(full_filename)
    for i in os.listdir():
        jpg = os.path.splitext(i)
        print(jpg[-1])
        if jpg[-1] in ".jpg" or jpg[-1] in ".JPG":
            if jpg[-1] != "":
                print("삭제:" + full_filename + "\\" + i)
                os.remove(full_filename + "\\" + i)


path = tkinter.Entry(window, width="100")
path.pack()

#현재 사용중인 변환 기능 (맨위에 있는 버튼)
button = tkinter.Button(window, overrelief="solid", width=30, command=search_hwp_change_jpg_upgrade,
                         text="HWP->JPG 변환 시작 업그레이드")
button.pack()

button2 = tkinter.Button(window, overrelief="solid", width=30, command=search_pdf_change_jpg,
                          text="PDF->JPG 변환 시작")
button2.pack()

button3 = tkinter.Button(window, overrelief="solid", width=30, command=search_hwp_change_jpg,
                        text="HWP->JPG 변환 시작 (중요)!")
button3.pack()

button4 = tkinter.Button(window, overrelief="solid", width=30, command=search_delete_hwp,
                        text="HWP 삭제 ")
button4.pack()

button5 = tkinter.Button(window, overrelief="solid", width=30, command=search_delete_pdf,
                         text="PDF 삭제 ")
button5.pack()

#사용금지
button6 = tkinter.Button(window, overrelief="solid", width=30, command=search_delete_jpg,
                          text="JPG 삭제 ")

window.mainloop()
