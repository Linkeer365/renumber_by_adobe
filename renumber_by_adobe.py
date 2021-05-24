import win32gui
import win32api
import win32con

import os
import sys

import time

target_dir=r"D:\AllDowns\newbooks"
root_hd = None

ab_rf_path=r"D:\AllDowns\newbooks\page_mid_pic\finished_pages_ab_rf.txt"

already_path=r"D:\AllDowns\newbooks\already_trans.txt"

if not os.path.exists(already_path):
    open(already_path, 'a').close()

def click_on_pos(pos_list):
    btn_pos = pos_list
    win32api.SetCursorPos(btn_pos)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

def get_hd_from_child_hds(father_hd,some_idx,expect_name):
    child_hds=[]
    win32gui.EnumChildWindows(father_hd,lambda hwnd, param: param.append(hwnd),child_hds)

    names=[win32gui.GetWindowText(each) for each in child_hds]
    hds=[hex(each) for each in child_hds]
    print("ChildName List:",names)
    print("Child Hds List:",hds)

    name=names[some_idx]
    hd=hds[some_idx]

    print("The {} Child.".format(some_idx))
    print("The Name:{}".format(name))
    print("The HD:{}".format(hd))

    if name==expect_name:
        return child_hds[some_idx]
    else:
        print("窗口不对！")
        return None

def renumber(pdf_path,delta_page):
    # pdf_path=r"D:\AllDowns\newbooks\typetype1致歉信.pdf"
    os.startfile(pdf_path)
    time.sleep(2)

    adobe_str="Adobe Acrobat"
    adobe_hd=win32gui.FindWindowEx(root_hd,0,0,adobe_str)
    # if adobe_hd:
    #     queding_hd=get_hd_from_child_hds(adobe_hd,4,"")
    #     win32gui.SendMessage(queding_hd, win32con.BM_CLICK)

    bianpai_pos=[578,56]
    click_on_pos(bianpai_pos)

    time.sleep(1)

    bianpai_str="编排页码"
    bianpai_hd=win32gui.FindWindowEx(root_hd,0,0,bianpai_str)
    print("++--++")
    fst_hd=get_hd_from_child_hds(bianpai_hd,0,"")

    # start_hd=get_hd_from_child_hds(bianpai_hd,5,"")
    end_hd=get_hd_from_child_hds(bianpai_hd,7,"")

    yangshi_hd=get_hd_from_child_hds(bianpai_hd,14,"")

    queding_hd=get_hd_from_child_hds(bianpai_hd,23,"确定")

    # 一遍遍试出来的，应该是可以跟Edit窗体调用一样的函数...
    # https://stackoverflow.com/questions/6206656/sendmessage-from-a-delphi-app-to-a-java-app-richedit50w-control
    # 我是看到SendMessage is working find.才意识到这一点的
    win32api.SendMessage(end_hd,win32con.WM_SETTEXT,0,f"{delta_page}")

    # time.sleep(1)
    # https://stackoverflow.com/questions/20661733/select-a-combobox-via-winapi-in-python
    # 一律使用大写的ABC这种，然后这是第六个所以idx=5
    idx=5
    win32gui.SendMessage(yangshi_hd,win32con.CB_SETCURSEL,idx,0)

    # time.sleep(1)

    win32gui.SendMessage(queding_hd,win32con.BM_CLICK)

    # https://www.cnblogs.com/pylemon/archive/2011/09/07/2169972.html
    # 进行保存

    win32api.keybd_event(17,0,0,0)  #ctrl键位码是17
    win32api.keybd_event(83,0,0,0)  #S键位码是83
    win32api.keybd_event(83,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)

    # 有些保存相当耗时...

    time.sleep(5)

    win32gui.SendMessage(adobe_hd,win32con.WM_CLOSE,0,0)

    time.sleep(2)

    with open(already_path,"a",encoding="utf-8") as f:
        f.write(pdf_path+"\n")

    print("one done.")

def main():
    with open(ab_rf_path,"r",encoding="utf-8") as f:
        lines=f.readlines()

    lines=[each.strip("\n") for each in lines]

    with open(already_path,"r",encoding="utf-8") as g:
        already_set=set(g.readlines())

    for each_line in lines:
        delta_zone,_,_,filename_zone=each_line.split("\t\t\t")
        delta_page=delta_zone.split(":")[1]
        filename=filename_zone.split(":")[1]
        pdf_path=f"{target_dir}{os.sep}{filename}"
        if int(delta_page)<0:
            print("Delta page negative,please check:",pdf_path)
            continue
        if pdf_path+"\n" in already_set:
            print(f"already:{pdf_path}")
            continue
        if not os.path.exists(pdf_path):
            print("有点问题，不纠结...")
            continue
        if int(delta_page)!=0:
            renumber(pdf_path,delta_page)
        else:
            print("重合！跳过...")
            continue
    print("all done.")

if __name__=="__main__":
    main()




# # length = win32gui.SendMessage(bookmark_hd, win32con.WM_GETTEXTLENGTH)*2+2
# length = win32gui.SendMessage(end_hd, win32con.WM_GETTEXTLENGTH) * 2 + 1
# time.sleep(0.5)
# buf = win32gui.PyMakeBuffer(length)
# # 发送获取文本请求
# win32api.SendMessage(end_hd, win32con.WM_GETTEXT, length, buf)
# time.sleep(1)
# # 下面应该是将内存读取文本
# address, length = win32gui.PyGetBufferAddressAndLen(buf[:-1])
# # time.sleep(0.5)
# text = win32gui.PyGetString(address, length)




# while True:
#     tempt = win32api.GetCursorPos() # 记录鼠标所处位置的坐标
#     # x = tempt[0]-choose_rect[0] # 计算相对x坐标
#     # y = tempt[1]-choose_rect[1] # 计算相对y坐标
#     print(tempt)
#     # time.sleep(0.5) # 每0.5s输出一次


