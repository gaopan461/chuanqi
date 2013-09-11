# -*- coding: gb2312 -*-
import win32com.client
import sys
import time

class App:
    def __init__(self):
        self.m_objDM = win32com.client.Dispatch("dm.dmsoft")
        assert(self.m_objDM.SetPath("F:\\3.1233") == 1)
        assert(self.m_objDM.SetDict(0,"hanzi.txt") == 1)
        assert(self.m_objDM.SetDict(1,"shuzi.txt") == 1)

        # ���ҵ�½������
        self.m_hLoginWnd = self.m_objDM.GetMousePointWindow()
        if self.m_hLoginWnd == 0:
            print "Error:init,FindWindow login failed"
            sys.exit(-1)

        print u"��ȡ����½�����ھ��:",self.m_hLoginWnd

    def run(self):
        assert(self.m_objDM.UseDict(0) == 1)
        city = self.m_objDM.Ocr(250,715,297,731,"ffffff-000000",1.0)
        print "city",city

        assert(self.m_objDM.UseDict(1) == 1)
        pos = self.m_objDM.Ocr(299,717,344,728,"ffffff-000000",1.0)
        print "position:",pos

    def login(self):
        # ��󻯲������½������
        assert(self.m_objDM.SetWindowState(self.m_hLoginWnd,12) == 1)
        time.sleep(3)

        # �󶨵�½������
        assert(self.m_objDM.UnBindWindow() == 1)
        assert(self.m_objDM.BindWindow(self.m_hLoginWnd,"gdi","normal","normal",0) == 1)
    
        # ���ҷ�������·
        pos = self.m_objDM.FindStrFastE(232,84,273,165,"һ��","f1bb4b-000000",1.0)
        pos = pos.split('|')
        if int(pos[0]) == -1:
            print "Error:login,FindStrFastE failed"
            sys.exit(-1)

        # ���ѡ�з�����
        assert(self.m_objDM.MoveTo(int(pos[1]),int(pos[2])) == 1)
        time.sleep(0.2)
        assert(self.m_objDM.LeftClick() == 1)

        # �жϷ���������״̬
        pos = self.m_objDM.FindColorE(127,461,217,480,"00ff00-000000",1.0,0)
        pos = pos.split('|')
        count = 0
        while int(pos[0]) < 0:
            count = count + 1
            if count > 10:
                print "Error:login,server connect failed"
                sys.exit(-1)

            time.sleep(1)
            pos = self.m_objDM.FindColorE(127,461,217,480,"00ff00-000000",1.0,0)
            pos = pos.split('|')

        print u"������״̬����"

        # ���������Ϸ
        assert(self.m_objDM.MoveTo(378,467) == 1)
        time.sleep(0.2)
        assert(self.m_objDM.LeftClick() == 1)
        time.sleep(0.2)

        # �ж���Ϸ�����Ƿ��
        hwnd = self.m_objDM.FindWindow("","Legend of mir2")
        count = 0
        while hwnd == 0:
            count = count + 1
            if count > 120:
                print "Error:login,FindWindow timeout"
                sys.exit(-1)

            time.sleep(1)
            hwnd = self.m_objDM.FindWindow("","Legend of mir2")

        self.m_hGameWnd = hwnd
        print u"��ȡ����Ϸ���ھ��:",self.m_hGameWnd

        # ��󻯲�������Ϸ����
        assert(self.m_objDM.SetWindowState(self.m_hGameWnd,12) == 1)
        time.sleep(3)

        # �󶨵���Ϸ����
        assert(self.m_objDM.UnBindWindow() == 1)
        assert(self.m_objDM.BindWindow(self.m_hGameWnd,"gdi","normal","normal",0) == 1)

        # �����û��������룬�س�������Ϸ
        time.sleep(1)
        self.inputUsernameAndPassword("gaopan461","gaopan461")

    def inputUsernameAndPassword(self,username,password):
        for c in username:
            assert(self.m_objDM.KeyPressChar(c) == 1)
            time.sleep(0.2)

        assert(self.m_objDM.KeyPressChar("tab") == 1)
        time.sleep(0.2)

        for c in password:
            assert(self.m_objDM.KeyPressChar(c) == 1)
            time.sleep(0.2)

        assert(self.m_objDM.KeyPressChar("enter") == 1)
        time.sleep(0.2)

if __name__ == '__main__':
    app = App()
    app.login()