#coding:utf-8

#name:51jobV2.1.py
#progrom:51job批量投递简历,V2.1,数据增加薪资筛选,增加BlackList和Report
#notice:预制条件为,用户已注册51job,并且默认简历已经建立好
#user:bj122韩志超
#date:2016-10-\8

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import unittest, time, re
import xlrd


class Auto_send_resume_in51job(unittest.TestCase):

    def setUp(self):
        #调用浏览器
        self.driver=webdriver.Firefox()
        #self.driver.implicitly_wait(3)
        time.sleep(3)
        #最大化浏览器
        self.driver.maximize_window()

    def login(self,username,password):
        """登录"""
        
        driver=self.driver

        driver.find_element_by_id('username').send_keys(username)    #用户名    #需要判断异常吗?
        driver.find_element_by_id('userpwd').send_keys(password)     #密码
        driver.find_element_by_class_name('btn').click()             #点击确定

    def search(self,keyword,region):
        """搜索"""
        time.sleep(3)
        driver=self.driver
        driver.implicitly_wait(3)

        global search_window
        search_window=driver.current_window_handle  #获取搜索窗口句柄
        driver.find_element_by_class_name('tSearch_inp_text').clear()
        driver.find_element_by_class_name('tSearch_inp_text').send_keys(keyword) #汉字参数需要转化吗?

        #选择地区
        driver.find_element_by_id('btnJobarea').click()
        driver.find_element_by_link_text(region).click()
        driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[2]/form/div[2]/a').click()

    def send_resume(self):
        """投递简历"""
        
        driver=self.driver
        
        #获得当前所有打开的窗口的句柄
        all_handles = driver.window_handles

        #切换至投递简历窗口
        for handle in all_handles:
            if handle != search_window:
                driver.switch_to_window(handle)

        #筛选4501-8000
        driver.find_element_by_link_text('4501-8000').click()


        #获取结果页数
        pagesresult=driver.find_element_by_class_name('td').text;

        pages=int(pagesresult[1:-4])

        

        for i in range(1,pages+1):
            
            print('共有 %d 页') % (pages)
            print("第 %d 页") % (i)
            
            #申请职位
            driver.find_element_by_id('top_select_all_jobs_checkbox').click()#点击全选复选框
            driver.find_element_by_class_name('but_sq').click()#点击申请职位
            driver.implicitly_wait(5)

            #处理第一次投递弹出选择简历框
            try:
                driver.find_element_by_name('qpostset').click()#点击快速投递
                driver.find_element_by_id('apply_now').click()#点击立即申请
            except:
                pass

            #处理已投递简历对话框
            try:
                driver.switch_to_alert().accept()
            except:
                 pass

            #关闭投递完成对话框
            try:
                driver.find_element_by_id('window_close_apply').click()
            except:
                pass
            
            #点击下一页
            
            if i!=pages:
                print ('i=%d\tpages=%d\t') %(i,pages)
                #driver.implicitly_wait(3)
                try:
                    driver.find_element_by_id('rtNext').click()
                except:
                    print "已经最后一页"
            else:
                driver.close()
                #获得当前所有打开的窗口的句柄
                all_handles = driver.window_handles

                #切换至搜索窗口
                for handle in all_handles:
                    if handle == search_window:
                        driver.switch_to_window(handle)

            #下一页
            i+=1
        
    def logout(self):
        driver=self.driver
        driver.find_element_by_id('loginOutLink').click()

    def test_login_search_sendresume(self):
        driver=self.driver
        try:
            wb=xlrd.open_workbook('data.xls')
        except IOError:                #?使用什么异常?
            print "数据文件丢失或无法打开！"
        sh=wb.sheet_by_index(0)
        rows=sh.nrows
        for i in range(1,rows):
            username=sh.cell_value(i,0)
            password=sh.cell_value(i,1)
            keywords=sh.cell_value(i,2)
            region=sh.cell_value(i,3)

            print username,password  #测试用---------------

            time.sleep(5)
            #访问51job网址
            print "......................"
            driver.get('http://www.51job.com/')
            time.sleep(5) 
            print "......................"
            #调用登陆函数
            self.login(username,password)

            #判断是否登陆成功
            #act=driver.find_element_by_id('top_username').text
            #print act
            '''
            try:                    #反思,一定能定位到吗?
                driver.find_element_by_id('top_username')
                print "login sucess"
            except NoSuchElementException,e:
                print "login error"   #测试用--------------
                return False
            '''
            print keywords
            
            for keyword in keywords.split(','):

                #调用搜索函数
                self.search(keyword,region)
                #调用投递简历函数
                self.send_resume()
                
            self.logout()

            i+=1
            
    def tearDown(self):
        self.driver.quit()

if __name__ == "__main__":
    unittest.main()
