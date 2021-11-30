# install nerodia, selenium, pandas, numpy, xlr

from nerodia.browser import *
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import pandas as pd

# user = '90140981'
# password = 'Indonesia2021'
# checker = '00143365'
# user = input("PN anda : ")
# password = input("Password anda: ")
# checker = input("PN Checker: ")


def bot():
    browser = Browser(browser='chrome')
    browser.goto('bristars.bri.co.id')
    #
    # # browser.div(id = 'popup_info').div(class_name = 'modal-dialog').div(class_name = 'modal-content').div(class_name = 'modal-header').button(class_name = 'close').wait_until(method=lambda e: e.present).click()
    #
    # time.sleep(1)
    #
    # #loginn bristars
    # browser.text_field(name='pernr').wait_until(method=lambda e: e.present).send_keys(user)
    # browser.text_field(name='password').wait_until(method=lambda e: e.present).send_keys(password)
    # # browser.text_field(id = 'captcha').wait_until(method=lambda e: e.present).send_keys("cok")
    #
    time.sleep(15)
    # # browser.goto('brihc.bri.co.id')
    #
    browser.button(text='BRIHC').click()
    time.sleep(2)
    browser.span(text="Persetujuan").wait_until(method=lambda e: e.present).click()
    browser.link(text="Unggah Pemeliharan SDM").wait_until(method=lambda e: e.present).click()
    time.sleep(10)

    # browser.button(name = 'login').click()
    # time.sleep(1)
    # browser.span(text="Lain-Lain").wait_until(method=lambda e: e.present).click()
    # time.sleep(1)
    # browser.link(href = "/brihc.bri.co.id/").click()
    # browser.link(text= " BRIHC [A]").wait_until(method=lambda e: e.present).click()

    # browser.span(text="Struktur Organisasi").wait_until(method=lambda e: e.present).fire_event('click')
    # browser.span(text="Pemeliharaan Jabatan").wait_until(method=lambda e: e.present).fire_event('click')

    # browser.refresh()
    # browser.link(href= "https://brihc.bri.co.id/jobs/filter_jobs").wait_until(method=lambda e: e.present).click()
    i = 0
    for x in range(2):

        i = i + 1
        try:
            print("=========================================================")
            print('Iterate:' + str(i))
            # print('Sisa : ' + str(len(data) - i))
            # print('Object : ' + x[10])

            browser.button(id='selectall').click()
            browser.button(id='btn_tolak').click()

            # time.sleep(1000)
            #         browser.refresh()

            #         browser.button(name = 'save').click()
            time.sleep(1)
            if browser.alert.exists == True:
                browser.alert.ok()
            time.sleep(15)
            print("OK!")
        except Exception as e:
            print("EROOOOOOORRRRRRRRRRRRR!!!!!")
            print(e)
        # finally:
        #     browser.refresh()

    print("=========================================================")
    print("=========================================================")
    print("=========================================================")
    print("=========================================================")
    print("=========================================================")
    print("Done ")
    browser.close()


bot()
