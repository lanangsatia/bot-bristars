# install nerodia, selenium, pandas, numpy, xlr

from nerodia.browser import *
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import pandas as pd

data = pd.read_excel('data-org.xls')
df2 = pd.DataFrame(data).astype(str)
data = df2.to_numpy()

# user = '90140981'
# password = 'Indonesia2021'
# checker = '00143365'
# user = input("PN anda : ")
# password = input("Password anda: ")
checker = input("PN Checker: ")


def bot():
    browser = Browser(browser='chrome')
    browser.goto('bristars.bri.co.id')

    # browser.div(id = 'popup_info').div(class_name = 'modal-dialog').div(class_name = 'modal-content').div(class_name = 'modal-header').button(class_name = 'close').wait_until(method=lambda e: e.present).click()

    time.sleep(1)

    #loginn bristars
    # browser.text_field(name='pernr').wait_until(method=lambda e: e.present).send_keys(user)
    # browser.text_field(name='password').wait_until(method=lambda e: e.present).send_keys(password)
    # browser.text_field(id = 'captcha').wait_until(method=lambda e: e.present).send_keys("cok")

    time.sleep(20)
    # browser.goto('brihc.bri.co.id')

    browser.button(text='BRIHC').click()
    time.sleep(2)
    browser.span(text="Struktur Organisasi").wait_until(method=lambda e: e.present).click()
    browser.span(text="Pemeliharaan Struktur Organisasi").wait_until(method=lambda e: e.present).click()
    browser.link(text="Pohon Struktur Organisasi").wait_until(method=lambda e: e.present).click()

    time.sleep(1)
    # browser.button(name = 'login').click()
    # time.sleep(1)
    # browser.span(text="Lain-Lain").wait_until(method=lambda e: e.present).click()
    # time.sleep(1)
    # browser.link(href = "/brihc.bri.co.id/").click()
    # browser.link(text= " BRIHC [A]").wait_until(method=lambda e: e.present).click()

    # browser.span(text="Struktur Organisasi").wait_until(method=lambda e: e.present).fire_event('click')
    # browser.span(text="Pemeliharaan Jabatan").wait_until(method=lambda e: e.present).fire_event('click')

    browser.refresh()
    # browser.link(href= "https://brihc.bri.co.id/jobs/filter_jobs").wait_until(method=lambda e: e.present).click()
    i = 0
    for x in data:

        i = i + 1
        try:
            print("=========================================================")
            print('Usulan :' + str(i))
            print('Sisa : ' + str(len(data) - i))
            # print('Object : ' + x[i])

            # Validasi
            time.sleep(3)
            browser.text_field(name='f_orgeh').wait_until(method=lambda e: e.present).send_keys(x[0])
            print(x[0])
            browser.execute_script('document.getElementById("f_posisi_data").removeAttribute("readonly")')
            date1 = x[1]
            browser.execute_script("""
                                var date1 = arguments[0];
                                document.getElementById('f_posisi_data').value = date1;
                                """, date1)

            # browser.text_field(name='f_posisi_data').wait_until(method=lambda e: e.present).send_keys(x[1], Keys.ENTER)

            browser.button(name='search').click()

            browser.span(text="Add Child").wait_until(method=lambda e: e.present).click()
            time.sleep(2)

            #Form orgeh
            browser.execute_script('document.getElementById("startdate").removeAttribute("readonly")')
            # browser.text_field(name='startdate').wait_until(method=lambda e: e.present).send_keys(x[2], Keys.ENTER)
            date2 = x[2]
            browser.execute_script("""
                                            var date2 = arguments[0];
                                            document.getElementById('startdate').value = date2;
                                            """, date2)
            browser.text_field(name='abbrevation').wait_until(method=lambda e: e.present).send_keys(x[3])
            browser.text_field(name='short_text').wait_until(method=lambda e: e.present).send_keys(x[4])
            # browser.text_field(name='auto_costcenter').wait_until(method=lambda e: e.present).send_keys(x[5])
            # browser.text_field(name='auto_costcenter').wait_until(method=lambda e: e.present).send_keys(Keys.ARROW_DOWN, Keys.ENTER)
            browser.select_list(id="jenis_uker").select(x[6])
            # browser.select_list(id="divisi").select(x[7])
            browser.text_field(id='cs1').wait_until(method=lambda e: e.present).send_keys(checker)
            time.sleep(1)
            browser.text_field(id='cs1').wait_until(method=lambda e: e.present).send_keys(Keys.ARROW_DOWN, Keys.TAB)

            browser.button(name='save').click()
            time.sleep(1)

            if browser.alert.exists == True:
                browser.alert.ok()
            time.sleep(3)
            print("OK!")
            browser.refresh()
        except Exception as e:
            print("EROOOOOOORRRRRRRRRRRRR!!!!!")
            print(e)
        finally:
            browser.refresh()

    print("=========================================================")
    print("=========================================================")
    print("=========================================================")
    print("=========================================================")
    print("=========================================================")
    print("Done ")
    browser.close()


bot()
