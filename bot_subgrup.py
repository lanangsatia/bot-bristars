# install nerodia, selenium, pandas, numpy, xlr

from nerodia.browser import *
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import pandas as pd

data = pd.read_excel('data-subgrup.xls').fillna('')
df2 = pd.DataFrame(data).astype(str)
data = df2.to_numpy()


# user = '90140981'
# password = '-'
# checker = '78334'
pn = str(data[0][0]).zfill(8)
# print(data[0])
checker = input("PN Checker: ")


def bot():
    browser = Browser(browser='chrome')
    browser.goto('bristars.bri.co.id')

    # browser.div(id = 'popup_info').div(class_name = 'modal-dialog').div(class_name = 'modal-content').div(class_name = 'modal-header').button(class_name = 'close').wait_until(method=lambda e: e.present).click()

    #loginn bristars
    # browser.text_field(name='pernr').wait_until(method=lambda e: e.present).send_keys(user)
    # browser.text_field(name='password').wait_until(method=lambda e: e.present).send_keys(password)
    # browser.text_field(id = 'captcha').wait_until(method=lambda e: e.present).send_keys("cok")

    time.sleep(20)
    # browser.goto('brihc.bri.co.id')

    browser.button(text='BRIHC').click()
    time.sleep(2)
    browser.span(text="Data Utama SDM").wait_until(method=lambda e: e.present).click()
    browser.link(text="Pemeliharaan Data").wait_until(method=lambda e: e.present).click()

    time.sleep(1)

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
            time.sleep(1)
            pn = str(data[i-1][0]).zfill(8)
            browser.text_field(name='f_pn').wait_until(method=lambda e: e.present).send_keys(pn)
            print(x[0])
            time.sleep(2)
            browser.text_field(name='f_pn').wait_until(method=lambda e: e.present).send_keys(Keys.ARROW_DOWN, Keys.TAB)


            browser.button(name='check').click()

            browser.button(name='display').click()

            browser.select_list(id="opt_kelas").select('Basic Personal Data')
            browser.select_list(id="opt_infotype").select('Organizational Assignment')

            browser.button(name='display').click()

            time.sleep(1)
            browser.span(class_name='view').wait_until(method=lambda e: e.present).click()
            time.sleep(1)

            browser.select_list(id="employee_subgroup").select(x[1])

            browser.button(name='maintain').click()
            browser.button(name='maintain').send_keys(Keys.TAB, x[2])

            #checker
            browser.text_field(id='cs1').wait_until(method=lambda e: e.present).send_keys(checker)
            time.sleep(1)
            browser.text_field(id='cs1').wait_until(method=lambda e: e.present).send_keys(Keys.ARROW_DOWN, Keys.TAB)

            browser.button(name='edit').click()
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
