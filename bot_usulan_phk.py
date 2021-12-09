# install nerodia, selenium, pandas, numpy, xlr

from nerodia.browser import *
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import pandas as pd

data = pd.read_excel('data-phk.xls')
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

    time.sleep(10)
    # browser.goto('brihc.bri.co.id')

    browser.button(text='BRIHC').click()
    time.sleep(2)
    browser.span(text="Pengembangan Karir").wait_until(method=lambda e: e.present).click()
    browser.link(text="Mutasi").wait_until(method=lambda e: e.present).click()

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
            browser.text_field(id='f_pn').wait_until(method=lambda e: e.present).send_keys(x[0])
            print(x[0])
            time.sleep(7)
            browser.text_field(id='f_pn').wait_until(method=lambda e: e.present).send_keys(Keys.ARROW_DOWN, Keys.TAB)
            browser.execute_script('document.getElementById("startdate_input").removeAttribute("readonly")')
            browser.text_field(name='startdate_input').wait_until(method=lambda e: e.present).send_keys(x[1], Keys.ENTER)
            browser.select_list(id="opt_actions").select('Pemutusan Hubungan Kerja (PHK)')

            browser.button(name='validasi').click()
            time.sleep(1)

            # Deskripsi Posisi
            browser.select_list(id="opt_eg").select('Non Aktif')
            browser.select_list(id="opt_esg").select(x[4])


            # Dokumen Pendukung
            browser.text_field(name='nomor_sk').wait_until(method=lambda e: e.present).send_keys(x[5])
            browser.execute_script('document.getElementById("tanggal_sk").removeAttribute("readonly")')
            browser.text_field(name='tanggal_sk').wait_until(method=lambda e: e.present).send_keys(x[6], Keys.ENTER)
            browser.text_field(name='nomor_urut').wait_until(method=lambda e: e.present).send_keys(x[7])
            browser.select_list(id="kode_perihal_sk").select(x[8])

            time.sleep(5)

            browser.button(name='buat_usulan').click()
            time.sleep(1)

            if browser.alert.exists == True:
                browser.alert.ok()

            time.sleep(3)
            browser.text_field(id='cs1').wait_until(method=lambda e: e.present).send_keys(checker)
            time.sleep(1)
            browser.text_field(id='cs1').wait_until(method=lambda e: e.present).send_keys(Keys.ARROW_DOWN, Keys.TAB)

            time.sleep(300)
            # browser.button(name='kirim').click()


            # time.sleep(1000)
            #         browser.refresh()

            #         browser.button(name = 'save').click()
            time.sleep(1)
            if browser.alert.exists == True:
                browser.alert.ok()
            time.sleep(3)
            print("OK!")
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
