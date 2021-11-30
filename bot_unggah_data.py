# install nerodia, selenium, pandas, numpy, xlr

from nerodia.browser import *
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import pandas as pd

data = pd.read_excel('data.xls')
df2 = pd.DataFrame(data).astype(str)
data = df2.to_numpy()

# data = [["Penerimaan Trainee", "18-11-2021", "70112914",  # validasi
#          "PY1-Payroll Admin. 1", "T10-Div.MAT & Logsitik", "81837198",  # desc posisi
#          "121", "19-11-2021", "12121", "1-Promosi Pangkat (History)",  # dok pendukung
#          "nama", "2", "Dr.", "AMd.", "jakarta", "20-11-2000", "Indonesia", "Indonesian", "Indonesian", "A", "Belum", "",
#          "Islam", "10001 - Cl. Wakakanca I", "120",  # data pribadi
#          "Alamat tetap", "alamat rumah", "Indonesia", "BALI", "BADUNG", "KUTA", "KUTA", "09", "08", "12",  # alamat
#          "08921213331", "CELL", "08999998838",  # telepon
#          "228301001756509",  # rekening
#          "939082657416000", "No Dependent",  # pajak
#          "Kartu Tanda Penduduk", "2223331212", "jakarta", "Indonesia", "18-11-2021", "31-12-9999",  # nomor identity
#          "18-11-2021", "18-11-2021", "Strata 1", "ui", "Indonesia", "Negeri", "jurusan", "sub jurusan", "400", "2021",
#          # pendidikan
#          "Eselon 1", "Direksi", "Eselon 1", "Eselon 1"],  # hak dinas
#         ["Penerimaan Trainee", "18-11-2021", "70246747",
#          "PY1-Payroll Admin. 1", "T10-Div.MAT & Logsitik", "81837198",
#          "121", "19-11-2021", "12121", "1-Promosi Pangkat (History)",
#          "nama", "2", "Dr.", "AMd.", "jakarta", "20-11-2000", "Indonesia", "Indonesian", "Indonesian", "A", "Belum", "",
#          "Islam", "10001 - Cl. Wakakanca I", "120",
#          "Alamat tetap", "alamat rumah", "Indonesia", "BALI", "BADUNG", "KUTA", "KUTA", "09", "08", "12",
#          "08921213331", "CELL", "08999998838",
#          "228301001756509",
#          "93-9082657-416000", "No Dependent",
#          "Kartu Tanda Penduduk", "2223331212", "jakarta", "Indonesia", "18-11-2021", "18-11-2021",
#          "18-11-2021", "18-11-2021", "Strata 1", "ui", "Indonesia", "Negeri", "jurusan", "sub jurusan", "400", "2021",
#          "Eselon 1", "Direksi", "Eselon 1", "Eselon 1"],
#         ["Penerimaan Trainee", "18-11-2021", "70246747",
#          "PY1-Payroll Admin. 1", "T10-Div.MAT & Logsitik", "81837198",
#          "121", "19-11-2021", "12121", "1-Promosi Pangkat (History)",
#          "nama", "2", "Dr.", "AMd.", "jakarta", "20-11-2000", "Indonesia", "Indonesian", "Indonesian", "A", "Belum", "",
#          "Islam", "10001 - Cl. Wakakanca I", "120",
#          "Alamat tetap", "alamat rumah", "Indonesia", "BALI", "BADUNG", "KUTA", "KUTA", "09", "08", "12",
#          "08921213331", "CELL", "08999998838",
#          "228301001756509",
#          "93-9082657-416000", "No Dependent",
#          "Kartu Tanda Penduduk", "2223331212", "jakarta", "Indonesia", "18-11-2021", "18-11-2021",
#          "18-11-2021", "18-11-2021", "Strata 1", "ui", "Indonesia", "Negeri", "jurusan", "sub jurusan", "400", "2021",
#          "Eselon 1", "Direksi", "Eselon 1", "Eselon 1"]]

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
    browser.span(text="Penerimaan Pekerja").wait_until(method=lambda e: e.present).click()

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
            print('Object : ' + x[10])

            # Validasi
            time.sleep(3)
            browser.select_list(id="reason").select(x[0])
            browser.execute_script('document.getElementById("startdate").removeAttribute("readonly")')
            browser.text_field(name='startdate').wait_until(method=lambda e: e.present).send_keys(x[1], Keys.ENTER)
            browser.text_field(name='posisi').wait_until(method=lambda e: e.present).send_keys(x[2])

            browser.button(name='validasi').click()

            time.sleep(3)

            # Deskripsi Posisi
            browser.select_list(id="payroll_administrator").select(x[3])
            browser.select_list(id="time_administrator").select(x[4])
            browser.text_field(name='nip').wait_until(method=lambda e: e.present).send_keys(x[5])

            # Dokumen Pendukung
            browser.text_field(name='nomor_sk').wait_until(method=lambda e: e.present).send_keys(x[6])
            browser.execute_script('document.getElementById("tanggal_sk").removeAttribute("readonly")')
            browser.text_field(name='tanggal_sk').wait_until(method=lambda e: e.present).send_keys(x[7], Keys.ENTER)
            browser.text_field(name='nomor_urut').wait_until(method=lambda e: e.present).send_keys(x[8])
            browser.select_list(id="kode_perihal_sk").select(x[9])

            # Data Pribadi
            browser.text_field(name='nama_pekerja').wait_until(method=lambda e: e.present).send_keys(x[10])
            # radio input
            radioset = browser.radio_set(name='jk')
            radioset.select(x[11])
            browser.select_list(id="gelar_depan").select(x[12])
            browser.select_list(id="gelar_belakang").select(x[13])
            browser.text_field(name='tempat_lahir').wait_until(method=lambda e: e.present).send_keys(x[14])
            browser.execute_script('document.getElementById("tanggal_lahir").removeAttribute("readonly")')
            browser.text_field(name='tanggal_lahir').wait_until(method=lambda e: e.present).send_keys(x[15], Keys.ENTER)
            browser.select_list(id="negara_kelahiran").select(x[16])
            browser.select_list(id="kewarganegaraan").select(x[17])
            browser.select_list(id="bahasa").select(x[18])
            browser.select_list(id="golongan_darah").select(x[19])
            browser.select_list(id="status_kawin").select(x[20])
            browser.execute_script('document.getElementById("tanggal_status_kawin").removeAttribute("readonly")')
            browser.text_field(name='tanggal_status_kawin').wait_until(method=lambda e: e.present).send_keys(x[21],
                                                                                                             Keys.ENTER)
            browser.select_list(id="agama").select(x[22])
            browser.select_list(id="program_masuk").select(x[23])
            browser.text_field(name='angkatan').wait_until(method=lambda e: e.present).send_keys(x[24])

            # Data Keluarga
            # ....

            # Alamat
            browser.select_list(id="tipe_alamat").select(x[25])
            browser.text_field(name='alamat_rumah').wait_until(method=lambda e: e.present).send_keys(x[26])
            browser.select_list(id="negara_alamat").select(x[27])
            browser.select_list(id="provinsi").select(x[28])
            browser.select_list(id="kota").select(x[29])
            browser.select_list(id="kecamatan").select(x[30])
            browser.select_list(id="kelurahan").select(x[31])
            browser.text_field(name='rt').wait_until(method=lambda e: e.present).send_keys(x[32])
            browser.text_field(name='rw').wait_until(method=lambda e: e.present).send_keys(x[33])
            browser.text_field(name='kodepos').wait_until(method=lambda e: e.present).send_keys(x[34])

            # Komunikasi
            browser.text_field(name='nomor_komunikasi_utama').wait_until(method=lambda e: e.present).send_keys(x[35])
            browser.select_list(id="tipe_komunikasi_1").select(x[36])
            browser.text_field(name='nomor_komunikasi_1').wait_until(method=lambda e: e.present).send_keys(x[37])

            # Rekening
            browser.text_field(name='nomor_rekening').wait_until(method=lambda e: e.present).send_keys(x[38])

            # Data Pajak
            npwp = x[39]
            # browser.execute_script('document.getElementById("npwp").value = '{}'').format(npwp)
            # browser.text_field(name='npwp').wait_until(method=lambda e: e.present).send_keys(npwp)
            browser.execute_script("""
            var npwp = arguments[0];
            document.getElementById('npwp').value += npwp;
            """, npwp)
            browser.text_field(name='npwp').wait_until(method=lambda e: e.present).click()

            browser.select_list(id="tax_dependent").select(x[40])

            # checkbox
            # checkbox = browser.checkbox(id='menikah_pajak')
            #         # checkbox.set()

            # Data Jamsostek
            # ...

            # Nomor Identitas
            browser.select_list(id="jenis_identitas").select(x[41])
            browser.text_field(name='nomor_ktp').wait_until(method=lambda e: e.present).send_keys(x[42])
            browser.text_field(name='kota_dikeluarkan').wait_until(method=lambda e: e.present).send_keys(x[43])
            browser.select_list(id="negara_ktp").select(x[44])
            browser.execute_script('document.getElementById("tanggal_dikeluarkan").removeAttribute("readonly")')
            browser.text_field(name='tanggal_dikeluarkan').wait_until(method=lambda e: e.present).send_keys(x[45],
                                                                                                            Keys.ENTER)
            browser.execute_script('document.getElementById("berlaku_sampai").removeAttribute("readonly")')
            browser.text_field(name='berlaku_sampai').wait_until(method=lambda e: e.present).send_keys(x[46],
                                                                                                       Keys.ENTER)

            # Pendidikan
            browser.execute_script('document.getElementById("startdate_edu").removeAttribute("readonly")')
            browser.execute_script('document.getElementById("enddate_edu").removeAttribute("readonly")')
            browser.text_field(name='startdate_edu').wait_until(method=lambda e: e.present).send_keys(x[47], Keys.ENTER)
            browser.text_field(name='enddate_edu').wait_until(method=lambda e: e.present).send_keys(x[48], Keys.ENTER)
            browser.select_list(id="strata_edu").select(x[49])
            browser.text_field(name='institusi_edu').wait_until(method=lambda e: e.present).send_keys(x[50])
            browser.select_list(id="negara_edu").select(x[51])
            browser.select_list(id="status_institusi").select(x[52])
            browser.text_field(name='major').wait_until(method=lambda e: e.present).send_keys(x[53])
            browser.text_field(name='submajor').wait_until(method=lambda e: e.present).send_keys(x[54])
            browser.execute_script('document.getElementById("gpa").classList.remove("trim")')
            browser.execute_script('document.getElementById("gradyear").classList.remove("trim")')
            # browser.text_field(name='gpa').wait_until(method=lambda e: e.present).send_keys(x[55])
            # browser.text_field(name='gradyear').wait_until(method=lambda e: e.present).send_keys(x[56])
            gpa = x[55]
            gradyear = x[56]
            browser.execute_script("""
                    var gpa = arguments[0];
                    document.getElementById('gpa').value += gpa;
                    """, gpa)
            browser.text_field(name='gpa').wait_until(method=lambda e: e.present).click()

            browser.execute_script("""
                    var gradyear = arguments[0];
                    document.getElementById('gradyear').value += gradyear;
                    """, gradyear)
            # biayabri checkbox

            # Pengalaman Kerja
            # ...

            # Hak Perjalanan Dinas
            browser.select_list(id="statutory_travel").select(x[57])
            browser.select_list(id="enterprise_travel").select(x[58])
            browser.select_list(id="exp_type_travel").select(x[59])
            browser.select_list(id="egroup_travel").select(x[60])

            browser.button(name='konfirmasi').click()

            time.sleep(3)
            browser.text_field(id='cs1').wait_until(method=lambda e: e.present).send_keys(78334)
            time.sleep(1)
            browser.text_field(id='cs1').wait_until(method=lambda e: e.present).send_keys(Keys.ARROW_DOWN, Keys.TAB)

            browser.button(id='submit').click()

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
