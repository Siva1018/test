# -*- coding: utf-8 -*-
"""configuring settings"""
from django.conf import settings
if not settings.configured:
    settings.configure()
settings.EMAIL_USE_TLS=True
settings.EMAIL_HOST_USER = 'contact@aptuz.com'
settings.EMAIL_HOST_PASSWORD = 'Contact@123'
settings.EMAIL_HOST='smtp.gmail.com'
settings.EMAIL_PORT='587'


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re, os, xlwt, datetime, operator
from django.conf import settings
from django.core.mail import EmailMessage

class Test(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.implicitly_wait(30)
        self.base_url = "http://www.google.com/"
        print 'started'
        self.verificationErrors = []
        self.accept_next_alert = True

    def test_(self):
        driver = self.driver
        today = datetime.datetime.now()
        looped_hotel_data_1={}
        looped_hotel_data_2={}
        looped_hotel_data_3={}
        looped_hotel_data_4={}

        for i in range(7):

            if i==5:
                x=15
            elif i==6:
                x=30
            else:
                x=i

            """ set up required dates data """
            driver.implicitly_wait(10)
            dt = today + datetime.timedelta(days=x)
            ndt = today + datetime.timedelta(days=x+1)
            indate = str(dt.month).zfill(2)+"%2F"+str(dt.day).zfill(2)+"%2F"+str(dt.year)
            outdate = str(ndt.month).zfill(2)+"%2F"+str(ndt.day).zfill(2)+"%2F"+str(ndt.year)

            """ hotel laquinta """
            try:
                laquinta_url = 'http://www.lq.com/bin/lq-com/roomSearch.html?sessionId=a8e97e90-9e06-4823-ad46-01224fb96a20&from=%2Fcontent%2Flq%2Flq-com%2Fen%2Fnavigation%2Ffindandbook%2Fdynamic-pages%2Fhotel-details&indate='+indate+'&outdate='+outdate+'&location=New+Braunfels%2C+TX&searchType=address&bookingPage=address&lat=29.699459&lon=-98.09222&hotelId=0254&adults=1&numChildren=0&rooms=1&specialRates=RAC&promoCode=&currencyCode=USD&searchRadius=40&smokingPreference=NSMK&searchWithReturns=false&excludeFromStart=&corridorRadius=10&searchDCity=&searchDState=&searchOCity=&searchOState='
                driver.get(laquinta_url)
                if driver.current_url!=laquinta_url:
                    driver.get(laquinta_url)
                driver.implicitly_wait(10)
                hotel_list_1=driver.execute_script("price_list={}; $('.availableRatesContent').each(function(index){  var room=$(this).find('.availableRoom').find('h2').text(); var price=parseFloat($(this).find('.availableTotal').text().replace('USD','').replace('$','')); price_list[''+room+'']=price; }); return price_list;")
            except:
                hotel_list_1={}
            looped_hotel_data_1[indate.replace('%2F','/')]=hotel_list_1
            print 'laquinta iter_'+str(i+1)

            """ hotel hilton """

            driver.implicitly_wait(10)
            try:
                hilton_url="http://hamptoninn3.hilton.com/en/hotels/texas/hampton-inn-and-suites-new-braunfels-NBFELHX/index.html"
                driver.get(hilton_url)
                if driver.current_url!=hilton_url:
                    driver.get(hilton_url)
                month=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
                str_indate = str(dt.day).zfill(2)+" "+str(month[int(dt.month-1)])+" "+str(dt.year)
                str_outdate = str(ndt.day).zfill(2)+" "+str(month[int(ndt.month-1)])+" "+str(ndt.year)
                driver.execute_script('$("#checkin").val("'+str_indate+'");')
                driver.execute_script('$("#checkout").val("'+str_outdate+'");')
                driver.find_element_by_id("check_availability").click()
                handle=driver.window_handles
                driver.switch_to.window(handle[0])
                hotel_data_2 = driver.execute_script("return document.getElementsByClassName('fsRoom')")
                driver.implicitly_wait(10)
                hotel_list_2 = {}

                for room in hotel_data_2:
                    if room.value_of_css_property('display') == "block":
                        hotel_name =room.find_element_by_tag_name('h2').text
                        try:
                            hotel_list_2[hotel_name+' **EASY CANCELLATION']= float(room.find_elements_by_class_name('currencyCode-USD')[0].text.replace('$',''))
                            hotel_list_2[hotel_name+' **2X Points Package Special offer']= float(room.find_elements_by_class_name('currencyCode-USD')[1].text.replace('$',''))
                        except:
                            pass
            except:
                hotel_list_2 = {}
            looped_hotel_data_2[indate.replace('%2F','/')]=hotel_list_2

            print 'hilton iter_'+str(i+1)

            """ hotel marriott """
            try:
                marriott_url='https://www.marriott.com/reservation/availability.mi?isSearch=true&propertyCode=satbf&fromDate='+indate.replace('%2F','/')+'&toDate='+outdate.replace('%2F','/')+'&numberOfRooms=1&numberOfGuests=1'
                if driver.current_url!=marriott_url:
                    driver.get(marriott_url)
                driver.get(marriott_url)
                driver.implicitly_wait(5)
                driver.execute_script("$('.m-button-default').trigger('click');")
                driver.implicitly_wait(10)
                hotel_list_3=driver.execute_script("var hotel_list={}; $('.rph-row.l-rph-row').each(function(index,el){hotel_list[$(el).find('h3').text().trim()]=parseFloat($(el).find('.l-rate-display.rate-display.m-pricing-block').find('.t-price').text().trim())}); return hotel_list;")
            except:
                hotel_list_3={}
            looped_hotel_data_3[indate.replace('%2F','/')]=hotel_list_3
            
            print 'marriott iter_'+str(i+1)

            """ Win Gates """

            driver.implicitly_wait(10)
            try:
                wingate_url='http://www.wingatehotels.com/hotels/texas/new-braunfels/wingate-by-wyndham-new-braunfels/rooms-rates?srcDestination=&partner_id=&hotel_id=23976&srcBrand=&group_code=&campaign_code=&compare=false&propId=WG23976&checkout_date='+outdate+'&brand_id=WG&children=0&useWRPoints=false&ratePlan=BAR&teens=0&affiliate_id=&brand_code=BH%2CDI%2CRA%2CBU%2CHJ%2CKG%2CMT%2CSE%2CTL%2CWG%2CWY%2CPX%2CWT%2CWP%2CPN&iata=&childAgeParam=&adults=1&checkin_date='+indate+'&rooms=1'
                driver.get(wingate_url)
                if driver.current_url!=wingate_url:
                    driver.get(wingate_url)
                driver.implicitly_wait(10)
                hotel_list_4=driver.execute_script("var hotel_list={}; $('#tabs-0 ol.room_results li .room_info').each(function(index){  var room=$(this).find('h3').text(); var price=parseFloat($(this).find('tr.total td').find('span.actual_price').find('span.price_amt').text().trim()); if (price){hotel_list [''+room+'']=price;} }); return hotel_list;")
            except:
                hotel_list_4={}
            looped_hotel_data_4[indate.replace('%2F','/')]=hotel_list_4

            print 'wingates iter_'+str(i+1)

        """filing data"""

        header_hotel_data={}
        header_hotel_data['La Quinta Inn & Suites New Braunfels']=looped_hotel_data_1
        header_hotel_data['Hampton Inn & Suites New Braunfels']=looped_hotel_data_2
        header_hotel_data['Fairfield Inn & Suites New Braunfels']=looped_hotel_data_3
        header_hotel_data['Wingate by Wyndham New Braunfels']=looped_hotel_data_4
        heading=['Hotel Name','Check In Date','Room Type','Price Per Night']

        dt = today
        capture_dt=str(dt.month).zfill(2)+"/"+str(dt.day).zfill(2)+"/"+str(dt.year)
        time=str((dt-datetime.timedelta(hours=6)).time()).split(':')
        file_name=str(dt.year)+str(dt.month).zfill(2)+str(dt.day).zfill(2)+"_"+time[0]+time[1]+'.xls'
        wb=xlwt.Workbook()
        sheet=wb.add_sheet('worksheet')
        style=xlwt.XFStyle()
        font=xlwt.Font()
        font.name='Times New Roman'
        font.bold=True
        font.weight='700'
        style.font=font
        row=0
        for column,heading in enumerate(heading):
            sheet.write(row,column,heading,style=style)
        font.bold=False
        font.weight='400'
        row=row+1

        for hotel_name,hotel_data in header_hotel_data.items():
            sheet.write(row,0,hotel_name,style)
            date_list=sorted(hotel_data.items(), key=operator.itemgetter(0))
            for date_record in date_list:
                indate=date_record[0]
                data=date_record[1]
                data_hotel=sorted(data.items(), key=operator.itemgetter(1))
                sheet.write(row,1,indate,style)
                for record in data_hotel:
                    for column,value in enumerate(record):
                        if column:
                            sheet.write(row,column+2,'$'+str('%.2f'%value),style)
                        else:
                            sheet.write(row,column+2,value,style)
                    row=row+1
                if not len(data):
                    row=row+1
                row=row+1
            row=row+1
        wb.save(file_name)

        print "file created sucessfully"

        """mailing file"""

        from_email = settings.EMAIL_HOST_USER
        email_list = ['sivadhanamjay@aptuz.com']
        subject = 'Room Rates Captured On '+capture_dt
        message = 'Room Rates Attached'
        msg=EmailMessage(subject,message,to=email_list,from_email=from_email)
        msg.content_subtype='html'
        msg.attach_file(file_name)
        msg.send()
        print "mail sent sucessfully"
        if os.path.isfile(file_name):
            os.system("rm "+file_name)
            print "file removed sucessfully"

    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException, e: return False
        return True

    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException, e: return False
        return True

    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True

    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()
