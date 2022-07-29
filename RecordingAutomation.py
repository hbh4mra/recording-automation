import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class Automation:

    def __init__(self, set_timer=False, timer_duration=0):

        self.__row_count = 0
        self.__set_timer = set_timer
        self.__timer_duration = timer_duration
        self.__processed_recordings = list()
        self.__faulty_recordings = list()       # List containing Data To Be Mailed
        self.driver = webdriver.Chrome()

    def fetch_web_page(self):
        # The address of your page
        self.driver.get("http://10.10.162.5:8080/servlet/cs")
        # For max window
        self.driver.maximize_window()

        # Enters in user details and submits
        input_element = self.driver.find_element_by_id("username")
        input_element.send_keys('USERNAME')
        input_element = self.driver.find_element_by_id("j_password")
        input_element.send_keys('PASSWORD')
        input_element.send_keys(Keys.ENTER)

        # Gives an implicit wait for 'x' seconds
        self.driver.implicitly_wait(15)
        try:
            # Clicks on Trading Recordings using class tags
            nav = self.driver.find_elements_by_class_name('navNotSelected1')
            if len(nav) > 0:
                nav[0].click()
        except NoSuchElementException:
            return

    def refresh_content(self):
        try:
            search = self.driver.find_elements_by_id('searchButtonBottom')
            if len(search) > 0:
                search[0].click()
                time.sleep(3)
                alert = self.driver.switch_to.alert
                alert.accept()
                self.driver.implicitly_wait(10)
        except NoSuchElementException:
            return
        except NoAlertPresentException:
            return
        except UnexpectedAlertPresentException:
            alert = self.driver.switch_to.alert
            alert.dismiss()

    def show_all_content(self):
        try:
            # Clicks on the Show All link
            show = self.driver.find_elements_by_link_text("Show All")
            if len(show) > 0:
                show[0].click()
                self.driver.implicitly_wait(5)
        except NoSuchElementException:
            return

    def scroll_page(self):
        # Scrolls down the page
        self.driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
        self.driver.implicitly_wait(5)

    def process_recordings(self):
        try:
            # Fetch the table that contains recording
            table_container = self.driver.find_element_by_id("tableContainer").find_element_by_tag_name('table')

            # Fetching all the content rows from the table
            table_rows = table_container.find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')

            if self.__row_count == 0:
                rows_to_process = 0
                first_run = True
            else:
                rows_to_process = 1
                first_run = False

            if len(table_rows) > self.__row_count:
                rows_to_process = rows_to_process + (len(table_rows) - self.__row_count)
                self.__row_count = len(table_rows)

            row_num_list = list(range(rows_to_process))

            try:
                # Loop through all rows to access the data in each row.
                for row_num in row_num_list:
                    columns = table_rows[row_num].find_elements_by_tag_name('td')
                    col = 0

                    try:
                        date_time = " ".join(dt.strip() for dt in columns[3].text.strip().split())
                        device_name = columns[7].text.strip()
                    except IndexError:
                        continue

                    if self.__is_row_processed(date_time, device_name):
                        continue

                    if row_num == rows_to_process - 1 and not first_run:
                        row_num_list.append(row_num + 1)
                    # Loop through each column in a row to access column data.
                    for data in columns:
                        try:
                            if col == 0:
                                col += 1
                                data.find_element_by_tag_name('input').click()
                                time.sleep(25)
                                alert = self.driver.switch_to.alert
                                alert.dismiss()
                                if self.__is_faulty_row(date_time, device_name):
                                    break
                                else:
                                    self.__faulty_recordings.append([date_time])
                                continue
                            else:
                                # Collect Data For Mail
                                if 3 < col < 12:
                                    self.__faulty_recordings[-1].append(data.text.strip())
                                if col == 12:
                                    self.run_mail("individual")
                                col += 1
                        except NoAlertPresentException:
                            self.__processed_recordings.append([date_time, device_name])
                            break
                        except UnexpectedAlertPresentException as e1:
                            print(type(e1).__name__ + " : " + e1.msg)
                            try:
                                alert = self.driver.switch_to.alert
                                alert.dismiss()
                                if self.__is_faulty_row(date_time, device_name):
                                    break
                                else:
                                    self.__faulty_recordings.append([date_time])
                                continue
                            except NoAlertPresentException as ex:
                                print(type(ex).__name__ + " : " + ex.msg)
                                break
            except IndexError:
                print("Data Processed for this Cycle. Moving to Next Cycle After Refreshing.")
                return
        except NoSuchElementException as e0:
            print("No Recording To Load. Moving to Next Cycle After Refreshing.")
            print(type(e0).__name__ + " : " + e0.msg)
            return

    def run_mail(self, mail_type):
        try:
            from_email = 'FROM_EMAIL'
            username = 'USERNAME'
            password = 'PASSWORD'
            send_to_email = 'TO_EMAIL'

            header = "<span style='width:20ch; display: inline-table'>%s</span>" \
                     "<span style='width:14ch; display: inline-table'>%s</span>" \
                     "<span style='width:9ch; display: inline-table'>%s</span>" \
                     "<span style='width:14ch; display: inline-table'>%s</span>" \
                     "<span style='width:15ch; display: inline-table'>%s</span>" \
                     "<span style='width:16ch; display: inline-table'>%s</span>" \
                     "<span style='width:26ch; display: inline-table'>%s</span>" \
                     "<span style='width:10ch; display: inline-table'>%s</span>" \
                     "<span style='width:10ch; display: inline-table'>%s</span><br />" % \
                     ("CALL START", "CALL LENGTH", "TRADER", "TRADER NAME", "DEVICE", "DIALED NUMBER", "CALLER NAME",
                      "CPN", "DIR")

            if mail_type == "individual":
                subject = '[INDIVIDUAL RECORDING ERROR]'  # The subject line
                error_details = self.__faulty_recordings[-1]
                message = '<p style="font-family: Courier">Complete Detail of Single Recording Error<br />'
                separator = "-" * 45
                message += "%s<br />" % separator
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("CALL START : ", error_details[0])
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("CALL LENGTH : ", error_details[1])
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("TRADER : ", error_details[2])
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("TRADER NAME : ", error_details[3])
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("DEVICE : ", error_details[4])
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("DIALED NUMBER : ", error_details[5])
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("CALLER NUMBER : ", error_details[6])
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("CPN : ", error_details[7])
                message += "<span style='width:18ch; display: inline-table'>%s</span><span>%s</span><br />" % \
                           ("DIR : ", error_details[8])
                message += "%s<br /></p>" % separator
            else:
                subject = '[DAILY RECORDING ERROR REPORT]'  # The subject line
                message = '<p style="font-family: Courier">Daily Report of Recording Errors.<br />'
                message += '%s%s<br /><br />' % ('Total Errors Caught : ', len(self.__faulty_recordings))
                separator = "-" * 136
                message += "%s<br />" % separator
                message += header
                message += "%s<br />" % separator
                for error_details in self.__faulty_recordings:
                    message += "<span style='width:20ch; display: inline-table'>%s</span>" \
                               "<span style='width:14ch; display: inline-table'>%s</span>" \
                               "<span style='width:9ch; display: inline-table'>%s</span>" \
                               "<span style='width:14ch; display: inline-table'>%s</span>" \
                               "<span style='width:15ch; display: inline-table'>%s</span>" \
                               "<span style='width:16ch; display: inline-table'>%s</span>" \
                               "<span style='width:26ch; display: inline-table'>%s</span>" \
                               "<span style='width:10ch; display: inline-table'>%s</span>" \
                               "<span style='width:10ch; display: inline-table'>%s</span><br />" % \
                               (error_details[0][0:19], error_details[1][0:13], error_details[2][0:8],
                                error_details[3][0:13], error_details[4][0:14], error_details[5][0:15],
                                error_details[6][0:25], error_details[7][0:9], error_details[8][0:9])
                message += "%s<br /></p>" % separator

            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = send_to_email
            msg['Subject'] = subject

            # Attach the message to the MIMEMultipart object
            msg.attach(MIMEText(message, 'html'))

            server = smtplib.SMTP('SMTP_DETAILS', PORT)
            server.starttls()
            server.login(username, password)
            text = msg.as_string()  # Converts the MIMEMultipart object to a string to send
            server.sendmail(from_email, send_to_email, text)
            server.quit()
        except Exception as e2:
            print(e2)

    def __is_row_processed(self, date_time, device_name):
        for processed_data in self.__processed_recordings:
            if date_time in processed_data and device_name in processed_data:
                return True
        return False

    def __is_faulty_row(self, date_time, device_name):
        for faulty_row in self.__faulty_recordings:
            if date_time in faulty_row and device_name in faulty_row:
                return True
        return False

    def get_unprocessed_list(self):
        return self.__faulty_recordings

    def get_processed_list(self):
        return self.__processed_recordings

    def get_total_recordings_checked(self):
        return self.__row_count

    def set_timer_duration(self, timer_duration):
        self.__timer_duration = timer_duration

    def stop_timer(self):
        self.__timer_duration = 0
        self.__set_timer = False

    def start_timer(self, timer):
        self.__timer_duration = timer
        self.__set_timer = True

    def run_automation(self):
        try:
            self.fetch_web_page()
            while True:
                self.refresh_content()
                self.show_all_content()
                self.scroll_page()
                self.process_recordings()
                if self.__set_timer:
                    time.sleep(self.__timer_duration)
        except KeyboardInterrupt:
            print("Forcefully Ending Automation Run.")
            self.driver.quit()
            return
        except TimeoutException as e3:
            print(type(e3).__name__ + " : " + e3.msg)
            self.driver.quit()
            return
        except NoSuchElementException as e4:
            print(type(e4).__name__ + " : " + e4.msg)
            self.driver.quit()
            return

    def __del__(self):
        self.driver.quit()


if __name__ == "__main__":
    try:
        recording_automation_object = Automation(True, 1)
        recording_automation_object.run_automation()
        recording_automation_object.run_mail("daily")
    except (TypeError, KeyboardInterrupt) as e:
        print("Automation Program Ended")
