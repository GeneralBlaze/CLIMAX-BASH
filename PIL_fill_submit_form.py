import json
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class TestGetTAGS():
    
    file_path = '/Users/princewill/Desktop/aTDO PIL/'
    
    def setup_method(self, method):
        self.driver = webdriver.Chrome()
        self.driver.set_window_size(1157, 778)
        self.vars = {}

    def teardown_method(self, method):
        self.driver.quit()

    def test_getTAGS(self):
        # Read data from JSON
        with open('/Users/princewill/alx-interview/CLIMAX BASH/bl.json', mode='r') as file:
            bl_numbers = json.load(file)
            
            for bl_number in bl_numbers:
                file1 = os.path.join(self.file_path, f"{bl_number}.pdf")

                # Open the form page
                self.driver.get("https://ecommerce.pilnigeria.com/refund-request")
                
                # Shortened variable names for field IDs
                submit_form_to = "edit-submitted-submit-form-to-1"
                name_of_consignee = "edit-submitted-personal-data-name-of-consignee-company-individual"
                registered_address = "edit-submitted-personal-data-company-individual-registered-address"
                rc_no = "edit-submitted-personal-data-registration-certificate-rc-no"
                date_of_incorporation = "edit-submitted-personal-data-date-of-incorporation"
                email_ref = "edit-submitted-personal-data-email-ref"
                office_phone_no = "edit-submitted-personal-data-office-phone-no"
                mobile_no = "edit-submitted-personal-data-mobile-no"
                bl_no_con = "edit-submitted-consignee-bank-details-bl-no-con"
                account_name_con = "edit-submitted-consignee-bank-details-account-name-con"
                account_number_con = "edit-submitted-consignee-bank-details-account-number-nuban-only-con"
                sort_code_con = "edit-submitted-consignee-bank-details-sort-code-con"
                bank_address = "edit-submitted-consignee-bank-details-bank-address"
                upload_doc1 = "edit-submitted-consignee-bank-details-upload-relevant-documents-fieldset-upload-document-con-upload-file-con-upload"
                description_of_document = "edit-submitted-consignee-bank-details-upload-relevant-documents-fieldset-upload-document-con-description-of-document-con"
                
                # Dictionary to store form field IDs and their corresponding values
                form_data = {
                    submit_form_to: None,  # This is a button, so no value needed
                    name_of_consignee: "ORIENT LOGISTICS ENTERPRISES",
                    registered_address: "5/7 BEN OYEKA STREET,OLODI APAPA, LAGOS,AJEROMI-IFELODUN, LAGOS",
                    rc_no: "BN3053987",
                    date_of_incorporation: "-",
                    email_ref: "DONCLIMAX22@YAHOO.COM",
                    office_phone_no: "08069916073",
                    mobile_no: "08069916073",
                    bl_no_con: bl_number,  # Dynamic BL number
                    account_name_con: "ORIENT LOGISTICS ENTERPRISES",
                    account_number_con: "1023649496",
                    sort_code_con: "033153283",
                    bank_address: "LAGOS NIGERIA",
                    description_of_document: f"Terminal Delivery Order for {bl_number}"  # Dynamic BL number,
                }

                # Fill in the form fields using the dictionary
                for field_id, value in form_data.items():
                    element = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, field_id))
                    )
                    if value is not None:
                        element.send_keys(value)
                    else:
                        element.click()
                
                # Dictionary to store dropdown field IDs and their corresponding values
                dropdown_data = {
                    "edit-submitted-personal-data-country-of-incorporation": "Nigeria",
                    "edit-submitted-consignee-bank-details-port-of-discharge": "ONNE",
                    "edit-submitted-consignee-bank-details-bank-name-con": "United Bank for Africa (UBA)"
                }

                # Handle dropdowns using the dictionary
                for field_id, value in dropdown_data.items():
                    dropdown = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, field_id))
                    )
                    dropdown.click()
                    option = dropdown.find_element(By.XPATH, f"//option[. = '{value}']")
                    option.click()
                    
                # Click the body element (empty part of the screen)
                self.driver.find_element(By.TAG_NAME, "body").click()

                # Upload the files
                WebDriverWait(self.driver, 1300).until(
                    EC.presence_of_element_located((By.ID, upload_doc1))
                ).send_keys(file1)
                

                # Submit the form
                #self.driver.find_element(By.NAME, "op").click()

                # Wait for the form to submit and return, or add time.sleep if necessary
                # self.driver.implicitly_wait(5)  # or time.sleep(5)

# To run the test
if __name__ == "__main__":
    test = TestGetTAGS()
    test.setup_method(None)
    test.test_getTAGS()
    time.sleep(10000)
    # Comment out the teardown_method call to keep the browser open
    # test.teardown_method(None)