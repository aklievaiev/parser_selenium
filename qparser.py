# Loading libraries
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import xlsxwriter
from configparser import ConfigParser
from selenium.webdriver.support.ui import Select


def main():
    dl_dir = ''
    chrome_options = webdriver.ChromeOptions()
    dl_location = os.path.join(os.getcwd(), dl_dir)
    prefs = {"download.default_directory": dl_location}
    chrome_options.add_experimental_option("prefs", prefs)
    chromedriver = "./chromedriver"
    driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=chrome_options)
    driver.maximize_window()
    driver.get('https://itdashboard.gov/')
    time.sleep(2)
    # # Click on a button 'Dive-In'
    bnt_dive_in = driver.find_element(By.CLASS_NAME, 'trend_sans_oneregular')
    bnt_dive_in.click()
    time.sleep(5)

    # # Collect the spent money on Information Technology
    amount = driver.find_elements(By.CSS_SELECTOR, 'span.h1.w900')
    amount_list = []
    for elem in amount:
        if len(elem.text) != 0:
            amount_list.append(elem.text)

    # # Collect name of agencies
    agencies = driver.find_elements(By.CSS_SELECTOR, 'span.h4.w200')
    agencies_list_title = []
    agencies_list = []
    for elem in agencies:
        if len(elem.text) != 0:
            agencies_list.append(elem.text.lower())
            # For excel
            agencies_list_title.append(elem.text)

    # Create agencies sheet
    workbook = xlsxwriter.Workbook('Agencies.xlsx')
    worksheet = workbook.add_worksheet('Agencies')
    worksheet.write('A1', 'Agency')
    worksheet.write('B1', 'Spent on IT')
    worksheet.write_column('A2', agencies_list_title)
    worksheet.write_column('B2', amount_list)

    # # Collect link of agencies
    link = driver.find_elements(By.LINK_TEXT, 'view')
    href_list = []
    for elem in link:
        href_list.append(elem.get_attribute('href'))
    #
    # # Create dictionary from agencies and their expenses
    agencies_link = dict(zip(agencies_list, href_list))

    # # Config settings
    config = ConfigParser()
    config.read('config.ini')

    agency = config.get('main', 'agency')
    driver.get(agencies_link[agency.lower()])
    time.sleep(15)
    select = Select(driver.find_element(By.CSS_SELECTOR, 'select.c-select'))
    select.select_by_visible_text('All')
    time.sleep(10)

    # Create Individual Investments Table
    rows_odd = driver.find_elements(By.CSS_SELECTOR, "tr.odd")
    rows_even = driver.find_elements(By.CSS_SELECTOR, "tr.even")
    rows = rows_odd + rows_even

    uii_list = []
    bureau_list = []
    investment_title_list = []
    total_spending_list = []
    type_list = []
    cio_rating_list = []
    number_of_project_list = []
    href_list_uii = []

    for row in rows:
        uii = row.find_elements(By.TAG_NAME, "td")[0]

        if uii.find_elements(By.CSS_SELECTOR, "a[href]"):
            href_list_uii.append(uii.find_elements(By.CSS_SELECTOR, "a[href]")[0].get_attribute('href'))
        uii_list.append(uii.text)
        bureau = row.find_elements(By.TAG_NAME, "td")[1]
        bureau_list.append(bureau.text)
        investment_title = row.find_elements(By.TAG_NAME, "td")[2]
        investment_title_list.append(investment_title.text)
        total_spending = row.find_elements(By.TAG_NAME, "td")[3]
        total_spending_list.append(total_spending.text)
        type_amount = row.find_elements(By.TAG_NAME, "td")[4]
        type_list.append(type_amount.text)
        cio_rating = row.find_elements(By.TAG_NAME, "td")[5]
        cio_rating_list.append(cio_rating.text)
        number_of_project = row.find_elements(By.TAG_NAME, "td")[6]
        number_of_project_list.append(number_of_project.text)

    # Create Individual Investment Sheet
    worksheet_invest = workbook.add_worksheet('Individual Investment')
    worksheet_invest.write('A1', 'UII')
    worksheet_invest.write('B1', 'Bureau')
    worksheet_invest.write('C1', 'Investment Title')
    worksheet_invest.write('D1', 'Total FY2021 Spending ($M)')
    worksheet_invest.write('E1', 'Type')
    worksheet_invest.write('F1', 'CIO Rating')
    worksheet_invest.write('G1', '# of Projects')
    worksheet_invest.write_column('A2', uii_list)
    worksheet_invest.write_column('B2', bureau_list)
    worksheet_invest.write_column('C2', investment_title_list)
    worksheet_invest.write_column('D2', total_spending_list)
    worksheet_invest.write_column('E2', type_list)
    worksheet_invest.write_column('F2', cio_rating_list)
    worksheet_invest.write_column('G2', number_of_project_list)
    workbook.close()

    # Open and Download PDF-files
    for elem in href_list_uii:
        driver.get(elem)
        time.sleep(5)
        bnt_dwnld_pdf = driver.find_element(By.LINK_TEXT, 'Download Business Case PDF')
        bnt_dwnld_pdf.click()
        time.sleep(5)

    input()


if __name__ == '__main__':
    main()
