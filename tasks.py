from robocorp.tasks import task
from robocorp import browser

#from.robot import seleniumlibrary
#from selenium import webdriver

from RPA.Archive import Archive
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF

import os
from zipfile import ZipFile
import re
import time
import pickle
from pathlib import Path



@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    browser.configure(
        slowmo=500,
    )
    open_robot_order_website()
    close_annoying_modal()
    download_csv_file()
    fill_the_form_with_csv_data()
    archive_receipts()
    
def archive_receipts():
    """Archives all the receipt pdfs into a single zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts/", "./output/receipts.zip")

def close_annoying_modal():
    """Closes the annoying modal window"""
    page = browser.page()
    page.click("button:text('Yep')")

def download_csv_file():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    orders = Tables().read_table_from_csv("orders.csv", header=True)
    print(orders)

def fill_and_submit_order_form(order_rep):
    """Fills in the order data and click the 'Submit' button"""
    page = browser.page()
    page.select_option("#head", str(order_rep["Head"]))
    page.click("#id-body-1")
    body_locator = "#id-body-%s" % (order_rep["Body"])
    page.click(body_locator)
    page.fill('input[placeholder="Enter the part number for the legs"]', str(order_rep["Legs"]))
    page.fill("#address", str(order_rep["Address"]))
    retry_order_until_successful()

    time.sleep (1)
    
    store_receipt_as_pdf()
    time.sleep (3)
    
    page.click("#order-another")
    close_annoying_modal()


def fill_the_form_with_csv_data():
    """Read the data from csv and fill in order form"""
    library = Tables()
    orders = library.read_table_from_csv("orders.csv", columns=["Order number","Head", "Body", "Legs", "Address"])

    for row in orders:
        fill_and_submit_order_form(row)

def retry_order_until_successful():
    """search order recepien"""
    page = browser.page()
    for i in range(100):
        page.click("#order")
        order_completed = page.query_selector("#order-completion")
        if order_completed:
            return

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def store_receipt_as_pdf():
    """Export the order data to a pdf file"""
    page = browser.page()
    take_robot_screenshot()
    sales_results_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/receipts/receipt.pdf")
    text = pdf.get_text_from_pdf("output/receipts/receipt.pdf")
    file = open('output/filename.txt', 'wb')
    pickle.dump(text, file)
    file.close()
    pattern = "RSB-ROBO-ORDER"
    text_file = open("output/filename.txt", encoding='cp1252').readlines()

    for line in text_file:
        if re.search(pattern, line):
            #print(line[0:25])
            filename = line[0:25]
            pdf.html_to_pdf(sales_results_html, "output/receipts/Receipt-%s.pdf" %(filename))
            page.locator("#robot-preview-image").screenshot(path="output/receipts/robot_picture.png")
            pdf.add_watermark_image_to_pdf(image_path="output/receipts/robot_picture.png", source_path="output/receipts/Receipt-%s.pdf" %(filename), output_path="output/receipts/Receipt-%s.pdf" %(filename))                 
            os.remove("output/receipts/receipt.pdf")
            os.remove("output/filename.txt")
            os.remove("output/receipts/robot_picture.png")                
        else:
            print("Order number not found")

def take_robot_screenshot():
    """Take a screenshot of the page, tätä ei varsinaisesti käytetä missää. Mutta tekee robotista screenshotin logiin"""
    page = browser.page()
    locator = page.locator("#robot-preview-image")
    browser.screenshot(locator),


 

