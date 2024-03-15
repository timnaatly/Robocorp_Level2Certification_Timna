from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Desktop import Desktop
import time
import os
desktop = Desktop()
@task
def order_robots_from_RobotSparBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    message = "before open"

    browser.configure(
        slowmo=100
    )
    open_the_website()
    log_in()
    download_csv_file()
    close_browser()
    receipts_zipfile()


def open_the_website():

    page = browser.goto("https://robotsparebinindustries.com/#/robot-order")

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    page = browser.page()
    page.click("button:text('OK')")

def create_order_with_inputdata(order_data):
    page = browser.page()
    page.select_option('//*[@id="head"]', order_data["Head"])
    num = order_data["Body"]
    page.click(f'//*[@id="id-body-{num}"]')
    page.fill('//*[@placeholder="Enter the part number for the legs"]', order_data["Legs"])
    page.fill('//*[@id="address"]', order_data["Address"])


# Preview the robot

    page.click('//*[@id="preview"]')
    time.sleep(2)
# Submit the order

    page.click('//*[@id="order"]')
    time.sleep(2)
# Take screenshot of the robot
    ScreenshotPath=r"output/Screenshot/"+order_data["Order number"]+".png"
    #page.screenshot(path=ScreenshotPath)
    page.locator('//*[@id="robot-preview-image"]').screenshot(path=ScreenshotPath)

    export_as_PDF(order_data)

#Go to order another request
    page.click('//*[@id="order-another"]')
    log_in()

def download_csv_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", target_file="input/orders.csv", overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv("input/orders.csv")

    for row in orders:
        create_order_with_inputdata(row)

def export_as_PDF(order_data):
    page = browser.page()
    order_receipt_html = page.locator('//*[@id="receipt"]').inner_html()
   

    pdf = PDF()
    ScreenshotPath="output/Screenshot/"+order_data["Order number"]+".png"
    PDFPath="output/PDFReceipts/"+order_data["Order number"]+".pdf"
    pdf.html_to_pdf(order_receipt_html, PDFPath)

    pdf.open_pdf(PDFPath)
    
    list_of_files = [ScreenshotPath]
    pdf.add_files_to_pdf(files=list_of_files, target_document=PDFPath, append=True)

def receipts_zipfile():
    lib = Archive()
    PDFFolderPath= "output/PDFReceipts/"
    lib.archive_folder_with_zip(PDFFolderPath, 'output/pdf_archive.zip')

def close_browser():
    page = browser.page()
    page.close()
       





