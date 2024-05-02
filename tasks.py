from robocorp.tasks import task
from robocorp import browser
from RPA.Tables import Tables
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive
import shutil



@task
def orders_robots_from_RobotSpareBin():
    """Orders robots from RobotSpareBin Industries inc.
    Saves the order html receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the Screenshot of the robot th the PDF receipt.
    Creates ZIP archive of the receipts and the images."""
    browser.configure(slowmo=100)

    open_robot_order_website()
    open_CSV_and_create_table()
    download_CSV_file()
    move_to_zip()
    clean_up()


def open_robot_order_website():
    """Open webpage to begin automation"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()

def close_annoying_modal():
    """Click OK to close message"""
    page = browser.page()

    page.click("button:text('OK')")

def download_CSV_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def fill_and_submits_order(order):
    """Fills in form and selects order"""
    page = browser.page()

    head_names = {
        "1" : "Roll-a-thor head",
        "2" : "Peanut crusher head",
        "3" : "D.A.V.E head",
        "4" : "Andy Roid head",
        "5" : "Spanner mate head",
        "6" : "Drillbit 2000 head"
    }
    head_number = order["Head"]
    page.select_option("#head", head_names.get(head_number))  
    page.click('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{0}]/label'.format(order["Body"]))  
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])
    page.fill("#address", order["Address"])
    while True:
        page.click("#order")
        order_another = page.query_selector("#order-another")
        if order_another:
            pdf_path = store_PDF_receipt(int(order["Order number"]))
            screenshot_path = screenshot_robot(int(order["Order number"]))
            embed_screenshot_to_receipt(screenshot_path, pdf_path)
            order_another_robot()
            close_annoying_modal()
            break
        
    

def open_CSV_and_create_table():
    """Create table from download"""
    table = Tables()
    orders = table.read_table_from_csv("orders.csv")

    for rows in orders:
        fill_and_submits_order(rows)
        
def order_another_robot():
    """Clickes on order another robot button"""
    page = browser.page()
    page.click("#order-another")

def store_PDF_receipt(order_number):
    """Creates file of robot order receipt as PDF"""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = "output/receipts/{0}.pdf".format(order_number)
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    """Saves screenshot of robot picture"""
    page = browser.page()
    screenshot_path = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    """embeds the screenshot to the receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path, source_path=pdf_path, output_path=pdf_path)

def move_to_zip():
    """create archive of all receipts in a zip file"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")

def clean_up():
    """remove old files"""
    shutil.rmtree("./output/receipts")
    shutil.rmtree("./output/screenshots")
