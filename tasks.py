from robocorp.tasks import task
from robocorp import browser
from RPA.Tables import Tables
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from PIL import Image
import time
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    # browser.configure(
    #     slowmo=1000,
    # )
    open_robot_order_website()
    orders = get_orders()
    close_annoying_modal()
    for order in orders:
        # print(order)
        # click yes on popup
        fill_the_form(order)
    create_zip_file_of_all_receipts()

def fill_the_form(order):
    """Fills in the sales data and click the 'Submit' button"""
    # time.sleep(2)
    page = browser.page()
    # time.sleep(2)
    page.wait_for_selector('#head')
    page.select_option('#head', order["Order number"])
    # page.fill("#body", str(order["Body"]))
    body_id = "#id-body-"+str(order["Body"])
    page.click(body_id)    
    number_input = page.query_selector('input[type="number"]')
    number_input.fill(str(order["Legs"]))
    page.fill("#address", str(order["Address"]))
    page.click("button:text('Order')") 
    # time.sleep(2)

    ## store pdf
    store_receipt_as_pdf(order["Order number"])
    time.sleep(8)
    page.wait_for_selector('#order-another')
    page.click('#order-another')
    # time.sleep(2)
    # try:
    page.wait_for_selector("button:text('Yep')")
    page.click("button:text('Yep')") 
    # except:
    #     pass


def store_receipt_as_pdf(order_number):
    # take screenshot of the page
    screenshot = screenshot_robot(order_number)
    # time.sleep(2)
    pdf = PDF() 
    if screenshot:
        embed_screenshot_to_receipt(screenshot,pdf,order_number)

def screenshot_robot(order_number):
    page = browser.page()
    image_name = "output/"+str(order_number)+".png"
    page.screenshot(path=image_name)
    time.sleep(2)
    return image_name

def embed_screenshot_to_receipt(screenshot, pdf_file,order_number):
    image = Image.open(screenshot)
    pdf_name = "output/"+str(order_number)+".pdf"
    pdf = image.convert('RGB').save(pdf_name)
    time.sleep(2)
    return pdf

def close_annoying_modal():
    page = browser.page()
    page.click("a:text('Order your robot!')") 
    page.click("button:text('Yep')") 

def get_orders():
    # TODO: Implement your function here
    """Downloads excel file from the given URL"""
    http = HTTP()
    library = Tables()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    orders = library.read_table_from_csv(
            "orders.csv" 
        )
    return orders

def open_robot_order_website():
    # TODO: Implement your function here
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")
    log_in()


def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")    

def create_zip_file_of_all_receipts():
    zip_file_name = os.path.join(out_dir, "all_receipts.zip")
    Archive().archive_folder_with_zip(receipt_dir, zip_file_name)