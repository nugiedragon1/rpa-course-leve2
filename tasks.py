from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables 
from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF
from RPA.Archive import Archive
from time import sleep

lib = Selenium()

@task
def order_robots_from_RobotSpareBin():
    """
    Order robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    # download_csv()
    orders = get_orders()

    for row in orders:
        close_annoying_modal()
        # sleep(5)
        fill_the_form(row)
        sleep(3)
        store_receipt_as_pdf(str(row["Order number"]))
    archive_receipts()
def open_robot_order_website():
    try:
     """Open the HTML link"""
     lib.open_browser("https://robotsparebinindustries.com/#/robot-order")
     print("success")
    except:
        print("failed")

def download_csv():
    """Download the CSV"""
    http=HTTP()
    http.download(url='https://robotsparebinindustries.com/orders.csv',overwrite=True)

def get_orders():
    csv= Tables()
    orders= csv.read_table_from_csv("orders.csv", columns=["Order number", "Head","Body","Legs","Address"])
    return orders

def close_annoying_modal():
    """ Closing modal"""
    lib.click_button("OK")

def fill_the_form(row):
    """Filling out forms"""
    lib.select_from_list_by_value('head', str(row["Head"]))

    lib.select_radio_button('body', str(row["Body"]))
    sleep(5)
    lib.find_element('address')
    lib.input_text('address', row["Address"])

    lib.find_element('class:form-control')
    lib.input_text('class:form-control', row["Legs"])

    # locator_preview = ("preview")
    # lib.scroll_element_into_view(locator_preview)
    sleep(5)
    lib.click_button("preview")
    lib.click_button("order")


def screenshot_robot(order_number):
    lib.screenshot(locator="robot-preview-image", filename="output/receipts/"+order_number+".png")

def store_receipt_as_pdf(order_num):
    receipt=lib.get_element_attribute(locator="receipt", attribute="innerHTML")
    pdf=PDF()
    filesystem= FileSystem()
    filesystem.create_directory("output/receipts")
    path="output/receipts/order_"+order_num+".pdf"
    pdf.html_to_pdf(receipt,path)

    screenshot_robot(order_num)
    lib.click_button("order-another")
    sleep(2)
    # lib.click_button("OK")

def archive_receipts():
    lib_archive=Archive()
    lib_archive.archive_folder_with_zip("output/receipts","receipts.zip")