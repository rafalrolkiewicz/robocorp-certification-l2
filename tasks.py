from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    download_orders()
    open_order_website()
    orders = get_orders()

    for order in orders:
        close_annoying_modal()
        fill_the_form(order)
        store_receipt_as_pdf(order["Order number"])

    archive_receipts()


def download_orders():
    """Downloads csv file with orders from given URL"""
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def open_order_website():
    """Opens the order website"""
    browser.goto(url="https://robotsparebinindustries.com/#/robot-order")


def get_orders():
    """Gets the order data from csv, returns list of dicts"""
    table = Tables()
    orders = table.read_table_from_csv("orders.csv")
    return orders


def close_annoying_modal():
    """Closes the pop-up window"""
    page = browser.page()
    page.click("button:text('OK')")


def fill_the_form(order):
    """Fills the order form and clicks order button"""
    page = browser.page()
    page.select_option("#head", order["Head"])
    page.click(f'#id-body-{order["Body"]}')
    page.get_by_placeholder(
        "Enter the part number for the legs").fill(order["Legs"])
    page.fill("#address", order["Address"])
    # Keep clicking Order button until receipt is loaded
    while (page.query_selector("#receipt") is None):
        page.click("button:text('Order')")


def store_receipt_as_pdf(order_number):
    """Stores the order receipt as a PDF file"""
    page = browser.page()
    pdf = PDF()
    lib = FileSystem()
    lib.create_directory("output/receipts")
    pdf_path = f"output/receipts/order_{order_number}.pdf"
    page.locator("#receipt").wait_for(state='attached')
    order_receipt_html = page.locator("#receipt").inner_html(timeout=10)
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    screenshot_robot(order_number)
    embed_screenshot_to_receipt(
        f"output/receipts/{order_number}.png", pdf_path)
    page.click("button:text('Order another robot')")


def screenshot_robot(order_number):
    """Takes a screenshot of the robot"""
    page = browser.page()
    page.locator(selector="#robot-preview-image").screenshot(
        path=f"output/receipts/{order_number}.png")


def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Adds the robot screenshot to the receipt PDF file"""
    pdf = PDF()
    pdf.open_pdf(pdf_file)
    pdf.add_files_to_pdf(
        files=[screenshot], target_document=pdf_file, append=True)
    pdf.save_pdf(output_path=pdf_file)
    pdf.close_all_pdfs()


def archive_receipts():
    """Archives pdfs of receipts into a zip file"""
    lib = Archive()
    lib.archive_folder_with_zip(
        "output/receipts", "output/receipts.zip", include='*.pdf')
