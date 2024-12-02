from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import zipfile
import os

@task
def order_robots_from_RobotSpareBin():
    """Orders robots from RobotSpareBin Industries Inc. and processes them"""
    browser.configure(slowmo=100)
    open_robot_order_website()
    close_annoying_modal()
    orders_file_path = download_csv_file()
    orders = get_orders(orders_file_path)
    process_orders(orders)
    archive_receipts()

def open_robot_order_website():
    """Opens the RobotSpareBin Industries order website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    """Closes the annoying modal on the website"""
    page = browser.page()
    page.click("button:text('Ok')")

def download_csv_file():
    """Downloads the CSV file containing orders"""
    http = HTTP()
    url = "https://robotsparebinindustries.com/orders.csv"
    file_path = "orders.csv"
    http.download(url=url, target_file=file_path, overwrite=True)
    return file_path

def get_orders(file_path):
    """Reads the downloaded CSV file and returns the orders table"""
    tables = Tables()
    orders_table = tables.read_table_from_csv(file_path)
    return orders_table

def process_orders(orders):
    """Loops through each order, submits the order form, and processes the output."""
    for order in orders:
        print(f"Processing order: {order}")
        fill_and_submit_form(order)
        pdf_file = store_receipt_as_pdf(order['Order number'])
        screenshot_path = screenshot_robot(order['Order number'])
        embed_screenshot_to_receipt(screenshot_path, pdf_file)

def fill_and_submit_form(order):
    """Fills in the robot order form and submits it"""
    page = browser.page()

    body_options = {
        "1": "Roll-a-thor body",
        "2": "Peanut crusher body",
        "3": "D.A.V.E body",
        "4": "Andy Roid body",
        "5": "Spanner mate body",
        "6": "Drillbit 2000 body"
    }

    if 'Head' in order and order['Head']:
        page.select_option("#head", order['Head'])
    else:
        print(f"Head missing for order {order['Order number']}")

    if 'Body' in order and order['Body']:
        body_value = order['Body']
        body_name = body_options.get(body_value)
        if body_name:
            page.click(f'input[id="id-body-{body_value}"]')
        else:
            print(f"Invalid Body option for order {order['Order number']}")
    else:
        print(f"Body missing for order {order['Order number']}")

    if 'Legs' in order and order['Legs']:
        page.fill('input[type="number"][min="1"][max="6"]', str(order['Legs']))
    else:
        print(f"Legs missing for order {order['Order number']}")

    if 'Address' in order and order['Address']:
        page.fill("#address", order['Address'])
    else:
        print(f"Address missing for order {order['Order number']}")

    page.click("text=Preview")
    
    page.click("text=ORDER")
    
    print(f"Order {order['Order number']} has been submitted.")

def store_receipt_as_pdf(order_number):
    """Stores the order receipt as a PDF file"""
    page = browser.page()
    receipt_html = page.locator("#order").inner_html()
    output_directory = os.path.join(os.getcwd(), "output", "receipts")
    os.makedirs(output_directory, exist_ok=True)
    pdf_file = os.path.join(output_directory, f"order_{order_number}_receipt.pdf")
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, pdf_file)
    print(f"Receipt for order {order_number} saved as PDF: {pdf_file}")
    return pdf_file

def screenshot_robot(order_number):
    """Takes a screenshot of the robot and saves it to a file."""
    page = browser.page()
    screenshot_directory = os.path.join(os.getcwd(), "output", "screenshots")
    os.makedirs(screenshot_directory, exist_ok=True)
    
    screenshot_file = os.path.join(screenshot_directory, f"robot_{order_number}.png")
    page.screenshot(path=screenshot_file)
    
    print(f"Screenshot for order {order_number} saved at: {screenshot_file}")
    return screenshot_file

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the robot screenshot into the receipt PDF file."""
    pdf = PDF()

    files_to_add = [
        pdf_file,                    
        f"{screenshot}:align=center"
    ]

    pdf.add_files_to_pdf(
        files=files_to_add,
        target_document=pdf_file
    )

    print(f"Screenshot {screenshot} embedded into receipt {pdf_file}")
    return pdf_file

def archive_receipts(output_directory="output/receipts", archive_name="output/receipts.zip"):
    """Creates a ZIP archive of all receipt PDF files in the specified directory."""
    lib = Archive()

    if not os.path.exists(output_directory):
        print(f"Directory {output_directory} does not exist!")
        return
    
    lib.archive_folder_with_zip(output_directory, archive_name, recursive=True)

    print(f"Archive created successfully: {archive_name}")
    files = lib.list_archive(archive_name)
    for file in files:
        print(file)

