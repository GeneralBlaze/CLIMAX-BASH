import asyncio, json, random, openpyxl
from pyppeteer import launch
from openpyxl import load_workbook

# Helper function to open a new page and navigate to tracking page
accepted_cookies = False

async def open_tracking_page(browser, tracking_number):
    global accepted_cookies
    page = await browser.newPage()
    await page.setViewport({'width': 1168, 'height': 845})

    await page.goto("https://www.msc.com/en/track-a-shipment", {'waitUntil': "networkidle0"})

    # Wait for the cookies popup to appear and click the "Accept" button
    if not accepted_cookies:
        try:
            await page.waitForXPath('//button[normalize-space()="Accept All"]', {'timeout': 30000})
            buttons = await page.xpath('//button[normalize-space()="Accept All"]')
            await buttons[0].click()
            accepted_cookies = True
        except Exception:
            print("No cookies popup found")

    await page.waitForSelector("#trackingNumber")
    await page.type("#trackingNumber", tracking_number)

    await asyncio.sleep(2)

    await page.waitForSelector(".msc-search-autocomplete__field > .msc-cta-icon-simple > .msc-icon-search")
    await page.click(".msc-search-autocomplete__field > .msc-cta-icon-simple > .msc-icon-search")

    await page.waitForSelector(".msc-flow-tracking__details-subtitle")
    return page

# Function to get number of containers
async def get_number_of_containers(page):
    try:
        no_of_containers = await page.querySelectorEval(".msc-flow-tracking__details-subtitle > span:nth-child(2)", "(el) => parseInt(el.innerText)")
        return no_of_containers
    except Exception as error:
        print(error)
        return None

# Function to run script for one container
async def run_one_container_script(page):
    container = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__cell:nth-child(1) .data-value", "(el) => el.innerText")
    vessel = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__port:nth-child(4) .data-value:nth-child(2) > span:nth-child(2)", "(el) => el.innerText")
    voyage = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__port:nth-child(4) span:nth-child(3)", "(el) => el.innerText")

    return {'container': container, 'vessel': vessel, 'voyage': voyage}

# Function to run script for two containers
async def run_two_container_script(page):
    first_container = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__cell:nth-child(1) .data-value", "(el) => el.innerText")
    first_vessel = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__port:nth-child(4) .data-value:nth-child(2) > span:nth-child(2)", "(el) => el.innerText")
    first_voyage = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__port:nth-child(4) span:nth-child(3)", "(el) => el.innerText")

    second_container = await page.querySelectorEval("div:nth-child(5) .msc-flow-tracking__cell:nth-child(1) .data-value", "(el) => el.innerText")
    second_vessel = await page.querySelectorEval("div:nth-child(5) .msc-flow-tracking__port:nth-child(4) .data-value:nth-child(2) > span:nth-child(2)", "(el) => el.innerText")
    second_voyage = await page.querySelectorEval("div:nth-child(5) .msc-flow-tracking__port:nth-child(4) span:nth-child(3)", "(el) => el.innerText")

    return [
        {'container': first_container, 'vessel': first_vessel, 'voyage': first_voyage},
        {'container': second_container, 'vessel': second_vessel, 'voyage': second_voyage}
    ]

# Function to run script for three containers
async def run_three_container_script(page):
    first_container = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__cell:nth-child(1) .data-value", "(el) => el.innerText")
    first_vessel = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__port:nth-child(4) .data-value:nth-child(2) > span:nth-child(2)", "(el) => el.innerText")
    first_voyage = await page.querySelectorEval("div:nth-child(4) .msc-flow-tracking__port:nth-child(4) span:nth-child(3)", "(el) => el.innerText")

    second_container = await page.querySelectorEval("div:nth-child(5) .msc-flow-tracking__cell:nth-child(1) .data-value", "(el) => el.innerText")
    second_vessel = await page.querySelectorEval("div:nth-child(5) .msc-flow-tracking__port:nth-child(4) .data-value:nth-child(2) > span:nth-child(2)", "(el) => el.innerText")
    second_voyage = await page.querySelectorEval("div:nth-child(5) .msc-flow-tracking__port:nth-child(4) span:nth-child(3)", "(el) => el.innerText")

    third_container = await page.querySelectorEval("div:nth-child(6) .msc-flow-tracking__cell:nth-child(1) .data-value", "(el) => el.innerText")
    third_vessel = await page.querySelectorEval("div:nth-child(6) .msc-flow-tracking__port:nth-child(4) .data-value:nth-child(2) > span:nth-child(2)", "(el) => el.innerText")
    third_voyage = await page.querySelectorEval("div:nth-child(6) .msc-flow-tracking__port:nth-child(4) span:nth-child(3)", "(el) => el.innerText")

    return [
        {'container': first_container, 'vessel': first_vessel, 'voyage': first_voyage},
        {'container': second_container, 'vessel': second_vessel, 'voyage': second_voyage},
        {'container': third_container, 'vessel': third_vessel, 'voyage': third_voyage}
    ]

# Function to update Excel sheet
def update_excel_sheet(results):
    workbook = load_workbook(filename='SOA.xlsx')
    worksheet = workbook.active

    for row in range(6, 217):
        cell_address = 'H' + str(row)
        container_cell = worksheet[cell_address]

        if container_cell.value:
            container_name = container_cell.value

            result = next((result for result in results if result['container'] == container_name), None)
            if result:
                worksheet['I' + str(row)] = result['vessel']
                worksheet['J' + str(row)] = result['voyage']

    workbook.save('SOA.xlsx')
    

# Function to check if a tracking number already has a vessel and voyage number in the SOA sheet
def is_already_processed(tracking_number):
    # Open the SOA sheet
    workbook = openpyxl.load_workbook('SOA.xlsx')
    sheet = workbook.active

    # Iterate over the rows in the sheet
    for row in sheet.iter_rows(min_row=6, max_row=216):  # Rows 6 to 216
        # Check if the tracking number in the row matches the given tracking number
        if row[4].value == tracking_number:  # Column E
            # Check if the vessel and voyage number columns are not empty
            if row[8].value and row[9].value:  # Columns I and J
                return True

    return False

# Main function to run the appropriate script for each tracking number
async def main():
    # Load tracking numbers from a JSON file
    try:
        with open("tracking_numbers.json", "r") as file:
            tracking_numbers = json.load(file)
    except Exception as error:
        print("Error loading tracking numbers:", error)
        return

    # Convert list to set to remove duplicates
    tracking_numbers = list(set(tracking_numbers))

    browser = await launch(headless=False)

    for i in range(0, len(tracking_numbers), 15):
        results = []

        for tracking_number in tracking_numbers[i:i+15]:
            # Check if the tracking number already has a vessel and voyage number in the SOA sheet
            if is_already_processed(tracking_number):
                print(f"Tracking number {tracking_number} already processed, skipping...")
                continue
            # Wait for a random number of seconds between 5 and 15
            await asyncio.sleep(random.randint(5, 15))

            page = await open_tracking_page(browser, tracking_number)

            try:
                no_of_containers = await get_number_of_containers(page)

                if no_of_containers == 1:
                    result = await run_one_container_script(page)
                    results.append(result)
                elif no_of_containers == 2:
                    containers = await run_two_container_script(page)
                    results.extend(containers)
                elif no_of_containers == 3:
                    containers = await run_three_container_script(page)
                    results.extend(containers)
                else:
                    print(f"Tracking number {tracking_number}: Number of containers not supported by the script.")
            except Exception as error:
                print(f"Tracking number {tracking_number}:", error)
            finally:
                await page.close()


        # Update Excel sheet
        try:
            update_excel_sheet(results)
            print("Excel sheet updated")
        except Exception as error:
            print("Error updating Excel sheet:", error)

# Replace 'TRACKING_NUMBERS' with an array of actual tracking numbers
asyncio.run(main())