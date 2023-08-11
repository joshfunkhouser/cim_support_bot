import asyncio
from pyppeteer import launch
from config import login_credentials
import time
import json
import html
import os
import pandas as pd

# project location
proj_dir = os.getcwd()

# market matches
market_match = [['Washington DC','washington_dc'],['Philadelphia', 'philadelphia']]

# cim login credentials
credentials = login_credentials()
username = credentials[0]
password = credentials[1]


def get_option_count(element):
    decoded_data_tagger = html.unescape(element)
    options = json.loads(decoded_data_tagger)
    option_count = len(options)
    return option_count


async def cim_open(page):
    print("Opening CIM")
    await page.goto("https://supplierhub.homedepot.com/")


async def cim_login(page):
    print("Signing in to CIM")
    await page.click(
        "#sign-in-button")
    await asyncio.sleep(5)
    await page.type(
        "#inputUsername", username
    )
    await page.type(
        "#inputPassword", password
    )
    await page.click("#buttonSignOn")


async def cim_support_ticket(page):
    print("going to support ticket page")

    # navbar support button
    nav_support = await page.waitForXPath('//*[@id="support"]', timeout=30000)
    await nav_support.click()

    # navbar submit new ticket option
    new_ticket = await page.waitForXPath('//*[@id="new-ticket"]', timeout=30000)
    await new_ticket.click()
    await asyncio.sleep(3)

    # select supplier type
    print("selecting supplier type")
    number_of_options = get_option_count('[{&quot;label&quot;:&quot;-&quot;,&quot;value&quot;:&quot;&quot;},{&quot;label&quot;:&quot;Core/Retail&quot;,&quot;value&quot;:&quot;core/retail_supplier&quot;},{&quot;label&quot;:&quot;HD Pro&quot;,&quot;value&quot;:&quot;hd_pro_supplier&quot;},{&quot;label&quot;:&quot;QuoteCenter&quot;,&quot;value&quot;:&quot;quote_center_supplier&quot;},{&quot;label&quot;:&quot;Canada&quot;,&quot;value&quot;:&quot;canada_supplier&quot;},{&quot;label&quot;:&quot;Store/Associate Use only (Suppliers should not select)&quot;,&quot;value&quot;:&quot;store/associate&quot;}]')
    supplier_type = await page.waitForSelector(
        '#new_request > div.form-field.string.required.request_custom_fields_360029395831 > a')
    await asyncio.sleep(1)
    for a in range(0, number_of_options+1):
        await supplier_type.click()
        for b in range(0, a+1):
            await page.keyboard.down('ArrowDown')
        await page.keyboard.press('Enter')
        await asyncio.sleep(1)
        new_value = await page.evaluate('(element) => element.value',
                                        await page.querySelector('#request_custom_fields_360029395831'))
        if new_value == "core/retail_supplier":
            break
    await asyncio.sleep(1)

    # Selecting the value "transportation_dms_carrier_support"
    print("selecting issue")
    await page.evaluate('''() => {
        const input = document.querySelector('input[name="request[custom_fields][360014494051]"]');
        input.value = 'transportation__dms_carrier_support';
    }''')
    await asyncio.sleep(1)

    # Selecting I am a
    print("selecting i am a carrier")
    await page.evaluate('''() => {
        const input = document.querySelector('input[name="request[custom_fields][360024820012]"]');
        input.value = 'carrier_dms_recipient';
    }''')
    await asyncio.sleep(1)

    # Selecting network
    print("selecting network")
    await page.evaluate('''() => {
        const input = document.querySelector('input[name="request[custom_fields][360033229132]');
        input.value = 'dfs';
    }''')
    await asyncio.sleep(1)

    # Selecting category
    print("selecting category")
    await page.evaluate('''() => {
            const input = document.querySelector('input[name="request[custom_fields][360024820492]');
            input.value = 'territory';
        }''')
    await asyncio.sleep(1)

    # Selecting division
    print("selecting division")
    await page.evaluate('''() => {
            const input = document.querySelector('input[name="request[custom_fields][360025129731]');
            input.value = 'northern';
        }''')
    await asyncio.sleep(1)

    # selecting territory type
    print("selecting territory type")
    await page.evaluate('''() => {
            const input = document.querySelector('input[name="request[custom_fields][360024821152]');
            input.value = '22';
        }''')
    await asyncio.sleep(1)

    # specific trucks
    print("selecting specific trucks")
    await page.evaluate('''() => {
            const input = document.querySelector('input[name="request[custom_fields][1900000010427]');
            input.value = 'no_specific_trucks';
        }''')
    await asyncio.sleep(1)

    # add address
    print("adding address")
    await page.type('#request_custom_fields_360024821192', 'This is a test address.')
    await asyncio.sleep(1)

    # add emails
    print("adding emails")
    await page.type('#request_collaborators_', 'joe.simmons@estes-express.com')
    await asyncio.sleep(1)

    # add description
    print("adding description")
    await page.type('#request_description', 'This is a test description.')
    await asyncio.sleep(1)

    # add subject
    print("adding subject")
    await page.type('#request_subject', 'This is a test subject.')

    time.sleep(999)


async def main():

    # launch browser
    browser = await launch(
        headless=False,  # Set headless to False to display the browser window
        defaultViewport=None,  # Disable the default viewport
        args=['--start-maximized']  # Use the '--start-maximized' argument to maximize the window
    )
    page = (await browser.pages())[0]

    # open cim page
    await cim_open(page)

    # login to cim page
    await cim_login(page)

    # fill out support ticket
    await cim_support_ticket(page)

    # Close the browser
    await browser.close()


# get input file
input_dir = os.path.join(os.getcwd(), "input files")
file_list = os.listdir(input_dir)
excel_files = [file for file in file_list if file.endswith(".xlsx") or file.endswith(".xls")]
required_columns = ["Market", "CSA", "WO", "Address", "Notes", "INI Ticket number"]
# Read Excel files into Pandas DataFrames
valid_dataframes = []
for excel_file in excel_files:
    file_path = os.path.join(input_dir, excel_file)
    df = pd.read_excel(file_path, converters={'WO': str})
    if all(col in df.columns for col in required_columns):
        valid_dataframes.append(df)
    else:
        print(f"File '{excel_file}' does not have all required columns.")
ticket_list = pd.concat(valid_dataframes, ignore_index=True)

# clean up input file
ticket_list['Market'] = ticket_list['Market'].str.replace('Market: ', '')
ticket_list['Market'] = ticket_list['Market'].replace('unk', 'Washington DC')
ticket_list['Address'].fillna('X', inplace=True)
ticket_list['subject'] = '22 zone required ' + ticket_list['CSA']
ticket_list['description'] = ticket_list['WO'] + ticket_list['Notes'].apply(lambda x: f" - {x}" if pd.notnull(x) and x != "" else "")

# output file directory
output_dir = os.path.join(os.getcwd(), "input files\\processed\\output.xlsx")
ticket_list.to_excel(output_dir)
quit()


# run the asyncio event loop
asyncio.get_event_loop().run_until_complete(main())
