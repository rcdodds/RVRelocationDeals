import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl


# Open a Selenium browser while printing status updates. Return said browser for use in scraping.
def open_selenium_browser(nickname, website):
    print('----------Scraping ' + nickname + '----------')
    chrome_options = Options()
    chrome_options.headless = True
    print('Opening Selenium browser')
    sele = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    print('Selenium browser opened')
    print('Opening ' + nickname)
    sele.get(website)
    print(nickname + ' opened')
    return sele


# Close a Selenium browser while printing status updates.
def close_selenium(name, slnm):
    print('Data has been scraped & saved')
    slnm.close()
    print('----------' + name + ' Complete----------')


# Scrape IMoovA.com and clean data -- No inputs, creates CSV, & returns data frame
def scrape_imoova():
    # Start a selenium browser & navigate to IMoovA.com (USA tab)
    browser = open_selenium_browser('IMoovA.com', 'https://imoova.com/imoova/relocations?region_id=3')

    # Store each row of the table as a selenium element
    table_rows = browser.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr')

    # Store the contents of the table -- Texts in a list of lists -- Info & Orders in their own respective lists
    print('Scraping data')
    cell_texts, info_links, order_links = ([] for j in range(3))
    for row in range(len(table_rows)):
        # Store text of each cell
        xpath = '//*[@id="dataTable"]/tbody/tr[' + str(row + 1) + ']/td'
        cell_texts.append([cell.text.replace(',', '') for cell in browser.find_elements_by_xpath(xpath)][:-2])
        # Store the reference ID for the more info / book now pages
        link_xpath = xpath + '[6]/a'
        link = browser.find_element_by_xpath(link_xpath).get_attribute('href')
        info_links.append(link)
        order_links.append(link.replace('view', 'order'))

    # Convert the list of lists to a data frame
    col_names = ['From', 'To', 'Bonus', 'Earliest Pick Up', 'Latest Drop Off', 'RV Type',
                 'Rate', 'Days Allowed', 'Extra Days']
    imoova = pd.DataFrame(cell_texts, columns=col_names)
    col_names.extend(['More Info', 'Order'])
    imoova['More Info'] = info_links
    imoova['Order'] = order_links

    # Clean up data
    imoova['Rate'] = [rate.split('.')[0] for rate in imoova['Rate']]
    imoova = imoova.groupby(col_names).size().reset_index()
    imoova.rename(columns={0: 'RVs'}, inplace=True)

    # Add miles allowance column
    miles = []
    for page in imoova['More Info']:
        try:    # Attempt to pull the number of miles allowed
            browser.get(page)
            more_info = [box.text.replace(',', '') for box in browser.find_elements_by_css_selector('td')]
            miles.append(more_info[more_info.index('Miles allowance:') + 1])
        except ValueError:
            miles.append('??')      # Couldn't find miles allowed
    imoova['Miles Included'] = miles

    # Done
    close_selenium('IMoovA.com', browser)
    return imoova


# Scrape ElMonteRV.com and clean data -- No inputs, creates CSV, & returns data frame
def scrape_elmonte():
    # Start a selenium browser
    driver = open_selenium_browser('ElMonteRV.com',
                                   'https://www.elmonterv.com/rv-rental/cool-deal-detail/ONE-WAY-SPECIAL/')

    # Store each row of the table as a selenium element. Skip over table of airport codes
    print('Scraping data')
    rows = driver.find_elements_by_xpath('//div/div[1]/table[2]/tbody/tr')

    # Scrape each table row
    cell_texts = []
    for row in rows[1:-1]:
        # Scrape each cell
        cells = row.find_elements_by_css_selector('td')
        # Store relevant data
        cell_texts.append([" ".join(cll.text.split()).replace(',', '') for cll in cells])
        # Specify fuel credit if necessary
        if len(cell_texts[-1][-1]):
            cell_texts[-1][-1] = cell_texts[-1][-1] + ' fuel credit'

    # Convert the list of lists to a data frame
    cols = ['From', 'To', 'Earliest Pick Up', 'Latest Drop Off', 'RV Type',
            'RVs', 'Rate', 'Days Allowed', 'Miles Included', 'Bonus']
    elmonte = pd.DataFrame(cell_texts, columns=cols)

    # Clean up data
    elmonte['Earliest Pick Up'] = [' '.join(pickup.split('-')[:-1]) for pickup in elmonte['Earliest Pick Up']]
    elmonte['Latest Drop Off'] = [' '.join(pickup.split('-')[:-1]) for pickup in elmonte['Latest Drop Off']]
    elmonte = elmonte[~elmonte.RVs.str.contains('none')]    # Get rid of listings with no remaining RVs
    elmonte['From'].replace('', np.nan, inplace=True)       # Replace blank pick up spots with null value
    elmonte.dropna(subset=['From'], inplace=True)           # Drop rows with null pick up spots
    elmonte['Earliest Pick Up'].replace('', np.nan, inplace=True)       # Replace blank pick up dates with null value
    elmonte.dropna(subset=['Earliest Pick Up'], inplace=True)           # Drop rows with null pick up dates

    # Add link to RV info and order info
    rvs_url = 'https://www.elmonterv.com/rv-rental/rvs-we-rent/'
    driver.get(rvs_url)
    rv_links = []
    for rv in elmonte['RV Type']:
        try:
            rv_links.append(driver.find_element_by_partial_link_text(rv).get_attribute('href'))
        except:
            print('RV Type Not Found - ' + rv)
            rv_links.append('Not Found')
    elmonte['More Info'] = rv_links
    elmonte['Order'] = ['Email reservations@elmonterv.com and mention FRDS.' for i in range(len(elmonte))]

    # Done
    close_selenium('ElMonteRV.com', driver)
    return elmonte


# Add column to data frame containing link to Google Maps
def google_maps(df):
    gmaps = []
    for offer in range(len(df)):
        origin = df['From'][offer].replace(' ', '+')
        dest = df['To'][offer].replace(' ', '+')
        gmaps.append(
            'https://www.google.com/maps/dir/?api=1&travelmode=driving&origin=' + origin + '&destination=' + dest)
    df['Google Maps'] = gmaps
    return df


# (1) Scrape websites. (2) Stitch the data together. (3) Store data in CSV backed up to Google Drive.
def main():
    # Scrape websites, storing results in a dictionary. Keys = site names. Elements = data frames.
    scraped_data = {
        'I Moov A': scrape_imoova(),
        'El Monte RV': scrape_elmonte()
    }

    # Stitch data together
    relocations = pd.concat(scraped_data.values(), keys=scraped_data.keys(), names=['Website'])
    relocations = google_maps(relocations)  # Add google maps link

    # Rearrange columns
    columns = ['From', 'To', 'Rate', 'Earliest Pick Up', 'Latest Drop Off', 'RV Type', 'Days Allowed',
               'Extra Days', 'RVs', 'Miles Included', 'Bonus', 'Google Maps', 'More Info', 'Order']
    relocations = relocations[columns]

    # Store in file which is automatically backed up to Google Drive (and can be shared)
    path = 'C:\\Users\\rcdodds\\Google Drive\\Travel\\RV Relocation Deals.xlsx'
    print('Saving data to ' + path)
    relocations.to_excel(path, index=False)
    print('Data saved to ' + path)
    print('----------Program Complete----------')


# Let's get it going
if __name__ == "__main__":
    main()
