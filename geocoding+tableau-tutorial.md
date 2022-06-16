# Geocoding + Tableau Tutorial #
*Note: this tutorial is only for geocoding and for creating the Tableau visual, not for finding NAICS codes (since that can be done manually) or other functions.*
* Set up Python environment and download geocoding scripts.
    * Download Python (https://www.python.org/).
    * At a terminal, type `python --version` and press enter. If you see an output like `Python x.x.x`, you're fine.
    * Type `pip --version`. If you don't see an error, you're fine. If you do see an error, that means pip is not installed. Follow https://pip.pypa.io/en/stable/installation/ to install pip.
    * Install the necessary Python libraries by typing `pip install pandas numpy requests geopy`
        * Pandas and numpy are the Python heavyweights of data analysis, requests is a library that allows for easy web (HTTP, etc) communication in Python, and geopy is a library that makes querying geocoding APIs a whole lot easier.
    * In this repository, click the green "Code" button and select the zip archive option.![image](https://user-images.githubusercontent.com/80125711/160976565-c5478463-d377-4483-888b-f03b8e9fa55a.png)
    * Extract the zip archive to some place you will remember.
* Set up Tableau.
    * Download Tableau Public (https://public.tableau.com/en-us/s/). You will have to sign up for an account if you don't already have one.
    * Install and start Tableau.
* Set up company lists with full addresses.
    * In Google Sheets, remove all entities that don't have a full, valid address (or make up a valid address).
    * Find an empty column, and in the first cell of that empty column type `Full Address`.
    * In the cell right below it, type a formula to form a full address from the adjacent cells in the spreadsheet. For example, if "1234 Random Road" is in cell `A1`, "Random City" is in cell `B1`, and "12456" is in cell `D1`, to form the full address "1234 Random Road, Random City, OR 12456" (assuming this is the Oregon list), you would type into the cell `=A1&", "&B1&", OR "&D1`, where the `&` operator is the glue that joins together the strings.
        * Press enter to evaluate the formula, and copy that same formula down to all the other cells that have a full, valid address.
     * Download the spreadsheet to the same directory as the Python scripts.
* Nominatim geocoding.
    * In the Command Prompt/Terminal window, navigate to the directory where the scripts and the company list are. Use the `cd` command (for example, if I wanted to go to a folder named `Stuff` within the current directory I am in, I would type `cd Stuff` and press enter to go within the `Stuff` folder.
    * Type `python geocoding_nominatim.py` and press enter to execute the Nominatim geocoding script. Enter in the appropriate values when prompted.
        * __DO NOTE__ At least on Windows, if you want to use a directory path like C:\Users\user\Downloads\list.xlsx, you should input it as C:\\Users\\user\\Downloads\\list.xlsx. If you run into an error, try this first.
    * After the script completes, you should find a file titled `nominatim.xlsx` in the same folder. Open it, and copy everything from the row marked 1 down the the last marked row to the company Excel sheet in two empty columns. Title the columns `Latitude` and `Longitude`.
    * Save the Excel file.
    * Copy and paste the coordinates back to Google Sheets.
* Google geocoding (Handled by me due to increased complexity)
    * Re-download the Google Sheet with the Nominatim coordinates.
    * Likely, not all companies listed will have a latitude/longitude, since Nominatim is usually less powerful compared to other solutions, which require an API key. Google, as you might suspect, provides one of the best (if not the best) geocoding services. However, the process to obtain an API key is quite convoluted.
    * The official Google docs is at https://developers.google.com/maps/documentation/geocoding. The relevant section is Setup, but it's good to get context from the other sections. Do note, it is not necessary to restrict the use of the API key, since I assume you won't post it publicly ;) 
    * Copy the API key for use later.
    * Go back to the Command Prompt/Terminal, and this time run `python geocoding_gmaps.py`. Enter in the appropriate fields as requested.
        * *WATCH YOUR USAGE!* Make sure that you don't go over 40k requests/month, or you will be billed.
    * This time, a file named `googlev3.xlsx` will be created. Copy the non-blank rows into the corresponding locations in the company Excel sheet.
    * If there are still empty lat/longs, they will have to be found manually.
    * Save the Excel file.
    * Copy and paste the coordinates back to Google Sheets.
* Tableau!
    * Download the Google Sheets as an Excel file. DO NOT download as csv, Tableau has issues with refreshing csv data sources.
    * Remove all columns we do not want to be public.
    * Open Tableau and add a new Excel data source. Navigate to the Excel file.
    * If you do not see a screen like this: ![image](https://user-images.githubusercontent.com/80125711/162134035-f6f75f34-5d07-4bbc-b190-e89bcabac36f.png)
drag the appropriate sheet under "Sheets" on the left toolbar to the central gray space and release.
    * Switch to Sheet 1. Ensure that Latitude and Longitude fields are of type "Dimension." Check by clicking on this little dropdown arrow: ![image](https://user-images.githubusercontent.com/80125711/162134375-6cbe383f-5f83-47f4-837c-a13abe0e3e1e.png)
and ensuring "Dimension" is checked.
    * Drag and drop Latitude to the "Columns" area, and Longitude to the "Rows" area: ![image](https://user-images.githubusercontent.com/80125711/162134706-cb6e0473-6087-4736-ba03-2a706ef8b9c3.png). Click on the dropdowns next to Latitude and Longitude and select Dimension for both.
    * Click on the "Show Me" button at the top right corner and click the Symbol Map tile. ![image](https://user-images.githubusercontent.com/80125711/162135132-02df76c9-e1d8-4139-bfdf-735382047a08.png)
    * To filter data, drag a field like "City" to the Filters area. If it is properly set as a Dimension, a window like this should show up: ![image](https://user-images.githubusercontent.com/80125711/162135374-fd7e070b-012b-4e3f-a38f-c3ee6f6b4188.png). Ensure that "Select from list" is checked, click "All" if not all fields are selected, and press "OK" to exit. Repeat for other filters you may want to add.
    * In the Filters area, click on each field's dropdown menu and click "Show Filter." They should pop up on the right hand side.
    * Click the dropdown icon next to the filters on the right hand side and select "Only Relevant Values." ![image](https://user-images.githubusercontent.com/80125711/162135753-930bbc81-83ee-4cb5-a8b4-953386baee91.png)
    * Save the viz to Tableau Public by going to File > Save to Tableau Public. Choose a name to save it as (login if necessary), and the viz should automatically open in your default browser.
    * __To update data, as responses come in from Google form__
        * Close the Tableau Public app on your computer.
        * Open Google form link and go to the "Responses" tab on the top.
        * Open locally saved Excel file (linked to Tableau), and add new rows/modify information on the Excel sheet as necessary. Make the same changes to the "FINAL Company Lists" Google Sheet.
        * Reopen Tableau Public and open the Oregon_CT or Washington_CT project, whichever was edited. The edits should automatically be ingested into Tableau. If they are not automatically ingested, go to the "Data Source" tab in the bottom left hand corner and resolve issues there (re-link Excel workbook if necessary).
        * Re-publish the viz to Tableau Public by going to File > Save to Tableau Public.
