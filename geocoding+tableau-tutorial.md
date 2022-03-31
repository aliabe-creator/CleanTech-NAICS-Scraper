# Geocoding + Tableau Tutorial #
*Note: this tutorial is only for geocoding and for creating the Tableau visual, not for finding NAICS codes (since that can be done manually) or other functions.*
* Set up Python environment and download geocoding scripts.
    * Download Python (https://www.python.org/).
    * At a terminal, type `python --version` and press enter. If you see an output like `Python x.x.x`, you're fine.
    * Type `pip --version`. If you don't see an error, you're fine. If you do see an error, that means pip is not installed. Follow https://pip.pypa.io/en/stable/installation/ to install pip.
    * Install the necessary Python libraries by typing `pip install pandas numpy requests geopy`. Pandas and numpy are the Python heavyweights of data analysis, requests is a library that allows for easy web (HTTP, etc) communication in Python, and geopy is a library that makes querying geocoding APIs a whole lot easier.
    * In this repository, click the green "Download
2.	Download Tableau, sign in.
3.	Set up company lists with full addresses.
4.	Geocode, combine outputs.
a.	Todo: upload the new versions to GH
b.	For Google, need to add a billing account, so watch your requests.
5.	Insert into Tableau, set up filters and publish.
