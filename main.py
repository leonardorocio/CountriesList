from requests import get
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

class CountryList():

    def fetchData(self):
        '''
        params: self
        returns: A list of tuples containing name, capital, area and currencies of each country. The list
        is sorted based on the country names. 
        '''
        fetch_request = get("https://restcountries.com/v3.1/all")
        names = []
        capitals = []
        currencies = []
        areas = []
        if (fetch_request.status_code == 200):
            response_data = fetch_request.json();
            for k, v in enumerate(response_data):
                areas.append(v.get("area", "-"))
                capitals.append(v.get("capital", "-"))
                names.append(v["name"]["common"])
                # list comprehension to get all the country's currencies
                currencies.append(", ".join([i for i in v["currencies"]]) if v.get("currencies", "-") != "-" else "-")

        # First, the program zips the names, capitals, areas and currencies inside of an iterable
        # Then converts back to a list so it can apply the sorted function, with sorting key the first element
        # Which is the country's name
        zipped = sorted(list(zip(names, capitals, areas, currencies)), key= lambda x: x[0])
        return zipped

    def createSheet(self):
        '''
        params: self
        returns: None
        '''
        workbook = Workbook()

        # Creating the title cell
        active_worksheet = workbook.active
        active_worksheet["A1"] = "Countries List"
        active_worksheet.merge_cells("A1:D1")
        top_left_cell = active_worksheet["A1"]
        top_left_cell.font = Font(b=True, color="4F4F4F", size=16)
        top_left_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Creating the subtitle columns
        cells_dict = {"Name": "A2", "Capital":"B2", "Area":"C2", "Currencies":"D2"}
        for k, v in cells_dict.items():
            active_worksheet[v] = k
            cell = active_worksheet[v]
            cell.font = Font(b=True, color="808080", size=12)

        # Calling the fetchData function to insert all the data
        country_list = self.fetchData()
        for country in country_list:
            name, capital, area, currencies = country
            area = f'{area:,.2f}'
            active_worksheet.append([name, capital[0], area, currencies])

        workbook.save("countries.xlsx")

# Driver code
countries_list = CountryList()
countries_list.createSheet()
