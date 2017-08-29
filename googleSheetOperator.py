import gspread
import pprint
from oauth2client.service_account import ServiceAccountCredentials


class bulletin(object):
    result = None
    pp = None
    sheet = None

    def __init__(self):
        scope = ['https://spreadsheets.google.com/feeds']
        creds = ServiceAccountCredentials.from_json_keyfile_name("ppt project-ab4824ede4f2.json", scope)
        client = gspread.authorize(creds)
        self.sheet = client.open('Service Roles & Rotations').worksheet("Bulletin")
        self.pp = pprint.PrettyPrinter()
        self.result = self.sheet.col_values(1)

    def retrieveAnnouncements(self, dateString):

        #dateString = "5/14/2017"
        try:
            idx = self.result.index(dateString) + 1
        except:
            idx = -1

        if idx != -1:
            announcement_1 = self.sheet.cell(idx, 5).value.strip(' \t\n\r')
            announcement_2 = self.sheet.cell(idx, 6).value.strip(' \t\n\r')
            announcement_3 = self.sheet.cell(idx, 7).value.strip(' \t\n\r')
            announcement_4 = self.sheet.cell(idx, 8).value.strip(' \t\n\r')
            hasResults = True
            # self.pp.pprint(announcement_1) # Announcement 1
            # self.pp.pprint(announcement_2) # Announcement 2
            # self.pp.pprint(announcement_3) # Announcement 3
        else:
            print("Searchkey \"" + dateString + "\" does not exist.")
            announcement_1 = None
            announcement_2 = None
            announcement_3 = None
            announcement_4 = None
            hasResults = False

        return ([announcement_1, announcement_2, announcement_3, announcement_4], hasResults)
