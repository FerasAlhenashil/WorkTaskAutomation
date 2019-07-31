import openpyxl
import os
class WorkTaskAutomation(object):
    def __init__(self):
        pass

    def ws_titles(self, ws):
        titles = []
        for row in ws:
            for cell in row:
                if cell.column is 7:
                    if cell.value is None:
                        cell = cell.offset(0, 1)
                        titles.append(cell.value)
        return titles

    def ws_Melanie(self, ws):
        melanie = []
        for row in ws.values:
            for value in row:
                if value == 'Melanie':
                    melanie.append(row)
        return melanie

    def API(self):
        os.chdir('C:\\Users\\Feras\\Documents\\Work_Task_Automaiton')
        wb = openpyxl.load_workbook('Copy of 190510_IN_Rebrand_Schedule_50116 (1).xlsm')
        print(type(wb))
        ws = wb.active
        melanie = self.ws_Melanie(ws)

        for i in range(len(melanie)):
            print(melanie[i])

        titles = self.ws_titles(ws)
        print(titles)




def main():
    obj = WorkTaskAutomation()
    obj.API()


if __name__ == '__main__':
    main()