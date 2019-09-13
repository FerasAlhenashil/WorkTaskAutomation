import os
from datetime import datetime, timedelta
import re
from dateutil.relativedelta import relativedelta
import warnings
from openpyxl import load_workbook


class MelaniesToDo(object):
    project_indices = []
    ws_rows = []
    jobs = []
    matches = []
    todo = []

    def __init__(self):
        pass

    def extract(self, ws):
        fini = self.get_fini(ws)
        for row in ws:
            for cell in row:
                if cell.value == 'Status':
                    self.project_indices.append(cell.row - 1)
        self.project_indices.append(fini)

        for row in ws.values:
            self.ws_rows.append(list(row))

        for i in range(len(self.project_indices) - 1):
            temp = []
            start = self.project_indices[i]
            end = self.project_indices[i + 1]
            for j in range(end - start):
                temp.append(self.ws_rows[start + j])
            self.jobs.append(temp)

    def match(self):
        for i in range(len(self.jobs)):
            for j in range(len(self.jobs[i])):
                try:
                    temp = re.findall("((?:[1-9]|1[0-2])/(?:\\d)*)", self.jobs[i][j][4])
                except TypeError:
                    continue
                if temp:
                    #record loops position to be the key for finding a job
                    temp.append([i, j])
                    self.matches.append(temp)

    def date_format(self):
        for i in range(len(self.matches)):
            for j in range(len(self.matches[i]) - 1):
                self.matches[i][j] += '/19'
                self.matches[i][j] = datetime.strptime(self.matches[i][j], "%m/%d/%y")

        print(self.matches[26][4])

        for i in range(len(self.jobs)):
            for j in range(len(self.jobs[i])):
                try:
                    self.jobs[i][j][4] = self.jobs[i][j][4].split(",")
                except AttributeError:
                    continue

    def get_fini(self, tracker_ws):
        fini = tracker_ws.max_row
        while tracker_ws.cell(row=fini, column=1).value is None:
            fini -= 1
        return fini

    def check(self):
        today = datetime.today()
        tomorrow = today + relativedelta(days=1)
        two_days = today + relativedelta(days=2)
        print(len(self.jobs))
        print(self.matches)
        print(self.jobs[20][1][3])
        print(self.jobs[5][7][0])
        print(self.jobs[5][7][4][1])
        print(self.matches[20])
        print((self.matches[0][1] - today) > timedelta(days=4))
        for i in range(len(self.matches)):
            for j in range(len(self.matches[i]) - 1):
                if (self.matches[i][j] - today) > timedelta(days=2):
                    temp = [self.matches[i][-1][0], self.matches[i][-1][1], j]
                    self.todo.append(temp)
        print(self.todo)

    def to_do(self):
        pass

    def API(self):
        os.chdir('C:\\Users\\Feras\\Documents\\Work_Task_Automaiton')
        tracker = 'Melanies_Project_Tracker_MW (2) - Copy'
        tracker_file = tracker + '.xlsx'
        wb = load_workbook(tracker_file)
        ws = wb.active
        self.extract(ws)
        self.match()
        self.date_format()
        self.check()
        self.to_do()


def main():
    obj = MelaniesToDo()
    obj.API()


if __name__ == '__main__':
    main()
