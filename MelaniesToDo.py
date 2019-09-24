import os
from datetime import datetime, timedelta
import re
import warnings

from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook


class Melaniestodo_today(object):
    project_indices = []
    ws_rows = []
    jobs = []
    matches = []
    todo_today = []
    todo_tomorrow = []
    todo_2days = []
    output_today = []
    output_tomorrow = []

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
                    #record loops position to be the key for finding a task
                    temp.append([i, j])
                    self.matches.append(temp)

    def format_date(self):
        for i in range(len(self.matches)):
            for j in range(len(self.matches[i]) - 1):
                self.matches[i][j] += '/19'
                self.matches[i][j] = datetime.strptime(self.matches[i][j], "%m/%d/%y")
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
        monday = today + relativedelta(days=1)
        for i in range(len(self.matches)):
            for j in range(len(self.matches[i])-1):
                if timedelta(hours=-24) < (self.matches[i][j] - today) < timedelta(hours=0):
                    temp = [self.matches[i][-1][0], self.matches[i][-1][1], j]
                    self.todo_today.append(temp)

        for i in range(len(self.matches)):
            for j in range(len(self.matches[i])-1):
                if today.weekday() < 3:
                    if timedelta(hours=-6) < (self.matches[i][j] - today) < timedelta(days=1):
                        temp = [self.matches[i][-1][0], self.matches[i][-1][1], j]
                        self.todo_tomorrow.append(temp)
                else:
                    if timedelta(hours=1) < (self.matches[i][j] - today) < timedelta(days=3):
                        temp = [self.matches[i][-1][0], self.matches[i][-1][1], j]
                        self.todo_tomorrow.append(temp)

        for i in range(len(self.matches)):
            for j in range(len(self.matches[i])-1):
                if today.weekday() < 3:
                    if timedelta(days=1) < (self.matches[i][j] - today) < timedelta(days=2):
                        temp = [self.matches[i][-1][0], self.matches[i][-1][1], j]
                        self.todo_2days.append(temp)
                elif today.weekday() == 3:
                    if timedelta(days=2) < (self.matches[i][j] - today) < timedelta(days=4):
                        temp = [self.matches[i][-1][0], self.matches[i][-1][1], j]
                        self.todo_2days.append(temp)
                else:
                    if timedelta(days=3) < (self.matches[i][j] - today) < timedelta(days=4):
                        temp = [self.matches[i][-1][0], self.matches[i][-1][1], j]
                        self.todo_2days.append(temp)
        print(self.todo_2days)

    def to_do(self):
        print(" ")
        print("########################## today ###############################")
        for i in range(len(self.todo_today)):
            temp = [self.jobs[self.todo_today[i][0]][0][0], self.jobs[self.todo_today[i][0]][self.todo_today[i][1]][0],
                    self.jobs[self.todo_today[i][0]][self.todo_today[i][1]][4][self.todo_today[i][2]]]
            self.output_today.append(temp)
        #print(self.output_today[3][2].endswith("Xlation") or self.output_today[3][2].endswith("Xlation "))

        for i in range(len(self.output_today)):
            task = self.output_today[i][2]
            if task.endswith("PM Prep"):
                task += " \"Inputs due from PM\""
            elif task.endswith("Prep2"):
                task += " \"Due for Q&A\""
            elif task.endswith("Q&A"):
                task += " \"Feedback due from Cambridge - send to prep 3\""
            elif task.endswith("Prep3"):
                task += " \"Inputs for Format due from GK\""
            elif task.endswith("FQA"):
                task += " \"Due from FQA\""
            elif task.endswith("FQA/CEPR/PMTSGK") or task.endswith("CEPR/PMTSGK") or task.endswith("FQA/PMTSGK"):
                task += " \"Due from entry PMTSGK\""
            elif task.endswith("CE"):
                task += " \"Due from CE - send to entry or ICR\""
            elif task.endswith("CE/Upload for ICR"):
                task += " \"CE - send to ICR\""
            elif task.endswith("Finalize") or task.endswith("Delivery Prep"):
                task += " \"Due from delivery prep\""
            elif task.endswith("PDQA "):
                task += " \"Due from PDQA\""
            elif task.endswith("Delivery") or task.endswith("FFCH"):
                task += " \"Due from FFCH\""
            elif task.endswith("GK Format"):
                task += " \"Formatted files due. Send to proofreading/Check Entry\""
            elif task.endswith("Format"):
                task += " \"Formatted files due. Send to proofreading/Check Entry\""
            elif task.endswith("PR/CE") or task.endswith("Proof/CE") or task.endswith("CE/PR") or \
                    task.endswith("CE/Proof") or task.endswith("CE (LING)") or task.endswith("CE + PR"):
                task += " \"Complete CE and verify linguistic feedbackhas been incorporated." \
                        " Send to entry or upload for ICR/ICA\""
            elif task.endswith("PR/CE/Upload for ICA") or task.endswith("PR/CE/Upload for ICR"):
                task += " \"Complete CE and possible LSO send to ICR/ICA upload\""
            elif task.endswith(" PMTSGK") or task.endswith("PM/TS CH") or task.endswith("PM") or task.endswith("TS + GK"):
                task += " \"TS Check and send to CEPR or Entry\""
            self.output_today[i][2] = task
        print(self.output_today)

        print("########################## tomorrow ###############################")
        for i in range(len(self.todo_tomorrow)):
            temp = [self.jobs[self.todo_tomorrow[i][0]][0][0],
                    self.jobs[self.todo_tomorrow[i][0]][self.todo_tomorrow[i][1]][0],
                    self.jobs[self.todo_tomorrow[i][0]][self.todo_tomorrow[i][1]][4][self.todo_tomorrow[i][2]]]
            self.output_tomorrow.append(temp)
        print(self.output_tomorrow)
        print(self.output_tomorrow[0][2].endswith("Upload for ICR"))

        """print("########################## 2days ###############################")
        for i in range(len(self.todo_2days)):
            print(" ")
            print(self.jobs[self.todo_2days[i][0][0]][0][0])
            print(self.jobs[self.todo_2days[i][0][0]][self.todo_2days[i][0][1]][0])
            for j in range(len(self.todo_2days[i])):
                print(self.jobs[self.todo_2days[i][j][0]][self.todo_2days[i][j][1]][4][self.todo_2days[i][j][2]])"""

    def output(self):
        pass



    def API(self):
        os.chdir('C:\\Users\\Feras\\Documents\\Work_Task_Automaiton')
        tracker = 'Melanies_Project_Tracker_MW (2) - Copy'
        tracker_file = tracker + '.xlsx'
        wb = load_workbook(tracker_file)
        ws = wb.active
        self.extract(ws)
        self.match()
        self.format_date()
        self.check()
        self.to_do()


def main():
    obj = Melaniestodo_today()
    obj.API()


if __name__ == '__main__':
    main()
