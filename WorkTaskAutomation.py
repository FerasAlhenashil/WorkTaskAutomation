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

    def ws_art_works(self, ws):
        art_works = []
        for row in ws.values:
            for value in row:
                if value == 'Melanie':
                    art_works.append(row)
        return art_works

    def ws_schedule(self, art_works):
        schedule = []
        for i in range(len(art_works)):
            try:
                start = art_works[i].index('DV Shell due ')
                end = art_works[i].index('ICR')
                schedule.append(art_works[i][start:end])
                #print(schedule[i])
            except ValueError:
                start = self.odd_job(art_works[i])
                end = len(art_works[i])
                schedule.append(art_works[i][start:end])
                #print(schedule[i])
        return schedule

    def ws_schedule_dates(self, art_works, ws):
        schedule_dates = []
        dates = []
        for row in ws.values:
            for value in row:
                dates.append(value)
        for i in range(len(art_works)):
            try:
                start = art_works[i].index('DV Shell due ')
                end = art_works[i].index('ICR')
                schedule_dates.append(dates[start:end])
            except ValueError:
                start = self.odd_job(art_works[i])
                end = len(art_works[i])
                schedule_dates.append(dates[start:end])
        return schedule_dates

    def odd_job(self, job):
        try:
            start = job.index('PM Prep')
            return start
        except ValueError:
            print(job[7], 'Error: has no DV Shell due nor PM Prep tasks')

    def API(self):
        os.chdir('C:\\Users\\Feras\\Documents\\Work_Task_Automaiton')
        wb = openpyxl.load_workbook('50117_Rebrand_Schedule_FR_and_BEDE_190624.xlsm')
        ws = wb.active
        #each art_work is a job from the schedule
        art_works = self.ws_art_works(ws)
        titles = self.ws_titles(ws)
        schedule = self.ws_schedule(art_works)
        schedule_dates = self.ws_schedule_dates(art_works, ws)

        for i in range(len(art_works)):
            print(art_works[i])

        for i in range(len(schedule)):
            print(schedule[i])
            print(schedule_dates[i])
            # the last date on schedule is: schedule_dates[i][len(schedule_dates[i]) - 1]
            #print(schedule_dates[i][len(schedule_dates[i])-1].value)
        x = 0
        print(art_works[2][6])
        print(schedule_dates[2][x].month, '/', schedule_dates[2][x].day, end=' ')
        print(schedule[2][x])


def main():
    obj = WorkTaskAutomation()
    obj.API()

if __name__ == '__main__':
    main()