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
                end = len(art_works[i])-10
                schedule_dates.append(dates[start:end])
        return schedule_dates

    def odd_job(self, job):
        try:
            start = job.index('PM Prep')
            return start
        except ValueError:
            print(job[7], 'Error: has no DV Shell due nor PM Prep tasks')

    def ws_tracker_schedule(self, schedule, schedule_dates):
        print('we are in tracker')
        tracker_schedule = []
        for i in range(len(schedule_dates)):
            cell_schedule = []
            for j in range(len(schedule_dates[i])):
                schedule_str = str(schedule_dates[i][j].month) + "/" + str(schedule_dates[i][j].day) + str(schedule[i][j])
                if (schedule[i][j]) is not None:
                    cell_schedule.append(schedule_str)
            tracker_schedule.append(cell_schedule)
        print(tracker_schedule[0])
        return tracker_schedule

    def ws_append(self, tracker_ws, tracker_schedule, art_works, schedule, titles):
        last_row = tracker_ws.max_row
        print(last_row)



    def API(self):
        os.chdir('C:\\Users\\Feras\\Documents\\Work_Task_Automaiton')
        wb = openpyxl.load_workbook('Copy of 190510_IN_Rebrand_Schedule_50116 (1).xlsm')
        ws = wb.active
        tracker_wb = openpyxl.load_workbook('Version2_Melanies_Project_Tracker_MW - Copy.xlsx')
        tracker_ws = tracker_wb.active
        art_works = self.ws_art_works(ws)
        titles = self.ws_titles(ws)
        schedule = self.ws_schedule(art_works)
        schedule_dates = self.ws_schedule_dates(art_works, ws)
        tracker_schedule = self.ws_tracker_schedule(schedule, schedule_dates)
        self.ws_append(tracker_ws, tracker_schedule, art_works, schedule, titles)

        """for i in range(len(schedule)):
            print(schedule[i])
            print(schedule_dates[i])
            # the last date on schedule is: schedule_dates[i][len(schedule_dates[i]) - 1]
            #print(schedule_dates[i][len(schedule_dates[i])-1].value)
        x = len(schedule_dates[0])-1
        y = 0
        print(art_works[y][6])
        print(schedule_dates[y][x].month, '/', schedule_dates[y][x].day, end=' ')
        print(schedule[y][x])
        print(len(schedule[y]))"""


def main():
    obj = WorkTaskAutomation()
    obj.API()

if __name__ == '__main__':
    main()