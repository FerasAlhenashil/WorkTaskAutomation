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

    def ws_project_sizes(self, art_works):
        project_sizes = [0]
        j = 0
        for i in range(len(art_works)):
            if art_works[j][3] != art_works[i][3]:
                project_sizes.append(i)
                j = i
        return project_sizes

    def ws_schedule(self, art_works):
        schedule = []
        for i in range(len(art_works)):
            try:
                start = art_works[i].index('DV Shell due ')
                end = art_works[i].index('ICR')
                schedule.append(art_works[i][start:end])
            except ValueError:
                start = self.odd_job(art_works[i])
                end = len(art_works[i])
                schedule.append(art_works[i][start:end])
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
        tracker_schedule = []
        for i in range(len(schedule_dates)):
            cell_schedule = []
            for j in range(len(schedule_dates[i])):
                schedule_str = str(schedule_dates[i][j].month) + "/" + str(schedule_dates[i][j].day) + str(schedule[i][j])
                if (schedule[i][j]) is not None:
                    cell_schedule.append(schedule_str)
            tracker_schedule.append(cell_schedule)
        return tracker_schedule
    def get_fini(self, tracker_ws):
        fini = tracker_ws.max_row
        while tracker_ws.cell(row=fini, column=1).value is None:
            fini -= 1
        return fini

    def ws_append(self, tracker_ws, tracker_schedule, art_works, titles, project_sizes):
        fini = self.get_fini(tracker_ws)
        try:
            tracker_ws.unmerge_cells(start_row=fini, start_column=1, end_row=fini, end_column=6)
        except ValueError:
            print('cells are not merged')
        #tracker_ws.merge_cells(start_row=fini, start_column=1, end_row=fini, end_column=6)
        for i in range(len(titles)):
            tracker_ws.insert_rows(fini)
            fini = self.get_fini(tracker_ws)
            self.tracker_header(tracker_ws, titles, fini, art_works, i, project_sizes[i])
            self.tracker_body(tracker_ws, art_works, fini, tracker_schedule, project_sizes[i])

    def tracker_header(self, tracker_ws, titles, fini, art_works, i, project_sizes):
        print('tracker_header and i equals: ' + str(i))
        tracker_ws.cell(row=fini-1, column=1).value = titles[i] + '   ' + '(' + art_works[project_sizes][3] + ')'
        tracker_ws.cell(row=fini-1, column=2).value = 'GK'
        tracker_ws.cell(row=fini-1, column=3).value = 'DTP'
        tracker_ws.cell(row=fini-1, column=4).value = 'Status'
        tracker_ws.cell(row=fini-1, column=5).value = 'Schedule'
        tracker_ws.cell(row=fini-1, column=6).value = 'Notes'

    def tracker_body(self, tracker_ws, art_works, fini, tracker_schedule, project_size):
        print('tracker_body project size is: ' + str(project_size))
        i = project_size
        while art_works[project_size][3] == art_works[i][3]:
            tracker_ws.insert_rows(fini)
            fini = self.get_fini(tracker_ws)
            tracker_ws.cell(row=fini - 1, column=1).value = art_works[i][6] + '   ' + art_works[i][7]
            tracker_ws.cell(row=fini - 1, column=2).value = art_works[i][1]
            tracker_ws.cell(row=fini - 1, column=3).value = art_works[i][2]
            tracker_ws.cell(row=fini - 1, column=5).value = str(tracker_schedule[i])
            i += 1
            if i > len(art_works)-1:
                break
        print('i equals ' + str(i))
        return i

    def API(self):
        os.chdir('C:\\Users\\Feras\\Documents\\Work_Task_Automaiton')
        wb = openpyxl.load_workbook('190605_Frazier3_AU_schedule_48041.xlsm')
        ws = wb.active
        tracker_wb = openpyxl.load_workbook('Version2_Melanies_Project_Tracker_MW - Copy2.xlsx')
        tracker_ws = tracker_wb.active
        art_works = self.ws_art_works(ws)
        project_sizes = self.ws_project_sizes(art_works)
        titles = self.ws_titles(ws)
        schedule = self.ws_schedule(art_works)
        schedule_dates = self.ws_schedule_dates(art_works, ws)
        tracker_schedule = self.ws_tracker_schedule(schedule, schedule_dates)
        self.ws_append(tracker_ws, tracker_schedule, art_works, titles, project_sizes)
        try:
            tracker_wb.save('Version2_Melanies_Project_Tracker_MW - Copy2.xlsx')
        except PermissionError:
            print('Unable to save to the file. Check if it\'s open')


def main():
    obj = WorkTaskAutomation()
    obj.API()


if __name__ == '__main__':
    main()
