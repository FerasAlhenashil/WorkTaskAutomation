import os
import warnings
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Font, PatternFill, Alignment, Border, Side


class WorkTaskAutomation(object):
    def __init__(self):
        pass

    def ws_titles(self, ws):
        titles = []
        for row in ws:
            for cell in row:
                if cell.column is 7 and cell.value is None:
                    cell = cell.offset(0, 1)
                    if cell.value is not None:
                        titles.append(cell.value)
        return titles

    def ws_art_works(self, ws):
        art_works = []
        for row in ws.values:
            for value in row:
                if value == 'Melanie':
                    art_works.append(row)
        return art_works

    def ws_project_starting_index(self, art_works):
        project_starting_index = [0]
        j = 0
        for i in range(len(art_works)):
            if art_works[j][3] != art_works[i][3]:
                project_starting_index.append(i)
                j = i
        return project_starting_index

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
            try:
                start = job.index('PM prep')
                return start
            except ValueError:
                print(job[7], 'Error: has no DV Shell due nor PM Prep/prep tasks')

    def ws_tracker_schedule(self, schedule, schedule_dates):
        tracker_schedule = []
        for i in range(len(schedule_dates)):
            cell_schedule = []
            for j in range(len(schedule_dates[i])):
                try:
                    schedule_str = " " + str(schedule_dates[i][j].month) + "/" + str(schedule_dates[i][j].day) + " " + str(
                        schedule[i][j])
                    if (schedule[i][j]) is not None:
                        cell_schedule.append(schedule_str)
                except AttributeError:
                    print('An error with schedule dates')
            tracker_schedule.append(cell_schedule)
        return tracker_schedule

    def get_fini(self, tracker_ws):
        fini = tracker_ws.max_row
        while tracker_ws.cell(row=fini, column=1).value is None:
            fini -= 1
        return fini

    def ws_append(self, tracker_ws, tracker_schedule, art_works, titles, project_starting_index, tracker_wb):
        fini = self.get_fini(tracker_ws)
        try:
            tracker_ws.unmerge_cells(start_row=fini, start_column=1, end_row=fini, end_column=6)
        except ValueError:
            print('cells are not merged')
        thin = Side(border_style="thin", color="000000")
        top_style = NamedStyle(name="top_style", fill=PatternFill(patternType='solid', fill_type='solid', fgColor=
            "c5d9f1"), font=Font(name='Calibri', size=16), alignment=Alignment(horizontal="center", vertical="center"),
                               border=Border(top=thin, left=thin, right=thin, bottom=thin))
        body1 = NamedStyle(name="body1", font=Font(name="Calibri", size=16), border=Border(top=thin, left=thin, right=
            thin, bottom=thin), alignment=Alignment(horizontal="left", vertical="top"))
        try:
            tracker_wb.add_named_style(top_style)
            tracker_wb.add_named_style(body1)
        except ValueError:
            pass
        for i in range(len(titles)):
            fini = self.get_fini(tracker_ws)
            tracker_ws.insert_rows(fini)
            try:
                self.tracker_header(tracker_ws, titles, art_works, i, project_starting_index[i])
                self.tracker_body(tracker_ws, art_works, tracker_schedule, project_starting_index[i])
            except IndexError:
                print('There was an issue with the arrays that caused an error')

    def tracker_header(self, tracker_ws, titles, art_works, i, project_starting_index):
        fini = self.get_fini(tracker_ws)
        tracker_ws.row_dimensions[fini-1].height = 100
        cell1 = tracker_ws.cell(row=fini-1, column=1)
        cell2 = tracker_ws.cell(row=fini-1, column=2)
        cell3 = tracker_ws.cell(row=fini-1, column=3)
        cell4 = tracker_ws.cell(row=fini-1, column=4)
        cell5 = tracker_ws.cell(row=fini-1, column=5)
        cell6 = tracker_ws.cell(row=fini-1, column=6)

        cell1.value = str(titles[i]) + '   ' + '(' + str(art_works[project_starting_index][3]) + ')'
        cell2.value = 'GK'
        cell3.value = 'DTP'
        cell4.value = 'Status'
        cell5.value = 'Schedule'
        cell6.value = 'Notes'
        cell1.style = 'top_style'
        cell2.style = 'top_style'
        cell3.style = 'top_style'
        cell4.style = 'top_style'
        cell5.style = 'top_style'
        cell6.style = 'top_style'

    def tracker_body(self, tracker_ws, art_works, tracker_schedule, project_starting_index):
        i = project_starting_index
        fini = self.get_fini(tracker_ws)
        while art_works[project_starting_index][3] == art_works[i][3]:
            tracker_ws.insert_rows(fini)
            fini = self.get_fini(tracker_ws)
            tracker_ws.row_dimensions[fini - 1].height = 70
            cell1 = tracker_ws.cell(row=fini - 1, column=1)
            cell2 = tracker_ws.cell(row=fini - 1, column=2)
            cell3 = tracker_ws.cell(row=fini - 1, column=3)
            cell5 = tracker_ws.cell(row=fini - 1, column=5)
            cell1.value = art_works[i][6] + '   ' + art_works[i][7]
            cell2.value = art_works[i][1]
            cell3.value = art_works[i][2]
            cell5.value = ','.join(tracker_schedule[i])
            cell1.style = 'body1'
            cell2.style = 'body1'
            cell3.style = 'body1'
            cell5.style = 'body1'
            i += 1
            if i > len(art_works)-1:
                break
        return i

    def API(self):
        #'R:\\LifeScan\\_Lifescan_General\\20_Personal_Folders\\Melanie\\Schedules'
        os.chdir('C:\\Users\\Feras\\Documents\\Work_Task_Automaiton')
        schedule_file = input('Enter the schedule file name: ')
        schedule_file = schedule_file + '.xlsm'
        try:
            wb = load_workbook(schedule_file)
        except FileNotFoundError:
            print('Unable to find the schedule file. Make sure it exists at the expected location')
            quit()
        ws = wb.active
        tracker_file = input('Enter the tracker file to load: ')
        tracker_file = tracker_file + '.xlsx'
        try:
            tracker_wb = load_workbook(tracker_file)
        except FileNotFoundError:
            print('Unable to find the tracker file. Make sure it exists at the expected location')
            quit()
        tracker_ws = tracker_wb.active
        art_works = self.ws_art_works(ws)
        project_starting_index = self.ws_project_starting_index(art_works)
        titles = self.ws_titles(ws)
        schedule = self.ws_schedule(art_works)
        schedule_dates = self.ws_schedule_dates(art_works, ws)
        tracker_schedule = self.ws_tracker_schedule(schedule, schedule_dates)
        self.ws_append(tracker_ws, tracker_schedule, art_works, titles, project_starting_index, tracker_wb)
        fini = self.get_fini(tracker_ws)
        tracker_ws.merge_cells(start_row=fini, start_column=1, end_row=fini, end_column=6)
        tracker_file2 = input('Enter the tracker file to save: ')
        tracker_file2 = tracker_file2 + '.xlsx'
        try:
            tracker_wb.save(tracker_file2)
        except PermissionError:
            print('\nUnable to save to the tracker. Please make sure the file isn\'t open')
            input('\nPress any key to exit the program')


def main():
    obj = WorkTaskAutomation()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        obj.API()


if __name__ == '__main__':
    main()
