import os
import warnings
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Font, PatternFill, Alignment, Border, Side


class WorkTaskAutomation(object):
    titles = []
    art_works = []
    schedule = []
    schedule_dates = []

    def __init__(self):
        pass

    def ws_titles(self, ws):
        for row in ws:
            for cell in row:
                if cell.column is 7 and cell.value is None:
                    cell = cell.offset(0, 1)
                    if cell.value is not None:
                        self.titles.append(cell.value)

    def ws_art_works(self, ws):
        for row in ws.values:
            for value in row:
                if value == 'Melanie':
                    self.art_works.append(list(row))
        for i in range(len(self.art_works)):
            for j in range(len(self.art_works[i])):
                if type(self.art_works[i][j]) is not str:
                    self.art_works[i][j] = ''

    def ws_project_starting_indices(self):
        project_starting_indices = [0]
        j = 0
        for i in range(len(self.art_works)):
            if self.art_works[j][3] != self.art_works[i][3]:
                project_starting_indices.append(i)
                j = i
        return project_starting_indices

    def ws_schedule(self):
        for i in range(len(self.art_works)):
            try:
                start = [i for i, s in enumerate(self.art_works[i]) if 'DV' in s or 'PM ' in s]
                end = [i for i, s in enumerate(self.art_works[i]) if 'ICR' in s or 'ICA' in s]
                self.schedule.append(self.art_works[i][start[0]:end[0]+1])
            except ValueError:
                print('project might not contain DV, PM, ICR or ICA tasks')

    def ws_schedule_dates(self, ws):
        dates = []
        for row in ws.values:
            for value in row:
                dates.append(value)
        for i in range(len(self.art_works)):
            try:
                start = [i for i, s in enumerate(self.art_works[i]) if 'DV' in s or 'PM ' in s]
                end = [i for i, s in enumerate(self.art_works[i]) if 'ICR' in s or 'ICA' in s]
                self.schedule_dates.append(dates[start[0]:end[0]+1])
            except ValueError:
                pass

    def ws_tracker_schedule(self):
        tracker_schedule = []
        for i in range(len(self.schedule_dates)):
            cell_schedule = []
            for j in range(len(self.schedule_dates[i])):
                try:
                    schedule_str = " " + str(self.schedule_dates[i][j].month) + "/" \
                                   + str(self.schedule_dates[i][j].day)+ " " + str(self.schedule[i][j])
                    if (self.schedule[i][j]) != '':
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

    def ws_append(self, tracker_ws, tracker_schedule, project_starting_indices, tracker_wb):
        fini = self.get_fini(tracker_ws)
        try:
            tracker_ws.unmerge_cells(start_row=fini, start_column=1, end_row=fini, end_column=6)
        except ValueError:
            print('cells are not merged')
        thin = Side(border_style="thin", color="000000")
        ft = Font(name='Calibri', size=16)
        top_style = NamedStyle(name="top_style", fill=PatternFill(patternType='solid', fill_type='solid', fgColor=
            "c5d9f1"), font=ft, alignment=Alignment(horizontal="center", vertical="center"),
                               border=Border(top=thin, left=thin, right=thin, bottom=thin))
        body1 = NamedStyle(name="body1", font=ft, border=Border(top=thin, left=thin, right=
            thin, bottom=thin), alignment=Alignment(horizontal="left", vertical="top"))
        try:
            tracker_wb.add_named_style(top_style)
            tracker_wb.add_named_style(body1)
        except ValueError:
            pass
        for i in range(len(self.titles)):
            fini = self.get_fini(tracker_ws)
            tracker_ws.insert_rows(fini)
            try:
                self.tracker_header(tracker_ws, i, project_starting_indices[i])
                self.tracker_body(tracker_ws, tracker_schedule, project_starting_indices[i])
            except IndexError:
                print('There was an issue with the arrays that caused an error')

    def tracker_header(self, tracker_ws, i, project_starting_indices):
        fini = self.get_fini(tracker_ws)
        tracker_ws.row_dimensions[fini-1].height = 100
        cell1 = tracker_ws.cell(row=fini-1, column=1)
        cell2 = tracker_ws.cell(row=fini-1, column=2)
        cell3 = tracker_ws.cell(row=fini-1, column=3)
        cell4 = tracker_ws.cell(row=fini-1, column=4)
        cell5 = tracker_ws.cell(row=fini-1, column=5)
        cell6 = tracker_ws.cell(row=fini-1, column=6)

        cell1.value = str(self.titles[i]) + '   ' + '(' + str(self.art_works[project_starting_indices][3]) + ')'
        cell2.value = 'GK'
        cell3.value = 'DTP'
        cell4.value = 'Status'
        cell5.value = 'Schedule'
        cell6.value = 'Notes'
        cell1.style, cell2.style, cell3.style, cell4.style, cell5.style, cell6.style = 'top_style', 'top_style',\
                                                                                       'top_style', 'top_style',\
                                                                                       'top_style', 'top_style'

    def tracker_body(self, tracker_ws, tracker_schedule, project_starting_indices):
        i = project_starting_indices
        fini = self.get_fini(tracker_ws)
        while self.art_works[project_starting_indices][3] == self.art_works[i][3]:
            tracker_ws.insert_rows(fini)
            fini = self.get_fini(tracker_ws)
            tracker_ws.row_dimensions[fini - 1].height = 70
            cell1 = tracker_ws.cell(row=fini - 1, column=1)
            cell2 = tracker_ws.cell(row=fini - 1, column=2)
            cell3 = tracker_ws.cell(row=fini - 1, column=3)
            cell5 = tracker_ws.cell(row=fini - 1, column=5)
            cell1.value = self.art_works[i][6] + '   ' + self.art_works[i][7]
            cell2.value = self.art_works[i][1]
            cell3.value = self.art_works[i][2]
            cell5.value = ','.join(tracker_schedule[i])
            cell1.style, cell2.style, cell3.style, cell5.style = 'body1', 'body1', 'body1', 'body1'
            i += 1
            if i > len(self.art_works)-1:
                break

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
            input('\nPress any key to exit the program')
        tracker_ws = tracker_wb.active
        self.ws_art_works(ws)
        project_starting_indices = self.ws_project_starting_indices()
        self.ws_titles(ws)
        self.ws_schedule()
        self.ws_schedule_dates(ws)
        tracker_schedule = self.ws_tracker_schedule()
        self.ws_append(tracker_ws, tracker_schedule, project_starting_indices, tracker_wb)
        fini = self.get_fini(tracker_ws)
        tracker_ws.merge_cells(start_row=fini, start_column=1, end_row=fini, end_column=6)
        tracker_file2 = input('Enter the tracker file to load: ')
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
