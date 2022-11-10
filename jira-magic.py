from __future__ import print_function
from operator import truediv
from typing import SupportsIndex
from jira import JIRA
from atlassian import Jira
import pandas as pd
import time
import glob, argparse
import json
import pygsheets
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import os.path
import numpy as np

from pygsheets.custom_types import ChartType
from datetime import date
from dataclasses import make_dataclass
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from collections import namedtuple


## EXAMPLE COMMAND TO EXECUTE THIS for sprint 8:
## python3 jira-magic.py --sprint 8

space_required_per_sprint = 30

Range = namedtuple('Range', ['start', 'end'])

parser=argparse.ArgumentParser()

parser.add_argument("--sprint", help="The sprint number. This will be used to identify which tickets part of this sprint")

args = parser.parse_args()

sprint_number_from_command_line_args = args.sprint
date_format = '%Y-%m-%d'

first_sprint_number = 5 # first sprint for which we have data

sprint_0_start_date = datetime.date(2022, 6, 29)
sprint_0_end_date = datetime.date(2022, 7, 13)

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


projects = ["INT", "Integrations"]

# Issue : different names in Google Calendar vs Jira ie in Google I am 'Noah Noeh' but in Jira I am 'Noah-Vincenz Noeh'
# second name is Google Calendar name
devs_jira = [ "Andrew Myers", "Arun Bahra", "François Moureau", "Javier Azcurra", "Noah-Vincenz Noeh", "Ufuk Uste", "Vadim Sokolov" ]
qes_jira = [ "Niharika Pillai" ]
devs_google = [ "Andrew Myers", "Arun Bahra", "Francois Moureau", "Javier Azcurra Marilungo", "Noah Noeh", "Ufuk Üste", "Vadim Sokolov" ]
qes_google = [ "Niharika Pillai" ]

team_members_jira = sorted(devs_jira + qes_jira)
team_members_google = sorted(devs_google + qes_google)

devs_jira_str = '("' + '", "'.join(devs_jira) + '")'
qes_jira_str = '("' + '", "'.join(qes_jira) + '")'


server = 'https://yapily.atlassian.net'
email = 'noah-vincenz.noeh@yapily.com'
access_token = 'ohGwgFZJyzJpwUqLUNru742B'

# this uses https://pypi.org/project/jira/ - docs at https://jira.readthedocs.io/
# also see https://pypi.org/project/pygsheets/ and https://pygsheets.readthedocs.io/en/stable/reference.html#models
jira = JIRA(options={'server': server}, 
            basic_auth=(email, access_token))

# generate map of all fields including custom ones, from field name > field id
allfields = jira.fields()
for field in allfields:
    print(field['name']) # to print all ticket fields

nameMap = {field['name']:field['id'] for field in allfields} 

story_points_key = nameMap['Story Points']
current_story_points_key = nameMap['Current Story Points']
final_story_points_key = nameMap['Final Story Points']
issue_type_key = nameMap['Issue Type']

# Create a row object for DataFrames
SprintRow = make_dataclass("SprintRow", [("Name", str), ("Days_Worked", str), ("Story_Points_Started", str), ("Story_Points_Completed", str)])
SummaryRow = make_dataclass("SummaryRow", [("Name", str), ("Avg_Velocity", str), ("Avg_Story_Points_Started", str), ("Avg_Story_Points_Completed", str)])



print('===================================================')

def get_google_name_from_jira_name(jira_name):
    index_in_team_members_list = team_members_jira.index(str(jira_name))
    return team_members_google[index_in_team_members_list]

def get_sprint_start_date(sprint_number):
    days_to_add = 14 * sprint_number + 1
    return sprint_0_start_date + datetime.timedelta(days=days_to_add)

def get_sprint_end_date(sprint_number):
    days_to_add = 14 * sprint_number
    return sprint_0_end_date + datetime.timedelta(days=days_to_add)

def event_is_linked_to_team(event_summary):
    for team_member in team_members_google:
        if team_member in event_summary: 
            return True
    return False


def get_team_out_of_office_events(current_sprint):
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        start_of_sprint_date = get_sprint_start_date(int(current_sprint))
        start_of_sprint_datetime = datetime.datetime.strptime(str(start_of_sprint_date), date_format)
        start_of_sprint_datetime_utc = start_of_sprint_datetime.isoformat() + 'Z' # 'Z' indicates UTC time
        print('Start of sprint in UTC time:', start_of_sprint_datetime_utc)
        personal_calendar_id = 'primary' # or 'noah-vincenz.noeh@yapily.com'
        who_is_out_calendar_id = 'liis4au1bmg4dhnieis2vaa95rv3efte@import.calendar.google.com'
        events_result = service.events().list(calendarId=who_is_out_calendar_id, 
                                              timeMin=start_of_sprint_datetime_utc,
                                              singleEvents=True,
                                              orderBy='startTime').execute()
        print('finished 2')
        events = events_result.get('items', [])

        if not events:
            print('No upcoming events found.')
            return

        print('All Yapily out of office events')
        team_out_of_office_events = [event for event in events if event_is_linked_to_team(event['summary'])]
        return team_out_of_office_events

    except HttpError as error:
        print('An error occurred: %s' % error)



def create_tickets_dict():
    dict = {}
    for team_member in team_members_jira:
        dict[team_member] = {
            "started": 0,
            "completed": 0
        }
    return dict


def get_tickets_dictionaries(dict, current_sprint):
    unassigned_tickets_dict = {}
    print('Calculating data for devs')
    # DEVS
    # issues that are 'in progress' with development
    for singleIssue in jira.search_issues(jql_str='issuetype != Epic AND assignee IN {} AND Sprint = "INT Sprint {}" AND status IN ("SELECTED FOR DEVELOPMENT", "DOING") ORDER BY issuekey'.format(devs_jira_str, current_sprint)):
        ticket_number = singleIssue.key
        print(ticket_number)
        story_points = singleIssue.raw['fields'][story_points_key]
        current_story_points = singleIssue.raw['fields'][current_story_points_key]
        final_story_points = singleIssue.raw['fields'][final_story_points_key]

        if current_story_points != None:
            print('Current Story Points for Ticket {} exist, ticket size has been updated or ticket has been carried over from previous sprint'.format(ticket_number))
            if current_story_points != 0:
                story_points = current_story_points
            else:
                story_points = None
            #     if final_story_points != None:
            #         print('Final Story Points for Ticket {} exist, ticket has been carried over from previous sprint'.format(ticket_number))
            #         story_points = final_story_points
            #     else: 
            #         print('Current Story Points for Ticket {} exist and is 0 but ticket is missing Final Story Points'.format(ticket_number))
        
        print('story points: {}'.format(story_points))
        assignee = singleIssue.fields.assignee.displayName
        print('assignee', assignee)

        # Niharika is QE, for her all tickets that are not in status 'DONE' are in progress
        if assignee != None and story_points != None: # not unassigned
            try:
                current_started = dict[assignee]['started']
                if current_started != None:
                    dict[assignee]['started'] = current_started + story_points
                else:
                    dict[assignee]['started'] = story_points
            except KeyError:
                dict[assignee] = {
                    "started": story_points,
                    "completed": 0
                }
        elif assignee == None:
            if story_points == None:
                unassigned_tickets_dict[ticket_number] = 0
            else:
                unassigned_tickets_dict[ticket_number] = story_points


    # issues that are 'done' with development (not in 'DONE' Jira status)
    for singleIssue in jira.search_issues(jql_str='issuetype != Epic AND assignee IN {} AND Sprint = "INT Sprint {}" AND status IN ("CODE REVIEW", "TESTING", "AWAITING RELEASE", "DONE") ORDER BY issuekey'.format(devs_jira_str, current_sprint)):
        ticket_number = singleIssue.key
        print(ticket_number)
        story_points = singleIssue.raw['fields'][story_points_key]
        current_story_points = singleIssue.raw['fields'][current_story_points_key]
        final_story_points = singleIssue.raw['fields'][final_story_points_key]

        if current_story_points != None:
            print('Current Story Points for Ticket {} exist, ticket size has been updated or ticket has been carried over from previous sprint'.format(ticket_number))
            if current_story_points != 0:
                story_points = current_story_points
            else:
                story_points = None
            #     if final_story_points != None:
            #         print('Final Story Points for Ticket {} exist, ticket has been carried over from previous sprint'.format(ticket_number))
            #         story_points = final_story_points
            #     else: 
            #         print('Current Story Points for Ticket {} exist and is 0 but ticket is missing Final Story Points'.format(ticket_number))

        print('story points: {}'.format(story_points))
        assignee = singleIssue.fields.assignee.displayName
        print('assignee', assignee)
        if assignee != None and story_points != None: # not unassigned
            try:
                current_started = dict[assignee]['started']
                print('current started : {}'.format(current_started))
                if current_started != None:
                    dict[assignee]['started'] = current_started + story_points
                else:
                    dict[assignee]['started'] = story_points
                current_completed = dict[assignee]['completed']
                if current_completed != None:
                    dict[assignee]['completed'] = current_completed + story_points
                else:
                    dict[assignee]['completed'] = story_points
            except KeyError:
                dict[assignee] = {
                    "started": story_points,
                    "completed": story_points
                }
        elif assignee == None:
            if story_points == None:
                unassigned_tickets_dict[ticket_number] = 0
            else:
                unassigned_tickets_dict[ticket_number] = story_points

    print('Calculating data for QA')

    # QA
    # issues that are 'in progress' with qa
    for singleIssue in jira.search_issues(jql_str='issuetype != Epic AND assignee IN {} AND Sprint = "INT Sprint {}" AND status NOT IN ("AWAITING RELEASE", "DONE") ORDER BY issuekey'.format(qes_jira_str, current_sprint)):
        ticket_number = singleIssue.key
        print(ticket_number)
        story_points = singleIssue.raw['fields'][story_points_key]
        current_story_points = singleIssue.raw['fields'][current_story_points_key]
        final_story_points = singleIssue.raw['fields'][final_story_points_key]

        if current_story_points != None:
            print('Current Story Points for Ticket {} exist, ticket size has been updated or ticket has been carried over from previous sprint'.format(ticket_number))
            if current_story_points != 0:
                story_points = current_story_points
            else:
                story_points = None
            #     if final_story_points != None:
            #         print('Final Story Points for Ticket {} exist, ticket has been carried over from previous sprint'.format(ticket_number))
            #         story_points = final_story_points
            #     else: 
            #         print('Current Story Points for Ticket {} exist and is 0 but ticket is missing Final Story Points'.format(ticket_number))

        print('story points: {}'.format(story_points))
        assignee = singleIssue.fields.assignee.displayName
        print('assignee', assignee)

        # Niharika is QE, for her all tickets that are not in status 'DONE' are in progress
        if assignee != None and story_points != None: # not unassigned
            try:
                current_started = dict[assignee]['started']
                if current_started != None:
                    dict[assignee]['started'] = current_started + story_points
                else:
                    dict[assignee]['started'] = story_points
            except KeyError:
                dict[assignee] = {
                    "started": story_points,
                    "completed": 0
                }
        elif assignee == None:
            if story_points == None:
                unassigned_tickets_dict[ticket_number] = 0
            else:
                unassigned_tickets_dict[ticket_number] = story_points

    # issues that are 'done' with qa ('AWAITING RELEASE' and 'DONE' Jira status)
    for singleIssue in jira.search_issues(jql_str='issuetype != Epic AND assignee IN {} AND Sprint = "INT Sprint {}" AND status IN ("AWAITING RELEASE", "DONE") ORDER BY issuekey'.format(qes_jira_str, current_sprint)):
        ticket_number = singleIssue.key
        print(ticket_number)
        story_points = singleIssue.raw['fields'][story_points_key]
        current_story_points = singleIssue.raw['fields'][current_story_points_key]
        final_story_points = singleIssue.raw['fields'][final_story_points_key]

        if current_story_points != None:
            print('Current Story Points for Ticket {} exist, ticket size has been updated or ticket has been carried over from previous sprint'.format(ticket_number))
            if current_story_points != 0:
                story_points = current_story_points
            else:
                story_points = None
            #     if final_story_points != None:
            #         print('Final Story Points for Ticket {} exist, ticket has been carried over from previous sprint'.format(ticket_number))
            #         story_points = final_story_points
            #     else: 
            #         print('Current Story Points for Ticket {} exist and is 0 but ticket is missing Final Story Points'.format(ticket_number))

        print('story points: {}'.format(story_points))
        assignee = singleIssue.fields.assignee.displayName
        print('assignee', assignee)
        if assignee != None and story_points != None: # not unassigned
            try:
                current_started = dict[assignee]['started']
                print('current started : {}'.format(current_started))
                if current_started != None:
                    dict[assignee]['started'] = current_started + story_points
                else:
                    dict[assignee]['started'] = story_points
                current_completed = dict[assignee]['completed']
                if current_completed != None:
                    dict[assignee]['completed'] = current_completed + story_points
                else:
                    dict[assignee]['completed'] = story_points
            except KeyError:
                dict[assignee] = {
                    "started": story_points,
                    "completed": story_points
                }
        elif assignee == None:
            if story_points == None:
                unassigned_tickets_dict[ticket_number] = 0
            else:
                unassigned_tickets_dict[ticket_number] = story_points
    return [dict, unassigned_tickets_dict]

def get_days_worked_by_name(team_out_of_office_events, sprint_start_date, sprint_end_date, google_name):
    sprint_end_date = sprint_end_date + datetime.timedelta(days=1) # adding 1 day to sprint end because for OOO days we say end day is when an individual is no longer OOO
    days_worked = 10
    individual_out_of_office_events = [event for event in team_out_of_office_events if str(google_name) in event['summary']]
    for out_of_office_event in individual_out_of_office_events:
        start_ooo_str = out_of_office_event['start'].get('dateTime', out_of_office_event['start'].get('date'))
        end_ooo_str = out_of_office_event['end'].get('dateTime', out_of_office_event['end'].get('date'))
        print('Start OOO -', start_ooo_str)
        print('End OOO -', end_ooo_str)
        # get overlap
        start_ooo = datetime.datetime.strptime(start_ooo_str, date_format).date()
        end_ooo = datetime.datetime.strptime(end_ooo_str, date_format).date()
        r1 = Range(start=sprint_start_date, end=sprint_end_date)
        r2 = Range(start=start_ooo, end=end_ooo)
        latest_start = max(r1.start, r2.start)
        earliest_end = min(r1.end, r2.end)
        delta = (earliest_end - latest_start).days
        overlap = max(0, delta)
        print('overlap:', overlap)
        if overlap > 0:
            business_days = np.busday_count(latest_start, earliest_end)
            print('business days:', business_days)
            overlap = business_days
        days_worked = days_worked - overlap
    return days_worked


def fill_sprint_sheet_data(data_frame, worksheet, current_sprint, sprint_start_date, sprint_end_date):
    print('sprint start date:', sprint_start_date)
    print('sprint end date:', sprint_end_date)
        
    team_out_of_office_events = get_team_out_of_office_events(current_sprint)
    # filter all OOO to reduce to the ones during this sprint, ie the OOO start date is on or before the end date of the sprint
    team_out_of_office_events_during_sprint = [event for event in team_out_of_office_events if datetime.datetime.strptime(event['start'].get('dateTime', event['start'].get('date')), date_format).date() <= sprint_end_date]
    print('Out Of Office Events')
    for event in team_out_of_office_events_during_sprint:
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date'))
        print(start, event['summary'], end)

    tickets_dict = create_tickets_dict()

    [assigned_tickets_dict, unassigned_tickets_dict] = get_tickets_dictionaries(tickets_dict, current_sprint)

    print('Assigned Tickets')
    for key, value in assigned_tickets_dict.items():
        print(key)
        for key, value in value.items():
            print(key, ' : ', value)
        print('\n')

    print('Unassigned Tickets')
    for key, value in unassigned_tickets_dict.items():
        print('Ticket:', key, 'Points:', value)


    print(assigned_tickets_dict.items())
    for name, value in assigned_tickets_dict.items():
        print(name)
        print(value.items())
        ((_, started_value), (_, completed_value)) = value.items()
        google_name = get_google_name_from_jira_name(name)
        individual_days_worked = get_days_worked_by_name(team_out_of_office_events_during_sprint, sprint_start_date, sprint_end_date, google_name)
        df2 = pd.DataFrame([SprintRow(str(name), individual_days_worked, started_value, completed_value)])
        data_frame = pd.concat([data_frame, df2])

    #update the first sheet with df, starting in column C (= 3) and row Sprint x space_required_per_sprint, because space_required_per_sprint rows is the space required for each sprint's data.
    sprint_top_row_index = int(current_sprint) * space_required_per_sprint
    sprint_header_index = sprint_top_row_index - 3
    # for sprint length w/o weekend days
    business_days = np.busday_count(sprint_start_date, sprint_end_date) + 1
    # set sprint header
    worksheet.cell('C' + str(sprint_header_index)).set_text_format('bold', True).value = 'Sprint ' + str(current_sprint)
    worksheet.cell('C' + str(sprint_header_index + 1)).value = 'Date'
    worksheet.cell('D' + str(sprint_header_index + 1)).value = sprint_start_date.strftime("%d/%m/%Y") + '  -  ' + sprint_end_date.strftime("%d/%m/%Y")
    worksheet.cell('C' + str(sprint_header_index + 2)).value = 'Days in Sprint'
    worksheet.cell('D' + str(sprint_header_index + 2)).value = str(business_days)
    # set table headers to be bold font
    worksheet.cell('C' + str(sprint_top_row_index)).set_text_format('bold', True)
    worksheet.cell('D' + str(sprint_top_row_index)).set_text_format('bold', True)
    worksheet.cell('E' + str(sprint_top_row_index)).set_text_format('bold', True)
    worksheet.cell('F' + str(sprint_top_row_index)).set_text_format('bold', True)
    # get ranges needed for calculations
    sprint_first_data_row_index = sprint_top_row_index + 1
    sprint_last_data_row_index = sprint_first_data_row_index + len(team_members_jira) - 1 # number of team members - 1
    sprint_footer_data_row_index = sprint_last_data_row_index + 1
    days_worked_range_str = 'D' + str(sprint_first_data_row_index) + ':D' + str(sprint_last_data_row_index)
    story_points_started_range_str = 'E' + str(sprint_first_data_row_index) + ':E' + str(sprint_last_data_row_index)
    story_points_completed_range_str = 'F' + str(sprint_first_data_row_index) + ':F' + str(sprint_last_data_row_index)
    # add footers
    worksheet.cell('C' + str(sprint_footer_data_row_index + 1)).set_text_format('bold', True).value = 'MIN'
    worksheet.cell('C' + str(sprint_footer_data_row_index + 2)).set_text_format('bold', True).value = 'MAX'
    worksheet.cell('C' + str(sprint_footer_data_row_index + 3)).set_text_format('bold', True).value = 'AVG'
    worksheet.cell('C' + str(sprint_footer_data_row_index + 4)).set_text_format('bold', True).value = 'SUM'
    # set the min values
    worksheet.update_value('D' + str(sprint_footer_data_row_index + 1), '=min(' + days_worked_range_str + ')')
    worksheet.update_value('E' + str(sprint_footer_data_row_index + 1), '=min(' + story_points_started_range_str + ')')
    worksheet.update_value('F' + str(sprint_footer_data_row_index + 1), '=min(' + story_points_completed_range_str + ')')
    # set the max values
    worksheet.update_value('D' + str(sprint_footer_data_row_index + 2), '=max(' + days_worked_range_str + ')')
    worksheet.update_value('E' + str(sprint_footer_data_row_index + 2), '=max(' + story_points_started_range_str + ')')
    worksheet.update_value('F' + str(sprint_footer_data_row_index + 2), '=max(' + story_points_completed_range_str + ')')
    # set the avg values
    worksheet.update_value('D' + str(sprint_footer_data_row_index + 3), '=average(' + days_worked_range_str + ')')
    worksheet.update_value('E' + str(sprint_footer_data_row_index + 3), '=average(' + story_points_started_range_str + ')')
    worksheet.update_value('F' + str(sprint_footer_data_row_index + 3), '=average(' + story_points_completed_range_str + ')')
    # set the sum values
    worksheet.update_value('D' + str(sprint_footer_data_row_index + 4), '=sum(' + days_worked_range_str + ')')
    worksheet.update_value('E' + str(sprint_footer_data_row_index + 4), '=sum(' + story_points_started_range_str + ')')
    worksheet.update_value('F' + str(sprint_footer_data_row_index + 4), '=sum(' + story_points_completed_range_str + ')')
    
    worksheet.set_dataframe(data_frame, (sprint_top_row_index, 3))

    current_sprint_chart = worksheet.add_chart(domain=('C' + str(sprint_first_data_row_index), 'C' + str(sprint_last_data_row_index)), 
                            ranges=[('F' + str(sprint_first_data_row_index), 'F' + str(sprint_last_data_row_index))], 
                            title='Story Points Completed Sprint ' + str(current_sprint),
                            chart_type=ChartType.BAR, # taken from : BAR, LINE, AREA, COLUMN, SCATTER, COMBO, STEPPED_AREA
                            anchor_cell='H' + str(sprint_top_row_index))

    print('sprint start date:', sprint_start_date)
    print('sprint end date:', sprint_end_date)

def compute_velocities(worksheet, number_of_sprints_computed, index):
    # each sprint section is space_required_per_sprint spaces long

    # TODO: UPDATE TO USE SHEETS FORMULA INSTEAD OF ACTUAL VALUES

    first_sprint_data_row = index + 1 + first_sprint_number * space_required_per_sprint
    max_sprint_data_row = first_sprint_data_row + number_of_sprints_computed * space_required_per_sprint
    print(max_sprint_data_row)
    sum_of_days_worked = 0
    sum_of_started_points = 0
    sum_of_completed_points = 0
    while first_sprint_data_row < max_sprint_data_row:
        print('in data row: ', first_sprint_data_row)
        print('looking at cell D' + str(first_sprint_data_row))
        print(worksheet.cell('D' + str(first_sprint_data_row)))
        print(worksheet.cell('D' + str(first_sprint_data_row)).value)
        sum_of_days_worked = float(sum_of_days_worked) + float(worksheet.cell('D' + str(first_sprint_data_row)).value)
        sum_of_started_points = float(sum_of_started_points) + float(worksheet.cell('E' + str(first_sprint_data_row)).value)
        sum_of_completed_points = float(sum_of_completed_points) + float(worksheet.cell('F' + str(first_sprint_data_row)).value)
        first_sprint_data_row = first_sprint_data_row + space_required_per_sprint
    return (sum_of_days_worked, sum_of_started_points, sum_of_completed_points)


def draw_overall_chart():
    # all_sprints_line_chart = wks.add_chart(domain=('C' + str(sprint_first_data_row_index), 'C' + str(sprint_last_data_row_index)), 
    #                 ranges=[('F' + str(sprint_first_data_row_index), 'F' + str(sprint_last_data_row_index))], 
    #                 title='Story Points Completed Per Sprint',
    #                 chart_type=ChartType.LINE, # taken from : BAR, LINE, AREA, COLUMN, SCATTER, COMBO, STEPPED_AREA
    #                 anchor_cell='B5')
    return




def main():
    #authorise
    gc = pygsheets.authorize(service_file='/Users/noah-vincenz.noeh/Downloads/my-project-365311-76ce02b8b519.json')
    # Create empty dataframe
    df = pd.DataFrame()
    #open the google spreadsheet (where 'PY to Gsheet Test' is the name of my sheet)
    sh = gc.open('Integrations Team')
    #select the first sheet 
    wks = sh[0]
    wks.title = "Velocity Tracker"
    wks.adjust_column_width(start=3, end=6, pixel_size=200)
    existing_charts = wks.get_charts()
    for chart in existing_charts:
        chart.delete()
    wks.clear()
    number_of_sprints_computed = int(sprint_number_from_command_line_args) - first_sprint_number + 1
    sum_of_days_per_sprint = 0
    for current_sprint in range(first_sprint_number, int(sprint_number_from_command_line_args) + 1):
        sprint_start_date = get_sprint_start_date(int(current_sprint))
        sprint_end_date = get_sprint_end_date(int(current_sprint))
        business_days_in_sprint = np.busday_count(sprint_start_date, sprint_end_date) + 1
        sum_of_days_per_sprint = sum_of_days_per_sprint + business_days_in_sprint
        fill_sprint_sheet_data(df, wks, current_sprint, sprint_start_date, sprint_end_date)
    index = 0
    for team_member in team_members_jira:
        (sum_of_days_worked, sum_of_started_points, sum_of_completed_points) = compute_velocities(wks, number_of_sprints_computed, index)
        print('computed data for ', team_member)
        print(sum_of_days_worked, sum_of_started_points, sum_of_completed_points)
        avg_velocity = sum_of_completed_points / sum_of_days_worked * sum_of_days_per_sprint / number_of_sprints_computed
        df2 = pd.DataFrame([SummaryRow(str(team_member), avg_velocity, sum_of_started_points / number_of_sprints_computed, sum_of_completed_points / number_of_sprints_computed)])
        df = pd.concat([df, df2])
        index = index + 1
    wks.set_dataframe(df, (3, 3)) # inserting summary table at (3,3)
    draw_overall_chart()


if __name__ == '__main__':
    main()



