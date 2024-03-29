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
import requests
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
## python3 jira-magic.py --sprints 8 --max_sprint 8
## python3 jira-magic.py --sprints 5,6,7,8 --max_sprint 8
## python3 jira-magic.py --sprints 6 --max_sprint 8

space_required_per_sprint = 30

Range = namedtuple('Range', ['start', 'end'])

parser=argparse.ArgumentParser()

parser.add_argument("--sprints", help="The list of sprint numbers. You can pass in any arbitrary current or past sprint number here or multiple sprint numbers separated with a comma. This will be used to identify which tickets part of this sprint", default=[])
parser.add_argument("--max_sprint", help="The maximum sprint to be present in the spreadsheet. This should either be the highest existing sprint number in the sheet when creating data for old sprints or the highest spring number that is being passed in as part of the --sprints argument. This is necessary to calculate all the sums in the summary table at the top of the sheet", default=[])

args = parser.parse_args()

sprint_numbers_from_command_line_args = args.sprints.split(',')
max_sprint_from_command_line_args = args.max_sprint
date_format = '%Y-%m-%d'

first_sprint_number = 5 # first sprint for which we have data

sprint_0_start_date = datetime.date(2022, 6, 29)
sprint_0_end_date = datetime.date(2022, 7, 13)

gov_uk_bank_holidays_url = 'https://www.gov.uk/bank-holidays.json'
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


projects = ["INT", "Integrations"]

# Issue : different names in Google Calendar vs Jira ie in Google it might be 'Clare Catch' but in Jira it might be 'Cläre Catch'
# second name is Google Calendar name
devs_jira = [ "Max Mann", "Andrew Aia", "Bernd Bull", "Jack Julee", "Noah Norries", "Cläre Catch" ]
qes_jira = [ "Hugh Quality" ]
devs_google = [ "Max Mann", "Andrew Aia", "Bernd Bull", "Jack Julee", "Noah Norries", "Clare Catch" ]
qes_google = [ "Hugh Quality" ]

team_members_jira = sorted(devs_jira + qes_jira)
team_members_google = sorted(devs_google + qes_google)

devs_jira_str = '("' + '", "'.join(devs_jira) + '")'
qes_jira_str = '("' + '", "'.join(qes_jira) + '")'

start_row_of_summary_table = 3
end_row_of_summary_table = start_row_of_summary_table + len(team_members_jira)
start_column_of_summary_table = 3

squad_velocity_chart_name = 'Squad Velocity Per Sprint'

server = 'https://yapily.atlassian.net'
email = '....com'
access_token = '...'

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
sprint_key = nameMap['Sprint']
status_key = nameMap['Status']

# Create a row object for DataFrames
SprintRow = make_dataclass("SprintRow", [("Name", str), ("Days_Worked", str), ("Story_Points_Started", str), ("Story_Points_Completed", str)])
SummaryRow = make_dataclass("SummaryRow", [("Name", str), ("Avg_Velocity", str), ("Avg_Story_Points_Started", str), ("Avg_Story_Points_Completed", str)])


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
        personal_calendar_id = 'primary' # or 'noah-vincenz.noeh@yapily.com'
        who_is_out_calendar_id = '...@import.calendar.google.com'
        events_result = service.events().list(calendarId=who_is_out_calendar_id,
                                              timeMin=start_of_sprint_datetime_utc,
                                              singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            print('No upcoming events found.')
            return
        team_out_of_office_events = [event for event in events if event_is_linked_to_team(event['summary'])]
        return team_out_of_office_events

    except HttpError as error:
        print('An error occurred: %s' % error)



def create_tickets_dict():
    dict = {}
    for team_member in team_members_jira:
        dict[team_member] = {
            "started": {
                "sum": 0,
                "tickets": []
            },
            "completed": {
                "sum": 0,
                "tickets": []
            }
        }
    return dict


def sprint_list_contains_next_sprint_but_not_previous(sprint_list, current_sprint):
    sprint_list_contains_next_sprint_but_not_previous = False
    for sprint_obj in sprint_list:
        if (str(current_sprint - 1) in sprint_obj['name']):
            return False
        if (str(current_sprint + 1) in sprint_obj['name']):
            sprint_list_contains_next_sprint_but_not_previous = True
    return sprint_list_contains_next_sprint_but_not_previous


def remove_duplicates_and_sort_list(dict):
    # for name, dict_obj in dict.items():
    #     for started_completed_key, value in dict_obj.items():
    #         for sum_tickets_key, value in value.items():
    #             print(key, ' : ', value)
    return


def get_tickets_dictionaries(dict, current_sprint):
    unassigned_tickets_dict = {}
    print('\n\nCalculating data for devs')
    # DEVS
    # issues that are 'in progress' with development
    for singleIssue in jira.search_issues(jql_str='issuetype != Epic AND assignee IN {} AND Sprint = "INT Sprint {}" AND status IN ("SELECTED FOR DEVELOPMENT", "IN PROGRESS", "WAITING FOR BANK") ORDER BY issuekey'.format(devs_jira_str, current_sprint)):
        ticket_number = singleIssue.key
        print('---')
        print(ticket_number)
        story_points = singleIssue.raw['fields'][story_points_key]
        current_story_points = singleIssue.raw['fields'][current_story_points_key]
        final_story_points = singleIssue.raw['fields'][final_story_points_key]
        status = singleIssue.raw['fields'][status_key]['name']

        if current_story_points != None:
            print('Current Story Points for Ticket {} exist, ticket size has been updated or ticket has been carried over from previous sprint'.format(ticket_number))
            sprints = singleIssue.raw['fields'][sprint_key]
            if (len(sprints) > 1): # ticket has been carried over
                print('Ticket has been carried across sprints')
                # TODO: and same below
                # if (sprint_list_contains_next_sprint(sprints, int(current_sprint))):
                #     print('Ticket has been carried into next sprint, we are adding {} started points for this ticket'.format(str(story_points)))
                #     current_story_points = current_story_points + story_points
            if current_story_points != 0:
                story_points = current_story_points
            # else:
            elif status != 'Waiting for Bank':
                story_points = None

        print('story points: {}'.format(story_points))
        assignee = singleIssue.fields.assignee.displayName
        print('assignee', assignee)

        if assignee != None and story_points != None: # not unassigned
            current_started = dict[assignee]['started']['sum']
            current_tickets = dict[assignee]['started']['tickets']
            if current_started != None:
                dict[assignee]['started']['sum'] = current_started + story_points
                current_tickets.append(ticket_number)
                dict[assignee]['started']['tickets'] = current_tickets
            else:
                dict[assignee]['started']['sum'] = story_points
                dict[assignee]['started']['tickets'] = [ticket_number]
        elif assignee == None:
            if story_points == None:
                unassigned_tickets_dict[ticket_number] = 0
            else:
                unassigned_tickets_dict[ticket_number] = story_points


    # issues that are 'done' with development (not in 'DONE' Jira status)
    for singleIssue in jira.search_issues(jql_str='issuetype != Epic AND assignee IN {} AND Sprint = "INT Sprint {}" AND status IN ("CODE REVIEW", "TESTING", "AWAITING RELEASE", "DONE", "WAITING FOR BANK") ORDER BY issuekey'.format(devs_jira_str, current_sprint)):
        ticket_number = singleIssue.key
        print('---')
        print(ticket_number)
        story_points = singleIssue.raw['fields'][story_points_key]
        current_story_points = singleIssue.raw['fields'][current_story_points_key]
        final_story_points = singleIssue.raw['fields'][final_story_points_key]
        status = singleIssue.raw['fields'][status_key]['name']

        if current_story_points != None:
            print('Current Story Points for Ticket {} exist, ticket size has been updated or ticket has been carried over from previous sprint'.format(ticket_number))
            sprints = singleIssue.raw['fields'][sprint_key]
            if (len(sprints) > 1): # ticket has been carried over
                print('Ticket has been carried across sprints')
                if (sprint_list_contains_next_sprint_but_not_previous(sprints, int(current_sprint))):
                    print('Ticket has been carried into next sprint, we are adding {} completed points for this ticket'.format(str(story_points - current_story_points))) # can add story points - current story points
                    current_story_points = story_points - current_story_points
            if current_story_points != 0:
                story_points = current_story_points
            else:
            # elif status != 'Waiting for Bank':
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
            current_started = dict[assignee]['started']['sum']
            current_tickets_started = dict[assignee]['started']['tickets']
            if (ticket_number not in current_tickets_started):
                print('current started : {}'.format(current_started))
                if current_started != None:
                    dict[assignee]['started']['sum'] = current_started + story_points
                    current_tickets_started.append(ticket_number)
                    dict[assignee]['started']['tickets'] = current_tickets_started
                else:
                    dict[assignee]['started']['sum'] = story_points
                    dict[assignee]['started']['tickets'] = [ticket_number]
            current_completed = dict[assignee]['completed']['sum']
            current_tickets_completed = dict[assignee]['completed']['tickets']
            if current_completed != None:
                dict[assignee]['completed']['sum'] = current_completed + story_points
                current_tickets_completed.append(ticket_number)
                dict[assignee]['completed']['tickets'] = current_tickets_completed
            else:
                dict[assignee]['completed']['sum'] = story_points
                dict[assignee]['completed']['tickets'] = [ticket_number]
        elif assignee == None:
            if story_points == None:
                unassigned_tickets_dict[ticket_number] = 0
            else:
                unassigned_tickets_dict[ticket_number] = story_points
        elif story_points == None and sprint_list_contains_next_sprint_but_not_previous(sprints, int(current_sprint)):
            # ticket has been carried over and 0 points were completed this sprint - we need to add points to started points
            story_points = singleIssue.raw['fields'][story_points_key]
            current_started = dict[assignee]['started']['sum']
            current_tickets_started = dict[assignee]['started']['tickets']
            if (ticket_number not in current_tickets_started):
                print('current started : {}'.format(current_started))
                if current_started != None:
                    dict[assignee]['started']['sum'] = current_started + story_points
                    current_tickets_started.append(ticket_number)
                    dict[assignee]['started']['tickets'] = current_tickets_started
                else:
                    dict[assignee]['started']['sum'] = story_points
                    dict[assignee]['started']['tickets'] = [ticket_number]

    print('\n\nCalculating data for QA')

    # QA
    # issues that are 'in progress' with qa
    for singleIssue in jira.search_issues(jql_str='issuetype != Epic AND assignee IN {} AND Sprint = "INT Sprint {}" AND status NOT IN ("AWAITING RELEASE", "DONE") ORDER BY issuekey'.format(qes_jira_str, current_sprint)):
        ticket_number = singleIssue.key
        print('---')
        print(ticket_number)
        story_points = singleIssue.raw['fields'][story_points_key]
        current_story_points = singleIssue.raw['fields'][current_story_points_key]
        final_story_points = singleIssue.raw['fields'][final_story_points_key]

        if current_story_points != None:
            print('Current Story Points for Ticket {} exist, ticket size has been updated or ticket has been carried over from previous sprint'.format(ticket_number))
            sprints = singleIssue.raw['fields'][sprint_key]
            if (len(sprints) > 1): # ticket has been carried over
                print('Ticket has been carried across sprints')
                # TODO: and same above
                # if (sprint_list_contains_next_sprint(sprints, int(current_sprint))):
                #     print('Ticket has been carried into next sprint, we are adding {} started points for this ticket'.format(str(story_points)))
                #     current_story_points = current_story_points + story_points
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
            current_started = dict[assignee]['started']['sum']
            current_tickets = dict[assignee]['started']['tickets']
            if current_started != None:
                dict[assignee]['started']['sum'] = current_started + story_points
                current_tickets.append(ticket_number)
                dict[assignee]['started']['tickets'] = current_tickets
            else:
                dict[assignee]['started']['sum'] = story_points
                dict[assignee]['started']['tickets'] = [ticket_number]
        elif assignee == None:
            if story_points == None:
                unassigned_tickets_dict[ticket_number] = 0
            else:
                unassigned_tickets_dict[ticket_number] = story_points

    # issues that are 'done' with qa ('AWAITING RELEASE' and 'DONE' Jira status)
    for singleIssue in jira.search_issues(jql_str='issuetype != Epic AND assignee IN {} AND Sprint = "INT Sprint {}" AND status IN ("AWAITING RELEASE", "DONE", "WAITING FOR BANK") ORDER BY issuekey'.format(qes_jira_str, current_sprint)):
        ticket_number = singleIssue.key
        print('---')
        print(ticket_number)
        story_points = singleIssue.raw['fields'][story_points_key]
        current_story_points = singleIssue.raw['fields'][current_story_points_key]
        final_story_points = singleIssue.raw['fields'][final_story_points_key]

        if current_story_points != None:
            print('Current Story Points for Ticket {} exist, ticket size has been updated or ticket has been carried over from previous sprint'.format(ticket_number))
            sprints = singleIssue.raw['fields'][sprint_key]
            if (len(sprints) > 1): # ticket has been carried over
                print('Ticket has been carried across sprints')
                if (sprint_list_contains_next_sprint_but_not_previous(sprints, int(current_sprint))):
                    print('Ticket has been carried into next sprint, we are adding {} completed points for this ticket'.format(str(story_points - current_story_points))) # can add story points - current story points
                    current_story_points = story_points - current_story_points
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
            current_started = dict[assignee]['started']['sum']
            current_tickets_started = dict[assignee]['started']['tickets']
            if (ticket_number not in current_tickets_started):
                print('current started : {}'.format(current_started))
                if current_started != None:
                    dict[assignee]['started']['sum'] = current_started + story_points
                    current_tickets_started.append(ticket_number)
                    dict[assignee]['started']['tickets'] = current_tickets_started
                else:
                    dict[assignee]['started']['sum'] = story_points
                    dict[assignee]['started']['tickets'] = [ticket_number]
            current_completed = dict[assignee]['completed']['sum']
            current_tickets_completed = dict[assignee]['completed']['tickets']
            if current_completed != None:
                dict[assignee]['completed']['sum'] = current_completed + story_points
                current_tickets_completed.append(ticket_number)
                dict[assignee]['completed']['tickets'] = current_tickets_completed
            else:
                dict[assignee]['completed']['sum'] = story_points
                dict[assignee]['completed']['tickets'] = [ticket_number]
        elif assignee == None:
            if story_points == None:
                unassigned_tickets_dict[ticket_number] = 0
            else:
                unassigned_tickets_dict[ticket_number] = story_points
    remove_duplicates_and_sort_list(dict)
    return [dict, unassigned_tickets_dict]

def get_days_worked_by_name(team_out_of_office_events, sprint_start_date, sprint_end_date, google_name):
    sprint_end_date = sprint_end_date + datetime.timedelta(days=1) # adding 1 day to sprint end because for OOO days we say end day is when an individual is no longer OOO
    days_worked = 10
    for bank_holiday in get_uk_bank_holidays():
        if (sprint_start_date <= datetime.datetime.strptime(bank_holiday, '%Y-%m-%d').date() <= sprint_end_date):
            days_worked = days_worked - 1
    individual_out_of_office_events = [event for event in team_out_of_office_events if str(google_name) in event['summary']]
    for out_of_office_event in individual_out_of_office_events:
        start_ooo_str = out_of_office_event['start'].get('dateTime', out_of_office_event['start'].get('date'))
        end_ooo_str = out_of_office_event['end'].get('dateTime', out_of_office_event['end'].get('date'))
        # get overlap between OOO dates and sprint dates to find out how many days of the sprint the individual has been off
        start_ooo = datetime.datetime.strptime(start_ooo_str, date_format).date()
        end_ooo = datetime.datetime.strptime(end_ooo_str, date_format).date()
        r1 = Range(start=sprint_start_date, end=sprint_end_date)
        r2 = Range(start=start_ooo, end=end_ooo)
        latest_start = max(r1.start, r2.start)
        earliest_end = min(r1.end, r2.end)
        delta = (earliest_end - latest_start).days
        overlap = max(0, delta)
        if overlap > 0:
            business_days = np.busday_count(latest_start, earliest_end)
            overlap = business_days
        days_worked = days_worked - overlap
    return days_worked


def get_uk_bank_holidays():
    # sending get request and saving the response as response object
    r = requests.get(url = gov_uk_bank_holidays_url)
    # extracting data in json format for england and wales bank holidays
    bank_holidays_obj_array = r.json()['england-and-wales']['events']
    bank_holidays_dates_array = [bank_holiday_obj['date'] for bank_holiday_obj in bank_holidays_obj_array]
    return bank_holidays_dates_array


def get_sprint_length_from_sprint_dates(sprint_start_date, sprint_end_date):
    # for sprint length w/o weekend days
    business_days = np.busday_count(sprint_start_date, sprint_end_date) + 1
    for bank_holiday in get_uk_bank_holidays():
        if (sprint_start_date <= datetime.datetime.strptime(bank_holiday, '%Y-%m-%d').date() <= sprint_end_date):
            business_days = business_days - 1
    return business_days


def compute_individual_sprint_data_frame(data_frame, current_sprint, individual_name, individual_obj, squad_velocity, sprint_start_date, sprint_end_date):
    days_in_sprint = get_sprint_length_from_sprint_dates(sprint_start_date, sprint_end_date)
    (((_, individual_started_obj), (_, individual_completed_obj))) = individual_obj.items()
    individual_started_value = individual_started_obj['sum']
    individual_started_tickets = individual_started_obj['tickets']
    individual_completed_value = individual_completed_obj['sum']
    individual_completed_tickets = individual_completed_obj['tickets']
    google_name = get_google_name_from_jira_name(individual_name)

    team_out_of_office_events = get_team_out_of_office_events(current_sprint)
    # filter all OOO to reduce to the ones during this sprint, ie the OOO start date is on or before the end date of the sprint
    team_out_of_office_events_during_sprint = [event for event in team_out_of_office_events if datetime.datetime.strptime(event['start'].get('dateTime', event['start'].get('date')), date_format).date() <= sprint_end_date]

    individual_days_worked = get_days_worked_by_name(team_out_of_office_events_during_sprint, sprint_start_date, sprint_end_date, google_name)
    if (individual_days_worked != 0): # add individual's data to squad velocity for this sprint
        squad_velocity = squad_velocity + (individual_completed_value / individual_days_worked * days_in_sprint)
    df2 = pd.DataFrame([SprintRow(str(individual_name), individual_days_worked, individual_started_value, individual_completed_value)])
    data_frame = pd.concat([data_frame, df2])
    return (data_frame, squad_velocity)


def fill_sprint_sheet_data(data_frame, worksheet, current_sprint, sprint_start_date, sprint_end_date):
    print('sprint start date:', sprint_start_date)
    print('sprint end date:', sprint_end_date)

    tickets_dict = create_tickets_dict()

    [assigned_tickets_dict, unassigned_tickets_dict] = get_tickets_dictionaries(tickets_dict, current_sprint)

    print('\n\nAssigned Tickets')
    for key, value in assigned_tickets_dict.items():
        print(key)
        for key, value in value.items():
            print(key)
            for key, value in value.items():
                print(key, ' : ', value)
        print('\n')

    if len(unassigned_tickets_dict.items()) != 0:
        print('\n\nUnassigned Tickets')
        for key, value in unassigned_tickets_dict.items():
            print('Ticket:', key, 'Points:', value)

    business_days = get_sprint_length_from_sprint_dates(sprint_start_date, sprint_end_date)
    squad_velocity = 0
    for name, individual_data in assigned_tickets_dict.items():
        (data_frame, squad_velocity) = compute_individual_sprint_data_frame(data_frame, current_sprint, name, individual_data, squad_velocity, sprint_start_date, sprint_end_date)
    print(data_frame.to_string())
    #update the first sheet with df, starting in column C (= 3) and row Sprint x space_required_per_sprint, because space_required_per_sprint rows is the space required for each sprint's data.
    sprint_top_row_index = int(current_sprint) * space_required_per_sprint
    sprint_header_index = sprint_top_row_index - 3
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
    if (worksheet.cell('D' + str(sprint_first_data_row_index)).value) != '':
        days_worked_range_str = 'D' + str(sprint_first_data_row_index) + ':D' + str(sprint_last_data_row_index)
    else:
        days_worked_range_str = '0'
    if (worksheet.cell('E' + str(sprint_first_data_row_index)).value) != '':
        story_points_started_range_str = 'E' + str(sprint_first_data_row_index) + ':E' + str(sprint_last_data_row_index)
    else:
        story_points_started_range_str = '0'
    if (worksheet.cell('F' + str(sprint_first_data_row_index)).value) != '':
        story_points_completed_range_str = 'F' + str(sprint_first_data_row_index) + ':F' + str(sprint_last_data_row_index)
    else:
        story_points_completed_range_str = '0'

    # add footers
    worksheet.cell('C' + str(sprint_footer_data_row_index + 1)).set_text_format('bold', True).value = 'MIN'
    worksheet.cell('C' + str(sprint_footer_data_row_index + 2)).set_text_format('bold', True).value = 'MAX'
    worksheet.cell('C' + str(sprint_footer_data_row_index + 3)).set_text_format('bold', True).value = 'AVG'
    worksheet.cell('C' + str(sprint_footer_data_row_index + 4)).set_text_format('bold', True).value = 'SUM'
    worksheet.cell('C' + str(sprint_footer_data_row_index + 6)).set_text_format('bold', True).value = 'SQUAD VELOCITY'

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
    # set the squad velocity value
    cell_for_sprint_velocity = 'C' + str(sprint_footer_data_row_index + 7)
    worksheet.update_value(cell_for_sprint_velocity, str(squad_velocity))
    worksheet.update_value('C' + str(end_row_of_summary_table + int(current_sprint)), str(current_sprint))
    worksheet.update_value('D' + str(end_row_of_summary_table + int(current_sprint)), '=' + cell_for_sprint_velocity)

    worksheet.set_dataframe(data_frame, (sprint_top_row_index, 3))

    print('\nPlotting chart for sprint', current_sprint)
    current_sprint_chart = worksheet.add_chart(domain=('C' + str(sprint_first_data_row_index), 'C' + str(sprint_last_data_row_index)),
                            ranges=[('F' + str(sprint_first_data_row_index), 'F' + str(sprint_last_data_row_index))],
                            title='Story Points Completed Sprint ' + str(current_sprint),
                            chart_type=ChartType.BAR, # taken from : BAR, LINE, AREA, COLUMN, SCATTER, COMBO, STEPPED_AREA
                            anchor_cell='H' + str(sprint_top_row_index))

    print('sprint start date:', sprint_start_date)
    print('sprint end date:', sprint_end_date)


def append_to_sum_formula_str(sum_formula_str, cell_to_add):
    sum_formula_str = sum_formula_str[:-1] # remove last char from string, ie ')'
    sum_formula_str = sum_formula_str + ',' + cell_to_add + ')'
    return sum_formula_str


def compute_velocities(worksheet, max_sprint_from_command_line_args, index):
    # each sprint section is space_required_per_sprint spaces long
    first_sprint_data_row = index + 1 + first_sprint_number * space_required_per_sprint
    max_first_sprint_data_row = (max_sprint_from_command_line_args + 1) * space_required_per_sprint
    sum_of_days_worked_str = 'SUM(0)'
    sum_of_started_points_str = 'SUM(0)'
    sum_of_completed_points_str = 'SUM(0)'
    number_of_sprints_computed = 0
    while first_sprint_data_row < max_first_sprint_data_row:
        if (worksheet.cell('D' + str(first_sprint_data_row)).value != ''):
            sum_of_days_worked_str = append_to_sum_formula_str(sum_of_days_worked_str, 'D' + str(first_sprint_data_row))
            sum_of_started_points_str = append_to_sum_formula_str(sum_of_started_points_str, 'E' + str(first_sprint_data_row))
            sum_of_completed_points_str = append_to_sum_formula_str(sum_of_completed_points_str, 'F' + str(first_sprint_data_row))
            first_sprint_data_row = first_sprint_data_row + space_required_per_sprint
            number_of_sprints_computed = number_of_sprints_computed + 1
        else:
            first_sprint_data_row = first_sprint_data_row + space_required_per_sprint

    return (number_of_sprints_computed, sum_of_days_worked_str, sum_of_started_points_str, sum_of_completed_points_str)


def draw_overall_chart(worksheet):
    all_sprints_velocity_chart = worksheet.add_chart(domain=('C' + str(end_row_of_summary_table + first_sprint_number - 1), 'C' + str(end_row_of_summary_table + int(max_sprint_from_command_line_args))),
                            ranges=[('D' + str(end_row_of_summary_table + first_sprint_number - 1), 'D' + str(end_row_of_summary_table + int(max_sprint_from_command_line_args)))],
                            title=squad_velocity_chart_name,
                            chart_type=ChartType.LINE, # taken from : BAR, LINE, AREA, COLUMN, SCATTER, COMBO, STEPPED_AREA
                            anchor_cell='H' + str(start_row_of_summary_table))


def make_summary_table_headers_bold_and_set_velocity(worksheet):
    # make headers bold
    worksheet.cell('C' + str(start_row_of_summary_table)).set_text_format('bold', True)
    worksheet.cell('D' + str(start_row_of_summary_table)).set_text_format('bold', True)
    worksheet.cell('E' + str(start_row_of_summary_table)).set_text_format('bold', True)
    worksheet.cell('F' + str(start_row_of_summary_table)).set_text_format('bold', True)
    # set AVG velocity
    worksheet.cell('E' + str(end_row_of_summary_table + first_sprint_number - 1)).set_text_format('bold', True).value = 'SQUAD AVG VELOCITY (ACROSS ALL SPRINTS)'
    worksheet.cell('E' + str(end_row_of_summary_table + first_sprint_number)).value = '=SUM(D' + str(start_row_of_summary_table + 1) + ':D' + str(end_row_of_summary_table) + ')'
    # create squad velocity headers
    worksheet.cell('C' + str(end_row_of_summary_table + first_sprint_number - 1)).set_text_format('bold', True).value = 'Sprint'
    worksheet.cell('D' + str(end_row_of_summary_table + first_sprint_number - 1)).set_text_format('bold', True).value = 'Squad Velocity'


def delete_charts_that_are_being_replotted(worksheet):
    # delete charts that we are re-creating
    for chart in worksheet.get_charts():
        if (chart.title == squad_velocity_chart_name): # always being recreated
            chart.delete()
        else:
            for sprint_number in sprint_numbers_from_command_line_args:
                if (sprint_number in chart.title):
                    chart.delete()


def main():
    #authorise
    gc = pygsheets.authorize(service_file='my-project-...json')
    # Create empty dataframe
    df = pd.DataFrame()
    #open the google spreadsheet (where 'PY to Gsheet Test' is the name of my sheet)
    sh = gc.open('Integrations Team')
    #select the first sheet
    wks = sh[0]
    wks.title = 'Velocity Tracker'
    wks.adjust_column_width(start=3, end=6, pixel_size=200)
    delete_charts_that_are_being_replotted(wks)
    # wks.clear() # in case we want the worksheet to be cleared of all data
    sum_of_days_per_sprint = 0
    for current_sprint in sprint_numbers_from_command_line_args:
        if (int(current_sprint) >= first_sprint_number):
            print('\n\n\nComputing sprint data for sprint', current_sprint)
            sprint_start_date = get_sprint_start_date(int(current_sprint))
            sprint_end_date = get_sprint_end_date(int(current_sprint))
            business_days_in_sprint = np.busday_count(sprint_start_date, sprint_end_date) + 1
            for bank_holiday in get_uk_bank_holidays():
                if (sprint_start_date <= datetime.datetime.strptime(bank_holiday, '%Y-%m-%d').date() <= sprint_end_date):
                    print('Found a UK bank holiday during the sprint: ', bank_holiday)
                    business_days_in_sprint = business_days_in_sprint - 1
            sum_of_days_per_sprint = sum_of_days_per_sprint + business_days_in_sprint
            fill_sprint_sheet_data(df, wks, current_sprint, sprint_start_date, sprint_end_date)
    index = 0
    for team_member in team_members_jira:
        (number_of_sprints_computed, sum_of_days_worked_formula, sum_of_started_points_formula, sum_of_completed_points_formula) = compute_velocities(wks, int(max_sprint_from_command_line_args), index)
        print('computed data for', team_member)
        print(number_of_sprints_computed, sum_of_days_worked_formula, sum_of_started_points_formula, sum_of_completed_points_formula)
        avg_velocity = '=' + sum_of_completed_points_formula + '/' + sum_of_days_worked_formula + '*' + str(sum_of_days_per_sprint) + '/' + str(number_of_sprints_computed)
        df2 = pd.DataFrame([SummaryRow(str(team_member), avg_velocity, '=' + sum_of_started_points_formula + '/' + str(number_of_sprints_computed), '=' + sum_of_completed_points_formula + '/' + str(number_of_sprints_computed))])
        df = pd.concat([df, df2])
        index = index + 1
    wks.set_dataframe(df, (start_row_of_summary_table, start_column_of_summary_table)) # inserting summary table at (3,3)
    make_summary_table_headers_bold_and_set_velocity(wks)
    draw_overall_chart(wks)


if __name__ == '__main__':
    main()



