#!/usr/bin/env python

"""
Generate a monthly .xls report of Jira tasks assigned to me.

Usage:
$ poetry run jira-report [--month 2019/10] [--days 21] [--force-overwrite] [--blacklist <path>]

Configuration:
$ echo 'JIRA_SERVER_URL="https://mycompany.atlassian.net"' >> .env
$ echo 'JIRA_USERNAME="jdoe@mycompany.com"' >> .env
$ echo 'JIRA_API_TOKEN="qeYEtFiNUJ8FCSEbBp25jNKc"' >> .env

Interactive prompt appears if a local .env file is missing.
"""

import argparse
import calendar
import datetime
import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

import dateutil.parser
import environs
import jira
import workdays
import xlwt

logging.basicConfig(format='[%(levelname)s] %(message)s', level=logging.INFO)
LOGGER = logging.getLogger(__name__)


def run() -> None:
    """Command wrapper for Poetry."""
    try:
        main(parse_args())
    except ValueError as ex:
        LOGGER.error(str(ex))
    except KeyboardInterrupt:
        LOGGER.warning('Aborted')


def main(args: argparse.Namespace) -> None:
    """Script entry point."""

    title = args.date.strftime("%Y_%B")
    filename = f'Jira_{title}.xls'

    if os.path.exists(filename) and not args.force_overwrite:
        LOGGER.error('File already exists: "%s", use the -f flag to overwrite', filename)
    else:
        issues = blacklist(args.blacklist, find_issues(args.date, jira_config()))
        if len(issues) > 0:
            LOGGER.info('Found %d tasks assigned to you during that period.', len(issues))
            xls_export(issues, month_hours(args.date, args.business_days), title, filename)
        else:
            LOGGER.info('There were no tasks assigned to you during that period.')


def parse_args() -> argparse.Namespace:
    """Parse command line arguments."""

    def parse_month(text: str) -> datetime.date:
        """Return a datetime instance."""
        return datetime.datetime.strptime(text, '%Y/%m').date()

    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--force-overwrite', action='store_true', default=False)
    parser.add_argument('-d', '--days', dest='business_days', type=int)
    parser.add_argument('-b', '--blacklist')
    parser.add_argument('--month',
                        metavar='YYYY/MM',
                        dest='date',
                        type=parse_month,
                        default=datetime.date.today())

    args = parser.parse_args()

    if args.blacklist is not None:
        args.blacklist = Path(args.blacklist)
        if not (args.blacklist.exists() and args.blacklist.is_file()):
            raise ValueError(f'"{args.blacklist}" is not a valid path')

    return args


def blacklist(blacklist_path: Optional[Path], issues: List[jira.Issue]) -> List[jira.Issue]:
    """Filter out blacklisted issues."""

    if blacklist_path is None:
        return issues

    LOGGER.info('Using blacklist from: %s', blacklist_path.resolve())
    blacklisted = [x.strip() for x in blacklist_path.read_text().split('\n') if x != '']

    skipped = 0
    filtered_issues = []
    for issue in issues:
        if issue.key in blacklisted:
            skipped += 1
        else:
            filtered_issues.append(issue)

    if skipped > 0:
        LOGGER.info('Skipped %d blacklisted issues', skipped)

    return filtered_issues


def jira_config() -> Dict[str, str]:
    """Return a dict of Jira configuration options."""

    environs.load_dotenv()

    load_var('JIRA_SERVER_URL')
    load_var('JIRA_USERNAME')
    load_var('JIRA_API_TOKEN')

    return {
        'server': os.getenv('JIRA_SERVER_URL'),
        'basic_auth': (os.getenv('JIRA_USERNAME'), os.getenv('JIRA_API_TOKEN'))
    }


def load_var(name: str) -> None:
    """Ensure that Jira configuration is stored in .env file."""
    if name not in os.environ:
        prompt = ' '.join([x.title() for x in name.split('_')]) + ': '
        while True:
            value = input(prompt)
            if value.strip() != '':
                break
        with open('.env', 'a') as file_object:
            print(f'{name}="{value}"', file=file_object)
        environs.load_dotenv()


def find_issues(date: datetime.date, config: Dict[str, str]) -> jira.client.ResultList:
    """Return a list of Jira issues for the given month."""
    logging.info('Querying Jira...')
    api = jira.JIRA(**config)
    return api.search_issues(jql(date))


def jql(date: datetime.date) -> str:
    """Return a JQL query to get issues assigned to me in the given month."""
    start_date = f'{date.year}/{date.month:02}/01'
    end_date = f'{date.year}/{date.month:02}/{month_days(date):02}'
    return f'assignee was currentUser() DURING ("{start_date}", "{end_date}") ORDER BY created ASC'


def month_days(date: datetime.date) -> int:
    """Return the number of days in the given month."""
    _, num_days = calendar.monthrange(date.year, date.month)
    return num_days


def month_hours(date: datetime.date, business_days: Optional[int]) -> int:
    """Return the number of work hours in the given month."""

    if business_days is None:
        start_date = datetime.date(date.year, date.month, 1)
        end_date = datetime.date(date.year, date.month, month_days(date))
        business_days = workdays.networkdays(start_date, end_date)

    LOGGER.info('Business days=%d (%d hours)', business_days, business_days * 8)

    return business_days * 8


def xls_export(issues: List[jira.Issue],
               hours: int,
               title: str,
               filename: str) -> None:
    """Save Jira issues to a spreadsheet file."""

    class Styles:
        """A class whose attributes represent different styles."""

        bold = xlwt.easyxf('font: bold on; align: vert centre')
        middle = xlwt.easyxf('align: vert centre')

        date_format = xlwt.easyxf('align: vert centre, horiz left')
        date_format.num_format_str = 'yyyy-mm-dd, HH:MM'

        hours_format = xlwt.easyxf('align: vert centre, horiz right')
        hours_format.num_format_str = '#,#0.0 "h"'

        invisible = xlwt.easyxf('align: vert centre; font: color white')

    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet(title)

    row_height = sheet.row_default_height = 384

    column_headers = (
        'Task ID',
        'Task Key',
        'Task URL',
        'Project Name',
        'Created At',
        'Description',
        'Worklog',
    )

    styles = Styles()

    for column, header in enumerate(column_headers):
        write(sheet, 0, column, header, styles.bold)
        sheet.row(0).height = row_height

    for row, issue in enumerate(issues, 1):
        sheet.row(row).height = row_height
        write(sheet, row, 0, issue.id, styles.middle)
        write(sheet, row, 1, issue.key, styles.middle)
        write(sheet, row, 2, make_link(issue.permalink()), styles.middle)
        write(sheet, row, 3, issue.fields.project.name, styles.middle)
        write(sheet, row, 4, make_datetime(issue.fields.created), styles.date_format)
        write(sheet, row, 5, issue.fields.summary, styles.middle)
        sheet.write(row, 6, hours_worked(row, issues), styles.hours_format)
        write(sheet, row, 7, story_points(issue), styles.invisible)

    write(sheet, 0, 7, hours, styles.invisible)

    workbook.save(filename)
    logging.info('Exported file: "%s"', os.path.join(os.getcwd(), filename))


def write(sheet: xlwt.Worksheet, row: int, col: int, value: Any, style: xlwt.XFStyle) -> None:
    """Write text to a cell and auto-fit the column width."""

    sheet.write(row, col, value, style)

    char_width = 256
    text_width = len(str(value)) * char_width

    column = sheet.col(col)
    if column.get_width() < text_width:
        column.set_width(text_width)


def make_datetime(text: str) -> datetime.datetime:
    """Return an offset-naive datetime from an ISO-8601 string."""
    return dateutil.parser.parse(text).replace(tzinfo=None)


def make_link(url: str) -> xlwt.Formula:
    """Return an interactive hyperlink formula."""
    return xlwt.Formula(f'HYPERLINK("{url}")')


def story_points(issue: jira.Issue) -> Optional[float]:
    """Return the number of story points of None."""
    try:
        return issue.fields.customfield_10020
    except AttributeError:
        logging.warning('No story points assigned to %s', issue.key)
        return None


def hours_worked(row: int, issues: List) -> xlwt.Formula:
    """Return a math formula to calculate the number of hours worked on an issue."""
    return xlwt.Formula(f'H{row + 1}/SUM(H2:H{len(issues) + 1})*H1')
