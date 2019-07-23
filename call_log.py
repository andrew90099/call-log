#!/usr/bin/env python
# -*- coding: utf-8 -*-
import requests

import argparse
import csv
import datetime
import re
import smtplib
import string
import sys

from collections import OrderedDict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os import path

import config

BASE_DIR = path.dirname(path.abspath(__file__))
VERSION = '0.2'

BASE_HTML = '''<!DOCTYPE html>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
</head>
<body>
${html_body_tables}</body>
</html>'''
CALL_DATA_HTML = '''		<tr>
			<td width="50" style="border-right: thin solid black; border-bottom: thin solid black; padding: 0 1em;">${time}</td>
			<td width="75" style="border-right: thin solid black; border-bottom: thin solid black; padding: 0 1em;">${type}</td>
			<td width="75" style="border-right: thin solid black; border-bottom: thin solid black; padding: 0 1em;">${duration}</td>
			<td width="300" style="border-bottom: thin solid black; padding: 0 1em;">${tofrom}</td>
		</tr>'''
CALL_DATA_TXT = '''|${time}|${type}|${duration}|${tofrom}\n'''
USER_HTML = '''	<table cellspacing="0" cellpadding="0" border="0">
		<thead style="background-color: #ddd">
		<tr><th colspan=4 style="text-align: left;">${name}</th></tr>
		<tr><td colspan=4>${out_count} outgoing calls (${out_dur})</td></tr>
		<tr><td colspan=4>${in_count} ingoing calls (${in_dur})</td></tr>
		<tr>
			<th width="50" style="text-align: left; padding: 0 1em;">time</th>
			<th width="75" style="text-align: left; padding: 0 1em;">type</th>
			<th width="75" style="text-align: left; padding: 0 1em;">length</th>
			<th width="300" style="text-align: left; padding: 0 1em;">to/from</th>
		</tr>
		</thead>
		<tbody>
${call_data}
		</tbody>
		<tfoot style="background-color: #ddd">
		<tr><th colspan=4>total call time: ${total_dur}</th></tr>
		</tfoot>
	</table>
	<br />
'''
USER_TXT = '''|${name}
|${out_count} outgoing calls (${out_dur})
|${in_count} ingoing calls (${in_dur})
${call_data}|total call time: ${total_dur}

'''

def get_call_log_data(date, num_days=1):
    url = 'http://192.168.1.180:8080/logs.asp'
    # in the web interface, the months are zero-indexed
    url_params = {'Class': 'Logs', 'Subclass': 'cdr', 'op': 'View', 'month':
            date.month - 1, 'day': date.day, 'year': date.year, 'ndays':
            num_days, 'ExportTSVReport': 'Export TSV Report' }
    tsv = requests.get(url, url_params)
    # tsv has a trailing '\r\n'
    tsv_string = tsv.text.strip('\r\n')
    # the header has duplicate column headings.  If we run it through
    # csv.DialectictReader, the duplicate columns overwrite each other causing
    # a loss of data.
    reader = csv.reader(tsv_string.split('\t\r\n'), delimiter='\t')
    # next() acts like .pop() on the reader.  We'll construct our own header.
    header = next(reader)
    # These are more correct column headings.
    header = ('Id', 'Start Date', 'Start Time', 'Duration', 'From',
            'From CID Name', 'From CID Number', 'DNIS Name', 'DNIS Number',
            'From Port', 'To', 'To CID Name', 'To CID Number', 'To Port',
            'PIN', 'Digits', 'TC')
    # return a generator which produces a dictionary of only the
    # config.SALESPEOPLE's call data
    return (dict(zip(header, row)) for row in reader if is_sales(row))


def is_sales(row_data):
    """If either the To or From name is one of the SALESPEOPLE, return True.
    From is index 4, and To is index 10 in the row data."""
    return any([k in config.SALESPEOPLE for k in [row_data[4], row_data[10]]])


def process_data(data):
    """Once we've gotten all the data, we will cut it down to only the data we
    need, formatted properly."""
    # set up an OrderedDict to hold the data with initial data set to 0
    output = OrderedDict((s, {'calls': [], 'out_count': 0, 'out_dur':
        datetime.timedelta(seconds=0), 'in_count': 0, 'in_dur':
        datetime.timedelta(seconds=0), 'total_dur':
        datetime.timedelta(seconds=0)}) for s in config.SALESPEOPLE)
    for d in data:
        # assume it's an outgoing call.
        call_type = 'outgoing'
        key = d['From']
        tofrom = '{0} {1}'.format(d['To CID Name'], d['To CID Number'])
        if not tofrom.strip():
            tofrom = d['Digits'].strip() or d['To']
        out_count, in_count = 1, 0
        dur_h, dur_m, dur_s = map(int, d['Duration'].split(':'))
        duration = datetime.timedelta(hours=dur_h, minutes=dur_m,
                seconds=dur_s)
        out_dur = duration
        in_dur = datetime.timedelta(seconds=0)
        if key not in config.SALESPEOPLE:
            # it's an incoming call if the From name isn't one of the
            # config.SALESPEOPLE.  Adjust the data accordingly
            call_type = 'incoming'
            key = d['To']
            tofrom = '{0} {1}'.format(d['From CID Name'], d['From CID Number'])
            if not tofrom.strip():
                tofrom = d['To']
            out_count, in_count = 0, 1
            out_dur = datetime.timedelta(seconds=0)
            in_dur = duration

        # format the phone numbers
        tofrom = re.sub(r'1?(\d{3})(\d{3})(\d{4})$', r'\1-\2-\3', tofrom)

        output[key]['calls'].append({'time': d['Start Time'], 'type':
            call_type, 'duration': d['Duration'], 'tofrom': tofrom})
        output[key]['out_count'] = output[key]['out_count'] + out_count
        output[key]['out_dur'] = output[key]['out_dur'] + out_dur
        output[key]['in_count'] = output[key]['in_count'] + in_count
        output[key]['in_dur'] = output[key]['in_dur'] + in_dur
        output[key]['total_dur'] = output[key]['total_dur'] + duration

    return output


def generate_email(processed_data, date):
    recipients = config.RECIPIENTS.extend([
        (s, '{0}+calllogs@tomcodydesign.com'.format(s.split()[0].lower())) for
            s in config.SALESPEOPLE])
    email_recipients = ['{0} <{1}>'.format(name, email) for (name, email) in
            config.RECIPIENTS]
    html_body_tables, txt_body_tables = [], []
    for n, data in processed_data.items():
        call_list_h, call_list_t = [], []
        for call in data['calls']:
            call_data_template_h = string.Template(CALL_DATA_HTML)
            call_list_h.append(call_data_template_h.substitute(**call))
            call_data_template_t = string.Template(CALL_DATA_TXT)
            call_list_t.append(call_data_template_t.substitute(**call))
        call_data_h = '\n'.join(call_list_h)
        call_data_t = ''.join(call_list_t)
        user_template_h = string.Template(USER_HTML)
        user_template_t = string.Template(USER_TXT)
        template_data = {'name': n, 'out_count': data['out_count'], 'out_dur':
                data['out_dur'], 'in_count': data['in_count'], 'in_dur':
                data['in_dur'], 'total_dur': data['total_dur']}
        html_body_tables.append(user_template_h.substitute(
            call_data=call_data_h, **template_data))
        txt_body_tables.append(user_template_t.substitute(
            call_data=call_data_t, **template_data))
    html_body = string.Template(BASE_HTML)
    txt_body = ''.join(txt_body_tables)
    html_body_tables = ''.join(html_body_tables)
    html_body = html_body.substitute({'html_body_tables': html_body_tables})
    msg = MIMEMultipart('alternative')
    charset = 'us-ascii' if config.DEBUG else 'utf-8'
    txt_part = MIMEText(txt_body, 'plain', charset)
    html_part = MIMEText(html_body, 'html', charset)
    msg.attach(txt_part)
    msg.attach(html_part)

    msg.add_header('X-CALL_LOG-VERSION', VERSION)
    msg.add_header('From', 'Tom Cody Design <info@tomcodydesign.com>')
    msg.add_header('To', ', '.join(email_recipients))
    msg.add_header('Subject', 'CALL LOG: {0:%Y-%m-%d}'.format(date))
    recipients = [email_recipients[0],] if config.DEBUG else email_recipients
    return (msg, recipients)


def send_email(msg, recipients):
    s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    s.login('info@tomcodydesign.com', 'info2nite')
    s.sendmail('info@tomcodydesign.com', recipients, msg.as_string())
    s.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('date_string', type=str, nargs='?', help='date to '
            'query the phone system. format: YYYY-MM-DD  default: yesterday')
    parser.add_argument('-n', '--num_days', type=int, nargs='?', default=1,
            help='number of days to query phone system for.  default: 1')
    args = parser.parse_args()
    date_string = args.date_string
    if date_string is None:
        date = datetime.date.today() - datetime.timedelta(days=1)
    else:
        try:
            year, month, day = map(int, date_string.split('-'))
            date = datetime.date(year, month, day)
        except ValueError:
            print("incorrect format for date_string.  Please make sure it's "
                    "YYYY-MM-DD")
            return 1

    num_days = args.num_days

    data = get_call_log_data(date, num_days)
    processed_data = process_data(data)
    msg, recipients = generate_email(processed_data, date)
    if config.DEBUG:
        print(msg.as_string())
    else:
        send_email(msg, recipients)

    return 0


if __name__ == '__main__':
    sys.exit(main())