#!/usr/bin/python

import argparse
import getpass

import gdata
import gdata.data
import gdata.docs.client
import gdata.docs.data
import gdata.spreadsheet
import gdata.spreadsheet.service
from lxml import etree


def xml2dict(xml_file):

    report = list()

    with open(xml_file, 'r') as f:
        tree = etree.parse(f)

        all_tests = 0
        failed = 0
        skipped = 0

        for el in tree.findall('testcase'):
            all_tests += 1
            classname = el.attrib['classname']
            name_of_test = el.attrib['name']
            if classname == '':
                classname = name_of_test.split('(', 1)[1]
                classname = classname.split(')', 1)[0]
                name_of_test = ''
            status_list = el.getchildren()
            status = 'OK'
            comment = ''

            if not status_list:
                field = {'testclass': classname,
                         'testname': name_of_test,
                         'status': status,
                         'comment': ''}
                report.append(field)
                continue

            if 'skipped' in status_list[0].__str__():
                status = 'SKIP'
                skipped += 1
                comment = status_list[0].text
            if 'failure' in status_list[0].__str__():
                status = 'FAIL'
                failed += 1
                comment = status_list[0].text
                comment = comment.replace('\'', '"')

            field = {'testclass': classname,
                     'testname': name_of_test,
                     'status': status,
                     'comment': comment}
            report.append(field)

        passed = all_tests - skipped - failed
        summary = {'passed': str(passed),
                   'skipped': str(skipped),
                   'failed': str(failed),
                   'all': str(all_tests)}

    return summary, report


class Spreadsheet(object):

    def __init__(self, email, password, spreadsheetname):
        # Authorize
        self.client = gdata.docs.client.DocsClient(source='')
        self.client.http_client.debug = False
        self.client.client_login(email, password, source='', service='writely')
        # Create our doc
        document = gdata.docs.data.Resource(
            type='spreadsheet',
            title=spreadsheetname.split('.', 1)[0])
        self.document = self.client.CreateResource(document)

        self.spreadsheet_key = self.document.GetId().split("%3A")[1]

        self.gd_client = gdata.spreadsheet.service.SpreadsheetsService()
        self.gd_client.email = email
        self.gd_client.source = ''
        self.gd_client.password = password
        self.gd_client.ProgrammaticLogin()
        self.header_flag = False

    def fill_spreadsheet(self, report):
        ws = self.gd_client.AddWorksheet('Test results', 1, 4,
                                         self.spreadsheet_key)
        ws_id = ws.id.text
        ws_id = ws_id[(ws_id.rfind('/')+1):]

        if not self.header_flag:
            tmp = ['testclass', 'testname', 'status', 'comment']
            for i, header in enumerate(tmp):
                self.gd_client.UpdateCell(row=1, col=i+1, inputValue=header,
                                          key=self.spreadsheet_key,
                                          wksht_id=ws_id)
            self.header_flag = True

        for data in report:
            try:
                self.gd_client.InsertRow(data,
                                         self.spreadsheet_key,
                                         wksht_id=ws_id)
            except Exception:
                if len(data['comment']) > 50000:
                    data['comment'] = data['comment'][:50000]
                self.gd_client.InsertRow(data,
                                         self.spreadsheet_key,
                                         wksht_id=ws_id)

    def fill_summary(self, summary):
        for i, header in enumerate(summary.keys()):
                self.gd_client.UpdateCell(row=3, col=i+2, inputValue=header,
                                          key=self.spreadsheet_key)
                self.gd_client.UpdateCell(row=4, col=i+2,
                                          inputValue=summary.get(header),
                                          key=self.spreadsheet_key)


def main():
    parser = argparse.ArgumentParser(description="publish test results in "
                                                 "Google tables")
    parser.add_argument("-e",
                        "--email",
                        type=str,
                        help="set Google email.",
                        required=True)

    parser.add_argument("-f",
                        "--file",
                        type=str,
                        help="set path to xml with results.",
                        required=True)

    args = parser.parse_args()
    xml_file = args.file
    email = args.email
    password = getpass.getpass()
    if not password:
        raise Exception('You did not enter password!')

    summary, report = xml2dict(xml_file)

    sp = Spreadsheet(email, password, xml_file)
    sp.fill_summary(summary)
    sp.fill_spreadsheet(report)


if __name__ == "__main__":
    main()