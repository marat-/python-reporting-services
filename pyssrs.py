# -*- coding: utf-8 -*-
"""
    pyssrs.py
    SQL Server Reporting Services Python-library
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    :copyright: (c) 2014 by Shaidullin Marat, Bodunkov Anton.
    :license: Apache License.

"""
from urllib import quote

import requests


class SSRSReport(object):
    """ SQL Server Reportin Services Report object """

    def __init__(self, server, report_path, auth=(), params={}, output_format='EXCEL'):
        self._server = server
        self._report_path = report_path
        self._auth = ()
        if auth:
            self._auth = (
                auth[0].decode('utf-8').encode('cp1251').decode('latin1'),
                auth[1].decode('utf-8').encode('cp1251').decode('latin1'),
            )
        self._params = params
        self._format = output_format
        self._connection_string = self.get_connection_string(params)
        self._report_request = None
        self._output_file = None

    @property
    def server(self):
        """ server address """
        return self._server

    @server.setter
    def server(self, value):
        assert isinstance(value, (str, unicode))
        self._server = value

    @property
    def report_path(self):
        """ Path to report on server, eg. /Report/Report1 """
        return self._report_path

    @report_path.setter
    def report_path(self, value):
        assert isinstance(value, (str, unicode))
        self._report_path = value

    @property
    def auth(self):
        """ Tuple of user's auth data, eg. (user, pass) """
        return self._auth

    @auth.setter
    def auth(self, value):
        assert isinstance(value, tuple)
        self._auth = value

    @property
    def connection_string(self):
        """ Full URL-string for report """
        return self._connection_string

    @connection_string.setter
    def connection_string(self, value):
        self._connection_string = value

    @property
    def output_format(self):
        """
            Format of report's file, 'EXCEL' as default.
            Could be any format supported by Reporting Services
        """
        return self._format

    @output_format.setter
    def output_format(self, value):
        assert isinstance(value, (str, unicode))
        self._format = value

    def get_connection_string(self, params={}):
        """
            Builds full URL to connect to report, eg.
            http://your-reporting.com/ReportServer?/Report/Path&rs:FORMAT=EXCEL&item_id=666
        """
        url_params = []
        for key, value in params.iteritems():
            if isinstance(value, (str, unicode)):
                # Convert string params to url-format and add them to list
                value = quote("{0}".format(value).encode('cp1251'))
            url_params.append("{0}={1}".format(key, value))
        params_dict = {
            'server': self.server,
            'report_path': self.report_path,
            'format': self.output_format,
            'url_params': '&'.join(url_params),
        }
        connection_string = '{server}?{report_path}&rs:FORMAT={format}&{url_params}'.format(
            **params_dict
        )

        return connection_string

    def get_report(self):
        """ Get report's file """
        req = requests.get(self.connection_string, auth=self.auth)
        self._report_request = req

    def save_file(self, output_file):
        """ Write received content to your file """
        result = ''
        if not self._report_request:
            # Get report if it wasn't
            self.get_report()
        if self._report_request.status_code == 200:
            # Write file on success
            with open(output_file, 'wb') as xlsx_file:
                for chunk in self._report_request.iter_content(1024):
                    xlsx_file.write(chunk)

            self._output_file = output_file
        else:
            # Return error if it's happened
            result = 'Server error: {0}'.format(self._report_request.status_code)

        return result


if __name__ == '__main__':
    report = SSRSReport(
        'http://your-reporting.com/ReportServer',
        '/Report/Path',
        output_format='EXCEL',
        params={
            'item_id': 666
        }
    )
    print(report.connection_string)