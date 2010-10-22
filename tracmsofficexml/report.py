# -*- coding: utf-8 -*-

import re

from trac.core import *
from trac.mimeview import Context
from trac.resource import Resource
from trac.util.text import unicode_urlencode
from trac.web.api import IRequestFilter
from trac.web.chrome import add_link


__all__ = ['ExcelReportModule']


class ExcelReportModule(Component):
    implements(IRequestFilter)

    def pre_process_request(self, req, handler):
        if re.match(r'/report/[0-9]+', req.path_info) \
                and req.args.get('format') in ('xlshtml', 'excelxml') \
                and handler.__class__.__name__ == 'ReportModule':
            req.args['max'] = 0
        return handler

    def post_process_request(self, req, template, data, content_type):
        if template == 'report_view.html' and req.args.get('id'):
            format = req.args.get('format')
            if format == 'xlshtml':
                template = 'report_xlshtml.html'
            elif format == 'excelxml':
                template = 'report_excelxml.html'
            else:
                format = None
            if format is not None:
                content_type = 'application/vnd.ms-excel'
                resource = Resource('report', req.args['id'])
                data['context'] = Context.from_request(req, resource, absurls=True)
            else:
                self._add_alternate_links(req)
        return template, data, content_type

    def _add_alternate_links(self, req):
        params = {}
        for arg in req.args.keys():
            if not arg.isupper():
                continue
            params[arg] = req.args.get(arg)
        if 'USER' not in params:
            params['USER'] = req.authname
        if 'sort' in req.args:
            params['sort'] = req.args['sort']
        if 'asc' in req.args:
            params['asc'] = req.args['asc']
        href = ''
        if params:
            href = '&' + unicode_urlencode(params)
        add_link(req, 'alternate', '?format=xlshtml' + href, 'Excel HTML', 'application/vnd.ms-excel', 'html')
        # add_link(req, 'alternate', '?format=excelxml' + href, 'Excel XML', 'application/vnd.ms-excel', 'xml')


