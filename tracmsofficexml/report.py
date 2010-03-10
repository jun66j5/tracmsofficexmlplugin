# -*- coding: utf-8 -*-

import os
import pkg_resources
import re

from trac.core import *
from trac.mimeview import Context
from trac.resource import Resource
from trac.ticket.report import ReportModule
from trac.util.text import unicode_urlencode
from trac.web.api import IRequestFilter
from trac.web.chrome import add_link


__all__ = ['ExcelReportModule']


class ExcelReportModule(Component):
    implements(IRequestFilter)

    def pre_process_request(self, req, handler):
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
                locale = _detect_locale()
                if locale is not None:
                    template = '%s.%s.html' % (template[:-5], locale)
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


_page_names = {
    'TracJa': 'ja',
}

def _detect_locale():
    dir = pkg_resources.resource_filename('trac.wiki', 'default-pages') 
    for name, locale in _page_names.iteritems():
        if os.path.exists(os.path.join(dir, name)):
            return locale
    return None


def _execute_paginated_report(self, req, db, id, sql, args, limit=0, offset=0):
    excel = False
    if re.match(r'/report/[0-9]+', req.path_info) \
            and 'action' not in req.args \
            and req.args.get('format') in ('xlshtml', 'excelxml') \
            and self.env.is_component_enabled(ExcelReportModule):
        excel = True
        limit = 0
        offset = 0
    cols, results, num_items = _execute_paginated_report_orig(self, req, db, id, sql, args, limit, offset)
    if excel:
        num_items = len(results)
    return cols, results, num_items


_execute_paginated_report_orig = ReportModule.execute_paginated_report
ReportModule.execute_paginated_report = _execute_paginated_report


