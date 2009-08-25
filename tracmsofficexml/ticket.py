# -*- coding: utf-8 -*-

import locale
import re
import sys

from trac.core import *
from trac.web.chrome import Chrome
from trac.web.api import IRequestFilter
from trac.mimeview.api import Context, IContentConverter
from trac.util.translation import _
from trac.util.datefmt import to_datetime


__all__ = ['ExcelHTMLQueryModule']


class ExcelHTMLQueryModule(Component):
    implements(IContentConverter)

    def get_supported_conversions(self):
        yield ('xlshtml', 'Excel HTML', 'html',
               'trac.ticket.Query', 'application/vnd.ms-excel', 8)
        yield ('excelxml', 'Excel XML', 'html',
               'trac.ticket.Query', 'application/vnd.ms-excel', 8)

    def convert_content(self, req, mimetype, query, key):
        if key == 'xlshtml':
            return self.export_xlshtml(req, query)
        if key == 'excelxml':
            return self.export_excelxml(req, query)

    def export_xlshtml(self, req, query):
        return self._export(req, query, 'query_xlshtml.html')

    def export_excelxml(self, req, query):
        content, content_type = self._export(req, query, 'query_excelxml.html')
        # XXX workaround
        content = content.replace('<ss:Workbook ', '<Workbook ', 1)
        return content, content_type

    def _export(self, req, query, template):
        # no paginator
        query.max = 0
        query.has_more_pages = False
        query.offset = 0

        db = self.env.get_db_cnx()
        tickets = query.execute(req, db)
        context = Context.from_request(req, 'query', absurls=True)
        data = query.template_data(context, tickets)
        data['title'] = _('Custom Query')
        data['context'] = context
        data.setdefault('report', None)
        data.setdefault('description', None)

        def format_datetime(t=None, format='%Y-%m-%d %H:%M:%S', tzinfo=None):
            t = to_datetime(t, tzinfo)
            text = t.strftime(format)
            encoding = locale.getpreferredencoding() or sys.getdefaultencoding()
            if sys.platform != 'win32' or sys.version_info[:2] > (2, 3):
                encoding = locale.getlocale(locale.LC_TIME)[1] or encoding
            return unicode(text, encoding, 'replace')
        data['format_datetime'] = format_datetime

        content_type = 'application/vnd.ms-excel'
        content = Chrome(self.env).render_template(req, template, data, content_type)
        return content, content_type


