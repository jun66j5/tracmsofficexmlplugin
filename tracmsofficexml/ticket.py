# -*- coding: utf-8 -*-

import re

from trac.core import *
from trac.web.chrome import Chrome
from trac.mimeview.api import Context, IContentConverter
from trac.util.translation import _
from trac.util.datefmt import to_datetime


__all__ = ['ExcelQueryModule']


class ExcelQueryModule(Component):
    implements(IContentConverter)

    def get_supported_conversions(self):
        yield ('xlshtml', 'Excel HTML', 'html',
               'trac.ticket.Query', 'application/vnd.ms-excel', 8)
        # yield ('excelxml', 'Excel XML', 'xml',
        #        'trac.ticket.Query', 'application/vnd.ms-excel', 8)

    def convert_content(self, req, mimetype, query, key):
        if key == 'xlshtml':
            return self._export(req, query, 'query_xlshtml.html')
        if key == 'excelxml':
            content, content_type = self._export(req, query, 'query_excelxml.html')
            # XXX workaround
            content = re.sub(r'(</?)ss:', r'\1', content)
            return content, content_type

    def _export(self, req, query, template):
        # no paginator
        query.max = 0
        query.has_more_pages = False
        query.offset = 0

        # extract all fields
        cols = ['id']
        cols += [field['name'] for field in query.fields]
        for name in ('time', 'changetime'):
            if name not in cols:
                cols.append(name)
        query.cols = cols

        db = self.env.get_db_cnx()
        tickets = query.execute(req, db)
        context = Context.from_request(req, 'query', absurls=True)
        data = query.template_data(context, tickets)
        data['title'] = _('Custom Query')
        data['context'] = context

        content_type = 'application/vnd.ms-excel'
        content = Chrome(self.env).render_template(req, template, data, content_type)
        return content, content_type


