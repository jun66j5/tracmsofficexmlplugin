# -*- coding: utf-8 -*-

from trac.core import *
from trac.web.chrome import ITemplateProvider


__all__ = ['ExcelHTMLTemplateProvider']


class ExcelHTMLTemplateProvider(Component):
    implements(ITemplateProvider)

    def get_htdocs_dirs(self):
        return []

    def get_templates_dirs(self):
        from pkg_resources import resource_filename
        return [ resource_filename(__name__, 'templates') ]


