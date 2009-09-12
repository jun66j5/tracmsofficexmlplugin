#!/usr/bin/env python
# -*- coding: utf-8 -*-

from setuptools import setup, find_packages

setup(
    name = 'TracMsOfficeXmlPlugin', version = '0.1',
    description = 'Trac Microsoft Office XML Plugin',
    license='MIT License',
    packages = find_packages(exclude=['*.tests*']),
    package_data = {
        'tracmsofficexml' : [ 'templates/*.html' ],
    },
    entry_points = {
        'trac.plugins': [
            'tracmsofficexml.web_ui = tracmsofficexml.web_ui',
            'tracmsofficexml.ticket = tracmsofficexml.ticket',
            'tracmsofficexml.report = tracmsofficexml.report',
        ],
    },
)
