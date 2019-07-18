#!/usr/bin/env python
# -*- coding: utf-8 -*-

from setuptools import setup, find_packages

__author__ = 'Kanshi TANAIKE'

with open('README.md') as f:
    readme = f.read()

setup(
    name='gdoctableapppy',
    version='1.0.0',
    description='This is a python library to manage the tables on Google Document using Google Docs API.',
    long_description=readme,
    long_description_content_type="text/markdown",
    author='Kanshi TANAIKE',
    author_email='tanaike@hotmail.com',
    install_requires=['google-api-python-client'],
    url='https://github.com/tanaikech/gdoctableapppy',
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Programming Language :: Python :: 3.5',
    ],
    packages=find_packages(),
    keywords=['google document', 'google docs', 'docs api', 'table', 'manager', 'developer tools'],
    license='MIT License',
    test_suite='tests'
)
