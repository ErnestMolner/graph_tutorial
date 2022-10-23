#!/usr/bin/env python
from setuptools import setup, find_packages

setup(name='microsoft_graph_whatever',
      version='1.0.1',
      description='This package is a implementation of microsoft graph api utilities',
      author='Ernest Molner',
      author_email='whatever.piques.my.interest@gmail.com',
      url='https://www.youtube.com/channel/UCgoux3jJpq1EDknYbkzfvuw',
      install_requires=['keyring', 'requests', 'msal'],
      packages= find_packages(exclude=['tests'])
      )
