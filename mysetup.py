from distutils.core import setup
import py2exe
setup (
    console=[{'script': 'ForAimee.py',
              'icon_resources': [(1, u'favicon.ico')]
              }]
    )
