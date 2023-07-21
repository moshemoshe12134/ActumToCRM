from setuptools import setup

APP = ['main.py']

OPTIONS = {
    'argv_emulation': True,
    'iconfile': 'icons8-automation-64.png'
}

setup(
  name='AchToCRM',
  app=APP,
  options={'py2app': OPTIONS},
  setup_requires=['py2app'],
  install_requires=[
     'selenium',
     'pandas',
     'openpyxl',
     'datetime',
     'pyautogui',
     'time',
     'os',
     'glob',
     'pywin32'
  ]
)