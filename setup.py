from setuptools import setup

APP = ['main.py', 'gui.py', 'utils.py']
OPTIONS = {
    'argv_emulation': True,
    'packages': ['docx', 'lxml', 'json', 'os', 're', 'tkinter', 'threading', 'webbrowser'],
    'iconfile': 'resources/icon.icns',
}

setup(
    app=APP,
    options={'py2app': OPTIONS},
    use_scm_version=True,
    setup_requires=['py2app', 'setuptools_scm'],
)
