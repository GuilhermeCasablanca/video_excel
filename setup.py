from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': []}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('C:\\code\\video_excel\\main.py', base=base, target_name = 'video_excel_1_10')
]

setup(name='video_excel',
      version = '1.10.0',
      description = 'Casablanca Online',
      options = {'build_exe': build_options},
      executables = executables)
