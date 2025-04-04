@echo off
rem 事前にpythonをexe化するライブラリをインストールしておく「pip install pyinstaller」
pyinstaller --noconsole --onefile powerpoint_splitter.py
