rd /s /q __pycache__
rd /s /q build
rd /s /q dist
set PATH=C:\Users\katahiromz\AppData\Roaming\Python\Python312-32\Scripts;C:\Program Files (x86)\Python312-32;%PATH%
rem pyinstaller MyName.py --icon icon.ico
pyinstaller MyName.py --noconsole --icon icon.ico
mkdir dist\MyName\data
copy README.txt dist\MyName
copy LICENSE.txt dist\MyName
copy icon.ico dist\MyName\data
copy template.docx dist\MyName
