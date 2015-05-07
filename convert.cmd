@Echo off
CD C:\test
pandoc test.md -f markdown -t html -s -o C:\test\html\test.html
html2email.vbs C:\test\html\test.html
@echo on