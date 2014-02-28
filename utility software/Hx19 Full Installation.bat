@echo If you have not read Hexamite terms of sale.pdf document
@echo on your installation disk, or you disagree with the terms
@echo please close this window, otherwise hit any button to continue
pause

@echo This progam installs all the hx19 programs one by one
@echo and all the basic files needed for operation
@echo
@echo Continue until you see Installation Finished below
@echo
@echo off

cd hx19v2access
setup
cd..
@echo hx19v2access setup is completed

cd hx19v2xyzDDE
setup
cd..
@echo hx19v2xyzDDE setup is completed

cd hx19v2xyzlabDDE
setup
cd..
@echo hx19v2xyzLabDDE setup is completed


copy *.txt c:\Progra~1\HX19v2\*.txt

md c:\Progra~1\HX19v2\dataFiles

@echo *** Installation Finished ***



pause
