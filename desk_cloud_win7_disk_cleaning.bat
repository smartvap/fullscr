@echo off
FOR /F "usebackq delims==" %%i IN (`dir /A:D /B "C:\Users\heqiming\.Dynatrace\Dynatrace 6.5\client\sessions\offline"`) DO rd /s /q "C:\Users\heqiming\.Dynatrace\Dynatrace 6.5\client\sessions\offline\%%i"
FOR /F "usebackq delims==" %%i IN (`dir /A:D /B "C:\Users\heqiming\.Dynatrace\Dynatrace 6.5\log\client"`) DO rd /s /q "C:\Users\heqiming\.Dynatrace\Dynatrace 6.5\log\client\%%i"
del /f /s /q "C:\Users\heqiming\.Dynatrace\Dynatrace 6.5\log\client\*.*"
@echo on