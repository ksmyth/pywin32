Echo on
echo .
echo .
del *.pyc
echo .
REM unit test
adodbapitest.py
echo .
echo .
REM generic test
@echo .
test_adodbapi_dbapi20.py
echo .
echo .
db_print.py
echo .
del *.pyc
del *.pyo
del ..\adodbapi.pyc
echo .
call p3k.bat adodbapitest.py
echo .
echo .
REM generic test
@echo .
call p3k.bat test_adodbapi_dbapi20.py
echo .
call p3k.bat db_print.py
echo .
pause Testing complete
