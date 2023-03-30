set srv=%1
set in=%2
set out=%3
for /F %%i in (args.txt) do echo %%i

sqlcmd -S %srv% -i "%__CD__%sql\%in%" -s ',' -W -o "%__CD__%txt\%out%" 2> "%__CD__%err\tables.err"