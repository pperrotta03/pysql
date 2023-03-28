set in=%1
set out=%2
for /F %%i in (args.txt) do echo %%i

sqlcmd -i "%__CD__%sql\%in%" -s ',' -W -o "%__CD__%txt\%out%" 2> "%__CD__%err\tables.err"