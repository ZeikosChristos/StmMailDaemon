sc query "StmMailDaemon" | find "RUNNING"
if %errorlevel% neq 0 net start "StmMailDaemon"