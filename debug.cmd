@echo Running debug on %2..
chcp 866

@cd /D %1\projects\%2
@debug %2.com %3

Exit
