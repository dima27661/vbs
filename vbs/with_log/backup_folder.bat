@echo on
 
set dat=%date:~-4%.%date:~3,2%.%date:~,2%
"C:\Program Files\7-Zip\7z.exe" a -t7z -ssw -mx9 -r0 "D:\temp\to\%dat%.7z" "D:\temp\from" >>"D:\temp\to\log_file.%dat%.txt"
rem -x	Исключить файлы или папки из архива.
rem xcopy "\\server-1234\Information\BackUp\BackUp %dat%.7z" "D:\temp\to" /E /F /H /R /K /Y /D
 
pause