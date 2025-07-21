@echo off
rem === 把下面路径改成你要清空的目录 ===
set "TARGET_DIR=C:\Users\13684\Desktop\jophin\joplin_word"
set "TARGET_DIR2=C:\Users\13684\Desktop\jophin\joplindirectoutput"

rem 删除目录里的所有文件
DEL /Q /F "%TARGET_DIR%\*"

rem 删除目录里的所有子文件夹及其内容
FOR /D %%p IN ("%TARGET_DIR%\*") DO RMDIR /S /Q "%%p"

rem 删除目录里的所有文件
DEL /Q /F "%TARGET_DIR2%\*"

rem 删除目录里的所有子文件夹及其内容
FOR /D %%p IN ("%TARGET_DIR2%\*") DO RMDIR /S /Q "%%p"


call joplin --profile "C:\Users\13684\.config\joplin-desktop" export --format md_frontmatter "C:\Users\13684\Desktop\jophin\joplindirectoutput"



set CONDA_BASE_PATH=C:\Users\13684\anaconda3
call "%CONDA_BASE_PATH%\Scripts\conda.exe" run -n llms python addsvg.py
call "%CONDA_BASE_PATH%\Scripts\conda.exe" run -n llms python audiotoword.py
call "%CONDA_BASE_PATH%\Scripts\conda.exe" run -n llms python md2word.py
call "%CONDA_BASE_PATH%\Scripts\conda.exe" run -n llms python mergetoonefile_gudinglujing.py


