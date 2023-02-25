:start
@echo off
cls
echo.
echo ##    ######    #### #####**  **  $$      $$
echo ##  ##  ## ##  ## ####    **  **  $$  $$  $$ @ --------
echo ## ##   ##  ####  ## ###  **  **  $$  $$  $$ = ==     \\
echo ##  ##  ##        ##   ## ********$$  $$  $$ = ==     ==
echo ##   ## ##        ######      **   $$$$$$$$  = ==     == by 510208 Use Heart
echo Open_____Source_____Windows___And___Office_____Activeter
echo.
echo.
set KmsServer=kms.03k.org
set OfficePath="C:\Program Files\Microsoft Office\Office16"
echo =========================================
echo  [W] 啟用Microsoft Windows
echo  [O] 啟用Microsoft Office
echo _________________________________________
echo  [S] 設定
echo  [I] 關於本電腦
echo  [E] 離開軟體
echo.
echo KMS主機設為： %KmsServer%
echo Office路徑為： %OficePath%
echo.
echo  KMS4Win by 510208：免費Windows啟用工具
echo =========================================
set /P M=請選擇選項:
IF %M%==w GOTO WINDOWS
IF %M%==o GOTO OFFICE
IF %M%==s GOTO SETTING
IF %M%==i GOTO INFO
IF %M%==e EXIT
goto start

:WINDOWS
@echo on
slmgr.vbs /upk
slmgr /skms %KmsServer%
slmgr /ato
@echo off
echo.
echo =========================================
echo  啟用完畢
echo.
echo  如果上面沒有跳出啟用錯誤，代表啟用成功；
echo  否則，請嘗試更換KMS伺服器或檢查OS是否為
echo  VL版本。如不是，底下提供手動密鑰，將其貼
echo  到CMD中（slmgr /ipc 啟用密鑰）以取得授權
echo  。
echo _________________________________________
echo  以下網址為微軟官方放出的Key，請照上文方
echo  式啟用：
echo  https://docs.microsoft.com/zh-cn/windows-server/get-started/kmsclientkeys
echo _________________________________________
echo  感謝您的使用。
echo =========================================
pause > nul
goto start

:OFFICE
@echo on
cd %OfficePath%
cscript ospp.vbs /sethst:%KmsServer
cscript ospp.vbs /act
@echo off
echo.
echo =========================================
echo  啟用完畢
echo.
echo  如果上面有跳出
echo  <Product activation successful>，代表啟用
echo  成功。
echo.
echo  否則，請檢查你的Office路徑
echo  ：%OfficePath%
echo  與KMS主機
echo  ：%KmsServer%
echo  是否正常
echo _________________________________________
echo  感謝您的使用。
echo =========================================
pause > nul
goto start

:SETTING
cls
echo ##    ######    #### #####**  **  $$      $$
echo ##  ##  ## ##  ## ####    **  **  $$  $$  $$ @ --------
echo ## ##   ##  ####  ## ###  **  **  $$  $$  $$ = ==     \\
echo ##  ##  ##        ##   ## ********$$  $$  $$ = ==     ==
echo ##   ## ##        ######      **   $$$$$$$$  = ==     ==
echo.
echo =========================================
echo  [K] 設定KMS主機
echo  [O] 設定Office位址與路徑
echo _________________________________________
echo  [E] 離開設定
echo =========================================
IF %M%==k GOTO KMSSERV
IF %M%==o GOTO OFFICEPATH
IF %M%==e GOTO start

:OFFICEPATH
set /P OficePath=請輸入Office路徑（可參考此軟體Github頁面的Readme.md說明檔）
IF EXIST %OfficePath%\OSPP.vbs (echo 成功找到Office檔案，按任意鍵繼續 & pause > nul) ELSE (echo 找不到Office檔案，按任意鍵繼續。（請注意！路徑分隔必須要是反斜槓（數字鍵８上面） & pause > nul & goto OFFICEPATH)
goto start
:KMSSERV
set /P KmsServer=請輸入新主機IP或網域名稱（不須輸入通信協定）
echo KMS主機設為： %KmsServer%
set /P YesNoPing=是否先Ping該主機？(Y/N)
IF %YesNoPing%==n GOTO start
ping %KmsServer
echo 如果上面沒有顯示遺失（遺失為0%），代表此主機可用。
pause > nul
goto start

:INFO
cls
echo ##    ######    #### #####**  **  $$      $$
echo ##  ##  ## ##  ## ####    **  **  $$  $$  $$ @ --------
echo ## ##   ##  ####  ## ###  **  **  $$  $$  $$ = ==     \\
echo ##  ##  ##        ##   ## ********$$  $$  $$ = ==     ==
echo ##   ## ##        ######      **   $$$$$$$$  = ==     ==
echo.
echo =========================================
echo  系統檢視器 by 510208
echo _________________________________________
wmic os get caption
echo =========================================
pause
goto start