@echo off
setlocal EnableDelayedExpansion
mode 70,30
color 9f
title ics2csv4win V.0.2 ~ 2023 ~ emkadoc.de

rem this param should by modified by the user
set "cal_url=https://<domain>/<path_to_ics_file>"

rem fixed params - do not modify
set "cal_file=cal.tmp"
set "csv_file=cal.csv"
set "date_from="
set "date_to="
set "substring_SUMMARY=SUMMARY:"
set "substring_DTSTART=DTSTART:"
set "substring_DTEND=DTEND:"
set "substring_DTSTART_ALT=DTSTART;VALUE=DATE:"
set "substring_DTEND_ALT=DTEND;VALUE=DATE:"
set "substring_VEVENT_END=END:VEVENT"
set "exported=true"
set "dst_deviation=1"

echo -------------------------
echo    ics2csv4win V.0.2
echo -------------------------
echo    2023 ~ emkadoc.de
echo -------------------------
call :downloadfile %cal_url% %cal_file%

echo 1. Start (Format: YYYYMMDD [en] / JJJJMMTT [de]):
set /p date_from=
echo 2. End (Format: YYYYMMDD [en] / JJJJMMTT [de]):
set /p date_to=
echo 3. Downloading...

(for /f "delims=" %%i in (%cal_file%) do (
	set "line=%%i"

	if not "!line:%substring_DTSTART%=!"=="!line!" (
		set "line_new=!line:DTSTART:=!"
		set "final_dtstart=!line_new!"
	)

	if not "!line:%substring_DTSTART_ALT%=!"=="!line!" (
		set "line_new=!line:~19!"
		set "final_dtstart=!line_new!"
	)

	if not "!line:%substring_DTEND%=!"=="!line!" (
		set "line_new=!line:DTEND:=!"
		set "final_dtend=!line_new!"
	)

	if not "!line:%substring_DTEND_ALT%=!"=="!line!" (
		set "line_new=!line:~17!"
		set "final_dtend=!line_new!"
	)

	if not "!line:%substring_SUMMARY%=!"=="!line!" (
		set "line_new=!line:SUMMARY:=!"
		set "line_new=!line_new:\=!"
		set "final_summary=!line_new!"
	) 

     set "substring_CR=!line:~0,3!"

	set "found=false"

	for %%k in ("BEG", "PRO", "VER", "CAL", "MET", "X-W", "TZI", "X-L", "TZO", "TZN", "DTS", "RRU", "END", "DTE", "DTS", "UID", "CRE", "LAS", "LOC", "SEQ", "STA", "SUM", "TRA", "DES") do (
		if "%%~k" == "!substring_CR!" (
			rem echo "gleich !substring_CR! %%~k"
			set "found=true"
		) 
	)

	if "!found!" == "false" (
		set "line_new=!line:\=!"
		call :trim final_summery_line !line_new!
	)

	if not "!line:%substring_VEVENT_END%=!"=="!line!" (
		set "event_end=true"
	) else (
		set "event_end=false"
	)


	if "!event_end!" == "true" (
		if not "!final_dtend!"=="" (
			if not "!final_dtstart!"=="" (
				if not "!final_summary!"=="" (

					set "final_dtstart_cut=!final_dtstart:~0,8!"
					set "final_dtend_cut=!final_dtend:~0,8!"

					if !final_dtstart_cut! gtr !date_from! (
						if !date_to! gtr !final_dtstart_cut! (

							set i_year=!final_dtstart_cut:~0,4!
							set i_month=!final_dtstart_cut:~4,2!
							if "!i_month:~0,1!"=="0" SET i_month=!i_month:~1!
							set i_day=!final_dtstart_cut:~6,2!
							if "!i_day:~0,1!"=="0" SET i_day=!i_day:~1!

							call :calcdate final_date
							call :calc_daylight_saving !i_year! !i_month! !i_day!
										
							if not "!final_dtstart:~9,2!" == "" (
								set final_dstart_hour=!final_dtstart:~9,2!
								if "!final_dstart_hour:~0,1!"=="0" SET final_dstart_hour=!final_dstart_hour:~1!
								set /a final_dstart_hour_de_tz=!final_dstart_hour! + !dst_deviation!
								if 10 gtr !final_dstart_hour_de_tz! SET "final_dstart_hour_de_tz=0!final_dstart_hour_de_tz!"
								set "final_start_datetime=!final_date!!final_dstart_hour_de_tz!:!final_dtstart:~11,2!"
							) else (
								set "final_start_datetime=!final_date!"
							)
							call :trim final_start_datetime !final_start_datetime!

							set i_year=!final_dtend_cut:~0,4!
							set i_month=!final_dtend_cut:~4,2!
							if "!i_month:~0,1!"=="0" SET i_month=!i_month:~1!
							set i_day=!final_dtend_cut:~6,2!
							if "!i_day:~0,1!"=="0" SET i_day=!i_day:~1!

							call :calcdate final_date
							call :calc_daylight_saving !i_year! !i_month! !i_day!

							if not "!final_dtend:~9,2!" == "" (
								set final_dtend_hour=!final_dtend:~9,2!
								if "!final_dtend_hour:~0,1!"=="0" set final_dtend_hour=!final_dtend_hour:~1!
								set /a final_dtend_hour_de_tz=!final_dtend_hour! + !dst_deviation!
								if 10 gtr !final_dtend_hour_de_tz! set "final_dtend_hour_de_tz=0!final_dtend_hour_de_tz!"
								set "final_end_datetime=!final_date!!final_dtend_hour_de_tz!:!final_dtend:~11,2!"
							) else (
								set "final_end_datetime=!final_date!"
							)
							call :trim final_end_datetime !final_end_datetime!

							echo !final_summary!!final_summery_line!	!final_dtstart_cut!!final_dstart_hour_de_tz!	!final_start_datetime!	!final_end_datetime!


						)

					)
				)
			)
		)
		set "final_summery_line="
	)
				
)) >> "%csv_file%"
%SystemRoot%\notepad.exe %csv_file%
del %csv_file%
del %cal_file%
echo 4. -Exit-
exit /b

:calcdate
call :dgetl   y m d
call :dpack p y m d
set days=MoDiMiDoFrSaSo
for /l %%o in (-0,1,0) do (
  	set /a o=p+%%o, o3=o%%7*2
  	call :dunpk y m d o
  	for %%d in (!o3!) do (
		rem setlocal& 
		if 10 gtr !d! (
			set "d_new=0!d!"
		) else (
			set "d_new=!d!"
		)
		if 10 gtr !m! (
			set "m_new=0!m!"
		) else (
			set "m_new=!m!"
		)

		set final_date=!days:~%%d,2! !d_new!.!m_new!.!y!
	)
)
endlocal& set %1=%final_date% &exit /b

:dgetl
setlocal& set "z="
rem for /f "skip=1" %%a in ('wmic os get localdatetime') do set z=!z!%%a
rem set /a y=%z:~0,4%, m=1%z:~4,2% %%100, d=1%z:~6,2% %%100
endlocal& set /a %1=!i_year!, %2=!i_month!, %3=!i_day!& exit /b

:dpack
setlocal EnableDelayedExpansion&^
set /a y=(%2)*512+(%3)*32+(%4), d=y%%32, m=y/32%%16, m3=m*3, y/=512&^
set t=xxx  0 31 59 90120151181212243273304334
set /a r=y-(12-m)/10, r=365*(y-1)+d+!t:~%m3%,3!+r/4-(r/100-r/400)-1
endlocal& set %1=%r%& exit /b

:dunpk
setlocal& set /a y=(%4)+366, y+=y/146097*3+(y%%146097-60)/36524,^
y+=y/1461*3+(y%%1461-60)/365, d=y%%366+1, y/=366
set m=1& for %%x in (31 60 91 121 152 182 213 244 274 305 335
) do if %d% gtr %%x set /a m+=1, d=%d%-%%x
endlocal& set /a %1=%y%, %2=%m%, %3=%d%& exit /b

:trim
SetLocal EnableDelayedExpansion
set params=%*
for /f "tokens=1*" %%a in ("!params!") do EndLocal & set %1=%%b
exit /b

:downloadfile
SetLocal EnableDelayedExpansion
%SystemRoot%\system32\curl.exe -k -s %1 -o %2
endlocal& set %1=%cal_url%, %2=%cal_file%& exit /b

:calc_daylight_saving
For /F "delims=" %%G In ('PowerShell -Command "&{(Get-Date -Year %1 -Month %2 -Day %3).IsDaylightSavingTime()}"') Do Set "var=%%G"
if "%var%" == "True" (
	set "dst_deviation=2"
) else (
	set "dst_deviation=1"
)
exit /b 0
