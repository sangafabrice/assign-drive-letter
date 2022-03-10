@Echo OFF

:Main
: 1=: Disk Name
: 2=: Drive Letter to assign
: 3=: Matching Partition Number

SetLocal ENABLEDELAYEDEXPANSION

: DeviceID is an integer
Set DeviceID=
For /F "Tokens=2 Delims==" %%F In ('^
    WMIC DISKDRIVE ^
    WHERE "Model LIKE '%%%~1%%'" ^
    GET DeviceID /FORMAT:List ^| ^
    Find /I "DeviceID" ^
') Do Set DeviceID=%%~nF
If Not DEFINED DeviceID GoTo :EOF
Set DeviceID=%DeviceID:PHYSICALDRIVE=%

: List volumes
: The record of the volume attached
: to the partition starts with a star *
(
    Echo Select Disk %DeviceID%
    Echo Select Partition %~3
    Echo List Volume
) > DScript
DiskPart /S DScript > DScriptOut
Set "VolumePattern=  Volume [0-9][0-9]*  *"

: Quit if the letter is
: already assigned to the volume
Type DScriptOut | ^
FindStr /B /R /C:"\*%VolumePattern:~1%%~2  *" && ^
GoTo End

: If the letter is assigned to another partition
: copy the volume id to VolumeIDOther
For /F "Tokens=2" %%N in ('^
    Type DScriptOut ^|^
    FindStr /B /R /C:"%VolumePattern%%~2  *" ^
') Do Set VolumeIDOther=%%N

: The record of the partition
For /F "Tokens=*" %%L In ('^
    Type DScriptOut ^| ^
    FindStr /B /R /C:"\*%VolumePattern:~1%" ^
') Do Set VolumeLine=%%L

: Reassign letters
For /F "Tokens=2-3" %%N In ("%VolumeLine:~2,16%") Do (
    If DEFINED VolumeIDOther (
        If Not ""=="%%O" (
            Echo Select Volume %%N
            Echo Remove
            Echo Select Volume %VolumeIDOther%
            Echo Assign Letter=%%O
        ) Else (
            Set Letters=ABDEFGHJKLMNOPQRSTUVWXYZ
            For /F "Tokens=3" %%N in ('^
                Type DScriptOut ^|^
                FindStr /B /R /C:"%VolumePattern%[A-Z]  *" ^
            ') Do Set Letters=!Letters:%%N=!
            Echo Select Volume %VolumeIDOther%
            Echo Assign Letter=!Letters:~-1!
        )
    )
    Echo Select Volume %%N
    Echo Assign Letter=%~2
) > DScript
DiskPart /S DScript

:End
Del /F /Q DScript*
EndLocal