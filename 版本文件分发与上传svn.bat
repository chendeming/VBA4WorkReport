set src=H:\aa

set dst1=F:\imdocs\01_����\���ȸ���
set dst2=F:\imdocs\01_����\���ȸ���

xcopy %src% %dst1%  /s/y/D
echo xcopy %src% %dst2%  /s/y/D

echo TortoiseProc.exe /command:commit /path:%dst1% /logmsg:"HLM committed on %DATE% %TIME%" /closeonend:0
echo TortoiseProc.exe /command:commit /path:%dst2% /logmsg:"HLM committed on %DATE% %TIME%" /closeonend:0

pause