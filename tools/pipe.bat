@echo off
mkisofs.exe -J -l -r data | cdrecord.exe dev=0,2,0 gracetime=2 -v driveropts=burnfree speed=8 -eject -tao -data -
pause