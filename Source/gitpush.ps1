# Clean up binary folder since we will be generating new binaries
rm ..\Binary\*.* -Force -Recurse
rm -Force bin
rm -Force obj

# build new version
msbuild /verbosity:minimal .\SharepointListClient.sln 
# Compress bin\debug and move to binary folder
cd bin\debug
# clean the .config file from sensitive info
../../cleanconfig.ps1
cmd /c "c:\Program Files\7-Zip\7z.exe" a -tzip SharepointTaskReminder.zip *.* 
cd ../..
cmd /c copy bin\debug\SharepointTaskReminder.zip ..\Binary\ /Y

# remove bin and obj before comitting to git
rm -Force bin
rm -Force obj

# backup app.config file and clean it
$backup = gc app.config
./cleanconfig.ps1

cd ..
git add -A *.*
git commit -a -m "New binary"
git push origin master

# restore app.config
cd Source
sc app.config $backup 