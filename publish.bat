dotnet publish -c Release -r win-x64 -p:PublishAot=true --self-contained true -p:DebugType=None -o publish\win-x64

del /q "publish\win-x64\*.pdb"