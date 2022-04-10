::------------------------------------------------------------------------------------------------------------
:: runas /user:%computername%\administrator "jre-8u321-windows-x64.exe installdir=c:\printstation /s"
:: runas /user:%computername%\administrator "jre-8u321-windows-x64.exe installdir=c:\printstation /s"

::------------------------------------------------------------------------------------------------------------
:: forfiles -s -m *.jar -c "cp @file c:\printstation"
:: forfiles -s -m *.* -c "cmd /c echo The extension of @file is 0x09@ext"
mkdir c:\printstation\bin
forfiles -s -m *.jar -c "cmd /c copy @file c:\printstation\bin"

::------------------------------------------------------------------------------------------------------------
:: forfiles -s -m *.jar -c "cp @file c:\printstation"
:: forfiles -s -m *.* -c "cmd /c echo The extension of @file is 0x09@ext"
mkdir c:\printstation\res
forfiles -s -m *.xlsx -c "cmd /c copy @file c:\printstation\res"

::------------------------------------------------------------------------------------------------------------
:: forfiles -s -m *.jar -c "cp @file c:\printstation"
:: forfiles -s -m *.* -c "cmd /c echo The extension of @file is 0x09@ext"
forfiles -s -m *.pdf  -c "cmd /c copy @file c:\printstation\res"