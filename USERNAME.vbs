oi
sub oi

Set wshShell = CreateObject( "WScript.Shell" )

strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )

WScript.Echo strName
end sub