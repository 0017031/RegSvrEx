RegSvrEx
========

Overview
--------

This is a tiny tool capable of registering Windows COM DLLs (and EXEs in 32 bit
version) for the current user only, something that is inexplicably not
supported by the standard `regsvr32.exe` tool included in Windows itself.

See the [original article](http://www.codeproject.com/Articles/3505/RegSvrEx-An-Enchanced-COM-Server-Registration-Util)
for more information.


Instructions
------------

Get the sources and build them using MSVS 2017 (earlier versions can probably
be used as well, but the project files in the repository requires this one) or
just grab the
[binary](https://github.com/vadz/RegSvrEx/releases/download/v1.0.0.1/RegSvrEx.exe)
as this is normally all you need.


Credits
-------

The program was written by [Rama Krishna Vavilala](http://www.codeproject.com/script/Membership/View.aspx?mid=15383)
and I (Vadim Zeitlin) have only built it in 64 bits.

Extra Notes:
-------
https://www.codeproject.com/Messages/3361407/Solution-to-problems-with-Windows-Vista-Windows-7.aspx
The current/first version of regsvrex doesn't work with Windows Vista or later. This is due to increased security. For more information see Microsoft KB article 935200: Error message when an application calls the RegisterTypeLib API to register a type library in Windows Vista: "Access denied".

The article describes two solutions:
1. Set environment variable OAPERUSERTLIBREG to 1:
Hide   Copy Code
set OAPERUSERTLIBREG=1
regsvrex <parameters>

2. Add a call to OaEnablePerUserTLibRegistration to the regsvrex source code.

typedef void (*MYPROC)(); 
HINSTANCE hModForOle;

HRESULT CallOaEnable()
{ 
	HRESULT hr = S_OK;

	hModForOle = LoadLibrary(TEXT("Oleaut32.dll"));
	MYPROC ProcAdd; 

	if (hModForOle != NULL)
	{
		ProcAdd = (MYPROC) GetProcAddress(hModForOle, "OaEnablePerUserTLibRegistration"); 
		
		if (NULL != ProcAdd)
		{
			(ProcAdd)();
			hr = S_OK;
		}
		else
			hr = GetHresultFromWin32();
	}
	else
		hr = GetHresultFromWin32();
		
	return hr;
}

before calling RegisterExe/RegisterDll, you should call CallOaEnable(). and after registering dll/exe, you should also call FreeLibrary(hModForOle). if you call FreeLibrary(hModForOle) before registering, then you'll meet the 'access denied' problem again.



