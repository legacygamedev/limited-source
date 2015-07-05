
AtlSampleps.dll: dlldata.obj AtlSample_p.obj AtlSample_i.obj
	link /dll /out:AtlSampleps.dll /def:AtlSampleps.def /entry:DllMain dlldata.obj AtlSample_p.obj AtlSample_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del AtlSampleps.dll
	@del AtlSampleps.lib
	@del AtlSampleps.exp
	@del dlldata.obj
	@del AtlSample_p.obj
	@del AtlSample_i.obj
