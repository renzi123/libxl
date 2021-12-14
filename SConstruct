if ARGUMENTS.get('CC') == 'gcc':
	env=DefaultEnvironment(
		CPPFLAGS=['-I./include'],
		TARGET_ARCH='x86_64',
		#LIBS=["libxl64"],
		LINKFLAGS='-m64 -L ./lib -lxl -Wl,-rpath,./lib',
	)
else:
	env=DefaultEnvironment(
		#tools=['default','msvc','mslink','mslib'],
		platform='vc',
		MSVC_VERSION='17.0',
		CPPFLAGS='-I ./include  /O1 /Oi /Ob2 /Oy- /GR- /fp:precise /Zc:wchar_t /Zc:forScope /EHsc',
		TARGET_ARCH='x86_64',
		LIBPATH=['./lib'],
		LIBS=["libxl"],
		LINKFLAGS=['/subsystem:console']
	)

Program('invoice',['invoice.cpp'])

