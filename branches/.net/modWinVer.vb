Option Strict Off
Option Explicit On
Option Compare Binary
Module modWinVer
	'***************************************************
	'윈도우 이름, 버전 이름, 그리고 버전 번호를 구하는 함수 입니다.
	'MSDN을 참조해서 만들었습니다.
	'제작 : 리바이
	'일시 : 2011.01.22
	'***************************************************
	
	
	'함수 선언부
	
	'UPGRADE_WARNING: OSVERSIONINFOEX 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetVersionEx Lib "kernel32"  Alias "GetVersionExA"(ByRef lpVersionInformation As OSVERSIONINFOEX) As Integer
	'UPGRADE_WARNING: SYSTEM_INFO 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Sub GetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)
	Private Declare Function GetProductInfo Lib "kernel32" (ByVal dwOSMajorVersion As Integer, ByVal dwOSMinorVersion As Integer, ByVal dwSpMajorVersion As Integer, ByVal dwSpMinorVersion As Integer, ByRef pdwReturnedProductType As Integer) As Integer
	'UPGRADE_ISSUE: 매개 변수를 'As Any'로 선언할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Object, ByVal Length As Integer)
	Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer
	Private Const SM_SERVERR2 As Short = 89
	'구조체 타입 선언부
	
	Private Structure OSVERSIONINFOEX
		Dim dwOSVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
		'UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char
		Dim wServicePackMajor As Short
		Dim wServicePackMinor As Short
		Dim wSuiteMask As Short
		Dim wProductType As Byte
		Dim wReserved As Byte
	End Structure
	
	Private Structure SYSTEM_INFO
		Dim wProcessorArchitecture As Short
		Dim wReserved As Short
		Dim dwPageSize As Integer
		Dim lpMinimumApplicationAddress As Integer
		Dim lpMaximumApplicationAddress As Integer
		Dim dwActiveProcessorMask As Integer
		Dim dwNumberOfProcessors As Integer
		Dim dwProcessorType As Integer
		Dim dwAllocationGranularity As Integer
		Dim wProcessorLevel As Short
		Dim wProcessorRevision As Short
	End Structure
	
	'상수 선언부
	
	Private Const VER_NT_WORKSTATION As Short = &O1s
	Private Const VER_NT_DOMAIN_CONTROLLER As Integer = &H2
	Private Const VER_NT_SERVER As Integer = &H3
	
	Private Const VER_SUITE_BACKOFFICE As Integer = &H4 'Microsoft BackOffice
	Private Const VER_SUITE_BLADE As Integer = &H400 'Windows Server 2003 Web Edition
	Private Const VER_SUITE_COMPUTE_SERVER As Integer = &H4000 'Windows Server 2003 Compute Cluster Edition
	Private Const VER_SUITE_DATACENTER As Integer = &H80 'Windows Server 2008 Datacenter, Windows Server 2003 Datacenter Edition, or Windows 2000 Datacenter Server
	Private Const VER_SUITE_ENTERPRISE As Integer = &H2 'Windows Server 2008 Enterprise, Windows Server 2003 Enterprise Edition, or Windows 2000 Advanced Server
	Private Const VER_SUITE_EMBEDDEDNT As Integer = &H40 'Windows XP Embedded
	Private Const VER_SUITE_PERSONAL As Integer = &H200 'Windows Vista Home Premium, Windows Vista Home Basic, or Windows XP Home Edition
	Private Const VER_SUITE_SINGLEUSERTS As Integer = &H100 'Remote Desktop
	Private Const VER_SUITE_SMALLBUSINESS As Integer = &H1 'Microsoft Small Business Server
	Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Integer = &H20 'Microsoft Small Business Server
	Private Const VER_SUITE_STORAGE_SERVER As Integer = &H2000 'Windows Storage Server 2003 R2 or Windows Storage Server 2003
	Private Const VER_SUITE_TERMINAL As Integer = &H10 'Terminal Services
	Private Const VER_SUITE_WH_SERVER As Integer = &H8000 'Windows Home Server
	
	Private Const VER_PLATFORM_WIN32s As Short = 0
	Private Const VER_PLATFORM_WIN32_WINDOWS As Short = 1
	Private Const VER_PLATFORM_WIN32_NT As Short = 2
	
	Private Const PROCESSOR_ARCHITECTURE_AMD64 As Short = 9 'x64 (AMD Or Intel)
	Private Const PROCESSOR_ARCHITECTURE_IA64 As Short = 6 'Intel Itanium - based
	Private Const PROCESSOR_ARCHITECTURE_INTEL As Short = 0 'x86
	Private Const PROCESSOR_ARCHITECTURE_UNKNOWN As Integer = &HFFFF 'Unknown
	
	Private Const PRODUCT_BUSINESS As Integer = &H6 'Business
	Private Const PRODUCT_BUSINESS_N As Integer = &H10 'Business N
	Private Const PRODUCT_CLUSTER_SERVER As Integer = &H12 'Cluster Server Edition
	Private Const PRODUCT_DATACENTER_SERVER As Integer = &H8 'Server Datacenter Edition (full installation)
	Private Const PRODUCT_DATACENTER_SERVER_CORE As Integer = &HC 'Server Datacenter Edition (core installation)
	Private Const PRODUCT_DATACENTER_SERVER_CORE_V As Integer = &H27 'Server Datacenter Edition without Hyper-V (core installation)
	Private Const PRODUCT_DATACENTER_SERVER_V As Integer = &H25 'Server Datacenter Edition without Hyper-V (full installation)
	Private Const PRODUCT_ENTERPRISE As Integer = &H4 'Enterprise
	Private Const PRODUCT_ENTERPRISE_E As Integer = &H46 'Not supported
	Private Const PRODUCT_ENTERPRISE_N As Integer = &H1B 'Enterprise N
	Private Const PRODUCT_ENTERPRISE_SERVER As Integer = &HA 'Server Enterprise Edition (full installation)
	Private Const PRODUCT_ENTERPRISE_SERVER_CORE As Integer = &HE 'Server Enterprise Edition (core installation)
	Private Const PRODUCT_ENTERPRISE_SERVER_CORE_V As Integer = &H29 'Server Enterprise Edition without Hyper-V (core installation)
	Private Const PRODUCT_ENTERPRISE_SERVER_IA64 As Integer = &HF 'Server Enterprise Edition for Itanium-based Systems
	Private Const PRODUCT_ENTERPRISE_SERVER_V As Integer = &H26 'Server Enterprise Edition without Hyper-V (full installation)
	Private Const PRODUCT_HOME_BASIC As Integer = &H2 'Home Basic
	Private Const PRODUCT_HOME_BASIC_E As Integer = &H43 'Not supported
	Private Const PRODUCT_HOME_BASIC_N As Integer = &H5 'Home Basic N
	Private Const PRODUCT_HOME_PREMIUM As Integer = &H3 'Home Premium
	Private Const PRODUCT_HOME_PREMIUM_E As Integer = &H44 'Not supported
	Private Const PRODUCT_HOME_PREMIUM_N As Integer = &H1A 'Home Premium N
	Private Const PRODUCT_HOME_PREMIUM_SERVER As Integer = &H22 'Windows Home Server 2011
	Private Const PRODUCT_HOME_SERVER As Integer = &H13 'Windows Storage Server 2008 R2 Essentials
	Private Const PRODUCT_HYPERV As Integer = &H2A 'Microsoft Hyper-V Server
	Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT As Integer = &H1E 'Windows Essential Business Server Management Server
	Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING As Integer = &H20 'Windows Essential Business Server Messaging Server
	Private Const PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY As Integer = &H1F 'Windows Essential Business Server Security Server
	Private Const PRODUCT_PROFESSIONAL As Integer = &H30 'Professional
	Private Const PRODUCT_PROFESSIONAL_E As Integer = &H45 'Not supported
	Private Const PRODUCT_PROFESSIONAL_N As Integer = &H31 'Professional N
	Private Const PRODUCT_SB_SOLUTION_SERVER As Integer = &H32 'Windows Small Business Server 2011 Essentials
	Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS As Integer = &H18 'Windows Server 2008 for Windows Essential Server Solutions
	Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS_V As Integer = &H23 'Windows Server 2008 without Hyper-V for Windows Essential Server Solutions
	Private Const PRODUCT_SERVER_FOUNDATION As Integer = &H21 'Server Foundation
	Private Const PRODUCT_SMALLBUSINESS_SERVER As Integer = &H9 'Small Business Server
	Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM As Integer = &H19 'Small Business Server Premium Edition
	Private Const PRODUCT_SOLUTION_EMBEDDEDSERVER As Integer = &H38 'Windows MultiPoint Server
	Private Const PRODUCT_STANDARD_SERVER As Integer = &H7 'Server Standard Edition (full installation)
	Private Const PRODUCT_STANDARD_SERVER_CORE As Integer = &HD 'Server Standard Edition (core installation)
	Private Const PRODUCT_STANDARD_SERVER_CORE_V As Integer = &H28 'Server Standard Edition without Hyper-V (core installation)
	Private Const PRODUCT_STANDARD_SERVER_V As Integer = &H24 'Server Standard Edition without Hyper-V (full installation)
	Private Const PRODUCT_STARTER As Integer = &HB 'Starter
	Private Const PRODUCT_STARTER_E As Integer = &H42 'Not supported
	Private Const PRODUCT_STARTER_N As Integer = &H2F 'Starter N
	Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER As Integer = &H17 'Storage Server Enterprise Edition
	Private Const PRODUCT_STORAGE_EXPRESS_SERVER As Integer = &H14 'Storage Server Express Edition
	Private Const PRODUCT_STORAGE_STANDARD_SERVER As Integer = &H15 'Storage Server Standard Edition
	Private Const PRODUCT_STORAGE_WORKGROUP_SERVER As Integer = &H16 'Storage Server Workgroup Edition
	Private Const PRODUCT_ULTIMATE As Integer = &H1 'Ultimate
	Private Const PRODUCT_ULTIMATE_E As Integer = &H47 'Not supported
	Private Const PRODUCT_ULTIMATE_N As Integer = &H1C 'Ultimate N
	Private Const PRODUCT_UNDEFINED As Integer = &H0 'An unknown product
	Private Const PRODUCT_WEB_SERVER As Integer = &H11 'Web Server Edition (full installation)
	Private Const PRODUCT_WEB_SERVER_CORE As Integer = &H1D 'Web Server Edition (core installation)
	
	'힘수 본체
	
	Public Function fGetWindowVersion() As String
        fGetWindowVersion = "Windows"
    End Function
End Module