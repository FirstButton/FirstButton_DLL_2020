Option Strict Off
Option Explicit On
Module reghelp

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Integer

Public Structure cmdmapping_entry
   Dim cmd_valuename As String
   Dim cmd_valuedword As Integer
   Dim cmd_index As Integer
End Structure

Public Structure cmdmap_entry
  Dim ce_valuename As String
  Dim ce_valuelength As Integer
  Dim ce_valuedword As Integer
  Dim ce_valuehex As String
  Dim ce_IE7_cmdvalue As Integer
  Dim ce_b4_cmdvalue As Integer
  Dim ce_mapped_flag As Boolean
  Dim ce_b4_inst_pos As Integer
  Dim ce_enabled_flag As Boolean
  Dim ce_special_ext As Boolean
End Structure

 Public Structure ie7_cmdmap_entry
  Dim ce7_valuename As String
  Dim ce7_valuelength As Integer
  Dim ce7_b4_cmdvalue As Integer
  Dim ce7_IE7_cmdvalue As Integer
  Dim ce7_b4_mapped_flag As Boolean
  Dim ce7_b4_inst_pos As Integer
  Dim ce7_enabled_flag As Boolean
  Dim ce7_special_ext As Boolean
 End Structure

 Public Structure extension_array
   Dim ea_extension_name As String
 End Structure

 Public Structure toolbar_map_entry
   Dim tm_value_hi_dword As Integer
   Dim tm_value_lo_dword As Integer
   Dim tm_valuehex As String
 End Structure

Public Structure OSVERSIONINFO
  Dim dwOSVersionInfoSize As Integer
  Dim dwMajorVersion As Integer
  Dim dwMinorVersion As Integer
  Dim dwBuildNumber As Integer
  Dim dwPlatformId As Integer
  <VBFixedString(128), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=128)> Public szCSDVersion() As Char
 End Structure

 Public Const ERROR_ACCESS_DENIED As Short = 5
 Public Const ERROR_FILE_NOT_FOUND As Short = 2
 Public Const ERROR_MORE_DATA As Short = 234
 Public Const ERROR_NO_DATA As Short = 232
 Public Const ERROR_SUCCESS As Short = 0
 Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
 Public Const HKEY_CURRENT_CONFIG As Integer = &H80000005
 Public Const HKEY_CURRENT_USER As Integer = &H80000001
 Public Const HKEY_DYN_DATA As Integer = &H80000006
 Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002

 Public Const cmd_id_cutoff_value_IE7 As Integer = 11
 Public Const cmd_id_cutoff_value_IE8 As Integer = 104
 Public Const cmd_id_cutoff_value_IE9 As Integer = 104
 'Public Const cmd_id_cutoff_value_IE9 As Integer = 117      IE9RC'
 Public Const cmd_id_compare_value_IE7 As Integer = 15
 Public Const cmd_id_compare_value_IE8 As Integer = 104
 Public Const cmd_id_compare_value_IE9 As Integer = 104
 'Public Const cmd_id_compare_value_IE9 As Integer = 117     IE9RC'
 Public Const cmd_id_increment_value_IE7 As Integer = 16
 Public Const cmd_id_increment_value_IE8 As Integer = 105
 Public Const cmd_id_increment_value_IE9 As Integer = 105
 'Public Const cmd_id_increment_value_IE9 As Integer = 118   IE9RC'     
 '----------------------------------'
End Module