#ifndef __KERNELIO_H__
#define __KERNELIO_H__


#pragma pack(4)
typedef struct os_iobuf {
    String ib_ptr;
    Long   ib_cnt;
    String ib_base;
    Long   ib_flag;
    Long   ib_file;
    Long   ib_charbuf;
    Long   ib_bufsiz;
    String ib_tmpfname;
} os_iobuf;

typedef struct API_FILETIME {
  Long dwLowDateTime;
  Long dwHighDateTime;
} API_FILETIME;

typedef enum API_DOSDEVICE_FLAGS {	
	DDD_RAW_TARGET_PATH         = 0x00000001,
	DDD_REMOVE_DEFINITION       = 0x00000002,
	DDD_EXACT_MATCH_ON_REMOVE   = 0x00000004,
	DDD_NO_BROADCAST_SYSTEM     = 0x00000008,
} API_DOSDEVICE_FLAGS;


#pragma pack()
[
    dllname("Kernel32.dll"),
    helpstring("Access to API functions within the Kernel32.dll system file.")
]
module KernelIO {
[entry("AreFileApisANSI"), usesgetlasterror]
    long API_AreFileApisANSI();
//========================================
[entry("ReadFile"), usesgetlasterror]
    long API_ReadFileOVR([in] long hFile, [out] Any lpBuffer, [in] long nNumberOfBytesToRead, [out] long* lpNumberOfBytesRead, [in] API_OVERLAPPED lpOverlapped);
[entry("ReadFile"), usesgetlasterror]
    long API_ReadFile([in] long hFile, [out] Any lpBuffer, [in] long nNumberOfBytesToRead, [out] long* lpNumberOfBytesRead, [in] Any lpOverlapped);
[entry("ReadFileEx"), usesgetlasterror]
    long API_ReadFileExOVR([in] long hFile, [out] Any lpBuffer, [in] long nNumberOfBytesToRead, [out] long* lpNumberOfBytesRead, [in] API_OVERLAPPED lpOverlapped, [in] long lpCompletionRoutine);
[entry("ReadFileEx"), usesgetlasterror]
    long API_ReadFileEx([in] long hFile, [out] Any lpBuffer, [in] long nNumberOfBytesToRead, [out] long* lpNumberOfBytesRead, [in] Any lpOverlapped, [in] long lpCompletionRoutine);
[entry("WriteFile"), usesgetlasterror]
    long API_WriteFileOVR([in] long hFile, [out] Any lpBuffer, [in] long nNumberOfBytesToWrite, [out] long* lpNumberOfBytesWritten, [in] API_OVERLAPPED lpOverlapped);
[entry("WriteFile"), usesgetlasterror]
    long API_WriteFile([in] long hFile, [out] Any lpBuffer, [in] long nNumberOfBytesToWrite, [out] long* lpNumberOfBytesWritten, [in] Any lpOverlapped);
[entry("WriteFileEx"), usesgetlasterror]
    long API_WriteFileExOVR([in] long hFile, [out] Any lpBuffer, [in] long nNumberOfBytesToWrite, [out] long* lpNumberOfBytesWritten, [in] API_OVERLAPPED lpOverlapped, [out] long* lpCompletionRoutine);
[entry("WriteFileEx"), usesgetlasterror]
    long API_WriteFileEx([in] long hFile, [out] Any lpBuffer, [in] long nNumberOfBytesToWrite, [out] long* lpNumberOfBytesWritten, [in] Any lpOverlapped, [out] long* lpCompletionRoutine);
//========================================
[entry("FlushFileBuffers"), usesgetlasterror]
    long API_FlushFileBuffers([in] long hFile);
//========================================
[entry("ReadFileScatter"), usesgetlasterror]
    long API_ReadFileScatter([in] long hFile, [out] API_FILE_SEGMENT_ELEMENT* aSegmentArray, [in] long nNumberOfBytesToRead, [out] long* lpReserved, [in] API_OVERLAPPED lpOverlapped, [in] long lpCompletionRoutine);
[entry("WriteFileGather"), usesgetlasterror]
    long API_WriteFileGather([in] long hFile, [out] API_FILE_SEGMENT_ELEMENT* aSegmentArray, [in] long nNumberOfBytesToWrite, [out] long* lpReserved, [in] API_OVERLAPPED lpOverlapped, [in] long lpCompletionRoutine);
//========================================
[entry("GetStdHandle"), usesgetlasterror]
    long API_GetStdHandle([in] long nStdHandle);
[entry("SetStdHandle"), usesgetlasterror]
    long API_SetStdHandle([in] long nStdHandle, [in] long nHandle);
//========================================
[entry("FileTimeToSystemTime"), usesgetlasterror]
    long API_FileTimeToSystemTime([out] API_FILETIME* lpFileTime, [out] API_SYSTEMTIME* lpSystemTime);
[entry("FileTimeToLocalFileTime"), usesgetlasterror]
    long API_FileTimeToLocalFileTime([out] API_FILETIME* lpFileTime, [out] API_FILETIME* lpLocalFileTime);
[entry("FileTimeToDosDateTime"), usesgetlasterror]
    long API_FileTimeToDosDateTime([out] API_FILETIME* lpFileTime, [out] Integer* lpFatDate, [out] Integer* lpFatTime);
}

#pragma pack()
[
    dllname("Kernel32.dll"),
    helpstring("Access to API functions within the Kernel32.dll system file.")
]
module KernelDiskDrive {
[entry("GetLogicalDriveStringsA"), usesgetlasterror]
    long API_GetLogicalDriveStrings([in] long nBufferLength, [in] String lpBuffer);
[entry("GetLogicalDriveStringsW"), usesgetlasterror]
    long API_GetLogicalDriveStringsUnicode([in] long nBufferLength, [in] WString lpBuffer);

[entry("GetVolumeInformationA"), usesgetlasterror]
    long API_GetVolumeInformation([in] String lpRootPathName, [in] String lpVolumeNameBuffer, [in] long nVolumeNameSize, [out] long* lpVolumeSerialNumber, [out] long* lpMaximumComponentLength, [out] long* lpFileSystemFlags, [in] String lpFileSystemNameBuffer, [in] long nFileSystemNameSize);
[entry("GetVolumeInformationW"), usesgetlasterror]
    long API_GetVolumeInformationUnicode([in] WString lpRootPathName, [in] WString lpVolumeNameBuffer, [in] long nVolumeNameSize, [out] long* lpVolumeSerialNumber, [out] long* lpMaximumComponentLength, [out] long* lpFileSystemFlags, [in] WString lpFileSystemNameBuffer, [in] long nFileSystemNameSize);

[entry("SetVolumeLabelA"), usesgetlasterror]
    long API_SetVolumeLabel([in] String lpRootPathName, [in] String lpVolumeName);
[entry("SetVolumeLabelW"), usesgetlasterror]
    long API_SetVolumeLabelUnicode([in] String lpRootPathName, [in] String lpVolumeName);

[entry("DefineDosDeviceA"), usesgetlasterror]
    long API_DefineDosDevice([in] API_DOSDEVICE_FLAGS dwFlags, [in] String lpDeviceName, [in] String lpTargetPath);
[entry("DefineDosDeviceW"), usesgetlasterror]
    long API_DefineDosDeviceUnicode([in] API_DOSDEVICE_FLAGS dwFlags, [in] WString lpDeviceName, [in] WString lpTargetPath);

[entry("SetVolumeMountPoint"), usesgetlasterror]
	long API_SetVolumeMountPoint([in] LPSTR lpszVolumeMountPoint, [in] LPSTR lpszVolumeName);

[entry("GetDriveTypeA"), usesgetlasterror]
    long API_GetDriveType([in] String nDrive);
[entry("GetDriveTypeW"), usesgetlasterror]
    long API_GetDriveTypeUnicode([in] WString nDrive);

//========================================
}

#endif //__KERNELIO_H__