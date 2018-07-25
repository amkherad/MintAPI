#ifndef __SHELL32_H__
#define __SHELL32_H__


// interface marshaling definitions
//#define MARSHALINTERFACE_MIN 500 // minimum number of bytes for interface marshl

//typedef enum API_CLSCTX {
//} API_CLSCTX;

#pragma pack(4)

typedef struct API_NOTIFYICONDATA {
	long cbSize;
	long hwnd;
	long uID;
	long uFlags;
	long uCallbackMessage;
	long hIcon;
	char szTip[64];// As String * 64
} API_NOTIFYICONDATA;

typedef struct API_SHFILEINFO {
    long hIcon;//out: icon
    long iIcon;//out: icon index
    long dwAttributes;//out: SFGAO_ flags
    char szDisplayName[256];//String * MAX_PATH '  out: display name (or path)
    char szTypeName[80];//String * 80         '  out: type name
} API_SHFILEINFO;

typedef struct API_SHFILEOPSTRUCT {
	long hwnd;
	long wFunc;
	String pFrom;
	String pTo;
	short fFlags;
	long fAnyOperationsAborted;
	long hNameMappings;
	String lpszProgressTitle; // only used if FOF_SIMPLEPROGRESS
} API_SHFILEOPSTRUCT;

typedef struct API_SHNAMEMAPPING {
	String pszOldPath;
	String pszNewPath;
	long cchOldPath;
	long cchNewPath;
} API_SHNAMEMAPPING;

typedef struct API_SHELLEXECUTEINFO {
	long cbSize;
	long fMask;
	long hwnd;
	String lpVerb;
	String lpFile;
	String lpParameters;
	String lpDirectory;
	long nShow;
	long hInstApp;
	//Optional fields
	long lpIDList;
	String lpClass;
	long hkeyClass;
	long dwHotKey;
	long hIcon;
	long hProcess;
} API_SHELLEXECUTEINFO;

typedef struct API_BROWSEINFO {
    long hOwner;
    long pidlRoot;
    String pszDisplayName;
    String lpszTitle;
    long ulFlags;
    long lpfn;
    long lParam;
    long iImage;
} API_BROWSEINFO;

typedef struct API_AppBarData
{
    long cbSize;
    long hWnd;
    long uCallbackMessage;
    long uEdge;
    API_RECT rc;
    long lParam; // message specific
} API_AppBarData;

typedef struct API_SH_ITEM_ID {
    long cb;
    Byte abID;
} API_SH_ITEM_ID;
typedef struct API_ITEMIDLIST {
    API_SH_ITEM_ID mkid;
} API_ITEMIDLIST;

#pragma pack()
[
    dllname("shell32.dll"),
    helpstring("Access to API functions within the Shell32.dll system file.")
]
module Shell32 {
[entry("ShellAboutA")]
    long	API_ShellAbout([in] long hwnd, [in] String szApp, [in] String szOtherStuff, [in] long hIcon);
[entry("ShellExecuteA")]
    long	API_ShellExecute([in] long hwnd, [in] String lpOperation, [in] String lpFile, [in] String lpParameters, [in] String lpDirectory, [in] long nShowCmd);
//========================================
[entry("Shell_NotifyIconA")]
    long	API_ShellNotifyIcon([in] long dwMessage, [in] API_NOTIFYICONDATA* lpData);
[entry("SHFileOperationA")]
    long	API_SHFileOperation([in] API_SHFILEOPSTRUCT* lpFileOp);
[entry("SHGetFileInfoA")]
    long	API_SHGetFileInfo([in] String pszPath, [in] long dwFileAttributes, [in] API_SHFILEINFO* psfi, [in] long cbFileInfo, [in] long uFlags);
[entry("SHGetSpecialFolderLocation")]
    long	API_SHGetSpecialFolderLocation([in] long hWndOwner, [in] long nFolder, [in] API_ITEMIDLIST* pidl);
[entry("#162")]
    long	API_SHSimpleIDListFromPath([in] String szPath);
[entry("SHGetPathFromIDList")]
    long	API_SHGetPathFromIDList([in] long pidList, [in] String lpBuffer);
[entry("SHBrowseForFolder")]
    long	API_SHBrowseForFolder([in] API_BROWSEINFO* lpbi);
//========================================
[entry("SHAppBarMessage")]
    long	API_SHAppBarMessage([in] long dwMessage, [in] API_APPBARDATA pData);
[entry("SHGetNewLinkInfoA")]
    long	API_SHGetNewLinkInfo([in] String pszLinkto, [in] String pszDir, [in] String pszName, [in] long* pfMustCopy, [in] long uFlags);
[entry("SHFreeNameMappings")]
    long	API_SHFreeNameMappings([in] long hNameMappings);
//========================================
[entry("SHGetMalloc")]
    HRESULT API_SHGetMalloc([out, retval] IMalloc** ppMalloc);
//========================================
}

#endif //__SHELL32_H__