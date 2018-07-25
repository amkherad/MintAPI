Attribute VB_Name = "win_errors"
Option Explicit
Option Base 0

''Constant values exported from winerror.h
''By MintAPI's "miccy Ultimate Tools".

'
'  Values are 32 bit values layed out as follows:
'
'   3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
'   1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
'  +---+-+-+-----------------------+-------------------------------+
'  |Sev|C|R|     Facility          |               Code            |
'  +---+-+-+-----------------------+-------------------------------+
'
'  where
'
'      Sev - is the severity code
'
'          00 - Success
'          01 - Informational
'          10 - Warning
'          11 - Error
'
'      C - is the Customer code flag
'
'      R - is a reserved bit
'
'      Facility - is the facility code
'
'      Code - is the facility's status code
'
'====================
' Define the facility codes
'
Public Const FACILITY_WINDOWSUPDATE           As Long = 36
Public Const FACILITY_WINDOWS_CE              As Long = 24
Public Const FACILITY_WINDOWS                 As Long = 8
Public Const FACILITY_URT                     As Long = 19
Public Const FACILITY_UMI                     As Long = 22
Public Const FACILITY_SXS                     As Long = 23
Public Const FACILITY_STORAGE                 As Long = 3
Public Const FACILITY_STATE_MANAGEMENT        As Long = 34
Public Const FACILITY_SSPI                    As Long = 9
Public Const FACILITY_SCARD                   As Long = 16
Public Const FACILITY_SETUPAPI                As Long = 15
Public Const FACILITY_SECURITY                As Long = 9
Public Const FACILITY_RPC                     As Long = 1
Public Const FACILITY_WIN32                   As Long = 7
Public Const FACILITY_CONTROL                 As Long = 10
Public Const FACILITY_NULL                    As Long = 0
Public Const FACILITY_METADIRECTORY           As Long = 35
Public Const FACILITY_MSMQ                    As Long = 14
Public Const FACILITY_MEDIASERVER             As Long = 13
Public Const FACILITY_INTERNET                As Long = 12
Public Const FACILITY_ITF                     As Long = 4
Public Const FACILITY_HTTP                    As Long = 25
Public Const FACILITY_DPLAY                   As Long = 21
Public Const FACILITY_DISPATCH                As Long = 2
Public Const FACILITY_DIRECTORYSERVICE        As Long = 37
Public Const FACILITY_CONFIGURATION           As Long = 33
Public Const FACILITY_COMPLUS                 As Long = 17
Public Const FACILITY_CERT                    As Long = 11
Public Const FACILITY_BACKGROUNDCOPY          As Long = 32
Public Const FACILITY_ACS                     As Long = 20
Public Const FACILITY_AAF                     As Long = 18


'=============================
' Define the severity codes
'=============================
Public Const SEVC_SUCCESS                     As Long = &H0
Public Const SEVC_INFORMATIONAL               As Long = &H40000000
Public Const SEVC_WARNING                     As Long = &H80000000
Public Const SEVC_ERROR                       As Long = &HC0000000

Public Const ERROR_USERERROR                  As Long = SEVC_INFORMATIONAL

Public Const ERROR_SUCCESS                    As Long = 0       'The operation completed successfully.

'Public Const NO_ERROR                         As Long = 0       ' dderror
Public Const SEC_E_OK                         As Long = &H0

Public Const NO_VALUE                         As Long = 0


'=============================
' Define the codes
'=============================

Public Const ERROR_INVALID_FUNCTION           As Long = 1       'Incorrect function.  // dderror
Public Const ERROR_FILE_NOT_FOUND             As Long = 2       'The system cannot find the file specified.

Public Const ERROR_PATH_NOT_FOUND             As Long = 3       'The system cannot find the path specified.
Public Const ERROR_TOO_MANY_OPEN_FILES        As Long = 4       'The system cannot open the file.
Public Const ERROR_ACCESS_DENIED              As Long = 5       'Access is denied.
Public Const ERROR_INVALID_HANDLE             As Long = 6       'The handle is invalid.
Public Const ERROR_ARENA_TRASHED              As Long = 7       'The storage control blocks were destroyed.
Public Const ERROR_NOT_ENOUGH_MEMORY          As Long = 8       'Not enough storage is available to process this command.  // dderror
Public Const ERROR_INVALID_BLOCK              As Long = 9       'The storage control block address is invalid.
Public Const ERROR_BAD_ENVIRONMENT            As Long = 10      'The environment is incorrect.
Public Const ERROR_BAD_FORMAT                 As Long = 11      'An attempt was made to load a program with an incorrect format.
Public Const ERROR_INVALID_ACCESS             As Long = 12      'The access code is invalid.
Public Const ERROR_INVALID_DATA               As Long = 13      'The data is invalid.
Public Const ERROR_OUTOFMEMORY                As Long = 14      'Not enough storage is available to complete this operation.
Public Const ERROR_INVALID_DRIVE              As Long = 15      'The system cannot find the drive specified.
Public Const ERROR_CURRENT_DIRECTORY          As Long = 16      'The directory cannot be removed.
Public Const ERROR_NOT_SAME_DEVICE            As Long = 17      'The system cannot move the file to a different disk drive.
Public Const ERROR_NO_MORE_FILES              As Long = 18      'There are no more files.
Public Const ERROR_WRITE_PROTECT              As Long = 19      'The media is write protected.
Public Const ERROR_BAD_UNIT                   As Long = 20      'The system cannot find the device specified.
Public Const ERROR_NOT_READY                  As Long = 21      'The device is not ready.
Public Const ERROR_BAD_COMMAND                As Long = 22      'The device does not recognize the command.
Public Const ERROR_CRC                        As Long = 23      'Data error (cyclic redundancy check).
Public Const ERROR_BAD_LENGTH                 As Long = 24      'The program issued a command but the command length is incorrect.
Public Const ERROR_SEEK                       As Long = 25      'The drive cannot locate a specific area or track on the disk.
Public Const ERROR_NOT_DOS_DISK               As Long = 26      'The specified disk or diskette cannot be accessed.
Public Const ERROR_SECTOR_NOT_FOUND           As Long = 27      'The drive cannot find the sector requested.
Public Const ERROR_OUT_OF_PAPER               As Long = 28      'The printer is out of paper.
Public Const ERROR_WRITE_FAULT                As Long = 29      'The system cannot write to the specified device.
Public Const ERROR_READ_FAULT                 As Long = 30      'The system cannot read from the specified device.
Public Const ERROR_GEN_FAILURE                As Long = 31      'A device attached to the system is not functioning.
Public Const ERROR_SHARING_VIOLATION          As Long = 32      'The process cannot access the file because it is being used by another process.
Public Const ERROR_LOCK_VIOLATION             As Long = 33      'The process cannot access the file because another process has locked a portion of the file.
Public Const ERROR_WRONG_DISK                 As Long = 34      'The wrong diskette is in the drive.'Insert %2 (Volume Serial Number: %3) into drive %1.
Public Const ERROR_SHARING_BUFFER_EXCEEDED    As Long = 36      'Too many files opened for sharing.
Public Const ERROR_HANDLE_EOF                 As Long = 38      'Reached the end of the file.
Public Const ERROR_HANDLE_DISK_FULL           As Long = 39      'The disk is full.
Public Const ERROR_NOT_SUPPORTED              As Long = 50      'The request is not supported.
Public Const ERROR_REM_NOT_LIST               As Long = 51      'Windows cannot find the network path. Verify that the network path is correct and the destination computer is not busy or turned off. If Windows still cannot find the network path, contact your network administrator.
Public Const ERROR_DUP_NAME                   As Long = 52      'You were not connected because a duplicate name exists on the network. Go to System in Control Panel to change the computer name and try again.
Public Const ERROR_BAD_NETPATH                As Long = 53      'The network path was not found.
Public Const ERROR_NETWORK_BUSY               As Long = 54      'The network is busy.
Public Const ERROR_DEV_NOT_EXIST              As Long = 55      'The specified network resource or device is no longer available.    // dderror
Public Const ERROR_TOO_MANY_CMDS              As Long = 56      'The network BIOS command limit has been reached.
Public Const ERROR_ADAP_HDW_ERR               As Long = 57      'A network adapter hardware error occurred.
Public Const ERROR_BAD_NET_RESP               As Long = 58      'The specified server cannot perform the requested operation.
Public Const ERROR_UNEXP_NET_ERR              As Long = 59      'An unexpected network error occurred.
Public Const ERROR_BAD_REM_ADAP               As Long = 60      'The remote adapter is not compatible.
Public Const ERROR_PRINTQ_FULL                As Long = 61      'The printer queue is full.
Public Const ERROR_NO_SPOOL_SPACE             As Long = 62      'Space to store the file waiting to be printed is not available on the server.
Public Const ERROR_PRINT_CANCELLED            As Long = 63      'Your file waiting to be printed was deleted.
Public Const ERROR_NETNAME_DELETED            As Long = 64      'The specified network name is no longer available.
Public Const ERROR_NETWORK_ACCESS_DENIED      As Long = 65      'Network access is denied.
Public Const ERROR_BAD_DEV_TYPE               As Long = 66      'The network resource type is not correct.
Public Const ERROR_BAD_NET_NAME               As Long = 67      'The network name cannot be found.
Public Const ERROR_TOO_MANY_NAMES             As Long = 68      'The name limit for the local computer network adapter card was exceeded.
Public Const ERROR_TOO_MANY_SESS              As Long = 69      'The network BIOS session limit was exceeded.
Public Const ERROR_SHARING_PAUSED             As Long = 70      'The remote server has been paused or is in the process of being started.
Public Const ERROR_REQ_NOT_ACCEP              As Long = 71      'No more connections can be made to this remote computer at this time because there are already as many connections as the computer can accept.
Public Const ERROR_REDIR_PAUSED               As Long = 72      'The specified printer or disk device has been paused.
Public Const ERROR_FILE_EXISTS                As Long = 80      'The file exists.
Public Const ERROR_CANNOT_MAKE                As Long = 82      'The directory or file cannot be created.
Public Const ERROR_FAIL_I24                   As Long = 83      'Fail on INT 24.
Public Const ERROR_OUT_OF_STRUCTURES          As Long = 84      'Storage to process this request is not available.
Public Const ERROR_ALREADY_ASSIGNED           As Long = 85      'The local device name is already in use.
Public Const ERROR_INVALID_PASSWORD           As Long = 86      'The specified network password is not correct.
Public Const ERROR_INVALID_PARAMETER          As Long = 87      'The parameter is incorrect.    // dderror
Public Const ERROR_NET_WRITE_FAULT            As Long = 88      'A write fault occurred on the network.
Public Const ERROR_NO_PROC_SLOTS              As Long = 89      'The system cannot start another process at this time.
Public Const ERROR_TOO_MANY_SEMAPHORES        As Long = 100     'Cannot create another system semaphore.
Public Const ERROR_EXCL_SEM_ALREADY_OWNED     As Long = 101     'The exclusive semaphore is owned by another process.
Public Const ERROR_SEM_IS_SET                 As Long = 102     'The semaphore is set and cannot be closed.
Public Const ERROR_TOO_MANY_SEM_REQUESTS      As Long = 103     'The semaphore cannot be set again.
Public Const ERROR_INVALID_AT_INTERRUPT_TIME  As Long = 104     'Cannot request exclusive semaphores at interrupt time.
Public Const ERROR_SEM_OWNER_DIED             As Long = 105     'The previous ownership of this semaphore has ended.
Public Const ERROR_SEM_USER_LIMIT             As Long = 106     'Insert the diskette for drive %1.
Public Const ERROR_DISK_CHANGE                As Long = 107     'The program stopped because an alternate diskette was not inserted.
Public Const ERROR_DRIVE_LOCKED               As Long = 108     'The disk is in use or locked by another process.
Public Const ERROR_BROKEN_PIPE                As Long = 109     'The pipe has been ended.
Public Const ERROR_OPEN_FAILED                As Long = 110     'The system cannot open the device or file specified.
Public Const ERROR_BUFFER_OVERFLOW            As Long = 111     'The file name is too long.
Public Const ERROR_DISK_FULL                  As Long = 112     'There is not enough space on the disk.
Public Const ERROR_NO_MORE_SEARCH_HANDLES     As Long = 113     'No more internal file identifiers available.
Public Const ERROR_INVALID_TARGET_HANDLE      As Long = 114     'The target internal file identifier is incorrect.
Public Const ERROR_INVALID_CATEGORY           As Long = 117     'The IOCTL call made by the application program is not correct.
Public Const ERROR_INVALID_VERIFY_SWITCH      As Long = 118     'The verify-on-write switch parameter value is not correct.
Public Const ERROR_BAD_DRIVER_LEVEL           As Long = 119     'The system does not support the command requested.
Public Const ERROR_CALL_NOT_IMPLEMENTED       As Long = 120     'This function is not supported on this system.
Public Const ERROR_SEM_TIMEOUT                As Long = 121     'The semaphore timeout period has expired.
Public Const ERROR_INSUFFICIENT_BUFFER        As Long = 122     'The data area passed to a system call is too small.    // dderror
Public Const ERROR_INVALID_NAME               As Long = 123     'The filename, directory name, or volume label syntax is incorrect.    // dderror
Public Const ERROR_INVALID_LEVEL              As Long = 124     'The system call level is not correct.
Public Const ERROR_NO_VOLUME_LABEL            As Long = 125     'The disk has no volume label.
Public Const ERROR_MOD_NOT_FOUND              As Long = 126     'The specified module could not be found.
Public Const ERROR_PROC_NOT_FOUND             As Long = 127     'The specified procedure could not be found.
Public Const ERROR_WAIT_NO_CHILDREN           As Long = 128     'There are no child processes to wait for.
Public Const ERROR_CHILD_NOT_COMPLETE         As Long = 129     'The %1 application cannot be run in Win32 mode.
Public Const ERROR_DIRECT_ACCESS_HANDLE       As Long = 130     'Attempt to use a file handle to an open disk partition for an operation other than raw disk I/O.
Public Const ERROR_NEGATIVE_SEEK              As Long = 131     'An attempt was made to move the file pointer before the beginning of the file.
Public Const ERROR_SEEK_ON_DEVICE             As Long = 132     'The file pointer cannot be set on the specified device or file.
Public Const ERROR_IS_JOIN_TARGET             As Long = 133     'A JOIN or SUBST command cannot be used for a drive that contains previously joined drives.
Public Const ERROR_IS_JOINED                  As Long = 134     'An attempt was made to use a JOIN or SUBST command on a drive that has already been joined.
Public Const ERROR_IS_SUBSTED                 As Long = 135     'An attempt was made to use a JOIN or SUBST command on a drive that has already been substituted.
Public Const ERROR_NOT_JOINED                 As Long = 136     'The system tried to delete the JOIN of a drive that is not joined.
Public Const ERROR_NOT_SUBSTED                As Long = 137     'The system tried to delete the substitution of a drive that is not substituted.
Public Const ERROR_JOIN_TO_JOIN               As Long = 138     'The system tried to join a drive to a directory on a joined drive.
Public Const ERROR_SUBST_TO_SUBST             As Long = 139     'The system tried to substitute a drive to a directory on a substituted drive.
Public Const ERROR_JOIN_TO_SUBST              As Long = 140     'The system tried to join a drive to a directory on a substituted drive.
Public Const ERROR_SUBST_TO_JOIN              As Long = 141     'The system tried to SUBST a drive to a directory on a joined drive.
Public Const ERROR_BUSY_DRIVE                 As Long = 142     'The system cannot perform a JOIN or SUBST at this time.
Public Const ERROR_SAME_DRIVE                 As Long = 143     'The system cannot join or substitute a drive to or for a directory on the same drive.
Public Const ERROR_DIR_NOT_ROOT               As Long = 144     'The directory is not a subdirectory of the root directory.
Public Const ERROR_DIR_NOT_EMPTY              As Long = 145     'The directory is not empty.
Public Const ERROR_IS_SUBST_PATH              As Long = 146     'The path specified is being used in a substitute.
Public Const ERROR_IS_JOIN_PATH               As Long = 147     'Not enough resources are available to process this command.
Public Const ERROR_PATH_BUSY                  As Long = 148     'The path specified cannot be used at this time.
Public Const ERROR_IS_SUBST_TARGET            As Long = 149     'An attempt was made to join or substitute a drive for which a directory on the drive is the target of a previous substitute.
Public Const ERROR_SYSTEM_TRACE               As Long = 150     'System trace information was not specified in your CONFIG.SYS file, or tracing is disallowed.
Public Const ERROR_INVALID_EVENT_COUNT        As Long = 151     'The number of specified semaphore events for DosMuxSemWait is not correct.
Public Const ERROR_TOO_MANY_MUXWAITERS        As Long = 152     'DosMuxSemWait did not execute; too many semaphores are already set.
Public Const ERROR_INVALID_LIST_FORMAT        As Long = 153     'The DosMuxSemWait list is not correct.
Public Const ERROR_LABEL_TOO_LONG             As Long = 154     'The volume label you entered exceeds the label character limit of the target file system.
Public Const ERROR_TOO_MANY_TCBS              As Long = 155     'Cannot create another thread.
Public Const ERROR_SIGNAL_REFUSED             As Long = 156     'The recipient process has refused the signal.
Public Const ERROR_DISCARDED                  As Long = 157     'The segment is already discarded and cannot be locked.
Public Const ERROR_NOT_LOCKED                 As Long = 158     'The segment is already unlocked.
Public Const ERROR_BAD_THREADID_ADDR          As Long = 159     'The address for the thread ID is not correct.
Public Const ERROR_BAD_ARGUMENTS              As Long = 160     'One or more arguments are not correct.
Public Const ERROR_BAD_PATHNAME               As Long = 161     'The specified path is invalid.
Public Const ERROR_SIGNAL_PENDING             As Long = 162     'A signal is already pending.
Public Const ERROR_MAX_THRDS_REACHED          As Long = 164     'No more threads can be created in the system.
Public Const ERROR_LOCK_FAILED                As Long = 167     'Unable to lock a region of a file.
Public Const ERROR_BUSY                       As Long = 170     'The requested resource is in use.    // dderror
Public Const ERROR_CANCEL_VIOLATION           As Long = 173     'A lock request was not outstanding for the supplied cancel region.
Public Const ERROR_ATOMIC_LOCKS_NOT_SUPPORTED As Long = 174     'The file system does not support atomic changes to the lock type.
Public Const ERROR_INVALID_SEGMENT_NUMBER     As Long = 180     'The system detected a segment number that was not correct.
Public Const ERROR_INVALID_ORDINAL            As Long = 182     'The operating system cannot run %1.
Public Const ERROR_ALREADY_EXISTS             As Long = 183     'Cannot create a file when that file already exists.
Public Const ERROR_INVALID_FLAG_NUMBER        As Long = 186     'The flag passed is not correct.
Public Const ERROR_SEM_NOT_FOUND              As Long = 187     'The specified system semaphore name was not found.
Public Const ERROR_INVALID_STARTING_CODESEG   As Long = 188     'The operating system cannot run %1.
Public Const ERROR_INVALID_STACKSEG           As Long = 189     'The operating system cannot run %1.
Public Const ERROR_INVALID_MODULETYPE         As Long = 190     'The operating system cannot run %1.
Public Const ERROR_INVALID_EXE_SIGNATURE      As Long = 191     'Cannot run %1 in Win32 mode.
Public Const ERROR_EXE_MARKED_INVALID         As Long = 192     'The operating system cannot run %1.
Public Const ERROR_BAD_EXE_FORMAT             As Long = 193     '%1 is not a valid Win32 application.
Public Const ERROR_ITERATED_DATA_EXCEEDS_64k  As Long = 194     'The operating system cannot run %1.
Public Const ERROR_INVALID_MINALLOCSIZE       As Long = 195     'The operating system cannot run %1.
Public Const ERROR_DYNLINK_FROM_INVALID_RING  As Long = 196     'The operating system cannot run this application program.
Public Const ERROR_IOPL_NOT_ENABLED           As Long = 197     'The operating system is not presently configured to run this application.
Public Const ERROR_INVALID_SEGDPL             As Long = 198     'The operating system cannot run %1.
Public Const ERROR_AUTODATASEG_EXCEEDS_64k    As Long = 199     'The operating system cannot run this application program.
Public Const ERROR_RING2SEG_MUST_BE_MOVABLE   As Long = 200     'The code segment cannot be greater than or equal to 64K.
Public Const ERROR_RELOC_CHAIN_XEEDS_SEGLIM   As Long = 201     'The operating system cannot run %1.
Public Const ERROR_INFLOOP_IN_RELOC_CHAIN     As Long = 202     'The operating system cannot run %1.
Public Const ERROR_ENVVAR_NOT_FOUND           As Long = 203     'The system could not find the environment option that was entered.
Public Const ERROR_NO_SIGNAL_SENT             As Long = 205     'No process in the command subtree has a signal handler.
Public Const ERROR_FILENAME_EXCED_RANGE       As Long = 206     'The filename or extension is too long.
Public Const ERROR_RING2_STACK_IN_USE         As Long = 207     'The ring 2 stack is in use.
Public Const ERROR_META_EXPANSION_TOO_LONG    As Long = 208     'The global filename characters, * or ?, are entered incorrectly or too many global filename characters are specified.
Public Const ERROR_INVALID_SIGNAL_NUMBER      As Long = 209     'The signal being posted is not correct.
Public Const ERROR_THREAD_1_INACTIVE          As Long = 210     'The signal handler cannot be set.
Public Const ERROR_LOCKED                     As Long = 212     'The segment is locked and cannot be reallocated.
Public Const ERROR_TOO_MANY_MODULES           As Long = 214     'Too many dynamic-link modules are attached to this program or dynamic-link module.
Public Const ERROR_NESTING_NOT_ALLOWED        As Long = 215     'Cannot nest calls to LoadModule.
Public Const ERROR_EXE_MACHINE_TYPE_MISMATCH  As Long = 216     'The image file %1 is valid, but is for a machine type other than the current machine.
Public Const ERROR_EXE_CANNOT_MODIFY_SIGNED_BINARY As Long = 217     'The image file %1 is signed, unable to modify.
Public Const ERROR_EXE_CANNOT_MODIFY_STRONG_SIGNED_BINARY As Long = 218     'The image file %1 is strong signed, unable to modify.
Public Const ERROR_BAD_PIPE                   As Long = 230     'The pipe state is invalid.
Public Const ERROR_PIPE_BUSY                  As Long = 231     'All pipe instances are busy.
Public Const ERROR_NO_DATA                    As Long = 232     'The pipe is being closed.
Public Const ERROR_PIPE_NOT_CONNECTED         As Long = 233     'No process is on the other end of the pipe.
Public Const ERROR_MORE_DATA                  As Long = 234     'More data is available.    // dderror
Public Const ERROR_VC_DISCONNECTED            As Long = 240     'The session was canceled.
Public Const ERROR_INVALID_EA_NAME            As Long = 254     'The specified extended attribute name was invalid.
Public Const ERROR_EA_LIST_INCONSISTENT       As Long = 255     'The extended attributes are inconsistent.
Public Const WAIT_TIMEOUT                     As Long = 258     'The wait operation timed out.    // dderror
Public Const ERROR_NO_MORE_ITEMS              As Long = 259     'No more data is available.
Public Const ERROR_CANNOT_COPY                As Long = 266     'The copy functions cannot be used.
Public Const ERROR_DIRECTORY                  As Long = 267     'The directory name is invalid.
Public Const ERROR_EAS_DIDNT_FIT              As Long = 275     'The extended attributes did not fit in the buffer.
Public Const ERROR_EA_FILE_CORRUPT            As Long = 276     'The extended attribute file on the mounted file system is corrupt.
Public Const ERROR_EA_TABLE_FULL              As Long = 277     'The extended attribute table file is full.
Public Const ERROR_INVALID_EA_HANDLE          As Long = 278     'The specified extended attribute handle is invalid.
Public Const ERROR_EAS_NOT_SUPPORTED          As Long = 282     'The mounted file system does not support extended attributes.
Public Const ERROR_NOT_OWNER                  As Long = 288     'Attempt to release mutex not owned by caller.
Public Const ERROR_TOO_MANY_POSTS             As Long = 298     'Too many posts were made to a semaphore.
Public Const ERROR_PARTIAL_COPY               As Long = 299     'Only part of a ReadProcessMemory or WriteProcessMemory request was completed.
Public Const ERROR_OPLOCK_NOT_GRANTED         As Long = 300     'The oplock request is denied.
Public Const ERROR_INVALID_OPLOCK_PROTOCOL    As Long = 301     'An invalid oplock acknowledgment was received by the system.
Public Const ERROR_DISK_TOO_FRAGMENTED        As Long = 302     'The volume is too fragmented to complete this operation.
Public Const ERROR_DELETE_PENDING             As Long = 303     'The file cannot be opened because it is in the process of being deleted.
Public Const ERROR_MR_MID_NOT_FOUND           As Long = 317     'The system cannot find message text for message number 0x%1 in the message file for %2.
Public Const ERROR_SCOPE_NOT_FOUND            As Long = 318     'The scope specified was not found.
Public Const ERROR_INVALID_ADDRESS            As Long = 487     'Attempt to access invalid address.
Public Const ERROR_ARITHMETIC_OVERFLOW        As Long = 534     'Arithmetic result exceeded 32 bits.
Public Const ERROR_PIPE_CONNECTED             As Long = 535     'There is a process on other end of the pipe.
Public Const ERROR_PIPE_LISTENING             As Long = 536     'Waiting for a process to open the other end of the pipe.
Public Const ERROR_EA_ACCESS_DENIED           As Long = 994     'Access to the extended attribute was denied.
Public Const ERROR_OPERATION_ABORTED          As Long = 995     'The I/O operation has been aborted because of either a thread exit or an application request.
Public Const ERROR_IO_INCOMPLETE              As Long = 996     'Overlapped I/O event is not in a signaled state.
Public Const ERROR_IO_PENDING                 As Long = 997     'Overlapped I/O operation is in progress.    // dderror
Public Const ERROR_NOACCESS                   As Long = 998     'Invalid access to memory location.
Public Const ERROR_SWAPERROR                  As Long = 999     'Error performing inpage operation.
Public Const ERROR_STACK_OVERFLOW             As Long = 1001    'Recursion too deep; the stack overflowed.
Public Const ERROR_INVALID_MESSAGE            As Long = 1002    'The window cannot act on the sent message.
Public Const ERROR_CAN_NOT_COMPLETE           As Long = 1003    'Cannot complete this function.
Public Const ERROR_INVALID_FLAGS              As Long = 1004    'Invalid flags.
Public Const ERROR_UNRECOGNIZED_VOLUME        As Long = 1005    'The volume does not contain a recognized file system.   Please make sure that all required file system drivers are loaded and that the volume is not corrupted.
Public Const ERROR_FILE_INVALID               As Long = 1006    'The volume for a file has been externally altered so that the opened file is no longer valid.
Public Const ERROR_FULLSCREEN_MODE            As Long = 1007    'The requested operation cannot be performed in full-screen mode.
Public Const ERROR_NO_TOKEN                   As Long = 1008    'An attempt was made to reference a token that does not exist.
Public Const ERROR_BADDB                      As Long = 1009    'The configuration registry database is corrupt.
Public Const ERROR_BADKEY                     As Long = 1010    'The configuration registry key is invalid.
Public Const ERROR_CANTOPEN                   As Long = 1011    'The configuration registry key could not be opened.
Public Const ERROR_CANTREAD                   As Long = 1012    'The configuration registry key could not be read.
Public Const ERROR_CANTWRITE                  As Long = 1013    'The configuration registry key could not be written.
Public Const ERROR_REGISTRY_RECOVERED         As Long = 1014    'One of the files in the registry database had to be recovered by use of a log or alternate copy. The recovery was successful.
Public Const ERROR_REGISTRY_CORRUPT           As Long = 1015    'The registry is corrupted. The structure of one of the files containing registry data is corrupted, or the system's memory image of the file is corrupted, or the file could not be recovered because the alternate copy or log was absent or corrupted.
Public Const ERROR_REGISTRY_IO_FAILED         As Long = 1016    'An I/O operation initiated by the registry failed unrecoverably. The registry could not read in, or write out, or flush, one of the files that contain the system's image of the registry.
Public Const ERROR_NOT_REGISTRY_FILE          As Long = 1017    'The system has attempted to load or restore a file into the registry, but the specified file is not in a registry file format.
Public Const ERROR_KEY_DELETED                As Long = 1018    'Illegal operation attempted on a registry key that has been marked for deletion.
Public Const ERROR_NO_LOG_SPACE               As Long = 1019    'System could not allocate the required space in a registry log.
Public Const ERROR_KEY_HAS_CHILDREN           As Long = 1020    'Cannot create a symbolic link in a registry key that already has subkeys or values.
Public Const ERROR_CHILD_MUST_BE_VOLATILE     As Long = 1021    'Cannot create a stable subkey under a volatile parent key.
Public Const ERROR_NOTIFY_ENUM_DIR            As Long = 1022    'A notify change request is being completed and the information is not being returned in the caller's buffer. The caller now needs to enumerate the files to find the changes.
Public Const ERROR_DEPENDENT_SERVICES_RUNNING As Long = 1051    'A stop control has been sent to a service that other running services are dependent on.
Public Const ERROR_INVALID_SERVICE_CONTROL    As Long = 1052    'The requested control is not valid for this service.
Public Const ERROR_SERVICE_REQUEST_TIMEOUT    As Long = 1053    'The service did not respond to the start or control request in a timely fashion.
Public Const ERROR_SERVICE_NO_THREAD          As Long = 1054    'A thread could not be created for the service.
Public Const ERROR_SERVICE_DATABASE_LOCKED    As Long = 1055    'The service database is locked.
Public Const ERROR_SERVICE_ALREADY_RUNNING    As Long = 1056    'An instance of the service is already running.
Public Const ERROR_INVALID_SERVICE_ACCOUNT    As Long = 1057    'The account name is invalid or does not exist, or the password is invalid for the account name specified.
Public Const ERROR_SERVICE_DISABLED           As Long = 1058    'The service cannot be started, either because it is disabled or because it has no enabled devices associated with it.
Public Const ERROR_CIRCULAR_DEPENDENCY        As Long = 1059    'Circular service dependency was specified.
Public Const ERROR_SERVICE_DOES_NOT_EXIST     As Long = 1060    'The specified service does not exist as an installed service.
Public Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL As Long = 1061    'The service cannot accept control messages at this time.
Public Const ERROR_SERVICE_NOT_ACTIVE         As Long = 1062    'The service has not been started.
Public Const ERROR_FAILED_SERVICE_CONTROLLER_CONNECT As Long = 1063 'The service process could not connect to the service controller.
Public Const ERROR_EXCEPTION_IN_SERVICE       As Long = 1064    'An exception occurred in the service when handling the control request.
Public Const ERROR_DATABASE_DOES_NOT_EXIST    As Long = 1065    'The database specified does not exist.
Public Const ERROR_SERVICE_SPECIFIC_ERROR     As Long = 1066    'The service has returned a service-specific error code.
Public Const ERROR_PROCESS_ABORTED            As Long = 1067    'The process terminated unexpectedly.
Public Const ERROR_SERVICE_DEPENDENCY_FAIL    As Long = 1068    'The dependency service or group failed to start.
Public Const ERROR_SERVICE_LOGON_FAILED       As Long = 1069    'The service did not start due to a logon failure.
Public Const ERROR_SERVICE_START_HANG         As Long = 1070    'After starting, the service hung in a start-pending state.
Public Const ERROR_INVALID_SERVICE_LOCK       As Long = 1071    'The specified service database lock is invalid.
Public Const ERROR_SERVICE_MARKED_FOR_DELETE  As Long = 1072    'The specified service has been marked for deletion.
Public Const ERROR_SERVICE_EXISTS             As Long = 1073    'The specified service already exists.
Public Const ERROR_ALREADY_RUNNING_LKG        As Long = 1074    'The system is currently running with the last-known-good configuration.
Public Const ERROR_SERVICE_DEPENDENCY_DELETED As Long = 1075    'The dependency service does not exist or has been marked for deletion.
Public Const ERROR_BOOT_ALREADY_ACCEPTED      As Long = 1076    'The current boot has already been accepted for use as the last-known-good control set.
Public Const ERROR_SERVICE_NEVER_STARTED      As Long = 1077    'No attempts to start the service have been made since the last boot.
Public Const ERROR_DUPLICATE_SERVICE_NAME     As Long = 1078    'The name is already in use as either a service name or a service display name.
Public Const ERROR_DIFFERENT_SERVICE_ACCOUNT  As Long = 1079    'The account specified for this service is different from the account specified for other services running in the same process.
Public Const ERROR_CANNOT_DETECT_DRIVER_FAILURE As Long = 1080  'Failure actions can only be set for Win32 services, not for drivers.
Public Const ERROR_CANNOT_DETECT_PROCESS_ABORT As Long = 1081   'This service runs in the same process as the service control manager.    'Therefore, the service control manager cannot take action if this service's process terminates unexpectedly.
Public Const ERROR_NO_RECOVERY_PROGRAM        As Long = 1082    'No recovery program has been configured for this service.
Public Const ERROR_SERVICE_NOT_IN_EXE         As Long = 1083    'The executable program that this service is configured to run in does not implement the service.
Public Const ERROR_NOT_SAFEBOOT_SERVICE       As Long = 1084    'This service cannot be started in Safe Mode
Public Const ERROR_END_OF_MEDIA               As Long = 1100    'The physical end of the tape has been reached.
Public Const ERROR_FILEMARK_DETECTED          As Long = 1101    'A tape access reached a filemark.
Public Const ERROR_BEGINNING_OF_MEDIA         As Long = 1102    'The beginning of the tape or a partition was encountered.
Public Const ERROR_SETMARK_DETECTED           As Long = 1103    'A tape access reached the end of a set of files.
Public Const ERROR_NO_DATA_DETECTED           As Long = 1104    'No more data is on the tape.
Public Const ERROR_PARTITION_FAILURE          As Long = 1105    'Tape could not be partitioned.
Public Const ERROR_INVALID_BLOCK_LENGTH       As Long = 1106    'When accessing a new tape of a multivolume partition, the current block size is incorrect.
Public Const ERROR_DEVICE_NOT_PARTITIONED     As Long = 1107    'Tape partition information could not be found when loading a tape.
Public Const ERROR_UNABLE_TO_LOCK_MEDIA       As Long = 1108    'Unable to lock the media eject mechanism.
Public Const ERROR_UNABLE_TO_UNLOAD_MEDIA     As Long = 1109    'Unable to unload the media.
Public Const ERROR_MEDIA_CHANGED              As Long = 1110    'The media in the drive may have changed.
Public Const ERROR_BUS_RESET                  As Long = 1111    'The I/O bus was reset.
Public Const ERROR_NO_MEDIA_IN_DRIVE          As Long = 1112    'No media in drive.
Public Const ERROR_NO_UNICODE_TRANSLATION     As Long = 1113    'No mapping for the Unicode character exists in the target multi-byte code page.
Public Const ERROR_DLL_INIT_FAILED            As Long = 1114    'A dynamic link library (DLL) initialization routine failed.
Public Const ERROR_SHUTDOWN_IN_PROGRESS       As Long = 1115    'A system shutdown is in progress.
Public Const ERROR_NO_SHUTDOWN_IN_PROGRESS    As Long = 1116    'Unable to abort the system shutdown because no shutdown was in progress.
Public Const ERROR_IO_DEVICE                  As Long = 1117    'The request could not be performed because of an I/O device error.
Public Const ERROR_SERIAL_NO_DEVICE           As Long = 1118    'No serial device was successfully initialized. The serial driver will unload.
Public Const ERROR_IRQ_BUSY                   As Long = 1119    'Unable to open a device that was sharing an interrupt request (IRQ) with other devices. At least one other device that uses that IRQ was already opened.
Public Const ERROR_MORE_WRITES                As Long = 1120    'A serial I/O operation was completed by another write to the serial port.  '(The IOCTL_SERIAL_XOFF_COUNTER reached zero.)
Public Const ERROR_COUNTER_TIMEOUT            As Long = 1121    'A serial I/O operation completed because the timeout period expired.   '(The IOCTL_SERIAL_XOFF_COUNTER did not reach zero.)
Public Const ERROR_FLOPPY_ID_MARK_NOT_FOUND   As Long = 1122    'No ID address mark was found on the floppy disk.
Public Const ERROR_FLOPPY_WRONG_CYLINDER      As Long = 1123    'Mismatch between the floppy disk sector ID field and the floppy disk controller track address.
Public Const ERROR_FLOPPY_UNKNOWN_ERROR       As Long = 1124    'The floppy disk controller reported an error that is not recognized by the floppy disk driver.
Public Const ERROR_FLOPPY_BAD_REGISTERS       As Long = 1125    'The floppy disk controller returned inconsistent results in its registers.
Public Const ERROR_DISK_RECALIBRATE_FAILED    As Long = 1126    'While accessing the hard disk, a recalibrate operation failed, even after retries.
Public Const ERROR_DISK_OPERATION_FAILED      As Long = 1127    'While accessing the hard disk, a disk operation failed even after retries.
Public Const ERROR_DISK_RESET_FAILED          As Long = 1128    'While accessing the hard disk, a disk controller reset was needed, but even that failed.
Public Const ERROR_EOM_OVERFLOW               As Long = 1129    'Physical end of tape encountered.
Public Const ERROR_NOT_ENOUGH_SERVER_MEMORY   As Long = 1130    'Not enough server storage is available to process this command.
Public Const ERROR_POSSIBLE_DEADLOCK          As Long = 1131    'A potential deadlock condition has been detected.
Public Const ERROR_MAPPED_ALIGNMENT           As Long = 1132    'The base address or the file offset specified does not have the proper alignment.
Public Const ERROR_SET_POWER_STATE_VETOED     As Long = 1140    'An attempt to change the system power state was vetoed by another application or driver.
Public Const ERROR_SET_POWER_STATE_FAILED     As Long = 1141    'The system BIOS failed an attempt to change the system power state.
Public Const ERROR_TOO_MANY_LINKS             As Long = 1142    'An attempt was made to create more links on a file than the file system supports.
Public Const ERROR_OLD_WIN_VERSION            As Long = 1150    'The specified program requires a newer version of Windows.
Public Const ERROR_APP_WRONG_OS               As Long = 1151    'The specified program is not a Windows or MS-DOS program.
Public Const ERROR_SINGLE_INSTANCE_APP        As Long = 1152    'Cannot start more than one instance of the specified program.
Public Const ERROR_RMODE_APP                  As Long = 1153    'The specified program was written for an earlier version of Windows.
Public Const ERROR_INVALID_DLL                As Long = 1154    'One of the library files needed to run this application is damaged.
Public Const ERROR_NO_ASSOCIATION             As Long = 1155    'No application is associated with the specified file for this operation.
Public Const ERROR_DDE_FAIL                   As Long = 1156    'An error occurred in sending the command to the application.
Public Const ERROR_DLL_NOT_FOUND              As Long = 1157    'One of the library files needed to run this application cannot be found.
Public Const ERROR_NO_MORE_USER_HANDLES       As Long = 1158    'The current process has used all of its system allowance of handles for Window Manager objects.
Public Const ERROR_MESSAGE_SYNC_ONLY          As Long = 1159    'The message can be used only with synchronous operations.
Public Const ERROR_SOURCE_ELEMENT_EMPTY       As Long = 1160    'The indicated source element has no media.
Public Const ERROR_DESTINATION_ELEMENT_FULL   As Long = 1161    'The indicated destination element already contains media.
Public Const ERROR_ILLEGAL_ELEMENT_ADDRESS    As Long = 1162    'The indicated element does not exist.
Public Const ERROR_MAGAZINE_NOT_PRESENT       As Long = 1163    'The indicated element is part of a magazine that is not present.
Public Const ERROR_DEVICE_REINITIALIZATION_NEEDED As Long = 1164 'The indicated device requires reinitialization due to hardware errors.    // dderror
Public Const ERROR_DEVICE_REQUIRES_CLEANING   As Long = 1165    'The device has indicated that cleaning is required before further operations are attempted.
Public Const ERROR_DEVICE_DOOR_OPEN           As Long = 1166    'The device has indicated that its door is open.
Public Const ERROR_DEVICE_NOT_CONNECTED       As Long = 1167    'The device is not connected.
Public Const ERROR_NOT_FOUND                  As Long = 1168    'Element not found.
Public Const ERROR_NO_MATCH                   As Long = 1169    'There was no match for the specified key in the index.
Public Const ERROR_SET_NOT_FOUND              As Long = 1170    'The property set specified does not exist on the object.
Public Const ERROR_POINT_NOT_FOUND            As Long = 1171    'The point passed to GetMouseMovePoints is not in the buffer.
Public Const ERROR_NO_TRACKING_SERVICE        As Long = 1172    'The tracking (workstation) service is not running.
Public Const ERROR_NO_VOLUME_ID               As Long = 1173    'The Volume ID could not be found.
Public Const ERROR_UNABLE_TO_REMOVE_REPLACED  As Long = 1175    'Unable to remove the file to be replaced.
Public Const ERROR_UNABLE_TO_MOVE_REPLACEMENT As Long = 1176    'Unable to move the replacement file to the file to be replaced. The file to be replaced has retained its original name.
Public Const ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 As Long = 1177    'Unable to move the replacement file to the file to be replaced. The file to be replaced has been renamed using the backup name.
Public Const ERROR_JOURNAL_DELETE_IN_PROGRESS As Long = 1178    'The volume change journal is being deleted.
Public Const ERROR_JOURNAL_NOT_ACTIVE         As Long = 1179    'The volume change journal is not active.
Public Const ERROR_POTENTIAL_FILE_FOUND       As Long = 1180    'A file was found, but it may not be the correct file.
Public Const ERROR_JOURNAL_ENTRY_DELETED      As Long = 1181    'The journal entry has been deleted from the journal.
Public Const ERROR_BAD_DEVICE                 As Long = 1200    'The specified device name is invalid.
Public Const ERROR_CONNECTION_UNAVAIL         As Long = 1201    'The device is not currently connected but it is a remembered connection.
Public Const ERROR_DEVICE_ALREADY_REMEMBERED  As Long = 1202    'The local device name has a remembered connection to another network resource.
Public Const ERROR_NO_NET_OR_BAD_PATH         As Long = 1203    'No network provider accepted the given network path.
Public Const ERROR_BAD_PROVIDER               As Long = 1204    'The specified network provider name is invalid.
Public Const ERROR_CANNOT_OPEN_PROFILE        As Long = 1205    'Unable to open the network connection profile.
Public Const ERROR_BAD_PROFILE                As Long = 1206    'The network connection profile is corrupted.
Public Const ERROR_NOT_CONTAINER              As Long = 1207    'Cannot enumerate a noncontainer.
Public Const ERROR_EXTENDED_ERROR             As Long = 1208    'An extended error has occurred.
Public Const ERROR_INVALID_GROUPNAME          As Long = 1209    'The format of the specified group name is invalid.
Public Const ERROR_INVALID_COMPUTERNAME       As Long = 1210    'The format of the specified computer name is invalid.
Public Const ERROR_INVALID_EVENTNAME          As Long = 1211    'The format of the specified event name is invalid.
Public Const ERROR_INVALID_DOMAINNAME         As Long = 1212    'The format of the specified domain name is invalid.
Public Const ERROR_INVALID_SERVICENAME        As Long = 1213    'The format of the specified service name is invalid.
Public Const ERROR_INVALID_NETNAME            As Long = 1214    'The format of the specified network name is invalid.
Public Const ERROR_INVALID_SHARENAME          As Long = 1215    'The format of the specified share name is invalid.
Public Const ERROR_INVALID_PASSWORDNAME       As Long = 1216    'The format of the specified password is invalid.
Public Const ERROR_INVALID_MESSAGENAME        As Long = 1217    'The format of the specified message name is invalid.
Public Const ERROR_INVALID_MESSAGEDEST        As Long = 1218    'The format of the specified message destination is invalid.
Public Const ERROR_SESSION_CREDENTIAL_CONFLICT As Long = 1219    'Multiple connections to a server or shared resource by the same user, using more than one user name, are not allowed. Disconnect all previous connections to the server or shared resource and try again.
Public Const ERROR_REMOTE_SESSION_LIMIT_EXCEEDED As Long = 1220    'An attempt was made to establish a session to a network server, but there are already too many sessions established to that server.
Public Const ERROR_DUP_DOMAINNAME             As Long = 1221    'The workgroup or domain name is already in use by another computer on the network.
Public Const ERROR_NO_NETWORK                 As Long = 1222    'The network is not present or not started.
Public Const ERROR_CANCELLED                  As Long = 1223    'The operation was canceled by the user.
Public Const ERROR_USER_MAPPED_FILE           As Long = 1224    'The requested operation cannot be performed on a file with a user-mapped section open.
Public Const ERROR_CONNECTION_REFUSED         As Long = 1225    'The remote system refused the network connection.
Public Const ERROR_GRACEFUL_DISCONNECT        As Long = 1226    'The network connection was gracefully closed.
Public Const ERROR_ADDRESS_ALREADY_ASSOCIATED As Long = 1227    'The network transport endpoint already has an address associated with it.
Public Const ERROR_ADDRESS_NOT_ASSOCIATED     As Long = 1228    'An address has not yet been associated with the network endpoint.
Public Const ERROR_CONNECTION_INVALID         As Long = 1229    'An operation was attempted on a nonexistent network connection.
Public Const ERROR_CONNECTION_ACTIVE          As Long = 1230    'An invalid operation was attempted on an active network connection.
Public Const ERROR_NETWORK_UNREACHABLE        As Long = 1231    'The network location cannot be reached. For information about network troubleshooting, see Windows Help.
Public Const ERROR_HOST_UNREACHABLE           As Long = 1232    'The network location cannot be reached. For information about network troubleshooting, see Windows Help.
Public Const ERROR_PROTOCOL_UNREACHABLE       As Long = 1233    'The network location cannot be reached. For information about network troubleshooting, see Windows Help.
Public Const ERROR_PORT_UNREACHABLE           As Long = 1234    'No service is operating at the destination network endpoint on the remote system.
Public Const ERROR_REQUEST_ABORTED            As Long = 1235    'The request was aborted.
Public Const ERROR_CONNECTION_ABORTED         As Long = 1236    'The network connection was aborted by the local system.
Public Const ERROR_RETRY                      As Long = 1237    'The operation could not be completed. A retry should be performed.
Public Const ERROR_CONNECTION_COUNT_LIMIT     As Long = 1238    'A connection to the server could not be made because the limit on the number of concurrent connections for this account has been reached.
Public Const ERROR_LOGIN_TIME_RESTRICTION     As Long = 1239    'Attempting to log in during an unauthorized time of day for this account.
Public Const ERROR_LOGIN_WKSTA_RESTRICTION    As Long = 1240    'The account is not authorized to log in from this station.
Public Const ERROR_INCORRECT_ADDRESS          As Long = 1241    'The network address could not be used for the operation requested.
Public Const ERROR_ALREADY_REGISTERED         As Long = 1242    'The service is already registered.
Public Const ERROR_SERVICE_NOT_FOUND          As Long = 1243    'The specified service does not exist.
Public Const ERROR_NOT_AUTHENTICATED          As Long = 1244    'The operation being requested was not performed because the user has not been authenticated.
Public Const ERROR_NOT_LOGGED_ON              As Long = 1245    'The operation being requested was not performed because the user has not logged on to the network.   'The specified service does not exist.
Public Const ERROR_CONTINUE                   As Long = 1246    'Continue with work in progress.    // dderror
Public Const ERROR_ALREADY_INITIALIZED        As Long = 1247    'An attempt was made to perform an initialization operation when initialization has already been completed.
Public Const ERROR_NO_MORE_DEVICES            As Long = 1248    'No more local devices.    // dderror
Public Const ERROR_NO_SUCH_SITE               As Long = 1249    'The specified site does not exist.
Public Const ERROR_DOMAIN_CONTROLLER_EXISTS   As Long = 1250    'A domain controller with the specified name already exists.
Public Const ERROR_ONLY_IF_CONNECTED          As Long = 1251    'This operation is supported only when you are connected to the server.
Public Const ERROR_OVERRIDE_NOCHANGES         As Long = 1252    'The group policy framework should call the extension even if there are no changes.
Public Const ERROR_BAD_USER_PROFILE           As Long = 1253    'The specified user does not have a valid profile.
Public Const ERROR_NOT_SUPPORTED_ON_SBS       As Long = 1254    'This operation is not supported on a computer running Windows Server 2003 for Small Business Server
Public Const ERROR_SERVER_SHUTDOWN_IN_PROGRESS As Long = 1255    'The server machine is shutting down.
Public Const ERROR_HOST_DOWN                  As Long = 1256    'The remote system is not available. For information about network troubleshooting, see Windows Help.
Public Const ERROR_NON_ACCOUNT_SID            As Long = 1257    'The security identifier provided is not from an account domain.
Public Const ERROR_NON_DOMAIN_SID             As Long = 1258    'The security identifier provided does not have a domain component.
Public Const ERROR_APPHELP_BLOCK              As Long = 1259    'AppHelp dialog canceled thus preventing the application from starting.
Public Const ERROR_ACCESS_DISABLED_BY_POLICY  As Long = 1260    'Windows cannot open this program because it has been prevented by a software restriction policy. For more information, open Event Viewer or contact your system administrator.
Public Const ERROR_REG_NAT_CONSUMPTION        As Long = 1261    'A program attempt to use an invalid register value.  Normally caused by an uninitialized register. This error is Itanium specific.
Public Const ERROR_CSCSHARE_OFFLINE           As Long = 1262    'The share is currently offline or does not exist.
Public Const ERROR_PKINIT_FAILURE             As Long = 1263    'The kerberos protocol encountered an error while validating the KDC certificate during smartcard logon.  There is more information in the system event log.
Public Const ERROR_SMARTCARD_SUBSYSTEM_FAILURE As Long = 1264    'The kerberos protocol encountered an error while attempting to utilize the smartcard subsystem.
Public Const ERROR_DOWNGRADE_DETECTED         As Long = 1265    'The system detected a possible attempt to compromise security. Please ensure that you can contact the server that authenticated you.
Public Const ERROR_MACHINE_LOCKED             As Long = 1271    'The machine is locked and can not be shut down without the force option.
Public Const ERROR_CALLBACK_SUPPLIED_INVALID_DATA As Long = 1273    'An application-defined callback gave invalid data when called.
Public Const ERROR_SYNC_FOREGROUND_REFRESH_REQUIRED As Long = 1274    'The group policy framework should call the extension in the synchronous foreground policy refresh.
Public Const ERROR_DRIVER_BLOCKED             As Long = 1275    'This driver has been blocked from loading
Public Const ERROR_INVALID_IMPORT_OF_NON_DLL  As Long = 1276    'A dynamic link library (DLL) referenced a module that was neither a DLL nor the process's executable image.
Public Const ERROR_ACCESS_DISABLED_WEBBLADE   As Long = 1277    'Windows cannot open this program since it has been disabled.
Public Const ERROR_ACCESS_DISABLED_WEBBLADE_TAMPER As Long = 1278    'Windows cannot open this program because the license enforcement system has been tampered with or become corrupted.
Public Const ERROR_RECOVERY_FAILURE           As Long = 1279    'A transaction recover failed.
Public Const ERROR_ALREADY_FIBER              As Long = 1280    'The current thread has already been converted to a fiber.
Public Const ERROR_ALREADY_THREAD             As Long = 1281    'The current thread has already been converted from a fiber.
Public Const ERROR_STACK_BUFFER_OVERRUN       As Long = 1282    'The system detected an overrun of a stack-based buffer in this application.  This overrun could potentially allow a malicious user to gain control of this application.
Public Const ERROR_PARAMETER_QUOTA_EXCEEDED   As Long = 1283    'Data present in one of the parameters is more than the function can operate on.
Public Const ERROR_DEBUGGER_INACTIVE          As Long = 1284    'An attempt to do an operation on a debug object failed because the object is in the process of being deleted.
Public Const ERROR_DELAY_LOAD_FAILED          As Long = 1285    'An attempt to delay-load a .dll or get a function address in a delay-loaded .dll failed.
Public Const ERROR_VDM_DISALLOWED             As Long = 1286    '%1 is a 16-bit application. You do not have permissions to execute 16-bit applications. Check your permissions with your system administrator.
Public Const ERROR_UNIDENTIFIED_ERROR         As Long = 1287    'Insufficient information exists to identify the cause of failure.
'
'
'///////////////////////////
'//
'// Add new status codes before this point unless there is a component specific section below.
'//
'///////////////////////////
'
'
'///////////////////////////
'//                       //
'// Security Status Codes //
'//                       //
'///////////////////////////
'
Public Const ERROR_NOT_ALL_ASSIGNED           As Long = 1300    'Not all privileges referenced are assigned to the caller.
Public Const ERROR_SOME_NOT_MAPPED            As Long = 1301    'Some mapping between account names and security IDs was not done.
Public Const ERROR_NO_QUOTAS_FOR_ACCOUNT      As Long = 1302    'No system quota limits are specifically set for this account.
Public Const ERROR_LOCAL_USER_SESSION_KEY     As Long = 1303    'No encryption key is available. A well-known encryption key was returned.
Public Const ERROR_NULL_LM_PASSWORD           As Long = 1304    'The password is too complex to be converted to a LAN Manager password. The LAN Manager password returned is a NULL string.
Public Const ERROR_UNKNOWN_REVISION           As Long = 1305    'The revision level is unknown.
Public Const ERROR_REVISION_MISMATCH          As Long = 1306    'Indicates two revision levels are incompatible.
Public Const ERROR_INVALID_OWNER              As Long = 1307    'This security ID may not be assigned as the owner of this object.
Public Const ERROR_INVALID_PRIMARY_GROUP      As Long = 1308    'This security ID may not be assigned as the primary group of an object.
Public Const ERROR_NO_IMPERSONATION_TOKEN     As Long = 1309    'An attempt has been made to operate on an impersonation token by a thread that is not currently impersonating a client.
Public Const ERROR_CANT_DISABLE_MANDATORY     As Long = 1310    'The group may not be disabled.
Public Const ERROR_NO_LOGON_SERVERS           As Long = 1311    'There are currently no logon servers available to service the logon request.
Public Const ERROR_NO_SUCH_LOGON_SESSION      As Long = 1312    'A specified logon session does not exist. It may already have been terminated.
Public Const ERROR_NO_SUCH_PRIVILEGE          As Long = 1313    'A specified privilege does not exist.
Public Const ERROR_PRIVILEGE_NOT_HELD         As Long = 1314    'A required privilege is not held by the client.
Public Const ERROR_INVALID_ACCOUNT_NAME       As Long = 1315    'The name provided is not a properly formed account name.
Public Const ERROR_USER_EXISTS                As Long = 1316    'The specified user already exists.
Public Const ERROR_NO_SUCH_USER               As Long = 1317    'The specified user does not exist.
Public Const ERROR_GROUP_EXISTS               As Long = 1318    'The specified group already exists.
Public Const ERROR_NO_SUCH_GROUP              As Long = 1319    'The specified group does not exist.
Public Const ERROR_MEMBER_IN_GROUP            As Long = 1320    'Either the specified user account is already a member of the specified group, or the specified group cannot be deleted because it contains a member.
Public Const ERROR_MEMBER_NOT_IN_GROUP        As Long = 1321    'The specified user account is not a member of the specified group account.
Public Const ERROR_LAST_ADMIN                 As Long = 1322    'The last remaining administration account cannot be disabled or deleted.
Public Const ERROR_WRONG_PASSWORD             As Long = 1323    'Unable to update the password. The value provided as the current password is incorrect.
Public Const ERROR_ILL_FORMED_PASSWORD        As Long = 1324    'Unable to update the password. The value provided for the new password contains values that are not allowed in passwords.
Public Const ERROR_PASSWORD_RESTRICTION       As Long = 1325    'Unable to update the password. The value provided for the new password does not meet the length, complexity, or history requirement of the domain.
Public Const ERROR_LOGON_FAILURE              As Long = 1326    'Logon failure: unknown user name or bad password.
Public Const ERROR_ACCOUNT_RESTRICTION        As Long = 1327    'Logon failure: user account restriction.  Possible reasons are blank passwords not allowed, logon hour restrictions, or a policy restriction has been enforced.
Public Const ERROR_INVALID_LOGON_HOURS        As Long = 1328    'Logon failure: account logon time restriction violation.
Public Const ERROR_INVALID_WORKSTATION        As Long = 1329    'Logon failure: user not allowed to log on to this computer.
Public Const ERROR_PASSWORD_EXPIRED           As Long = 1330    'Logon failure: the specified account password has expired.
Public Const ERROR_ACCOUNT_DISABLED           As Long = 1331    'Logon failure: account currently disabled.
Public Const ERROR_NONE_MAPPED                As Long = 1332    'No mapping between account names and security IDs was done.
Public Const ERROR_TOO_MANY_LUIDS_REQUESTED   As Long = 1333    'Too many local user identifiers (LUIDs) were requested at one time.
Public Const ERROR_LUIDS_EXHAUSTED            As Long = 1334    'No more local user identifiers (LUIDs) are available.
Public Const ERROR_INVALID_SUB_AUTHORITY      As Long = 1335    'The subauthority part of a security ID is invalid for this particular use.
Public Const ERROR_INVALID_ACL                As Long = 1336    'The access control list (ACL) structure is invalid.
Public Const ERROR_INVALID_SID                As Long = 1337    'The security ID structure is invalid.
Public Const ERROR_INVALID_SECURITY_DESCR     As Long = 1338    'The security descriptor structure is invalid.
Public Const ERROR_BAD_INHERITANCE_ACL        As Long = 1340    'The inherited access control list (ACL) or access control entry (ACE) could not be built.
Public Const ERROR_SERVER_DISABLED            As Long = 1341    'The server is currently disabled.
Public Const ERROR_SERVER_NOT_DISABLED        As Long = 1342    'The server is currently enabled.
Public Const ERROR_INVALID_ID_AUTHORITY       As Long = 1343    'The value provided was an invalid value for an identifier authority.
Public Const ERROR_ALLOTTED_SPACE_EXCEEDED    As Long = 1344    'The value provided was an invalid value for an identifier authority.
Public Const ERROR_INVALID_GROUP_ATTRIBUTES   As Long = 1345    'The specified attributes are invalid, or incompatible with the attributes for the group as a whole.
Public Const ERROR_BAD_IMPERSONATION_LEVEL    As Long = 1346    'Either a required impersonation level was not provided, or the provided impersonation level is invalid.
Public Const ERROR_CANT_OPEN_ANONYMOUS        As Long = 1347    'Cannot open an anonymous level security token.
Public Const ERROR_BAD_VALIDATION_CLASS       As Long = 1348    'The validation information class requested was invalid.
Public Const ERROR_BAD_TOKEN_TYPE             As Long = 1349    'The type of the token is inappropriate for its attempted use.
Public Const ERROR_NO_SECURITY_ON_OBJECT      As Long = 1350    'Unable to perform a security operation on an object that has no associated security.
Public Const ERROR_CANT_ACCESS_DOMAIN_INFO    As Long = 1351    'Configuration information could not be read from the domain controller, either because the machine is unavailable, or access has been denied.
Public Const ERROR_INVALID_SERVER_STATE       As Long = 1352    'The security account manager (SAM) or local security authority (LSA) server was in the wrong state to perform the security operation.
Public Const ERROR_INVALID_DOMAIN_STATE       As Long = 1353    'The domain was in the wrong state to perform the security operation.
Public Const ERROR_INVALID_DOMAIN_ROLE        As Long = 1354    'This operation is only allowed for the Primary Domain Controller of the domain.
Public Const ERROR_NO_SUCH_DOMAIN             As Long = 1355    'The specified domain either does not exist or could not be contacted.
'Public Const ERROR_DOMAIN_EXISTS              1356L'The specified domain already exists.
Public Const ERROR_DOMAIN_LIMIT_EXCEEDED      As Long = 1357 'An attempt was made to exceed the limit on the number of domains per server.
Public Const ERROR_INTERNAL_DB_CORRUPTION     As Long = 1358 'Unable to complete the requested operation because of either a catastrophic media failure or a data structure corruption on the disk.
Public Const ERROR_INTERNAL_ERROR             As Long = 1359 'An internal error occurred.
Public Const ERROR_GENERIC_NOT_MAPPED         As Long = 1360 'Generic access types were contained in an access mask which should already be mapped to nongeneric types.
Public Const ERROR_BAD_DESCRIPTOR_FORMAT      As Long = 1361 'A security descriptor is not in the right format (absolute or self-relative).
Public Const ERROR_NOT_LOGON_PROCESS          As Long = 1362 'The requested action is restricted for use by logon processes only. The calling process has not registered as a logon process.
Public Const ERROR_LOGON_SESSION_EXISTS       As Long = 1363 'Cannot start a new logon session with an ID that is already in use.
Public Const ERROR_NO_SUCH_PACKAGE            As Long = 1364 'A specified authentication package is unknown.
Public Const ERROR_BAD_LOGON_SESSION_STATE    As Long = 1365 'The logon session is not in a state that is consistent with the requested operation.
Public Const ERROR_LOGON_SESSION_COLLISION    As Long = 1366 'The logon session ID is already in use.
Public Const ERROR_INVALID_LOGON_TYPE         As Long = 1367 'A logon request contained an invalid logon type value.
Public Const ERROR_CANNOT_IMPERSONATE         As Long = 1368 'Unable to impersonate using a named pipe until data has been read from that pipe.
Public Const ERROR_RXACT_INVALID_STATE        As Long = 1369 'The transaction state of a registry subtree is incompatible with the requested operation.
Public Const ERROR_RXACT_COMMIT_FAILURE       As Long = 1370 'An internal security database corruption has been encountered.
Public Const ERROR_SPECIAL_ACCOUNT            As Long = 1371 'Cannot perform this operation on built-in accounts.
Public Const ERROR_SPECIAL_GROUP              As Long = 1372 'Cannot perform this operation on this built-in special group.
Public Const ERROR_SPECIAL_USER               As Long = 1373 'Cannot perform this operation on this built-in special user.
Public Const ERROR_MEMBERS_PRIMARY_GROUP      As Long = 1374 'The user cannot be removed from a group because the group is currently the user's primary group.
Public Const ERROR_TOKEN_ALREADY_IN_USE       As Long = 1375 'The token is already in use as a primary token.
Public Const ERROR_NO_SUCH_ALIAS              As Long = 1376 'The specified local group does not exist.
Public Const ERROR_MEMBER_NOT_IN_ALIAS        As Long = 1377 'The specified account name is not a member of the local group.
Public Const ERROR_MEMBER_IN_ALIAS            As Long = 1378 'The specified account name is already a member of the local group.
Public Const ERROR_ALIAS_EXISTS               As Long = 1379 'The specified local group already exists.
Public Const ERROR_LOGON_NOT_GRANTED          As Long = 1380 'Logon failure: the user has not been granted the requested logon type at this computer.
Public Const ERROR_TOO_MANY_SECRETS           As Long = 1381 'The maximum number of secrets that may be stored in a single system has been exceeded.
Public Const ERROR_SECRET_TOO_LONG            As Long = 1382 'The length of a secret exceeds the maximum length allowed.
Public Const ERROR_INTERNAL_DB_ERROR          As Long = 1383 'The local security authority database contains an internal inconsistency.
Public Const ERROR_TOO_MANY_CONTEXT_IDS       As Long = 1384 'During a logon attempt, the user's security context accumulated too many security IDs.
Public Const ERROR_LOGON_TYPE_NOT_GRANTED     As Long = 1385 'Logon failure: the user has not been granted the requested logon type at this computer.
Public Const ERROR_NT_CROSS_ENCRYPTION_REQUIRED As Long = 1386 'A cross-encrypted password is necessary to change a user password.
Public Const ERROR_NO_SUCH_MEMBER             As Long = 1387 'A member could not be added to or removed from the local group because the member does not exist.
Public Const ERROR_INVALID_MEMBER             As Long = 1388 'A new member could not be added to a local group because the member has the wrong account type.
Public Const ERROR_TOO_MANY_SIDS              As Long = 1389 'Too many security IDs have been specified.
Public Const ERROR_LM_CROSS_ENCRYPTION_REQUIRED As Long = 1390 'A cross-encrypted password is necessary to change this user password.
Public Const ERROR_NO_INHERITANCE             As Long = 1391 'Indicates an ACL contains no inheritable components.
Public Const ERROR_FILE_CORRUPT               As Long = 1392 'The file or directory is corrupted and unreadable.
Public Const ERROR_DISK_CORRUPT               As Long = 1393 'The disk structure is corrupted and unreadable.
Public Const ERROR_NO_USER_SESSION_KEY        As Long = 1394 'There is no user session key for the specified logon session.
Public Const ERROR_LICENSE_QUOTA_EXCEEDED     As Long = 1395 'The service being accessed is licensed for a particular number of connections.    'No more connections can be made to the service at this time because there are already as many connections as the service can accept.
Public Const ERROR_WRONG_TARGET_NAME          As Long = 1396 'Logon Failure: The target account name is incorrect.
Public Const ERROR_MUTUAL_AUTH_FAILED         As Long = 1397 'Mutual Authentication failed. The server's password is out of date at the domain controller.
Public Const ERROR_TIME_SKEW                  As Long = 1398 'There is a time and/or date difference between the client and server.
Public Const ERROR_CURRENT_DOMAIN_NOT_ALLOWED As Long = 1399 'This operation can not be performed on the current domain.

'
'// End of security error codes
'
'
'
'///////////////////////////
'//                       //
'// WinUser Error Codes   //
'//                       //
'///////////////////////////
'
Public Const ERROR_INVALID_WINDOW_HANDLE      As Long = 1400 'Invalid window handle.
Public Const ERROR_INVALID_MENU_HANDLE        As Long = 1401 'Invalid menu handle.
Public Const ERROR_INVALID_CURSOR_HANDLE      As Long = 1402 'Invalid cursor handle.
Public Const ERROR_INVALID_ACCEL_HANDLE       As Long = 1403 'Invalid accelerator table handle.
Public Const ERROR_INVALID_HOOK_HANDLE        As Long = 1404 'Invalid hook handle.
Public Const ERROR_INVALID_DWP_HANDLE         As Long = 1405 'Invalid handle to a multiple-window position structure.
Public Const ERROR_TLW_WITH_WSCHILD           As Long = 1406 'Cannot create a top-level child window.
Public Const ERROR_CANNOT_FIND_WND_CLASS      As Long = 1407 'Cannot find window class.
Public Const ERROR_WINDOW_OF_OTHER_THREAD     As Long = 1408 'Invalid window; it belongs to other thread.
Public Const ERROR_HOTKEY_ALREADY_REGISTERED  As Long = 1409 'Hot key is already registered.
Public Const ERROR_CLASS_ALREADY_EXISTS       As Long = 1410 'Class already exists.
Public Const ERROR_CLASS_DOES_NOT_EXIST       As Long = 1411 'Class does not exist.
Public Const ERROR_CLASS_HAS_WINDOWS          As Long = 1412 'Class still has open windows.
Public Const ERROR_INVALID_INDEX              As Long = 1413 'Invalid index.
Public Const ERROR_INVALID_ICON_HANDLE        As Long = 1414 'Invalid icon handle.
Public Const ERROR_PRIVATE_DIALOG_INDEX       As Long = 1415 'Using private DIALOG window words.
Public Const ERROR_LISTBOX_ID_NOT_FOUND       As Long = 1416 'The list box identifier was not found.
Public Const ERROR_NO_WILDCARD_CHARACTERS     As Long = 1417 'No wildcards were found.
Public Const ERROR_CLIPBOARD_NOT_OPEN         As Long = 1418 'Thread does not have a clipboard open.
Public Const ERROR_HOTKEY_NOT_REGISTERED      As Long = 1419 'Hot key is not registered.
Public Const ERROR_WINDOW_NOT_DIALOG          As Long = 1420 'The window is not a valid dialog window.
Public Const ERROR_CONTROL_ID_NOT_FOUND       As Long = 1421 'Control ID not found.
Public Const ERROR_INVALID_COMBOBOX_MESSAGE   As Long = 1422 'Invalid message for a combo box because it does not have an edit control.
Public Const ERROR_WINDOW_NOT_COMBOBOX        As Long = 1423 'The window is not a combo box.
Public Const ERROR_INVALID_EDIT_HEIGHT        As Long = 1424 'Height must be less than 256.
Public Const ERROR_DC_NOT_FOUND               As Long = 1425 'Invalid device context (DC) handle.
Public Const ERROR_INVALID_HOOK_FILTER        As Long = 1426 'Invalid hook procedure type.
Public Const ERROR_INVALID_FILTER_PROC        As Long = 1427 'Invalid hook procedure.
Public Const ERROR_HOOK_NEEDS_HMOD            As Long = 1428 'Cannot set nonlocal hook without a module handle.
Public Const ERROR_GLOBAL_ONLY_HOOK           As Long = 1429 'This hook procedure can only be set globally.
Public Const ERROR_JOURNAL_HOOK_SET           As Long = 1430 'The journal hook procedure is already installed.
Public Const ERROR_HOOK_NOT_INSTALLED         As Long = 1431 'The hook procedure is not installed.
Public Const ERROR_INVALID_LB_MESSAGE         As Long = 1432 'Invalid message for single-selection list box.
Public Const ERROR_SETCOUNT_ON_BAD_LB         As Long = 1433 'LB_SETCOUNT sent to non-lazy list box.
Public Const ERROR_LB_WITHOUT_TABSTOPS        As Long = 1434 'This list box does not support tab stops.
Public Const ERROR_DESTROY_OBJECT_OF_OTHER_THREAD As Long = 1435 'Cannot destroy object created by another thread.
Public Const ERROR_CHILD_WINDOW_MENU          As Long = 1436 'Child windows cannot have menus.
Public Const ERROR_NO_SYSTEM_MENU             As Long = 1437 'The window does not have a system menu.
Public Const ERROR_INVALID_MSGBOX_STYLE       As Long = 1438 'Invalid message box style.
Public Const ERROR_INVALID_SPI_VALUE          As Long = 1439 'Invalid system-wide (SPI_*) parameter.
Public Const ERROR_SCREEN_ALREADY_LOCKED      As Long = 1440 'Screen already locked.
Public Const ERROR_HWNDS_HAVE_DIFF_PARENT     As Long = 1441 'All handles to windows in a multiple-window position structure must have the same parent.
Public Const ERROR_NOT_CHILD_WINDOW           As Long = 1442 'The window is not a child window.
Public Const ERROR_INVALID_GW_COMMAND         As Long = 1443 'Invalid GW_* command.
Public Const ERROR_INVALID_THREAD_ID          As Long = 1444 'Invalid thread identifier.
Public Const ERROR_NON_MDICHILD_WINDOW        As Long = 1445 'Cannot process a message from a window that is not a multiple document interface (MDI) window.
Public Const ERROR_POPUP_ALREADY_ACTIVE       As Long = 1446 'Popup menu already active.
Public Const ERROR_NO_SCROLLBARS              As Long = 1447 'The window does not have scroll bars.
Public Const ERROR_INVALID_SCROLLBAR_RANGE    As Long = 1448 'Scroll bar range cannot be greater than MAXLONG.
Public Const ERROR_INVALID_SHOWWIN_COMMAND    As Long = 1449 'Cannot show or remove the window in the way specified.
Public Const ERROR_NO_SYSTEM_RESOURCES        As Long = 1450 'Insufficient system resources exist to complete the requested service.
Public Const ERROR_NONPAGED_SYSTEM_RESOURCES  As Long = 1451 'Insufficient system resources exist to complete the requested service.
Public Const ERROR_PAGED_SYSTEM_RESOURCES     As Long = 1452 'Insufficient system resources exist to complete the requested service.
Public Const ERROR_WORKING_SET_QUOTA          As Long = 1453 'Insufficient quota to complete the requested service.
Public Const ERROR_PAGEFILE_QUOTA             As Long = 1454 'Insufficient quota to complete the requested service.
Public Const ERROR_COMMITMENT_LIMIT           As Long = 1455 'The paging file is too small for this operation to complete.
Public Const ERROR_MENU_ITEM_NOT_FOUND        As Long = 1456 'A menu item was not found.
Public Const ERROR_INVALID_KEYBOARD_HANDLE    As Long = 1457 'Invalid keyboard layout handle.
Public Const ERROR_HOOK_TYPE_NOT_ALLOWED      As Long = 1458 'Hook type not allowed.
Public Const ERROR_REQUIRES_INTERACTIVE_WINDOWSTATION As Long = 1459 'This operation requires an interactive window station.
Public Const ERROR_TIMEOUT                    As Long = 1460 'This operation returned because the timeout period expired.
Public Const ERROR_INVALID_MONITOR_HANDLE     As Long = 1461 'Invalid monitor handle.
Public Const ERROR_INCORRECT_SIZE             As Long = 1462 'Incorrect size argument.

'
'// End of WinUser error codes
'
'
'
'///////////////////////////
'//                       //
'// Eventlog Status Codes //
'//                       //
'///////////////////////////
'
'Public Const ERROR_EVENTLOG_FILE_CORRUPT      1500L'The event log file is corrupted.
'Public Const ERROR_EVENTLOG_CANT_START        1501L'No event log file could be opened, so the event logging service did not start.
'Public Const ERROR_LOG_FILE_FULL              1502L'The event log file is full.
'Public Const ERROR_EVENTLOG_FILE_CHANGED      1503L'The event log file has changed between read operations.
'
'// End of eventlog error codes
'
'
'
'///////////////////////////
'//                       //
'// MSI Error Codes       //
'//                       //
'///////////////////////////
'
Public Const ERROR_INSTALL_SERVICE_FAILURE    As Long = 1601 'The Windows Installer Service could not be accessed. This can occur if you are running Windows in safe mode, or if the Windows Installer is not correctly installed. Contact your support personnel for assistance.
Public Const ERROR_INSTALL_USEREXIT           As Long = 1602 'User cancelled installation.
Public Const ERROR_INSTALL_FAILURE            As Long = 1603 'Fatal error during installation.
Public Const ERROR_INSTALL_SUSPEND            As Long = 1604 'Installation suspended, incomplete.
Public Const ERROR_UNKNOWN_PRODUCT            As Long = 1605 'This action is only valid for products that are currently installed.
Public Const ERROR_UNKNOWN_FEATURE            As Long = 1606 'Feature ID not registered.
Public Const ERROR_UNKNOWN_COMPONENT          As Long = 1607 'Component ID not registered.
Public Const ERROR_UNKNOWN_PROPERTY           As Long = 1608 'Unknown property.
Public Const ERROR_INVALID_HANDLE_STATE       As Long = 1609 'Handle is in an invalid state.
Public Const ERROR_BAD_CONFIGURATION          As Long = 1610 'The configuration data for this product is corrupt.  Contact your support personnel.
Public Const ERROR_INDEX_ABSENT               As Long = 1611 'Component qualifier not present.
Public Const ERROR_INSTALL_SOURCE_ABSENT      As Long = 1612 'The installation source for this product is not available.  Verify that the source exists and that you can access it.
Public Const ERROR_INSTALL_PACKAGE_VERSION    As Long = 1613 'This installation package cannot be installed by the Windows Installer service.  You must install a Windows service pack that contains a newer version of the Windows Installer service.
Public Const ERROR_PRODUCT_UNINSTALLED        As Long = 1614 'Product is uninstalled.
Public Const ERROR_BAD_QUERY_SYNTAX           As Long = 1615 'SQL query syntax invalid or unsupported.
Public Const ERROR_INVALID_FIELD              As Long = 1616 'Record field does not exist.
Public Const ERROR_DEVICE_REMOVED             As Long = 1617 'The device has been removed.
Public Const ERROR_INSTALL_ALREADY_RUNNING    As Long = 1618 'Another installation is already in progress.  Complete that installation before proceeding with this install.
Public Const ERROR_INSTALL_PACKAGE_OPEN_FAILED As Long = 1619 'This installation package could not be opened.  Verify that the package exists and that you can access it, or contact the application vendor to verify that this is a valid Windows Installer package.
Public Const ERROR_INSTALL_PACKAGE_INVALID    As Long = 1620 'This installation package could not be opened.  Contact the application vendor to verify that this is a valid Windows Installer package.
Public Const ERROR_INSTALL_UI_FAILURE         As Long = 1621 'There was an error starting the Windows Installer service user interface.  Contact your support personnel.
Public Const ERROR_INSTALL_LOG_FAILURE        As Long = 1622 'Error opening installation log file. Verify that the specified log file location exists and that you can write to it.
Public Const ERROR_INSTALL_LANGUAGE_UNSUPPORTED As Long = 1623 'The language of this installation package is not supported by your system.
Public Const ERROR_INSTALL_TRANSFORM_FAILURE  As Long = 1624 'Error applying transforms.  Verify that the specified transform paths are valid.
Public Const ERROR_INSTALL_PACKAGE_REJECTED   As Long = 1625 'This installation is forbidden by system policy.  Contact your system administrator.
Public Const ERROR_FUNCTION_NOT_CALLED        As Long = 1626 'Function could not be executed.
Public Const ERROR_FUNCTION_FAILED            As Long = 1627 'Function failed during execution.
Public Const ERROR_INVALID_TABLE              As Long = 1628 'Invalid or unknown table specified.
Public Const ERROR_DATATYPE_MISMATCH          As Long = 1629 'Data supplied is of wrong type.
Public Const ERROR_UNSUPPORTED_TYPE           As Long = 1630 'Data of this type is not supported.
Public Const ERROR_CREATE_FAILED              As Long = 1631 'The Windows Installer service failed to start.  Contact your support personnel.
Public Const ERROR_INSTALL_TEMP_UNWRITABLE    As Long = 1632 'The Temp folder is on a drive that is full or is inaccessible. Free up space on the drive or verify that you have write permission on the Temp folder.
Public Const ERROR_INSTALL_PLATFORM_UNSUPPORTED As Long = 1633 'This installation package is not supported by this processor type. Contact your product vendor.
Public Const ERROR_INSTALL_NOTUSED            As Long = 1634 'Component not used on this computer.
Public Const ERROR_PATCH_PACKAGE_OPEN_FAILED  As Long = 1635 'This patch package could not be opened.  Verify that the patch package exists and that you can access it, or contact the application vendor to verify that this is a valid Windows Installer patch package.
Public Const ERROR_PATCH_PACKAGE_INVALID      As Long = 1636 'This patch package could not be opened.  Contact the application vendor to verify that this is a valid Windows Installer patch package.
Public Const ERROR_PATCH_PACKAGE_UNSUPPORTED  As Long = 1637 'This patch package cannot be processed by the Windows Installer service.  You must install a Windows service pack that contains a newer version of the Windows Installer service.
Public Const ERROR_PRODUCT_VERSION            As Long = 1638 'Another version of this product is already installed.  Installation of this version cannot continue.  To configure or remove the existing version of this product, use Add/Remove Programs on the Control Panel.
Public Const ERROR_INVALID_COMMAND_LINE       As Long = 1639 'Invalid command line argument.  Consult the Windows Installer SDK for detailed command line help.
Public Const ERROR_INSTALL_REMOTE_DISALLOWED  As Long = 1640 'Only administrators have permission to add, remove, or configure server software during a Terminal services remote session. If you want to install or configure software on the server, contact your network administrator.
Public Const ERROR_SUCCESS_REBOOT_INITIATED   As Long = 1641 'The requested operation completed successfully.  The system will be restarted so the changes can take effect.
Public Const ERROR_PATCH_TARGET_NOT_FOUND     As Long = 1642 'The upgrade patch cannot be installed by the Windows Installer service because the program to be upgraded may be missing, or the upgrade patch may update a different version of the program. Verify that the program to be upgraded exists on your computer an    'd that you have the correct upgrade patch.
Public Const ERROR_PATCH_PACKAGE_REJECTED     As Long = 1643 'The patch package is not permitted by software restriction policy.
Public Const ERROR_INSTALL_TRANSFORM_REJECTED As Long = 1644 'One or more customizations are not permitted by software restriction policy.
Public Const ERROR_INSTALL_REMOTE_PROHIBITED  As Long = 1645 'The Windows Installer does not permit installation from a Remote Desktop Connection.

'
'// End of MSI error codes
'
'
'
'///////////////////////////
'//                       //
'//   RPC Status Codes    //
'//                       //
'///////////////////////////
'
Public Const RPC_S_INVALID_STRING_BINDING     As Long = 1700    'The string binding is invalid.
Public Const RPC_S_WRONG_KIND_OF_BINDING      As Long = 1701    'The binding handle is not the correct type.
Public Const RPC_S_INVALID_BINDING            As Long = 1702    'The binding handle is invalid.
Public Const RPC_S_PROTSEQ_NOT_SUPPORTED      As Long = 1703    'The RPC protocol sequence is not supported.
Public Const RPC_S_INVALID_RPC_PROTSEQ        As Long = 1704    'The RPC protocol sequence is invalid.
Public Const RPC_S_INVALID_STRING_UUID        As Long = 1705    'The string universal unique identifier (UUID) is invalid.
Public Const RPC_S_INVALID_ENDPOINT_FORMAT    As Long = 1706    'The endpoint format is invalid.
Public Const RPC_S_INVALID_NET_ADDR           As Long = 1707    'The network address is invalid.
Public Const RPC_S_NO_ENDPOINT_FOUND          As Long = 1708    'No endpoint was found.
Public Const RPC_S_INVALID_TIMEOUT            As Long = 1709    'The timeout value is invalid.
Public Const RPC_S_OBJECT_NOT_FOUND           As Long = 1710    'The object universal unique identifier (UUID) was not found.
Public Const RPC_S_ALREADY_REGISTERED         As Long = 1711    'The object universal unique identifier (UUID) has already been registered.
Public Const RPC_S_TYPE_ALREADY_REGISTERED    As Long = 1712    'The type universal unique identifier (UUID) has already been registered.
Public Const RPC_S_ALREADY_LISTENING          As Long = 1713    'The RPC server is already listening.
Public Const RPC_S_NO_PROTSEQS_REGISTERED     As Long = 1714    'No protocol sequences have been registered.
Public Const RPC_S_NOT_LISTENING              As Long = 1715    'The RPC server is not listening.
Public Const RPC_S_UNKNOWN_MGR_TYPE           As Long = 1716    'The manager type is unknown.
Public Const RPC_S_UNKNOWN_IF                 As Long = 1717    'The interface is unknown.
Public Const RPC_S_NO_BINDINGS                As Long = 1718    'There are no bindings.
Public Const RPC_S_NO_PROTSEQS                As Long = 1719    'There are no protocol sequences.
Public Const RPC_S_CANT_CREATE_ENDPOINT       As Long = 1720    'The endpoint cannot be created.
Public Const RPC_S_OUT_OF_RESOURCES           As Long = 1721    'Not enough resources are available to complete this operation.
Public Const RPC_S_SERVER_UNAVAILABLE         As Long = 1722    'The RPC server is unavailable.
Public Const RPC_S_SERVER_TOO_BUSY            As Long = 1723    'The RPC server is too busy to complete this operation.
Public Const RPC_S_INVALID_NETWORK_OPTIONS    As Long = 1724    'The network options are invalid.
Public Const RPC_S_NO_CALL_ACTIVE             As Long = 1725    'There are no remote procedure calls active on this thread.
Public Const RPC_S_CALL_FAILED                As Long = 1726    'The remote procedure call failed.
Public Const RPC_S_CALL_FAILED_DNE            As Long = 1727    'The remote procedure call failed and did not execute.
Public Const RPC_S_PROTOCOL_ERROR             As Long = 1728    'A remote procedure call (RPC) protocol error occurred.
Public Const RPC_S_UNSUPPORTED_TRANS_SYN      As Long = 1730    'The transfer syntax is not supported by the RPC server.
Public Const RPC_S_UNSUPPORTED_TYPE           As Long = 1732    'The universal unique identifier (UUID) type is not supported.
Public Const RPC_S_INVALID_TAG                As Long = 1733    'The tag is invalid.
Public Const RPC_S_INVALID_BOUND              As Long = 1734    'The array bounds are invalid.
Public Const RPC_S_NO_ENTRY_NAME              As Long = 1735    'The binding does not contain an entry name.
Public Const RPC_S_INVALID_NAME_SYNTAX        As Long = 1736    'The name syntax is invalid.
Public Const RPC_S_UNSUPPORTED_NAME_SYNTAX    As Long = 1737    'The name syntax is not supported.
Public Const RPC_S_UUID_NO_ADDRESS            As Long = 1739    'No network address is available to use to construct a universal unique identifier (UUID).
Public Const RPC_S_DUPLICATE_ENDPOINT         As Long = 1740    'The endpoint is a duplicate.
Public Const RPC_S_UNKNOWN_AUTHN_TYPE         As Long = 1741    'The authentication type is unknown.
Public Const RPC_S_MAX_CALLS_TOO_SMALL        As Long = 1742    'The maximum number of calls is too small.
Public Const RPC_S_STRING_TOO_LONG            As Long = 1743    'The string is too long.
Public Const RPC_S_PROTSEQ_NOT_FOUND          As Long = 1744    'The RPC protocol sequence was not found.
Public Const RPC_S_PROCNUM_OUT_OF_RANGE       As Long = 1745    'The procedure number is out of range.
Public Const RPC_S_BINDING_HAS_NO_AUTH        As Long = 1746    'The binding does not contain any authentication information.
Public Const RPC_S_UNKNOWN_AUTHN_SERVICE      As Long = 1747    'The authentication service is unknown.
Public Const RPC_S_UNKNOWN_AUTHN_LEVEL        As Long = 1748    'The authentication level is unknown.
Public Const RPC_S_INVALID_AUTH_IDENTITY      As Long = 1749    'The security context is invalid.
Public Const RPC_S_UNKNOWN_AUTHZ_SERVICE      As Long = 1750    'The authorization service is unknown.
Public Const EPT_S_INVALID_ENTRY              As Long = 1751    'The entry is invalid.
Public Const EPT_S_CANT_PERFORM_OP            As Long = 1752    'The server endpoint cannot perform the operation.
Public Const EPT_S_NOT_REGISTERED             As Long = 1753    'There are no more endpoints available from the endpoint mapper.
Public Const RPC_S_NOTHING_TO_EXPORT          As Long = 1754    'No interfaces have been exported.
Public Const RPC_S_INCOMPLETE_NAME            As Long = 1755    'The entry name is incomplete.
Public Const RPC_S_INVALID_VERS_OPTION        As Long = 1756    'The version option is invalid.
Public Const RPC_S_NO_MORE_MEMBERS            As Long = 1757    'There are no more members.
Public Const RPC_S_NOT_ALL_OBJS_UNEXPORTED    As Long = 1758    'There is nothing to unexport.
Public Const RPC_S_INTERFACE_NOT_FOUND        As Long = 1759    'The interface was not found.
Public Const RPC_S_ENTRY_ALREADY_EXISTS       As Long = 1760    'The entry already exists.
Public Const RPC_S_ENTRY_NOT_FOUND            As Long = 1761    'The entry is not found.
Public Const RPC_S_NAME_SERVICE_UNAVAILABLE   As Long = 1762    'The name service is unavailable.
Public Const RPC_S_INVALID_NAF_ID             As Long = 1763    'The network address family is invalid.
Public Const RPC_S_CANNOT_SUPPORT             As Long = 1764    'The requested operation is not supported.
Public Const RPC_S_NO_CONTEXT_AVAILABLE       As Long = 1765    'No security context is available to allow impersonation.
Public Const RPC_S_INTERNAL_ERROR             As Long = 1766    'An internal error occurred in a remote procedure call (RPC).
Public Const RPC_S_ZERO_DIVIDE                As Long = 1767    'The RPC server attempted an integer division by zero.
Public Const RPC_S_ADDRESS_ERROR              As Long = 1768    'An addressing error occurred in the RPC server.
Public Const RPC_S_FP_DIV_ZERO                As Long = 1769    'A floating-point operation at the RPC server caused a division by zero.
Public Const RPC_S_FP_UNDERFLOW               As Long = 1770    'A floating-point underflow occurred at the RPC server.
Public Const RPC_S_FP_OVERFLOW                As Long = 1771    'A floating-point overflow occurred at the RPC server.
Public Const RPC_X_NO_MORE_ENTRIES            As Long = 1772    'The list of RPC servers available for the binding of auto handles has been exhausted.
Public Const RPC_X_SS_CHAR_TRANS_OPEN_FAIL    As Long = 1773    'Unable to open the character translation table file.
Public Const RPC_X_SS_CHAR_TRANS_SHORT_FILE   As Long = 1774    'The file containing the character translation table has fewer than 512 bytes.
Public Const RPC_X_SS_IN_NULL_CONTEXT         As Long = 1775    'A null context handle was passed from the client to the host during a remote procedure call.
Public Const RPC_X_SS_CONTEXT_DAMAGED         As Long = 1777    'The context handle changed during a remote procedure call.
Public Const RPC_X_SS_HANDLES_MISMATCH        As Long = 1778    'The binding handles passed to a remote procedure call do not match.
Public Const RPC_X_SS_CANNOT_GET_CALL_HANDLE  As Long = 1779    'The stub is unable to get the remote procedure call handle.
Public Const RPC_X_NULL_REF_POINTER           As Long = 1780    'A null reference pointer was passed to the stub.
Public Const RPC_X_ENUM_VALUE_OUT_OF_RANGE    As Long = 1781    'The enumeration value is out of range.
Public Const RPC_X_BYTE_COUNT_TOO_SMALL       As Long = 1782    'The byte count is too small.
Public Const RPC_X_BAD_STUB_DATA              As Long = 1783    'The stub received bad data.
Public Const ERROR_INVALID_USER_BUFFER        As Long = 1784    'The supplied user buffer is not valid for the requested operation.
Public Const ERROR_UNRECOGNIZED_MEDIA         As Long = 1785    'The disk media is not recognized. It may not be formatted.
Public Const ERROR_NO_TRUST_LSA_SECRET        As Long = 1786    'The workstation does not have a trust secret.
Public Const ERROR_NO_TRUST_SAM_ACCOUNT       As Long = 1787    'The security database on the server does not have a computer account for this workstation trust relationship.
Public Const ERROR_TRUSTED_DOMAIN_FAILURE     As Long = 1788    'The trust relationship between the primary domain and the trusted domain failed.
Public Const ERROR_TRUSTED_RELATIONSHIP_FAILURE As Long = 1789  'The trust relationship between this workstation and the primary domain failed.
Public Const ERROR_TRUST_FAILURE              As Long = 1790    'The network logon failed.
Public Const RPC_S_CALL_IN_PROGRESS           As Long = 1791    'A remote procedure call is already in progress for this thread.
Public Const ERROR_NETLOGON_NOT_STARTED       As Long = 1792    'An attempt was made to logon, but the network logon service was not started.
Public Const ERROR_ACCOUNT_EXPIRED            As Long = 1793    'The user's account has expired.
Public Const ERROR_REDIRECTOR_HAS_OPEN_HANDLES As Long = 1794   'The redirector is in use and cannot be unloaded.
Public Const ERROR_PRINTER_DRIVER_ALREADY_INSTALLED As Long = 1795    'The specified printer driver is already installed.
Public Const ERROR_UNKNOWN_PORT               As Long = 1796    'The specified port is unknown.
Public Const ERROR_UNKNOWN_PRINTER_DRIVER     As Long = 1797    'The printer driver is unknown.
Public Const ERROR_UNKNOWN_PRINTPROCESSOR     As Long = 1798    'The print processor is unknown.
Public Const ERROR_INVALID_SEPARATOR_FILE     As Long = 1799    'The specified separator file is invalid.
Public Const ERROR_INVALID_PRIORITY           As Long = 1800    'The specified priority is invalid.
Public Const ERROR_INVALID_PRINTER_NAME       As Long = 1801    'The printer name is invalid.
Public Const ERROR_PRINTER_ALREADY_EXISTS     As Long = 1802    'The printer already exists.
Public Const ERROR_INVALID_PRINTER_COMMAND    As Long = 1803    'The printer command is invalid.
Public Const ERROR_INVALID_DATATYPE           As Long = 1804    'The specified datatype is invalid.
Public Const ERROR_INVALID_ENVIRONMENT        As Long = 1805    'The environment specified is invalid.
Public Const RPC_S_NO_MORE_BINDINGS           As Long = 1806    'There are no more bindings.
Public Const ERROR_NOLOGON_INTERDOMAIN_TRUST_ACCOUNT As Long = 1807    'The account used is an interdomain trust account. Use your global user account or local user account to access this server.
Public Const ERROR_NOLOGON_WORKSTATION_TRUST_ACCOUNT As Long = 1808    'The account used is a computer account. Use your global user account or local user account to access this server.
Public Const ERROR_NOLOGON_SERVER_TRUST_ACCOUNT As Long = 1809  'The account used is a server trust account. Use your global user account or local user account to access this server.
Public Const ERROR_DOMAIN_TRUST_INCONSISTENT  As Long = 1810    'The name or security ID (SID) of the domain specified is inconsistent with the trust information for that domain.
Public Const ERROR_SERVER_HAS_OPEN_HANDLES    As Long = 1811    'The server is in use and cannot be unloaded.
Public Const ERROR_RESOURCE_DATA_NOT_FOUND    As Long = 1812    'The specified image file did not contain a resource section.
Public Const ERROR_RESOURCE_TYPE_NOT_FOUND    As Long = 1813    'The specified resource type cannot be found in the image file.
Public Const ERROR_RESOURCE_NAME_NOT_FOUND    As Long = 1814    'The specified resource name cannot be found in the image file.
Public Const ERROR_RESOURCE_LANG_NOT_FOUND    As Long = 1815    'The specified resource language ID cannot be found in the image file.
Public Const ERROR_NOT_ENOUGH_QUOTA           As Long = 1816    'Not enough quota is available to process this command.
Public Const RPC_S_NO_INTERFACES              As Long = 1817    'No interfaces have been registered.
Public Const RPC_S_CALL_CANCELLED             As Long = 1818    'The remote procedure call was cancelled.
Public Const RPC_S_BINDING_INCOMPLETE         As Long = 1819    'The binding handle does not contain all required information.
Public Const RPC_S_COMM_FAILURE               As Long = 1820    'A communications failure occurred during a remote procedure call.
Public Const RPC_S_UNSUPPORTED_AUTHN_LEVEL    As Long = 1821    'The requested authentication level is not supported.
Public Const RPC_S_NO_PRINC_NAME              As Long = 1822    'No principal name registered.
Public Const RPC_S_NOT_RPC_ERROR              As Long = 1823    'The error specified is not a valid Windows RPC error code.
Public Const RPC_S_UUID_LOCAL_ONLY            As Long = 1824    'A UUID that is valid only on this computer has been allocated.
Public Const RPC_S_SEC_PKG_ERROR              As Long = 1825    'A security package specific error occurred.
Public Const RPC_S_NOT_CANCELLED              As Long = 1826    'Thread is not canceled.
Public Const RPC_X_INVALID_ES_ACTION          As Long = 1827    'Invalid operation on the encoding/decoding handle.
Public Const RPC_X_WRONG_ES_VERSION           As Long = 1828    'Incompatible version of the serializing package.
Public Const RPC_X_WRONG_STUB_VERSION         As Long = 1829    'Incompatible version of the RPC stub.
Public Const RPC_X_INVALID_PIPE_OBJECT        As Long = 1830    'The RPC pipe object is invalid or corrupted.
Public Const RPC_X_WRONG_PIPE_ORDER           As Long = 1831    'An invalid operation was attempted on an RPC pipe object.
Public Const RPC_X_WRONG_PIPE_VERSION         As Long = 1832    'Unsupported RPC pipe version.
Public Const RPC_S_GROUP_MEMBER_NOT_FOUND     As Long = 1898    'The group member was not found.
Public Const EPT_S_CANT_CREATE                As Long = 1899    'The endpoint mapper database entry could not be created.
Public Const RPC_S_INVALID_OBJECT             As Long = 1900    'The object universal unique identifier (UUID) is the nil UUID.
Public Const ERROR_INVALID_TIME               As Long = 1901    'The specified time is invalid.
Public Const ERROR_INVALID_FORM_NAME          As Long = 1902    'The specified form name is invalid.
Public Const ERROR_INVALID_FORM_SIZE          As Long = 1903    'The specified form size is invalid.
Public Const ERROR_ALREADY_WAITING            As Long = 1904    'The specified printer handle is already being waited on
Public Const ERROR_PRINTER_DELETED            As Long = 1905    'The specified printer has been deleted.
Public Const ERROR_INVALID_PRINTER_STATE      As Long = 1906    'The state of the printer is invalid.
Public Const ERROR_PASSWORD_MUST_CHANGE       As Long = 1907    'The user's password must be changed before logging on the first time.
Public Const ERROR_DOMAIN_CONTROLLER_NOT_FOUND As Long = 1908   'Could not find the domain controller for this domain.
Public Const ERROR_ACCOUNT_LOCKED_OUT         As Long = 1909    'The referenced account is currently locked out and may not be logged on to.
Public Const OR_INVALID_OXID                  As Long = 1910    'The object exporter specified was not found.
Public Const OR_INVALID_OID                   As Long = 1911    'The object specified was not found.
Public Const OR_INVALID_SET                   As Long = 1912    'The object resolver set specified was not found.
Public Const RPC_S_SEND_INCOMPLETE            As Long = 1913    'Some data remains to be sent in the request buffer.
Public Const RPC_S_INVALID_ASYNC_HANDLE       As Long = 1914    'Invalid asynchronous remote procedure call handle.
Public Const RPC_S_INVALID_ASYNC_CALL         As Long = 1915    'Invalid asynchronous RPC call handle for this operation.
Public Const RPC_X_PIPE_CLOSED                As Long = 1916    'The RPC pipe object has already been closed.
Public Const RPC_X_PIPE_DISCIPLINE_ERROR      As Long = 1917    'The RPC call completed before all pipes were processed.
Public Const RPC_X_PIPE_EMPTY                 As Long = 1918    'No more data is available from the RPC pipe.
Public Const ERROR_NO_SITENAME                As Long = 1919    'No site name is available for this machine.
Public Const ERROR_CANT_ACCESS_FILE           As Long = 1920    'The file can not be accessed by the system.
Public Const ERROR_CANT_RESOLVE_FILENAME      As Long = 1921    'The name of the file cannot be resolved by the system.
Public Const RPC_S_ENTRY_TYPE_MISMATCH        As Long = 1922    'The entry is not of the expected type.
Public Const RPC_S_NOT_ALL_OBJS_EXPORTED      As Long = 1923    'Not all object UUIDs could be exported to the specified entry.
Public Const RPC_S_INTERFACE_NOT_EXPORTED     As Long = 1924    'Interface could not be exported to the specified entry.
Public Const RPC_S_PROFILE_NOT_ADDED          As Long = 1925    'The specified profile entry could not be added.
Public Const RPC_S_PRF_ELT_NOT_ADDED          As Long = 1926    'The specified profile element could not be added.
Public Const RPC_S_PRF_ELT_NOT_REMOVED        As Long = 1927    'The specified profile element could not be removed.
Public Const RPC_S_GRP_ELT_NOT_ADDED          As Long = 1928    'The group element could not be added.
Public Const RPC_S_GRP_ELT_NOT_REMOVED        As Long = 1929    'The group element could not be removed.
Public Const ERROR_KM_DRIVER_BLOCKED          As Long = 1930    'The printer driver is not compatible with a policy enabled on your computer that blocks NT 4.0 drivers.
Public Const ERROR_CONTEXT_EXPIRED            As Long = 1931    'The context has expired and can no longer be used.
Public Const ERROR_PER_USER_TRUST_QUOTA_EXCEEDED As Long = 1932 'The current user's delegated trust creation quota has been exceeded.
Public Const ERROR_ALL_USER_TRUST_QUOTA_EXCEEDED As Long = 1933 'The total delegated trust creation quota has been exceeded.
Public Const ERROR_USER_DELETE_TRUST_QUOTA_EXCEEDED As Long = 1934 'The current user's delegated trust deletion quota has been exceeded.
Public Const ERROR_AUTHENTICATION_FIREWALL_FAILED As Long = 1935 'Logon Failure: The machine you are logging onto is protected by an authentication firewall.  The specified account is not allowed to authenticate to the machine.
Public Const ERROR_REMOTE_PRINT_CONNECTIONS_BLOCKED As Long = 1936 'Remote connections to the Print Spooler are blocked by a policy set on your machine.
'
'
'
'
'///////////////////////////
'//                       //
'//   OpenGL Error Code   //
'//                       //
'///////////////////////////
'
'
Public Const ERROR_INVALID_PIXEL_FORMAT       As Long = 2000    'The pixel format is invalid.
Public Const ERROR_BAD_DRIVER                 As Long = 2001    'The specified driver is invalid.
Public Const ERROR_INVALID_WINDOW_STYLE       As Long = 2002    'The window style or class attribute is invalid for this operation.
Public Const ERROR_METAFILE_NOT_SUPPORTED     As Long = 2003    'The requested metafile operation is not supported.
Public Const ERROR_TRANSFORM_NOT_SUPPORTED    As Long = 2004    'The requested transformation operation is not supported.
Public Const ERROR_CLIPPING_NOT_SUPPORTED     As Long = 2005    'The requested clipping operation is not supported.
'
'// End of OpenGL error codes
'
'
'
'///////////////////////////////////////////
'//                                       //
'//   Image Color Management Error Code   //
'//                                       //
'///////////////////////////////////////////
'
Public Const ERROR_INVALID_CMM                As Long = 2010    'The specified color management module is invalid.
Public Const ERROR_INVALID_PROFILE            As Long = 2011    'The specified color profile is invalid.
Public Const ERROR_TAG_NOT_FOUND              As Long = 2012    'The specified tag was not found.
Public Const ERROR_TAG_NOT_PRESENT            As Long = 2013    'A required tag is not present.
Public Const ERROR_DUPLICATE_TAG              As Long = 2014    'The specified tag is already present.
Public Const ERROR_PROFILE_NOT_ASSOCIATED_WITH_DEVICE As Long = 2015    'The specified color profile is not associated with any device.
Public Const ERROR_PROFILE_NOT_FOUND          As Long = 2016    'The specified color profile was not found.
Public Const ERROR_INVALID_COLORSPACE         As Long = 2017    'The specified color space is invalid.
Public Const ERROR_ICM_NOT_ENABLED            As Long = 2018    'Image Color Management is not enabled.
Public Const ERROR_DELETING_ICM_XFORM         As Long = 2019    'There was an error while deleting the color transform.
Public Const ERROR_INVALID_TRANSFORM          As Long = 2020    'The specified color transform is invalid.
Public Const ERROR_COLORSPACE_MISMATCH        As Long = 2021    'The specified transform does not match the bitmap's color space.
Public Const ERROR_INVALID_COLORINDEX         As Long = 2022    'The specified named color index is not present in the profile.

'
'
'
'
'///////////////////////////
'//                       //
'// Winnet32 Status Codes //
'//                       //
'// The range 2100 through 2999 is reserved for network status codes.
'// See lmerr.h for a complete listing
'///////////////////////////
'
Public Const ERROR_CONNECTED_OTHER_PASSWORD   As Long = 2108    'The network connection was made successfully, but the user had to be prompted for a password other than the one originally specified.
Public Const ERROR_CONNECTED_OTHER_PASSWORD_DEFAULT As Long = 2109 'The network connection was made successfully using default credentials.
Public Const ERROR_BAD_USERNAME               As Long = 2202    'The specified username is invalid.
Public Const ERROR_NOT_CONNECTED              As Long = 2250    'This network connection does not exist.
Public Const ERROR_OPEN_FILES                 As Long = 2401    'This network connection has files open or requests pending.
Public Const ERROR_ACTIVE_CONNECTIONS         As Long = 2402    'Active connections still exist.
Public Const ERROR_DEVICE_IN_USE              As Long = 2404    'The device is in use by an active process and cannot be disconnected.

'
'
'////////////////////////////////////
'//                                //
'//     Win32 Spooler Error Codes  //
'//                                //
'////////////////////////////////////
Public Const ERROR_UNKNOWN_PRINT_MONITOR      As Long = 3000    'The specified print monitor is unknown.
Public Const ERROR_PRINTER_DRIVER_IN_USE      As Long = 3001    'The specified printer driver is currently in use.
Public Const ERROR_SPOOL_FILE_NOT_FOUND       As Long = 3002    'The spool file was not found.
Public Const ERROR_SPL_NO_STARTDOC            As Long = 3003    'A StartDocPrinter call was not issued.
Public Const ERROR_SPL_NO_ADDJOB              As Long = 3004    'An AddJob call was not issued.
Public Const ERROR_PRINT_PROCESSOR_ALREADY_INSTALLED As Long = 3005    'The specified print processor has already been installed.
Public Const ERROR_PRINT_MONITOR_ALREADY_INSTALLED As Long = 3006    'The specified print monitor has already been installed.
Public Const ERROR_INVALID_PRINT_MONITOR      As Long = 3007    'The specified print monitor does not have the required functions.
Public Const ERROR_PRINT_MONITOR_IN_USE       As Long = 3008    'The specified print monitor is currently in use.
Public Const ERROR_PRINTER_HAS_JOBS_QUEUED    As Long = 3009    'The requested operation is not allowed when there are jobs queued to the printer.
Public Const ERROR_SUCCESS_REBOOT_REQUIRED    As Long = 3010    'The requested operation is successful. Changes will not be effective until the system is rebooted.
Public Const ERROR_SUCCESS_RESTART_REQUIRED   As Long = 3011    'The requested operation is successful. Changes will not be effective until the service is restarted.
Public Const ERROR_PRINTER_NOT_FOUND          As Long = 3012    'No printers were found.
Public Const ERROR_PRINTER_DRIVER_WARNED      As Long = 3013    'The printer driver is known to be unreliable.
Public Const ERROR_PRINTER_DRIVER_BLOCKED     As Long = 3014    'The printer driver is known to harm the system.

'
'////////////////////////////////////
'//                                //
'//     Wins Error Codes           //
'//                                //
'////////////////////////////////////
Public Const ERROR_WINS_INTERNAL              As Long = 4000    'WINS encountered an error while processing the command.
Public Const ERROR_CAN_NOT_DEL_LOCAL_WINS     As Long = 4001    'The local WINS can not be deleted.
Public Const ERROR_STATIC_INIT                As Long = 4002    'The importation from the file failed.
Public Const ERROR_INC_BACKUP                 As Long = 4003    'The backup failed. Was a full backup done before?
Public Const ERROR_FULL_BACKUP                As Long = 4004    'The backup failed. Check the directory to which you are backing the database.
Public Const ERROR_REC_NON_EXISTENT           As Long = 4005    'The name does not exist in the WINS database.
Public Const ERROR_RPL_NOT_ALLOWED            As Long = 4006    'Replication with a nonconfigured partner is not allowed.

'
'////////////////////////////////////
'//                                //
'//     DHCP Error Codes           //
'//                                //
'////////////////////////////////////
'//
'// MessageId: ERROR_DHCP_ADDRESS_CONFLICT
'//
'// MessageText:
'//
'//  The DHCP client has obtained an IP address that is already in use on the network. The local interface will be disabled until the DHCP client can obtain a new address.
'//
'Public Const ERROR_DHCP_ADDRESS_CONFLICT      4100L
'
'////////////////////////////////////
'//                                //
'//     WMI Error Codes            //
'//                                //
'////////////////////////////////////
Public Const ERROR_WMI_GUID_NOT_FOUND         As Long = 4200    'The Guid passed was not recognized as valid by a WMI data provider.
Public Const ERROR_WMI_INSTANCE_NOT_FOUND     As Long = 4201    'The instance name passed was not recognized as valid by a WMI data provider.
Public Const ERROR_WMI_ITEMID_NOT_FOUND       As Long = 4202    'The data item ID passed was not recognized as valid by a WMI data provider.
Public Const ERROR_WMI_TRY_AGAIN              As Long = 4203    'The WMI request could not be completed and should be retried.
Public Const ERROR_WMI_DP_NOT_FOUND           As Long = 4204    'The WMI data provider could not be located.
Public Const ERROR_WMI_UNRESOLVED_INSTANCE_REF As Long = 4205    'The WMI data provider references an instance set that has not been registered.
Public Const ERROR_WMI_ALREADY_ENABLED        As Long = 4206    'The WMI data block or event notification has already been enabled.
Public Const ERROR_WMI_GUID_DISCONNECTED      As Long = 4207    'The WMI data block is no longer available.
Public Const ERROR_WMI_SERVER_UNAVAILABLE     As Long = 4208    'The WMI data service is not available.
Public Const ERROR_WMI_DP_FAILED              As Long = 4209    'The WMI data provider failed to carry out the request.
Public Const ERROR_WMI_INVALID_MOF            As Long = 4210    'The WMI MOF information is not valid.
Public Const ERROR_WMI_INVALID_REGINFO        As Long = 4211    'The WMI registration information is not valid.
Public Const ERROR_WMI_ALREADY_DISABLED       As Long = 4212    'The WMI data block or event notification has already been disabled.
Public Const ERROR_WMI_READ_ONLY              As Long = 4213    'The WMI data item or data block is read only.
Public Const ERROR_WMI_SET_FAILURE            As Long = 4214    'The WMI data item or data block could not be changed.

'
'//////////////////////////////////////////
'//                                      //
'// NT Media Services (RSM) Error Codes  //
'//                                      //
'//////////////////////////////////////////
Public Const ERROR_INVALID_MEDIA              As Long = 4300    'The media identifier does not represent a valid medium.
Public Const ERROR_INVALID_LIBRARY            As Long = 4301    'The library identifier does not represent a valid library.
Public Const ERROR_INVALID_MEDIA_POOL         As Long = 4302    'The media pool identifier does not represent a valid media pool.
Public Const ERROR_DRIVE_MEDIA_MISMATCH       As Long = 4303    'The drive and medium are not compatible or exist in different libraries.
Public Const ERROR_MEDIA_OFFLINE              As Long = 4304    'The medium currently exists in an offline library and must be online to perform this operation.
Public Const ERROR_LIBRARY_OFFLINE            As Long = 4305    'The operation cannot be performed on an offline library.
Public Const ERROR_EMPTY                      As Long = 4306    'The library, drive, or media pool is empty.
Public Const ERROR_NOT_EMPTY                  As Long = 4307    'The library, drive, or media pool must be empty to perform this operation.
Public Const ERROR_MEDIA_UNAVAILABLE          As Long = 4308    'No media is currently available in this media pool or library.
Public Const ERROR_RESOURCE_DISABLED          As Long = 4309    'A resource required for this operation is disabled.
Public Const ERROR_INVALID_CLEANER            As Long = 4310    'The media identifier does not represent a valid cleaner.
Public Const ERROR_UNABLE_TO_CLEAN            As Long = 4311    'The drive cannot be cleaned or does not support cleaning.
Public Const ERROR_OBJECT_NOT_FOUND           As Long = 4312    'The object identifier does not represent a valid object.
Public Const ERROR_DATABASE_FAILURE           As Long = 4313    'Unable to read from or write to the database.
Public Const ERROR_DATABASE_FULL              As Long = 4314    'The database is full.
Public Const ERROR_MEDIA_INCOMPATIBLE         As Long = 4315    'The medium is not compatible with the device or media pool.
Public Const ERROR_RESOURCE_NOT_PRESENT       As Long = 4316    'The resource required for this operation does not exist.
Public Const ERROR_INVALID_OPERATION          As Long = 4317    'The operation identifier is not valid.
Public Const ERROR_MEDIA_NOT_AVAILABLE        As Long = 4318    'The media is not mounted or ready for use.
Public Const ERROR_DEVICE_NOT_AVAILABLE       As Long = 4319    'The device is not ready for use.
Public Const ERROR_REQUEST_REFUSED            As Long = 4320    'The operator or administrator has refused the request.
Public Const ERROR_INVALID_DRIVE_OBJECT       As Long = 4321    'The drive identifier does not represent a valid drive.
Public Const ERROR_LIBRARY_FULL               As Long = 4322    'Library is full.  No slot is available for use.
Public Const ERROR_MEDIUM_NOT_ACCESSIBLE      As Long = 4323    'The transport cannot access the medium.
Public Const ERROR_UNABLE_TO_LOAD_MEDIUM      As Long = 4324    'Unable to load the medium into the drive.
Public Const ERROR_UNABLE_TO_INVENTORY_DRIVE  As Long = 4325    'Unable to retrieve the drive status.
Public Const ERROR_UNABLE_TO_INVENTORY_SLOT   As Long = 4326    'Unable to retrieve the slot status.
Public Const ERROR_UNABLE_TO_INVENTORY_TRANSPORT As Long = 4327 'Unable to retrieve status about the transport.
Public Const ERROR_TRANSPORT_FULL             As Long = 4328    'Cannot use the transport because it is already in use.
Public Const ERROR_CONTROLLING_IEPORT         As Long = 4329    'Unable to open or close the inject/eject port.
Public Const ERROR_UNABLE_TO_EJECT_MOUNTED_MEDIA As Long = 4330 'Unable to eject the medium because it is in a drive.
Public Const ERROR_CLEANER_SLOT_SET           As Long = 4331    'A cleaner slot is already reserved.
Public Const ERROR_CLEANER_SLOT_NOT_SET       As Long = 4332    'A cleaner slot is not reserved.
Public Const ERROR_CLEANER_CARTRIDGE_SPENT    As Long = 4333    'The cleaner cartridge has performed the maximum number of drive cleanings.
Public Const ERROR_UNEXPECTED_OMID            As Long = 4334    'Unexpected on-medium identifier.
Public Const ERROR_CANT_DELETE_LAST_ITEM      As Long = 4335    'The last remaining item in this group or resource cannot be deleted.
Public Const ERROR_MESSAGE_EXCEEDS_MAX_SIZE   As Long = 4336    'The message provided exceeds the maximum size allowed for this parameter.
Public Const ERROR_VOLUME_CONTAINS_SYS_FILES  As Long = 4337    'The volume contains system or paging files.
Public Const ERROR_INDIGENOUS_TYPE            As Long = 4338    'The media type cannot be removed from this library since at least one drive in the library reports it can support this media type.
Public Const ERROR_NO_SUPPORTING_DRIVES       As Long = 4339    'This offline media cannot be mounted on this system since no enabled drives are present which can be used.
Public Const ERROR_CLEANER_CARTRIDGE_INSTALLED As Long = 4340   'A cleaner cartridge is present in the tape library.
Public Const ERROR_IEPORT_FULL                As Long = 4341    'Cannot use the ieport because it is not empty.

'
'////////////////////////////////////////////
'//                                        //
'// NT Remote Storage Service Error Codes  //
'//                                        //
'////////////////////////////////////////////
'
'Public Const ERROR_FILE_OFFLINE               4350L'The remote storage service was not able to recall the file.
'Public Const ERROR_REMOTE_STORAGE_NOT_ACTIVE  4351L'The remote storage service is not operational at this time.
'Public Const ERROR_REMOTE_STORAGE_MEDIA_ERROR 4352L'The remote storage service encountered a media error.
'
'////////////////////////////////////////////
'//                                        //
'// NT Reparse Points Error Codes          //
'//                                        //
'////////////////////////////////////////////
'
Public Const ERROR_NOT_A_REPARSE_POINT        As Long = 4390    'The file or directory is not a reparse point.
Public Const ERROR_REPARSE_ATTRIBUTE_CONFLICT As Long = 4391    'The reparse point attribute cannot be set because it conflicts with an existing attribute.
Public Const ERROR_INVALID_REPARSE_DATA       As Long = 4392    'The data present in the reparse point buffer is invalid.
Public Const ERROR_REPARSE_TAG_INVALID        As Long = 4393    'The tag present in the reparse point buffer is invalid.
Public Const ERROR_REPARSE_TAG_MISMATCH       As Long = 4394    'There is a mismatch between the tag specified in the request and the tag present in the reparse point.
'
'////////////////////////////////////////////
'//                                        //
'// NT Single Instance Store Error Codes   //
'//                                        //
'////////////////////////////////////////////
'
'Public Const ERROR_VOLUME_NOT_SIS_ENABLED     4500L'Single Instance Storage is not available on this volume.
'
'////////////////////////////////////
'//                                //
'//     Cluster Error Codes        //
'//                                //
'////////////////////////////////////
'
Public Const ERROR_DEPENDENT_RESOURCE_EXISTS  As Long = 5001    'The cluster resource cannot be moved to another group because other resources are dependent on it.
Public Const ERROR_DEPENDENCY_NOT_FOUND       As Long = 5002    'The cluster resource dependency cannot be found.
Public Const ERROR_DEPENDENCY_ALREADY_EXISTS  As Long = 5003    'The cluster resource cannot be made dependent on the specified resource because it is already dependent.
Public Const ERROR_RESOURCE_NOT_ONLINE        As Long = 5004    'The cluster resource is not online.
Public Const ERROR_HOST_NODE_NOT_AVAILABLE    As Long = 5005    'A cluster node is not available for this operation.
Public Const ERROR_RESOURCE_NOT_AVAILABLE     As Long = 5006    'The cluster resource is not available.
Public Const ERROR_RESOURCE_NOT_FOUND         As Long = 5007    'The cluster resource could not be found.
Public Const ERROR_SHUTDOWN_CLUSTER           As Long = 5008    'The cluster is being shut down.
Public Const ERROR_CANT_EVICT_ACTIVE_NODE     As Long = 5009    'A cluster node cannot be evicted from the cluster unless the node is down or it is the last node.
Public Const ERROR_OBJECT_ALREADY_EXISTS      As Long = 5010    'The object already exists.
Public Const ERROR_OBJECT_IN_LIST             As Long = 5011    'The object is already in the list.
Public Const ERROR_GROUP_NOT_AVAILABLE        As Long = 5012    'The cluster group is not available for any new requests.
Public Const ERROR_GROUP_NOT_FOUND            As Long = 5013    'The cluster group could not be found.
Public Const ERROR_GROUP_NOT_ONLINE           As Long = 5014    'The operation could not be completed because the cluster group is not online.
Public Const ERROR_HOST_NODE_NOT_RESOURCE_OWNER As Long = 5015  'The cluster node is not the owner of the resource.
Public Const ERROR_HOST_NODE_NOT_GROUP_OWNER  As Long = 5016    'The cluster node is not the owner of the group.
Public Const ERROR_RESMON_CREATE_FAILED       As Long = 5017    'The cluster resource could not be created in the specified resource monitor.
Public Const ERROR_RESMON_ONLINE_FAILED       As Long = 5018    'The cluster resource could not be brought online by the resource monitor.
Public Const ERROR_RESOURCE_ONLINE            As Long = 5019    'The operation could not be completed because the cluster resource is online.
Public Const ERROR_QUORUM_RESOURCE            As Long = 5020    'The cluster resource could not be deleted or brought offline because it is the quorum resource.
Public Const ERROR_NOT_QUORUM_CAPABLE         As Long = 5021    'The cluster could not make the specified resource a quorum resource because it is not capable of being a quorum resource.
Public Const ERROR_CLUSTER_SHUTTING_DOWN      As Long = 5022    'The cluster software is shutting down.
Public Const ERROR_INVALID_STATE              As Long = 5023    'The group or resource is not in the correct state to perform the requested operation.
Public Const ERROR_RESOURCE_PROPERTIES_STORED As Long = 5024    'The properties were stored but not all changes will take effect until the next time the resource is brought online.
Public Const ERROR_NOT_QUORUM_CLASS           As Long = 5025    'The cluster could not make the specified resource a quorum resource because it does not belong to a shared storage class.
Public Const ERROR_CORE_RESOURCE              As Long = 5026    'The cluster resource could not be deleted since it is a core resource.
Public Const ERROR_QUORUM_RESOURCE_ONLINE_FAILED As Long = 5027 'The quorum resource failed to come online.
Public Const ERROR_QUORUMLOG_OPEN_FAILED      As Long = 5028    'The quorum log could not be created or mounted successfully.
Public Const ERROR_CLUSTERLOG_CORRUPT         As Long = 5029    'The cluster log is corrupt.
Public Const ERROR_CLUSTERLOG_RECORD_EXCEEDS_MAXSIZE As Long = 5030    'The record could not be written to the cluster log since it exceeds the maximum size.
Public Const ERROR_CLUSTERLOG_EXCEEDS_MAXSIZE As Long = 5031    'The cluster log exceeds its maximum size.
Public Const ERROR_CLUSTERLOG_CHKPOINT_NOT_FOUND As Long = 5032 'No checkpoint record was found in the cluster log.
Public Const ERROR_CLUSTERLOG_NOT_ENOUGH_SPACE As Long = 5033   'The minimum required disk space needed for logging is not available.
Public Const ERROR_QUORUM_OWNER_ALIVE         As Long = 5034    'The cluster node failed to take control of the quorum resource because the resource is owned by another active node.
Public Const ERROR_NETWORK_NOT_AVAILABLE      As Long = 5035    'A cluster network is not available for this operation.
Public Const ERROR_NODE_NOT_AVAILABLE         As Long = 5036    'A cluster node is not available for this operation.
Public Const ERROR_ALL_NODES_NOT_AVAILABLE    As Long = 5037    'All cluster nodes must be running to perform this operation.
Public Const ERROR_RESOURCE_FAILED            As Long = 5038    'A cluster resource failed.
Public Const ERROR_CLUSTER_INVALID_NODE       As Long = 5039    'The cluster node is not valid.
Public Const ERROR_CLUSTER_NODE_EXISTS        As Long = 5040    'The cluster node already exists.
Public Const ERROR_CLUSTER_JOIN_IN_PROGRESS   As Long = 5041    'A node is in the process of joining the cluster.
Public Const ERROR_CLUSTER_NODE_NOT_FOUND     As Long = 5042    'The cluster node was not found.
Public Const ERROR_CLUSTER_LOCAL_NODE_NOT_FOUND As Long = 5043  'The cluster local node information was not found.
Public Const ERROR_CLUSTER_NETWORK_EXISTS     As Long = 5044    'The cluster network already exists.
Public Const ERROR_CLUSTER_NETWORK_NOT_FOUND  As Long = 5045    'The cluster network was not found.
Public Const ERROR_CLUSTER_NETINTERFACE_EXISTS As Long = 5046   'The cluster network interface already exists.
Public Const ERROR_CLUSTER_NETINTERFACE_NOT_FOUND As Long = 5047 'The cluster network interface was not found.
Public Const ERROR_CLUSTER_INVALID_REQUEST    As Long = 5048    'The cluster request is not valid for this object.
Public Const ERROR_CLUSTER_INVALID_NETWORK_PROVIDER As Long = 5049    'The cluster network provider is not valid.
Public Const ERROR_CLUSTER_NODE_DOWN          As Long = 5050    'The cluster node is down.
Public Const ERROR_CLUSTER_NODE_UNREACHABLE   As Long = 5051    'The cluster node is not reachable.
Public Const ERROR_CLUSTER_NODE_NOT_MEMBER    As Long = 5052    'The cluster node is not a member of the cluster.
Public Const ERROR_CLUSTER_JOIN_NOT_IN_PROGRESS As Long = 5053  'A cluster join operation is not in progress.
Public Const ERROR_CLUSTER_INVALID_NETWORK    As Long = 5054    'The cluster network is not valid.
Public Const ERROR_CLUSTER_NODE_UP            As Long = 5056    'The cluster node is up.
Public Const ERROR_CLUSTER_IPADDR_IN_USE      As Long = 5057    'The cluster IP address is already in use.
Public Const ERROR_CLUSTER_NODE_NOT_PAUSED    As Long = 5058    'The cluster node is not paused.
Public Const ERROR_CLUSTER_NO_SECURITY_CONTEXT As Long = 5059   'No cluster security context is available.
Public Const ERROR_CLUSTER_NETWORK_NOT_INTERNAL As Long = 5060  'The cluster network is not configured for internal cluster communication.
Public Const ERROR_CLUSTER_NODE_ALREADY_UP    As Long = 5061    'The cluster node is already up.
Public Const ERROR_CLUSTER_NODE_ALREADY_DOWN  As Long = 5062    'The cluster node is already down.
Public Const ERROR_CLUSTER_NETWORK_ALREADY_ONLINE As Long = 5063 'The cluster network is already online.
Public Const ERROR_CLUSTER_NETWORK_ALREADY_OFFLINE As Long = 5064 'The cluster network is already offline.
Public Const ERROR_CLUSTER_NODE_ALREADY_MEMBER As Long = 5065    'The cluster node is already a member of the cluster.
Public Const ERROR_CLUSTER_LAST_INTERNAL_NETWORK As Long = 5066  'The cluster network is the only one configured for internal cluster communication between two or more active cluster nodes. The internal communication capability cannot be removed from the network.
Public Const ERROR_CLUSTER_NETWORK_HAS_DEPENDENTS As Long = 5067 'One or more cluster resources depend on the network to provide service to clients. The client access capability cannot be removed from the network.
Public Const ERROR_INVALID_OPERATION_ON_QUORUM As Long = 5068   'This operation cannot be performed on the cluster resource as it the quorum resource. You may not bring the quorum resource offline or modify its possible owners list.
Public Const ERROR_DEPENDENCY_NOT_ALLOWED     As Long = 5069    'The cluster quorum resource is not allowed to have any dependencies.
Public Const ERROR_CLUSTER_NODE_PAUSED        As Long = 5070    'The cluster node is paused.
Public Const ERROR_NODE_CANT_HOST_RESOURCE    As Long = 5071    'The cluster resource cannot be brought online. The owner node cannot run this resource.
Public Const ERROR_CLUSTER_NODE_NOT_READY     As Long = 5072    'The cluster node is not ready to perform the requested operation.
Public Const ERROR_CLUSTER_NODE_SHUTTING_DOWN As Long = 5073    'The cluster node is shutting down.
Public Const ERROR_CLUSTER_JOIN_ABORTED       As Long = 5074    'The cluster join operation was aborted.
Public Const ERROR_CLUSTER_INCOMPATIBLE_VERSIONS As Long = 5075 'The cluster join operation failed due to incompatible software versions between the joining node and its sponsor.
Public Const ERROR_CLUSTER_MAXNUM_OF_RESOURCES_EXCEEDED As Long = 5076    'This resource cannot be created because the cluster has reached the limit on the number of resources it can monitor.
Public Const ERROR_CLUSTER_SYSTEM_CONFIG_CHANGED As Long = 5077 'The system configuration changed during the cluster join or form operation. The join or form operation was aborted.
Public Const ERROR_CLUSTER_RESOURCE_TYPE_NOT_FOUND As Long = 5078 'The specified resource type was not found.
Public Const ERROR_CLUSTER_RESTYPE_NOT_SUPPORTED As Long = 5079 'The specified node does not support a resource of this type.  This may be due to version inconsistencies or due to the absence of the resource DLL on this node.
Public Const ERROR_CLUSTER_RESNAME_NOT_FOUND  As Long = 5080    'The specified resource name is not supported by this resource DLL. This may be due to a bad (or changed) name supplied to the resource DLL.
Public Const ERROR_CLUSTER_NO_RPC_PACKAGES_REGISTERED As Long = 5081    'No authentication package could be registered with the RPC server.
Public Const ERROR_CLUSTER_OWNER_NOT_IN_PREFLIST As Long = 5082 'You cannot bring the group online because the owner of the group is not in the preferred list for the group. To change the owner node for the group, move the group.
Public Const ERROR_CLUSTER_DATABASE_SEQMISMATCH As Long = 5083  'The join operation failed because the cluster database sequence number has changed or is incompatible with the locker node. This may happen during a join operation if the cluster database was changing during the join.
Public Const ERROR_RESMON_INVALID_STATE       As Long = 5084    'The resource monitor will not allow the fail operation to be performed while the resource is in its current state. This may happen if the resource is in a pending state.
Public Const ERROR_CLUSTER_GUM_NOT_LOCKER     As Long = 5085    'A non locker code got a request to reserve the lock for making global updates.
Public Const ERROR_QUORUM_DISK_NOT_FOUND      As Long = 5086    'The quorum disk could not be located by the cluster service.
Public Const ERROR_DATABASE_BACKUP_CORRUPT    As Long = 5087    'The backed up cluster database is possibly corrupt.
Public Const ERROR_CLUSTER_NODE_ALREADY_HAS_DFS_ROOT As Long = 5088    'A DFS root already exists in this cluster node.
Public Const ERROR_RESOURCE_PROPERTY_UNCHANGEABLE As Long = 5089 'An attempt to modify a resource property failed because it conflicts with another existing property.
Public Const ERROR_CLUSTER_MEMBERSHIP_INVALID_STATE As Long = 5890    'An operation was attempted that is incompatible with the current membership state of the node.
Public Const ERROR_CLUSTER_QUORUMLOG_NOT_FOUND As Long = 5891   'The quorum resource does not contain the quorum log.
Public Const ERROR_CLUSTER_MEMBERSHIP_HALT    As Long = 5892    'The membership engine requested shutdown of the cluster service on this node.
Public Const ERROR_CLUSTER_INSTANCE_ID_MISMATCH As Long = 5893  'The join operation failed because the cluster instance ID of the joining node does not match the cluster instance ID of the sponsor node.
Public Const ERROR_CLUSTER_NETWORK_NOT_FOUND_FOR_IP As Long = 5894    'A matching network for the specified IP address could not be found. Please also specify a subnet mask and a cluster network.
Public Const ERROR_CLUSTER_PROPERTY_DATA_TYPE_MISMATCH As Long = 5895    'The actual data type of the property did not match the expected data type of the property.
Public Const ERROR_CLUSTER_EVICT_WITHOUT_CLEANUP As Long = 5896 'The cluster node was evicted from the cluster successfully, but the node was not cleaned up.  Extended status information explaining why the node was not cleaned up is available.
Public Const ERROR_CLUSTER_PARAMETER_MISMATCH As Long = 5897    'Two or more parameter values specified for a resource's properties are in conflict.
Public Const ERROR_NODE_CANNOT_BE_CLUSTERED   As Long = 5898    'This computer cannot be made a member of a cluster.
Public Const ERROR_CLUSTER_WRONG_OS_VERSION   As Long = 5899    'This computer cannot be made a member of a cluster because it does not have the correct version of Windows installed.
Public Const ERROR_CLUSTER_CANT_CREATE_DUP_CLUSTER_NAME As Long = 5900    'A cluster cannot be created with the specified cluster name because that cluster name is already in use. Specify a different name for the cluster.
Public Const ERROR_CLUSCFG_ALREADY_COMMITTED  As Long = 5901    'The cluster configuration action has already been committed.
Public Const ERROR_CLUSCFG_ROLLBACK_FAILED    As Long = 5902    'The cluster configuration action could not be rolled back.
Public Const ERROR_CLUSCFG_SYSTEM_DISK_DRIVE_LETTER_CONFLICT As Long = 5903    'The drive letter assigned to a system disk on one node conflicted with the drive letter assigned to a disk on another node.
Public Const ERROR_CLUSTER_OLD_VERSION        As Long = 5904    'One or more nodes in the cluster are running a version of Windows that does not support this operation.
Public Const ERROR_CLUSTER_MISMATCHED_COMPUTER_ACCT_NAME As Long = 5905    'The name of the corresponding computer account doesn't match the Network Name for this resource.

'
'////////////////////////////////////
'//                                //
'//     EFS Error Codes            //
'//                                //
'////////////////////////////////////
'
Public Const ERROR_ENCRYPTION_FAILED          As Long = 6000    'The specified file could not be encrypted.
Public Const ERROR_DECRYPTION_FAILED          As Long = 6001    'The specified file could not be decrypted.
Public Const ERROR_FILE_ENCRYPTED             As Long = 6002    'The specified file is encrypted and the user does not have the ability to decrypt it.
Public Const ERROR_NO_RECOVERY_POLICY         As Long = 6003    'There is no valid encryption recovery policy configured for this system.
Public Const ERROR_NO_EFS                     As Long = 6004    'The required encryption driver is not loaded for this system.
Public Const ERROR_WRONG_EFS                  As Long = 6005    'The file was encrypted with a different encryption driver than is currently loaded.
Public Const ERROR_NO_USER_KEYS               As Long = 6006    'There are no EFS keys defined for the user.
Public Const ERROR_FILE_NOT_ENCRYPTED         As Long = 6007    'The specified file is not encrypted.
Public Const ERROR_NOT_EXPORT_FORMAT          As Long = 6008    'The specified file is not in the defined EFS export format.
Public Const ERROR_FILE_READ_ONLY             As Long = 6009    'The specified file is read only.
Public Const ERROR_DIR_EFS_DISALLOWED         As Long = 6010    'The directory has been disabled for encryption.
Public Const ERROR_EFS_SERVER_NOT_TRUSTED     As Long = 6011    'The server is not trusted for remote encryption operation.
Public Const ERROR_BAD_RECOVERY_POLICY        As Long = 6012    'Recovery policy configured for this system contains invalid recovery certificate.
Public Const ERROR_EFS_ALG_BLOB_TOO_BIG       As Long = 6013    'The encryption algorithm used on the source file needs a bigger key buffer than the one on the destination file.
Public Const ERROR_VOLUME_NOT_SUPPORT_EFS     As Long = 6014    'The disk partition does not support file encryption.
Public Const ERROR_EFS_DISABLED               As Long = 6015    'This machine is disabled for file encryption.
Public Const ERROR_EFS_VERSION_NOT_SUPPORT    As Long = 6016    'A newer system is required to decrypt this encrypted file.
Public Const ERROR_NO_BROWSER_SERVERS_FOUND   As Long = 6118    'The list of servers for this workgroup is not currently available

'
'//////////////////////////////////////////////////////////////////
'//                                                              //
'// Task Scheduler Error Codes that NET START must understand    //
'//                                                              //
'//////////////////////////////////////////////////////////////////
'
'Public Const SCHED_E_SERVICE_NOT_LOCALSYSTEM  6200L'The Task Scheduler service must be configured to run in the System account to function properly.  Individual tasks may be configured to run in other accounts.
'
'////////////////////////////////////
'//                                //
'// Terminal Server Error Codes    //
'//                                //
'////////////////////////////////////
Public Const ERROR_CTX_WINSTATION_NAME_INVALID As Long = 7001   'The specified session name is invalid.
Public Const ERROR_CTX_INVALID_PD             As Long = 7002    'The specified protocol driver is invalid.
Public Const ERROR_CTX_PD_NOT_FOUND           As Long = 7003    'The specified protocol driver was not found in the system path.
Public Const ERROR_CTX_WD_NOT_FOUND           As Long = 7004    'The specified terminal connection driver was not found in the system path.
Public Const ERROR_CTX_CANNOT_MAKE_EVENTLOG_ENTRY As Long = 7005 'A registry key for event logging could not be created for this session.
Public Const ERROR_CTX_SERVICE_NAME_COLLISION As Long = 7006    'A service with the same name already exists on the system.
Public Const ERROR_CTX_CLOSE_PENDING          As Long = 7007    'A close operation is pending on the session.
Public Const ERROR_CTX_NO_OUTBUF              As Long = 7008    'There are no free output buffers available.
Public Const ERROR_CTX_MODEM_INF_NOT_FOUND    As Long = 7009    'The MODEM.INF file was not found.
Public Const ERROR_CTX_INVALID_MODEMNAME      As Long = 7010    'The modem name was not found in MODEM.INF.
Public Const ERROR_CTX_MODEM_RESPONSE_ERROR   As Long = 7011    'The modem did not accept the command sent to it. Verify that the configured modem name matches the attached modem.
Public Const ERROR_CTX_MODEM_RESPONSE_TIMEOUT As Long = 7012    'The modem did not respond to the command sent to it. Verify that the modem is properly cabled and powered on.
Public Const ERROR_CTX_MODEM_RESPONSE_NO_CARRIER As Long = 7013 'Carrier detect has failed or carrier has been dropped due to disconnect.
Public Const ERROR_CTX_MODEM_RESPONSE_NO_DIALTONE As Long = 7014 'Dial tone not detected within the required time. Verify that the phone cable is properly attached and functional.
Public Const ERROR_CTX_MODEM_RESPONSE_BUSY    As Long = 7015    'Busy signal detected at remote site on callback.
Public Const ERROR_CTX_MODEM_RESPONSE_VOICE   As Long = 7016    'Voice detected at remote site on callback.
Public Const ERROR_CTX_TD_ERROR               As Long = 7017    'Transport driver error
Public Const ERROR_CTX_WINSTATION_NOT_FOUND   As Long = 7022    'The specified session cannot be found.
Public Const ERROR_CTX_WINSTATION_ALREADY_EXISTS As Long = 7023 'The specified session name is already in use.
Public Const ERROR_CTX_WINSTATION_BUSY        As Long = 7024    'The requested operation cannot be completed because the terminal connection is currently busy processing a connect, disconnect, reset, or delete operation.
Public Const ERROR_CTX_BAD_VIDEO_MODE         As Long = 7025    'An attempt has been made to connect to a session whose video mode is not supported by the current client.
Public Const ERROR_CTX_GRAPHICS_INVALID       As Long = 7035    'The application attempted to enable DOS graphics mode.    'DOS graphics mode is not supported.
Public Const ERROR_CTX_LOGON_DISABLED         As Long = 7037    'Your interactive logon privilege has been disabled.    'Please contact your administrator.
Public Const ERROR_CTX_NOT_CONSOLE            As Long = 7038    'The requested operation can be performed only on the system console.    'This is most often the result of a driver or system DLL requiring direct console access.
Public Const ERROR_CTX_CLIENT_QUERY_TIMEOUT   As Long = 7040    'The client failed to respond to the server connect message.
Public Const ERROR_CTX_CONSOLE_DISCONNECT     As Long = 7041    'Disconnecting the console session is not supported.
Public Const ERROR_CTX_CONSOLE_CONNECT        As Long = 7042    'Reconnecting a disconnected session to the console is not supported.
Public Const ERROR_CTX_SHADOW_DENIED          As Long = 7044    'The request to control another session remotely was denied.
Public Const ERROR_CTX_WINSTATION_ACCESS_DENIED As Long = 7045  'The requested session access is denied.
Public Const ERROR_CTX_INVALID_WD             As Long = 7049    'The specified terminal connection driver is invalid.
Public Const ERROR_CTX_SHADOW_INVALID         As Long = 7050    'The requested session cannot be controlled remotely.    'This may be because the session is disconnected or does not currently have a user logged on.
Public Const ERROR_CTX_SHADOW_DISABLED        As Long = 7051    'The requested session is not configured to allow remote control.
Public Const ERROR_CTX_CLIENT_LICENSE_IN_USE  As Long = 7052    'Your request to connect to this Terminal Server has been rejected. Your Terminal Server client license number is currently being used by another user.    'Please call your system administrator to obtain a unique license number.
Public Const ERROR_CTX_CLIENT_LICENSE_NOT_SET As Long = 7053    'Your request to connect to this Terminal Server has been rejected. Your Terminal Server client license number has not been entered for this copy of the Terminal Server client.    'Please contact your system administrator.
Public Const ERROR_CTX_LICENSE_NOT_AVAILABLE  As Long = 7054    'The system has reached its licensed logon limit.    'Please try again later.
Public Const ERROR_CTX_LICENSE_CLIENT_INVALID As Long = 7055    'The client you are using is not licensed to use this system.  Your logon request is denied.
Public Const ERROR_CTX_LICENSE_EXPIRED        As Long = 7056    'The system license has expired.  Your logon request is denied.
Public Const ERROR_CTX_SHADOW_NOT_RUNNING     As Long = 7057    'Remote control could not be terminated because the specified session is not currently being remotely controlled.
Public Const ERROR_CTX_SHADOW_ENDED_BY_MODE_CHANGE As Long = 7058 'The remote control of the console was terminated because the display mode was changed. Changing the display mode in a remote control session is not supported.
Public Const ERROR_ACTIVATION_COUNT_EXCEEDED  As Long = 7059    'Activation has already been reset the maximum number of times for this installation. Your activation timer will not be cleared.

'
'///////////////////////////////////////////////////
'//                                                /
'//             Traffic Control Error Codes        /
'//                                                /
'//                  7500 to  7999                 /
'//                                                /
'//         defined in: tcerror.h                  /
'///////////////////////////////////////////////////
'///////////////////////////////////////////////////
'//                                                /
'//             Active Directory Error Codes       /
'//                                                /
'//                  8000 to  8999                 /
'///////////////////////////////////////////////////
'// *****************
'// FACILITY_FILE_REPLICATION_SERVICE
'// *****************
Public Const FRS_ERR_INVALID_API_SEQUENCE     As Long = 8001    'The file replication service API was called incorrectly.
Public Const FRS_ERR_STARTING_SERVICE         As Long = 8002    'The file replication service cannot be started.
Public Const FRS_ERR_STOPPING_SERVICE         As Long = 8003    'The file replication service cannot be stopped.
Public Const FRS_ERR_INTERNAL_API             As Long = 8004    'The file replication service API terminated the request.    'The event log may have more information.
Public Const FRS_ERR_INTERNAL                 As Long = 8005    'The file replication service terminated the request.    'The event log may have more information.
Public Const FRS_ERR_SERVICE_COMM             As Long = 8006    'The file replication service cannot be contacted.    'The event log may have more information.
Public Const FRS_ERR_INSUFFICIENT_PRIV        As Long = 8007    'The file replication service cannot satisfy the request because the user has insufficient privileges.    'The event log may have more information.
Public Const FRS_ERR_AUTHENTICATION           As Long = 8008    'The file replication service cannot satisfy the request because authenticated RPC is not available.    'The event log may have more information.
Public Const FRS_ERR_PARENT_INSUFFICIENT_PRIV As Long = 8009    'The file replication service cannot satisfy the request because the user has insufficient privileges on the domain controller.    'The event log may have more information.
Public Const FRS_ERR_PARENT_AUTHENTICATION    As Long = 8010    'The file replication service cannot satisfy the request because authenticated RPC is not available on the domain controller.    'The event log may have more information.
Public Const FRS_ERR_CHILD_TO_PARENT_COMM     As Long = 8011    'The file replication service cannot communicate with the file replication service on the domain controller.    'The event log may have more information.
Public Const FRS_ERR_PARENT_TO_CHILD_COMM     As Long = 8012    'The file replication service on the domain controller cannot communicate with the file replication service on this computer.    'The event log may have more information.
Public Const FRS_ERR_SYSVOL_POPULATE          As Long = 8013    'The file replication service cannot populate the system volume because of an internal error.    'The event log may have more information.
Public Const FRS_ERR_SYSVOL_POPULATE_TIMEOUT  As Long = 8014    'The file replication service cannot populate the system volume because of an internal timeout.    'The event log may have more information.
Public Const FRS_ERR_SYSVOL_IS_BUSY           As Long = 8015    'The file replication service cannot process the request. The system volume is busy with a previous request.
Public Const FRS_ERR_SYSVOL_DEMOTE            As Long = 8016    'The file replication service cannot stop replicating the system volume because of an internal error.    'The event log may have more information.
Public Const FRS_ERR_INVALID_SERVICE_PARAMETER As Long = 8017   'The file replication service detected an invalid parameter.
Public Const ERROR_DS_NOT_INSTALLED           As Long = 8200    'An error occurred while installing the directory service. For more information, see the event log.
Public Const ERROR_DS_MEMBERSHIP_EVALUATED_LOCALLY As Long = 8201 'The directory service evaluated group memberships locally.
Public Const ERROR_DS_NO_ATTRIBUTE_OR_VALUE   As Long = 8202    'The specified directory service attribute or value does not exist.
Public Const ERROR_DS_INVALID_ATTRIBUTE_SYNTAX As Long = 8203   'The attribute syntax specified to the directory service is invalid.
Public Const ERROR_DS_ATTRIBUTE_TYPE_UNDEFINED As Long = 8204   'The attribute type specified to the directory service is not defined.
Public Const ERROR_DS_ATTRIBUTE_OR_VALUE_EXISTS As Long = 8205  'The specified directory service attribute or value already exists.
Public Const ERROR_DS_BUSY                    As Long = 8206    'The directory service is busy.
Public Const ERROR_DS_UNAVAILABLE             As Long = 8207    'The directory service is unavailable.
Public Const ERROR_DS_NO_RIDS_ALLOCATED       As Long = 8208    'The directory service was unable to allocate a relative identifier.
Public Const ERROR_DS_NO_MORE_RIDS            As Long = 8209    'The directory service has exhausted the pool of relative identifiers.
Public Const ERROR_DS_INCORRECT_ROLE_OWNER    As Long = 8210    'The requested operation could not be performed because the directory service is not the master for that type of operation.
Public Const ERROR_DS_RIDMGR_INIT_ERROR       As Long = 8211    'The directory service was unable to initialize the subsystem that allocates relative identifiers.
Public Const ERROR_DS_OBJ_CLASS_VIOLATION     As Long = 8212    'The requested operation did not satisfy one or more constraints associated with the class of the object.
Public Const ERROR_DS_CANT_ON_NON_LEAF        As Long = 8213    'The directory service can perform the requested operation only on a leaf object.
Public Const ERROR_DS_CANT_ON_RDN             As Long = 8214    'The directory service cannot perform the requested operation on the RDN attribute of an object.
Public Const ERROR_DS_CANT_MOD_OBJ_CLASS      As Long = 8215    'The directory service detected an attempt to modify the object class of an object.
Public Const ERROR_DS_CROSS_DOM_MOVE_ERROR    As Long = 8216    'The requested cross-domain move operation could not be performed.
Public Const ERROR_DS_GC_NOT_AVAILABLE        As Long = 8217    'Unable to contact the global catalog server.
Public Const ERROR_SHARED_POLICY              As Long = 8218    'The policy object is shared and can only be modified at the root.
Public Const ERROR_POLICY_OBJECT_NOT_FOUND    As Long = 8219    'The policy object does not exist.
Public Const ERROR_POLICY_ONLY_IN_DS          As Long = 8220    'The requested policy information is only in the directory service.
Public Const ERROR_PROMOTION_ACTIVE           As Long = 8221    'A domain controller promotion is currently active.
Public Const ERROR_NO_PROMOTION_ACTIVE        As Long = 8222    'A domain controller promotion is not currently active
Public Const ERROR_DS_OPERATIONS_ERROR        As Long = 8224    'An operations error occurred.
Public Const ERROR_DS_PROTOCOL_ERROR          As Long = 8225    'A protocol error occurred.
Public Const ERROR_DS_TIMELIMIT_EXCEEDED      As Long = 8226    'The time limit for this request was exceeded.
Public Const ERROR_DS_SIZELIMIT_EXCEEDED      As Long = 8227    'The size limit for this request was exceeded.
Public Const ERROR_DS_ADMIN_LIMIT_EXCEEDED    As Long = 8228    'The administrative limit for this request was exceeded.
Public Const ERROR_DS_COMPARE_FALSE           As Long = 8229    'The compare response was false.
Public Const ERROR_DS_COMPARE_TRUE            As Long = 8230    'The compare response was true.
Public Const ERROR_DS_AUTH_METHOD_NOT_SUPPORTED As Long = 8231  'The requested authentication method is not supported by the server.
Public Const ERROR_DS_STRONG_AUTH_REQUIRED    As Long = 8232    'A more secure authentication method is required for this server.
Public Const ERROR_DS_INAPPROPRIATE_AUTH      As Long = 8233    'Inappropriate authentication.
Public Const ERROR_DS_AUTH_UNKNOWN            As Long = 8234    'The authentication mechanism is unknown.
Public Const ERROR_DS_REFERRAL                As Long = 8235    'A referral was returned from the server.
Public Const ERROR_DS_UNAVAILABLE_CRIT_EXTENSION As Long = 8236 'The server does not support the requested critical extension.
Public Const ERROR_DS_CONFIDENTIALITY_REQUIRED As Long = 8237   'This request requires a secure connection.
Public Const ERROR_DS_INAPPROPRIATE_MATCHING  As Long = 8238    'Inappropriate matching.
Public Const ERROR_DS_CONSTRAINT_VIOLATION    As Long = 8239    'A constraint violation occurred.
Public Const ERROR_DS_NO_SUCH_OBJECT          As Long = 8240    'There is no such object on the server.
Public Const ERROR_DS_ALIAS_PROBLEM           As Long = 8241    'There is an alias problem.
Public Const ERROR_DS_INVALID_DN_SYNTAX       As Long = 8242    'An invalid dn syntax has been specified.
Public Const ERROR_DS_IS_LEAF                 As Long = 8243    'The object is a leaf object.
Public Const ERROR_DS_ALIAS_DEREF_PROBLEM     As Long = 8244    'There is an alias dereferencing problem.
Public Const ERROR_DS_UNWILLING_TO_PERFORM    As Long = 8245    'The server is unwilling to process the request.
Public Const ERROR_DS_LOOP_DETECT             As Long = 8246    'A loop has been detected.
Public Const ERROR_DS_NAMING_VIOLATION        As Long = 8247    'There is a naming violation.
Public Const ERROR_DS_OBJECT_RESULTS_TOO_LARGE As Long = 8248   'The result set is too large.
Public Const ERROR_DS_AFFECTS_MULTIPLE_DSAS   As Long = 8249    'The operation affects multiple DSAs
Public Const ERROR_DS_SERVER_DOWN             As Long = 8250    'The server is not operational.
Public Const ERROR_DS_LOCAL_ERROR             As Long = 8251    'A local error has occurred.
Public Const ERROR_DS_ENCODING_ERROR          As Long = 8252    'An encoding error has occurred.
Public Const ERROR_DS_DECODING_ERROR          As Long = 8253    'A decoding error has occurred.
Public Const ERROR_DS_FILTER_UNKNOWN          As Long = 8254    'The search filter cannot be recognized.
Public Const ERROR_DS_PARAM_ERROR             As Long = 8255    'One or more parameters are illegal.
Public Const ERROR_DS_NOT_SUPPORTED           As Long = 8256    'The specified method is not supported.
Public Const ERROR_DS_NO_RESULTS_RETURNED     As Long = 8257    'No results were returned.
Public Const ERROR_DS_CONTROL_NOT_FOUND       As Long = 8258    'The specified control is not supported by the server.
Public Const ERROR_DS_CLIENT_LOOP             As Long = 8259    'A referral loop was detected by the client.
Public Const ERROR_DS_REFERRAL_LIMIT_EXCEEDED As Long = 8260    'The preset referral limit was exceeded.
Public Const ERROR_DS_SORT_CONTROL_MISSING    As Long = 8261    'The search requires a SORT control.
Public Const ERROR_DS_OFFSET_RANGE_ERROR      As Long = 8262    'The search results exceed the offset range specified.
Public Const ERROR_DS_ROOT_MUST_BE_NC         As Long = 8301    'The root object must be the head of a naming context. The root object cannot have an instantiated parent.
Public Const ERROR_DS_ADD_REPLICA_INHIBITED   As Long = 8302    'The add replica operation cannot be performed. The naming context must be writeable in order to create the replica.
Public Const ERROR_DS_ATT_NOT_DEF_IN_SCHEMA   As Long = 8303    'A reference to an attribute that is not defined in the schema occurred.
Public Const ERROR_DS_MAX_OBJ_SIZE_EXCEEDED   As Long = 8304    'The maximum size of an object has been exceeded.
Public Const ERROR_DS_OBJ_STRING_NAME_EXISTS  As Long = 8305    'An attempt was made to add an object to the directory with a name that is already in use.
Public Const ERROR_DS_NO_RDN_DEFINED_IN_SCHEMA As Long = 8306   'An attempt was made to add an object of a class that does not have an RDN defined in the schema.
Public Const ERROR_DS_RDN_DOESNT_MATCH_SCHEMA As Long = 8307    'An attempt was made to add an object using an RDN that is not the RDN defined in the schema.
Public Const ERROR_DS_NO_REQUESTED_ATTS_FOUND As Long = 8308    'None of the requested attributes were found on the objects.
Public Const ERROR_DS_USER_BUFFER_TO_SMALL    As Long = 8309    'The user buffer is too small.
Public Const ERROR_DS_ATT_IS_NOT_ON_OBJ       As Long = 8310    'The attribute specified in the operation is not present on the object.
Public Const ERROR_DS_ILLEGAL_MOD_OPERATION   As Long = 8311    'Illegal modify operation. Some aspect of the modification is not permitted.
Public Const ERROR_DS_OBJ_TOO_LARGE           As Long = 8312    'The specified object is too large.
Public Const ERROR_DS_BAD_INSTANCE_TYPE       As Long = 8313    'The specified instance type is not valid.
Public Const ERROR_DS_MASTERDSA_REQUIRED      As Long = 8314    'The operation must be performed at a master DSA.
Public Const ERROR_DS_OBJECT_CLASS_REQUIRED   As Long = 8315    'The object class attribute must be specified.
Public Const ERROR_DS_MISSING_REQUIRED_ATT    As Long = 8316    'A required attribute is missing.
Public Const ERROR_DS_ATT_NOT_DEF_FOR_CLASS   As Long = 8317    'An attempt was made to modify an object to include an attribute that is not legal for its class.
Public Const ERROR_DS_ATT_ALREADY_EXISTS      As Long = 8318    'The specified attribute is already present on the object.
Public Const ERROR_DS_CANT_ADD_ATT_VALUES     As Long = 8320    'The specified attribute is not present, or has no values.
Public Const ERROR_DS_SINGLE_VALUE_CONSTRAINT As Long = 8321    'Multiple values were specified for an attribute that can have only one value.
Public Const ERROR_DS_RANGE_CONSTRAINT        As Long = 8322    'A value for the attribute was not in the acceptable range of values.
Public Const ERROR_DS_ATT_VAL_ALREADY_EXISTS  As Long = 8323    'The specified value already exists.
Public Const ERROR_DS_CANT_REM_MISSING_ATT    As Long = 8324    'The attribute cannot be removed because it is not present on the object.
Public Const ERROR_DS_CANT_REM_MISSING_ATT_VAL As Long = 8325   'The attribute value cannot be removed because it is not present on the object.
Public Const ERROR_DS_ROOT_CANT_BE_SUBREF     As Long = 8326    'The specified root object cannot be a subref.
Public Const ERROR_DS_NO_CHAINING             As Long = 8327    'Chaining is not permitted.
Public Const ERROR_DS_NO_CHAINED_EVAL         As Long = 8328    'Chained evaluation is not permitted.
Public Const ERROR_DS_NO_PARENT_OBJECT        As Long = 8329    'The operation could not be performed because the object's parent is either uninstantiated or deleted.
Public Const ERROR_DS_PARENT_IS_AN_ALIAS      As Long = 8330    'Having a parent that is an alias is not permitted. Aliases are leaf objects.
Public Const ERROR_DS_CANT_MIX_MASTER_AND_REPS As Long = 8331   'The object and parent must be of the same type, either both masters or both replicas.
Public Const ERROR_DS_CHILDREN_EXIST          As Long = 8332    'The operation cannot be performed because child objects exist. This operation can only be performed on a leaf object.
Public Const ERROR_DS_OBJ_NOT_FOUND           As Long = 8333    'Directory object not found.
Public Const ERROR_DS_ALIASED_OBJ_MISSING     As Long = 8334    'The aliased object is missing.
Public Const ERROR_DS_BAD_NAME_SYNTAX         As Long = 8335    'The object name has bad syntax.
Public Const ERROR_DS_ALIAS_POINTS_TO_ALIAS   As Long = 8336    'It is not permitted for an alias to refer to another alias.
Public Const ERROR_DS_CANT_DEREF_ALIAS        As Long = 8337    'The alias cannot be dereferenced.
Public Const ERROR_DS_OUT_OF_SCOPE            As Long = 8338    'The operation is out of scope.
Public Const ERROR_DS_OBJECT_BEING_REMOVED    As Long = 8339    'The operation cannot continue because the object is in the process of being removed.
Public Const ERROR_DS_CANT_DELETE_DSA_OBJ     As Long = 8340    'The DSA object cannot be deleted.
Public Const ERROR_DS_GENERIC_ERROR           As Long = 8341    'A directory service error has occurred.
Public Const ERROR_DS_DSA_MUST_BE_INT_MASTER  As Long = 8342    'The operation can only be performed on an internal master DSA object.
Public Const ERROR_DS_CLASS_NOT_DSA           As Long = 8343    'The object must be of class DSA.
Public Const ERROR_DS_INSUFF_ACCESS_RIGHTS    As Long = 8344    'Insufficient access rights to perform the operation.
Public Const ERROR_DS_ILLEGAL_SUPERIOR        As Long = 8345    'The object cannot be added because the parent is not on the list of possible superiors.
Public Const ERROR_DS_ATTRIBUTE_OWNED_BY_SAM  As Long = 8346    'Access to the attribute is not permitted because the attribute is owned by the Security Accounts Manager (SAM).
Public Const ERROR_DS_NAME_TOO_MANY_PARTS     As Long = 8347    'The name has too many parts.
Public Const ERROR_DS_NAME_TOO_LONG           As Long = 8348    'The name is too long.
Public Const ERROR_DS_NAME_VALUE_TOO_LONG     As Long = 8349    'The name value is too long.
Public Const ERROR_DS_NAME_UNPARSEABLE        As Long = 8350    'The directory service encountered an error parsing a name.
Public Const ERROR_DS_NAME_TYPE_UNKNOWN       As Long = 8351    'The directory service cannot get the attribute type for a name.
Public Const ERROR_DS_NOT_AN_OBJECT           As Long = 8352    'The name does not identify an object; the name identifies a phantom.
Public Const ERROR_DS_SEC_DESC_TOO_SHORT      As Long = 8353    'The security descriptor is too short.
Public Const ERROR_DS_SEC_DESC_INVALID        As Long = 8354    'The security descriptor is invalid.
Public Const ERROR_DS_NO_DELETED_NAME         As Long = 8355    'Failed to create name for deleted object.
Public Const ERROR_DS_SUBREF_MUST_HAVE_PARENT As Long = 8356    'The parent of a new subref must exist.
Public Const ERROR_DS_NCNAME_MUST_BE_NC       As Long = 8357    'The object must be a naming context.
Public Const ERROR_DS_CANT_ADD_SYSTEM_ONLY    As Long = 8358    'It is not permitted to add an attribute which is owned by the system.
Public Const ERROR_DS_CLASS_MUST_BE_CONCRETE  As Long = 8359    'The class of the object must be structural; you cannot instantiate an abstract class.
Public Const ERROR_DS_INVALID_DMD             As Long = 8360    'The schema object could not be found.
Public Const ERROR_DS_OBJ_GUID_EXISTS         As Long = 8361    'A local object with this Guid (dead or alive) already exists.
Public Const ERROR_DS_NOT_ON_BACKLINK         As Long = 8362    'The operation cannot be performed on a back link.
Public Const ERROR_DS_NO_CROSSREF_FOR_NC      As Long = 8363    'The cross reference for the specified naming context could not be found.
Public Const ERROR_DS_SHUTTING_DOWN           As Long = 8364    'The operation could not be performed because the directory service is shutting down.
Public Const ERROR_DS_UNKNOWN_OPERATION       As Long = 8365    'The directory service request is invalid.
Public Const ERROR_DS_INVALID_ROLE_OWNER      As Long = 8366    'The role owner attribute could not be read.
Public Const ERROR_DS_COULDNT_CONTACT_FSMO    As Long = 8367    'The requested FSMO operation failed. The current FSMO holder could not be contacted.
Public Const ERROR_DS_CROSS_NC_DN_RENAME      As Long = 8368    'Modification of a DN across a naming context is not permitted.
Public Const ERROR_DS_CANT_MOD_SYSTEM_ONLY    As Long = 8369    'The attribute cannot be modified because it is owned by the system.
Public Const ERROR_DS_REPLICATOR_ONLY         As Long = 8370    'Only the replicator can perform this function.
Public Const ERROR_DS_OBJ_CLASS_NOT_DEFINED   As Long = 8371    'The specified class is not defined.
Public Const ERROR_DS_OBJ_CLASS_NOT_SUBCLASS  As Long = 8372    'The specified class is not a subclass.
Public Const ERROR_DS_NAME_REFERENCE_INVALID  As Long = 8373    'The name reference is invalid.
Public Const ERROR_DS_CROSS_REF_EXISTS        As Long = 8374    'A cross reference already exists.
Public Const ERROR_DS_CANT_DEL_MASTER_CROSSREF As Long = 8375   'It is not permitted to delete a master cross reference.
Public Const ERROR_DS_SUBTREE_NOTIFY_NOT_NC_HEAD As Long = 8376 'Subtree notifications are only supported on NC heads.
Public Const ERROR_DS_NOTIFY_FILTER_TOO_COMPLEX As Long = 8377  'Notification filter is too complex.
Public Const ERROR_DS_DUP_RDN                 As Long = 8378    'Schema update failed: duplicate RDN.
Public Const ERROR_DS_DUP_OID                 As Long = 8379    'Schema update failed: duplicate OID.
Public Const ERROR_DS_DUP_MAPI_ID             As Long = 8380    'Schema update failed: duplicate MAPI identifier.
Public Const ERROR_DS_DUP_SCHEMA_ID_GUID      As Long = 8381    'Schema update failed: duplicate schema-id Guid.
Public Const ERROR_DS_DUP_LDAP_DISPLAY_NAME   As Long = 8382    'Schema update failed: duplicate LDAP display name.
Public Const ERROR_DS_SEMANTIC_ATT_TEST       As Long = 8383    'Schema update failed: range-lower less than range upper.
Public Const ERROR_DS_SYNTAX_MISMATCH         As Long = 8384    'Schema update failed: syntax mismatch.
Public Const ERROR_DS_EXISTS_IN_MUST_HAVE     As Long = 8385    'Schema deletion failed: attribute is used in must-contain.
Public Const ERROR_DS_EXISTS_IN_MAY_HAVE      As Long = 8386    'Schema deletion failed: attribute is used in may-contain.
Public Const ERROR_DS_NONEXISTENT_MAY_HAVE    As Long = 8387    'Schema update failed: attribute in may-contain does not exist.
Public Const ERROR_DS_NONEXISTENT_MUST_HAVE   As Long = 8388    'Schema update failed: attribute in must-contain does not exist.
Public Const ERROR_DS_AUX_CLS_TEST_FAIL       As Long = 8389    'Schema update failed: class in aux-class list does not exist or is not an auxiliary class.
Public Const ERROR_DS_NONEXISTENT_POSS_SUP    As Long = 8390    'Schema update failed: class in poss-superiors does not exist.
Public Const ERROR_DS_SUB_CLS_TEST_FAIL       As Long = 8391    'Schema update failed: class in subclassof list does not exist or does not satisfy hierarchy rules.
Public Const ERROR_DS_BAD_RDN_ATT_ID_SYNTAX   As Long = 8392    'Schema update failed: Rdn-Att-Id has wrong syntax.
Public Const ERROR_DS_EXISTS_IN_AUX_CLS       As Long = 8393    'Schema deletion failed: class is used as auxiliary class.
Public Const ERROR_DS_EXISTS_IN_SUB_CLS       As Long = 8394    'Schema deletion failed: class is used as sub class.
Public Const ERROR_DS_EXISTS_IN_POSS_SUP      As Long = 8395    'Schema deletion failed: class is used as poss superior.
Public Const ERROR_DS_RECALCSCHEMA_FAILED     As Long = 8396    'Schema update failed in recalculating validation cache.
Public Const ERROR_DS_TREE_DELETE_NOT_FINISHED As Long = 8397   'The tree deletion is not finished.  The request must be made again to continue deleting the tree.
Public Const ERROR_DS_CANT_DELETE             As Long = 8398    'The requested delete operation could not be performed.
Public Const ERROR_DS_ATT_SCHEMA_REQ_ID       As Long = 8399    'Cannot read the governs class identifier for the schema record.
Public Const ERROR_DS_BAD_ATT_SCHEMA_SYNTAX   As Long = 8400    'The attribute schema has bad syntax.
Public Const ERROR_DS_CANT_CACHE_ATT          As Long = 8401    'The attribute could not be cached.
Public Const ERROR_DS_CANT_CACHE_CLASS        As Long = 8402    'The class could not be cached.
Public Const ERROR_DS_CANT_REMOVE_ATT_CACHE   As Long = 8403    'The attribute could not be removed from the cache.
Public Const ERROR_DS_CANT_REMOVE_CLASS_CACHE As Long = 8404    'The class could not be removed from the cache.
Public Const ERROR_DS_CANT_RETRIEVE_DN        As Long = 8405    'The distinguished name attribute could not be read.
Public Const ERROR_DS_MISSING_SUPREF          As Long = 8406    'No superior reference has been configured for the directory service. The directory service is therefore unable to issue referrals to objects outside this forest.
Public Const ERROR_DS_CANT_RETRIEVE_INSTANCE  As Long = 8407    'The instance type attribute could not be retrieved.
Public Const ERROR_DS_CODE_INCONSISTENCY      As Long = 8408    'An internal error has occurred.
Public Const ERROR_DS_DATABASE_ERROR          As Long = 8409    'A database error has occurred.
Public Const ERROR_DS_GOVERNSID_MISSING       As Long = 8410    'The attribute GOVERNSID is missing.
Public Const ERROR_DS_MISSING_EXPECTED_ATT    As Long = 8411    'An expected attribute is missing.
Public Const ERROR_DS_NCNAME_MISSING_CR_REF   As Long = 8412    'The specified naming context is missing a cross reference.
Public Const ERROR_DS_SECURITY_CHECKING_ERROR As Long = 8413    'A security checking error has occurred.
Public Const ERROR_DS_SCHEMA_NOT_LOADED       As Long = 8414    'The schema is not loaded.
Public Const ERROR_DS_SCHEMA_ALLOC_FAILED     As Long = 8415    'Schema allocation failed. Please check if the machine is running low on memory.
Public Const ERROR_DS_ATT_SCHEMA_REQ_SYNTAX   As Long = 8416    'Failed to obtain the required syntax for the attribute schema.
Public Const ERROR_DS_GCVERIFY_ERROR          As Long = 8417    'The global catalog verification failed. The global catalog is not available or does not support the operation. Some part of the directory is currently not available.
Public Const ERROR_DS_DRA_SCHEMA_MISMATCH     As Long = 8418    'The replication operation failed because of a schema mismatch between the servers involved.
Public Const ERROR_DS_CANT_FIND_DSA_OBJ       As Long = 8419    'The DSA object could not be found.
Public Const ERROR_DS_CANT_FIND_EXPECTED_NC   As Long = 8420    'The naming context could not be found.
Public Const ERROR_DS_CANT_FIND_NC_IN_CACHE   As Long = 8421    'The naming context could not be found in the cache.
Public Const ERROR_DS_CANT_RETRIEVE_CHILD     As Long = 8422    'The child object could not be retrieved.
Public Const ERROR_DS_SECURITY_ILLEGAL_MODIFY As Long = 8423    'The modification was not permitted for security reasons.
Public Const ERROR_DS_CANT_REPLACE_HIDDEN_REC As Long = 8424    'The operation cannot replace the hidden record.
Public Const ERROR_DS_BAD_HIERARCHY_FILE      As Long = 8425    'The hierarchy file is invalid.
Public Const ERROR_DS_BUILD_HIERARCHY_TABLE_FAILED As Long = 8426 'The attempt to build the hierarchy table failed.
Public Const ERROR_DS_CONFIG_PARAM_MISSING    As Long = 8427    'The directory configuration parameter is missing from the registry.
Public Const ERROR_DS_COUNTING_AB_INDICES_FAILED As Long = 8428 'The attempt to count the address book indices failed.
Public Const ERROR_DS_HIERARCHY_TABLE_MALLOC_FAILED As Long = 8429    'The allocation of the hierarchy table failed.
Public Const ERROR_DS_INTERNAL_FAILURE        As Long = 8430    'The directory service encountered an internal failure.
Public Const ERROR_DS_UNKNOWN_ERROR           As Long = 8431    'The directory service encountered an unknown failure.
Public Const ERROR_DS_ROOT_REQUIRES_CLASS_TOP As Long = 8432    'A root object requires a class of 'top'.
Public Const ERROR_DS_REFUSING_FSMO_ROLES     As Long = 8433    'This directory server is shutting down, and cannot take ownership of new floating single-master operation roles.
Public Const ERROR_DS_MISSING_FSMO_SETTINGS   As Long = 8434    'The directory service is missing mandatory configuration information, and is unable to determine the ownership of floating single-master operation roles.
Public Const ERROR_DS_UNABLE_TO_SURRENDER_ROLES As Long = 8435  'The directory service was unable to transfer ownership of one or more floating single-master operation roles to other servers.
Public Const ERROR_DS_DRA_GENERIC             As Long = 8436    'The replication operation failed.
Public Const ERROR_DS_DRA_INVALID_PARAMETER   As Long = 8437    'An invalid parameter was specified for this replication operation.
Public Const ERROR_DS_DRA_BUSY                As Long = 8438    'The directory service is too busy to complete the replication operation at this time.
Public Const ERROR_DS_DRA_BAD_DN              As Long = 8439    'The distinguished name specified for this replication operation is invalid.
Public Const ERROR_DS_DRA_BAD_NC              As Long = 8440    'The naming context specified for this replication operation is invalid.
Public Const ERROR_DS_DRA_DN_EXISTS           As Long = 8441    'The distinguished name specified for this replication operation already exists.
Public Const ERROR_DS_DRA_INTERNAL_ERROR      As Long = 8442    'The replication system encountered an internal error.
Public Const ERROR_DS_DRA_INCONSISTENT_DIT    As Long = 8443    'The replication operation encountered a database inconsistency.
Public Const ERROR_DS_DRA_CONNECTION_FAILED   As Long = 8444    'The server specified for this replication operation could not be contacted.
Public Const ERROR_DS_DRA_BAD_INSTANCE_TYPE   As Long = 8445    'The replication operation encountered an object with an invalid instance type.
Public Const ERROR_DS_DRA_OUT_OF_MEM          As Long = 8446    'The replication operation failed to allocate memory.
Public Const ERROR_DS_DRA_MAIL_PROBLEM        As Long = 8447    'The replication operation encountered an error with the mail system.
Public Const ERROR_DS_DRA_REF_ALREADY_EXISTS  As Long = 8448    'The replication reference information for the target server already exists.
Public Const ERROR_DS_DRA_REF_NOT_FOUND       As Long = 8449    'The replication reference information for the target server does not exist.
Public Const ERROR_DS_DRA_OBJ_IS_REP_SOURCE   As Long = 8450    'The naming context cannot be removed because it is replicated to another server.
Public Const ERROR_DS_DRA_DB_ERROR            As Long = 8451    'The replication operation encountered a database error.
Public Const ERROR_DS_DRA_NO_REPLICA          As Long = 8452    'The naming context is in the process of being removed or is not replicated from the specified server.
Public Const ERROR_DS_DRA_ACCESS_DENIED       As Long = 8453    'Replication access was denied.
Public Const ERROR_DS_DRA_NOT_SUPPORTED       As Long = 8454    'The requested operation is not supported by this version of the directory service.
Public Const ERROR_DS_DRA_RPC_CANCELLED       As Long = 8455    'The replication remote procedure call was cancelled.
Public Const ERROR_DS_DRA_SOURCE_DISABLED     As Long = 8456    'The source server is currently rejecting replication requests.
Public Const ERROR_DS_DRA_SINK_DISABLED       As Long = 8457    'The destination server is currently rejecting replication requests.
Public Const ERROR_DS_DRA_NAME_COLLISION      As Long = 8458    'The replication operation failed due to a collision of object names.
Public Const ERROR_DS_DRA_SOURCE_REINSTALLED  As Long = 8459    'The replication source has been reinstalled.
Public Const ERROR_DS_DRA_MISSING_PARENT      As Long = 8460    'The replication operation failed because a required parent object is missing.
Public Const ERROR_DS_DRA_PREEMPTED           As Long = 8461    'The replication operation was preempted.
Public Const ERROR_DS_DRA_ABANDON_SYNC        As Long = 8462    'The replication synchronization attempt was abandoned because of a lack of updates.
Public Const ERROR_DS_DRA_SHUTDOWN            As Long = 8463    'The replication operation was terminated because the system is shutting down.
Public Const ERROR_DS_DRA_INCOMPATIBLE_PARTIAL_SET As Long = 8464 'Synchronization attempt failed because the destination DC is currently waiting to synchronize new partial attributes from source. This condition is normal if a recent schema change modified the partial attribute set. The destination partial attribute set is not a subset of source partial attribute set.
Public Const ERROR_DS_DRA_SOURCE_IS_PARTIAL_REPLICA As Long = 8465 'The replication synchronization attempt failed because a master replica attempted to sync from a partial replica.
Public Const ERROR_DS_DRA_EXTN_CONNECTION_FAILED As Long = 8466 'The server specified for this replication operation was contacted, but that server was unable to contact an additional server needed to complete the operation.
Public Const ERROR_DS_INSTALL_SCHEMA_MISMATCH As Long = 8467    'The version of the Active Directory schema of the source forest is not compatible with the version of Active Directory on this computer.
Public Const ERROR_DS_DUP_LINK_ID             As Long = 8468    'Schema update failed: An attribute with the same link identifier already exists.
Public Const ERROR_DS_NAME_ERROR_RESOLVING    As Long = 8469    'Name translation: Generic processing error.
Public Const ERROR_DS_NAME_ERROR_NOT_FOUND    As Long = 8470    'Name translation: Could not find the name or insufficient right to see name.
Public Const ERROR_DS_NAME_ERROR_NOT_UNIQUE   As Long = 8471    'Name translation: Input name mapped to more than one output name.
Public Const ERROR_DS_NAME_ERROR_NO_MAPPING   As Long = 8472    'Name translation: Input name found, but not the associated output format.
Public Const ERROR_DS_NAME_ERROR_DOMAIN_ONLY  As Long = 8473    'Name translation: Unable to resolve completely, only the domain was found.
Public Const ERROR_DS_NAME_ERROR_NO_SYNTACTICAL_MAPPING As Long = 8474 'Name translation: Unable to perform purely syntactical mapping at the client without going out to the wire.
Public Const ERROR_DS_CONSTRUCTED_ATT_MOD     As Long = 8475    'Modification of a constructed attribute is not allowed.
Public Const ERROR_DS_WRONG_OM_OBJ_CLASS      As Long = 8476    'The OM-Object-Class specified is incorrect for an attribute with the specified syntax.
Public Const ERROR_DS_DRA_REPL_PENDING        As Long = 8477    'The replication request has been posted; waiting for reply.
Public Const ERROR_DS_DS_REQUIRED             As Long = 8478    'The requested operation requires a directory service, and none was available.
Public Const ERROR_DS_INVALID_LDAP_DISPLAY_NAME As Long = 8479  'The LDAP display name of the class or attribute contains non-ASCII characters.
Public Const ERROR_DS_NON_BASE_SEARCH         As Long = 8480    'The requested search operation is only supported for base searches.
Public Const ERROR_DS_CANT_RETRIEVE_ATTS      As Long = 8481    'The search failed to retrieve attributes from the database.
Public Const ERROR_DS_BACKLINK_WITHOUT_LINK   As Long = 8482    'The schema update operation tried to add a backward link attribute that has no corresponding forward link.
Public Const ERROR_DS_EPOCH_MISMATCH          As Long = 8483    'Source and destination of a cross-domain move do not agree on the object's epoch number.  Either source or destination does not have the latest version of the object.
Public Const ERROR_DS_SRC_NAME_MISMATCH       As Long = 8484    'Source and destination of a cross-domain move do not agree on the object's current name.  Either source or destination does not have the latest version of the object.
Public Const ERROR_DS_SRC_AND_DST_NC_IDENTICAL As Long = 8485   'Source and destination for the cross-domain move operation are identical.  Caller should use local move operation instead of cross-domain move operation.
Public Const ERROR_DS_DST_NC_MISMATCH         As Long = 8486    'Source and destination for a cross-domain move are not in agreement on the naming contexts in the forest.  Either source or destination does not have the latest version of the Partitions container.
Public Const ERROR_DS_NOT_AUTHORITIVE_FOR_DST_NC As Long = 8487 'Destination of a cross-domain move is not authoritative for the destination naming context.
Public Const ERROR_DS_SRC_GUID_MISMATCH       As Long = 8488    'Source and destination of a cross-domain move do not agree on the identity of the source object.  Either source or destination does not have the latest version of the source object.
Public Const ERROR_DS_CANT_MOVE_DELETED_OBJECT As Long = 8489   'Object being moved across-domains is already known to be deleted by the destination server.  The source server does not have the latest version of the source object.
Public Const ERROR_DS_PDC_OPERATION_IN_PROGRESS As Long = 8490  'Another operation which requires exclusive access to the PDC FSMO is already in progress.
Public Const ERROR_DS_CROSS_DOMAIN_CLEANUP_REQD As Long = 8491  'A cross-domain move operation failed such that two versions of the moved object exist - one each in the source and destination domains.  The destination object needs to be removed to restore the system to a consistent state.
Public Const ERROR_DS_ILLEGAL_XDOM_MOVE_OPERATION As Long = 8492 'This object may not be moved across domain boundaries either because cross-domain moves for this class are disallowed, or the object has some special characteristics, e.g.: trust account or restricted RID, which prevent its move.
Public Const ERROR_DS_CANT_WITH_ACCT_GROUP_MEMBERSHPS As Long = 8493 'Can't move objects with memberships across domain boundaries as once moved, this would violate the membership conditions of the account group.  Remove the object from any account group memberships and retry.
Public Const ERROR_DS_NC_MUST_HAVE_NC_PARENT  As Long = 8494    'A naming context head must be the immediate child of another naming context head, not of an interior node.
Public Const ERROR_DS_CR_IMPOSSIBLE_TO_VALIDATE As Long = 8495  'The directory cannot validate the proposed naming context name because it does not hold a replica of the naming context above the proposed naming context.  Please ensure that the domain naming master role is held by a server that is configured as a global catalog server, and that the server is up to date with its replication partners. (Applies only to Windows 2000 Domain Naming masters)
Public Const ERROR_DS_DST_DOMAIN_NOT_NATIVE   As Long = 8496    'Destination domain must be in native mode.
Public Const ERROR_DS_MISSING_INFRASTRUCTURE_CONTAINER As Long = 8497 'The operation can not be performed because the server does not have an infrastructure container in the domain of interest.
Public Const ERROR_DS_CANT_MOVE_ACCOUNT_GROUP As Long = 8498    'Cross-domain move of non-empty account groups is not allowed.
Public Const ERROR_DS_CANT_MOVE_RESOURCE_GROUP As Long = 8499   'Cross-domain move of non-empty resource groups is not allowed.
Public Const ERROR_DS_INVALID_SEARCH_FLAG     As Long = 8500    'The search flags for the attribute are invalid. The ANR bit is valid only on attributes of Unicode or Teletex strings.
Public Const ERROR_DS_NO_TREE_DELETE_ABOVE_NC As Long = 8501    'Tree deletions starting at an object which has an NC head as a descendant are not allowed.
Public Const ERROR_DS_COULDNT_LOCK_TREE_FOR_DELETE As Long = 8502 'The directory service failed to lock a tree in preparation for a tree deletion because the tree was in use.
Public Const ERROR_DS_COULDNT_IDENTIFY_OBJECTS_FOR_TREE_DELETE As Long = 8503    'The directory service failed to identify the list of objects to delete while attempting a tree deletion.
Public Const ERROR_DS_SAM_INIT_FAILURE        As Long = 8504    'Security Accounts Manager initialization failed because of the following error: %1.    'Error Status: 0x%2. Click OK to shut down the system and reboot into Directory Services Restore Mode. Check the event log for detailed information.
Public Const ERROR_DS_SENSITIVE_GROUP_VIOLATION As Long = 8505  'Only an administrator can modify the membership list of an administrative group.
Public Const ERROR_DS_CANT_MOD_PRIMARYGROUPID As Long = 8506    'Cannot change the primary group ID of a domain controller account.
Public Const ERROR_DS_ILLEGAL_BASE_SCHEMA_MOD As Long = 8507    'An attempt is made to modify the base schema.
Public Const ERROR_DS_NONSAFE_SCHEMA_CHANGE   As Long = 8508    'Adding a new mandatory attribute to an existing class, deleting a mandatory attribute from an existing class, or adding an optional attribute to the special class Top that is not a backlink attribute (directly or through inheritance, for example, by adding or deleting an auxiliary class) is not allowed.
Public Const ERROR_DS_SCHEMA_UPDATE_DISALLOWED As Long = 8509   'Schema update is not allowed on this DC because the DC is not the schema FSMO Role Owner.
Public Const ERROR_DS_CANT_CREATE_UNDER_SCHEMA As Long = 8510   'An object of this class cannot be created under the schema container. You can only create attribute-schema and class-schema objects under the schema container.
Public Const ERROR_DS_INSTALL_NO_SRC_SCH_VERSION As Long = 8511 'The replica/child install failed to get the objectVersion attribute on the schema container on the source DC. Either the attribute is missing on the schema container or the credentials supplied do not have permission to read it.
Public Const ERROR_DS_INSTALL_NO_SCH_VERSION_IN_INIFILE As Long = 8512    'The replica/child install failed to read the objectVersion attribute in the SCHEMA section of the file schema.ini in the system32 directory.
Public Const ERROR_DS_INVALID_GROUP_TYPE      As Long = 8513    'The specified group type is invalid.
Public Const ERROR_DS_NO_NEST_GLOBALGROUP_IN_MIXEDDOMAIN As Long = 8514 'You cannot nest global groups in a mixed domain if the group is security-enabled.
Public Const ERROR_DS_NO_NEST_LOCALGROUP_IN_MIXEDDOMAIN As Long = 8515 'You cannot nest local groups in a mixed domain if the group is security-enabled.
Public Const ERROR_DS_GLOBAL_CANT_HAVE_LOCAL_MEMBER As Long = 8516 'A global group cannot have a local group as a member.
Public Const ERROR_DS_GLOBAL_CANT_HAVE_UNIVERSAL_MEMBER As Long = 8517 'A global group cannot have a universal group as a member.
Public Const ERROR_DS_UNIVERSAL_CANT_HAVE_LOCAL_MEMBER As Long = 8518 'A universal group cannot have a local group as a member.
Public Const ERROR_DS_GLOBAL_CANT_HAVE_CROSSDOMAIN_MEMBER As Long = 8519    'A global group cannot have a cross-domain member.
Public Const ERROR_DS_LOCAL_CANT_HAVE_CROSSDOMAIN_LOCAL_MEMBER As Long = 8520    'A local group cannot have another cross domain local group as a member.
Public Const ERROR_DS_HAVE_PRIMARY_MEMBERS    As Long = 8521    'A group with primary members cannot change to a security-disabled group.
Public Const ERROR_DS_STRING_SD_CONVERSION_FAILED As Long = 8522 'The schema cache load failed to convert the string default SD on a class-schema object.
Public Const ERROR_DS_NAMING_MASTER_GC        As Long = 8523    'Only DSAs configured to be Global Catalog servers should be allowed to hold the Domain Naming Master FSMO role. (Applies only to Windows 2000 servers)
Public Const ERROR_DS_DNS_LOOKUP_FAILURE      As Long = 8524    'The DSA operation is unable to proceed because of a DNS lookup failure.
Public Const ERROR_DS_COULDNT_UPDATE_SPNS     As Long = 8525    'While processing a change to the DNS Host Name for an object, the Service Principal Name values could not be kept in sync.
Public Const ERROR_DS_CANT_RETRIEVE_SD        As Long = 8526    'The Security Descriptor attribute could not be read.
Public Const ERROR_DS_KEY_NOT_UNIQUE          As Long = 8527    'The object requested was not found, but an object with that key was found.
Public Const ERROR_DS_WRONG_LINKED_ATT_SYNTAX As Long = 8528    'The syntax of the linked attribute being added is incorrect. Forward links can only have syntax 2.5.5.1, 2.5.5.7, and 2.5.5.14, and backlinks can only have syntax 2.5.5.1
Public Const ERROR_DS_SAM_NEED_BOOTKEY_PASSWORD As Long = 8529  'Security Account Manager needs to get the boot password.
Public Const ERROR_DS_SAM_NEED_BOOTKEY_FLOPPY As Long = 8530    'Security Account Manager needs to get the boot key from floppy disk.
Public Const ERROR_DS_CANT_START              As Long = 8531    'Directory Service cannot start.
Public Const ERROR_DS_INIT_FAILURE            As Long = 8532    'Directory Services could not start.
Public Const ERROR_DS_NO_PKT_PRIVACY_ON_CONNECTION As Long = 8533 'The connection between client and server requires packet privacy or better.
Public Const ERROR_DS_SOURCE_DOMAIN_IN_FOREST As Long = 8534    'The source domain may not be in the same forest as destination.
Public Const ERROR_DS_DESTINATION_DOMAIN_NOT_IN_FOREST As Long = 8535 'The destination domain must be in the forest.
Public Const ERROR_DS_DESTINATION_AUDITING_NOT_ENABLED As Long = 8536 'The operation requires that destination domain auditing be enabled.
Public Const ERROR_DS_CANT_FIND_DC_FOR_SRC_DOMAIN As Long = 8537 'The operation couldn't locate a DC for the source domain.
Public Const ERROR_DS_SRC_OBJ_NOT_GROUP_OR_USER As Long = 8538  'The source object must be a group or user.
Public Const ERROR_DS_SRC_SID_EXISTS_IN_FOREST As Long = 8539   'The source object's SID already exists in destination forest.
Public Const ERROR_DS_SRC_AND_DST_OBJECT_CLASS_MISMATCH As Long = 8540 'The source and destination object must be of the same type.
Public Const ERROR_SAM_INIT_FAILURE           As Long = 8541    'Security Accounts Manager initialization failed because of the following error: %1.    'Error Status: 0x%2. Click OK to shut down the system and reboot into Safe Mode. Check the event log for detailed information.
Public Const ERROR_DS_DRA_SCHEMA_INFO_SHIP    As Long = 8542    'Schema information could not be included in the replication request.
Public Const ERROR_DS_DRA_SCHEMA_CONFLICT     As Long = 8543    'The replication operation could not be completed due to a schema incompatibility.
Public Const ERROR_DS_DRA_EARLIER_SCHEMA_CONFLICT As Long = 8544 'The replication operation could not be completed due to a previous schema incompatibility.
Public Const ERROR_DS_DRA_OBJ_NC_MISMATCH     As Long = 8545    'The replication update could not be applied because either the source or the destination has not yet received information regarding a recent cross-domain move operation.
Public Const ERROR_DS_NC_STILL_HAS_DSAS       As Long = 8546    'The requested domain could not be deleted because there exist domain controllers that still host this domain.
Public Const ERROR_DS_GC_REQUIRED             As Long = 8547    'The requested operation can be performed only on a global catalog server.
Public Const ERROR_DS_LOCAL_MEMBER_OF_LOCAL_ONLY As Long = 8548 'A local group can only be a member of other local groups in the same domain.
Public Const ERROR_DS_NO_FPO_IN_UNIVERSAL_GROUPS As Long = 8549 'Foreign security principals cannot be members of universal groups.
Public Const ERROR_DS_CANT_ADD_TO_GC          As Long = 8550    'The attribute is not allowed to be replicated to the GC because of security reasons.
Public Const ERROR_DS_NO_CHECKPOINT_WITH_PDC  As Long = 8551    'The checkpoint with the PDC could not be taken because there too many modifications being processed currently.
Public Const ERROR_DS_SOURCE_AUDITING_NOT_ENABLED As Long = 8552 'The operation requires that source domain auditing be enabled.
Public Const ERROR_DS_CANT_CREATE_IN_NONDOMAIN_NC As Long = 8553 'Security principal objects can only be created inside domain naming contexts.
Public Const ERROR_DS_INVALID_NAME_FOR_SPN    As Long = 8554    'A Service Principal Name (SPN) could not be constructed because the provided hostname is not in the necessary format.
Public Const ERROR_DS_FILTER_USES_CONTRUCTED_ATTRS As Long = 8555 'A Filter was passed that uses constructed attributes.
Public Const ERROR_DS_UNICODEPWD_NOT_IN_QUOTES As Long = 8556   'The unicodePwd attribute value must be enclosed in double quotes.
Public Const ERROR_DS_MACHINE_ACCOUNT_QUOTA_EXCEEDED As Long = 8557 'Your computer could not be joined to the domain. You have exceeded the maximum number of computer accounts you are allowed to create in this domain. Contact your system administrator to have this limit reset or increased.
Public Const ERROR_DS_MUST_BE_RUN_ON_DST_DC   As Long = 8558    'For security reasons, the operation must be run on the destination DC.
Public Const ERROR_DS_SRC_DC_MUST_BE_SP4_OR_GREATER As Long = 8559 'For security reasons, the source DC must be NT4SP4 or greater.
Public Const ERROR_DS_CANT_TREE_DELETE_CRITICAL_OBJ As Long = 8560 'Critical Directory Service System objects cannot be deleted during tree delete operations.  The tree delete may have been partially performed.
Public Const ERROR_DS_INIT_FAILURE_CONSOLE    As Long = 8561    'Directory Services could not start because of the following error: %1.    'Error Status: 0x%2. Please click OK to shutdown the system. You can use the recovery console to diagnose the system further.
Public Const ERROR_DS_SAM_INIT_FAILURE_CONSOLE As Long = 8562   'Security Accounts Manager initialization failed because of the following error: %1.    'Error Status: 0x%2. Please click OK to shutdown the system. You can use the recovery console to diagnose the system further.
Public Const ERROR_DS_FOREST_VERSION_TOO_HIGH As Long = 8563    'The version of the operating system installed is incompatible with the current forest functional level. You must upgrade to a new version of the operating system before this server can become a domain controller in this forest.
Public Const ERROR_DS_DOMAIN_VERSION_TOO_HIGH As Long = 8564    'The version of the operating system installed is incompatible with the current domain functional level. You must upgrade to a new version of the operating system before this server can become a domain controller in this domain.
Public Const ERROR_DS_FOREST_VERSION_TOO_LOW  As Long = 8565    'The version of the operating system installed on this server no longer supports the current forest functional level. You must raise the forest functional level before this server can become a domain controller in this forest.
Public Const ERROR_DS_DOMAIN_VERSION_TOO_LOW  As Long = 8566    'The version of the operating system installed on this server no longer supports the current domain functional level. You must raise the domain functional level before this server can become a domain controller in this domain.
Public Const ERROR_DS_INCOMPATIBLE_VERSION    As Long = 8567    'The version of the operating system installed on this server is incompatible with the functional level of the domain or forest.
Public Const ERROR_DS_LOW_DSA_VERSION         As Long = 8568    'The functional level of the domain (or forest) cannot be raised to the requested value, because there exist one or more domain controllers in the domain (or forest) that are at a lower incompatible functional level.
Public Const ERROR_DS_NO_BEHAVIOR_VERSION_IN_MIXEDDOMAIN As Long = 8569 'The forest functional level cannot be raised to the requested value since one or more domains are still in mixed domain mode. All domains in the forest must be in native mode, for you to raise the forest functional level.
Public Const ERROR_DS_NOT_SUPPORTED_SORT_ORDER As Long = 8570   'The sort order requested is not supported.
Public Const ERROR_DS_NAME_NOT_UNIQUE         As Long = 8571    'The requested name already exists as a unique identifier.
Public Const ERROR_DS_MACHINE_ACCOUNT_CREATED_PRENT4 As Long = 8572 'The machine account was created pre-NT4.  The account needs to be recreated.
Public Const ERROR_DS_OUT_OF_VERSION_STORE    As Long = 8573    'The database is out of version store.
Public Const ERROR_DS_INCOMPATIBLE_CONTROLS_USED As Long = 8574 'Unable to continue operation because multiple conflicting controls were used.
Public Const ERROR_DS_NO_REF_DOMAIN           As Long = 8575    'Unable to find a valid security descriptor reference domain for this partition.
Public Const ERROR_DS_RESERVED_LINK_ID        As Long = 8576    'Schema update failed: The link identifier is reserved.
Public Const ERROR_DS_LINK_ID_NOT_AVAILABLE   As Long = 8577    'Schema update failed: There are no link identifiers available.
Public Const ERROR_DS_AG_CANT_HAVE_UNIVERSAL_MEMBER As Long = 8578 'An account group can not have a universal group as a member.
Public Const ERROR_DS_MODIFYDN_DISALLOWED_BY_INSTANCE_TYPE As Long = 8579 'Rename or move operations on naming context heads or read-only objects are not allowed.
Public Const ERROR_DS_NO_OBJECT_MOVE_IN_SCHEMA_NC As Long = 8580 'Move operations on objects in the schema naming context are not allowed.
Public Const ERROR_DS_MODIFYDN_DISALLOWED_BY_FLAG As Long = 8581 'A system flag has been set on the object and does not allow the object to be moved or renamed.
Public Const ERROR_DS_MODIFYDN_WRONG_GRANDPARENT As Long = 8582 'This object is not allowed to change its grandparent container. Moves are not forbidden on this object, but are restricted to sibling containers.
Public Const ERROR_DS_NAME_ERROR_TRUST_REFERRAL As Long = 8583  'Unable to resolve completely, a referral to another forest is generated.
Public Const ERROR_NOT_SUPPORTED_ON_STANDARD_SERVER As Long = 8584 'The requested action is not supported on standard server.
Public Const ERROR_DS_CANT_ACCESS_REMOTE_PART_OF_AD As Long = 8585 'Could not access a partition of the Active Directory located on a remote server.  Make sure at least one server is running for the partition in question.
Public Const ERROR_DS_CR_IMPOSSIBLE_TO_VALIDATE_V2 As Long = 8586 'The directory cannot validate the proposed naming context (or partition) name because it does not hold a replica nor can it contact a replica of the naming context above the proposed naming context.  Please ensure that the parent naming context is properly registered in DNS, and at least one replica of this naming context is reachable by the Domain Naming master.
Public Const ERROR_DS_THREAD_LIMIT_EXCEEDED   As Long = 8587    'The thread limit for this request was exceeded.
Public Const ERROR_DS_NOT_CLOSEST             As Long = 8588    'The Global catalog server is not in the closest site.
Public Const ERROR_DS_CANT_DERIVE_SPN_WITHOUT_SERVER_REF As Long = 8589 'The DS cannot derive a service principal name (SPN) with which to mutually authenticate the target server because the corresponding server object in the local DS database has no serverReference attribute.
Public Const ERROR_DS_SINGLE_USER_MODE_FAILED As Long = 8590    'The Directory Service failed to enter single user mode.
Public Const ERROR_DS_NTDSCRIPT_SYNTAX_ERROR  As Long = 8591    'The Directory Service cannot parse the script because of a syntax error.
Public Const ERROR_DS_NTDSCRIPT_PROCESS_ERROR As Long = 8592    'The Directory Service cannot process the script because of an error.
Public Const ERROR_DS_DIFFERENT_REPL_EPOCHS   As Long = 8593    'The directory service cannot perform the requested operation because the servers    'involved are of different replication epochs (which is usually related to a    'domain rename that is in progress).
Public Const ERROR_DS_DRS_EXTENSIONS_CHANGED  As Long = 8594    'The directory service binding must be renegotiated due to a change in the server    'extensions information.
Public Const ERROR_DS_REPLICA_SET_CHANGE_NOT_ALLOWED_ON_DISABLED_CR As Long = 8595    'Operation not allowed on a disabled cross ref.
Public Const ERROR_DS_NO_MSDS_INTID           As Long = 8596    'Schema update failed: No values for msDS-IntId are available.
Public Const ERROR_DS_DUP_MSDS_INTID          As Long = 8597    'Schema update failed: Duplicate msDS-INtId. Retry the operation.
Public Const ERROR_DS_EXISTS_IN_RDNATTID      As Long = 8598    'Schema deletion failed: attribute is used in rDNAttID.
Public Const ERROR_DS_AUTHORIZATION_FAILED    As Long = 8599    'The directory service failed to authorize the request.
Public Const ERROR_DS_INVALID_SCRIPT          As Long = 8600    'The Directory Service cannot process the script because it is invalid.
Public Const ERROR_DS_REMOTE_CROSSREF_OP_FAILED As Long = 8601  'The remote create cross reference operation failed on the Domain Naming Master FSMO.  The operation's error is in the extended data.
Public Const ERROR_DS_CROSS_REF_BUSY          As Long = 8602    'A cross reference is in use locally with the same name.
Public Const ERROR_DS_CANT_DERIVE_SPN_FOR_DELETED_DOMAIN As Long = 8603 'The DS cannot derive a service principal name (SPN) with which to mutually authenticate the target server because the server's domain has been deleted from the forest.
Public Const ERROR_DS_CANT_DEMOTE_WITH_WRITEABLE_NC As Long = 8604 'Writeable NCs prevent this DC from demoting.
Public Const ERROR_DS_DUPLICATE_ID_FOUND      As Long = 8605    'The requested object has a non-unique identifier and cannot be retrieved.
Public Const ERROR_DS_INSUFFICIENT_ATTR_TO_CREATE_OBJECT As Long = 8606 'Insufficient attributes were given to create an object.  This object may not exist because it may have been deleted and already garbage collected.
Public Const ERROR_DS_GROUP_CONVERSION_ERROR  As Long = 8607    'The group cannot be converted due to attribute restrictions on the requested group type.
Public Const ERROR_DS_CANT_MOVE_APP_BASIC_GROUP As Long = 8608  'Cross-domain move of non-empty basic application groups is not allowed.
Public Const ERROR_DS_CANT_MOVE_APP_QUERY_GROUP As Long = 8609  'Cross-domain move of non-empty query based application groups is not allowed.
Public Const ERROR_DS_ROLE_NOT_VERIFIED       As Long = 8610    'The FSMO role ownership could not be verified because its directory partition has not replicated successfully with atleast one replication partner.
Public Const ERROR_DS_WKO_CONTAINER_CANNOT_BE_SPECIAL As Long = 8611 'The target container for a redirection of a well known object container cannot already be a special container.
Public Const ERROR_DS_DOMAIN_RENAME_IN_PROGRESS As Long = 8612  'The Directory Service cannot perform the requested operation because a domain rename operation is in progress.
Public Const ERROR_DS_EXISTING_AD_CHILD_NC    As Long = 8613    'The Active Directory detected an Active Directory child partition below the    'requested new partition name.  The Active Directory's partition heiarchy must    'be created in a top down method.
Public Const ERROR_DS_REPL_LIFETIME_EXCEEDED  As Long = 8614    'The Active Directory cannot replicate with this server because the time since the last replication with this server has exceeded the tombstone lifetime.
Public Const ERROR_DS_DISALLOWED_IN_SYSTEM_CONTAINER As Long = 8615 'The requested operation is not allowed on an object under the system container.
Public Const ERROR_DS_LDAP_SEND_QUEUE_FULL    As Long = 8616    'The LDAP servers network send queue has filled up because the client is not    'processing the results of it's requests fast enough.  No more requests will    'be processed until the client catches up.  If the client does not catch up    'then it will be disconnected.
Public Const ERROR_DS_DRA_OUT_SCHEDULE_WINDOW As Long = 8617    'The scheduled replication did not take place because the system was too busy to execute the request within the schedule window.  The replication queue is overloaded. Consider reducing the number of partners or decreasing the scheduled replication frequency.
'
'///////////////////////////////////////////////////
'//                                                /
'//     End of Active Directory Error Codes        /
'//                                                /
'//                  8000 to  8999                 /
'///////////////////////////////////////////////////
'
'
'///////////////////////////////////////////////////
'//                                               //
'//                  DNS Error Codes              //
'//                                               //
'//                   9000 to 9999                //
'///////////////////////////////////////////////////
'
'// =============================
'// Facility DNS Error Messages
'// =============================
'
'//
'//  DNS response codes.
'//
'
'Public Const DNS_ERROR_RESPONSE_CODES_BASE 9000
'
'Public Const DNS_ERROR_RCODE_NO_ERROR NO_ERROR
'
'Public Const DNS_ERROR_MASK 0x00002328 // 9000 or DNS_ERROR_RESPONSE_CODES_BASE
Public Const DNS_ERROR_RCODE_FORMAT_ERROR     As Long = 9001    'DNS server unable to interpret format.
Public Const DNS_ERROR_RCODE_SERVER_FAILURE   As Long = 9002    'DNS server failure.
Public Const DNS_ERROR_RCODE_NAME_ERROR       As Long = 9003    'DNS name does not exist.
Public Const DNS_ERROR_RCODE_NOT_IMPLEMENTED  As Long = 9004    'DNS request not supported by name server.
Public Const DNS_ERROR_RCODE_REFUSED          As Long = 9005    'DNS operation refused.
Public Const DNS_ERROR_RCODE_YXDOMAIN         As Long = 9006    'DNS name that ought not exist, does exist.
Public Const DNS_ERROR_RCODE_YXRRSET          As Long = 9007    'DNS RR set that ought not exist, does exist.
Public Const DNS_ERROR_RCODE_NXRRSET          As Long = 9008    'DNS RR set that ought to exist, does not exist.
Public Const DNS_ERROR_RCODE_NOTAUTH          As Long = 9009    'DNS server not authoritative for zone.
Public Const DNS_ERROR_RCODE_NOTZONE          As Long = 9010    'DNS name in update or prereq is not in zone.
Public Const DNS_ERROR_RCODE_BADSIG           As Long = 9016    'DNS signature failed to verify.
Public Const DNS_ERROR_RCODE_BADKEY           As Long = 9017    'DNS bad key.
Public Const DNS_ERROR_RCODE_BADTIME          As Long = 9018    'DNS signature validity expired.
Public Const DNS_INFO_NO_RECORDS              As Long = 9501    'No records found for given DNS query.
Public Const DNS_ERROR_BAD_PACKET             As Long = 9502    'Bad DNS packet.
Public Const DNS_ERROR_NO_PACKET              As Long = 9503    'No DNS packet.
Public Const DNS_ERROR_RCODE                  As Long = 9504    'DNS error, check rcode.
Public Const DNS_ERROR_UNSECURE_PACKET        As Long = 9505    'Unsecured DNS packet.

'
'Public Const DNS_STATUS_PACKET_UNSECURE DNS_ERROR_UNSECURE_PACKET
'
'
'//
'//  General API errors
'//
'
'Public Const DNS_ERROR_NO_MEMORY            ERROR_OUTOFMEMORY
'Public Const DNS_ERROR_INVALID_NAME         ERROR_INVALID_NAME
'Public Const DNS_ERROR_INVALID_DATA         ERROR_INVALID_DATA
'
'Public Const DNS_ERROR_GENERAL_API_BASE As Long = 9550
Public Const DNS_ERROR_INVALID_TYPE           As Long = 9551    'Invalid DNS type.
Public Const DNS_ERROR_INVALID_IP_ADDRESS     As Long = 9552    'Invalid IP address.
Public Const DNS_ERROR_INVALID_PROPERTY       As Long = 9553    'Invalid property.
Public Const DNS_ERROR_TRY_AGAIN_LATER        As Long = 9554    'Try DNS operation again later.
Public Const DNS_ERROR_NOT_UNIQUE             As Long = 9555    'Record for given name and type is not unique.
Public Const DNS_ERROR_NON_RFC_NAME           As Long = 9556    'DNS name does not comply with RFC specifications.
Public Const DNS_STATUS_FQDN                  As Long = 9557    'DNS name is a fully-qualified DNS name.
Public Const DNS_STATUS_DOTTED_NAME           As Long = 9558    'DNS name is dotted (multi-label).
Public Const DNS_STATUS_SINGLE_PART_NAME      As Long = 9559    'DNS name is a single-part name.
Public Const DNS_ERROR_INVALID_NAME_CHAR      As Long = 9560    'DNS name contains an invalid character.
Public Const DNS_ERROR_NUMERIC_NAME           As Long = 9561    'DNS name is entirely numeric.
Public Const DNS_ERROR_NOT_ALLOWED_ON_ROOT_SERVER As Long = 9562 'The operation requested is not permitted on a DNS root server.
Public Const DNS_ERROR_NOT_ALLOWED_UNDER_DELEGATION As Long = 9563 'The record could not be created because this part of the DNS namespace has    'been delegated to another server.
Public Const DNS_ERROR_CANNOT_FIND_ROOT_HINTS As Long = 9564    'The DNS server could not find a set of root hints.
Public Const DNS_ERROR_INCONSISTENT_ROOT_HINTS As Long = 9565   'The DNS server found root hints but they were not consistent across    'all adapters.
Public Const DNS_ERROR_ZONE_DOES_NOT_EXIST    As Long = 9601    'DNS zone does not exist.
Public Const DNS_ERROR_NO_ZONE_INFO           As Long = 9602    'DNS zone information not available.
Public Const DNS_ERROR_INVALID_ZONE_OPERATION As Long = 9603    'Invalid operation for DNS zone.
Public Const DNS_ERROR_ZONE_CONFIGURATION_ERROR As Long = 9604  'Invalid DNS zone configuration.
Public Const DNS_ERROR_ZONE_HAS_NO_SOA_RECORD As Long = 9605    'DNS zone has no start of authority (SOA) record.
Public Const DNS_ERROR_ZONE_HAS_NO_NS_RECORDS As Long = 9606    'DNS zone has no Name Server (NS) record.
Public Const DNS_ERROR_ZONE_LOCKED            As Long = 9607    'DNS zone is locked.
Public Const DNS_ERROR_ZONE_CREATION_FAILED   As Long = 9608    'DNS zone creation failed.
Public Const DNS_ERROR_ZONE_ALREADY_EXISTS    As Long = 9609    'DNS zone already exists.
Public Const DNS_ERROR_AUTOZONE_ALREADY_EXISTS As Long = 9610   'DNS automatic zone already exists.
Public Const DNS_ERROR_INVALID_ZONE_TYPE      As Long = 9611    'Invalid DNS zone type.
Public Const DNS_ERROR_SECONDARY_REQUIRES_MASTER_IP As Long = 9612 'Secondary DNS zone requires master IP address.
Public Const DNS_ERROR_ZONE_NOT_SECONDARY     As Long = 9613    'DNS zone not secondary.
Public Const DNS_ERROR_NEED_SECONDARY_ADDRESSES As Long = 9614  'Need secondary IP address.
Public Const DNS_ERROR_WINS_INIT_FAILED       As Long = 9615    'WINS initialization failed.
Public Const DNS_ERROR_NEED_WINS_SERVERS      As Long = 9616    'Need WINS servers.
Public Const DNS_ERROR_NBSTAT_INIT_FAILED     As Long = 9617    'NBTSTAT initialization call failed.
Public Const DNS_ERROR_SOA_DELETE_INVALID     As Long = 9618    'Invalid delete of start of authority (SOA)
Public Const DNS_ERROR_FORWARDER_ALREADY_EXISTS As Long = 9619  'A conditional forwarding zone already exists for that name.
Public Const DNS_ERROR_ZONE_REQUIRES_MASTER_IP As Long = 9620   'This zone must be configured with one or more master DNS server IP addresses.
Public Const DNS_ERROR_ZONE_IS_SHUTDOWN       As Long = 9621    'The operation cannot be performed because this zone is shutdown.
Public Const DNS_ERROR_PRIMARY_REQUIRES_DATAFILE As Long = 9651 'Primary DNS zone requires datafile.
Public Const DNS_ERROR_INVALID_DATAFILE_NAME  As Long = 9652    'Invalid datafile name for DNS zone.
Public Const DNS_ERROR_DATAFILE_OPEN_FAILURE  As Long = 9653    'Failed to open datafile for DNS zone.
Public Const DNS_ERROR_FILE_WRITEBACK_FAILED  As Long = 9654    'Failed to write datafile for DNS zone.
Public Const DNS_ERROR_DATAFILE_PARSING       As Long = 9655    'Failure while reading datafile for DNS zone.
Public Const DNS_ERROR_RECORD_DOES_NOT_EXIST  As Long = 9701    'DNS record does not exist.
Public Const DNS_ERROR_RECORD_FORMAT          As Long = 9702    'DNS record format error.
Public Const DNS_ERROR_NODE_CREATION_FAILED   As Long = 9703    'Node creation failure in DNS.
Public Const DNS_ERROR_UNKNOWN_RECORD_TYPE    As Long = 9704    'Unknown DNS record type.
Public Const DNS_ERROR_RECORD_TIMED_OUT       As Long = 9705    'DNS record timed out.
Public Const DNS_ERROR_NAME_NOT_IN_ZONE       As Long = 9706    'Name not in DNS zone.
Public Const DNS_ERROR_CNAME_LOOP             As Long = 9707    'CNAME loop detected.
Public Const DNS_ERROR_NODE_IS_CNAME          As Long = 9708    'Node is a CNAME DNS record.
Public Const DNS_ERROR_CNAME_COLLISION        As Long = 9709    'A CNAME record already exists for given name.
Public Const DNS_ERROR_RECORD_ONLY_AT_ZONE_ROOT As Long = 9710  'Record only at DNS zone root.
Public Const DNS_ERROR_RECORD_ALREADY_EXISTS  As Long = 9711    'DNS record already exists.
Public Const DNS_ERROR_SECONDARY_DATA         As Long = 9712    'Secondary DNS zone data error.
Public Const DNS_ERROR_NO_CREATE_CACHE_DATA   As Long = 9713    'Could not create DNS cache data.
Public Const DNS_ERROR_NAME_DOES_NOT_EXIST    As Long = 9714    'DNS name does not exist.
Public Const DNS_WARNING_PTR_CREATE_FAILED    As Long = 9715    'Could not create pointer (PTR) record.
Public Const DNS_WARNING_DOMAIN_UNDELETED     As Long = 9716    'DNS domain was undeleted.
Public Const DNS_ERROR_DS_UNAVAILABLE         As Long = 9717    'The directory service is unavailable.
Public Const DNS_ERROR_DS_ZONE_ALREADY_EXISTS As Long = 9718    'DNS zone already exists in the directory service.
Public Const DNS_ERROR_NO_BOOTFILE_IF_DS_ZONE As Long = 9719    'DNS server not creating or reading the boot file for the directory service integrated DNS zone.

'
'
'//
'//  Operation errors
'//
'
'Public Const DNS_ERROR_OPERATION_BASE 9750
'DNS_INFO_AXFR_COMPLETE                0x00002617
'Public Const DNS_INFO_AXFR_COMPLETE           9751L'DNS AXFR (zone transfer) complete.
'DNS_ERROR_AXFR                        0x00002618
'Public Const DNS_ERROR_AXFR                   9752L'DNS zone transfer failed.
'DNS_INFO_ADDED_LOCAL_WINS             0x00002619
'Public Const DNS_INFO_ADDED_LOCAL_WINS        9753L'Added local WINS server.
'
'//
'//  Secure update
'//
'
'Public Const DNS_ERROR_SECURE_BASE 9800
'
'// DNS_STATUS_CONTINUE_NEEDED            0x00002649
'//
'// MessageId: DNS_STATUS_CONTINUE_NEEDED
'//
'// MessageText:
'//
'//  Secure update call needs to continue update request.
'//
'Public Const DNS_STATUS_CONTINUE_NEEDED       9801L
'
'
'//
'//  Setup errors
'//
'
'Public Const DNS_ERROR_SETUP_BASE 9850
'
'// DNS_ERROR_NO_TCPIP                    0x0000267b
'//
'// MessageId: DNS_ERROR_NO_TCPIP
'//
'// MessageText:
'//
'//  TCP/IP network protocol not installed.
'//
'Public Const DNS_ERROR_NO_TCPIP               9851L
'
'// DNS_ERROR_NO_DNS_SERVERS              0x0000267c
'//
'// MessageId: DNS_ERROR_NO_DNS_SERVERS
'//
'// MessageText:
'//
'//  No DNS servers configured for local system.
'//
'Public Const DNS_ERROR_NO_DNS_SERVERS         9852L
'
'
'//
'//  Directory partition (DP) errors
'//
'
'Public Const DNS_ERROR_DP_BASE 9900
'
'// DNS_ERROR_DP_DOES_NOT_EXIST           0x000026ad
'//
'// MessageId: DNS_ERROR_DP_DOES_NOT_EXIST
'//
'// MessageText:
'//
'//  The specified directory partition does not exist.
'//
'Public Const DNS_ERROR_DP_DOES_NOT_EXIST      9901L
'
'// DNS_ERROR_DP_ALREADY_EXISTS           0x000026ae
'//
'// MessageId: DNS_ERROR_DP_ALREADY_EXISTS
'//
'// MessageText:
'//
'//  The specified directory partition already exists.
'//
'Public Const DNS_ERROR_DP_ALREADY_EXISTS      9902L
'
'// DNS_ERROR_DP_NOT_ENLISTED             0x000026af
'//
'// MessageId: DNS_ERROR_DP_NOT_ENLISTED
'//
'// MessageText:
'//
'//  This DNS server is not enlisted in the specified directory partition.
'//
'Public Const DNS_ERROR_DP_NOT_ENLISTED        9903L
'
'// DNS_ERROR_DP_ALREADY_ENLISTED         0x000026b0
'//
'// MessageId: DNS_ERROR_DP_ALREADY_ENLISTED
'//
'// MessageText:
'//
'//  This DNS server is already enlisted in the specified directory partition.
'//
'Public Const DNS_ERROR_DP_ALREADY_ENLISTED    9904L
'
'// DNS_ERROR_DP_NOT_AVAILABLE            0x000026b1
'//
'// MessageId: DNS_ERROR_DP_NOT_AVAILABLE
'//
'// MessageText:
'//
'//  The directory partition is not available at this time. Please wait
'//  a few minutes and try again.
'//
'Public Const DNS_ERROR_DP_NOT_AVAILABLE       9905L
'
'// DNS_ERROR_DP_FSMO_ERROR               0x000026b2
'//
'// MessageId: DNS_ERROR_DP_FSMO_ERROR
'//
'// MessageText:
'//
'//  The application directory partition operation failed. The domain controller
'//  holding the domain naming master role is down or unable to service the
'//  request or is not running Windows Server 2003.
'//
'Public Const DNS_ERROR_DP_FSMO_ERROR          9906L
'
'///////////////////////////////////////////////////
'//                                               //
'//             End of DNS Error Codes            //
'//                                               //
'//                  9000 to 9999                 //
'///////////////////////////////////////////////////
'
'
'///////////////////////////////////////////////////
'//                                               //
'//               WinSock Error Codes             //
'//                                               //
'//                 10000 to 11999                //
'///////////////////////////////////////////////////
'
'//
'// WinSock error codes are also defined in WinSock.h
'// and WinSock2.h, hence the IFDEF
'//
'#ifndef WSABASEERR
'Public Const WSABASEERR 10000
Public Const ERROR_WSAEINTR                         As Long = 10004    'A blocking operation was interrupted by a call to WSACancelBlockingCall.
Public Const ERROR_WSAEBADF                         As Long = 10009    'The file handle supplied is not valid.
Public Const ERROR_WSAEACCES                        As Long = 10013    'An attempt was made to access a socket in a way forbidden by its access permissions.
Public Const ERROR_WSAEFAULT                        As Long = 10014    'The system detected an invalid pointer address in attempting to use a pointer argument in a call.
Public Const ERROR_WSAEINVAL                        As Long = 10022    'An invalid argument was supplied.
Public Const ERROR_WSAEMFILE                        As Long = 10024    'Too many open sockets.
Public Const ERROR_WSAEWOULDBLOCK                   As Long = 10035    'A non-blocking socket operation could not be completed immediately.
Public Const ERROR_WSAEINPROGRESS                   As Long = 10036    'A blocking operation is currently executing.
Public Const ERROR_WSAEALREADY                      As Long = 10037    'An operation was attempted on a non-blocking socket that already had an operation in progress.
Public Const ERROR_WSAENOTSOCK                      As Long = 10038    'An operation was attempted on something that is not a socket.
Public Const ERROR_WSAEDESTADDRREQ                  As Long = 10039    'A required address was omitted from an operation on a socket.
Public Const ERROR_WSAEMSGSIZE                      As Long = 10040    'A message sent on a datagram socket was larger than the internal message buffer or some other network limit, or the buffer used to receive a datagram into was smaller than the datagram itself.
Public Const ERROR_WSAEPROTOTYPE                    As Long = 10041    'A protocol was specified in the socket function call that does not support the semantics of the socket type requested.
Public Const ERROR_WSAENOPROTOOPT                   As Long = 10042    'An unknown, invalid, or unsupported option or level was specified in a getsockopt or setsockopt call.
Public Const ERROR_WSAEPROTONOSUPPORT               As Long = 10043    'The requested protocol has not been configured into the system, or no implementation for it exists.
Public Const ERROR_WSAESOCKTNOSUPPORT               As Long = 10044    'The support for the specified socket type does not exist in this address family.
Public Const ERROR_WSAEOPNOTSUPP                    As Long = 10045    'The attempted operation is not supported for the type of object referenced.
Public Const ERROR_WSAEPFNOSUPPORT                  As Long = 10046    'The protocol family has not been configured into the system or no implementation for it exists.
Public Const ERROR_WSAEAFNOSUPPORT                  As Long = 10047    'An address incompatible with the requested protocol was used.
Public Const ERROR_WSAEADDRINUSE                    As Long = 10048    'Only one usage of each socket address (protocol/network address/port) is normally permitted.
Public Const ERROR_WSAEADDRNOTAVAIL                 As Long = 10049    'The requested address is not valid in its context.
Public Const ERROR_WSAENETDOWN                      As Long = 10050    'A socket operation encountered a dead network.
Public Const ERROR_WSAENETUNREACH                   As Long = 10051    'A socket operation was attempted to an unreachable network.
Public Const ERROR_WSAENETRESET                     As Long = 10052    'The connection has been broken due to keep-alive activity detecting a failure while the operation was in progress.
Public Const ERROR_WSAECONNABORTED                  As Long = 10053    'An established connection was aborted by the software in your host machine.
Public Const ERROR_WSAECONNRESET                    As Long = 10054    'An existing connection was forcibly closed by the remote host.
Public Const ERROR_WSAENOBUFS                       As Long = 10055    'An operation on a socket could not be performed because the system lacked sufficient buffer space or because a queue was full.
Public Const ERROR_WSAEISCONN                       As Long = 10056    'A connect request was made on an already connected socket.
Public Const ERROR_WSAENOTCONN                      As Long = 10057    'A request to send or receive data was disallowed because the socket is not connected and (when sending on a datagram socket using a sendto call) no address was supplied.
Public Const ERROR_WSAESHUTDOWN                     As Long = 10058    'A request to send or receive data was disallowed because the socket had already been shut down in that direction with a previous shutdown call.
Public Const ERROR_WSAETOOMANYREFS                  As Long = 10059    'Too many references to some kernel object.
Public Const ERROR_WSAETIMEDOUT                     As Long = 10060    'A connection attempt failed because the connected party did not properly respond after a period of time, or established connection failed because connected host has failed to respond.
Public Const ERROR_WSAECONNREFUSED                  As Long = 10061    'No connection could be made because the target machine actively refused it.
Public Const ERROR_WSAELOOP                         As Long = 10062    'Cannot translate name.
Public Const ERROR_WSAENAMETOOLONG                  As Long = 10063    'Name component or name was too long.
Public Const ERROR_WSAEHOSTDOWN                     As Long = 10064    'A socket operation failed because the destination host was down.
Public Const ERROR_WSAEHOSTUNREACH                  As Long = 10065    'A socket operation was attempted to an unreachable host.
Public Const ERROR_WSAENOTEMPTY                     As Long = 10066    'Cannot remove a directory that is not empty.
Public Const ERROR_WSAEPROCLIM                      As Long = 10067    'A Windows Sockets implementation may have a limit on the number of applications that may use it simultaneously.
Public Const ERROR_WSAEUSERS                        As Long = 10068    'Ran out of quota.
Public Const ERROR_WSAEDQUOT                        As Long = 10069    'Ran out of disk quota.
Public Const ERROR_WSAESTALE                        As Long = 10070    'File handle reference is no longer available.
Public Const ERROR_WSAEREMOTE                       As Long = 10071    'Item is not available locally.
Public Const ERROR_WSASYSNOTREADY                   As Long = 10091    'WSAStartup cannot function at this time because the underlying system it uses to provide network services is currently unavailable.
Public Const ERROR_WSAVERNOTSUPPORTED               As Long = 10092    'The Windows Sockets version requested is not supported.
Public Const ERROR_WSANOTINITIALISED                As Long = 10093    'Either the application has not called WSAStartup, or WSAStartup failed.
Public Const ERROR_WSAEDISCON                       As Long = 10101    'Returned by WSARecv or WSARecvFrom to indicate the remote party has initiated a graceful shutdown sequence.
Public Const ERROR_WSAENOMORE                       As Long = 10102    'No more results can be returned by WSALookupServiceNext.
Public Const ERROR_WSAECANCELLED                    As Long = 10103    'A call to WSALookupServiceEnd was made while this call was still processing. The call has been canceled.
Public Const ERROR_WSAEINVALIDPROCTABLE             As Long = 10104    'The procedure call table is invalid.
Public Const ERROR_WSAEINVALIDPROVIDER              As Long = 10105    'The requested service provider is invalid.
Public Const ERROR_WSAEPROVIDERFAILEDINIT           As Long = 10106    'The requested service provider could not be loaded or initialized.
Public Const ERROR_WSASYSCALLFAILURE                As Long = 10107    'A system call that should never fail has failed.
Public Const ERROR_WSASERVICE_NOT_FOUND             As Long = 10108    'No such service is known. The service cannot be found in the specified name space.
Public Const ERROR_WSATYPE_NOT_FOUND                As Long = 10109    'The specified class was not found.
Public Const ERROR_WSA_E_NO_MORE                    As Long = 10110    'No more results can be returned by WSALookupServiceNext.
Public Const ERROR_WSA_E_CANCELLED                  As Long = 10111    'A call to WSALookupServiceEnd was made while this call was still processing. The call has been canceled.
Public Const ERROR_WSAEREFUSED                      As Long = 10112    'A database query failed because it was actively refused.
Public Const ERROR_WSAHOST_NOT_FOUND                As Long = 11001    'No such host is known.
Public Const ERROR_WSATRY_AGAIN                     As Long = 11002    'This is usually a temporary error during hostname resolution and means that the local server did not receive a response from an authoritative server.
Public Const ERROR_WSANO_RECOVERY                   As Long = 11003    'A non-recoverable error occurred during a database lookup.
Public Const ERROR_WSANO_DATA                       As Long = 11004    'The requested name is valid, but no data of the requested type was found.
Public Const ERROR_WSA_QOS_RECEIVERS                As Long = 11005    'At least one reserve has arrived.
Public Const ERROR_WSA_QOS_SENDERS                  As Long = 11006    'At least one path has arrived.
Public Const ERROR_WSA_QOS_NO_SENDERS               As Long = 11007    'There are no senders.
Public Const ERROR_WSA_QOS_NO_RECEIVERS             As Long = 11008    'There are no receivers.
Public Const ERROR_WSA_QOS_REQUEST_CONFIRMED        As Long = 11009    'Reserve has been confirmed.
Public Const ERROR_WSA_QOS_ADMISSION_FAILURE        As Long = 11010    'Error due to lack of resources.
Public Const ERROR_WSA_QOS_POLICY_FAILURE           As Long = 11011    'Rejected for administrative reasons - bad credentials.
Public Const ERROR_WSA_QOS_BAD_STYLE                As Long = 11012    'Unknown or conflicting style.
Public Const ERROR_WSA_QOS_BAD_OBJECT               As Long = 11013    'Problem with some part of the filterspec or providerspecific buffer in general.
Public Const ERROR_WSA_QOS_TRAFFIC_CTRL_ERROR       As Long = 11014    'Problem with some part of the flowspec.
Public Const ERROR_WSA_QOS_GENERIC_ERROR            As Long = 11015    'General QOS error.
Public Const ERROR_WSA_QOS_ESERVICETYPE             As Long = 11016    'An invalid or unrecognized service type was found in the flowspec.
Public Const ERROR_WSA_QOS_EFLOWSPEC                As Long = 11017    'An invalid or inconsistent flowspec was found in the QOS structure.
Public Const ERROR_WSA_QOS_EPROVSPECBUF             As Long = 11018    'Invalid QOS provider-specific buffer.
Public Const ERROR_WSA_QOS_EFILTERSTYLE             As Long = 11019    'An invalid QOS filter style was used.
Public Const ERROR_WSA_QOS_EFILTERTYPE              As Long = 11020    'An invalid QOS filter type was used.
Public Const ERROR_WSA_QOS_EFILTERCOUNT             As Long = 11021    'An incorrect number of QOS FILTERSPECs were specified in the FLOWDESCRIPTOR.
Public Const ERROR_WSA_QOS_EOBJLENGTH               As Long = 11022    'An object with an invalid ObjectLength field was specified in the QOS provider-specific buffer.
Public Const ERROR_WSA_QOS_EFLOWCOUNT               As Long = 11023    'An incorrect number of flow descriptors was specified in the QOS structure.
Public Const ERROR_WSA_QOS_EUNKOWNPSOBJ             As Long = 11024    'An unrecognized object was found in the QOS provider-specific buffer.
Public Const ERROR_WSA_QOS_EPOLICYOBJ               As Long = 11025    'An invalid policy object was found in the QOS provider-specific buffer.
Public Const ERROR_WSA_QOS_EFLOWDESC                As Long = 11026    'An invalid QOS flow descriptor was found in the flow descriptor list.
Public Const ERROR_WSA_QOS_EPSFLOWSPEC              As Long = 11027    'An invalid or inconsistent flowspec was found in the QOS provider specific buffer.
Public Const ERROR_WSA_QOS_EPSFILTERSPEC            As Long = 11028    'An invalid FILTERSPEC was found in the QOS provider-specific buffer.
Public Const ERROR_WSA_QOS_ESDMODEOBJ               As Long = 11029    'An invalid shape discard mode object was found in the QOS provider specific buffer.
Public Const ERROR_WSA_QOS_ESHAPERATEOBJ            As Long = 11030    'An invalid shaping rate object was found in the QOS provider-specific buffer.
Public Const ERROR_WSA_QOS_RESERVED_PETYPE          As Long = 11031    'A reserved policy element was found in the QOS provider-specific buffer.

'
'#endif // defined(WSABASEERR)
'
'///////////////////////////////////////////////////
'//                                               //
'//           End of WinSock Error Codes          //
'//                                               //
'//                 10000 to 11999                //
'///////////////////////////////////////////////////
'
'
'
'///////////////////////////////////////////////////
'//                                               //
'//             Side By Side Error Codes          //
'//                                               //
'//                 14000 to 14999                //
'///////////////////////////////////////////////////
Public Const ERROR_SXS_SECTION_NOT_FOUND      As Long = 14000    'The requested section was not present in the activation context.
Public Const ERROR_SXS_CANT_GEN_ACTCTX        As Long = 14001    'This application has failed to start because the application configuration is incorrect. Reinstalling the application may fix this problem.
Public Const ERROR_SXS_INVALID_ACTCTXDATA_FORMAT As Long = 14002 'The application binding data format is invalid.
Public Const ERROR_SXS_ASSEMBLY_NOT_FOUND     As Long = 14003    'The referenced assembly is not installed on your system.
Public Const ERROR_SXS_MANIFEST_FORMAT_ERROR  As Long = 14004    'The manifest file does not begin with the required tag and format information.
Public Const ERROR_SXS_MANIFEST_PARSE_ERROR   As Long = 14005    'The manifest file contains one or more syntax errors.
Public Const ERROR_SXS_ACTIVATION_CONTEXT_DISABLED As Long = 14006 'The application attempted to activate a disabled activation context.
Public Const ERROR_SXS_KEY_NOT_FOUND          As Long = 14007    'The requested lookup key was not found in any active activation context.
Public Const ERROR_SXS_VERSION_CONFLICT       As Long = 14008    'A component version required by the application conflicts with another component version already active.
Public Const ERROR_SXS_WRONG_SECTION_TYPE     As Long = 14009    'The type requested activation context section does not match the query API used.
Public Const ERROR_SXS_THREAD_QUERIES_DISABLED As Long = 14010   'Lack of system resources has required isolated activation to be disabled for the current thread of execution.
Public Const ERROR_SXS_PROCESS_DEFAULT_ALREADY_SET As Long = 14011 'An attempt to set the process default activation context failed because the process default activation context was already set.
Public Const ERROR_SXS_UNKNOWN_ENCODING_GROUP As Long = 14012    'The encoding group identifier specified is not recognized.
Public Const ERROR_SXS_UNKNOWN_ENCODING       As Long = 14013    'The encoding requested is not recognized.
Public Const ERROR_SXS_INVALID_XML_NAMESPACE_URI As Long = 14014 'The manifest contains a reference to an invalid URI.
Public Const ERROR_SXS_ROOT_MANIFEST_DEPENDENCY_NOT_INSTALLED As Long = 14015 'The application manifest contains a reference to a dependent assembly which is not installed
Public Const ERROR_SXS_LEAF_MANIFEST_DEPENDENCY_NOT_INSTALLED As Long = 14016 'The manifest for an assembly used by the application has a reference to a dependent assembly which is not installed
Public Const ERROR_SXS_INVALID_ASSEMBLY_IDENTITY_ATTRIBUTE As Long = 14017 'The manifest contains an attribute for the assembly identity which is not valid.
Public Const ERROR_SXS_MANIFEST_MISSING_REQUIRED_DEFAULT_NAMESPACE As Long = 14018 'The manifest is missing the required default namespace specification on the assembly element.
Public Const ERROR_SXS_MANIFEST_INVALID_REQUIRED_DEFAULT_NAMESPACE As Long = 14019 'The manifest has a default namespace specified on the assembly element but its value is not "urn:schemas-microsoft-com:asm.v1".
Public Const ERROR_SXS_PRIVATE_MANIFEST_CROSS_PATH_WITH_REPARSE_POINT As Long = 14020 'The private manifest probed has crossed reparse-point-associated path
Public Const ERROR_SXS_DUPLICATE_DLL_NAME     As Long = 14021    'Two or more components referenced directly or indirectly by the application manifest have files by the same name.
Public Const ERROR_SXS_DUPLICATE_WINDOWCLASS_NAME As Long = 14022 'Two or more components referenced directly or indirectly by the application manifest have window classes with the same name.
Public Const ERROR_SXS_DUPLICATE_CLSID        As Long = 14023    'Two or more components referenced directly or indirectly by the application manifest have the same COM server CLSIDs.
Public Const ERROR_SXS_DUPLICATE_IID          As Long = 14024    'Two or more components referenced directly or indirectly by the application manifest have proxies for the same COM interface IIDs.
Public Const ERROR_SXS_DUPLICATE_TLBID        As Long = 14025    'Two or more components referenced directly or indirectly by the application manifest have the same COM type library TLBIDs.
Public Const ERROR_SXS_DUPLICATE_PROGID       As Long = 14026    'Two or more components referenced directly or indirectly by the application manifest have the same COM ProgIDs.
Public Const ERROR_SXS_DUPLICATE_ASSEMBLY_NAME As Long = 14027   'Two or more components referenced directly or indirectly by the application manifest are different versions of the same component which is not permitted.
Public Const ERROR_SXS_FILE_HASH_MISMATCH     As Long = 14028    'A component's file does not match the verification information present in the    'component manifest.
Public Const ERROR_SXS_POLICY_PARSE_ERROR     As Long = 14029    'The policy manifest contains one or more syntax errors.
Public Const ERROR_SXS_XML_E_MISSINGQUOTE     As Long = 14030    'Manifest Parse Error : A string literal was expected, but no opening quote character was found.
Public Const ERROR_SXS_XML_E_COMMENTSYNTAX    As Long = 14031    'Manifest Parse Error : Incorrect syntax was used in a comment.
Public Const ERROR_SXS_XML_E_BADSTARTNAMECHAR As Long = 14032    'Manifest Parse Error : A name was started with an invalid character.
Public Const ERROR_SXS_XML_E_BADNAMECHAR      As Long = 14033    'Manifest Parse Error : A name contained an invalid character.
Public Const ERROR_SXS_XML_E_BADCHARINSTRING  As Long = 14034    'Manifest Parse Error : A string literal contained an invalid character.
Public Const ERROR_SXS_XML_E_XMLDECLSYNTAX    As Long = 14035    'Manifest Parse Error : Invalid syntax for an xml declaration.
Public Const ERROR_SXS_XML_E_BADCHARDATA      As Long = 14036    'Manifest Parse Error : An Invalid character was found in text content.
Public Const ERROR_SXS_XML_E_MISSINGWHITESPACE As Long = 14037   'Manifest Parse Error : Required white space was missing.
Public Const ERROR_SXS_XML_E_EXPECTINGTAGEND  As Long = 14038    'Manifest Parse Error : The character '>' was expected.
Public Const ERROR_SXS_XML_E_MISSINGSEMICOLON As Long = 14039    'Manifest Parse Error : A semi colon character was expected.
Public Const ERROR_SXS_XML_E_UNBALANCEDPAREN  As Long = 14040    'Manifest Parse Error : Unbalanced parentheses.
Public Const ERROR_SXS_XML_E_INTERNALERROR    As Long = 14041    'Manifest Parse Error : Internal error.
Public Const ERROR_SXS_XML_E_UNEXPECTED_WHITESPACE As Long = 14042 'Manifest Parse Error : Whitespace is not allowed at this location.
Public Const ERROR_SXS_XML_E_INCOMPLETE_ENCODING As Long = 14043 'Manifest Parse Error : End of file reached in invalid state for current encoding.
Public Const ERROR_SXS_XML_E_MISSING_PAREN    As Long = 14044    'Manifest Parse Error : Missing parenthesis.
Public Const ERROR_SXS_XML_E_EXPECTINGCLOSEQUOTE As Long = 14045 'Manifest Parse Error : A single or double closing quote character (\' or \") is missing.
Public Const ERROR_SXS_XML_E_MULTIPLE_COLONS  As Long = 14046    'Manifest Parse Error : Multiple colons are not allowed in a name.
Public Const ERROR_SXS_XML_E_INVALID_DECIMAL  As Long = 14047    'Manifest Parse Error : Invalid character for decimal digit.
Public Const ERROR_SXS_XML_E_INVALID_HEXIDECIMAL As Long = 14048 'Manifest Parse Error : Invalid character for hexidecimal digit.
Public Const ERROR_SXS_XML_E_INVALID_UNICODE  As Long = 14049    'Manifest Parse Error : Invalid unicode character value for this platform.
Public Const ERROR_SXS_XML_E_WHITESPACEORQUESTIONMARK As Long = 14050 'Manifest Parse Error : Expecting whitespace or '?'.
Public Const ERROR_SXS_XML_E_UNEXPECTEDENDTAG As Long = 14051    'Manifest Parse Error : End tag was not expected at this location.
Public Const ERROR_SXS_XML_E_UNCLOSEDTAG      As Long = 14052    'Manifest Parse Error : The following tags were not closed: %1.
Public Const ERROR_SXS_XML_E_DUPLICATEATTRIBUTE As Long = 14053  'Manifest Parse Error : Duplicate attribute.
Public Const ERROR_SXS_XML_E_MULTIPLEROOTS    As Long = 14054    'Manifest Parse Error : Only one top level element is allowed in an XML document.
Public Const ERROR_SXS_XML_E_INVALIDATROOTLEVEL As Long = 14055  'Manifest Parse Error : Invalid at the top level of the document.
Public Const ERROR_SXS_XML_E_BADXMLDECL       As Long = 14056    'Manifest Parse Error : Invalid xml declaration.
Public Const ERROR_SXS_XML_E_MISSINGROOT      As Long = 14057    'Manifest Parse Error : XML document must have a top level element.
Public Const ERROR_SXS_XML_E_UNEXPECTEDEOF    As Long = 14058    'Manifest Parse Error : Unexpected end of file.
Public Const ERROR_SXS_XML_E_BADPEREFINSUBSET As Long = 14059    'Manifest Parse Error : Parameter entities cannot be used inside markup declarations in an internal subset.
Public Const ERROR_SXS_XML_E_UNCLOSEDSTARTTAG As Long = 14060    'Manifest Parse Error : Element was not closed.
Public Const ERROR_SXS_XML_E_UNCLOSEDENDTAG   As Long = 14061    'Manifest Parse Error : End element was missing the character '>'.
Public Const ERROR_SXS_XML_E_UNCLOSEDSTRING   As Long = 14062    'Manifest Parse Error : A string literal was not closed.
Public Const ERROR_SXS_XML_E_UNCLOSEDCOMMENT  As Long = 14063    'Manifest Parse Error : A comment was not closed.
Public Const ERROR_SXS_XML_E_UNCLOSEDDECL     As Long = 14064    'Manifest Parse Error : A declaration was not closed.
Public Const ERROR_SXS_XML_E_UNCLOSEDCDATA    As Long = 14065    'Manifest Parse Error : A CDATA section was not closed.
Public Const ERROR_SXS_XML_E_RESERVEDNAMESPACE As Long = 14066   'Manifest Parse Error : The namespace prefix is not allowed to start with the reserved string "xml".
Public Const ERROR_SXS_XML_E_INVALIDENCODING  As Long = 14067    'Manifest Parse Error : System does not support the specified encoding.
Public Const ERROR_SXS_XML_E_INVALIDSWITCH    As Long = 14068    'Manifest Parse Error : Switch from current encoding to specified encoding not supported.
Public Const ERROR_SXS_XML_E_BADXMLCASE       As Long = 14069    'Manifest Parse Error : The name 'xml' is reserved and must be lower case.
Public Const ERROR_SXS_XML_E_INVALID_STANDALONE As Long = 14070  'Manifest Parse Error : The standalone attribute must have the value 'yes' or 'no'.
Public Const ERROR_SXS_XML_E_UNEXPECTED_STANDALONE As Long = 14071 'Manifest Parse Error : The standalone attribute cannot be used in external entities.
Public Const ERROR_SXS_XML_E_INVALID_VERSION  As Long = 14072    'Manifest Parse Error : Invalid version number.
Public Const ERROR_SXS_XML_E_MISSINGEQUALS    As Long = 14073    'Manifest Parse Error : Missing equals sign between attribute and attribute value.
Public Const ERROR_SXS_PROTECTION_RECOVERY_FAILED As Long = 14074 'Assembly Protection Error : Unable to recover the specified assembly.
Public Const ERROR_SXS_PROTECTION_PUBLIC_KEY_TOO_SHORT As Long = 14075 'Assembly Protection Error : The public key for an assembly was too short to be allowed.
Public Const ERROR_SXS_PROTECTION_CATALOG_NOT_VALID As Long = 14076 'Assembly Protection Error : The catalog for an assembly is not valid, or does not match the assembly's manifest.
Public Const ERROR_SXS_UNTRANSLATABLE_HRESULT As Long = 14077    'An HRESULT could not be translated to a corresponding Win32 error code.
Public Const ERROR_SXS_PROTECTION_CATALOG_FILE_MISSING As Long = 14078 'Assembly Protection Error : The catalog for an assembly is missing.
Public Const ERROR_SXS_MISSING_ASSEMBLY_IDENTITY_ATTRIBUTE As Long = 14079 'The supplied assembly identity is missing one or more attributes which must be present in this context.
Public Const ERROR_SXS_INVALID_ASSEMBLY_IDENTITY_ATTRIBUTE_NAME As Long = 14080 'The supplied assembly identity has one or more attribute names that contain characters not permitted in XML names.

'
'
'///////////////////////////////////////////////////
'//                                               //
'//           End of Side By Side Error Codes     //
'//                                               //
'//                 14000 to 14999                //
'///////////////////////////////////////////////////
'
'
'
'///////////////////////////////////////////////////
'//                                               //
'//           Start of IPSec Error codes          //
'//                                               //
'//                 13000 to 13999                //
'///////////////////////////////////////////////////
'
Public Const ERROR_IPSEC_QM_POLICY_EXISTS     As Long = 13000    'The specified quick mode policy already exists.
Public Const ERROR_IPSEC_QM_POLICY_NOT_FOUND  As Long = 13001    'The specified quick mode policy was not found.
Public Const ERROR_IPSEC_QM_POLICY_IN_USE     As Long = 13002    'The specified quick mode policy is being used.
Public Const ERROR_IPSEC_MM_POLICY_EXISTS     As Long = 13003    'The specified main mode policy already exists.
Public Const ERROR_IPSEC_MM_POLICY_NOT_FOUND  As Long = 13004    'The specified main mode policy was not found
Public Const ERROR_IPSEC_MM_POLICY_IN_USE     As Long = 13005    'The specified main mode policy is being used.
Public Const ERROR_IPSEC_MM_FILTER_EXISTS     As Long = 13006    'The specified main mode filter already exists.
Public Const ERROR_IPSEC_MM_FILTER_NOT_FOUND  As Long = 13007    'The specified main mode filter was not found.
Public Const ERROR_IPSEC_TRANSPORT_FILTER_EXISTS As Long = 13008 'The specified transport mode filter already exists.
Public Const ERROR_IPSEC_TRANSPORT_FILTER_NOT_FOUND As Long = 13009 'The specified transport mode filter does not exist.
Public Const ERROR_IPSEC_MM_AUTH_EXISTS       As Long = 13010    'The specified main mode authentication list exists.
Public Const ERROR_IPSEC_MM_AUTH_NOT_FOUND    As Long = 13011    'The specified main mode authentication list was not found.
Public Const ERROR_IPSEC_MM_AUTH_IN_USE       As Long = 13012    'The specified quick mode policy is being used.
Public Const ERROR_IPSEC_DEFAULT_MM_POLICY_NOT_FOUND As Long = 13013 'The specified main mode policy was not found.
Public Const ERROR_IPSEC_DEFAULT_MM_AUTH_NOT_FOUND As Long = 13014 'The specified quick mode policy was not found
Public Const ERROR_IPSEC_DEFAULT_QM_POLICY_NOT_FOUND As Long = 13015 'The manifest file contains one or more syntax errors.
Public Const ERROR_IPSEC_TUNNEL_FILTER_EXISTS As Long = 13016    'The application attempted to activate a disabled activation context.
Public Const ERROR_IPSEC_TUNNEL_FILTER_NOT_FOUND As Long = 13017 'The requested lookup key was not found in any active activation context.
Public Const ERROR_IPSEC_MM_FILTER_PENDING_DELETION As Long = 13018 'The Main Mode filter is pending deletion.
Public Const ERROR_IPSEC_TRANSPORT_FILTER_PENDING_DELETION As Long = 13019    'The transport filter is pending deletion.
Public Const ERROR_IPSEC_TUNNEL_FILTER_PENDING_DELETION As Long = 13020    'The tunnel filter is pending deletion.
Public Const ERROR_IPSEC_MM_POLICY_PENDING_DELETION As Long = 13021 'The Main Mode policy is pending deletion.
Public Const ERROR_IPSEC_MM_AUTH_PENDING_DELETION As Long = 13022 'The Main Mode authentication bundle is pending deletion.
Public Const ERROR_IPSEC_QM_POLICY_PENDING_DELETION As Long = 13023 'The Quick Mode policy is pending deletion.
Public Const WARNING_IPSEC_MM_POLICY_PRUNED   As Long = 13024    'The Main Mode policy was successfully added, but some of the requested offers are not supported.
Public Const WARNING_IPSEC_QM_POLICY_PRUNED   As Long = 13025    'The Quick Mode policy was successfully added, but some of the requested offers are not supported.
Public Const ERROR_IPSEC_IKE_NEG_STATUS_BEGIN As Long = 13800    'ERROR_IPSEC_IKE_NEG_STATUS_BEGIN
Public Const ERROR_IPSEC_IKE_AUTH_FAIL        As Long = 13801    'IKE authentication credentials are unacceptable
Public Const ERROR_IPSEC_IKE_ATTRIB_FAIL      As Long = 13802    'IKE security attributes are unacceptable
Public Const ERROR_IPSEC_IKE_NEGOTIATION_PENDING As Long = 13803 'IKE Negotiation in progress
Public Const ERROR_IPSEC_IKE_GENERAL_PROCESSING_ERROR As Long = 13804    'General processing error
Public Const ERROR_IPSEC_IKE_TIMED_OUT        As Long = 13805    'Negotiation timed out
Public Const ERROR_IPSEC_IKE_NO_CERT          As Long = 13806    'IKE failed to find valid machine certificate
Public Const ERROR_IPSEC_IKE_SA_DELETED       As Long = 13807    'IKE SA deleted by peer before establishment completed
Public Const ERROR_IPSEC_IKE_SA_REAPED        As Long = 13808    'IKE SA deleted before establishment completed
Public Const ERROR_IPSEC_IKE_MM_ACQUIRE_DROP  As Long = 13809    'Negotiation request sat in Queue too long
Public Const ERROR_IPSEC_IKE_QM_ACQUIRE_DROP  As Long = 13810    'Negotiation request sat in Queue too long
Public Const ERROR_IPSEC_IKE_QUEUE_DROP_MM    As Long = 13811    'Negotiation request sat in Queue too long
Public Const ERROR_IPSEC_IKE_QUEUE_DROP_NO_MM As Long = 13812    'Negotiation request sat in Queue too long
Public Const ERROR_IPSEC_IKE_DROP_NO_RESPONSE As Long = 13813    'No response from peer
Public Const ERROR_IPSEC_IKE_MM_DELAY_DROP    As Long = 13814    'Negotiation took too long
Public Const ERROR_IPSEC_IKE_QM_DELAY_DROP    As Long = 13815    'Negotiation took too long
Public Const ERROR_IPSEC_IKE_ERROR            As Long = 13816    'Unknown error occurred
Public Const ERROR_IPSEC_IKE_CRL_FAILED       As Long = 13817    'Certificate Revocation Check failed
Public Const ERROR_IPSEC_IKE_INVALID_KEY_USAGE As Long = 13818   'Invalid certificate key usage
Public Const ERROR_IPSEC_IKE_INVALID_CERT_TYPE As Long = 13819   'Invalid certificate type
Public Const ERROR_IPSEC_IKE_NO_PRIVATE_KEY   As Long = 13820    'No private key associated with machine certificate
Public Const ERROR_IPSEC_IKE_DH_FAIL          As Long = 13822    'Failure in Diffie-Helman computation
Public Const ERROR_IPSEC_IKE_INVALID_HEADER   As Long = 13824    'Invalid header
Public Const ERROR_IPSEC_IKE_NO_POLICY        As Long = 13825    'No policy configured
Public Const ERROR_IPSEC_IKE_INVALID_SIGNATURE As Long = 13826   'Failed to verify signature
Public Const ERROR_IPSEC_IKE_KERBEROS_ERROR   As Long = 13827    'Failed to authenticate using kerberos
Public Const ERROR_IPSEC_IKE_NO_PUBLIC_KEY    As Long = 13828    'Peer's certificate did not have a public key
Public Const ERROR_IPSEC_IKE_PROCESS_ERR      As Long = 13829    'Error processing error payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_SA   As Long = 13830    'Error processing SA payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_PROP As Long = 13831    'Error processing Proposal payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_TRANS As Long = 13832   'Error processing Transform payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_KE   As Long = 13833    'Error processing KE payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_ID   As Long = 13834    'Error processing ID payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_CERT As Long = 13835    'Error processing Cert payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_CERT_REQ As Long = 13836 'Error processing Certificate Request payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_HASH As Long = 13837    'Error processing Hash payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_SIG  As Long = 13838    'Error processing Signature payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_NONCE As Long = 13839   'Error processing Nonce payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_NOTIFY As Long = 13840  'Error processing Notify payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_DELETE As Long = 13841  'Error processing Delete Payload
Public Const ERROR_IPSEC_IKE_PROCESS_ERR_VENDOR As Long = 13842  'Error processing VendorId payload
Public Const ERROR_IPSEC_IKE_INVALID_PAYLOAD  As Long = 13843    'Invalid payload received
Public Const ERROR_IPSEC_IKE_LOAD_SOFT_SA     As Long = 13844    'Soft SA loaded
Public Const ERROR_IPSEC_IKE_SOFT_SA_TORN_DOWN As Long = 13845   'Soft SA torn down
Public Const ERROR_IPSEC_IKE_INVALID_COOKIE   As Long = 13846    'Invalid cookie received.
Public Const ERROR_IPSEC_IKE_NO_PEER_CERT     As Long = 13847    'Peer failed to send valid machine certificate
Public Const ERROR_IPSEC_IKE_PEER_CRL_FAILED  As Long = 13848    'Certification Revocation check of peer's certificate failed
Public Const ERROR_IPSEC_IKE_POLICY_CHANGE    As Long = 13849    'New policy invalidated SAs formed with old policy
Public Const ERROR_IPSEC_IKE_NO_MM_POLICY     As Long = 13850    'There is no available Main Mode IKE policy.
Public Const ERROR_IPSEC_IKE_NOTCBPRIV        As Long = 13851    'Failed to enabled TCB privilege.
Public Const ERROR_IPSEC_IKE_SECLOADFAIL      As Long = 13852    'Failed to load SECURITY.DLL.
Public Const ERROR_IPSEC_IKE_FAILSSPINIT      As Long = 13853    'Failed to obtain security function table dispatch address from SSPI.
Public Const ERROR_IPSEC_IKE_FAILQUERYSSP     As Long = 13854    'Failed to query Kerberos package to obtain max token size.
Public Const ERROR_IPSEC_IKE_SRVACQFAIL       As Long = 13855    'Failed to obtain Kerberos server credentials for ISAKMP/ERROR_IPSEC_IKE service.  Kerberos authentication will not function.  The most likely reason for this is lack of domain membership.  This is normal if your computer is a member of a workgroup.
Public Const ERROR_IPSEC_IKE_SRVQUERYCRED     As Long = 13856    'Failed to determine SSPI principal name for ISAKMP/ERROR_IPSEC_IKE service (QueryCredentialsAttributes).
Public Const ERROR_IPSEC_IKE_GETSPIFAIL       As Long = 13857    'Failed to obtain new SPI for the inbound SA from Ipsec driver.  The most common cause for this is that the driver does not have the correct filter.  Check your policy to verify the filters.
Public Const ERROR_IPSEC_IKE_INVALID_FILTER   As Long = 13858    'Given filter is invalid
Public Const ERROR_IPSEC_IKE_OUT_OF_MEMORY    As Long = 13859    'Memory allocation failed.
Public Const ERROR_IPSEC_IKE_ADD_UPDATE_KEY_FAILED As Long = 13860 'Failed to add Security Association to IPSec Driver.  The most common cause for this is if the IKE negotiation took too long to complete.  If the problem persists, reduce the load on the faulting machine.
Public Const ERROR_IPSEC_IKE_INVALID_POLICY   As Long = 13861    'Invalid policy
Public Const ERROR_IPSEC_IKE_UNKNOWN_DOI      As Long = 13862    'Invalid DOI
Public Const ERROR_IPSEC_IKE_INVALID_SITUATION As Long = 13863   'Invalid situation
Public Const ERROR_IPSEC_IKE_DH_FAILURE       As Long = 13864    'Diffie-Hellman failure
Public Const ERROR_IPSEC_IKE_INVALID_GROUP    As Long = 13865    'Invalid Diffie-Hellman group
Public Const ERROR_IPSEC_IKE_ENCRYPT          As Long = 13866    'Error encrypting payload
Public Const ERROR_IPSEC_IKE_DECRYPT          As Long = 13867    'Error decrypting payload
Public Const ERROR_IPSEC_IKE_POLICY_MATCH     As Long = 13868    'Policy match error
Public Const ERROR_IPSEC_IKE_UNSUPPORTED_ID   As Long = 13869    'Unsupported ID
Public Const ERROR_IPSEC_IKE_INVALID_HASH     As Long = 13870    'Hash verification failed
Public Const ERROR_IPSEC_IKE_INVALID_HASH_ALG As Long = 13871    'Invalid hash algorithm
Public Const ERROR_IPSEC_IKE_INVALID_HASH_SIZE As Long = 13872   'Invalid hash size
Public Const ERROR_IPSEC_IKE_INVALID_ENCRYPT_ALG As Long = 13873 'Invalid encryption algorithm
Public Const ERROR_IPSEC_IKE_INVALID_AUTH_ALG As Long = 13874    'Invalid authentication algorithm
Public Const ERROR_IPSEC_IKE_INVALID_SIG      As Long = 13875    'Invalid certificate signature
Public Const ERROR_IPSEC_IKE_LOAD_FAILED      As Long = 13876    'Load failed
Public Const ERROR_IPSEC_IKE_RPC_DELETE       As Long = 13877    'Deleted via RPC call
Public Const ERROR_IPSEC_IKE_BENIGN_REINIT    As Long = 13878    'Temporary state created to perform reinit. This is not a real failure.
Public Const ERROR_IPSEC_IKE_INVALID_RESPONDER_LIFETIME_NOTIFY As Long = 13879    'The lifetime value received in the Responder Lifetime Notify is below the Windows 2000 configured minimum value.  Please fix the policy on the peer machine.
Public Const ERROR_IPSEC_IKE_INVALID_CERT_KEYLEN As Long = 13881 'Key length in certificate is too small for configured security requirements.
Public Const ERROR_IPSEC_IKE_MM_LIMIT         As Long = 13882    'Max number of established MM SAs to peer exceeded.
Public Const ERROR_IPSEC_IKE_NEGOTIATION_DISABLED As Long = 13883 'IKE received a policy that disables negotiation.
Public Const ERROR_IPSEC_IKE_NEG_STATUS_END As Long = 13884 'ERROR_IPSEC_IKE_NEG_STATUS_END 'ERROR_IPSEC_IKE_NEG_STATUS_END
'
'////////////////////////////////////
'//                                //
'//     COM Error Codes            //
'//                                //
'////////////////////////////////////
'
'//
'// The return value of COM functions and methods is an HRESULT.
'// This is not a handle to anything, but is merely a 32-bit value
'// with several fields encoded in the value.  The parts of an
'// HRESULT are shown below.
'//
'// Many of the macros and functions below were orginally defined to
'// operate on SCODEs.  SCODEs are no longer used.  The macros are
'// still present for compatibility and easy porting of Win16 code.
'// Newly written code should use the HRESULT macros and functions.
'//
'
'//
'//  HRESULTs are 32 bit values layed out as follows:
'//
'//   3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
'//   1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
'//  +-+-+-+-+-+---------------------+-------------------------------+
'//  |S|R|C|N|r|    Facility         |               Code            |
'//  +-+-+-+-+-+---------------------+-------------------------------+
'//
'//  where
'//
'//      S - Severity - indicates success/fail
'//
'//          0 - Success
'//          1 - Fail (COERROR)
'//
'//      R - reserved portion of the facility code, corresponds to NT's
'//              second severity bit.
'//
'//      C - reserved portion of the facility code, corresponds to NT's
'//              C field.
'//
'//      N - reserved portion of the facility code. Used to indicate a
'//              mapped NT status value.
'//
'//      r - reserved portion of the facility code. Reserved for internal
'//              use. Used to indicate HRESULT values that are not status
'//              values, but are instead message ids for display strings.
'//
'//      Facility - is the facility code
'//
'//      Code - is the facility's status code
'//
'
'//
'// Severity values
'//
'
Public Const SEVERITY_SUCCESS    As Long = 0
Public Const SEVERITY_ERROR      As Long = 1
'
'
'//
'// Generic test for success on any status value (non-negative numbers
'// indicate success).
'//
'
'Public Const SUCCEEDED(hr) ((HRESULT)(hr) >= 0)
'
'//
'// and the inverse
'//
'
'Public Const FAILED(hr) ((HRESULT)(hr) < 0)
'
'
'//
'// Generic test for error on any status value.
'//
'
'Public Const IS_ERROR(Status) ((unsigned long)(Status) >> 31 == SEVERITY_ERROR)
'
'//
'// Return the code
'//
'
'Public Const HRESULT_CODE(hr)    ((hr) & 0xFFFF)
'Public Const SCODE_CODE(sc)      ((sc) & 0xFFFF)
'
'//
'//  Return the facility
'//
'
'Public Const HRESULT_FACILITY(hr)  (((hr) >> 16) & 0x1fff)
'Public Const SCODE_FACILITY(sc)    (((sc) >> 16) & 0x1fff)
'
'//
'//  Return the severity
'//
'
'Public Const HRESULT_SEVERITY(hr)  (((hr) >> 31) & 0x1)
'Public Const SCODE_SEVERITY(sc)    (((sc) >> 31) & 0x1)
'
'//
'// Create an HRESULT value from component pieces
'//
'
'Public Const MAKE_HRESULT(sev,fac,code) \
'    ((HRESULT) (((unsigned long)(sev)<<31) | ((unsigned long)(fac)<<16) | ((unsigned long)(code))) )
'Public Const MAKE_SCODE(sev,fac,code) \
'    ((SCODE) (((unsigned long)(sev)<<31) | ((unsigned long)(fac)<<16) | ((unsigned long)(code))) )
'
'
'//
'// Map a WIN32 error value into a HRESULT
'// Note: This assumes that WIN32 errors fall in the range -32k to 32k.
'//
'// Define bits here so macros are guaranteed to work
'
'Public Const FACILITY_NT_BIT                 0x10000000
'
'// __HRESULT_FROM_WIN32 will always be a macro.
'// The goal will be to enable INLINE_HRESULT_FROM_WIN32 all the time,
'// but there's too much code to change to do that at this time.
'
'Public Const __HRESULT_FROM_WIN32(x) ((HRESULT)(x) <= 0 ? ((HRESULT)(x)) : ((HRESULT) (((x) & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000)))
'
'#ifdef INLINE_HRESULT_FROM_WIN32
'#ifndef _HRESULT_DEFINED
'Public Const _HRESULT_DEFINED
'typedef long HRESULT;
'#End If
'#ifndef __midl
'__inline HRESULT HRESULT_FROM_WIN32(long x) { return x <= 0 ? (HRESULT)x : (HRESULT) (((x) & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000);}
'#Else
'Public Const HRESULT_FROM_WIN32(x) __HRESULT_FROM_WIN32(x)
'#End If
'#Else
'Public Const HRESULT_FROM_WIN32(x) __HRESULT_FROM_WIN32(x)
'#End If
'
'//
'// Map an NT status value into a HRESULT
'//
'
'Public Const HRESULT_FROM_NT(x)      ((HRESULT) ((x) | FACILITY_NT_BIT))
'
'
'// ****** OBSOLETE functions
'
'// HRESULT functions
'// As noted above, these functions are obsolete and should not be used.
'
'
'// Extract the SCODE from a HRESULT
'
'Public Const GetScode(hr) ((SCODE) (hr))
'
'// Convert an SCODE into an HRESULT.
'
'Public Const ResultFromScode(sc) ((HRESULT) (sc))
'
'
'// PropagateResult is a noop
'Public Const PropagateResult(hrPrevious, scBase) ((HRESULT) scBase)
'
'
'// ****** End of OBSOLETE functions.
'
'
'// ---------------------- HRESULT value definitions -----------------
'//
'// HRESULT definitions
'//
'
'#ifdef RC_INVOKED
'Public Const _HRESULT_TYPEDEF_(_sc) _sc
'#else // RC_INVOKED
'Public Const _HRESULT_TYPEDEF_(_sc) ((HRESULT)_sc)
'#endif // RC_INVOKED
'
'Public Const NOERROR             0
'
'//
'// Error definitions follow
'//
'
'//
'// Codes 0x4000-0x40ff are reserved for OLE
'//
'//
'// Error codes
'//
Public Const ERROR_E_UNEXPECTED                                           As Long = &H8000FFFF 'Catastrophic failure 'E_UNEXPECTED
Public Const ERROR_E_NOTIMPL                                              As Long = &H80000001 'Not implemented 'E_NOTIMPL
Public Const ERROR_E_OUTOFMEMORY                                          As Long = &H80000002 'Ran out of memory 'E_OUTOFMEMORY
Public Const ERROR_E_INVALIDARG                                           As Long = &H80000003 'One or more arguments are invalid 'E_INVALIDARG
Public Const ERROR_E_NOINTERFACE                                          As Long = &H80000004 'No such interface supported 'E_NOINTERFACE
Public Const ERROR_E_POINTER                                              As Long = &H80000005 'Invalid pointer 'E_POINTER
Public Const ERROR_E_HANDLE                                               As Long = &H80000006 'Invalid handle 'E_HANDLE
Public Const ERROR_E_ABORT                                                As Long = &H80000007 'Operation aborted 'E_ABORT
Public Const ERROR_E_FAIL                                                 As Long = &H80000008 'Unspecified error 'E_FAIL
Public Const ERROR_E_ACCESSDENIED                                         As Long = &H80000009 'General access denied error 'E_ACCESSDENIED
Public Const ERROR_E_PENDING                                              As Long = &H8000000A 'The data necessary to complete this operation is not yet available. 'E_PENDING
Public Const ERROR_CO_E_INIT_TLS                                          As Long = &H80004006 'Thread local storage failure 'CO_E_INIT_TLS
Public Const ERROR_CO_E_INIT_SHARED_ALLOCATOR                             As Long = &H80004007 'Get shared memory allocator failure 'CO_E_INIT_SHARED_ALLOCATOR
Public Const ERROR_CO_E_INIT_MEMORY_ALLOCATOR                             As Long = &H80004008 'Get memory allocator failure 'CO_E_INIT_MEMORY_ALLOCATOR
Public Const ERROR_CO_E_INIT_CLASS_CACHE                                  As Long = &H80004009 'Unable to initialize class cache 'CO_E_INIT_CLASS_CACHE
Public Const ERROR_CO_E_INIT_RPC_CHANNEL                                  As Long = &H8000400A 'Unable to initialize RPC services 'CO_E_INIT_RPC_CHANNEL
Public Const ERROR_CO_E_INIT_TLS_SET_CHANNEL_CONTROL                      As Long = &H8000400B 'Cannot set thread local storage channel control 'CO_E_INIT_TLS_SET_CHANNEL_CONTROL
Public Const ERROR_CO_E_INIT_TLS_CHANNEL_CONTROL                          As Long = &H8000400C 'Could not allocate thread local storage channel control 'CO_E_INIT_TLS_CHANNEL_CONTROL
Public Const ERROR_CO_E_INIT_UNACCEPTED_USER_ALLOCATOR                    As Long = &H8000400D 'The user supplied memory allocator is unacceptable 'CO_E_INIT_UNACCEPTED_USER_ALLOCATOR
Public Const ERROR_CO_E_INIT_SCM_MUTEX_EXISTS                             As Long = &H8000400E 'The OLE service mutex already exists 'CO_E_INIT_SCM_MUTEX_EXISTS
Public Const ERROR_CO_E_INIT_SCM_FILE_MAPPING_EXISTS                      As Long = &H8000400F 'The OLE service file mapping already exists 'CO_E_INIT_SCM_FILE_MAPPING_EXISTS
Public Const ERROR_CO_E_INIT_SCM_MAP_VIEW_OF_FILE                         As Long = &H80004010 'Unable to map view of file for OLE service 'CO_E_INIT_SCM_MAP_VIEW_OF_FILE
Public Const ERROR_CO_E_INIT_SCM_EXEC_FAILURE                             As Long = &H80004011 'Failure attempting to launch OLE service 'CO_E_INIT_SCM_EXEC_FAILURE
Public Const ERROR_CO_E_INIT_ONLY_SINGLE_THREADED                         As Long = &H80004012 'There was an attempt to call CoInitialize a second time while single threaded 'CO_E_INIT_ONLY_SINGLE_THREADED
Public Const ERROR_CO_E_CANT_REMOTE                                       As Long = &H80004013 'A Remote activation was necessary but was not allowed 'CO_E_CANT_REMOTE
Public Const ERROR_CO_E_BAD_SERVER_NAME                                   As Long = &H80004014 'A Remote activation was necessary but the server name provided was invalid 'CO_E_BAD_SERVER_NAME
Public Const ERROR_CO_E_WRONG_SERVER_IDENTITY                             As Long = &H80004015 'The class is configured to run as a security id different from the caller 'CO_E_WRONG_SERVER_IDENTITY
Public Const ERROR_CO_E_OLE1DDE_DISABLED                                  As Long = &H80004016 'Use of Ole1 services requiring DDE windows is disabled 'CO_E_OLE1DDE_DISABLED
Public Const ERROR_CO_E_RUNAS_SYNTAX                                      As Long = &H80004017 'A RunAs specification must be <domain name>\<user name> or simply <user name> 'CO_E_RUNAS_SYNTAX
Public Const ERROR_CO_E_CREATEPROCESS_FAILURE                             As Long = &H80004018 'The server process could not be started.  The pathname may be incorrect. 'CO_E_CREATEPROCESS_FAILURE
Public Const ERROR_CO_E_RUNAS_CREATEPROCESS_FAILURE                       As Long = &H80004019 'The server process could not be started as the configured identity.  The pathname may be incorrect or unavailable. 'CO_E_RUNAS_CREATEPROCESS_FAILURE
Public Const ERROR_CO_E_RUNAS_LOGON_FAILURE                               As Long = &H8000401A 'The server process could not be started because the configured identity is incorrect.  Check the username and password. 'CO_E_RUNAS_LOGON_FAILURE
Public Const ERROR_CO_E_LAUNCH_PERMSSION_DENIED                           As Long = &H8000401B 'The client is not allowed to launch this server. 'CO_E_LAUNCH_PERMSSION_DENIED
Public Const ERROR_CO_E_START_SERVICE_FAILURE                             As Long = &H8000401C 'The service providing this server could not be started. 'CO_E_START_SERVICE_FAILURE
Public Const ERROR_CO_E_REMOTE_COMMUNICATION_FAILURE                      As Long = &H8000401D 'This computer was unable to communicate with the computer providing the server. 'CO_E_REMOTE_COMMUNICATION_FAILURE
Public Const ERROR_CO_E_SERVER_START_TIMEOUT                              As Long = &H8000401E 'The server did not respond after being launched. 'CO_E_SERVER_START_TIMEOUT
Public Const ERROR_CO_E_CLSREG_INCONSISTENT                               As Long = &H8000401F 'The registration information for this server is inconsistent or incomplete. 'CO_E_CLSREG_INCONSISTENT
Public Const ERROR_CO_E_IIDREG_INCONSISTENT                               As Long = &H80004020 'The registration information for this interface is inconsistent or incomplete. 'CO_E_IIDREG_INCONSISTENT
Public Const ERROR_CO_E_NOT_SUPPORTED                                     As Long = &H80004021 'The operation attempted is not supported. 'CO_E_NOT_SUPPORTED
Public Const ERROR_CO_E_RELOAD_DLL                                        As Long = &H80004022 'A dll must be loaded. 'CO_E_RELOAD_DLL
Public Const ERROR_CO_E_MSI_ERROR                                         As Long = &H80004023 'A Microsoft Software Installer error was encountered. 'CO_E_MSI_ERROR
Public Const ERROR_CO_E_ATTEMPT_TO_CREATE_OUTSIDE_CLIENT_CONTEXT          As Long = &H80004024 'The specified activation could not occur in the client context as specified. 'CO_E_ATTEMPT_TO_CREATE_OUTSIDE_CLIENT_CONTEXT
Public Const ERROR_CO_E_SERVER_PAUSED                                     As Long = &H80004025 'Activations on the server are paused. 'CO_E_SERVER_PAUSED
Public Const ERROR_CO_E_SERVER_NOT_PAUSED                                 As Long = &H80004026 'Activations on the server are not paused. 'CO_E_SERVER_NOT_PAUSED
Public Const ERROR_CO_E_CLASS_DISABLED                                    As Long = &H80004027 'The component or application containing the component has been disabled. 'CO_E_CLASS_DISABLED
Public Const ERROR_CO_E_CLRNOTAVAILABLE                                   As Long = &H80004028 'The common language runtime is not available 'CO_E_CLRNOTAVAILABLE
Public Const ERROR_CO_E_ASYNC_WORK_REJECTED                               As Long = &H80004029 'The thread-pool rejected the submitted asynchronous work. 'CO_E_ASYNC_WORK_REJECTED
Public Const ERROR_CO_E_SERVER_INIT_TIMEOUT                               As Long = &H8000402A 'The server started, but did not finish initializing in a timely fashion. 'CO_E_SERVER_INIT_TIMEOUT
Public Const ERROR_CO_E_NO_SECCTX_IN_ACTIVATE                             As Long = &H8000402B 'Unable to complete the call since there is no COM+ security context inside IObjectControl.Activate. 'CO_E_NO_SECCTX_IN_ACTIVATE
Public Const ERROR_CO_E_TRACKER_CONFIG                                    As Long = &H80004030 'The provided tracker configuration is invalid 'CO_E_TRACKER_CONFIG
Public Const ERROR_CO_E_THREADPOOL_CONFIG                                 As Long = &H80004031 'The provided thread pool configuration is invalid 'CO_E_THREADPOOL_CONFIG
Public Const ERROR_CO_E_SXS_CONFIG                                        As Long = &H80004032 'The provided side-by-side configuration is invalid 'CO_E_SXS_CONFIG
Public Const ERROR_CO_E_MALFORMED_SPN                                     As Long = &H80004033 'The server principal name (SPN) obtained during security negotiation is malformed. 'CO_E_MALFORMED_SPN
Public Const ERROR_OLE_E_OLEVERB                                          As Long = &H80040000 'Invalid OLEVERB structure 'OLE_E_OLEVERB
Public Const ERROR_OLE_E_ADVF                                             As Long = &H80040001 'Invalid advise flags 'OLE_E_ADVF
Public Const ERROR_OLE_E_ENUM_NOMORE                                      As Long = &H80040002 'Can't enumerate any more, because the associated data is missing 'OLE_E_ENUM_NOMORE
Public Const ERROR_OLE_E_ADVISENOTSUPPORTED                               As Long = &H80040003 'This implementation doesn't take advises 'OLE_E_ADVISENOTSUPPORTED
Public Const ERROR_OLE_E_NOCONNECTION                                     As Long = &H80040004 'There is no connection for this connection ID 'OLE_E_NOCONNECTION
Public Const ERROR_OLE_E_NOTRUNNING                                       As Long = &H80040005 'Need to run the object to perform this operation 'OLE_E_NOTRUNNING
Public Const ERROR_OLE_E_NOCACHE                                          As Long = &H80040006 'There is no cache to operate on 'OLE_E_NOCACHE
Public Const ERROR_OLE_E_BLANK                                            As Long = &H80040007 'Uninitialized object 'OLE_E_BLANK
Public Const ERROR_OLE_E_CLASSDIFF                                        As Long = &H80040008 'Linked object's source class has changed 'OLE_E_CLASSDIFF
Public Const ERROR_OLE_E_CANT_GETMONIKER                                  As Long = &H80040009 'Not able to get the moniker of the object 'OLE_E_CANT_GETMONIKER
Public Const ERROR_OLE_E_CANT_BINDTOSOURCE                                As Long = &H8004000A 'Not able to bind to the source 'OLE_E_CANT_BINDTOSOURCE
Public Const ERROR_OLE_E_STATIC                                           As Long = &H8004000B 'Object is static; operation not allowed 'OLE_E_STATIC
Public Const ERROR_OLE_E_PROMPTSAVECANCELLED                              As Long = &H8004000C 'User canceled out of save dialog 'OLE_E_PROMPTSAVECANCELLED
Public Const ERROR_OLE_E_INVALIDRECT                                      As Long = &H8004000D 'Invalid rectangle 'OLE_E_INVALIDRECT
Public Const ERROR_OLE_E_WRONGCOMPOBJ                                     As Long = &H8004000E 'compobj.dll is too old for the ole2.dll initialized 'OLE_E_WRONGCOMPOBJ
Public Const ERROR_OLE_E_INVALIDHWND                                      As Long = &H8004000F 'Invalid window handle 'OLE_E_INVALIDHWND
Public Const ERROR_OLE_E_NOT_INPLACEACTIVE                                As Long = &H80040010 'Object is not in any of the inplace active states 'OLE_E_NOT_INPLACEACTIVE
Public Const ERROR_OLE_E_CANTCONVERT                                      As Long = &H80040011 'Not able to convert object 'OLE_E_CANTCONVERT
Public Const ERROR_OLE_E_NOSTORAGE                                        As Long = &H80040012 'Not able to perform the operation because object is not given storage yet 'OLE_E_NOSTORAGE
Public Const ERROR_DV_E_FORMATETC                                         As Long = &H80040064 'Invalid FORMATETC structure 'DV_E_FORMATETC
Public Const ERROR_DV_E_DVTARGETDEVICE                                    As Long = &H80040065 'Invalid DVTARGETDEVICE structure 'DV_E_DVTARGETDEVICE
Public Const ERROR_DV_E_STGMEDIUM                                         As Long = &H80040066 'Invalid STDGMEDIUM structure 'DV_E_STGMEDIUM
Public Const ERROR_DV_E_STATDATA                                          As Long = &H80040067 'Invalid STATDATA structure 'DV_E_STATDATA
Public Const ERROR_DV_E_LINDEX                                            As Long = &H80040068 'Invalid lindex 'DV_E_LINDEX
Public Const ERROR_DV_E_TYMED                                             As Long = &H80040069 'Invalid tymed 'DV_E_TYMED
Public Const ERROR_DV_E_CLIPFORMAT                                        As Long = &H8004006A 'Invalid clipboard format 'DV_E_CLIPFORMAT
Public Const ERROR_DV_E_DVASPECT                                          As Long = &H8004006B 'Invalid aspect(s) 'DV_E_DVASPECT
Public Const ERROR_DV_E_DVTARGETDEVICE_SIZE                               As Long = &H8004006C 'tdSize parameter of the DVTARGETDEVICE structure is invalid 'DV_E_DVTARGETDEVICE_SIZE
Public Const ERROR_DV_E_NOIVIEWOBJECT                                     As Long = &H8004006D 'Object doesn't support IViewObject interface 'DV_E_NOIVIEWOBJECT
Public Const ERROR_DRAGDROP_E_NOTREGISTERED                               As Long = &H80040100 'Trying to revoke a drop target that has not been registered 'DRAGDROP_E_NOTREGISTERED
Public Const ERROR_DRAGDROP_E_ALREADYREGISTERED                           As Long = &H80040101 'This window has already been registered as a drop target 'DRAGDROP_E_ALREADYREGISTERED
Public Const ERROR_DRAGDROP_E_INVALIDHWND                                 As Long = &H80040102 'Invalid window handle 'DRAGDROP_E_INVALIDHWND
Public Const ERROR_CLASS_E_NOAGGREGATION                                  As Long = &H80040110 'Class does not support aggregation (or class object is remote) 'CLASS_E_NOAGGREGATION
Public Const ERROR_CLASS_E_CLASSNOTAVAILABLE                              As Long = &H80040111 'ClassFactory cannot supply requested class 'CLASS_E_CLASSNOTAVAILABLE
Public Const ERROR_CLASS_E_NOTLICENSED                                    As Long = &H80040112 'Class is not licensed for use 'CLASS_E_NOTLICENSED
Public Const ERROR_VIEW_E_DRAW                                            As Long = &H80040140 'Error drawing view 'VIEW_E_DRAW
Public Const ERROR_REGDB_E_INVALIDVALUE                                   As Long = &H80040153 'Could not read key from registry 'REGDB_E_READREGDB
Public Const ERROR_REGDB_E_CLASSNOTREG                                    As Long = &H80040154 'Class not registered 'REGDB_E_CLASSNOTREG
Public Const ERROR_REGDB_E_IIDNOTREG                                      As Long = &H80040155 'Interface not registered 'REGDB_E_IIDNOTREG
Public Const ERROR_REGDB_E_BADTHREADINGMODEL                              As Long = &H80040156 'Threading model entry is not valid 'REGDB_E_BADTHREADINGMODEL
Public Const ERROR_REGDB_E_READREGDB                                      As Long = &H80040150 'The Registry could not be read. 'REGDB_E_READREGDB
Public Const ERROR_CAT_E_CATIDNOEXIST                                     As Long = &H80040160 'CATID does not exist 'CAT_E_CATIDNOEXIST
Public Const ERROR_CAT_E_NODESCRIPTION                                    As Long = &H80040161 'Description not found 'CAT_E_NODESCRIPTION
Public Const ERROR_CS_E_PACKAGE_NOTFOUND                                  As Long = &H80040164 'No package in the software installation data in the Active Directory meets this criteria. 'CS_E_PACKAGE_NOTFOUND
Public Const ERROR_CS_E_NOT_DELETABLE                                     As Long = &H80040165 'Deleting this will break the referential integrity of the software installation data in the Active Directory. 'CS_E_NOT_DELETABLE
Public Const ERROR_CS_E_CLASS_NOTFOUND                                    As Long = &H80040166 'The CLSID was not found in the software installation data in the Active Directory. 'CS_E_CLASS_NOTFOUND
Public Const ERROR_CS_E_INVALID_VERSION                                   As Long = &H80040167 'The software installation data in the Active Directory is corrupt. 'CS_E_INVALID_VERSION
Public Const ERROR_CS_E_NO_CLASSSTORE                                     As Long = &H80040168 'There is no software installation data in the Active Directory. 'CS_E_NO_CLASSSTORE
Public Const ERROR_CS_E_OBJECT_NOTFOUND                                   As Long = &H80040169 'There is no software installation data object in the Active Directory. 'CS_E_OBJECT_NOTFOUND
Public Const ERROR_CS_E_OBJECT_ALREADY_EXISTS                             As Long = &H8004016A 'The software installation data object in the Active Directory already exists. 'CS_E_OBJECT_ALREADY_EXISTS
Public Const ERROR_CS_E_INVALID_PATH                                      As Long = &H8004016B 'The path to the software installation data in the Active Directory is not correct. 'CS_E_INVALID_PATH
Public Const ERROR_CS_E_NETWORK_ERROR                                     As Long = &H8004016C 'A network error interrupted the operation. 'CS_E_NETWORK_ERROR
Public Const ERROR_CS_E_ADMIN_LIMIT_EXCEEDED                              As Long = &H8004016D 'The size of this object exceeds the maximum size set by the Administrator. 'CS_E_ADMIN_LIMIT_EXCEEDED
Public Const ERROR_CS_E_SCHEMA_MISMATCH                                   As Long = &H8004016E 'The schema for the software installation data in the Active Directory does not match the required schema. 'CS_E_SCHEMA_MISMATCH
Public Const ERROR_CS_E_INTERNAL_ERROR                                    As Long = &H8004016F 'An error occurred in the software installation data in the Active Directory. 'CS_E_INTERNAL_ERROR
Public Const ERROR_CACHE_E_NOCACHE_UPDATED                                As Long = &H80040170 'Cache not updated 'CACHE_E_NOCACHE_UPDATED
Public Const ERROR_OLEOBJ_E_NOVERBS                                       As Long = &H80040180 'No verbs for OLE object 'OLEOBJ_E_NOVERBS
Public Const ERROR_OLEOBJ_E_INVALIDVERB                                   As Long = &H80040181 'Invalid verb for OLE object 'OLEOBJ_E_INVALIDVERB
Public Const ERROR_INPLACE_E_NOTUNDOABLE                                  As Long = &H800401A0 'Undo is not available 'INPLACE_E_NOTUNDOABLE
Public Const ERROR_INPLACE_E_NOTOOLSPACE                                  As Long = &H800401A1 'Space for tools is not available 'INPLACE_E_NOTOOLSPACE
Public Const ERROR_CONVERT10_E_OLESTREAM_GET                              As Long = &H800401C0 'OLESTREAM Get method failed 'CONVERT10_E_OLESTREAM_GET
Public Const ERROR_CONVERT10_E_OLESTREAM_PUT                              As Long = &H800401C1 'OLESTREAM Put method failed 'CONVERT10_E_OLESTREAM_PUT
Public Const ERROR_CONVERT10_E_OLESTREAM_FMT                              As Long = &H800401C2 'Contents of the OLESTREAM not in correct format 'CONVERT10_E_OLESTREAM_FMT
Public Const ERROR_CONVERT10_E_OLESTREAM_BITMAP_TO_DIB                    As Long = &H800401C3 'There was an error in a Windows GDI call while converting the bitmap to a DIB 'CONVERT10_E_OLESTREAM_BITMAP_TO_DIB
Public Const ERROR_CONVERT10_E_STG_FMT                                    As Long = &H800401C4 'Contents of the IStorage not in correct format 'CONVERT10_E_STG_FMT
Public Const ERROR_CONVERT10_E_STG_NO_STD_STREAM                          As Long = &H800401C5 'Contents of IStorage is missing one of the standard streams 'CONVERT10_E_STG_NO_STD_STREAM
Public Const ERROR_CONVERT10_E_STG_DIB_TO_BITMAP                          As Long = &H800401C6 'There was an error in a Windows GDI call while converting the DIB to a bitmap. 'CONVERT10_E_STG_DIB_TO_BITMAP
Public Const ERROR_CLIPBRD_E_CANT_OPEN                                    As Long = &H800401D0 'OpenClipboard Failed 'CLIPBRD_E_CANT_OPEN
Public Const ERROR_CLIPBRD_E_CANT_EMPTY                                   As Long = &H800401D1 'EmptyClipboard Failed 'CLIPBRD_E_CANT_EMPTY
Public Const ERROR_CLIPBRD_E_CANT_SET                                     As Long = &H800401D2 'SetClipboard Failed 'CLIPBRD_E_CANT_SET
Public Const ERROR_CLIPBRD_E_BAD_DATA                                     As Long = &H800401D3 'Data on clipboard is invalid 'CLIPBRD_E_BAD_DATA
Public Const ERROR_CLIPBRD_E_CANT_CLOSE                                   As Long = &H800401D4 'CloseClipboard Failed 'CLIPBRD_E_CANT_CLOSE
Public Const ERROR_MK_E_CONNECTMANUALLY                                   As Long = &H800401E0 'Moniker needs to be connected manually 'MK_E_CONNECTMANUALLY
Public Const ERROR_MK_E_EXCEEDEDDEADLINE                                  As Long = &H800401E1 'Operation exceeded deadline 'MK_E_EXCEEDEDDEADLINE
Public Const ERROR_MK_E_NEEDGENERIC                                       As Long = &H800401E2 'Moniker needs to be generic 'MK_E_NEEDGENERIC
Public Const ERROR_MK_E_UNAVAILABLE                                       As Long = &H800401E3 'Operation unavailable 'MK_E_UNAVAILABLE
Public Const ERROR_MK_E_SYNTAX                                            As Long = &H800401E4 'Invalid syntax 'MK_E_SYNTAX
Public Const ERROR_MK_E_NOOBJECT                                          As Long = &H800401E5 'No object for moniker 'MK_E_NOOBJECT
Public Const ERROR_MK_E_INVALIDEXTENSION                                  As Long = &H800401E6 'Bad extension for file 'MK_E_INVALIDEXTENSION
Public Const ERROR_MK_E_INTERMEDIATEINTERFACENOTSUPPORTED                 As Long = &H800401E7 'Intermediate operation failed 'MK_E_INTERMEDIATEINTERFACENOTSUPPORTED
Public Const ERROR_MK_E_NOTBINDABLE                                       As Long = &H800401E8 'Moniker is not bindable 'MK_E_NOTBINDABLE
Public Const ERROR_MK_E_NOTBOUND                                          As Long = &H800401E9 'Moniker is not bound 'MK_E_NOTBOUND
Public Const ERROR_MK_E_CANTOPENFILE                                      As Long = &H800401EA 'Moniker cannot open file 'MK_E_CANTOPENFILE
Public Const ERROR_MK_E_MUSTBOTHERUSER                                    As Long = &H800401EB 'User input required for operation to succeed 'MK_E_MUSTBOTHERUSER
Public Const ERROR_MK_E_NOINVERSE                                         As Long = &H800401EC 'Moniker class has no inverse 'MK_E_NOINVERSE
Public Const ERROR_MK_E_NOSTORAGE                                         As Long = &H800401ED 'Moniker does not refer to storage 'MK_E_NOSTORAGE
Public Const ERROR_MK_E_NOPREFIX                                          As Long = &H800401EE 'No common prefix 'MK_E_NOPREFIX
Public Const ERROR_MK_E_ENUMERATION_FAILED                                As Long = &H800401EF 'Moniker could not be enumerated 'MK_E_ENUMERATION_FAILED
Public Const ERROR_CO_E_NOTINITIALIZED                                    As Long = &H800401F0 'CoInitialize has not been called. 'CO_E_NOTINITIALIZED
Public Const ERROR_CO_E_ALREADYINITIALIZED                                As Long = &H800401F1 'CoInitialize has already been called. 'CO_E_ALREADYINITIALIZED
Public Const ERROR_CO_E_CANTDETERMINECLASS                                As Long = &H800401F2 'Class of object cannot be determined 'CO_E_CANTDETERMINECLASS
Public Const ERROR_CO_E_CLASSSTRING                                       As Long = &H800401F3 'Invalid class string 'CO_E_CLASSSTRING
Public Const ERROR_CO_E_IIDSTRING                                         As Long = &H800401F4 'Invalid interface string 'CO_E_IIDSTRING
Public Const ERROR_CO_E_APPNOTFOUND                                       As Long = &H800401F5 'Application not found 'CO_E_APPNOTFOUND
Public Const ERROR_CO_E_APPSINGLEUSE                                      As Long = &H800401F6 'Application cannot be run more than once 'CO_E_APPSINGLEUSE
Public Const ERROR_CO_E_ERRORINAPP                                        As Long = &H800401F7 'Some error in application program 'CO_E_ERRORINAPP
Public Const ERROR_CO_E_DLLNOTFOUND                                       As Long = &H800401F8 'DLL for class not found 'CO_E_DLLNOTFOUND
Public Const ERROR_CO_E_ERRORINDLL                                        As Long = &H800401F9 'Error in the DLL 'CO_E_ERRORINDLL
Public Const ERROR_CO_E_WRONGOSFORAPP                                     As Long = &H800401FA 'Wrong OS or OS version for application 'CO_E_WRONGOSFORAPP
Public Const ERROR_CO_E_OBJNOTREG                                         As Long = &H800401FB 'Object is not registered 'CO_E_OBJNOTREG
Public Const ERROR_CO_E_OBJISREG                                          As Long = &H800401FC 'Object is already registered 'CO_E_OBJISREG
Public Const ERROR_CO_E_OBJNOTCONNECTED                                   As Long = &H800401FD 'Object is not connected to server 'CO_E_OBJNOTCONNECTED
Public Const ERROR_CO_E_APPDIDNTREG                                       As Long = &H800401FE 'Application was launched but it didn't register a class factory 'CO_E_APPDIDNTREG
Public Const ERROR_CO_E_RELEASED                                          As Long = &H800401FF 'Object has been released 'CO_E_RELEASED
Public Const ERROR_EVENT_S_SOME_SUBSCRIBERS_FAILED                        As Long = &H40200    'An event was able to invoke some but not all of the subscribers 'EVENT_S_SOME_SUBSCRIBERS_FAILED
Public Const ERROR_EVENT_E_ALL_SUBSCRIBERS_FAILED                         As Long = &H80040201 'An event was unable to invoke any of the subscribers 'EVENT_E_ALL_SUBSCRIBERS_FAILED
Public Const ERROR_EVENT_S_NOSUBSCRIBERS                                  As Long = &H40202    'An event was delivered but there were no subscribers 'EVENT_S_NOSUBSCRIBERS
Public Const ERROR_EVENT_E_QUERYSYNTAX                                    As Long = &H80040203 'A syntax error occurred trying to evaluate a query string 'EVENT_E_QUERYSYNTAX
Public Const ERROR_EVENT_E_QUERYFIELD                                     As Long = &H80040204 'An invalid field name was used in a query string 'EVENT_E_QUERYFIELD
Public Const ERROR_EVENT_E_INTERNALEXCEPTION                              As Long = &H80040205 'An unexpected exception was raised 'EVENT_E_INTERNALEXCEPTION
Public Const ERROR_EVENT_E_INTERNALERROR                                  As Long = &H80040206 'An unexpected internal error was detected 'EVENT_E_INTERNALERROR
Public Const ERROR_EVENT_E_INVALID_PER_USER_SID                           As Long = &H80040207 'The owner SID on a per-user subscription doesn't exist 'EVENT_E_INVALID_PER_USER_SID
Public Const ERROR_EVENT_E_USER_EXCEPTION                                 As Long = &H80040208 'A user-supplied component or subscriber raised an exception 'EVENT_E_USER_EXCEPTION
Public Const ERROR_EVENT_E_TOO_MANY_METHODS                               As Long = &H80040209 'An interface has too many methods to fire events from 'EVENT_E_TOO_MANY_METHODS
Public Const ERROR_EVENT_E_MISSING_EVENTCLASS                             As Long = &H8004020A 'A subscription cannot be stored unless its event class already exists 'EVENT_E_MISSING_EVENTCLASS
Public Const ERROR_EVENT_E_NOT_ALL_REMOVED                                As Long = &H8004020B 'Not all the objects requested could be removed 'EVENT_E_NOT_ALL_REMOVED
Public Const ERROR_EVENT_E_COMPLUS_NOT_INSTALLED                          As Long = &H8004020C 'COM+ is required for this operation, but is not installed 'EVENT_E_COMPLUS_NOT_INSTALLED
Public Const ERROR_EVENT_E_CANT_MODIFY_OR_DELETE_UNCONFIGURED_OBJECT      As Long = &H8004020D 'Cannot modify or delete an object that was not added using the COM+ Admin SDK 'EVENT_E_CANT_MODIFY_OR_DELETE_UNCONFIGURED_OBJECT
Public Const ERROR_EVENT_E_CANT_MODIFY_OR_DELETE_CONFIGURED_OBJECT        As Long = &H8004020E 'Cannot modify or delete an object that was added using the COM+ Admin SDK 'EVENT_E_CANT_MODIFY_OR_DELETE_CONFIGURED_OBJECT
Public Const ERROR_EVENT_E_INVALID_EVENT_CLASS_PARTITION                  As Long = &H8004020F 'The event class for this subscription is in an invalid partition 'EVENT_E_INVALID_EVENT_CLASS_PARTITION
Public Const ERROR_EVENT_E_PER_USER_SID_NOT_LOGGED_ON                     As Long = &H80040210 'The owner of the PerUser subscription is not logged on to the system specified 'EVENT_E_PER_USER_SID_NOT_LOGGED_ON
Public Const ERROR_XACT_E_ALREADYOTHERSINGLEPHASE                         As Long = &H8004D000 'Another single phase resource manager has already been enlisted in this transaction. 'XACT_E_ALREADYOTHERSINGLEPHASE
Public Const ERROR_XACT_E_CANTRETAIN                                      As Long = &H8004D001 'A retaining commit or abort is not supported 'XACT_E_CANTRETAIN
Public Const ERROR_XACT_E_COMMITFAILED                                    As Long = &H8004D002 'The transaction failed to commit for an unknown reason. The transaction was aborted. 'XACT_E_COMMITFAILED
Public Const ERROR_XACT_E_COMMITPREVENTED                                 As Long = &H8004D003 'Cannot call commit on this transaction object because the calling application did not initiate the transaction. 'XACT_E_COMMITPREVENTED
Public Const ERROR_XACT_E_HEURISTICABORT                                  As Long = &H8004D004 'Instead of committing, the resource heuristically aborted. 'XACT_E_HEURISTICABORT
Public Const ERROR_XACT_E_HEURISTICCOMMIT                                 As Long = &H8004D005 'Instead of aborting, the resource heuristically committed. 'XACT_E_HEURISTICCOMMIT
Public Const ERROR_XACT_E_HEURISTICDAMAGE                                 As Long = &H8004D006 'Some of the states of the resource were committed while others were aborted, likely because of heuristic decisions. 'XACT_E_HEURISTICDAMAGE
Public Const ERROR_XACT_E_HEURISTICDANGER                                 As Long = &H8004D007 'Some of the states of the resource may have been committed while others may have been aborted, likely because of heuristic decisions. 'XACT_E_HEURISTICDANGER
Public Const ERROR_XACT_E_ISOLATIONLEVEL                                  As Long = &H8004D008 'The requested isolation level is not valid or supported. 'XACT_E_ISOLATIONLEVEL
Public Const ERROR_XACT_E_NOASYNC                                         As Long = &H8004D009 'The transaction manager doesn't support an asynchronous operation for this method. 'XACT_E_NOASYNC
Public Const ERROR_XACT_E_NOENLIST                                        As Long = &H8004D00A 'Unable to enlist in the transaction. 'XACT_E_NOENLIST
Public Const ERROR_XACT_E_NOISORETAIN                                     As Long = &H8004D00B 'The requested semantics of retention of isolation across retaining commit and abort boundaries cannot be supported by this transaction implementation, or isoFlags was not equal to zero. 'XACT_E_NOISORETAIN
Public Const ERROR_XACT_E_NORESOURCE                                      As Long = &H8004D00C 'There is no resource presently associated with this enlistment 'XACT_E_NORESOURCE
Public Const ERROR_XACT_E_NOTCURRENT                                      As Long = &H8004D00D 'The transaction failed to commit due to the failure of optimistic concurrency control in at least one of the resource managers. 'XACT_E_NOTCURRENT
Public Const ERROR_XACT_E_NOTRANSACTION                                   As Long = &H8004D00E 'The transaction has already been implicitly or explicitly committed or aborted 'XACT_E_NOTRANSACTION
Public Const ERROR_XACT_E_NOTSUPPORTED                                    As Long = &H8004D00F 'An invalid combination of flags was specified 'XACT_E_NOTSUPPORTED
Public Const ERROR_XACT_E_UNKNOWNRMGRID                                   As Long = &H8004D010 'The resource manager id is not associated with this transaction or the transaction manager. 'XACT_E_UNKNOWNRMGRID
Public Const ERROR_XACT_E_WRONGSTATE                                      As Long = &H8004D011 'This method was called in the wrong state 'XACT_E_WRONGSTATE
Public Const ERROR_XACT_E_WRONGUOW                                        As Long = &H8004D012 'The indicated unit of work does not match the unit of work expected by the resource manager. 'XACT_E_WRONGUOW
Public Const ERROR_XACT_E_XTIONEXISTS                                     As Long = &H8004D013 'An enlistment in a transaction already exists. 'XACT_E_XTIONEXISTS
Public Const ERROR_XACT_E_NOIMPORTOBJECT                                  As Long = &H8004D014 'An import object for the transaction could not be found. 'XACT_E_NOIMPORTOBJECT
Public Const ERROR_XACT_E_INVALIDCOOKIE                                   As Long = &H8004D015 'The transaction cookie is invalid. 'XACT_E_INVALIDCOOKIE
Public Const ERROR_XACT_E_INDOUBT                                         As Long = &H8004D016 'The transaction status is in doubt. A communication failure occurred, or a transaction manager or resource manager has failed 'XACT_E_INDOUBT
Public Const ERROR_XACT_E_NOTIMEOUT                                       As Long = &H8004D017 'A time-out was specified, but time-outs are not supported. 'XACT_E_NOTIMEOUT
Public Const ERROR_XACT_E_ALREADYINPROGRESS                               As Long = &H8004D018 'The requested operation is already in progress for the transaction. 'XACT_E_ALREADYINPROGRESS
Public Const ERROR_XACT_E_ABORTED                                         As Long = &H8004D019 'The transaction has already been aborted. 'XACT_E_ABORTED
Public Const ERROR_XACT_E_LOGFULL                                         As Long = &H8004D01A 'The Transaction Manager returned a log full error. 'XACT_E_LOGFULL
Public Const ERROR_XACT_E_TMNOTAVAILABLE                                  As Long = &H8004D01B 'The Transaction Manager is not available. 'XACT_E_TMNOTAVAILABLE
Public Const ERROR_XACT_E_CONNECTION_DOWN                                 As Long = &H8004D01C 'A connection with the transaction manager was lost. 'XACT_E_CONNECTION_DOWN
Public Const ERROR_XACT_E_CONNECTION_DENIED                               As Long = &H8004D01D 'A request to establish a connection with the transaction manager was denied. 'XACT_E_CONNECTION_DENIED
Public Const ERROR_XACT_E_REENLISTTIMEOUT                                 As Long = &H8004D01E 'Resource manager reenlistment to determine transaction status timed out. 'XACT_E_REENLISTTIMEOUT
Public Const ERROR_XACT_E_TIP_CONNECT_FAILED                              As Long = &H8004D01F 'This transaction manager failed to establish a connection with another TIP transaction manager. 'XACT_E_TIP_CONNECT_FAILED
Public Const ERROR_XACT_E_TIP_PROTOCOL_ERROR                              As Long = &H8004D020 'This transaction manager encountered a protocol error with another TIP transaction manager. 'XACT_E_TIP_PROTOCOL_ERROR
Public Const ERROR_XACT_E_TIP_PULL_FAILED                                 As Long = &H8004D021 'This transaction manager could not propagate a transaction from another TIP transaction manager. 'XACT_E_TIP_PULL_FAILED
Public Const ERROR_XACT_E_DEST_TMNOTAVAILABLE                             As Long = &H8004D022 'The Transaction Manager on the destination machine is not available. 'XACT_E_DEST_TMNOTAVAILABLE
Public Const ERROR_XACT_E_TIP_DISABLED                                    As Long = &H8004D023 'The Transaction Manager has disabled its support for TIP. 'XACT_E_TIP_DISABLED
Public Const ERROR_XACT_E_NETWORK_TX_DISABLED                             As Long = &H8004D024 'The transaction manager has disabled its support for remote/network transactions. 'XACT_E_NETWORK_TX_DISABLED
Public Const ERROR_XACT_E_PARTNER_NETWORK_TX_DISABLED                     As Long = &H8004D025 'The partner transaction manager has disabled its support for remote/network transactions. 'XACT_E_PARTNER_NETWORK_TX_DISABLED
Public Const ERROR_XACT_E_XA_TX_DISABLED                                  As Long = &H8004D026 'The transaction manager has disabled its support for XA transactions. 'XACT_E_XA_TX_DISABLED
Public Const ERROR_XACT_E_UNABLE_TO_READ_DTC_CONFIG                       As Long = &H8004D027 'MSDTC was unable to read its configuration information. 'XACT_E_UNABLE_TO_READ_DTC_CONFIG
Public Const ERROR_XACT_E_UNABLE_TO_LOAD_DTC_PROXY                        As Long = &H8004D028 'MSDTC was unable to load the dtc proxy dll. 'XACT_E_UNABLE_TO_LOAD_DTC_PROXY
Public Const ERROR_XACT_E_ABORTING                                        As Long = &H8004D029 'The local transaction has aborted. 'XACT_E_ABORTING
Public Const ERROR_XACT_E_CLERKNOTFOUND                                   As Long = &H8004D080 'XACT_E_CLERKNOTFOUND 'XACT_E_CLERKNOTFOUND
Public Const ERROR_XACT_E_CLERKEXISTS                                     As Long = &H8004D081 'XACT_E_CLERKEXISTS 'XACT_E_CLERKEXISTS
Public Const ERROR_XACT_E_RECOVERYINPROGRESS                              As Long = &H8004D082 'XACT_E_RECOVERYINPROGRESS 'XACT_E_RECOVERYINPROGRESS
Public Const ERROR_XACT_E_TRANSACTIONCLOSED                               As Long = &H8004D083 'XACT_E_TRANSACTIONCLOSED 'XACT_E_TRANSACTIONCLOSED
Public Const ERROR_XACT_E_INVALIDLSN                                      As Long = &H8004D084 'XACT_E_INVALIDLSN 'XACT_E_INVALIDLSN
Public Const ERROR_XACT_E_REPLAYREQUEST                                   As Long = &H8004D085 'XACT_E_REPLAYREQUEST 'XACT_E_REPLAYREQUEST
Public Const ERROR_XACT_S_ASYNC                                           As Long = &H4D000    'An asynchronous operation was specified. The operation has begun, but its outcome is not known yet. 'XACT_S_ASYNC
Public Const ERROR_XACT_S_DEFECT                                          As Long = &H4D001    'XACT_S_DEFECT 'XACT_S_DEFECT
Public Const ERROR_XACT_S_READONLY                                        As Long = &H4D002    'The method call succeeded because the transaction was read-only. 'XACT_S_READONLY
Public Const ERROR_XACT_S_SOMENORETAIN                                    As Long = &H4D003    'The transaction was successfully aborted. However, this is a coordinated transaction, and some number of enlisted resources were aborted outright because they could not support abort-retaining semantics 'XACT_S_SOMENORETAIN
Public Const ERROR_XACT_S_OKINFORM                                        As Long = &H4D004    'No changes were made during this call, but the sink wants another chance to look if any other sinks make further changes. 'XACT_S_OKINFORM
Public Const ERROR_XACT_S_MADECHANGESCONTENT                              As Long = &H4D005    'The sink is content and wishes the transaction to proceed. Changes were made to one or more resources during this call. 'XACT_S_MADECHANGESCONTENT
Public Const ERROR_XACT_S_MADECHANGESINFORM                               As Long = &H4D006    'The sink is for the moment and wishes the transaction to proceed, but if other changes are made following this return by other event sinks then this sink wants another chance to look 'XACT_S_MADECHANGESINFORM
Public Const ERROR_XACT_S_ALLNORETAIN                                     As Long = &H4D007    'The transaction was successfully aborted. However, the abort was non-retaining. 'XACT_S_ALLNORETAIN
Public Const ERROR_XACT_S_ABORTING                                        As Long = &H4D008    'An abort operation was already in progress. 'XACT_S_ABORTING
Public Const ERROR_XACT_S_SINGLEPHASE                                     As Long = &H4D009    'The resource manager has performed a single-phase commit of the transaction. 'XACT_S_SINGLEPHASE
Public Const ERROR_XACT_S_LOCALLY_OK                                      As Long = &H4D00A    'The local transaction has not aborted. 'XACT_S_LOCALLY_OK
Public Const ERROR_XACT_S_LASTRESOURCEMANAGER                             As Long = &H4D010    'The resource manager has requested to be the coordinator (last resource manager) for the transaction. 'XACT_S_LASTRESOURCEMANAGER
Public Const ERROR_CONTEXT_E_ABORTED                                      As Long = &H8004E002 'The root transaction wanted to commit, but transaction aborted 'CONTEXT_E_ABORTED
Public Const ERROR_CONTEXT_E_ABORTING                                     As Long = &H8004E003 'You made a method call on a COM+ component that has a transaction that has already aborted or in the process of aborting. 'CONTEXT_E_ABORTING
Public Const ERROR_CONTEXT_E_NOCONTEXT                                    As Long = &H8004E004 'There is no MTS object context 'CONTEXT_E_NOCONTEXT
Public Const ERROR_CONTEXT_E_WOULD_DEADLOCK                               As Long = &H8004E005 'The component is configured to use synchronization and this method call would cause a deadlock to occur. 'CONTEXT_E_WOULD_DEADLOCK
Public Const ERROR_CONTEXT_E_SYNCH_TIMEOUT                                As Long = &H8004E006 'The component is configured to use synchronization and a thread has timed out waiting to enter the context. 'CONTEXT_E_SYNCH_TIMEOUT
Public Const ERROR_CONTEXT_E_OLDREF                                       As Long = &H8004E007 'You made a method call on a COM+ component that has a transaction that has already committed or aborted. 'CONTEXT_E_OLDREF
Public Const ERROR_CONTEXT_E_ROLENOTFOUND                                 As Long = &H8004E00C 'The specified role was not configured for the application 'CONTEXT_E_ROLENOTFOUND
Public Const ERROR_CONTEXT_E_TMNOTAVAILABLE                               As Long = &H8004E00F 'COM+ was unable to talk to the Microsoft Distributed Transaction Coordinator 'CONTEXT_E_TMNOTAVAILABLE
Public Const ERROR_CO_E_ACTIVATIONFAILED                                  As Long = &H8004E021 'An unexpected error occurred during COM+ Activation. 'CO_E_ACTIVATIONFAILED
Public Const ERROR_CO_E_ACTIVATIONFAILED_EVENTLOGGED                      As Long = &H8004E022 'COM+ Activation failed. Check the event log for more information 'CO_E_ACTIVATIONFAILED_EVENTLOGGED
Public Const ERROR_CO_E_ACTIVATIONFAILED_CATALOGERROR                     As Long = &H8004E023 'COM+ Activation failed due to a catalog or configuration error. 'CO_E_ACTIVATIONFAILED_CATALOGERROR
Public Const ERROR_CO_E_ACTIVATIONFAILED_TIMEOUT                          As Long = &H8004E024 'COM+ activation failed because the activation could not be completed in the specified amount of time. 'CO_E_ACTIVATIONFAILED_TIMEOUT
Public Const ERROR_CO_E_INITIALIZATIONFAILED                              As Long = &H8004E025 'COM+ Activation failed because an initialization function failed.  Check the event log for more information. 'CO_E_INITIALIZATIONFAILED
Public Const ERROR_CONTEXT_E_NOJIT                                        As Long = &H8004E026 'The requested operation requires that JIT be in the current context and it is not 'CONTEXT_E_NOJIT
Public Const ERROR_CONTEXT_E_NOTRANSACTION                                As Long = &H8004E027 'The requested operation requires that the current context have a Transaction, and it does not 'CONTEXT_E_NOTRANSACTION
Public Const ERROR_CO_E_THREADINGMODEL_CHANGED                            As Long = &H8004E028 'The components threading model has changed after install into a COM+ Application.  Please re-install component. 'CO_E_THREADINGMODEL_CHANGED
Public Const ERROR_CO_E_NOIISINTRINSICS                                   As Long = &H8004E029 'IIS intrinsics not available.  Start your work with IIS. 'CO_E_NOIISINTRINSICS
Public Const ERROR_CO_E_NOCOOKIES                                         As Long = &H8004E02A 'An attempt to write a cookie failed. 'CO_E_NOCOOKIES
Public Const ERROR_CO_E_DBERROR                                           As Long = &H8004E02B 'An attempt to use a database generated a database specific error. 'CO_E_DBERROR
Public Const ERROR_CO_E_NOTPOOLED                                         As Long = &H8004E02C 'The COM+ component you created must use object pooling to work. 'CO_E_NOTPOOLED
Public Const ERROR_CO_E_NOTCONSTRUCTED                                    As Long = &H8004E02D 'The COM+ component you created must use object construction to work correctly. 'CO_E_NOTCONSTRUCTED
Public Const ERROR_CO_E_NOSYNCHRONIZATION                                 As Long = &H8004E02E 'The COM+ component requires synchronization, and it is not configured for it. 'CO_E_NOSYNCHRONIZATION
Public Const ERROR_CO_E_ISOLEVELMISMATCH                                  As Long = &H8004E02F 'The TxIsolation Level property for the COM+ component being created is stronger than the TxIsolationLevel for the "root" component for the transaction.  The creation failed. 'CO_E_ISOLEVELMISMATCH
Public Const ERROR_OLE_S_USEREG                                           As Long = &H40000    'Use the registry database to provide the requested information 'OLE_S_USEREG
Public Const ERROR_OLE_S_STATIC                                           As Long = &H40001    'Success, but static 'OLE_S_STATIC
Public Const ERROR_OLE_S_MAC_CLIPFORMAT                                   As Long = &H40002    'Macintosh clipboard format 'OLE_S_MAC_CLIPFORMAT
Public Const ERROR_DRAGDROP_S_DROP                                        As Long = &H40100    'Successful drop took place 'DRAGDROP_S_DROP
Public Const ERROR_DRAGDROP_S_CANCEL                                      As Long = &H40101    'Drag-drop operation canceled 'DRAGDROP_S_CANCEL
Public Const ERROR_DRAGDROP_S_USEDEFAULTCURSORS                           As Long = &H40102    'Use the default cursor 'DRAGDROP_S_USEDEFAULTCURSORS
Public Const ERROR_DATA_S_SAMEFORMATETC                                   As Long = &H40130    'Data has same FORMATETC 'DATA_S_SAMEFORMATETC
Public Const ERROR_VIEW_S_ALREADY_FROZEN                                  As Long = &H40140    'View is already frozen 'VIEW_S_ALREADY_FROZEN
Public Const ERROR_CACHE_S_FORMATETC_NOTSUPPORTED                         As Long = &H40170    'FORMATETC not supported 'CACHE_S_FORMATETC_NOTSUPPORTED
Public Const ERROR_CACHE_S_SAMECACHE                                      As Long = &H40171    'Same cache 'CACHE_S_SAMECACHE
Public Const ERROR_CACHE_S_SOMECACHES_NOTUPDATED                          As Long = &H40172    'Some cache(s) not updated 'CACHE_S_SOMECACHES_NOTUPDATED
Public Const ERROR_OLEOBJ_S_INVALIDVERB                                   As Long = &H40180    'Invalid verb for OLE object 'OLEOBJ_S_INVALIDVERB
Public Const ERROR_OLEOBJ_S_CANNOT_DOVERB_NOW                             As Long = &H40181    'Verb number is valid but verb cannot be done now 'OLEOBJ_S_CANNOT_DOVERB_NOW
Public Const ERROR_OLEOBJ_S_INVALIDHWND                                   As Long = &H40182    'Invalid window handle passed 'OLEOBJ_S_INVALIDHWND
Public Const ERROR_INPLACE_S_TRUNCATED                                    As Long = &H401A0    'Message is too long; some of it had to be truncated before displaying 'INPLACE_S_TRUNCATED
Public Const ERROR_CONVERT10_S_NO_PRESENTATION                            As Long = &H401C0    'Unable to convert OLESTREAM to IStorage 'CONVERT10_S_NO_PRESENTATION
Public Const ERROR_MK_S_REDUCED_TO_SELF                                   As Long = &H401E2    'Moniker reduced to itself 'MK_S_REDUCED_TO_SELF
Public Const ERROR_MK_S_ME                                                As Long = &H401E4    'Common prefix is this moniker 'MK_S_ME
Public Const ERROR_MK_S_HIM                                               As Long = &H401E5    'Common prefix is input moniker 'MK_S_HIM
Public Const ERROR_MK_S_US                                                As Long = &H401E6    'Common prefix is both monikers 'MK_S_US
Public Const ERROR_MK_S_MONIKERALREADYREGISTERED                          As Long = &H401E7    'Moniker is already registered in running object table 'MK_S_MONIKERALREADYREGISTERED
Public Const ERROR_SCHED_S_TASK_READY                                     As Long = &H41300    'The task is ready to run at its next scheduled time. 'SCHED_S_TASK_READY
Public Const ERROR_SCHED_S_TASK_RUNNING                                   As Long = &H41301    'The task is currently running. 'SCHED_S_TASK_RUNNING
Public Const ERROR_SCHED_S_TASK_DISABLED                                  As Long = &H41302    'The task will not run at the scheduled times because it has been disabled. 'SCHED_S_TASK_DISABLED
Public Const ERROR_SCHED_S_TASK_HAS_NOT_RUN                               As Long = &H41303    'The task has not yet run. 'SCHED_S_TASK_HAS_NOT_RUN
Public Const ERROR_SCHED_S_TASK_NO_MORE_RUNS                              As Long = &H41304    'There are no more runs scheduled for this task. 'SCHED_S_TASK_NO_MORE_RUNS
Public Const ERROR_SCHED_S_TASK_NOT_SCHEDULED                             As Long = &H41305    'One or more of the properties that are needed to run this task on a schedule have not been set. 'SCHED_S_TASK_NOT_SCHEDULED
Public Const ERROR_SCHED_S_TASK_TERMINATED                                As Long = &H41306    'The last run of the task was terminated by the user. 'SCHED_S_TASK_TERMINATED
Public Const ERROR_SCHED_S_TASK_NO_VALID_TRIGGERS                         As Long = &H41307    'Either the task has no triggers or the existing triggers are disabled or not set. 'SCHED_S_TASK_NO_VALID_TRIGGERS
Public Const ERROR_SCHED_S_EVENT_TRIGGER                                  As Long = &H41308    'Event triggers don't have set run times. 'SCHED_S_EVENT_TRIGGER
Public Const ERROR_SCHED_E_TRIGGER_NOT_FOUND                              As Long = &H80041309 'Trigger not found. 'SCHED_E_TRIGGER_NOT_FOUND
Public Const ERROR_SCHED_E_TASK_NOT_READY                                 As Long = &H8004130A 'One or more of the properties that are needed to run this task have not been set. 'SCHED_E_TASK_NOT_READY
Public Const ERROR_SCHED_E_TASK_NOT_RUNNING                               As Long = &H8004130B 'There is no running instance of the task to terminate. 'SCHED_E_TASK_NOT_RUNNING
Public Const ERROR_SCHED_E_SERVICE_NOT_INSTALLED                          As Long = &H8004130C 'The Task Scheduler Service is not installed on this computer. 'SCHED_E_SERVICE_NOT_INSTALLED
Public Const ERROR_SCHED_E_CANNOT_OPEN_TASK                               As Long = &H8004130D 'The task object could not be opened. 'SCHED_E_CANNOT_OPEN_TASK
Public Const ERROR_SCHED_E_INVALID_TASK                                   As Long = &H8004130E 'The object is either an invalid task object or is not a task object. 'SCHED_E_INVALID_TASK
Public Const ERROR_SCHED_E_ACCOUNT_INFORMATION_NOT_SET                    As Long = &H8004130F 'No account information could be found in the Task Scheduler security database for the task indicated. 'SCHED_E_ACCOUNT_INFORMATION_NOT_SET
Public Const ERROR_SCHED_E_ACCOUNT_NAME_NOT_FOUND                         As Long = &H80041310 'Unable to establish existence of the account specified. 'SCHED_E_ACCOUNT_NAME_NOT_FOUND
Public Const ERROR_SCHED_E_ACCOUNT_DBASE_CORRUPT                          As Long = &H80041311 'Corruption was detected in the Task Scheduler security database; the database has been reset. 'SCHED_E_ACCOUNT_DBASE_CORRUPT
Public Const ERROR_SCHED_E_NO_SECURITY_SERVICES                           As Long = &H80041312 'Task Scheduler security services are available only on Windows NT. 'SCHED_E_NO_SECURITY_SERVICES
Public Const ERROR_SCHED_E_UNKNOWN_OBJECT_VERSION                         As Long = &H80041313 'The task object version is either unsupported or invalid. 'SCHED_E_UNKNOWN_OBJECT_VERSION
Public Const ERROR_SCHED_E_UNSUPPORTED_ACCOUNT_OPTION                     As Long = &H80041314 'The task has been configured with an unsupported combination of account settings and run time options. 'SCHED_E_UNSUPPORTED_ACCOUNT_OPTION
Public Const ERROR_SCHED_E_SERVICE_NOT_RUNNING                            As Long = &H80041315 'The Task Scheduler Service is not running. 'SCHED_E_SERVICE_NOT_RUNNING
Public Const ERROR_CO_E_CLASS_CREATE_FAILED                               As Long = &H80080001 'Attempt to create a class object failed 'CO_E_CLASS_CREATE_FAILED
Public Const ERROR_CO_E_SCM_ERROR                                         As Long = &H80080002 'OLE service could not bind object 'CO_E_SCM_ERROR
Public Const ERROR_CO_E_SCM_RPC_FAILURE                                   As Long = &H80080003 'RPC communication failed with OLE service 'CO_E_SCM_RPC_FAILURE
Public Const ERROR_CO_E_BAD_PATH                                          As Long = &H80080004 'Bad path to object 'CO_E_BAD_PATH
Public Const ERROR_CO_E_SERVER_EXEC_FAILURE                               As Long = &H80080005 'Server execution failed 'CO_E_SERVER_EXEC_FAILURE
Public Const ERROR_CO_E_OBJSRV_RPC_FAILURE                                As Long = &H80080006 'OLE service could not communicate with the object server 'CO_E_OBJSRV_RPC_FAILURE
Public Const ERROR_MK_E_NO_NORMALIZED                                     As Long = &H80080007 'Moniker path could not be normalized 'MK_E_NO_NORMALIZED
Public Const ERROR_CO_E_SERVER_STOPPING                                   As Long = &H80080008 'Object server is stopping when OLE service contacts it 'CO_E_SERVER_STOPPING
Public Const ERROR_MEM_E_INVALID_ROOT                                     As Long = &H80080009 'An invalid root block pointer was specified 'MEM_E_INVALID_ROOT
Public Const ERROR_MEM_E_INVALID_LINK                                     As Long = &H80080010 'An allocation chain contained an invalid link pointer 'MEM_E_INVALID_LINK
Public Const ERROR_MEM_E_INVALID_SIZE                                     As Long = &H80080011 'The requested allocation size was too large 'MEM_E_INVALID_SIZE
Public Const ERROR_CO_S_NOTALLINTERFACES                                  As Long = &H80012    'Not all the requested interfaces were available 'CO_S_NOTALLINTERFACES
Public Const ERROR_CO_S_MACHINENAMENOTFOUND                               As Long = &H80013    'The specified machine name was not found in the cache. 'CO_S_MACHINENAMENOTFOUND
Public Const ERROR_DISP_E_UNKNOWNINTERFACE                                As Long = &H80020001 'Unknown interface. 'DISP_E_UNKNOWNINTERFACE
Public Const ERROR_DISP_E_MEMBERNOTFOUND                                  As Long = &H80020003 'Member not found. 'DISP_E_MEMBERNOTFOUND
Public Const ERROR_DISP_E_PARAMNOTFOUND                                   As Long = &H80020004 'Parameter not found. 'DISP_E_PARAMNOTFOUND
Public Const ERROR_DISP_E_TYPEMISMATCH                                    As Long = &H80020005 'Type mismatch. 'DISP_E_TYPEMISMATCH
Public Const ERROR_DISP_E_UNKNOWNNAME                                     As Long = &H80020006 'Unknown name. 'DISP_E_UNKNOWNNAME
Public Const ERROR_DISP_E_NONAMEDARGS                                     As Long = &H80020007 'No named arguments. 'DISP_E_NONAMEDARGS
Public Const ERROR_DISP_E_BADVARTYPE                                      As Long = &H80020008 'Bad variable type. 'DISP_E_BADVARTYPE
Public Const ERROR_DISP_E_EXCEPTION                                       As Long = &H80020009 'Exception occurred. 'DISP_E_EXCEPTION
Public Const ERROR_DISP_E_OVERFLOW                                        As Long = &H8002000A 'Out of present range. 'DISP_E_OVERFLOW
Public Const ERROR_DISP_E_BADINDEX                                        As Long = &H8002000B 'Invalid index. 'DISP_E_BADINDEX
Public Const ERROR_DISP_E_UNKNOWNLCID                                     As Long = &H8002000C 'Unknown language. 'DISP_E_UNKNOWNLCID
Public Const ERROR_DISP_E_ARRAYISLOCKED                                   As Long = &H8002000D 'Memory is locked. 'DISP_E_ARRAYISLOCKED
Public Const ERROR_DISP_E_BADPARAMCOUNT                                   As Long = &H8002000E 'Invalid number of parameters. 'DISP_E_BADPARAMCOUNT
Public Const ERROR_DISP_E_PARAMNOTOPTIONAL                                As Long = &H8002000F 'Parameter not optional. 'DISP_E_PARAMNOTOPTIONAL
Public Const ERROR_DISP_E_BADCALLEE                                       As Long = &H80020010 'Invalid callee. 'DISP_E_BADCALLEE
Public Const ERROR_DISP_E_NOTACOLLECTION                                  As Long = &H80020011 'Does not support a collection. 'DISP_E_NOTACOLLECTION
Public Const ERROR_DISP_E_DIVBYZERO                                       As Long = &H80020012 'Division by zero. 'DISP_E_DIVBYZERO
Public Const ERROR_DISP_E_BUFFERTOOSMALL                                  As Long = &H80020013 'Buffer too small 'DISP_E_BUFFERTOOSMALL
Public Const ERROR_TYPE_E_BUFFERTOOSMALL                                  As Long = &H80028016 'Buffer too small. 'TYPE_E_BUFFERTOOSMALL
Public Const ERROR_TYPE_E_FIELDNOTFOUND                                   As Long = &H80028017 'Field name not defined in the record. 'TYPE_E_FIELDNOTFOUND
Public Const ERROR_TYPE_E_INVDATAREAD                                     As Long = &H80028018 'Old format or invalid type library. 'TYPE_E_INVDATAREAD
Public Const ERROR_TYPE_E_UNSUPFORMAT                                     As Long = &H80028019 'Old format or invalid type library. 'TYPE_E_UNSUPFORMAT
Public Const ERROR_TYPE_E_REGISTRYACCESS                                  As Long = &H8002801C 'Error accessing the OLE registry. 'TYPE_E_REGISTRYACCESS
Public Const ERROR_TYPE_E_LIBNOTREGISTERED                                As Long = &H8002801D 'Library not registered. 'TYPE_E_LIBNOTREGISTERED
Public Const ERROR_TYPE_E_UNDEFINEDTYPE                                   As Long = &H80028027 'Bound to unknown type. 'TYPE_E_UNDEFINEDTYPE
Public Const ERROR_TYPE_E_QUALIFIEDNAMEDISALLOWED                         As Long = &H80028028 'Qualified name disallowed. 'TYPE_E_QUALIFIEDNAMEDISALLOWED
Public Const ERROR_TYPE_E_INVALIDSTATE                                    As Long = &H80028029 'Invalid forward reference, or reference to uncompiled type. 'TYPE_E_INVALIDSTATE
Public Const ERROR_TYPE_E_WRONGTYPEKIND                                   As Long = &H8002802A 'Type mismatch. 'TYPE_E_WRONGTYPEKIND
Public Const ERROR_TYPE_E_ELEMENTNOTFOUND                                 As Long = &H8002802B 'Element not found. 'TYPE_E_ELEMENTNOTFOUND
Public Const ERROR_TYPE_E_AMBIGUOUSNAME                                   As Long = &H8002802C 'Ambiguous name. 'TYPE_E_AMBIGUOUSNAME
Public Const ERROR_TYPE_E_NAMECONFLICT                                    As Long = &H8002802D 'Name already exists in the library. 'TYPE_E_NAMECONFLICT
Public Const ERROR_TYPE_E_UNKNOWNLCID                                     As Long = &H8002802E 'Unknown LCID. 'TYPE_E_UNKNOWNLCID
Public Const ERROR_TYPE_E_DLLFUNCTIONNOTFOUND                             As Long = &H8002802F 'Function not defined in specified DLL. 'TYPE_E_DLLFUNCTIONNOTFOUND
Public Const ERROR_TYPE_E_BADMODULEKIND                                   As Long = &H800288BD 'Wrong module kind for the operation. 'TYPE_E_BADMODULEKIND
Public Const ERROR_TYPE_E_SIZETOOBIG                                      As Long = &H800288C5 'Size may not exceed 64K. 'TYPE_E_SIZETOOBIG
Public Const ERROR_TYPE_E_DUPLICATEID                                     As Long = &H800288C6 'Duplicate ID in inheritance hierarchy. 'TYPE_E_DUPLICATEID
Public Const ERROR_TYPE_E_INVALIDID                                       As Long = &H800288CF 'Incorrect inheritance depth in standard OLE hmember. 'TYPE_E_INVALIDID
Public Const ERROR_TYPE_E_TYPEMISMATCH                                    As Long = &H80028CA0 'Type mismatch. 'TYPE_E_TYPEMISMATCH
Public Const ERROR_TYPE_E_OUTOFBOUNDS                                     As Long = &H80028CA1 'Invalid number of arguments. 'TYPE_E_OUTOFBOUNDS
Public Const ERROR_TYPE_E_IOERROR                                         As Long = &H80028CA2 'I/O Error. 'TYPE_E_IOERROR
Public Const ERROR_TYPE_E_CANTCREATETMPFILE                               As Long = &H80028CA3 'Error creating unique tmp file. 'TYPE_E_CANTCREATETMPFILE
Public Const ERROR_TYPE_E_CANTLOADLIBRARY                                 As Long = &H80029C4A 'Error loading type library/DLL. 'TYPE_E_CANTLOADLIBRARY
Public Const ERROR_TYPE_E_INCONSISTENTPROPFUNCS                           As Long = &H80029C83 'Inconsistent property functions. 'TYPE_E_INCONSISTENTPROPFUNCS
Public Const ERROR_TYPE_E_CIRCULARTYPE                                    As Long = &H80029C84 'Circular dependency between types/modules. 'TYPE_E_CIRCULARTYPE
Public Const ERROR_STG_E_INVALIDFUNCTION                                  As Long = &H80030001 'Unable to perform requested operation. 'STG_E_INVALIDFUNCTION
Public Const ERROR_STG_E_FILENOTFOUND                                     As Long = &H80030002 '%1 could not be found. 'STG_E_FILENOTFOUND
Public Const ERROR_STG_E_PATHNOTFOUND                                     As Long = &H80030003 'The path %1 could not be found. 'STG_E_PATHNOTFOUND
Public Const ERROR_STG_E_TOOMANYOPENFILES                                 As Long = &H80030004 'There are insufficient resources to open another file. 'STG_E_TOOMANYOPENFILES
Public Const ERROR_STG_E_ACCESSDENIED                                     As Long = &H80030005 'Access Denied. 'STG_E_ACCESSDENIED
Public Const ERROR_STG_E_INVALIDHANDLE                                    As Long = &H80030006 'Attempted an operation on an invalid object. 'STG_E_INVALIDHANDLE
Public Const ERROR_STG_E_INSUFFICIENTMEMORY                               As Long = &H80030008 'There is insufficient memory available to complete operation. 'STG_E_INSUFFICIENTMEMORY
Public Const ERROR_STG_E_INVALIDPOINTER                                   As Long = &H80030009 'Invalid pointer error. 'STG_E_INVALIDPOINTER
Public Const ERROR_STG_E_NOMOREFILES                                      As Long = &H80030012 'There are no more entries to return. 'STG_E_NOMOREFILES
Public Const ERROR_STG_E_DISKISWRITEPROTECTED                             As Long = &H80030013 'Disk is write-protected. 'STG_E_DISKISWRITEPROTECTED
Public Const ERROR_STG_E_SEEKERROR                                        As Long = &H80030019 'An error occurred during a seek operation. 'STG_E_SEEKERROR
Public Const ERROR_STG_E_WRITEFAULT                                       As Long = &H8003001D 'A disk error occurred during a write operation. 'STG_E_WRITEFAULT
Public Const ERROR_STG_E_READFAULT                                        As Long = &H8003001E 'A disk error occurred during a read operation. 'STG_E_READFAULT
Public Const ERROR_STG_E_SHAREVIOLATION                                   As Long = &H80030020 'A share violation has occurred. 'STG_E_SHAREVIOLATION
Public Const ERROR_STG_E_LOCKVIOLATION                                    As Long = &H80030021 'A lock violation has occurred. 'STG_E_LOCKVIOLATION
Public Const ERROR_STG_E_FILEALREADYEXISTS                                As Long = &H80030050 '%1 already exists. 'STG_E_FILEALREADYEXISTS
Public Const ERROR_STG_E_INVALIDPARAMETER                                 As Long = &H80030057 'Invalid parameter error. 'STG_E_INVALIDPARAMETER
Public Const ERROR_STG_E_MEDIUMFULL                                       As Long = &H80030070 'There is insufficient disk space to complete operation. 'STG_E_MEDIUMFULL
Public Const ERROR_STG_E_PROPSETMISMATCHED                                As Long = &H800300F0 'Illegal write of non-simple property to simple property set. 'STG_E_PROPSETMISMATCHED
Public Const ERROR_STG_E_ABNORMALAPIEXIT                                  As Long = &H800300FA 'An API call exited abnormally. 'STG_E_ABNORMALAPIEXIT
Public Const ERROR_STG_E_INVALIDHEADER                                    As Long = &H800300FB 'The file %1 is not a valid compound file. 'STG_E_INVALIDHEADER
Public Const ERROR_STG_E_INVALIDNAME                                      As Long = &H800300FC 'The name %1 is not valid. 'STG_E_INVALIDNAME
Public Const ERROR_STG_E_UNKNOWN                                          As Long = &H800300FD 'An unexpected error occurred. 'STG_E_UNKNOWN
Public Const ERROR_STG_E_UNIMPLEMENTEDFUNCTION                            As Long = &H800300FE 'That function is not implemented. 'STG_E_UNIMPLEMENTEDFUNCTION
Public Const ERROR_STG_E_INVALIDFLAG                                      As Long = &H800300FF 'Invalid flag error. 'STG_E_INVALIDFLAG
Public Const ERROR_STG_E_INUSE                                            As Long = &H80030100 'Attempted to use an object that is busy. 'STG_E_INUSE
Public Const ERROR_STG_E_NOTCURRENT                                       As Long = &H80030101 'The storage has been changed since the last commit. 'STG_E_NOTCURRENT
Public Const ERROR_STG_E_REVERTED                                         As Long = &H80030102 'Attempted to use an object that has ceased to exist. 'STG_E_REVERTED
Public Const ERROR_STG_E_CANTSAVE                                         As Long = &H80030103 'Can't save. 'STG_E_CANTSAVE
Public Const ERROR_STG_E_OLDFORMAT                                        As Long = &H80030104 'The compound file %1 was produced with an incompatible version of storage. 'STG_E_OLDFORMAT
Public Const ERROR_STG_E_OLDDLL                                           As Long = &H80030105 'The compound file %1 was produced with a newer version of storage. 'STG_E_OLDDLL
Public Const ERROR_STG_E_SHAREREQUIRED                                    As Long = &H80030106 'Share.exe or equivalent is required for operation. 'STG_E_SHAREREQUIRED
Public Const ERROR_STG_E_NOTFILEBASEDSTORAGE                              As Long = &H80030107 'Illegal operation called on non-file based storage. 'STG_E_NOTFILEBASEDSTORAGE
Public Const ERROR_STG_E_EXTANTMARSHALLINGS                               As Long = &H80030108 'Illegal operation called on object with extant marshallings. 'STG_E_EXTANTMARSHALLINGS
Public Const ERROR_STG_E_DOCFILECORRUPT                                   As Long = &H80030109 'The docfile has been corrupted. 'STG_E_DOCFILECORRUPT
Public Const ERROR_STG_E_BADBASEADDRESS                                   As Long = &H80030110 'OLE32.DLL has been loaded at the wrong address. 'STG_E_BADBASEADDRESS
Public Const ERROR_STG_E_DOCFILETOOLARGE                                  As Long = &H80030111 'The compound file is too large for the current implementation 'STG_E_DOCFILETOOLARGE
Public Const ERROR_STG_E_NOTSIMPLEFORMAT                                  As Long = &H80030112 'The compound file was not created with the STGM_SIMPLE flag 'STG_E_NOTSIMPLEFORMAT
Public Const ERROR_STG_E_INCOMPLETE                                       As Long = &H80030201 'The file download was aborted abnormally.  The file is incomplete. 'STG_E_INCOMPLETE
Public Const ERROR_STG_E_TERMINATED                                       As Long = &H80030202 'The file download has been terminated. 'STG_E_TERMINATED
Public Const ERROR_STG_S_CONVERTED                                        As Long = &H30200    'The underlying file was converted to compound file format. 'STG_S_CONVERTED
Public Const ERROR_STG_S_BLOCK                                            As Long = &H30201    'The storage operation should block until more data is available. 'STG_S_BLOCK
Public Const ERROR_STG_S_RETRYNOW                                         As Long = &H30202    'The storage operation should retry immediately. 'STG_S_RETRYNOW
Public Const ERROR_STG_S_MONITORING                                       As Long = &H30203    'The notified event sink will not influence the storage operation. 'STG_S_MONITORING
Public Const ERROR_STG_S_MULTIPLEOPENS                                    As Long = &H30204    'Multiple opens prevent consolidated. (commit succeeded). 'STG_S_MULTIPLEOPENS
Public Const ERROR_STG_S_CONSOLIDATIONFAILED                              As Long = &H30205    'Consolidation of the storage file failed. (commit succeeded). 'STG_S_CONSOLIDATIONFAILED
Public Const ERROR_STG_S_CANNOTCONSOLIDATE                                As Long = &H30206    'Consolidation of the storage file is inappropriate. (commit succeeded). 'STG_S_CANNOTCONSOLIDATE
Public Const ERROR_STG_E_STATUS_COPY_PROTECTION_FAILURE                   As Long = &H80030305 'Generic Copy Protection Error. 'STG_E_STATUS_COPY_PROTECTION_FAILURE
Public Const ERROR_STG_E_CSS_AUTHENTICATION_FAILURE                       As Long = &H80030306 'Copy Protection Error - DVD CSS Authentication failed. 'STG_E_CSS_AUTHENTICATION_FAILURE
Public Const ERROR_STG_E_CSS_KEY_NOT_PRESENT                              As Long = &H80030307 'Copy Protection Error - The given sector does not have a valid CSS key. 'STG_E_CSS_KEY_NOT_PRESENT
Public Const ERROR_STG_E_CSS_KEY_NOT_ESTABLISHED                          As Long = &H80030308 'Copy Protection Error - DVD session key not established. 'STG_E_CSS_KEY_NOT_ESTABLISHED
Public Const ERROR_STG_E_CSS_SCRAMBLED_SECTOR                             As Long = &H80030309 'Copy Protection Error - The read failed because the sector is encrypted. 'STG_E_CSS_SCRAMBLED_SECTOR
Public Const ERROR_STG_E_CSS_REGION_MISMATCH                              As Long = &H8003030A 'Copy Protection Error - The current DVD's region does not correspond to the region setting of the drive. 'STG_E_CSS_REGION_MISMATCH
Public Const ERROR_STG_E_RESETS_EXHAUSTED                                 As Long = &H8003030B 'Copy Protection Error - The drive's region setting may be permanent or the number of user resets has been exhausted. 'STG_E_RESETS_EXHAUSTED
Public Const ERROR_RPC_E_CALL_REJECTED                                    As Long = &H80010001 'Call was rejected by callee. 'RPC_E_CALL_REJECTED
Public Const ERROR_RPC_E_CALL_CANCELED                                    As Long = &H80010002 'Call was canceled by the message filter. 'RPC_E_CALL_CANCELED
Public Const ERROR_RPC_E_CANTPOST_INSENDCALL                              As Long = &H80010003 'The caller is dispatching an intertask SendMessage call and cannot call out via PostMessage. 'RPC_E_CANTPOST_INSENDCALL
Public Const ERROR_RPC_E_CANTCALLOUT_INASYNCCALL                          As Long = &H80010004 'The caller is dispatching an asynchronous call and cannot make an outgoing call on behalf of this call. 'RPC_E_CANTCALLOUT_INASYNCCALL
Public Const ERROR_RPC_E_CANTCALLOUT_INEXTERNALCALL                       As Long = &H80010005 'It is illegal to call out while inside message filter. 'RPC_E_CANTCALLOUT_INEXTERNALCALL
Public Const ERROR_RPC_E_CONNECTION_TERMINATED                            As Long = &H80010006 'The connection terminated or is in a bogus state and cannot be used any more. Other connections are still valid. 'RPC_E_CONNECTION_TERMINATED
Public Const ERROR_RPC_E_SERVER_DIED                                      As Long = &H80010007 'The callee (server [not server application]) is not available and disappeared; all connections are invalid. The call may have executed. 'RPC_E_SERVER_DIED
Public Const ERROR_RPC_E_CLIENT_DIED                                      As Long = &H80010008 'The caller (client) disappeared while the callee (server) was processing a call. 'RPC_E_CLIENT_DIED
Public Const ERROR_RPC_E_INVALID_DATAPACKET                               As Long = &H80010009 'The data packet with the marshalled parameter data is incorrect. 'RPC_E_INVALID_DATAPACKET
Public Const ERROR_RPC_E_CANTTRANSMIT_CALL                                As Long = &H8001000A 'The call was not transmitted properly; the message queue was full and was not emptied after yielding. 'RPC_E_CANTTRANSMIT_CALL
Public Const ERROR_RPC_E_CLIENT_CANTMARSHAL_DATA                          As Long = &H8001000B 'The client (caller) cannot marshall the parameter data - low memory, etc. 'RPC_E_CLIENT_CANTMARSHAL_DATA
Public Const ERROR_RPC_E_CLIENT_CANTUNMARSHAL_DATA                        As Long = &H8001000C 'The client (caller) cannot unmarshall the return data - low memory, etc. 'RPC_E_CLIENT_CANTUNMARSHAL_DATA
Public Const ERROR_RPC_E_SERVER_CANTMARSHAL_DATA                          As Long = &H8001000D 'The server (callee) cannot marshall the return data - low memory, etc. 'RPC_E_SERVER_CANTMARSHAL_DATA
Public Const ERROR_RPC_E_SERVER_CANTUNMARSHAL_DATA                        As Long = &H8001000E 'The server (callee) cannot unmarshall the parameter data - low memory, etc. 'RPC_E_SERVER_CANTUNMARSHAL_DATA
Public Const ERROR_RPC_E_INVALID_DATA                                     As Long = &H8001000F 'Received data is invalid; could be server or client data. 'RPC_E_INVALID_DATA
Public Const ERROR_RPC_E_INVALID_PARAMETER                                As Long = &H80010010 'A particular parameter is invalid and cannot be (un)marshalled. 'RPC_E_INVALID_PARAMETER
Public Const ERROR_RPC_E_CANTCALLOUT_AGAIN                                As Long = &H80010011 'There is no second outgoing call on same channel in DDE conversation. 'RPC_E_CANTCALLOUT_AGAIN
Public Const ERROR_RPC_E_SERVER_DIED_DNE                                  As Long = &H80010012 'The callee (server [not server application]) is not available and disappeared; all connections are invalid. The call did not execute. 'RPC_E_SERVER_DIED_DNE
Public Const ERROR_RPC_E_SYS_CALL_FAILED                                  As Long = &H80010100 'System call failed. 'RPC_E_SYS_CALL_FAILED
Public Const ERROR_RPC_E_OUT_OF_RESOURCES                                 As Long = &H80010101 'Could not allocate some required resource (memory, events, ...) 'RPC_E_OUT_OF_RESOURCES
Public Const ERROR_RPC_E_ATTEMPTED_MULTITHREAD                            As Long = &H80010102 'Attempted to make calls on more than one thread in single threaded mode. 'RPC_E_ATTEMPTED_MULTITHREAD
Public Const ERROR_RPC_E_NOT_REGISTERED                                   As Long = &H80010103 'The requested interface is not registered on the server object. 'RPC_E_NOT_REGISTERED
Public Const ERROR_RPC_E_FAULT                                            As Long = &H80010104 'RPC could not call the server or could not return the results of calling the server. 'RPC_E_FAULT
Public Const ERROR_RPC_E_SERVERFAULT                                      As Long = &H80010105 'The server threw an exception. 'RPC_E_SERVERFAULT
Public Const ERROR_RPC_E_CHANGED_MODE                                     As Long = &H80010106 'Cannot change thread mode after it is set. 'RPC_E_CHANGED_MODE
Public Const ERROR_RPC_E_INVALIDMETHOD                                    As Long = &H80010107 'The method called does not exist on the server. 'RPC_E_INVALIDMETHOD
Public Const ERROR_RPC_E_DISCONNECTED                                     As Long = &H80010108 'The object invoked has disconnected from its clients. 'RPC_E_DISCONNECTED
Public Const ERROR_RPC_E_RETRY                                            As Long = &H80010109 'The object invoked chose not to process the call now.  Try again later. 'RPC_E_RETRY
Public Const ERROR_RPC_E_SERVERCALL_RETRYLATER                            As Long = &H8001010A 'The message filter indicated that the application is busy. 'RPC_E_SERVERCALL_RETRYLATER
Public Const ERROR_RPC_E_SERVERCALL_REJECTED                              As Long = &H8001010B 'The message filter rejected the call. 'RPC_E_SERVERCALL_REJECTED
Public Const ERROR_RPC_E_INVALID_CALLDATA                                 As Long = &H8001010C 'A call control interfaces was called with invalid data. 'RPC_E_INVALID_CALLDATA
Public Const ERROR_RPC_E_CANTCALLOUT_ININPUTSYNCCALL                      As Long = &H8001010D 'An outgoing call cannot be made since the application is dispatching an input-synchronous call. 'RPC_E_CANTCALLOUT_ININPUTSYNCCALL
Public Const ERROR_RPC_E_WRONG_THREAD                                     As Long = &H8001010E 'The application called an interface that was marshalled for a different thread. 'RPC_E_WRONG_THREAD
Public Const ERROR_RPC_E_THREAD_NOT_INIT                                  As Long = &H8001010F 'CoInitialize has not been called on the current thread. 'RPC_E_THREAD_NOT_INIT
Public Const ERROR_RPC_E_VERSION_MISMATCH                                 As Long = &H80010110 'The version of OLE on the client and server machines does not match. 'RPC_E_VERSION_MISMATCH
Public Const ERROR_RPC_E_INVALID_HEADER                                   As Long = &H80010111 'OLE received a packet with an invalid header. 'RPC_E_INVALID_HEADER
Public Const ERROR_RPC_E_INVALID_EXTENSION                                As Long = &H80010112 'OLE received a packet with an invalid extension. 'RPC_E_INVALID_EXTENSION
Public Const ERROR_RPC_E_INVALID_IPID                                     As Long = &H80010113 'The requested object or interface does not exist. 'RPC_E_INVALID_IPID
Public Const ERROR_RPC_E_INVALID_OBJECT                                   As Long = &H80010114 'The requested object does not exist. 'RPC_E_INVALID_OBJECT
Public Const ERROR_RPC_S_CALLPENDING                                      As Long = &H80010115 'OLE has sent a request and is waiting for a reply. 'RPC_S_CALLPENDING
Public Const ERROR_RPC_S_WAITONTIMER                                      As Long = &H80010116 'OLE is waiting before retrying a request. 'RPC_S_WAITONTIMER
Public Const ERROR_RPC_E_CALL_COMPLETE                                    As Long = &H80010117 'Call context cannot be accessed after call completed. 'RPC_E_CALL_COMPLETE
Public Const ERROR_RPC_E_UNSECURE_CALL                                    As Long = &H80010118 'Impersonate on unsecure calls is not supported. 'RPC_E_UNSECURE_CALL
Public Const ERROR_RPC_E_TOO_LATE                                         As Long = &H80010119 'Security must be initialized before any interfaces are marshalled or unmarshalled. It cannot be changed once initialized. 'RPC_E_TOO_LATE
Public Const ERROR_RPC_E_NO_GOOD_SECURITY_PACKAGES                        As Long = &H8001011A 'No security packages are installed on this machine or the user is not logged on or there are no compatible security packages between the client and server. 'RPC_E_NO_GOOD_SECURITY_PACKAGES
Public Const ERROR_RPC_E_ACCESS_DENIED                                    As Long = &H8001011B 'Access is denied. 'RPC_E_ACCESS_DENIED
Public Const ERROR_RPC_E_REMOTE_DISABLED                                  As Long = &H8001011C 'Remote calls are not allowed for this process. 'RPC_E_REMOTE_DISABLED
Public Const ERROR_RPC_E_INVALID_OBJREF                                   As Long = &H8001011D 'The marshaled interface data packet (OBJREF) has an invalid or unknown format. 'RPC_E_INVALID_OBJREF
Public Const ERROR_RPC_E_NO_CONTEXT                                       As Long = &H8001011E 'No context is associated with this call. This happens for some custom marshalled calls and on the client side of the call. 'RPC_E_NO_CONTEXT
Public Const ERROR_RPC_E_TIMEOUT                                          As Long = &H8001011F 'This operation returned because the timeout period expired. 'RPC_E_TIMEOUT
Public Const ERROR_RPC_E_NO_SYNC                                          As Long = &H80010120 'There are no synchronize objects to wait on. 'RPC_E_NO_SYNC
Public Const ERROR_RPC_E_FULLSIC_REQUIRED                                 As Long = &H80010121 'Full subject issuer chain SSL principal name expected from the server. 'RPC_E_FULLSIC_REQUIRED
Public Const ERROR_RPC_E_INVALID_STD_NAME                                 As Long = &H80010122 'Principal name is not a valid MSSTD name. 'RPC_E_INVALID_STD_NAME
Public Const ERROR_CO_E_FAILEDTOIMPERSONATE                               As Long = &H80010123 'Unable to impersonate DCOM client 'CO_E_FAILEDTOIMPERSONATE
Public Const ERROR_CO_E_FAILEDTOGETSECCTX                                 As Long = &H80010124 'Unable to obtain server's security context 'CO_E_FAILEDTOGETSECCTX
Public Const ERROR_CO_E_FAILEDTOOPENTHREADTOKEN                           As Long = &H80010125 'Unable to open the access token of the current thread 'CO_E_FAILEDTOOPENTHREADTOKEN
Public Const ERROR_CO_E_FAILEDTOGETTOKENINFO                              As Long = &H80010126 'Unable to obtain user info from an access token 'CO_E_FAILEDTOGETTOKENINFO
Public Const ERROR_CO_E_TRUSTEEDOESNTMATCHCLIENT                          As Long = &H80010127 'The client who called IAccessControl::IsAccessPermitted was not the trustee provided to the method 'CO_E_TRUSTEEDOESNTMATCHCLIENT
Public Const ERROR_CO_E_FAILEDTOQUERYCLIENTBLANKET                        As Long = &H80010128 'Unable to obtain the client's security blanket 'CO_E_FAILEDTOQUERYCLIENTBLANKET
Public Const ERROR_CO_E_FAILEDTOSETDACL                                   As Long = &H80010129 'Unable to set a discretionary ACL into a security descriptor 'CO_E_FAILEDTOSETDACL
Public Const ERROR_CO_E_ACCESSCHECKFAILED                                 As Long = &H8001012A 'The system function, AccessCheck, returned false 'CO_E_ACCESSCHECKFAILED
Public Const ERROR_CO_E_NETACCESSAPIFAILED                                As Long = &H8001012B 'Either NetAccessDel or NetAccessAdd returned an error code. 'CO_E_NETACCESSAPIFAILED
Public Const ERROR_CO_E_WRONGTRUSTEENAMESYNTAX                            As Long = &H8001012C 'One of the trustee strings provided by the user did not conform to the <Domain>\<Name> syntax and it was not the "*" string 'CO_E_WRONGTRUSTEENAMESYNTAX
Public Const ERROR_CO_E_INVALIDSID                                        As Long = &H8001012D 'One of the security identifiers provided by the user was invalid 'CO_E_INVALIDSID
Public Const ERROR_CO_E_CONVERSIONFAILED                                  As Long = &H8001012E 'Unable to convert a wide character trustee string to a multibyte trustee string 'CO_E_CONVERSIONFAILED
Public Const ERROR_CO_E_NOMATCHINGSIDFOUND                                As Long = &H8001012F 'Unable to find a security identifier that corresponds to a trustee string provided by the user 'CO_E_NOMATCHINGSIDFOUND
Public Const ERROR_CO_E_LOOKUPACCSIDFAILED                                As Long = &H80010130 'The system function, LookupAccountSID, failed 'CO_E_LOOKUPACCSIDFAILED
Public Const ERROR_CO_E_NOMATCHINGNAMEFOUND                               As Long = &H80010131 'Unable to find a trustee name that corresponds to a security identifier provided by the user 'CO_E_NOMATCHINGNAMEFOUND
Public Const ERROR_CO_E_LOOKUPACCNAMEFAILED                               As Long = &H80010132 'The system function, LookupAccountName, failed 'CO_E_LOOKUPACCNAMEFAILED
Public Const ERROR_CO_E_SETSERLHNDLFAILED                                 As Long = &H80010133 'Unable to set or reset a serialization handle 'CO_E_SETSERLHNDLFAILED
Public Const ERROR_CO_E_FAILEDTOGETWINDIR                                 As Long = &H80010134 'Unable to obtain the Windows directory 'CO_E_FAILEDTOGETWINDIR
Public Const ERROR_CO_E_PATHTOOLONG                                       As Long = &H80010135 'Path too long 'CO_E_PATHTOOLONG
Public Const ERROR_CO_E_FAILEDTOGENUUID                                   As Long = &H80010136 'Unable to generate a uuid. 'CO_E_FAILEDTOGENUUID
Public Const ERROR_CO_E_FAILEDTOCREATEFILE                                As Long = &H80010137 'Unable to create file 'CO_E_FAILEDTOCREATEFILE
Public Const ERROR_CO_E_FAILEDTOCLOSEHANDLE                               As Long = &H80010138 'Unable to close a serialization handle or a file handle. 'CO_E_FAILEDTOCLOSEHANDLE
Public Const ERROR_CO_E_EXCEEDSYSACLLIMIT                                 As Long = &H80010139 'The number of ACEs in an ACL exceeds the system limit. 'CO_E_EXCEEDSYSACLLIMIT
Public Const ERROR_CO_E_ACESINWRONGORDER                                  As Long = &H8001013A 'Not all the DENY_ACCESS ACEs are arranged in front of the GRANT_ACCESS ACEs in the stream. 'CO_E_ACESINWRONGORDER
Public Const ERROR_CO_E_INCOMPATIBLESTREAMVERSION                         As Long = &H8001013B 'The version of ACL format in the stream is not supported by this implementation of IAccessControl 'CO_E_INCOMPATIBLESTREAMVERSION
Public Const ERROR_CO_E_FAILEDTOOPENPROCESSTOKEN                          As Long = &H8001013C 'Unable to open the access token of the server process 'CO_E_FAILEDTOOPENPROCESSTOKEN
Public Const ERROR_CO_E_DECODEFAILED                                      As Long = &H8001013D 'Unable to decode the ACL in the stream provided by the user 'CO_E_DECODEFAILED
Public Const ERROR_CO_E_ACNOTINITIALIZED                                  As Long = &H8001013F 'The COM IAccessControl object is not initialized 'CO_E_ACNOTINITIALIZED
Public Const ERROR_CO_E_CANCEL_DISABLED                                   As Long = &H80010140 'Call Cancellation is disabled 'CO_E_CANCEL_DISABLED
Public Const ERROR_RPC_E_UNEXPECTED                                       As Long = &H8001FFFF 'An internal error occurred. 'RPC_E_UNEXPECTED
Public Const ERROR_AUDITING_DISABLED                                      As Long = &HC0090001 'The specified event is currently not being audited. 'ERROR_AUDITING_DISABLED
Public Const ERROR_ALL_SIDS_FILTERED                                      As Long = &HC0090002 'The SID filtering operation removed all SIDs. 'ERROR_ALL_SIDS_FILTERED
Public Const ERROR_NTE_BAD_UID                                            As Long = &H80090001 'Bad UID. 'NTE_BAD_UID
Public Const ERROR_NTE_BAD_HASH                                           As Long = &H80090002 'Bad Hash. 'NTE_BAD_HASH
Public Const ERROR_NTE_BAD_KEY                                            As Long = &H80090003 'Bad Key. 'NTE_BAD_KEY
Public Const ERROR_NTE_BAD_LEN                                            As Long = &H80090004 'Bad Length. 'NTE_BAD_LEN
Public Const ERROR_NTE_BAD_DATA                                           As Long = &H80090005 'Bad Data. 'NTE_BAD_DATA
Public Const ERROR_NTE_BAD_SIGNATURE                                      As Long = &H80090006 'Invalid Signature. 'NTE_BAD_SIGNATURE
Public Const ERROR_NTE_BAD_VER                                            As Long = &H80090007 'Bad Version of provider. 'NTE_BAD_VER
Public Const ERROR_NTE_BAD_ALGID                                          As Long = &H80090008 'Invalid algorithm specified. 'NTE_BAD_ALGID
Public Const ERROR_NTE_BAD_FLAGS                                          As Long = &H80090009 'Invalid flags specified. 'NTE_BAD_FLAGS
Public Const ERROR_NTE_BAD_TYPE                                           As Long = &H8009000A 'Invalid type specified. 'NTE_BAD_TYPE
Public Const ERROR_NTE_BAD_KEY_STATE                                      As Long = &H8009000B 'Key not valid for use in specified state. 'NTE_BAD_KEY_STATE
Public Const ERROR_NTE_BAD_HASH_STATE                                     As Long = &H8009000C 'Hash not valid for use in specified state. 'NTE_BAD_HASH_STATE
Public Const ERROR_NTE_NO_KEY                                             As Long = &H8009000D 'Key does not exist. 'NTE_NO_KEY
Public Const ERROR_NTE_NO_MEMORY                                          As Long = &H8009000E 'Insufficient memory available for the operation. 'NTE_NO_MEMORY
Public Const ERROR_NTE_EXISTS                                             As Long = &H8009000F 'Object already exists. 'NTE_EXISTS
Public Const ERROR_NTE_PERM                                               As Long = &H80090010 'Access denied. 'NTE_PERM
Public Const ERROR_NTE_NOT_FOUND                                          As Long = &H80090011 'Object was not found. 'NTE_NOT_FOUND
Public Const ERROR_NTE_DOUBLE_ENCRYPT                                     As Long = &H80090012 'Data already encrypted. 'NTE_DOUBLE_ENCRYPT
Public Const ERROR_NTE_BAD_PROVIDER                                       As Long = &H80090013 'Invalid provider specified. 'NTE_BAD_PROVIDER
Public Const ERROR_NTE_BAD_PROV_TYPE                                      As Long = &H80090014 'Invalid provider type specified. 'NTE_BAD_PROV_TYPE
Public Const ERROR_NTE_BAD_PUBLIC_KEY                                     As Long = &H80090015 'Provider's public key is invalid. 'NTE_BAD_PUBLIC_KEY
Public Const ERROR_NTE_BAD_KEYSET                                         As Long = &H80090016 'Keyset does not exist 'NTE_BAD_KEYSET
Public Const ERROR_NTE_PROV_TYPE_NOT_DEF                                  As Long = &H80090017 'Provider type not defined. 'NTE_PROV_TYPE_NOT_DEF
Public Const ERROR_NTE_PROV_TYPE_ENTRY_BAD                                As Long = &H80090018 'Provider type as registered is invalid. 'NTE_PROV_TYPE_ENTRY_BAD
Public Const ERROR_NTE_KEYSET_NOT_DEF                                     As Long = &H80090019 'The keyset is not defined. 'NTE_KEYSET_NOT_DEF
Public Const ERROR_NTE_KEYSET_ENTRY_BAD                                   As Long = &H8009001A 'Keyset as registered is invalid. 'NTE_KEYSET_ENTRY_BAD
Public Const ERROR_NTE_PROV_TYPE_NO_MATCH                                 As Long = &H8009001B 'Provider type does not match registered value. 'NTE_PROV_TYPE_NO_MATCH
Public Const ERROR_NTE_SIGNATURE_FILE_BAD                                 As Long = &H8009001C 'The digital signature file is corrupt. 'NTE_SIGNATURE_FILE_BAD
Public Const ERROR_NTE_PROVIDER_DLL_FAIL                                  As Long = &H8009001D 'Provider DLL failed to initialize correctly. 'NTE_PROVIDER_DLL_FAIL
Public Const ERROR_NTE_PROV_DLL_NOT_FOUND                                 As Long = &H8009001E 'Provider DLL could not be found. 'NTE_PROV_DLL_NOT_FOUND
Public Const ERROR_NTE_BAD_KEYSET_PARAM                                   As Long = &H8009001F 'The Keyset parameter is invalid. 'NTE_BAD_KEYSET_PARAM
Public Const ERROR_NTE_FAIL                                               As Long = &H80090020 'An internal error occurred. 'NTE_FAIL
Public Const ERROR_NTE_SYS_ERR                                            As Long = &H80090021 'A base error occurred. 'NTE_SYS_ERR
Public Const ERROR_NTE_SILENT_CONTEXT                                     As Long = &H80090022 'Provider could not perform the action since the context was acquired as silent. 'NTE_SILENT_CONTEXT
Public Const ERROR_NTE_TOKEN_KEYSET_STORAGE_FULL                          As Long = &H80090023 'The security token does not have storage space available for an additional container. 'NTE_TOKEN_KEYSET_STORAGE_FULL
Public Const ERROR_NTE_TEMPORARY_PROFILE                                  As Long = &H80090024 'The profile for the user is a temporary profile. 'NTE_TEMPORARY_PROFILE
Public Const ERROR_NTE_FIXEDPARAMETER                                     As Long = &H80090025 'The key parameters could not be set because the CSP uses fixed parameters. 'NTE_FIXEDPARAMETER
Public Const ERROR_SEC_E_INSUFFICIENT_MEMORY                              As Long = &H80090300 'Not enough memory is available to complete this request 'SEC_E_INSUFFICIENT_MEMORY
Public Const ERROR_SEC_E_INVALID_HANDLE                                   As Long = &H80090301 'The handle specified is invalid 'SEC_E_INVALID_HANDLE
Public Const ERROR_SEC_E_UNSUPPORTED_FUNCTION                             As Long = &H80090302 'The function requested is not supported 'SEC_E_UNSUPPORTED_FUNCTION
Public Const ERROR_SEC_E_TARGET_UNKNOWN                                   As Long = &H80090303 'The specified target is unknown or unreachable 'SEC_E_TARGET_UNKNOWN
Public Const ERROR_SEC_E_INTERNAL_ERROR                                   As Long = &H80090304 'The Local Security Authority cannot be contacted 'SEC_E_INTERNAL_ERROR
Public Const ERROR_SEC_E_SECPKG_NOT_FOUND                                 As Long = &H80090305 'The requested security package does not exist 'SEC_E_SECPKG_NOT_FOUND
Public Const ERROR_SEC_E_NOT_OWNER                                        As Long = &H80090306 'The caller is not the owner of the desired credentials 'SEC_E_NOT_OWNER
Public Const ERROR_SEC_E_CANNOT_INSTALL                                   As Long = &H80090307 'The security package failed to initialize, and cannot be installed 'SEC_E_CANNOT_INSTALL
Public Const ERROR_SEC_E_INVALID_TOKEN                                    As Long = &H80090308 'The token supplied to the function is invalid 'SEC_E_INVALID_TOKEN
Public Const ERROR_SEC_E_CANNOT_PACK                                      As Long = &H80090309 'The security package is not able to marshall the logon buffer, so the logon attempt has failed 'SEC_E_CANNOT_PACK
Public Const ERROR_SEC_E_QOP_NOT_SUPPORTED                                As Long = &H8009030A 'The per-message Quality of Protection is not supported by the security package 'SEC_E_QOP_NOT_SUPPORTED
Public Const ERROR_SEC_E_NO_IMPERSONATION                                 As Long = &H8009030B 'The security context does not allow impersonation of the client 'SEC_E_NO_IMPERSONATION
Public Const ERROR_SEC_E_LOGON_DENIED                                     As Long = &H8009030C 'The logon attempt failed 'SEC_E_LOGON_DENIED
Public Const ERROR_SEC_E_UNKNOWN_CREDENTIALS                              As Long = &H8009030D 'The credentials supplied to the package were not recognized 'SEC_E_UNKNOWN_CREDENTIALS
Public Const ERROR_SEC_E_NO_CREDENTIALS                                   As Long = &H8009030E 'No credentials are available in the security package 'SEC_E_NO_CREDENTIALS
Public Const ERROR_SEC_E_MESSAGE_ALTERED                                  As Long = &H8009030F 'The message or signature supplied for verification has been altered 'SEC_E_MESSAGE_ALTERED
Public Const ERROR_SEC_E_OUT_OF_SEQUENCE                                  As Long = &H80090310 'The message supplied for verification is out of sequence 'SEC_E_OUT_OF_SEQUENCE
Public Const ERROR_SEC_E_NO_AUTHENTICATING_AUTHORITY                      As Long = &H80090311 'No authority could be contacted for authentication. 'SEC_E_NO_AUTHENTICATING_AUTHORITY
Public Const ERROR_SEC_I_CONTINUE_NEEDED                                  As Long = &H90312    'The function completed successfully, but must be called again to complete the context 'SEC_I_CONTINUE_NEEDED
Public Const ERROR_SEC_I_COMPLETE_NEEDED                                  As Long = &H90313    'The function completed successfully, but CompleteToken must be called 'SEC_I_COMPLETE_NEEDED
Public Const ERROR_SEC_I_COMPLETE_AND_CONTINUE                            As Long = &H90314    'The function completed successfully, but both CompleteToken and this function must be called to complete the context 'SEC_I_COMPLETE_AND_CONTINUE
Public Const ERROR_SEC_I_LOCAL_LOGON                                      As Long = &H90315    'The logon was completed, but no network authority was available. The logon was made using locally known information 'SEC_I_LOCAL_LOGON
Public Const ERROR_SEC_E_BAD_PKGID                                        As Long = &H80090316 'The requested security package does not exist 'SEC_E_BAD_PKGID
Public Const ERROR_SEC_E_CONTEXT_EXPIRED                                  As Long = &H80090317 'The context has expired and can no longer be used. 'SEC_E_CONTEXT_EXPIRED
Public Const ERROR_SEC_I_CONTEXT_EXPIRED                                  As Long = &H90317    'The context has expired and can no longer be used. 'SEC_I_CONTEXT_EXPIRED
Public Const ERROR_SEC_E_INCOMPLETE_MESSAGE                               As Long = &H80090318 'The supplied message is incomplete.  The signature was not verified. 'SEC_E_INCOMPLETE_MESSAGE
Public Const ERROR_SEC_E_INCOMPLETE_CREDENTIALS                           As Long = &H80090320 'The credentials supplied were not complete, and could not be verified. The context could not be initialized. 'SEC_E_INCOMPLETE_CREDENTIALS
Public Const ERROR_SEC_E_BUFFER_TOO_SMALL                                 As Long = &H80090321 'The buffers supplied to a function was too small. 'SEC_E_BUFFER_TOO_SMALL
Public Const ERROR_SEC_I_INCOMPLETE_CREDENTIALS                           As Long = &H90320    'The credentials supplied were not complete, and could not be verified. Additional information can be returned from the context. 'SEC_I_INCOMPLETE_CREDENTIALS
Public Const ERROR_SEC_I_RENEGOTIATE                                      As Long = &H90321    'The context data must be renegotiated with the peer. 'SEC_I_RENEGOTIATE
Public Const ERROR_SEC_E_WRONG_PRINCIPAL                                  As Long = &H80090322 'The target principal name is incorrect. 'SEC_E_WRONG_PRINCIPAL
Public Const ERROR_SEC_I_NO_LSA_CONTEXT                                   As Long = &H90323    'There is no LSA mode context associated with this context. 'SEC_I_NO_LSA_CONTEXT
Public Const ERROR_SEC_E_TIME_SKEW                                        As Long = &H80090324 'The clocks on the client and server machines are skewed. 'SEC_E_TIME_SKEW
Public Const ERROR_SEC_E_UNTRUSTED_ROOT                                   As Long = &H80090325 'The certificate chain was issued by an authority that is not trusted. 'SEC_E_UNTRUSTED_ROOT
Public Const ERROR_SEC_E_ILLEGAL_MESSAGE                                  As Long = &H80090326 'The message received was unexpected or badly formatted. 'SEC_E_ILLEGAL_MESSAGE
Public Const ERROR_SEC_E_CERT_UNKNOWN                                     As Long = &H80090327 'An unknown error occurred while processing the certificate. 'SEC_E_CERT_UNKNOWN
Public Const ERROR_SEC_E_CERT_EXPIRED                                     As Long = &H80090328 'The received certificate has expired. 'SEC_E_CERT_EXPIRED
Public Const ERROR_SEC_E_ENCRYPT_FAILURE                                  As Long = &H80090329 'The specified data could not be encrypted. 'SEC_E_ENCRYPT_FAILURE
Public Const ERROR_SEC_E_DECRYPT_FAILURE                                  As Long = &H80090330 'The specified data could not be decrypted. 'SEC_E_DECRYPT_FAILURE
Public Const ERROR_SEC_E_ALGORITHM_MISMATCH                               As Long = &H80090331 'The client and server cannot communicate, because they do not possess a common algorithm. 'SEC_E_ALGORITHM_MISMATCH
Public Const ERROR_SEC_E_SECURITY_QOS_FAILED                              As Long = &H80090332 'The security context could not be established due to a failure in the requested quality of service (e.g. mutual authentication or delegation). 'SEC_E_SECURITY_QOS_FAILED
Public Const ERROR_SEC_E_UNFINISHED_CONTEXT_DELETED                       As Long = &H80090333 'A security context was deleted before the context was completed.  This is considered a logon failure. 'SEC_E_UNFINISHED_CONTEXT_DELETED
Public Const ERROR_SEC_E_NO_TGT_REPLY                                     As Long = &H80090334 'The client is trying to negotiate a context and the server requires user-to-user but didn't send a TGT reply. 'SEC_E_NO_TGT_REPLY
Public Const ERROR_SEC_E_NO_IP_ADDRESSES                                  As Long = &H80090335 'Unable to accomplish the requested task because the local machine does not have any IP addresses. 'SEC_E_NO_IP_ADDRESSES
Public Const ERROR_SEC_E_WRONG_CREDENTIAL_HANDLE                          As Long = &H80090336 'The supplied credential handle does not match the credential associated with the security context. 'SEC_E_WRONG_CREDENTIAL_HANDLE
Public Const ERROR_SEC_E_CRYPTO_SYSTEM_INVALID                            As Long = &H80090337 'The crypto system or checksum function is invalid because a required function is unavailable. 'SEC_E_CRYPTO_SYSTEM_INVALID
Public Const ERROR_SEC_E_MAX_REFERRALS_EXCEEDED                           As Long = &H80090338 'The number of maximum ticket referrals has been exceeded. 'SEC_E_MAX_REFERRALS_EXCEEDED
Public Const ERROR_SEC_E_MUST_BE_KDC                                      As Long = &H80090339 'The local machine must be a Kerberos KDC (domain controller) and it is not. 'SEC_E_MUST_BE_KDC
Public Const ERROR_SEC_E_STRONG_CRYPTO_NOT_SUPPORTED                      As Long = &H8009033A 'The other end of the security negotiation is requires strong crypto but it is not supported on the local machine. 'SEC_E_STRONG_CRYPTO_NOT_SUPPORTED
Public Const ERROR_SEC_E_TOO_MANY_PRINCIPALS                              As Long = &H8009033B 'The KDC reply contained more than one principal name. 'SEC_E_TOO_MANY_PRINCIPALS
Public Const ERROR_SEC_E_NO_PA_DATA                                       As Long = &H8009033C 'Expected to find PA data for a hint of what etype to use, but it was not found. 'SEC_E_NO_PA_DATA
Public Const ERROR_SEC_E_PKINIT_NAME_MISMATCH                             As Long = &H8009033D 'The client certificate does not contain a valid UPN, or does not match the client name 'SEC_E_PKINIT_NAME_MISMATCH
Public Const ERROR_SEC_E_SMARTCARD_LOGON_REQUIRED                         As Long = &H8009033E 'Smartcard logon is required and was not used. 'SEC_E_SMARTCARD_LOGON_REQUIRED
Public Const ERROR_SEC_E_SHUTDOWN_IN_PROGRESS                             As Long = &H8009033F 'A system shutdown is in progress. 'SEC_E_SHUTDOWN_IN_PROGRESS
Public Const ERROR_SEC_E_KDC_INVALID_REQUEST                              As Long = &H80090340 'An invalid request was sent to the KDC. 'SEC_E_KDC_INVALID_REQUEST
Public Const ERROR_SEC_E_KDC_UNABLE_TO_REFER                              As Long = &H80090341 'The KDC was unable to generate a referral for the service requested. 'SEC_E_KDC_UNABLE_TO_REFER
Public Const ERROR_SEC_E_KDC_UNKNOWN_ETYPE                                As Long = &H80090342 'The encryption type requested is not supported by the KDC. 'SEC_E_KDC_UNKNOWN_ETYPE
Public Const ERROR_SEC_E_UNSUPPORTED_PREAUTH                              As Long = &H80090343 'An unsupported preauthentication mechanism was presented to the kerberos package. 'SEC_E_UNSUPPORTED_PREAUTH
Public Const ERROR_SEC_E_DELEGATION_REQUIRED                              As Long = &H80090345 'The requested operation cannot be completed.  The computer must be trusted for delegation and the current user account must be configured to allow delegation. 'SEC_E_DELEGATION_REQUIRED
Public Const ERROR_SEC_E_BAD_BINDINGS                                     As Long = &H80090346 'Client's supplied SSPI channel bindings were incorrect. 'SEC_E_BAD_BINDINGS
Public Const ERROR_SEC_E_MULTIPLE_ACCOUNTS                                As Long = &H80090347 'The received certificate was mapped to multiple accounts. 'SEC_E_MULTIPLE_ACCOUNTS
Public Const ERROR_SEC_E_NO_KERB_KEY                                      As Long = &H80090348 'SEC_E_NO_KERB_KEY 'SEC_E_NO_KERB_KEY
Public Const ERROR_SEC_E_CERT_WRONG_USAGE                                 As Long = &H80090349 'The certificate is not valid for the requested usage. 'SEC_E_CERT_WRONG_USAGE
Public Const ERROR_SEC_E_DOWNGRADE_DETECTED                               As Long = &H80090350 'The system detected a possible attempt to compromise security.  Please ensure that you can contact the server that authenticated you. 'SEC_E_DOWNGRADE_DETECTED
Public Const ERROR_SEC_E_SMARTCARD_CERT_REVOKED                           As Long = &H80090351 'The smartcard certificate used for authentication has been revoked. 'SEC_E_SMARTCARD_CERT_REVOKED
Public Const ERROR_SEC_E_ISSUING_CA_UNTRUSTED                             As Long = &H80090352 'An untrusted certificate authority was detected While processing the 'SEC_E_ISSUING_CA_UNTRUSTED
Public Const ERROR_SEC_E_REVOCATION_OFFLINE_C                             As Long = &H80090353 'The revocation status of the smartcard certificate used for 'SEC_E_REVOCATION_OFFLINE_C
Public Const ERROR_SEC_E_PKINIT_CLIENT_FAILURE                            As Long = &H80090354 'The smartcard certificate used for authentication was not trusted.  Please 'SEC_E_PKINIT_CLIENT_FAILURE
Public Const ERROR_SEC_E_SMARTCARD_CERT_EXPIRED                           As Long = &H80090355 'The smartcard certificate used for authentication has expired.  Please 'SEC_E_SMARTCARD_CERT_EXPIRED
Public Const ERROR_SEC_E_NO_S4U_PROT_SUPPORT                              As Long = &H80090356 'The Kerberos subsystem encountered an error.  A service for user protocol request was made 'SEC_E_NO_S4U_PROT_SUPPORT
Public Const ERROR_SEC_E_CROSSREALM_DELEGATION_FAILURE                    As Long = &H80090357 'An attempt was made by this server to make a Kerberos constrained delegation request for a target 'SEC_E_CROSSREALM_DELEGATION_FAILURE
Public Const ERROR_SEC_E_REVOCATION_OFFLINE_KDC                           As Long = &H80090358 'The revocation status of the domain controller certificate used for smartcard 'SEC_E_REVOCATION_OFFLINE_KDC
Public Const ERROR_SEC_E_ISSUING_CA_UNTRUSTED_KDC                         As Long = &H80090359 'An untrusted certificate authority was detected while processing the 'SEC_E_ISSUING_CA_UNTRUSTED_KDC
Public Const ERROR_SEC_E_KDC_CERT_EXPIRED                                 As Long = &H8009035A 'The domain controller certificate used for smartcard logon has expired. 'SEC_E_KDC_CERT_EXPIRED
Public Const ERROR_SEC_E_KDC_CERT_REVOKED                                 As Long = &H8009035B 'The domain controller certificate used for smartcard logon has been revoked. 'SEC_E_KDC_CERT_REVOKED
Public Const ERROR_CRYPT_E_MSG_ERROR                                      As Long = &H80091001 'An error occurred while performing an operation on a cryptographic message. 'CRYPT_E_MSG_ERROR
Public Const ERROR_CRYPT_E_UNKNOWN_ALGO                                   As Long = &H80091002 'Unknown cryptographic algorithm. 'CRYPT_E_UNKNOWN_ALGO
Public Const ERROR_CRYPT_E_OID_FORMAT                                     As Long = &H80091003 'The object identifier is poorly formatted. 'CRYPT_E_OID_FORMAT
Public Const ERROR_CRYPT_E_INVALID_MSG_TYPE                               As Long = &H80091004 'Invalid cryptographic message type. 'CRYPT_E_INVALID_MSG_TYPE
Public Const ERROR_CRYPT_E_UNEXPECTED_ENCODING                            As Long = &H80091005 'Unexpected cryptographic message encoding. 'CRYPT_E_UNEXPECTED_ENCODING
Public Const ERROR_CRYPT_E_AUTH_ATTR_MISSING                              As Long = &H80091006 'The cryptographic message does not contain an expected authenticated attribute. 'CRYPT_E_AUTH_ATTR_MISSING
Public Const ERROR_CRYPT_E_HASH_VALUE                                     As Long = &H80091007 'The hash value is not correct. 'CRYPT_E_HASH_VALUE
Public Const ERROR_CRYPT_E_INVALID_INDEX                                  As Long = &H80091008 'The index value is not valid. 'CRYPT_E_INVALID_INDEX
Public Const ERROR_CRYPT_E_ALREADY_DECRYPTED                              As Long = &H80091009 'The content of the cryptographic message has already been decrypted. 'CRYPT_E_ALREADY_DECRYPTED
Public Const ERROR_CRYPT_E_NOT_DECRYPTED                                  As Long = &H8009100A 'The content of the cryptographic message has not been decrypted yet. 'CRYPT_E_NOT_DECRYPTED
Public Const ERROR_CRYPT_E_RECIPIENT_NOT_FOUND                            As Long = &H8009100B 'The enveloped-data message does not contain the specified recipient. 'CRYPT_E_RECIPIENT_NOT_FOUND
Public Const ERROR_CRYPT_E_CONTROL_TYPE                                   As Long = &H8009100C 'Invalid control type. 'CRYPT_E_CONTROL_TYPE
Public Const ERROR_CRYPT_E_ISSUER_SERIALNUMBER                            As Long = &H8009100D 'Invalid issuer and/or serial number. 'CRYPT_E_ISSUER_SERIALNUMBER
Public Const ERROR_CRYPT_E_SIGNER_NOT_FOUND                               As Long = &H8009100E 'Cannot find the original signer. 'CRYPT_E_SIGNER_NOT_FOUND
Public Const ERROR_CRYPT_E_ATTRIBUTES_MISSING                             As Long = &H8009100F 'The cryptographic message does not contain all of the requested attributes. 'CRYPT_E_ATTRIBUTES_MISSING
Public Const ERROR_CRYPT_E_STREAM_MSG_NOT_READY                           As Long = &H80091010 'The streamed cryptographic message is not ready to return data. 'CRYPT_E_STREAM_MSG_NOT_READY
Public Const ERROR_CRYPT_E_STREAM_INSUFFICIENT_DATA                       As Long = &H80091011 'The streamed cryptographic message requires more data to complete the decode operation. 'CRYPT_E_STREAM_INSUFFICIENT_DATA
Public Const ERROR_CRYPT_I_NEW_PROTECTION_REQUIRED                        As Long = &H91012    'The protected data needs to be re-protected. 'CRYPT_I_NEW_PROTECTION_REQUIRED
Public Const ERROR_CRYPT_E_BAD_LEN                                        As Long = &H80092001 'The length specified for the output data was insufficient. 'CRYPT_E_BAD_LEN
Public Const ERROR_CRYPT_E_BAD_ENCODE                                     As Long = &H80092002 'An error occurred during encode or decode operation. 'CRYPT_E_BAD_ENCODE
Public Const ERROR_CRYPT_E_FILE_ERROR                                     As Long = &H80092003 'An error occurred while reading or writing to a file. 'CRYPT_E_FILE_ERROR
Public Const ERROR_CRYPT_E_NOT_FOUND                                      As Long = &H80092004 'Cannot find object or property. 'CRYPT_E_NOT_FOUND
Public Const ERROR_CRYPT_E_EXISTS                                         As Long = &H80092005 'The object or property already exists. 'CRYPT_E_EXISTS
Public Const ERROR_CRYPT_E_NO_PROVIDER                                    As Long = &H80092006 'No provider was specified for the store or object. 'CRYPT_E_NO_PROVIDER
Public Const ERROR_CRYPT_E_SELF_SIGNED                                    As Long = &H80092007 'The specified certificate is self signed. 'CRYPT_E_SELF_SIGNED
Public Const ERROR_CRYPT_E_DELETED_PREV                                   As Long = &H80092008 'The previous certificate or CRL context was deleted. 'CRYPT_E_DELETED_PREV
Public Const ERROR_CRYPT_E_NO_MATCH                                       As Long = &H80092009 'Cannot find the requested object. 'CRYPT_E_NO_MATCH
Public Const ERROR_CRYPT_E_UNEXPECTED_MSG_TYPE                            As Long = &H8009200A 'The certificate does not have a property that references a private key. 'CRYPT_E_UNEXPECTED_MSG_TYPE
Public Const ERROR_CRYPT_E_NO_KEY_PROPERTY                                As Long = &H8009200B 'Cannot find the certificate and private key for decryption. 'CRYPT_E_NO_KEY_PROPERTY
Public Const ERROR_CRYPT_E_NO_DECRYPT_CERT                                As Long = &H8009200C 'Cannot find the certificate and private key to use for decryption. 'CRYPT_E_NO_DECRYPT_CERT
Public Const ERROR_CRYPT_E_BAD_MSG                                        As Long = &H8009200D 'Not a cryptographic message or the cryptographic message is not formatted correctly. 'CRYPT_E_BAD_MSG
Public Const ERROR_CRYPT_E_NO_SIGNER                                      As Long = &H8009200E 'The signed cryptographic message does not have a signer for the specified signer index. 'CRYPT_E_NO_SIGNER
Public Const ERROR_CRYPT_E_PENDING_CLOSE                                  As Long = &H8009200F 'Final closure is pending until additional frees or closes. 'CRYPT_E_PENDING_CLOSE
Public Const ERROR_CRYPT_E_REVOKED                                        As Long = &H80092010 'The certificate is revoked. 'CRYPT_E_REVOKED
Public Const ERROR_CRYPT_E_NO_REVOCATION_DLL                              As Long = &H80092011 'No Dll or exported function was found to verify revocation. 'CRYPT_E_NO_REVOCATION_DLL
Public Const ERROR_CRYPT_E_NO_REVOCATION_CHECK                            As Long = &H80092012 'The revocation function was unable to check revocation for the certificate. 'CRYPT_E_NO_REVOCATION_CHECK
Public Const ERROR_CRYPT_E_REVOCATION_OFFLINE                             As Long = &H80092013 'The revocation function was unable to check revocation because the revocation server was offline. 'CRYPT_E_REVOCATION_OFFLINE
Public Const ERROR_CRYPT_E_NOT_IN_REVOCATION_DATABASE                     As Long = &H80092014 'The certificate is not in the revocation server's database. 'CRYPT_E_NOT_IN_REVOCATION_DATABASE
Public Const ERROR_CRYPT_E_INVALID_NUMERIC_STRING                         As Long = &H80092020 'The string contains a non-numeric character. 'CRYPT_E_INVALID_NUMERIC_STRING
Public Const ERROR_CRYPT_E_INVALID_PRINTABLE_STRING                       As Long = &H80092021 'The string contains a non-printable character. 'CRYPT_E_INVALID_PRINTABLE_STRING
Public Const ERROR_CRYPT_E_INVALID_IA5_STRING                             As Long = &H80092022 'The string contains a character not in the 7 bit ASCII character set. 'CRYPT_E_INVALID_IA5_STRING
Public Const ERROR_CRYPT_E_INVALID_X500_STRING                            As Long = &H80092023 'The string contains an invalid X500 name attribute key, oid, value or delimiter. 'CRYPT_E_INVALID_X500_STRING
Public Const ERROR_CRYPT_E_NOT_CHAR_STRING                                As Long = &H80092024 'The dwValueType for the CERT_NAME_VALUE is not one of the character strings.  Most likely it is either a CERT_RDN_ENCODED_BLOB or CERT_TDN_OCTED_STRING. 'CRYPT_E_NOT_CHAR_STRING
Public Const ERROR_CRYPT_E_FILERESIZED                                    As Long = &H80092025 'The Put operation can not continue.  The file needs to be resized.  However, there is already a signature present.  A complete signing operation must be done. 'CRYPT_E_FILERESIZED
Public Const ERROR_CRYPT_E_SECURITY_SETTINGS                              As Long = &H80092026 'The cryptographic operation failed due to a local security option setting. 'CRYPT_E_SECURITY_SETTINGS
Public Const ERROR_CRYPT_E_NO_VERIFY_USAGE_DLL                            As Long = &H80092027 'No DLL or exported function was found to verify subject usage. 'CRYPT_E_NO_VERIFY_USAGE_DLL
Public Const ERROR_CRYPT_E_NO_VERIFY_USAGE_CHECK                          As Long = &H80092028 'The called function was unable to do a usage check on the subject. 'CRYPT_E_NO_VERIFY_USAGE_CHECK
Public Const ERROR_CRYPT_E_VERIFY_USAGE_OFFLINE                           As Long = &H80092029 'Since the server was offline, the called function was unable to complete the usage check. 'CRYPT_E_VERIFY_USAGE_OFFLINE
Public Const ERROR_CRYPT_E_NOT_IN_CTL                                     As Long = &H8009202A 'The subject was not found in a Certificate Trust List (CTL). 'CRYPT_E_NOT_IN_CTL
Public Const ERROR_CRYPT_E_NO_TRUSTED_SIGNER                              As Long = &H8009202B 'None of the signers of the cryptographic message or certificate trust list is trusted. 'CRYPT_E_NO_TRUSTED_SIGNER
Public Const ERROR_CRYPT_E_MISSING_PUBKEY_PARA                            As Long = &H8009202C 'The public key's algorithm parameters are missing. 'CRYPT_E_MISSING_PUBKEY_PARA
Public Const ERROR_CRYPT_E_OSS_ERROR                                      As Long = &H80093000 'OSS Certificate encode/decode error code base 'CRYPT_E_OSS_ERROR
Public Const ERROR_OSS_MORE_BUF                                           As Long = &H80093001 'OSS ASN.1 Error: Output Buffer is too small. 'OSS_MORE_BUF
Public Const ERROR_OSS_NEGATIVE_UINTEGER                                  As Long = &H80093002 'OSS ASN.1 Error: Signed integer is encoded as a unsigned integer. 'OSS_NEGATIVE_UINTEGER
Public Const ERROR_OSS_PDU_RANGE                                          As Long = &H80093003 'OSS ASN.1 Error: Unknown ASN.1 data type. 'OSS_PDU_RANGE
Public Const ERROR_OSS_MORE_INPUT                                         As Long = &H80093004 'OSS ASN.1 Error: Output buffer is too small, the decoded data has been truncated. 'OSS_MORE_INPUT
Public Const ERROR_OSS_DATA_ERROR                                         As Long = &H80093005 'OSS ASN.1 Error: Invalid data. 'OSS_DATA_ERROR
Public Const ERROR_OSS_BAD_ARG                                            As Long = &H80093006 'OSS ASN.1 Error: Invalid argument. 'OSS_BAD_ARG
Public Const ERROR_OSS_BAD_VERSION                                        As Long = &H80093007 'OSS ASN.1 Error: Encode/Decode version mismatch. 'OSS_BAD_VERSION
Public Const ERROR_OSS_OUT_MEMORY                                         As Long = &H80093008 'OSS ASN.1 Error: Out of memory. 'OSS_OUT_MEMORY
Public Const ERROR_OSS_PDU_MISMATCH                                       As Long = &H80093009 'OSS ASN.1 Error: Encode/Decode Error. 'OSS_PDU_MISMATCH
Public Const ERROR_OSS_LIMITED                                            As Long = &H8009300A 'OSS ASN.1 Error: Internal Error. 'OSS_LIMITED
Public Const ERROR_OSS_BAD_PTR                                            As Long = &H8009300B 'OSS ASN.1 Error: Invalid data. 'OSS_BAD_PTR
Public Const ERROR_OSS_BAD_TIME                                           As Long = &H8009300C 'OSS ASN.1 Error: Invalid data. 'OSS_BAD_TIME
Public Const ERROR_OSS_INDEFINITE_NOT_SUPPORTED                           As Long = &H8009300D 'OSS ASN.1 Error: Unsupported BER indefinite-length encoding. 'OSS_INDEFINITE_NOT_SUPPORTED
Public Const ERROR_OSS_MEM_ERROR                                          As Long = &H8009300E 'OSS ASN.1 Error: Access violation. 'OSS_MEM_ERROR
Public Const ERROR_OSS_BAD_TABLE                                          As Long = &H8009300F 'OSS ASN.1 Error: Invalid data. 'OSS_BAD_TABLE
Public Const ERROR_OSS_TOO_LONG                                           As Long = &H80093010 'OSS ASN.1 Error: Invalid data. 'OSS_TOO_LONG
Public Const ERROR_OSS_CONSTRAINT_VIOLATED                                As Long = &H80093011 'OSS ASN.1 Error: Invalid data. 'OSS_CONSTRAINT_VIOLATED
Public Const ERROR_OSS_FATAL_ERROR                                        As Long = &H80093012 'OSS ASN.1 Error: Internal Error. 'OSS_FATAL_ERROR
Public Const ERROR_OSS_ACCESS_SERIALIZATION_ERROR                         As Long = &H80093013 'OSS ASN.1 Error: Multi-threading conflict. 'OSS_ACCESS_SERIALIZATION_ERROR
Public Const ERROR_OSS_NULL_TBL                                           As Long = &H80093014 'OSS ASN.1 Error: Invalid data. 'OSS_NULL_TBL
Public Const ERROR_OSS_NULL_FCN                                           As Long = &H80093015 'OSS ASN.1 Error: Invalid data. 'OSS_NULL_FCN
Public Const ERROR_OSS_BAD_ENCRULES                                       As Long = &H80093016 'OSS ASN.1 Error: Invalid data. 'OSS_BAD_ENCRULES
Public Const ERROR_OSS_UNAVAIL_ENCRULES                                   As Long = &H80093017 'OSS ASN.1 Error: Encode/Decode function not implemented. 'OSS_UNAVAIL_ENCRULES
Public Const ERROR_OSS_CANT_OPEN_TRACE_WINDOW                             As Long = &H80093018 'OSS ASN.1 Error: Trace file error. 'OSS_CANT_OPEN_TRACE_WINDOW
Public Const ERROR_OSS_UNIMPLEMENTED                                      As Long = &H80093019 'OSS ASN.1 Error: Function not implemented. 'OSS_UNIMPLEMENTED
Public Const ERROR_OSS_OID_DLL_NOT_LINKED                                 As Long = &H8009301A 'OSS ASN.1 Error: Program link error. 'OSS_OID_DLL_NOT_LINKED
Public Const ERROR_OSS_CANT_OPEN_TRACE_FILE                               As Long = &H8009301B 'OSS ASN.1 Error: Trace file error. 'OSS_CANT_OPEN_TRACE_FILE
Public Const ERROR_OSS_TRACE_FILE_ALREADY_OPEN                            As Long = &H8009301C 'OSS ASN.1 Error: Trace file error. 'OSS_TRACE_FILE_ALREADY_OPEN
Public Const ERROR_OSS_TABLE_MISMATCH                                     As Long = &H8009301D 'OSS ASN.1 Error: Invalid data. 'OSS_TABLE_MISMATCH
Public Const ERROR_OSS_TYPE_NOT_SUPPORTED                                 As Long = &H8009301E 'OSS ASN.1 Error: Invalid data. 'OSS_TYPE_NOT_SUPPORTED
Public Const ERROR_OSS_REAL_DLL_NOT_LINKED                                As Long = &H8009301F 'OSS ASN.1 Error: Program link error. 'OSS_REAL_DLL_NOT_LINKED
Public Const ERROR_OSS_REAL_CODE_NOT_LINKED                               As Long = &H80093020 'OSS ASN.1 Error: Program link error. 'OSS_REAL_CODE_NOT_LINKED
Public Const ERROR_OSS_OUT_OF_RANGE                                       As Long = &H80093021 'OSS ASN.1 Error: Program link error. 'OSS_OUT_OF_RANGE
Public Const ERROR_OSS_COPIER_DLL_NOT_LINKED                              As Long = &H80093022 'OSS ASN.1 Error: Program link error. 'OSS_COPIER_DLL_NOT_LINKED
Public Const ERROR_OSS_CONSTRAINT_DLL_NOT_LINKED                          As Long = &H80093023 'OSS ASN.1 Error: Program link error. 'OSS_CONSTRAINT_DLL_NOT_LINKED
Public Const ERROR_OSS_COMPARATOR_DLL_NOT_LINKED                          As Long = &H80093024 'OSS ASN.1 Error: Program link error. 'OSS_COMPARATOR_DLL_NOT_LINKED
Public Const ERROR_OSS_COMPARATOR_CODE_NOT_LINKED                         As Long = &H80093025 'OSS ASN.1 Error: Program link error. 'OSS_COMPARATOR_CODE_NOT_LINKED
Public Const ERROR_OSS_MEM_MGR_DLL_NOT_LINKED                             As Long = &H80093026 'OSS ASN.1 Error: Program link error. 'OSS_MEM_MGR_DLL_NOT_LINKED
Public Const ERROR_OSS_PDV_DLL_NOT_LINKED                                 As Long = &H80093027 'OSS ASN.1 Error: Program link error. 'OSS_PDV_DLL_NOT_LINKED
Public Const ERROR_OSS_PDV_CODE_NOT_LINKED                                As Long = &H80093028 'OSS ASN.1 Error: Program link error. 'OSS_PDV_CODE_NOT_LINKED
Public Const ERROR_OSS_API_DLL_NOT_LINKED                                 As Long = &H80093029 'OSS ASN.1 Error: Program link error. 'OSS_API_DLL_NOT_LINKED
Public Const ERROR_OSS_BERDER_DLL_NOT_LINKED                              As Long = &H8009302A 'OSS ASN.1 Error: Program link error. 'OSS_BERDER_DLL_NOT_LINKED
Public Const ERROR_OSS_PER_DLL_NOT_LINKED                                 As Long = &H8009302B 'OSS ASN.1 Error: Program link error. 'OSS_PER_DLL_NOT_LINKED
Public Const ERROR_OSS_OPEN_TYPE_ERROR                                    As Long = &H8009302C 'OSS ASN.1 Error: Program link error. 'OSS_OPEN_TYPE_ERROR
Public Const ERROR_OSS_MUTEX_NOT_CREATED                                  As Long = &H8009302D 'OSS ASN.1 Error: System resource error. 'OSS_MUTEX_NOT_CREATED
Public Const ERROR_OSS_CANT_CLOSE_TRACE_FILE                              As Long = &H8009302E 'OSS ASN.1 Error: Trace file error. 'OSS_CANT_CLOSE_TRACE_FILE
Public Const ERROR_CRYPT_E_ASN1_ERROR                                     As Long = &H80093100 'ASN1 Certificate encode/decode error code base. 'CRYPT_E_ASN1_ERROR
Public Const ERROR_CRYPT_E_ASN1_INTERNAL                                  As Long = &H80093101 'ASN1 internal encode or decode error. 'CRYPT_E_ASN1_INTERNAL
Public Const ERROR_CRYPT_E_ASN1_EOD                                       As Long = &H80093102 'ASN1 unexpected end of data. 'CRYPT_E_ASN1_EOD
Public Const ERROR_CRYPT_E_ASN1_CORRUPT                                   As Long = &H80093103 'ASN1 corrupted data. 'CRYPT_E_ASN1_CORRUPT
Public Const ERROR_CRYPT_E_ASN1_LARGE                                     As Long = &H80093104 'ASN1 value too large. 'CRYPT_E_ASN1_LARGE
Public Const ERROR_CRYPT_E_ASN1_CONSTRAINT                                As Long = &H80093105 'ASN1 constraint violated. 'CRYPT_E_ASN1_CONSTRAINT
Public Const ERROR_CRYPT_E_ASN1_MEMORY                                    As Long = &H80093106 'ASN1 out of memory. 'CRYPT_E_ASN1_MEMORY
Public Const ERROR_CRYPT_E_ASN1_OVERFLOW                                  As Long = &H80093107 'ASN1 buffer overflow. 'CRYPT_E_ASN1_OVERFLOW
Public Const ERROR_CRYPT_E_ASN1_BADPDU                                    As Long = &H80093108 'ASN1 function not supported for this PDU. 'CRYPT_E_ASN1_BADPDU
Public Const ERROR_CRYPT_E_ASN1_BADARGS                                   As Long = &H80093109 'ASN1 bad arguments to function call. 'CRYPT_E_ASN1_BADARGS
Public Const ERROR_CRYPT_E_ASN1_BADREAL                                   As Long = &H8009310A 'ASN1 bad real value. 'CRYPT_E_ASN1_BADREAL
Public Const ERROR_CRYPT_E_ASN1_BADTAG                                    As Long = &H8009310B 'ASN1 bad tag value met. 'CRYPT_E_ASN1_BADTAG
Public Const ERROR_CRYPT_E_ASN1_CHOICE                                    As Long = &H8009310C 'ASN1 bad choice value. 'CRYPT_E_ASN1_CHOICE
Public Const ERROR_CRYPT_E_ASN1_RULE                                      As Long = &H8009310D 'ASN1 bad encoding rule. 'CRYPT_E_ASN1_RULE
Public Const ERROR_CRYPT_E_ASN1_UTF8                                      As Long = &H8009310E 'ASN1 bad unicode (UTF8). 'CRYPT_E_ASN1_UTF8
Public Const ERROR_CRYPT_E_ASN1_PDU_TYPE                                  As Long = &H80093133 'ASN1 bad PDU type. 'CRYPT_E_ASN1_PDU_TYPE
Public Const ERROR_CRYPT_E_ASN1_NYI                                       As Long = &H80093134 'ASN1 not yet implemented. 'CRYPT_E_ASN1_NYI
Public Const ERROR_CRYPT_E_ASN1_EXTENDED                                  As Long = &H80093201 'ASN1 skipped unknown extension(s). 'CRYPT_E_ASN1_EXTENDED
Public Const ERROR_CRYPT_E_ASN1_NOEOD                                     As Long = &H80093202 'ASN1 end of data expected 'CRYPT_E_ASN1_NOEOD
Public Const ERROR_CERTSRV_E_BAD_REQUESTSUBJECT                           As Long = &H80094001 'The request subject name is invalid or too long. 'CERTSRV_E_BAD_REQUESTSUBJECT
Public Const ERROR_CERTSRV_E_NO_REQUEST                                   As Long = &H80094002 'The request does not exist. 'CERTSRV_E_NO_REQUEST
Public Const ERROR_CERTSRV_E_BAD_REQUESTSTATUS                            As Long = &H80094003 'The request's current status does not allow this operation. 'CERTSRV_E_BAD_REQUESTSTATUS
Public Const ERROR_CERTSRV_E_PROPERTY_EMPTY                               As Long = &H80094004 'The requested property value is empty. 'CERTSRV_E_PROPERTY_EMPTY
Public Const ERROR_CERTSRV_E_INVALID_CA_CERTIFICATE                       As Long = &H80094005 'The certification authority's certificate contains invalid data. 'CERTSRV_E_INVALID_CA_CERTIFICATE
Public Const ERROR_CERTSRV_E_SERVER_SUSPENDED                             As Long = &H80094006 'Certificate service has been suspended for a database restore operation. 'CERTSRV_E_SERVER_SUSPENDED
Public Const ERROR_CERTSRV_E_ENCODING_LENGTH                              As Long = &H80094007 'The certificate contains an encoded length that is potentially incompatible with older enrollment software. 'CERTSRV_E_ENCODING_LENGTH
Public Const ERROR_CERTSRV_E_ROLECONFLICT                                 As Long = &H80094008 'The operation is denied. The user has multiple roles assigned and the certification authority is configured to enforce role separation. 'CERTSRV_E_ROLECONFLICT
Public Const ERROR_CERTSRV_E_RESTRICTEDOFFICER                            As Long = &H80094009 'The operation is denied. It can only be performed by a certificate manager that is allowed to manage certificates for the current requester. 'CERTSRV_E_RESTRICTEDOFFICER
Public Const ERROR_CERTSRV_E_KEY_ARCHIVAL_NOT_CONFIGURED                  As Long = &H8009400A 'Cannot archive private key.  The certification authority is not configured for key archival. 'CERTSRV_E_KEY_ARCHIVAL_NOT_CONFIGURED
Public Const ERROR_CERTSRV_E_NO_VALID_KRA                                 As Long = &H8009400B 'Cannot archive private key.  The certification authority could not verify one or more key recovery certificates. 'CERTSRV_E_NO_VALID_KRA
Public Const ERROR_CERTSRV_E_BAD_REQUEST_KEY_ARCHIVAL                     As Long = &H8009400C 'The request is incorrectly formatted.  The encrypted private key must be in an unauthenticated attribute in an outermost signature. 'CERTSRV_E_BAD_REQUEST_KEY_ARCHIVAL
Public Const ERROR_CERTSRV_E_NO_CAADMIN_DEFINED                           As Long = &H8009400D 'At least one security principal must have the permission to manage this CA. 'CERTSRV_E_NO_CAADMIN_DEFINED
Public Const ERROR_CERTSRV_E_BAD_RENEWAL_CERT_ATTRIBUTE                   As Long = &H8009400E 'The request contains an invalid renewal certificate attribute. 'CERTSRV_E_BAD_RENEWAL_CERT_ATTRIBUTE
Public Const ERROR_CERTSRV_E_NO_DB_SESSIONS                               As Long = &H8009400F 'An attempt was made to open a Certification Authority database session, but there are already too many active sessions.  The server may need to be configured to allow additional sessions. 'CERTSRV_E_NO_DB_SESSIONS
Public Const ERROR_CERTSRV_E_ALIGNMENT_FAULT                              As Long = &H80094010 'A memory reference caused a data alignment fault. 'CERTSRV_E_ALIGNMENT_FAULT
Public Const ERROR_CERTSRV_E_ENROLL_DENIED                                As Long = &H80094011 'The permissions on this certification authority do not allow the current user to enroll for certificates. 'CERTSRV_E_ENROLL_DENIED
Public Const ERROR_CERTSRV_E_TEMPLATE_DENIED                              As Long = &H80094012 'The permissions on the certificate template do not allow the current user to enroll for this type of certificate. 'CERTSRV_E_TEMPLATE_DENIED
Public Const ERROR_CERTSRV_E_DOWNLEVEL_DC_SSL_OR_UPGRADE                  As Long = &H80094013 'The contacted domain controller cannot support signed LDAP traffic.  Update the domain controller or configure Certificate Services to use SSL for Active Directory access. 'CERTSRV_E_DOWNLEVEL_DC_SSL_OR_UPGRADE
Public Const ERROR_CERTSRV_E_UNSUPPORTED_CERT_TYPE                        As Long = &H80094800 'The requested certificate template is not supported by this CA. 'CERTSRV_E_UNSUPPORTED_CERT_TYPE
Public Const ERROR_CERTSRV_E_NO_CERT_TYPE                                 As Long = &H80094801 'The request contains no certificate template information. 'CERTSRV_E_NO_CERT_TYPE
Public Const ERROR_CERTSRV_E_TEMPLATE_CONFLICT                            As Long = &H80094802 'The request contains conflicting template information. 'CERTSRV_E_TEMPLATE_CONFLICT
Public Const ERROR_CERTSRV_E_SUBJECT_ALT_NAME_REQUIRED                    As Long = &H80094803 'The request is missing a required Subject Alternate name extension. 'CERTSRV_E_SUBJECT_ALT_NAME_REQUIRED
Public Const ERROR_CERTSRV_E_ARCHIVED_KEY_REQUIRED                        As Long = &H80094804 'The request is missing a required private key for archival by the server. 'CERTSRV_E_ARCHIVED_KEY_REQUIRED
Public Const ERROR_CERTSRV_E_SMIME_REQUIRED                               As Long = &H80094805 'The request is missing a required SMIME capabilities extension. 'CERTSRV_E_SMIME_REQUIRED
Public Const ERROR_CERTSRV_E_BAD_RENEWAL_SUBJECT                          As Long = &H80094806 'The request was made on behalf of a subject other than the caller.  The certificate template must be configured to require at least one signature to authorize the request. 'CERTSRV_E_BAD_RENEWAL_SUBJECT
Public Const ERROR_CERTSRV_E_BAD_TEMPLATE_VERSION                         As Long = &H80094807 'The request template version is newer than the supported template version. 'CERTSRV_E_BAD_TEMPLATE_VERSION
Public Const ERROR_CERTSRV_E_TEMPLATE_POLICY_REQUIRED                     As Long = &H80094808 'The template is missing a required signature policy attribute. 'CERTSRV_E_TEMPLATE_POLICY_REQUIRED
Public Const ERROR_CERTSRV_E_SIGNATURE_POLICY_REQUIRED                    As Long = &H80094809 'The request is missing required signature policy information. 'CERTSRV_E_SIGNATURE_POLICY_REQUIRED
Public Const ERROR_CERTSRV_E_SIGNATURE_COUNT                              As Long = &H8009480A 'The request is missing one or more required signatures. 'CERTSRV_E_SIGNATURE_COUNT
Public Const ERROR_CERTSRV_E_SIGNATURE_REJECTED                           As Long = &H8009480B 'One or more signatures did not include the required application or issuance policies.  The request is missing one or more required valid signatures. 'CERTSRV_E_SIGNATURE_REJECTED
Public Const ERROR_CERTSRV_E_ISSUANCE_POLICY_REQUIRED                     As Long = &H8009480C 'The request is missing one or more required signature issuance policies. 'CERTSRV_E_ISSUANCE_POLICY_REQUIRED
Public Const ERROR_CERTSRV_E_SUBJECT_UPN_REQUIRED                         As Long = &H8009480D 'The UPN is unavailable and cannot be added to the Subject Alternate name. 'CERTSRV_E_SUBJECT_UPN_REQUIRED
Public Const ERROR_CERTSRV_E_SUBJECT_DIRECTORY_GUID_REQUIRED              As Long = &H8009480E 'The Active Directory Guid is unavailable and cannot be added to the Subject Alternate name. 'CERTSRV_E_SUBJECT_DIRECTORY_GUID_REQUIRED
Public Const ERROR_CERTSRV_E_SUBJECT_DNS_REQUIRED                         As Long = &H8009480F 'The DNS name is unavailable and cannot be added to the Subject Alternate name. 'CERTSRV_E_SUBJECT_DNS_REQUIRED
Public Const ERROR_CERTSRV_E_ARCHIVED_KEY_UNEXPECTED                      As Long = &H80094810 'The request includes a private key for archival by the server, but key archival is not enabled for the specified certificate template. 'CERTSRV_E_ARCHIVED_KEY_UNEXPECTED
Public Const ERROR_CERTSRV_E_KEY_LENGTH                                   As Long = &H80094811 'The public key does not meet the minimum size required by the specified certificate template. 'CERTSRV_E_KEY_LENGTH
Public Const ERROR_CERTSRV_E_SUBJECT_EMAIL_REQUIRED                       As Long = &H80094812 'The EMail name is unavailable and cannot be added to the Subject or Subject Alternate name. 'CERTSRV_E_SUBJECT_EMAIL_REQUIRED
Public Const ERROR_CERTSRV_E_UNKNOWN_CERT_TYPE                            As Long = &H80094813 'One or more certificate templates to be enabled on this certification authority could not be found. 'CERTSRV_E_UNKNOWN_CERT_TYPE
Public Const ERROR_CERTSRV_E_CERT_TYPE_OVERLAP                            As Long = &H80094814 'The certificate template renewal period is longer than the certificate validity period.  The template should be reconfigured or the CA certificate renewed. 'CERTSRV_E_CERT_TYPE_OVERLAP
Public Const ERROR_XENROLL_E_KEY_NOT_EXPORTABLE                           As Long = &H80095000 'The key is not exportable. 'XENROLL_E_KEY_NOT_EXPORTABLE
Public Const ERROR_XENROLL_E_CANNOT_ADD_ROOT_CERT                         As Long = &H80095001 'You cannot add the root CA certificate into your local store. 'XENROLL_E_CANNOT_ADD_ROOT_CERT
Public Const ERROR_XENROLL_E_RESPONSE_KA_HASH_NOT_FOUND                   As Long = &H80095002 'The key archival hash attribute was not found in the response. 'XENROLL_E_RESPONSE_KA_HASH_NOT_FOUND
Public Const ERROR_XENROLL_E_RESPONSE_UNEXPECTED_KA_HASH                  As Long = &H80095003 'An unexpected key archival hash attribute was found in the response. 'XENROLL_E_RESPONSE_UNEXPECTED_KA_HASH
Public Const ERROR_XENROLL_E_RESPONSE_KA_HASH_MISMATCH                    As Long = &H80095004 'There is a key archival hash mismatch between the request and the response. 'XENROLL_E_RESPONSE_KA_HASH_MISMATCH
Public Const ERROR_XENROLL_E_KEYSPEC_SMIME_MISMATCH                       As Long = &H80095005 'Signing certificate cannot include SMIME extension. 'XENROLL_E_KEYSPEC_SMIME_MISMATCH
Public Const ERROR_TRUST_E_SYSTEM_ERROR                                   As Long = &H80096001 'A system-level error occurred while verifying trust. 'TRUST_E_SYSTEM_ERROR
Public Const ERROR_TRUST_E_NO_SIGNER_CERT                                 As Long = &H80096002 'The certificate for the signer of the message is invalid or not found. 'TRUST_E_NO_SIGNER_CERT
Public Const ERROR_TRUST_E_COUNTER_SIGNER                                 As Long = &H80096003 'One of the counter signatures was invalid. 'TRUST_E_COUNTER_SIGNER
Public Const ERROR_TRUST_E_CERT_SIGNATURE                                 As Long = &H80096004 'The signature of the certificate can not be verified. 'TRUST_E_CERT_SIGNATURE
Public Const ERROR_TRUST_E_TIME_STAMP                                     As Long = &H80096005 'The timestamp signature and/or certificate could not be verified or is malformed. 'TRUST_E_TIME_STAMP
Public Const ERROR_TRUST_E_BAD_DIGEST                                     As Long = &H80096010 'The digital signature of the object did not verify. 'TRUST_E_BAD_DIGEST
Public Const ERROR_TRUST_E_BASIC_CONSTRAINTS                              As Long = &H80096019 'A certificate's basic constraint extension has not been observed. 'TRUST_E_BASIC_CONSTRAINTS
Public Const ERROR_TRUST_E_FINANCIAL_CRITERIA                             As Long = &H8009601E 'The certificate does not meet or contain the Authenticode(tm) financial extensions. 'TRUST_E_FINANCIAL_CRITERIA
Public Const ERROR_MSSIPOTF_E_OUTOFMEMRANGE                               As Long = &H80097001 'Tried to reference a part of the file outside the proper range. 'MSSIPOTF_E_OUTOFMEMRANGE
Public Const ERROR_MSSIPOTF_E_CANTGETOBJECT                               As Long = &H80097002 'Could not retrieve an object from the file. 'MSSIPOTF_E_CANTGETOBJECT
Public Const ERROR_MSSIPOTF_E_NOHEADTABLE                                 As Long = &H80097003 'Could not find the head table in the file. 'MSSIPOTF_E_NOHEADTABLE
Public Const ERROR_MSSIPOTF_E_BAD_MAGICNUMBER                             As Long = &H80097004 'The magic number in the head table is incorrect. 'MSSIPOTF_E_BAD_MAGICNUMBER
Public Const ERROR_MSSIPOTF_E_BAD_OFFSET_TABLE                            As Long = &H80097005 'The offset table has incorrect values. 'MSSIPOTF_E_BAD_OFFSET_TABLE
Public Const ERROR_MSSIPOTF_E_TABLE_TAGORDER                              As Long = &H80097006 'Duplicate table tags or tags out of alphabetical order. 'MSSIPOTF_E_TABLE_TAGORDER
Public Const ERROR_MSSIPOTF_E_TABLE_LONGWORD                              As Long = &H80097007 'A table does not start on a long word boundary. 'MSSIPOTF_E_TABLE_LONGWORD
Public Const ERROR_MSSIPOTF_E_BAD_FIRST_TABLE_PLACEMENT                   As Long = &H80097008 'First table does not appear after header information. 'MSSIPOTF_E_BAD_FIRST_TABLE_PLACEMENT
Public Const ERROR_MSSIPOTF_E_TABLES_OVERLAP                              As Long = &H80097009 'Two or more tables overlap. 'MSSIPOTF_E_TABLES_OVERLAP
Public Const ERROR_MSSIPOTF_E_TABLE_PADBYTES                              As Long = &H8009700A 'Too many pad bytes between tables or pad bytes are not 0. 'MSSIPOTF_E_TABLE_PADBYTES
Public Const ERROR_MSSIPOTF_E_FILETOOSMALL                                As Long = &H8009700B 'File is too small to contain the last table. 'MSSIPOTF_E_FILETOOSMALL
Public Const ERROR_MSSIPOTF_E_TABLE_CHECKSUM                              As Long = &H8009700C 'A table checksum is incorrect. 'MSSIPOTF_E_TABLE_CHECKSUM
Public Const ERROR_MSSIPOTF_E_FILE_CHECKSUM                               As Long = &H8009700D 'The file checksum is incorrect. 'MSSIPOTF_E_FILE_CHECKSUM
Public Const ERROR_MSSIPOTF_E_FAILED_POLICY                               As Long = &H80097010 'The signature does not have the correct attributes for the policy. 'MSSIPOTF_E_FAILED_POLICY
Public Const ERROR_MSSIPOTF_E_FAILED_HINTS_CHECK                          As Long = &H80097011 'The file did not pass the hints check. 'MSSIPOTF_E_FAILED_HINTS_CHECK
Public Const ERROR_MSSIPOTF_E_NOT_OPENTYPE                                As Long = &H80097012 'The file is not an OpenType file. 'MSSIPOTF_E_NOT_OPENTYPE
Public Const ERROR_MSSIPOTF_E_FILE                                        As Long = &H80097013 'Failed on a file operation (open, map, read, write). 'MSSIPOTF_E_FILE
Public Const ERROR_MSSIPOTF_E_CRYPT                                       As Long = &H80097014 'A call to a CryptoAPI function failed. 'MSSIPOTF_E_CRYPT
Public Const ERROR_MSSIPOTF_E_BADVERSION                                  As Long = &H80097015 'There is a bad version number in the file. 'MSSIPOTF_E_BADVERSION
Public Const ERROR_MSSIPOTF_E_DSIG_STRUCTURE                              As Long = &H80097016 'The structure of the DSIG table is incorrect. 'MSSIPOTF_E_DSIG_STRUCTURE
Public Const ERROR_MSSIPOTF_E_PCONST_CHECK                                As Long = &H80097017 'A check failed in a partially constant table. 'MSSIPOTF_E_PCONST_CHECK
Public Const ERROR_MSSIPOTF_E_STRUCTURE                                   As Long = &H80097018 'Some kind of structural error. 'MSSIPOTF_E_STRUCTURE
Public Const ERROR_TRUST_E_PROVIDER_UNKNOWN                               As Long = &H800B0001 'Unknown trust provider. 'TRUST_E_PROVIDER_UNKNOWN
Public Const ERROR_TRUST_E_ACTION_UNKNOWN                                 As Long = &H800B0002 'The trust verification action specified is not supported by the specified trust provider. 'TRUST_E_ACTION_UNKNOWN
Public Const ERROR_TRUST_E_SUBJECT_FORM_UNKNOWN                           As Long = &H800B0003 'The form specified for the subject is not one supported or known by the specified trust provider. 'TRUST_E_SUBJECT_FORM_UNKNOWN
Public Const ERROR_TRUST_E_SUBJECT_NOT_TRUSTED                            As Long = &H800B0004 'The subject is not trusted for the specified action. 'TRUST_E_SUBJECT_NOT_TRUSTED
Public Const ERROR_DIGSIG_E_ENCODE                                        As Long = &H800B0005 'Error due to problem in ASN.1 encoding process. 'DIGSIG_E_ENCODE
Public Const ERROR_DIGSIG_E_DECODE                                        As Long = &H800B0006 'Error due to problem in ASN.1 decoding process. 'DIGSIG_E_DECODE
Public Const ERROR_DIGSIG_E_EXTENSIBILITY                                 As Long = &H800B0007 'Reading / writing Extensions where Attributes are appropriate, and visa versa. 'DIGSIG_E_EXTENSIBILITY
Public Const ERROR_DIGSIG_E_CRYPTO                                        As Long = &H800B0008 'Unspecified cryptographic failure. 'DIGSIG_E_CRYPTO
Public Const ERROR_PERSIST_E_SIZEDEFINITE                                 As Long = &H800B0009 'The size of the data could not be determined. 'PERSIST_E_SIZEDEFINITE
Public Const ERROR_PERSIST_E_SIZEINDEFINITE                               As Long = &H800B000A 'The size of the indefinite-sized data could not be determined. 'PERSIST_E_SIZEINDEFINITE
Public Const ERROR_PERSIST_E_NOTSELFSIZING                                As Long = &H800B000B 'This object does not read and write self-sizing data. 'PERSIST_E_NOTSELFSIZING
Public Const ERROR_TRUST_E_NOSIGNATURE                                    As Long = &H800B0100 'No signature was present in the subject. 'TRUST_E_NOSIGNATURE
Public Const ERROR_CERT_E_EXPIRED                                         As Long = &H800B0101 'A required certificate is not within its validity period when verifying against the current system clock or the timestamp in the signed file. 'CERT_E_EXPIRED
Public Const ERROR_CERT_E_VALIDITYPERIODNESTING                           As Long = &H800B0102 'The validity periods of the certification chain do not nest correctly. 'CERT_E_VALIDITYPERIODNESTING
Public Const ERROR_CERT_E_ROLE                                            As Long = &H800B0103 'A certificate that can only be used as an end-entity is being used as a CA or visa versa. 'CERT_E_ROLE
Public Const ERROR_CERT_E_PATHLENCONST                                    As Long = &H800B0104 'A path length constraint in the certification chain has been violated. 'CERT_E_PATHLENCONST
Public Const ERROR_CERT_E_CRITICAL                                        As Long = &H800B0105 'A certificate contains an unknown extension that is marked 'critical'. 'CERT_E_CRITICAL
Public Const ERROR_CERT_E_PURPOSE                                         As Long = &H800B0106 'A certificate being used for a purpose other than the ones specified by its CA. 'CERT_E_PURPOSE
Public Const ERROR_CERT_E_ISSUERCHAINING                                  As Long = &H800B0107 'A parent of a given certificate in fact did not issue that child certificate. 'CERT_E_ISSUERCHAINING
Public Const ERROR_CERT_E_MALFORMED                                       As Long = &H800B0108 'A certificate is missing or has an empty value for an important field, such as a subject or issuer name. 'CERT_E_MALFORMED
Public Const ERROR_CERT_E_UNTRUSTEDROOT                                   As Long = &H800B0109 'A certificate chain processed, but terminated in a root certificate which is not trusted by the trust provider. 'CERT_E_UNTRUSTEDROOT
Public Const ERROR_CERT_E_CHAINING                                        As Long = &H800B010A 'A certificate chain could not be built to a trusted root authority. 'CERT_E_CHAINING
Public Const ERROR_TRUST_E_FAIL                                           As Long = &H800B010B 'Generic trust failure. 'TRUST_E_FAIL
Public Const ERROR_CERT_E_REVOKED                                         As Long = &H800B010C 'A certificate was explicitly revoked by its issuer. 'CERT_E_REVOKED
Public Const ERROR_CERT_E_UNTRUSTEDTESTROOT                               As Long = &H800B010D 'The certification path terminates with the test root which is not trusted with the current policy settings. 'CERT_E_UNTRUSTEDTESTROOT
Public Const ERROR_CERT_E_REVOCATION_FAILURE                              As Long = &H800B010E 'The revocation process could not continue - the certificate(s) could not be checked. 'CERT_E_REVOCATION_FAILURE
Public Const ERROR_CERT_E_CN_NO_MATCH                                     As Long = &H800B010F 'The certificate's CN name does not match the passed value. 'CERT_E_CN_NO_MATCH
Public Const ERROR_CERT_E_WRONG_USAGE                                     As Long = &H800B0110 'The certificate is not valid for the requested usage. 'CERT_E_WRONG_USAGE
Public Const ERROR_TRUST_E_EXPLICIT_DISTRUST                              As Long = &H800B0111 'The certificate was explicitly marked as untrusted by the user. 'TRUST_E_EXPLICIT_DISTRUST
Public Const ERROR_CERT_E_UNTRUSTEDCA                                     As Long = &H800B0112 'A certification chain processed correctly, but one of the CA certificates is not trusted by the policy provider. 'CERT_E_UNTRUSTEDCA
Public Const ERROR_CERT_E_INVALID_POLICY                                  As Long = &H800B0113 'The certificate has invalid policy. 'CERT_E_INVALID_POLICY
Public Const ERROR_CERT_E_INVALID_NAME                                    As Long = &H800B0114 'The certificate has an invalid name. The name is not included in the permitted list or is explicitly excluded. 'CERT_E_INVALID_NAME
Public Const ERROR_SPAPI_E_EXPECTED_SECTION_NAME                          As Long = &H800F0000 'A non-empty line was encountered in the INF before the start of a section. 'SPAPI_E_EXPECTED_SECTION_NAME
Public Const ERROR_SPAPI_E_BAD_SECTION_NAME_LINE                          As Long = &H800F0001 'A section name marker in the INF is not complete, or does not exist on a line by itself. 'SPAPI_E_BAD_SECTION_NAME_LINE
Public Const ERROR_SPAPI_E_SECTION_NAME_TOO_LONG                          As Long = &H800F0002 'An INF section was encountered whose name exceeds the maximum section name length. 'SPAPI_E_SECTION_NAME_TOO_LONG
Public Const ERROR_SPAPI_E_GENERAL_SYNTAX                                 As Long = &H800F0003 'The syntax of the INF is invalid. 'SPAPI_E_GENERAL_SYNTAX
Public Const ERROR_SPAPI_E_WRONG_INF_STYLE                                As Long = &H800F0100 'The style of the INF is different than what was requested. 'SPAPI_E_WRONG_INF_STYLE
Public Const ERROR_SPAPI_E_SECTION_NOT_FOUND                              As Long = &H800F0101 'The required section was not found in the INF. 'SPAPI_E_SECTION_NOT_FOUND
Public Const ERROR_SPAPI_E_LINE_NOT_FOUND                                 As Long = &H800F0102 'The required line was not found in the INF. 'SPAPI_E_LINE_NOT_FOUND
Public Const ERROR_SPAPI_E_NO_BACKUP                                      As Long = &H800F0103 'The files affected by the installation of this file queue have not been backed up for uninstall. 'SPAPI_E_NO_BACKUP
Public Const ERROR_SPAPI_E_NO_ASSOCIATED_CLASS                            As Long = &H800F0200 'The INF or the device information set or element does not have an associated install class. 'SPAPI_E_NO_ASSOCIATED_CLASS
Public Const ERROR_SPAPI_E_CLASS_MISMATCH                                 As Long = &H800F0201 'The INF or the device information set or element does not match the specified install class. 'SPAPI_E_CLASS_MISMATCH
Public Const ERROR_SPAPI_E_DUPLICATE_FOUND                                As Long = &H800F0202 'An existing device was found that is a duplicate of the device being manually installed. 'SPAPI_E_DUPLICATE_FOUND
Public Const ERROR_SPAPI_E_NO_DRIVER_SELECTED                             As Long = &H800F0203 'There is no driver selected for the device information set or element. 'SPAPI_E_NO_DRIVER_SELECTED
Public Const ERROR_SPAPI_E_KEY_DOES_NOT_EXIST                             As Long = &H800F0204 'The requested device registry key does not exist. 'SPAPI_E_KEY_DOES_NOT_EXIST
Public Const ERROR_SPAPI_E_INVALID_DEVINST_NAME                           As Long = &H800F0205 'The device instance name is invalid. 'SPAPI_E_INVALID_DEVINST_NAME
Public Const ERROR_SPAPI_E_INVALID_CLASS                                  As Long = &H800F0206 'The install class is not present or is invalid. 'SPAPI_E_INVALID_CLASS
Public Const ERROR_SPAPI_E_DEVINST_ALREADY_EXISTS                         As Long = &H800F0207 'The device instance cannot be created because it already exists. 'SPAPI_E_DEVINST_ALREADY_EXISTS
Public Const ERROR_SPAPI_E_DEVINFO_NOT_REGISTERED                         As Long = &H800F0208 'The operation cannot be performed on a device information element that has not been registered. 'SPAPI_E_DEVINFO_NOT_REGISTERED
Public Const ERROR_SPAPI_E_INVALID_REG_PROPERTY                           As Long = &H800F0209 'The device property code is invalid. 'SPAPI_E_INVALID_REG_PROPERTY
Public Const ERROR_SPAPI_E_NO_INF                                         As Long = &H800F020A 'The INF from which a driver list is to be built does not exist. 'SPAPI_E_NO_INF
Public Const ERROR_SPAPI_E_NO_SUCH_DEVINST                                As Long = &H800F020B 'The device instance does not exist in the hardware tree. 'SPAPI_E_NO_SUCH_DEVINST
Public Const ERROR_SPAPI_E_CANT_LOAD_CLASS_ICON                           As Long = &H800F020C 'The icon representing this install class cannot be loaded. 'SPAPI_E_CANT_LOAD_CLASS_ICON
Public Const ERROR_SPAPI_E_INVALID_CLASS_INSTALLER                        As Long = &H800F020D 'The class installer registry entry is invalid. 'SPAPI_E_INVALID_CLASS_INSTALLER
Public Const ERROR_SPAPI_E_DI_DO_DEFAULT                                  As Long = &H800F020E 'The class installer has indicated that the default action should be performed for this installation request. 'SPAPI_E_DI_DO_DEFAULT
Public Const ERROR_SPAPI_E_DI_NOFILECOPY                                  As Long = &H800F020F 'The operation does not require any files to be copied. 'SPAPI_E_DI_NOFILECOPY
Public Const ERROR_SPAPI_E_INVALID_HWPROFILE                              As Long = &H800F0210 'The specified hardware profile does not exist. 'SPAPI_E_INVALID_HWPROFILE
Public Const ERROR_SPAPI_E_NO_DEVICE_SELECTED                             As Long = &H800F0211 'There is no device information element currently selected for this device information set. 'SPAPI_E_NO_DEVICE_SELECTED
Public Const ERROR_SPAPI_E_DEVINFO_LIST_LOCKED                            As Long = &H800F0212 'The operation cannot be performed because the device information set is locked. 'SPAPI_E_DEVINFO_LIST_LOCKED
Public Const ERROR_SPAPI_E_DEVINFO_DATA_LOCKED                            As Long = &H800F0213 'The operation cannot be performed because the device information element is locked. 'SPAPI_E_DEVINFO_DATA_LOCKED
Public Const ERROR_SPAPI_E_DI_BAD_PATH                                    As Long = &H800F0214 'The specified path does not contain any applicable device INFs. 'SPAPI_E_DI_BAD_PATH
Public Const ERROR_SPAPI_E_NO_CLASSINSTALL_PARAMS                         As Long = &H800F0215 'No class installer parameters have been set for the device information set or element. 'SPAPI_E_NO_CLASSINSTALL_PARAMS
Public Const ERROR_SPAPI_E_FILEQUEUE_LOCKED                               As Long = &H800F0216 'The operation cannot be performed because the file queue is locked. 'SPAPI_E_FILEQUEUE_LOCKED
Public Const ERROR_SPAPI_E_BAD_SERVICE_INSTALLSECT                        As Long = &H800F0217 'A service installation section in this INF is invalid. 'SPAPI_E_BAD_SERVICE_INSTALLSECT
Public Const ERROR_SPAPI_E_NO_CLASS_DRIVER_LIST                           As Long = &H800F0218 'There is no class driver list for the device information element. 'SPAPI_E_NO_CLASS_DRIVER_LIST
Public Const ERROR_SPAPI_E_NO_ASSOCIATED_SERVICE                          As Long = &H800F0219 'The installation failed because a function driver was not specified for this device instance. 'SPAPI_E_NO_ASSOCIATED_SERVICE
Public Const ERROR_SPAPI_E_NO_DEFAULT_DEVICE_INTERFACE                    As Long = &H800F021A 'There is presently no default device interface designated for this interface class. 'SPAPI_E_NO_DEFAULT_DEVICE_INTERFACE
Public Const ERROR_SPAPI_E_DEVICE_INTERFACE_ACTIVE                        As Long = &H800F021B 'The operation cannot be performed because the device interface is currently active. 'SPAPI_E_DEVICE_INTERFACE_ACTIVE
Public Const ERROR_SPAPI_E_DEVICE_INTERFACE_REMOVED                       As Long = &H800F021C 'The operation cannot be performed because the device interface has been removed from the system. 'SPAPI_E_DEVICE_INTERFACE_REMOVED
Public Const ERROR_SPAPI_E_BAD_INTERFACE_INSTALLSECT                      As Long = &H800F021D 'An interface installation section in this INF is invalid. 'SPAPI_E_BAD_INTERFACE_INSTALLSECT
Public Const ERROR_SPAPI_E_NO_SUCH_INTERFACE_CLASS                        As Long = &H800F021E 'This interface class does not exist in the system. 'SPAPI_E_NO_SUCH_INTERFACE_CLASS
Public Const ERROR_SPAPI_E_INVALID_REFERENCE_STRING                       As Long = &H800F021F 'The reference string supplied for this interface device is invalid. 'SPAPI_E_INVALID_REFERENCE_STRING
Public Const ERROR_SPAPI_E_INVALID_MACHINENAME                            As Long = &H800F0220 'The specified machine name does not conform to UNC naming conventions. 'SPAPI_E_INVALID_MACHINENAME
Public Const ERROR_SPAPI_E_REMOTE_COMM_FAILURE                            As Long = &H800F0221 'A general remote communication error occurred. 'SPAPI_E_REMOTE_COMM_FAILURE
Public Const ERROR_SPAPI_E_MACHINE_UNAVAILABLE                            As Long = &H800F0222 'The machine selected for remote communication is not available at this time. 'SPAPI_E_MACHINE_UNAVAILABLE
Public Const ERROR_SPAPI_E_NO_CONFIGMGR_SERVICES                          As Long = &H800F0223 'The Plug and Play service is not available on the remote machine. 'SPAPI_E_NO_CONFIGMGR_SERVICES
Public Const ERROR_SPAPI_E_INVALID_PROPPAGE_PROVIDER                      As Long = &H800F0224 'The property page provider registry entry is invalid. 'SPAPI_E_INVALID_PROPPAGE_PROVIDER
Public Const ERROR_SPAPI_E_NO_SUCH_DEVICE_INTERFACE                       As Long = &H800F0225 'The requested device interface is not present in the system. 'SPAPI_E_NO_SUCH_DEVICE_INTERFACE
Public Const ERROR_SPAPI_E_DI_POSTPROCESSING_REQUIRED                     As Long = &H800F0226 'The device's co-installer has additional work to perform after installation is complete. 'SPAPI_E_DI_POSTPROCESSING_REQUIRED
Public Const ERROR_SPAPI_E_INVALID_COINSTALLER                            As Long = &H800F0227 'The device's co-installer is invalid. 'SPAPI_E_INVALID_COINSTALLER
Public Const ERROR_SPAPI_E_NO_COMPAT_DRIVERS                              As Long = &H800F0228 'There are no compatible drivers for this device. 'SPAPI_E_NO_COMPAT_DRIVERS
Public Const ERROR_SPAPI_E_NO_DEVICE_ICON                                 As Long = &H800F0229 'There is no icon that represents this device or device type. 'SPAPI_E_NO_DEVICE_ICON
Public Const ERROR_SPAPI_E_INVALID_INF_LOGCONFIG                          As Long = &H800F022A 'A logical configuration specified in this INF is invalid. 'SPAPI_E_INVALID_INF_LOGCONFIG
Public Const ERROR_SPAPI_E_DI_DONT_INSTALL                                As Long = &H800F022B 'The class installer has denied the request to install or upgrade this device. 'SPAPI_E_DI_DONT_INSTALL
Public Const ERROR_SPAPI_E_INVALID_FILTER_DRIVER                          As Long = &H800F022C 'One of the filter drivers installed for this device is invalid. 'SPAPI_E_INVALID_FILTER_DRIVER
Public Const ERROR_SPAPI_E_NON_WINDOWS_NT_DRIVER                          As Long = &H800F022D 'The driver selected for this device does not support Windows XP. 'SPAPI_E_NON_WINDOWS_NT_DRIVER
Public Const ERROR_SPAPI_E_NON_WINDOWS_DRIVER                             As Long = &H800F022E 'The driver selected for this device does not support Windows. 'SPAPI_E_NON_WINDOWS_DRIVER
Public Const ERROR_SPAPI_E_NO_CATALOG_FOR_OEM_INF                         As Long = &H800F022F 'The third-party INF does not contain digital signature information. 'SPAPI_E_NO_CATALOG_FOR_OEM_INF
Public Const ERROR_SPAPI_E_DEVINSTALL_QUEUE_NONNATIVE                     As Long = &H800F0230 'An invalid attempt was made to use a device installation file queue for verification of digital signatures relative to other platforms. 'SPAPI_E_DEVINSTALL_QUEUE_NONNATIVE
Public Const ERROR_SPAPI_E_NOT_DISABLEABLE                                As Long = &H800F0231 'The device cannot be disabled. 'SPAPI_E_NOT_DISABLEABLE
Public Const ERROR_SPAPI_E_CANT_REMOVE_DEVINST                            As Long = &H800F0232 'The device could not be dynamically removed. 'SPAPI_E_CANT_REMOVE_DEVINST
Public Const ERROR_SPAPI_E_INVALID_TARGET                                 As Long = &H800F0233 'Cannot copy to specified target. 'SPAPI_E_INVALID_TARGET
Public Const ERROR_SPAPI_E_DRIVER_NONNATIVE                               As Long = &H800F0234 'Driver is not intended for this platform. 'SPAPI_E_DRIVER_NONNATIVE
Public Const ERROR_SPAPI_E_IN_WOW64                                       As Long = &H800F0235 'Operation not allowed in WOW64. 'SPAPI_E_IN_WOW64
Public Const ERROR_SPAPI_E_SET_SYSTEM_RESTORE_POINT                       As Long = &H800F0236 'The operation involving unsigned file copying was rolled back, so that a system restore point could be set. 'SPAPI_E_SET_SYSTEM_RESTORE_POINT
Public Const ERROR_SPAPI_E_INCORRECTLY_COPIED_INF                         As Long = &H800F0237 'An INF was copied into the Windows INF directory in an improper manner. 'SPAPI_E_INCORRECTLY_COPIED_INF
Public Const ERROR_SPAPI_E_SCE_DISABLED                                   As Long = &H800F0238 'The Security Configuration Editor (SCE) APIs have been disabled on this Embedded product. 'SPAPI_E_SCE_DISABLED
Public Const ERROR_SPAPI_E_UNKNOWN_EXCEPTION                              As Long = &H800F0239 'An unknown exception was encountered. 'SPAPI_E_UNKNOWN_EXCEPTION
Public Const ERROR_SPAPI_E_PNP_REGISTRY_ERROR                             As Long = &H800F023A 'A problem was encountered when accessing the Plug and Play registry database. 'SPAPI_E_PNP_REGISTRY_ERROR
Public Const ERROR_SPAPI_E_REMOTE_REQUEST_UNSUPPORTED                     As Long = &H800F023B 'The requested operation is not supported for a remote machine. 'SPAPI_E_REMOTE_REQUEST_UNSUPPORTED
Public Const ERROR_SPAPI_E_NOT_AN_INSTALLED_OEM_INF                       As Long = &H800F023C 'The specified file is not an installed OEM INF. 'SPAPI_E_NOT_AN_INSTALLED_OEM_INF
Public Const ERROR_SPAPI_E_INF_IN_USE_BY_DEVICES                          As Long = &H800F023D 'One or more devices are presently installed using the specified INF. 'SPAPI_E_INF_IN_USE_BY_DEVICES
Public Const ERROR_SPAPI_E_DI_FUNCTION_OBSOLETE                           As Long = &H800F023E 'The requested device install operation is obsolete. 'SPAPI_E_DI_FUNCTION_OBSOLETE
Public Const ERROR_SPAPI_E_NO_AUTHENTICODE_CATALOG                        As Long = &H800F023F 'A file could not be verified because it does not have an associated catalog signed via Authenticode(tm). 'SPAPI_E_NO_AUTHENTICODE_CATALOG
Public Const ERROR_SPAPI_E_AUTHENTICODE_DISALLOWED                        As Long = &H800F0240 'Authenticode(tm) signature verification is not supported for the specified INF. 'SPAPI_E_AUTHENTICODE_DISALLOWED
Public Const ERROR_SPAPI_E_AUTHENTICODE_TRUSTED_PUBLISHER                 As Long = &H800F0241 'The INF was signed with an Authenticode(tm) catalog from a trusted publisher. 'SPAPI_E_AUTHENTICODE_TRUSTED_PUBLISHER
Public Const ERROR_SPAPI_E_AUTHENTICODE_TRUST_NOT_ESTABLISHED             As Long = &H800F0242 'The publisher of an Authenticode(tm) signed catalog has not yet been established as trusted. 'SPAPI_E_AUTHENTICODE_TRUST_NOT_ESTABLISHED
Public Const ERROR_SPAPI_E_AUTHENTICODE_PUBLISHER_NOT_TRUSTED             As Long = &H800F0243 'The publisher of an Authenticode(tm) signed catalog was not established as trusted. 'SPAPI_E_AUTHENTICODE_PUBLISHER_NOT_TRUSTED
Public Const ERROR_SPAPI_E_SIGNATURE_OSATTRIBUTE_MISMATCH                 As Long = &H800F0244 'The software was tested for compliance with Windows Logo requirements on a different version of Windows, and may not be compatible with this version. 'SPAPI_E_SIGNATURE_OSATTRIBUTE_MISMATCH
Public Const ERROR_SPAPI_E_ONLY_VALIDATE_VIA_AUTHENTICODE                 As Long = &H800F0245 'The file may only be validated by a catalog signed via Authenticode(tm). 'SPAPI_E_ONLY_VALIDATE_VIA_AUTHENTICODE
Public Const ERROR_SPAPI_E_UNRECOVERABLE_STACK_OVERFLOW                   As Long = &H800F0300 'An unrecoverable stack overflow was encountered. 'SPAPI_E_UNRECOVERABLE_STACK_OVERFLOW
Public Const ERROR_SPAPI_E_ERROR_NOT_INSTALLED                            As Long = &H800F1000 'No installed components were detected. 'SPAPI_E_ERROR_NOT_INSTALLED
Public Const ERROR_SCARD_F_INTERNAL_ERROR                                 As Long = &H80100001 'An internal consistency check failed. 'SCARD_F_INTERNAL_ERROR
Public Const ERROR_SCARD_E_CANCELLED                                      As Long = &H80100002 'The action was cancelled by an SCardCancel request. 'SCARD_E_CANCELLED
Public Const ERROR_SCARD_E_INVALID_HANDLE                                 As Long = &H80100003 'The supplied handle was invalid. 'SCARD_E_INVALID_HANDLE
Public Const ERROR_SCARD_E_INVALID_PARAMETER                              As Long = &H80100004 'One or more of the supplied parameters could not be properly interpreted. 'SCARD_E_INVALID_PARAMETER
Public Const ERROR_SCARD_E_INVALID_TARGET                                 As Long = &H80100005 'Registry startup information is missing or invalid. 'SCARD_E_INVALID_TARGET
Public Const ERROR_SCARD_E_NO_MEMORY                                      As Long = &H80100006 'Not enough memory available to complete this command. 'SCARD_E_NO_MEMORY
Public Const ERROR_SCARD_F_WAITED_TOO_LONG                                As Long = &H80100007 'An internal consistency timer has expired. 'SCARD_F_WAITED_TOO_LONG
Public Const ERROR_SCARD_E_INSUFFICIENT_BUFFER                            As Long = &H80100008 'The data buffer to receive returned data is too small for the returned data. 'SCARD_E_INSUFFICIENT_BUFFER
Public Const ERROR_SCARD_E_UNKNOWN_READER                                 As Long = &H80100009 'The specified reader name is not recognized. 'SCARD_E_UNKNOWN_READER
Public Const ERROR_SCARD_E_TIMEOUT                                        As Long = &H8010000A 'The user-specified timeout value has expired. 'SCARD_E_TIMEOUT
Public Const ERROR_SCARD_E_SHARING_VIOLATION                              As Long = &H8010000B 'The smart card cannot be accessed because of other connections outstanding. 'SCARD_E_SHARING_VIOLATION
Public Const ERROR_SCARD_E_NO_SMARTCARD                                   As Long = &H8010000C 'The operation requires a Smart Card, but no Smart Card is currently in the device. 'SCARD_E_NO_SMARTCARD
Public Const ERROR_SCARD_E_UNKNOWN_CARD                                   As Long = &H8010000D 'The specified smart card name is not recognized. 'SCARD_E_UNKNOWN_CARD
Public Const ERROR_SCARD_E_CANT_DISPOSE                                   As Long = &H8010000E 'The system could not dispose of the media in the requested manner. 'SCARD_E_CANT_DISPOSE
Public Const ERROR_SCARD_E_PROTO_MISMATCH                                 As Long = &H8010000F 'The requested protocols are incompatible with the protocol currently in use with the smart card. 'SCARD_E_PROTO_MISMATCH
Public Const ERROR_SCARD_E_NOT_READY                                      As Long = &H80100010 'The reader or smart card is not ready to accept commands. 'SCARD_E_NOT_READY
Public Const ERROR_SCARD_E_INVALID_VALUE                                  As Long = &H80100011 'One or more of the supplied parameters values could not be properly interpreted. 'SCARD_E_INVALID_VALUE
Public Const ERROR_SCARD_E_SYSTEM_CANCELLED                               As Long = &H80100012 'The action was cancelled by the system, presumably to log off or shut down. 'SCARD_E_SYSTEM_CANCELLED
Public Const ERROR_SCARD_F_COMM_ERROR                                     As Long = &H80100013 'An internal communications error has been detected. 'SCARD_F_COMM_ERROR
Public Const ERROR_SCARD_F_UNKNOWN_ERROR                                  As Long = &H80100014 'An internal error has been detected, but the source is unknown. 'SCARD_F_UNKNOWN_ERROR
Public Const ERROR_SCARD_E_INVALID_ATR                                    As Long = &H80100015 'An ATR obtained from the registry is not a valid ATR string. 'SCARD_E_INVALID_ATR
Public Const ERROR_SCARD_E_NOT_TRANSACTED                                 As Long = &H80100016 'An attempt was made to end a non-existent transaction. 'SCARD_E_NOT_TRANSACTED
Public Const ERROR_SCARD_E_READER_UNAVAILABLE                             As Long = &H80100017 'The specified reader is not currently available for use. 'SCARD_E_READER_UNAVAILABLE
Public Const ERROR_SCARD_P_SHUTDOWN                                       As Long = &H80100018 'The operation has been aborted to allow the server application to exit. 'SCARD_P_SHUTDOWN
Public Const ERROR_SCARD_E_PCI_TOO_SMALL                                  As Long = &H80100019 'The PCI Receive buffer was too small. 'SCARD_E_PCI_TOO_SMALL
Public Const ERROR_SCARD_E_READER_UNSUPPORTED                             As Long = &H8010001A 'The reader driver does not meet minimal requirements for support. 'SCARD_E_READER_UNSUPPORTED
Public Const ERROR_SCARD_E_DUPLICATE_READER                               As Long = &H8010001B 'The reader driver did not produce a unique reader name. 'SCARD_E_DUPLICATE_READER
Public Const ERROR_SCARD_E_CARD_UNSUPPORTED                               As Long = &H8010001C 'The smart card does not meet minimal requirements for support. 'SCARD_E_CARD_UNSUPPORTED
Public Const ERROR_SCARD_E_NO_SERVICE                                     As Long = &H8010001D 'The Smart card resource manager is not running. 'SCARD_E_NO_SERVICE
Public Const ERROR_SCARD_E_SERVICE_STOPPED                                As Long = &H8010001E 'The Smart card resource manager has shut down. 'SCARD_E_SERVICE_STOPPED
Public Const ERROR_SCARD_E_UNEXPECTED                                     As Long = &H8010001F 'An unexpected card error has occurred. 'SCARD_E_UNEXPECTED
Public Const ERROR_SCARD_E_ICC_INSTALLATION                               As Long = &H80100020 'No Primary Provider can be found for the smart card. 'SCARD_E_ICC_INSTALLATION
Public Const ERROR_SCARD_E_ICC_CREATEORDER                                As Long = &H80100021 'The requested order of object creation is not supported. 'SCARD_E_ICC_CREATEORDER
Public Const ERROR_SCARD_E_UNSUPPORTED_FEATURE                            As Long = &H80100022 'This smart card does not support the requested feature. 'SCARD_E_UNSUPPORTED_FEATURE
Public Const ERROR_SCARD_E_DIR_NOT_FOUND                                  As Long = &H80100023 'The identified directory does not exist in the smart card. 'SCARD_E_DIR_NOT_FOUND
Public Const ERROR_SCARD_E_FILE_NOT_FOUND                                 As Long = &H80100024 'The identified file does not exist in the smart card. 'SCARD_E_FILE_NOT_FOUND
Public Const ERROR_SCARD_E_NO_DIR                                         As Long = &H80100025 'The supplied path does not represent a smart card directory. 'SCARD_E_NO_DIR
Public Const ERROR_SCARD_E_NO_FILE                                        As Long = &H80100026 'The supplied path does not represent a smart card file. 'SCARD_E_NO_FILE
Public Const ERROR_SCARD_E_NO_ACCESS                                      As Long = &H80100027 'Access is denied to this file. 'SCARD_E_NO_ACCESS
Public Const ERROR_SCARD_E_WRITE_TOO_MANY                                 As Long = &H80100028 'The smartcard does not have enough memory to store the information. 'SCARD_E_WRITE_TOO_MANY
Public Const ERROR_SCARD_E_BAD_SEEK                                       As Long = &H80100029 'There was an error trying to set the smart card file object pointer. 'SCARD_E_BAD_SEEK
Public Const ERROR_SCARD_E_INVALID_CHV                                    As Long = &H8010002A 'The supplied PIN is incorrect. 'SCARD_E_INVALID_CHV
Public Const ERROR_SCARD_E_UNKNOWN_RES_MNG                                As Long = &H8010002B 'An unrecognized error code was returned from a layered component. 'SCARD_E_UNKNOWN_RES_MNG
Public Const ERROR_SCARD_E_NO_SUCH_CERTIFICATE                            As Long = &H8010002C 'The requested certificate does not exist. 'SCARD_E_NO_SUCH_CERTIFICATE
Public Const ERROR_SCARD_E_CERTIFICATE_UNAVAILABLE                        As Long = &H8010002D 'The requested certificate could not be obtained. 'SCARD_E_CERTIFICATE_UNAVAILABLE
Public Const ERROR_SCARD_E_NO_READERS_AVAILABLE                           As Long = &H8010002E 'Cannot find a smart card reader. 'SCARD_E_NO_READERS_AVAILABLE
Public Const ERROR_SCARD_E_COMM_DATA_LOST                                 As Long = &H8010002F 'A communications error with the smart card has been detected.  Retry the operation. 'SCARD_E_COMM_DATA_LOST
Public Const ERROR_SCARD_E_NO_KEY_CONTAINER                               As Long = &H80100030 'The requested key container does not exist on the smart card. 'SCARD_E_NO_KEY_CONTAINER
Public Const ERROR_SCARD_E_SERVER_TOO_BUSY                                As Long = &H80100031 'The Smart card resource manager is too busy to complete this operation. 'SCARD_E_SERVER_TOO_BUSY
Public Const ERROR_SCARD_W_UNSUPPORTED_CARD                               As Long = &H80100065 'The reader cannot communicate with the smart card, due to ATR configuration conflicts. 'SCARD_W_UNSUPPORTED_CARD
Public Const ERROR_SCARD_W_UNRESPONSIVE_CARD                              As Long = &H80100066 'The smart card is not responding to a reset. 'SCARD_W_UNRESPONSIVE_CARD
Public Const ERROR_SCARD_W_UNPOWERED_CARD                                 As Long = &H80100067 'Power has been removed from the smart card, so that further communication is not possible. 'SCARD_W_UNPOWERED_CARD
Public Const ERROR_SCARD_W_RESET_CARD                                     As Long = &H80100068 'The smart card has been reset, so any shared state information is invalid. 'SCARD_W_RESET_CARD
Public Const ERROR_SCARD_W_REMOVED_CARD                                   As Long = &H80100069 'The smart card has been removed, so that further communication is not possible. 'SCARD_W_REMOVED_CARD
Public Const ERROR_SCARD_W_SECURITY_VIOLATION                             As Long = &H8010006A 'Access was denied because of a security violation. 'SCARD_W_SECURITY_VIOLATION
Public Const ERROR_SCARD_W_WRONG_CHV                                      As Long = &H8010006B 'The card cannot be accessed because the wrong PIN was presented. 'SCARD_W_WRONG_CHV
Public Const ERROR_SCARD_W_CHV_BLOCKED                                    As Long = &H8010006C 'The card cannot be accessed because the maximum number of PIN entry attempts has been reached. 'SCARD_W_CHV_BLOCKED
Public Const ERROR_SCARD_W_EOF                                            As Long = &H8010006D 'The end of the smart card file has been reached. 'SCARD_W_EOF
Public Const ERROR_SCARD_W_CANCELLED_BY_USER                              As Long = &H8010006E 'The action was cancelled by the user. 'SCARD_W_CANCELLED_BY_USER
Public Const ERROR_SCARD_W_CARD_NOT_AUTHENTICATED                         As Long = &H8010006F 'No PIN was presented to the smart card. 'SCARD_W_CARD_NOT_AUTHENTICATED
Public Const ERROR_COMADMIN_E_OBJECTERRORS                                As Long = &H80110401 'Errors occurred accessing one or more objects - the ErrorInfo collection may have more detail 'COMADMIN_E_OBJECTERRORS
Public Const ERROR_COMADMIN_E_OBJECTINVALID                               As Long = &H80110402 'One or more of the object's properties are missing or invalid 'COMADMIN_E_OBJECTINVALID
Public Const ERROR_COMADMIN_E_KEYMISSING                                  As Long = &H80110403 'The object was not found in the catalog 'COMADMIN_E_KEYMISSING
Public Const ERROR_COMADMIN_E_ALREADYINSTALLED                            As Long = &H80110404 'The object is already registered 'COMADMIN_E_ALREADYINSTALLED
Public Const ERROR_COMADMIN_E_APP_FILE_WRITEFAIL                          As Long = &H80110407 'Error occurred writing to the application file 'COMADMIN_E_APP_FILE_WRITEFAIL
Public Const ERROR_COMADMIN_E_APP_FILE_READFAIL                           As Long = &H80110408 'Error occurred reading the application file 'COMADMIN_E_APP_FILE_READFAIL
Public Const ERROR_COMADMIN_E_APP_FILE_VERSION                            As Long = &H80110409 'Invalid version number in application file 'COMADMIN_E_APP_FILE_VERSION
Public Const ERROR_COMADMIN_E_BADPATH                                     As Long = &H8011040A 'The file path is invalid 'COMADMIN_E_BADPATH
Public Const ERROR_COMADMIN_E_APPLICATIONEXISTS                           As Long = &H8011040B 'The application is already installed 'COMADMIN_E_APPLICATIONEXISTS
Public Const ERROR_COMADMIN_E_ROLEEXISTS                                  As Long = &H8011040C 'The role already exists 'COMADMIN_E_ROLEEXISTS
Public Const ERROR_COMADMIN_E_CANTCOPYFILE                                As Long = &H8011040D 'An error occurred copying the file 'COMADMIN_E_CANTCOPYFILE
Public Const ERROR_COMADMIN_E_NOUSER                                      As Long = &H8011040F 'One or more users are not valid 'COMADMIN_E_NOUSER
Public Const ERROR_COMADMIN_E_INVALIDUSERIDS                              As Long = &H80110410 'One or more users in the application file are not valid 'COMADMIN_E_INVALIDUSERIDS
Public Const ERROR_COMADMIN_E_NOREGISTRYCLSID                             As Long = &H80110411 'The component's CLSID is missing or corrupt 'COMADMIN_E_NOREGISTRYCLSID
Public Const ERROR_COMADMIN_E_BADREGISTRYPROGID                           As Long = &H80110412 'The component's progID is missing or corrupt 'COMADMIN_E_BADREGISTRYPROGID
Public Const ERROR_COMADMIN_E_AUTHENTICATIONLEVEL                         As Long = &H80110413 'Unable to set required authentication level for update request 'COMADMIN_E_AUTHENTICATIONLEVEL
Public Const ERROR_COMADMIN_E_USERPASSWDNOTVALID                          As Long = &H80110414 'The identity or password set on the application is not valid 'COMADMIN_E_USERPASSWDNOTVALID
Public Const ERROR_COMADMIN_E_CLSIDORIIDMISMATCH                          As Long = &H80110418 'Application file CLSIDs or IIDs do not match corresponding DLLs 'COMADMIN_E_CLSIDORIIDMISMATCH
Public Const ERROR_COMADMIN_E_REMOTEINTERFACE                             As Long = &H80110419 'Interface information is either missing or changed 'COMADMIN_E_REMOTEINTERFACE
Public Const ERROR_COMADMIN_E_DLLREGISTERSERVER                           As Long = &H8011041A 'DllRegisterServer failed on component install 'COMADMIN_E_DLLREGISTERSERVER
Public Const ERROR_COMADMIN_E_NOSERVERSHARE                               As Long = &H8011041B 'No server file share available 'COMADMIN_E_NOSERVERSHARE
Public Const ERROR_COMADMIN_E_DLLLOADFAILED                               As Long = &H8011041D 'DLL could not be loaded 'COMADMIN_E_DLLLOADFAILED
Public Const ERROR_COMADMIN_E_BADREGISTRYLIBID                            As Long = &H8011041E 'The registered TypeLib ID is not valid 'COMADMIN_E_BADREGISTRYLIBID
Public Const ERROR_COMADMIN_E_APPDIRNOTFOUND                              As Long = &H8011041F 'Application install directory not found 'COMADMIN_E_APPDIRNOTFOUND
Public Const ERROR_COMADMIN_E_REGISTRARFAILED                             As Long = &H80110423 'Errors occurred while in the component registrar 'COMADMIN_E_REGISTRARFAILED
Public Const ERROR_COMADMIN_E_COMPFILE_DOESNOTEXIST                       As Long = &H80110424 'The file does not exist 'COMADMIN_E_COMPFILE_DOESNOTEXIST
Public Const ERROR_COMADMIN_E_COMPFILE_LOADDLLFAIL                        As Long = &H80110425 'The DLL could not be loaded 'COMADMIN_E_COMPFILE_LOADDLLFAIL
Public Const ERROR_COMADMIN_E_COMPFILE_GETCLASSOBJ                        As Long = &H80110426 'GetClassObject failed in the DLL 'COMADMIN_E_COMPFILE_GETCLASSOBJ
Public Const ERROR_COMADMIN_E_COMPFILE_CLASSNOTAVAIL                      As Long = &H80110427 'The DLL does not support the components listed in the TypeLib 'COMADMIN_E_COMPFILE_CLASSNOTAVAIL
Public Const ERROR_COMADMIN_E_COMPFILE_BADTLB                             As Long = &H80110428 'The TypeLib could not be loaded 'COMADMIN_E_COMPFILE_BADTLB
Public Const ERROR_COMADMIN_E_COMPFILE_NOTINSTALLABLE                     As Long = &H80110429 'The file does not contain components or component information 'COMADMIN_E_COMPFILE_NOTINSTALLABLE
Public Const ERROR_COMADMIN_E_NOTCHANGEABLE                               As Long = &H8011042A 'Changes to this object and its sub-objects have been disabled 'COMADMIN_E_NOTCHANGEABLE
Public Const ERROR_COMADMIN_E_NOTDELETEABLE                               As Long = &H8011042B 'The delete function has been disabled for this object 'COMADMIN_E_NOTDELETEABLE
Public Const ERROR_COMADMIN_E_SESSION                                     As Long = &H8011042C 'The server catalog version is not supported 'COMADMIN_E_SESSION
Public Const ERROR_COMADMIN_E_COMP_MOVE_LOCKED                            As Long = &H8011042D 'The component move was disallowed, because the source or destination application is either a system application or currently locked against changes 'COMADMIN_E_COMP_MOVE_LOCKED
Public Const ERROR_COMADMIN_E_COMP_MOVE_BAD_DEST                          As Long = &H8011042E 'The component move failed because the destination application no longer exists 'COMADMIN_E_COMP_MOVE_BAD_DEST
Public Const ERROR_COMADMIN_E_REGISTERTLB                                 As Long = &H80110430 'The system was unable to register the TypeLib 'COMADMIN_E_REGISTERTLB
Public Const ERROR_COMADMIN_E_SYSTEMAPP                                   As Long = &H80110433 'This operation can not be performed on the system application 'COMADMIN_E_SYSTEMAPP
Public Const ERROR_COMADMIN_E_COMPFILE_NOREGISTRAR                        As Long = &H80110434 'The component registrar referenced in this file is not available 'COMADMIN_E_COMPFILE_NOREGISTRAR
Public Const ERROR_COMADMIN_E_COREQCOMPINSTALLED                          As Long = &H80110435 'A component in the same DLL is already installed 'COMADMIN_E_COREQCOMPINSTALLED
Public Const ERROR_COMADMIN_E_SERVICENOTINSTALLED                         As Long = &H80110436 'The service is not installed 'COMADMIN_E_SERVICENOTINSTALLED
Public Const ERROR_COMADMIN_E_PROPERTYSAVEFAILED                          As Long = &H80110437 'One or more property settings are either invalid or in conflict with each other 'COMADMIN_E_PROPERTYSAVEFAILED
Public Const ERROR_COMADMIN_E_OBJECTEXISTS                                As Long = &H80110438 'The object you are attempting to add or rename already exists 'COMADMIN_E_OBJECTEXISTS
Public Const ERROR_COMADMIN_E_COMPONENTEXISTS                             As Long = &H80110439 'The component already exists 'COMADMIN_E_COMPONENTEXISTS
Public Const ERROR_COMADMIN_E_REGFILE_CORRUPT                             As Long = &H8011043B 'The registration file is corrupt 'COMADMIN_E_REGFILE_CORRUPT
Public Const ERROR_COMADMIN_E_PROPERTY_OVERFLOW                           As Long = &H8011043C 'The property value is too large 'COMADMIN_E_PROPERTY_OVERFLOW
Public Const ERROR_COMADMIN_E_NOTINREGISTRY                               As Long = &H8011043E 'Object was not found in registry 'COMADMIN_E_NOTINREGISTRY
Public Const ERROR_COMADMIN_E_OBJECTNOTPOOLABLE                           As Long = &H8011043F 'This object is not poolable 'COMADMIN_E_OBJECTNOTPOOLABLE
Public Const ERROR_COMADMIN_E_APPLID_MATCHES_CLSID                        As Long = &H80110446 'A CLSID with the same Guid as the new application ID is already installed on this machine 'COMADMIN_E_APPLID_MATCHES_CLSID
Public Const ERROR_COMADMIN_E_ROLE_DOES_NOT_EXIST                         As Long = &H80110447 'A role assigned to a component, interface, or method did not exist in the application 'COMADMIN_E_ROLE_DOES_NOT_EXIST
Public Const ERROR_COMADMIN_E_START_APP_NEEDS_COMPONENTS                  As Long = &H80110448 'You must have components in an application in order to start the application 'COMADMIN_E_START_APP_NEEDS_COMPONENTS
Public Const ERROR_COMADMIN_E_REQUIRES_DIFFERENT_PLATFORM                 As Long = &H80110449 'This operation is not enabled on this platform 'COMADMIN_E_REQUIRES_DIFFERENT_PLATFORM
Public Const ERROR_COMADMIN_E_CAN_NOT_EXPORT_APP_PROXY                    As Long = &H8011044A 'Application Proxy is not exportable 'COMADMIN_E_CAN_NOT_EXPORT_APP_PROXY
Public Const ERROR_COMADMIN_E_CAN_NOT_START_APP                           As Long = &H8011044B 'Failed to start application because it is either a library application or an application proxy 'COMADMIN_E_CAN_NOT_START_APP
Public Const ERROR_COMADMIN_E_CAN_NOT_EXPORT_SYS_APP                      As Long = &H8011044C 'System application is not exportable 'COMADMIN_E_CAN_NOT_EXPORT_SYS_APP
Public Const ERROR_COMADMIN_E_CANT_SUBSCRIBE_TO_COMPONENT                 As Long = &H8011044D 'Can not subscribe to this component (the component may have been imported) 'COMADMIN_E_CANT_SUBSCRIBE_TO_COMPONENT
Public Const ERROR_COMADMIN_E_EVENTCLASS_CANT_BE_SUBSCRIBER               As Long = &H8011044E 'An event class cannot also be a subscriber component 'COMADMIN_E_EVENTCLASS_CANT_BE_SUBSCRIBER
Public Const ERROR_COMADMIN_E_LIB_APP_PROXY_INCOMPATIBLE                  As Long = &H8011044F 'Library applications and application proxies are incompatible 'COMADMIN_E_LIB_APP_PROXY_INCOMPATIBLE
Public Const ERROR_COMADMIN_E_BASE_PARTITION_ONLY                         As Long = &H80110450 'This function is valid for the base partition only 'COMADMIN_E_BASE_PARTITION_ONLY
Public Const ERROR_COMADMIN_E_START_APP_DISABLED                          As Long = &H80110451 'You cannot start an application that has been disabled 'COMADMIN_E_START_APP_DISABLED
Public Const ERROR_COMADMIN_E_CAT_DUPLICATE_PARTITION_NAME                As Long = &H80110457 'The specified partition name is already in use on this computer 'COMADMIN_E_CAT_DUPLICATE_PARTITION_NAME
Public Const ERROR_COMADMIN_E_CAT_INVALID_PARTITION_NAME                  As Long = &H80110458 'The specified partition name is invalid. Check that the name contains at least one visible character 'COMADMIN_E_CAT_INVALID_PARTITION_NAME
Public Const ERROR_COMADMIN_E_CAT_PARTITION_IN_USE                        As Long = &H80110459 'The partition cannot be deleted because it is the default partition for one or more users 'COMADMIN_E_CAT_PARTITION_IN_USE
Public Const ERROR_COMADMIN_E_FILE_PARTITION_DUPLICATE_FILES              As Long = &H8011045A 'The partition cannot be exported, because one or more components in the partition have the same file name 'COMADMIN_E_FILE_PARTITION_DUPLICATE_FILES
Public Const ERROR_COMADMIN_E_CAT_IMPORTED_COMPONENTS_NOT_ALLOWED         As Long = &H8011045B 'Applications that contain one or more imported components cannot be installed into a non-base partition 'COMADMIN_E_CAT_IMPORTED_COMPONENTS_NOT_ALLOWED
Public Const ERROR_COMADMIN_E_AMBIGUOUS_APPLICATION_NAME                  As Long = &H8011045C 'The application name is not unique and cannot be resolved to an application id 'COMADMIN_E_AMBIGUOUS_APPLICATION_NAME
Public Const ERROR_COMADMIN_E_AMBIGUOUS_PARTITION_NAME                    As Long = &H8011045D 'The partition name is not unique and cannot be resolved to a partition id 'COMADMIN_E_AMBIGUOUS_PARTITION_NAME
Public Const ERROR_COMADMIN_E_REGDB_NOTINITIALIZED                        As Long = &H80110472 'The COM+ registry database has not been initialized 'COMADMIN_E_REGDB_NOTINITIALIZED
Public Const ERROR_COMADMIN_E_REGDB_NOTOPEN                               As Long = &H80110473 'The COM+ registry database is not open 'COMADMIN_E_REGDB_NOTOPEN
Public Const ERROR_COMADMIN_E_REGDB_SYSTEMERR                             As Long = &H80110474 'The COM+ registry database detected a system error 'COMADMIN_E_REGDB_SYSTEMERR
Public Const ERROR_COMADMIN_E_REGDB_ALREADYRUNNING                        As Long = &H80110475 'The COM+ registry database is already running 'COMADMIN_E_REGDB_ALREADYRUNNING
Public Const ERROR_COMADMIN_E_MIG_VERSIONNOTSUPPORTED                     As Long = &H80110480 'This version of the COM+ registry database cannot be migrated 'COMADMIN_E_MIG_VERSIONNOTSUPPORTED
Public Const ERROR_COMADMIN_E_MIG_SCHEMANOTFOUND                          As Long = &H80110481 'The schema version to be migrated could not be found in the COM+ registry database 'COMADMIN_E_MIG_SCHEMANOTFOUND
Public Const ERROR_COMADMIN_E_CAT_BITNESSMISMATCH                         As Long = &H80110482 'There was a type mismatch between binaries 'COMADMIN_E_CAT_BITNESSMISMATCH
Public Const ERROR_COMADMIN_E_CAT_UNACCEPTABLEBITNESS                     As Long = &H80110483 'A binary of unknown or invalid type was provided 'COMADMIN_E_CAT_UNACCEPTABLEBITNESS
Public Const ERROR_COMADMIN_E_CAT_WRONGAPPBITNESS                         As Long = &H80110484 'There was a type mismatch between a binary and an application 'COMADMIN_E_CAT_WRONGAPPBITNESS
Public Const ERROR_COMADMIN_E_CAT_PAUSE_RESUME_NOT_SUPPORTED              As Long = &H80110485 'The application cannot be paused or resumed 'COMADMIN_E_CAT_PAUSE_RESUME_NOT_SUPPORTED
Public Const ERROR_COMADMIN_E_CAT_SERVERFAULT                             As Long = &H80110486 'The COM+ Catalog Server threw an exception during execution 'COMADMIN_E_CAT_SERVERFAULT
Public Const ERROR_COMQC_E_APPLICATION_NOT_QUEUED                         As Long = &H80110600 'Only COM+ Applications marked "queued" can be invoked using the "queue" moniker 'COMQC_E_APPLICATION_NOT_QUEUED
Public Const ERROR_COMQC_E_NO_QUEUEABLE_INTERFACES                        As Long = &H80110601 'At least one interface must be marked "queued" in order to create a queued component instance with the "queue" moniker 'COMQC_E_NO_QUEUEABLE_INTERFACES
Public Const ERROR_COMQC_E_QUEUING_SERVICE_NOT_AVAILABLE                  As Long = &H80110602 'MSMQ is required for the requested operation and is not installed 'COMQC_E_QUEUING_SERVICE_NOT_AVAILABLE
Public Const ERROR_COMQC_E_NO_IPERSISTSTREAM                              As Long = &H80110603 'Unable to marshal an interface that does not support IPersistStream 'COMQC_E_NO_IPERSISTSTREAM
Public Const ERROR_COMQC_E_BAD_MESSAGE                                    As Long = &H80110604 'The message is improperly formatted or was damaged in transit 'COMQC_E_BAD_MESSAGE
Public Const ERROR_COMQC_E_UNAUTHENTICATED                                As Long = &H80110605 'An unauthenticated message was received by an application that accepts only authenticated messages 'COMQC_E_UNAUTHENTICATED
Public Const ERROR_COMQC_E_UNTRUSTED_ENQUEUER                             As Long = &H80110606 'The message was requeued or moved by a user not in the "QC Trusted User" role 'COMQC_E_UNTRUSTED_ENQUEUER
Public Const ERROR_MSDTC_E_DUPLICATE_RESOURCE                             As Long = &H80110701 'Cannot create a duplicate resource of type Distributed Transaction Coordinator 'MSDTC_E_DUPLICATE_RESOURCE
Public Const ERROR_COMADMIN_E_OBJECT_PARENT_MISSING                       As Long = &H80110808 'One of the objects being inserted or updated does not belong to a valid parent collection 'COMADMIN_E_OBJECT_PARENT_MISSING
Public Const ERROR_COMADMIN_E_OBJECT_DOES_NOT_EXIST                       As Long = &H80110809 'One of the specified objects cannot be found 'COMADMIN_E_OBJECT_DOES_NOT_EXIST
Public Const ERROR_COMADMIN_E_APP_NOT_RUNNING                             As Long = &H8011080A 'The specified application is not currently running 'COMADMIN_E_APP_NOT_RUNNING
Public Const ERROR_COMADMIN_E_INVALID_PARTITION                           As Long = &H8011080B 'The partition(s) specified are not valid. 'COMADMIN_E_INVALID_PARTITION
Public Const ERROR_COMADMIN_E_SVCAPP_NOT_POOLABLE_OR_RECYCLABLE           As Long = &H8011080D 'COM+ applications that run as NT service may not be pooled or recycled 'COMADMIN_E_SVCAPP_NOT_POOLABLE_OR_RECYCLABLE
Public Const ERROR_COMADMIN_E_USER_IN_SET                                 As Long = &H8011080E 'One or more users are already assigned to a local partition set. 'COMADMIN_E_USER_IN_SET
Public Const ERROR_COMADMIN_E_CANTRECYCLELIBRARYAPPS                      As Long = &H8011080F 'Library applications may not be recycled. 'COMADMIN_E_CANTRECYCLELIBRARYAPPS
Public Const ERROR_COMADMIN_E_CANTRECYCLESERVICEAPPS                      As Long = &H80110811 'Applications running as NT services may not be recycled. 'COMADMIN_E_CANTRECYCLESERVICEAPPS
Public Const ERROR_COMADMIN_E_PROCESSALREADYRECYCLED                      As Long = &H80110812 'The process has already been recycled. 'COMADMIN_E_PROCESSALREADYRECYCLED
Public Const ERROR_COMADMIN_E_PAUSEDPROCESSMAYNOTBERECYCLED               As Long = &H80110813 'A paused process may not be recycled. 'COMADMIN_E_PAUSEDPROCESSMAYNOTBERECYCLED
Public Const ERROR_COMADMIN_E_CANTMAKEINPROCSERVICE                       As Long = &H80110814 'Library applications may not be NT services. 'COMADMIN_E_CANTMAKEINPROCSERVICE
Public Const ERROR_COMADMIN_E_PROGIDINUSEBYCLSID                          As Long = &H80110815 'The ProgID provided to the copy operation is invalid. The ProgID is in use by another registered CLSID. 'COMADMIN_E_PROGIDINUSEBYCLSID
Public Const ERROR_COMADMIN_E_DEFAULT_PARTITION_NOT_IN_SET                As Long = &H80110816 'The partition specified as default is not a member of the partition set. 'COMADMIN_E_DEFAULT_PARTITION_NOT_IN_SET
Public Const ERROR_COMADMIN_E_RECYCLEDPROCESSMAYNOTBEPAUSED               As Long = &H80110817 'A recycled process may not be paused. 'COMADMIN_E_RECYCLEDPROCESSMAYNOTBEPAUSED
Public Const ERROR_COMADMIN_E_PARTITION_ACCESSDENIED                      As Long = &H80110818 'Access to the specified partition is denied. 'COMADMIN_E_PARTITION_ACCESSDENIED
Public Const ERROR_COMADMIN_E_PARTITION_MSI_ONLY                          As Long = &H80110819 'Only Application Files (*.MSI files) can be installed into partitions. 'COMADMIN_E_PARTITION_MSI_ONLY
Public Const ERROR_COMADMIN_E_LEGACYCOMPS_NOT_ALLOWED_IN_1_0_FORMAT       As Long = &H8011081A 'Applications containing one or more legacy components may not be exported to 1.0 format. 'COMADMIN_E_LEGACYCOMPS_NOT_ALLOWED_IN_1_0_FORMAT
Public Const ERROR_COMADMIN_E_LEGACYCOMPS_NOT_ALLOWED_IN_NONBASE_PARTITIONS As Long = &H8011081B 'Legacy components may not exist in non-base partitions. 'COMADMIN_E_LEGACYCOMPS_NOT_ALLOWED_IN_NONBASE_PARTITIONS

