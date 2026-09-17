# LeyBc COM Library — 完整 API 接口报告

> 生成日期：2026-07-24  
> 总接口数：38  
> 总方法数：~641（IDispatch 分发表条目）  
> 总 CLSID 组件数：30

---

## 目录

1. [接口总览](#1-接口总览)
2. [组件 CLSID 列表](#2-组件-clsid-列表)
3. [ILeyBc 统一接口（95+ 方法）](#3-ileybc-统一接口95-方法)
4. [31 个子模块接口详情](#4-31-个子模块接口详情)
5. [IJScriptEngine JavaScript 引擎](#5-ijscriptengine-javascript-引擎)
6. [IFileMapping / IShareMemory](#6-ifilemapping-isharememo-ry)
7. [额外转发命名空间](#7-额外转发命名空间)

---

## 1. 接口总览

| # | 接口 IID | 名称 | 方法数 | 说明 |
|---|----------|------|:------:|------|
| 1 | A1B2C3D4-... | IComponentBase | 2 | 组件基础（名称/版本） |
| 2 | B2C3D4E5-... | ITextOperation | 56 | 文本操作（查找/替换/随机/拼音） |
| 3 | C3D4E5F6-... | IImageOperation | 10 | 图片操作（截图/缩放/裁切） |
| 4 | D4E5F6A7-... | IWindowOperation | 37 | 窗口操作（查找/枚举/状态/坐标/控制） |
| 5 | E5F6A7B8-... | IProcessOperation | 115 | 进程操作（创建/注入/内存/线程） |
| 6 | F6A7B8C9-... | IFileOperation | 13 | 文件操作（读写/复制/枚举） |
| 7 | A7B8C9D0-... | IDiskOperation | 7 | 磁盘操作（容量/文件系统） |
| 8 | B8C9D0E1-... | IThreadOperation | 7 | 线程操作（暂停/恢复/终止） |
| 9 | C9D0E1F2-... | IMultiThreadOperation | 12 | 多线程（临界区/事件/信号量） |
| 10 | D0E1F2A3-... | ITextEncoding | 21 | 编码转换（Ansi/UTF8/Base64/URL） |
| 11 | E1F2A3B4-... | ICryptoOperation | 40 | 加解密（AES/RSA/DES/SHA/HMAC） |
| 12 | B6C7D8E9-... | ICryptoOperationEx | — | 加解密扩展（OpenSSL 后端） |
| 13 | 7C6D5E4F-... | IComInvoke | 8 | COM 动态调用（CreateObject/Invoke） |
| 14 | F2A3B4C5-... | IBaseConversion | 7 | 进制转换（2/8/10/16 进制互转） |
| 15 | A3B4C5D6-... | ITimeOperation | 10 | 时间操作（获取/格式化/定时） |
| 16 | B4C5D6E7-... | IMemoryOperation | 90+ | 内存操作（读写/分配/指针/特征码） |
| 17 | C5D6E7F8-... | IAssemblyOperation | 15 | 汇编操作（X86/X64/Push/Call/Jmp） |
| 18 | D6E7F8A9-... | IDisassemblyOperation | 4 | 反汇编操作（反汇编/长度获取） |
| 19 | 3A1B2C3D-... | IHttpOperation | 5 | HTTP 请求（Get/Post/证书/Cookie） |
| 20 | 8F9A0B1C-... | IFileHandleOperation | 11 | 文件句柄操作（打开/读写/定位） |
| 21 | A0B1C2D3-... | IFSOOperation | 22 | FSO 兼容操作（Drive/Folder/File） |
| 22 | C2D3E4F5-... | IRegexOperation | 10 | 正则表达式（Match/Replace/Split） |
| 23 | 5C6D7E8F-... | IDLLOperation | 18 | DLL 操作（注入/卸载/PE/IME） |
| 24 | 7E8F9A0B-... | IHookOperation | 10 | Hook 操作（Inline/IAT/EAT/VEH） |
| 25 | 9A0B1C2D-... | IFasmOperation | 1 | FASM 汇编（单入口） |
| 26 | 1C2D3E4F-... | ISmCrypto | 8 | 国密加密（SM2/SM3/SM4） |
| 27 | 3E4F5061-... | IDirectoryOperation | 22 | 目录操作（Create/Copy/Enum） |
| 28 | 50617283-... | IFileOperationEx | 37 | 文件扩展操作（Find/Icon/MIME） |
| 29 | 728394A5-... | ISystemOperation | 1 | 系统操作（执行命令） |
| 30 | 01A2B3C4-... | IGlobalHotkeyOperation | 5 | 全局热键（Register/Unregister） |
| 31 | 23C4D5E6-... | IWindowHotkeyOperation | 3 | 窗口热键（Register/Unregister） |
| 32 | 45E6F7A8-... | IDialogOperation | 3 | 对话框（OpenFile/SaveFile） |
| 33 | 67A8B9C0-... | IDynamicClassNameOperation | 4 | 动态类名（Register/Unregister） |
| 34 | 89C0D1E2-... | IMenuOperation | 15 | 菜单操作（Add/Delete/Enum/Icon） |
| 35 | 0F1E2D3C-... | IFileMapping | 17 | 文件映射（Create/Map/Read/Write） |
| 36 | 2F3E4D5C-... | IShareMemory | 17 | 共享内存（同上，不同 CLSID） |
| 37 | 4B2E1A3C-... | ILeyBc | ~641 | 统一管理器（唯一入口，转发表） |
| 38 | 5A1B2C3D-... | IJScriptEngine | 12 | JavaScript 引擎（Execute/Eval/Call） |

---

## 2. 组件 CLSID 列表

| # | CLSID | 组件名 | ProgID | 说明 |
|---|-------|--------|--------|------|
| 1 | 1A2B3C4D-... | CLeyBcManager | `LeyBc.Manager` | **统一入口，推荐使用** |
| 2 | E7F8A9B0-... | CTextOperation | — | 文本操作组件 |
| 3 | F8A9B0C1-... | CImageOperation | — | 图片操作组件 |
| 4 | A9B0C1D2-... | CWindowOperation | — | 窗口操作组件 |
| 5 | B0C1D2E3-... | CProcessOperation | — | 进程操作组件 |
| 6 | C1D2E3F4-... | CFileOperation | — | 文件操作组件 |
| 7 | D2E3F4A5-... | CDiskOperation | — | 磁盘操作组件 |
| 8 | E3F4A5B6-... | CThreadOperation | — | 线程操作组件 |
| 9 | F4A5B6C7-... | CMultiThreadOperation | — | 多线程操作组件 |
| 10 | A5B6C7D8-... | CTextEncoding | — | 编码转换组件 |
| 11 | B6C7D8E9-... | CCryptoOperation | — | 加解密组件 |
| 12 | B6C7D8E9-... | COpenSSLCrypto | — | OpenSSL 3.x 后端加解密 |
| 13 | D9C8B7A6-... | CComInvoke | — | COM 动态调用组件 |
| 14 | C7D8E9F0-... | CBaseConversion | — | 进制转换组件 |
| 15 | D8E9F0A1-... | CTimeOperation | — | 时间操作组件 |
| 16 | E9F0A1B2-... | CMemoryOperation | — | 内存操作组件 |
| 17 | F0A1B2C3-... | CAssemblyOperation | — | 汇编操作组件 |
| 18 | A1B2C3D4-... | CDisassemblyOperation | — | 反汇编操作组件 |
| 19 | 4B5C6D7E-... | CHttpOperation | `LeyBc.HttpOperation` | HTTP 请求组件 |
| 20 | 9A0B1C2D-... | CFileHandleOperation | — | 文件句柄操作组件 |
| 21 | B1C2D3E4-... | CFSOOperation | — | FSO 兼容操作组件 |
| 22 | D3E4F506-... | CRegexOperation | — | 正则表达式组件 |
| 23 | 6D7E8F9A-... | CDLLOperation | `LeyBc.DLLOperation` | DLL 注入组件 |
| 24 | 8F9A0B1C-... | CHookOperation | `LeyBc.HookOperation` | Hook 操作组件 |
| 25 | 0B1C2D3E-... | CFasmOperation | — | FASM 汇编组件 |
| 26 | 2D3E4F50-... | CSmCrypto | — | 国密加密组件 |
| 27 | 4F506172-... | CDirectoryOperation | — | 目录操作组件 |
| 28 | 61728394-... | CFileOperationEx | — | 文件扩展操作组件 |
| 29 | 8394A5B6-... | CSystemOperation | — | 系统操作组件 |
| 30 | 12B3C4D5-... | CGlobalHotkeyOperation | — | 全局热键组件 |
| 31 | 34D5E6F7-... | CWindowHotkeyOperation | — | 窗口热键组件 |
| 32 | 56F7A8B9-... | CDialogOperation | — | 对话框组件 |
| 33 | 78B9C0D1-... | CDynamicClassNameOperation | — | 动态类名组件 |
| 34 | 90D1E2F3-... | CMenuOperation | — | 菜单操作组件 |
| 35 | 1F2E3D4C-... | CFileMapping | — | 映射文件组件 |
| 36 | 3F4E5D6C-... | CShareMemory | — | 共享内存组件 |
| 37 | 6B7C8D9E-... | CJScriptEngine | `LeyBc.JScriptEngine` | JS 引擎组件 |

---

## 3. ILeyBc 统一接口（95+ 方法）

ILeyBc 统一管理器将 34 个子模块的方法全部通过一个 `IDispatch` 接口暴露，转发表共 **641 个条目**。

### 3.1 IComponentBase（2 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| GetComponentName | → LPOLESTR* name | 获取组件名称 |
| GetComponentVersion | → DWORD* major, DWORD* minor | 获取组件版本 |

### 3.2 ITextOperation（37 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| FindText | text, find, flags → position | 查找子串位置 |
| ReplaceText | text, old, new, flags → result, count | 替换文本 |
| SplitText | text, delimiter → parts[], count | 分割字符串 |
| JoinText | parts[], count, separator → result | 合并字符串 |
| Substring | text, start, length → result | 取子串 |
| Trim | text → result | 去除首尾空白 |
| ToUpper | text → result | 转大写 |
| ToLower | text → result | 转小写 |
| GetLength | text → length | 获取长度 |
| EncodeBase64 | data, dataSize → result | Base64 编码 |
| DecodeBase64 | base64 → data, dataSize | Base64 解码 |
| Compare | text1, text2, ignoreCase → result | 字符串比较 |
| GetLeft | text, subText → result | 取左边文本 |
| GetRight | text, subText → result | 取右边文本 |
| GetMid | text, left, right → result | 取中间文本 |
| GetMidReverse | text, left, right → result | 从右取中间文本 |
| GetLineContaining | text, subText → result | 获取包含子串的行 |
| GetLineText | text, line → result | 获取指定行 |
| RemoveLine | text, line → result | 删除指定行 |
| RemoveEmptyLines | text → result | 删除空行 |
| RemoveNewLines | text → result | 删除换行符 |
| PadLeftZero | text, totalWidth → result | 左侧补零 |
| RemoveAllSpaces | text → result | 删除所有空白 |
| StartsWithDigit | text → bool | 是否数字开头 |
| IsNumeric | text → bool | 是否全数字 |
| SendText | hWnd, text | 向窗口发送文本 |
| SelectAll | hWnd | 全选编辑框文本 |
| CapitalizeFirst | text → result | 首字母大写 |
| SplitLines | text → lines[], count | 按行分割 |
| IsRepeated | text, subText → bool | 是否重复包含 |
| RandomDigit | length → result | 随机数字串 |
| RandomLetter | length, capital → result | 随机字母串 |
| RandomChar | length → result | 随机字符 |
| RandomChinese | length → result | 随机中文 |
| RandomSurname | → result | 随机姓氏 |
| GetPinyin | text, firstOnly → result | 获取拼音 |
| CountOccurrences | text, subText, ignoreCase → count | 统计出现次数 |

### 3.3 IImageOperation（10 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| CaptureScreen | → bitmap | 全屏截图 |
| CaptureRect | left, top, right, bottom → bitmap | 区域截图 |
| CaptureWindow | window → bitmap | 窗口截图 |
| SaveToFile | bitmap, filePath, format | 保存图片 |
| LoadFromFile | filePath → bitmap | 加载图片 |
| Resize | src, newWidth, newHeight → dst | 缩放图片 |
| Rotate | src, angle → dst | 旋转图片 |
| Crop | src, left, top, right, bottom → dst | 裁切图片 |
| GetWidth | bitmap → width | 获取图片宽度 |
| GetHeight | bitmap → height | 获取图片高度 |

### 3.4 IWindowOperation（37 方法）

窗口句柄统一使用 `ULONGLONG` / `VT_UI8`；枚举结果为 `SAFEARRAY(VT_UI8)`，文本为 `BSTR`，布尔值为 `VARIANT_BOOL`。自动化分发 DISPID 范围为 67-103。

| 方法 | 参数 | 说明 |
|------|------|------|
| FindWindowByTitle | title → window | 按标题查找窗口 |
| FindWindowByClass | className → window | 按类名查找窗口 |
| FindWindowEx | parentWindow, childAfter, className, title → window | 查找匹配的顶层或子窗口 |
| EnumWindows | → windows | 枚举所有顶层窗口 |
| EnumChildWindows | parentWindow → windows | 枚举指定窗口的所有子窗口 |
| GetForegroundWindow | → window | 获取前台窗口 |
| GetThreadFocusWindow | → window | 获取当前线程焦点窗口 |
| GetDesktopWindow | → window | 获取桌面窗口 |
| GetShellWindow | → window | 获取 Shell 窗口 |
| GetTaskbarWindow | → window | 获取任务栏窗口 |
| GetParentWindow | window → parentWindow | 获取父窗口 |
| GetOwnerWindow | window → ownerWindow | 获取所有者窗口 |
| GetWindowText | window → text | 获取窗口标题 |
| GetWindowClassName | window → className | 获取窗口类名 |
| SetWindowText | window, text | 设置窗口标题 |
| IsWindowValid | window → result | 判断窗口句柄是否有效 |
| IsWindowVisible | window → result | 判断窗口是否可见 |
| IsWindowEnabled | window → result | 判断窗口是否启用 |
| IsWindowMinimized | window → result | 判断窗口是否最小化 |
| IsWindowMaximized | window → result | 判断窗口是否最大化 |
| GetWindowRect | window → left, top, right, bottom | 获取窗口矩形 |
| GetClientRect | window → left, top, right, bottom | 获取客户区矩形 |
| ClientToScreen | window, x, y → x, y | 将客户区坐标转换为屏幕坐标 |
| ScreenToClient | window, x, y → x, y | 将屏幕坐标转换为客户区坐标 |
| SetWindowRect | window, left, top, width, height | 设置窗口位置和大小 |
| ShowWindow | window, cmdShow | 设置窗口显示状态 |
| EnableWindow | window, enabled | 启用或禁用窗口 |
| SetTopMost | window, topMost | 设置窗口置顶状态 |
| SetForegroundWindow | window | 请求将窗口切换到前台 |
| GetWindowProcessId | window → processId | 获取窗口所属进程 ID |
| GetWindowThreadId | window → threadId | 获取窗口所属线程 ID |
| CloseWindow | window | 请求关闭窗口 |
| RefreshWindow | window | 刷新窗口 |
| SetCaptureExclusion | window, excluded | 设置窗口捕获排除状态 |
| SetCloseEnabled | window, enabled | 启用或禁用系统菜单关闭项 |
| RegisterWindowMessage | messageName → messageId | 注册全局窗口消息 |
| FlashWindow | window, flags, count, timeout | 闪烁窗口 |

### 3.5 IProcessOperation（115 方法，核心模块）

**进程管理：**
CreateProcess, CreateProcessAsSystem, CreateProcessWithParent, CreateProcessWithParentByHandle, TerminateProcess, TerminateProcessByHandle, TerminateProcessById, SuspendProcess, SuspendProcessByHandle, ResumeProcess, ResumeProcessByHandle, OpenProcess, OpenProcessByPid, CloseProcess, CloseProcessHandle, CloseListHandle, CloseListHandle2, GetPidByName, GetPidsByName, GetPidByHandle, GetCurrentProcessId, GetCurrentProcessPseudoHandle, IsPidValid, IsAdmin, IsDebugged, IsDebuggedEx, IsX64Process, IsX64ProcessByHandle, GetProcessName, GetProcessNameByHandle, GetProcessPath, GetProcessPathByHandle, GetProcessUserName, GetProcessUserNameByHandle, GetProcessWindowHandle, GetProcessWindowHandleByEnum, GetParentProcessId, GetParentProcessIdByHandle, GetParentProcessName, GetParentProcessNameByHandle, GetChildProcessId, GetChildProcessIdByHandle, GetChildProcesses, GetChildProcessesByHandle, GetCommandLine, GetCommandLineByHandle, GetProcessMemoryUsage, GetProcessMemoryUsageByHandle, GetPriorityClass, GetPriorityClassByHandle, SetPriorityClass, SetPriorityClassByHandle

**模块/内存/注入：**
EnumProcessModules, GetProcessModules, GetProcessModulesByHandle, GetModuleBaseAddress, GetModuleBaseAddressByHandle, GetModuleEntryPoint, GetModuleEntryPointByHandle, GetModulePeb, GetModulePebByHandle, GetModuleInfo, GetModuleInfoByHandle, GetExportFunctions, GetExportFunctionsByHandle, GetExportFunctionsEx, GetExportFunctionsExByHandle, GetFunctionAddress, GetFunctionAddressByHandle, GetFunctionAddressFast, GetFunctionAddressFastByHandle, GetPebBase, GetPebBaseByHandle, InjectDLL, PreventDllInjection, PreventDllInjectionByHandle, StartProtectProcess, StopProtectProcess

**远程代码执行：**
CreateRemoteThread, CreateRemoteThreadByHandle, ExecuteAssemblyCode, ExecuteAssemblyCodeByHandle, ExecuteAssemblyCodeByThread, ExecuteAssemblyCodeByThread32, ExecuteAssemblyCodeByThreadByHandle, ExecuteAssemblyCodeByTimer, ExecuteAssemblyCodeByTimerByHandle, ExecuteAssemblyCodeEx, ExecuteAssemblyCodeExByHandle, InvokeProcessFunction, InvokeProcessFunctionByHandle

**线程/ShellCode：**
CreateHijackThread, CreateHijackThreadByHandle, GetMainThreadId, GetMainThreadIdByHandle, GetIoCounters, DuplicateObjectHandle, EnumProcessHandles, EnumProcessHandlesByHandle, GetTokenInformation, OpenProcessToken, AdjustPrivilege, SetProcessVolume, SetProcessDEP, GetDEPPolicy, GetPortByPid, IsPortOccupied, OpenPortByPid

**监控：**
BeginMonitorProcess, StopMonitorProcess, GetPidByPort

### 3.6 IFileOperation（13 方法）

| 方法 | 说明 |
|------|------|
| CreateFile, DeleteFile, CopyFile, MoveFile | 文件创建/删除/复制/移动 |
| ReadTextFile, WriteTextFile | 文本文件读写 |
| ReadBinaryFile, WriteBinaryFile | 二进制文件读写 |
| FileExists | 文件是否存在 |
| GetFileSize | 获取文件大小 |
| CreateDirectory | 创建目录 |
| EnumFiles | 枚举文件 |

### 3.7 IDiskOperation（7 方法）

EnumDrives, GetDiskTotalSize, GetDiskFreeSize, GetFileSystemName, GetVolumeLabel, GetVolumeSerial, GetDriveType

### 3.8 IThreadOperation（7 方法）

CreateThread, SuspendThread, ResumeThread, WaitForThread, TerminateThread, GetExitCode, GetCurrentThreadId

### 3.9 IMultiThreadOperation（12 方法）

CreateCriticalSection, EnterCriticalSection, LeaveCriticalSection, DeleteCriticalSection, CreateEvent, SetEvent, ResetEvent, WaitForSingleObject, WaitForMultipleObjects, CreateMutex, ReleaseMutex, CreateSemaphore, ReleaseSemaphore

### 3.10 ITextEncoding（21 方法）

| 方法 | 说明 |
|------|------|
| AnsiToUtf8, Utf8ToAnsi | ANSI ↔ UTF-8 |
| AnsiToUtf16, Utf16ToAnsi | ANSI ↔ UTF-16 |
| Utf8ToUtf16, Utf16ToUtf8 | UTF-8 ↔ UTF-16 |
| Utf8ToUnicode, UnicodeToUtf8 | UTF-8 ↔ Unicode |
| Gb2312ToUtf8, Utf8ToGb2312 | GB2312 ↔ UTF-8 |
| UnicodeToAnsi, AnsiToUnicode | Unicode ↔ ANSI |
| UnicodeToAnsiText, AnsiToUnicodeText | Unicode ↔ ANSI (文本) |
| UrlEncode, UrlDecode | URL 编码/解码 |
| Base64Encode, Base64Decode | Base64 编码/解码 |
| Base64EncodeA, Base64DecodeA | Base64 ANSI 编码/解码 |
| Base64DecodeImage | Base64 解码为图片文件 |
| TextToUcs2Cn, TextToUcs2, Ucs2ToText | UCS2 编码/解码 |

### 3.11 ICryptoOperation（40 方法）

**哈希：** ComputeMD5, ComputeMD5String, ComputeMD2, ComputeMD4, ComputeSHA1, ComputeSHA224, ComputeSHA256, ComputeSHA384, ComputeSHA512, ComputeSHA3, ComputeCRC32

**HMAC：** HMACMD5, HMACSHA1, HMACSHA224, HMACSHA256, HMACSHA384, HMACSHA512

**对称加密：** AESEncrypt, AESDecrypt, DESEncrypt, DESDecrypt, ThreeDESEncrypt, ThreeDESDecrypt, TEAEncrypt, TEADecrypt, XTEAEncrypt, XTEADecrypt, RC4Encrypt, XOREncrypt

**非对称加密：** RSAKeyGenerate, RSAPublicEncrypt, RSAPrivateDecrypt, RSASign, RSAVerify

**椭圆曲线：** ECDSASign, ECDSAVerify, ECDHKeyExchange

**DSA/DH：** DSASign, DSAVerify, DHKeyExchange

**编码：** Base32Encode, Base32Decode, Base58Encode, Base58Decode

### 3.12 IComInvoke（8 方法）

CreateObject, CreateObjectFromFile, GetProperty, PutProperty, Invoke, InvokeNoResult, ReleaseObject, GetInterfaceInfo

### 3.13 IBaseConversion（7 方法）

DecToBin, DecToOct, DecToHex, StrToDec, DecToBase, BigStrToDec, — (6 total listed in .h)

### 3.14 ITimeOperation（10 方法）

GetSystemTime, GetLocalTime, FormatTime, GetTimestamp, GetHighResTimestamp, TimestampToSystemTime, SystemTimeToTimestamp, SleepMs, StartTimer, StopTimer

### 3.15 IMemoryOperation（90+ 方法）

**基本内存读写（PID 模式）：** ReadByte, ReadShort, ReadInt, ReadLong, ReadIntPtr, ReadFloat, ReadDouble, ReadTextW, ReadTextA, ReadBytes, WriteByte, WriteShort, WriteInt, WriteLong, WriteIntPtr, WriteFloat, WriteDouble, WriteTextW, WriteTextA, WriteBytes, WritePtr

**基本内存读写（句柄模式）：** ReadByteEx, ReadShortEx, ReadIntEx, ReadLongEx, ReadIntPtrEx, ReadFloatEx, ReadDoubleEx, ReadTextWEx, ReadTextAEx, ReadBytesEx, WriteByteEx, WriteShortEx, WriteIntEx, WriteLongEx, WriteIntPtrEx, WriteFloatEx, WriteDoubleEx, WriteTextWEx, WriteTextAEx, WriteBytesEx, WritePtrEx

**无符号读写（PID 模式）：** ReadUByte, ReadUShort, ReadUInt, ReadULong, WriteUByte, WriteUShort, WriteUInt, WriteULong

**无符号读写（句柄模式）：** ReadUByteEx, ReadUShortEx, ReadUIntEx, ReadULongEx, WriteUByteEx, WriteUShortEx, WriteUIntEx, WriteULongEx

**远程内存分配：** AllocProcessMemory, AllocProcessMemory2, VirtualFree

**本地内存分配：** HeapAlloc, HeapAllocGlobal, HeapAllocLocal, HeapReAlloc, HeapFree, HeapAllocBytes, HeapAllocInt, HeapAllocIntPtr, HeapAllocTextW, HeapAllocTextA

**内存查询：** QueryMemory, QueryMemoryEx, QueryMemorySelf, ProtectMemory, ProtectMemoryEx, ProtectMemorySelf

**模块基址：** ReadModuleBase, ReadModuleBaseEx

**多级指针表达式：** ResolveExpressionAddress, ReadExpressionInt, ReadExpressionIntEx, ReadExpressionFloat, ReadExpressionFloatEx, ReadExpressionDouble, ReadExpressionDoubleEx, ReadExpressionLong, ReadExpressionLongEx, ReadExpressionIntPtr, ReadExpressionIntPtrEx

**特征码搜索：** FindPattern, FindPatternEx, FindPatternArray, FindPatternArrayEx, FindPatternInModule, FindPatternInModuleEx, SearchBytes, SearchBytesEx

**进程内存：** ReadProcessMemory, WriteProcessMemory, VirtualAllocEx, VirtualFreeEx, VirtualQueryEx

### 3.16 IAssemblyOperation（15 方法）

AssembleX86, AssembleX64, ExecuteAsmX86, Push, Call, CallDwordPtr, CallQwordPtr, Jmp, JmpLong, Pushad_X64, Popad_X64, AsmEmpty, AsmAdd, AsmTakeShellcode, AsmGetText, AsmPushadX64, AsmPopadX64

### 3.17 IDisassemblyOperation（4 方法）

DisassembleX86, DisassembleX64, DisassembleBatch, GetInstructionLength

### 3.18 IHttpOperation（5 方法）

| 方法 | 说明 |
|------|------|
| HttpRequest | HTTP 请求（23 参数，支持 SSL/代理/Cookie） |
| HttpRequestCom | HTTP 请求（COM 兼容签名） |
| LoadCert | 加载 P12 证书 |
| FreeCert | 释放证书句柄 |
| MergeCookies | 合并 Cookies |
| GetCookie | 从 Cookie 字符串获取指定键值 |

### 3.19 IDLLOperation（18 方法）

InjectThreadDLL, UnloadThreadDLL, InjectHookDLL, UnloadHookDLL, InitThreadInject, InitHookInject, SendHookQuit, InjectHollowProcess, InjectReflectiveDLL, InjectReflectiveProcess, LoadPEFromMemory, FreePEModule, GetPEProcAddress, RemovePEHeader, UnlinkLdrModule, InjectImeDLLByWindow, InjectImeDLLByThread, InjectImePluginByWindow, InjectImePluginByThread

### 3.20 IHookOperation（10 方法）

InstallDetoursHook, InstallInlineHook, InstallVehHook, InstallSuperHook, InstallIATHook, InstallEATHook, UninstallHook, EnableHook, GetHookOriginalAddress, CallOriginal15, RemoveAllHooks

### 3.21 IRegexOperation（10 方法）

SetBackend, GetBackend, SetPattern, SetFlag, GetFlag, Test, Match, Matches, Replace, Split

### 3.22 IFSOOperation（22 方法）

DriveExistsFSO, GetDriveInfoFSO, EnumDrivesFSO, FolderExistsFSO, CreateFolderFSO, DeleteFolderFSO, CopyFolderFSO, MoveFolderFSO, GetFolderInfoFSO, FileExistsFSO, DeleteFileFSO, CopyFileFSO, MoveFileFSO, GetFileInfoFSO, GetSpecialFolderFSO, GetTempNameFSO, GetAbsolutePathNameFSO, GetBaseNameFSO, GetExtensionNameFSO, GetFileNameFSO, GetParentFolderNameFSO, GetShortPathFSO, CreateTextFileFSO, OpenTextFileFSO, AppendToFileFSO

### 3.23 ISmCrypto（8 方法）

SM3Hash, SM4Encrypt, SM4Decrypt, SM2GenerateKey, SM2Encrypt, SM2Decrypt, SM2Sign, SM2Verify

### 3.24 IFileMapping / IShareMemory（17 方法，两者同签名）

| 方法 | 说明 |
|------|------|
| ConstGet(index → value) | 获取常量值 |
| Create(protect, name, maxLow, maxHigh → handle) | 创建映射 |
| GetMax(→ value) | 获取最大容量 |
| Close(handle → ok) | 关闭映射 |
| Open(protect, name → handle) | 打开已有映射 |
| GetPoint(handle, access, len → address) | 映射视图 |
| RelasePoint(address → ok) | 释放视图 |
| WriteBin(address, data, size) | 写入二进制 |
| ReadBin(address, len → data, size) | 读取二进制 |
| WriteText(address, text) | 写入文本 |
| ReadText(address, len → text) | 读取文本 |
| WriteTextA, ReadTextA | ANSI 文本读写 |
| WriteInt, WriteIntP, WriteLong | 写入整数 |
| ReadInt, ReadIntP, ReadLong | 读取整数 |

---

## 4. 额外转发命名空间

这些方法通过 ILeyBc 统一接口的前缀命名方法转发：

| 前缀 | 对应模块 | 方法数 |
|------|---------|:------:|
| FH* | CFileHandleOperation | 11 |
| Dir* | CDirectoryOperation | 22 |
| Api* | CFileOperationEx | 37 |
| Sys* | CSystemOperation | 1 |
| HotkeyGlobal* | CGlobalHotkeyOperation | 5 |
| HotkeyWindow* | CWindowHotkeyOperation | 3 |
| Dialog* | CDialogOperation | 3 |
| DynClass* | CDynamicClassNameOperation | 4 |
| Menu* | CMenuOperation | 15 |
| Map* | CFileMapping | 17 |
| FSO* | CFSOOperation | 25 |

---

## 5. IJScriptEngine JavaScript 引擎

| 方法 | 说明 |
|------|------|
| AddCode(code) | 添加 JS 代码到引擎 |
| Execute(code → result) | 执行 JS 代码 |
| Eval(expression → result) | 求值 JS 表达式 |
| CallFunc(name, args → result) | 调用 JS 函数 |
| Run(→ result) | 运行引擎 |
| AddObject(name, obj) | 添加宿主对象 |
| get_Timeout(→ ms) | 获取超时 |
| put_Timeout(ms) | 设置超时 |
| get_LastError(→ desc) | 获取最后错误 |
| get_LastErrorLine(→ line) | 获取错误行号 |
| get_LastErrorColumn(→ col) | 获取错误列号 |
| Reset() | 重置引擎 |
| SetConsoleCallback(callback) | 设置控制台回调 |

内部注入的 JS 全局对象见 [JS Engine 注入对象报告](#)。

---

## 6. IFileMapping / IShareMemory

| 方法 | 说明 |
|------|------|
| ConstGet | 获取系统常量 |
| Create | 创建内存映射文件 |
| GetMax | 获取最大值 |
| Close | 关闭句柄 |
| Open | 打开已有映射 |
| GetPoint | 获取映射地址 |
| RelasePoint | 释放映射地址 |
| WriteBin / ReadBin | 二进制读写 |
| WriteText / ReadText | Unicode 文本读写 |
| WriteTextA / ReadTextA | ANSI 文本读写 |
| WriteInt / ReadInt | 32 位整数读写 |
| WriteIntP / ReadIntP | 指针读写 |
| WriteLong / ReadLong | 64 位整数读写 |

---

## 7. 额外转发命名空间

所有方法通过 `CLeyBcManager` 的 IDispatch 转发表暴露，共 **641 个 DISPID**。

### 文件句柄操作 (FH 前缀, 11 方法)

FHOpen, FHClose, FHSeek, FHGetPosition, FHToBegin, FHToEnd, FHWriteBytes, FHWriteText, FHReadBytes, FHReadText, FHGetSize

### 目录操作 (Dir 前缀, 22 方法)

DirCreate, DirDelete, DirExists, DirGetSubDirectoryCount, DirCopy, DirMove, DirGetCreateTime, DirGetModifyTime, DirGetAccessTime, DirGetShortPath, DirGetFullPath, DirGetShortName, DirGetDirName, DirMoveByCmd, IsEmptyDir, HasSubDir, GetTempPathDir, GetSystem32Dir, GetSysWOW64Dir, DirGetParentPath, DirGetSpecialFolder, DirEnum

### 文件扩展操作 (Api 后缀, 37 方法)

GetFileNameApi, CopyFileApi, GetExtensionApi, OpenFileApi, CloseFileApi, SeekFileApi, GetFilePositionApi, ToBeginApi, ToEndApi, WriteBytesApi, WriteTextApi, ReadBytesApi, ReadTextApi, GetSizeByHandleApi, IsDirectoryApi, GetDirectoryApi, FileExistsApi, GetAttributesApi, SetAttributesApi, RemoveAttributesApi, GetFileSizeApi, DeleteFileApi, OperateFileApi, DeleteToRecycleBinApi, ForceDeleteApi, MoveFileApi, EnumFilesApi, GetFileVersionApi, ToShortPathApi, ToLongPathApi, GetTextEncodingApi, ExecuteFileApi, GetMimeTypeApi, ExtractIconApi, ExtractIconHandleApi, GetIconFromHandleApi, FindFirstApi, FindNextApi, FindCloseApi

### 系统操作 (Sys 前缀, 1 方法)

ExecuteCmdSys

### 全局热键 (HotkeyGlobal 前缀, 5 方法)

HotkeyGlobalRegister, HotkeyGlobalUnregister, HotkeyGlobalUnregisterAll, HotkeyGlobalPeek, HotkeyGlobalWait

### 窗口热键 (HotkeyWindow 前缀, 3 方法)

HotkeyWindowRegister, HotkeyWindowUnregister, HotkeyWindowUnregisterAll

### 对话框 (Dialog 前缀, 3 方法)

DialogOpenFile, DialogSaveFile, DialogOpenFiles

### 动态类名 (DynClass 前缀, 4 方法)

DynClassRegister, DynClassUnregister, DynClassUnregisterAll, DynClassSetEnabled

### 菜单操作 (Menu 前缀, 15 方法)

MenuClick, MenuGetID, MenuClear, MenuRedraw, MenuDelete, MenuGetCount, MenuGetHandle, MenuGetSubHandle, MenuAddIcon, MenuGetTitle, MenuModifyTitle, MenuAdd, MenuEnum, MenuEnumSubMenu, MenuClickByTitle

### 文件映射 (Map 前缀, 13 方法)

MapConstGet, MapGetMax, MapGetPoint, MapRelasePoint, MapWriteBin, MapReadBin, MapWriteIntP, MapReadIntP, MapClose, MapCreate, MapOpen, MapReadText, MapWriteText

### FSO 兼容 (FSO 后缀, 25 方法)

DriveExistsFSO, GetDriveInfoFSO, EnumDrivesFSO, FolderExistsFSO, CreateFolderFSO, DeleteFolderFSO, CopyFolderFSO, MoveFolderFSO, GetFolderInfoFSO, FileExistsFSO, DeleteFileFSO, CopyFileFSO, MoveFileFSO, GetFileInfoFSO, GetSpecialFolderFSO, GetTempNameFSO, GetAbsolutePathNameFSO, GetBaseNameFSO, GetExtensionNameFSO, GetFileNameFSO, GetParentFolderNameFSO, GetShortPathFSO, CreateTextFileFSO, OpenTextFileFSO, AppendToFileFSO

---

> 最终统计：**38 个接口，30 个 CLSID 组件，641 个 IDispatch 方法，总计 ~1000+ 功能点**。