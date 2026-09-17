# LeyBc COM Library 功能说明书

## 概述

LeyBc COM Library 是一个基于 C++ 的 COM DLL，提供 **29 个接口、706 个公开方法**，涵盖文本处理、图像处理、窗口管理、进程管理、文件操作、磁盘操作、线程同步、编码转换、加解密、进制转换、时间操作、内存操作、汇编/反汇编、HTTP 请求、DLL 注入、Hook 注入、COM 动态调用、SM 国密、FASM 汇编、正则表达式、对话框、热键、菜单、JavaScript 引擎、TCP/HTTP/WebSocket 通信等底层功能。

- **目标系统**：Windows 7 ~ Windows 11（x86 & x64）
- **编译方式**：纯静态编译（/MT），不依赖任何外部 DLL
- 调用方式：COM（CoCreateInstance 或直接 LoadLibrary），推荐使用 CLSID_CLeyBcManager 单 CLSID 入口
- **脚本语言支持**：全部 736 个 `ILeyBc` 分发入口原生支持 IDispatch，可在 VBScript / JScript / PowerShell / Python / C# dynamic 中直接调用

---

## 1. IComponentBase — 组件基类（2 方法）

| 方法 | 原型 | 说明 |
|------|------|------|
| GetComponentName | `(LPOLESTR* name)` | 获取组件名称 |
| GetComponentVersion | `(DWORD* major, DWORD* minor)` | 获取组件版本号 |

---

## 2. ITextOperation — 文本操作（55 方法）

### 基础操作（12 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| FindText | `(text, find, flags, &position)` | 查找文本，返回位置（-1 未找到） |
| ReplaceText | `(text, oldText, newText, flags, &result, &count)` | 替换文本，返回替换次数 |
| SplitText | `(text, delimiter, &parts, &count)` | 分割字符串 |
| JoinText | `(parts, count, separator, &result)` | 合并字符串数组 |
| Substring | `(text, start, length, &result)` | 取子串 |
| Trim | `(text, &result)` | 去除首尾空白 |
| ToUpper | `(text, &result)` | 转大写 |
| ToLower | `(text, &result)` | 转小写 |
| GetLength | `(text, &length)` | 获取文本长度 |
| EncodeBase64 | `(data, dataSize, &result)` | Base64 编码 |
| DecodeBase64 | `(base64, &data, &dataSize)` | Base64 解码 |
| Compare | `(text1, text2, ignoreCase, &result)` | 字符串比较（0 相等） |

### 文本截取（7 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| GetLeft | `(text, subText, &result)` | 取指定文本左边的部分 |
| GetRight | `(text, subText, &result)` | 取指定文本右边的部分 |
| GetMid | `(text, left, right, &result)` | 取指定前后文本中间的部分 |
| GetMidReverse | `(text, left, right, &result)` | 从右边开始取中间 |
| GetMidBatch | `(text, left, right, &results, &count)` | 取中间_批量 |
| GetMidBatchRegex | `(text, left, right, &results, &count)` | 取中间_批量_正则 |

### 行操作（6 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| GetLineContaining | `(text, subText, &result)` | 取文本所在行 |
| GetLineText | `(text, line, &result)` | 取指定行（1-based） |
| RemoveLine | `(text, line, &result)` | 删除指定行 |
| RemoveEmptyLines | `(text, &result)` | 删除空行 |
| RemoveNewLines | `(text, &result)` | 删除换行符 |
| SplitLines | `(text, &lines, &count)` | 按换行符分行 |

### 字符处理（8 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| PadLeftZero | `(text, totalWidth, &result)` | 前补0到指定长度 |
| RemoveAllSpaces | `(text, &result)` | 删除全部空白 |
| CapitalizeFirst | `(text, &result)` | 首字母大写 |
| Reverse | `(text, &result)` | 颠倒文本 |
| SplitChars | `(text, &chars, &count)` | 逐字分割 |
| Escape | `(text, &result)` | 转义正则特殊字符 |

### 判断操作（9 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| StartsWithDigit | `(text, &result)` | 是否以数字开头 |
| IsNumeric | `(text, &result)` | 是否全为数字 |
| IsRepeated | `(text, subText, &result)` | 是否重复文本 |
| HasChinese | `(text, &result)` | 是否有中文 |
| IsAllChinese | `(text, &result)` | 是否全中文 |
| IsAlpha | `(text, &result)` | 是否全字母 |
| IsAllLower | `(text, &result)` | 是否全小写 |
| IsAllUpper | `(text, &result)` | 是否全大写 |
| IsUTF8 | `(text, &result)` | 是否 UTF-8 |

### 提取操作（4 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| ExtractChinese | `(text, &result)` | 只取汉字 |
| ExtractLetters | `(text, &result)` | 只取字母 |
| ExtractDigits | `(text, &result)` | 只取数字 |
| ExtractSymbols | `(text, &result)` | 只取符号 |

### 随机生成（5 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| RandomDigit | `(length, &result)` | 取随机数字字符串 |
| RandomLetter | `(length, capital, &result)` | 取随机字母字符串 |
| RandomChar | `(length, &result)` | 取随机字符（字母+数字） |
| RandomChinese | `(length, &result)` | 取随机汉字 |
| RandomSurname | `(&result)` | 取随机姓氏 |

### 功能增强（2 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| CountOccurrences | `(text, subText, ignoreCase, &count)` | 取出现次数 |
| GetPinyin | `(text, firstOnly, &result)` | 取拼音（字典查找，支持声调） |

### 窗口操作（2 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| SendText | `(hWnd, text)` | 发送文本到指定窗口 |
| SelectAll | `(hWnd)` | 全选编辑框文本 |

### 声音转换（2 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| TextToVoice | `(text, engine, &data, &dataSize)` | 文本转声音，支持 0 电信、1 移动、2 联通、3 QQ、4 百度、5 京东、6 阿里、7 微博；输出字节需 `CoTaskMemFree` 释放 |
| Str2Voice | `(text, engine, &data, &dataSize)` | 文本转声音兼容别名，音频为空返回 `S_FALSE` |

---

## 3. IImageOperation — 图片操作（10 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| CaptureScreen | `(&bitmap)` | 截取全屏 |
| CaptureRect | `(left, top, right, bottom, &bitmap)` | 截取指定区域 |
| CaptureWindow | `(hWnd, &bitmap)` | 截取指定窗口 |
| SaveToFile | `(bitmap, filePath, format)` | 保存图片（支持 jpg/png/bmp/gif） |
| LoadFromFile | `(filePath, &bitmap)` | 从文件加载图片 |
| Resize | `(src, newWidth, newHeight, &dst)` | 缩放图片 |
| Rotate | `(src, angle, &dst)` | 旋转图片 |
| Crop | `(src, left, top, right, bottom, &dst)` | 裁剪图片 |
| GetWidth | `(bitmap, &width)` | 获取图片宽度 |
| GetHeight | `(bitmap, &height)` | 获取图片高度 |

---

## 4. IWindowOperation — 窗口操作（37 方法）

公开窗口句柄统一使用 `ULONGLONG`；文本使用 `BSTR`；布尔值使用 `VARIANT_BOOL`；枚举结果为 `SAFEARRAY(VT_UI8)`。窗口方法占用 DISPID 67-103，通用 `SendMessage` 已从公开接口移除。

| DISPID | 方法 | 参数 | 说明 |
|--------|------|------|------|
| 67 | FindWindowByTitle | `(title, &window)` | 按完整标题查找顶层窗口 |
| 68 | FindWindowByClass | `(className, &window)` | 按类名查找顶层窗口 |
| 69 | FindWindowEx | `(parent, childAfter, className, title, &window)` | 查找匹配的顶层或子窗口 |
| 70 | EnumWindows | `(&windows)` | 枚举顶层窗口 |
| 71 | EnumChildWindows | `(parent, &windows)` | 枚举指定窗口的后代窗口 |
| 72 | GetForegroundWindow | `(&window)` | 获取前台窗口 |
| 73 | GetThreadFocusWindow | `(&window)` | 获取当前线程焦点窗口 |
| 74 | GetDesktopWindow | `(&window)` | 获取桌面窗口 |
| 75 | GetShellWindow | `(&window)` | 获取 Shell 窗口 |
| 76 | GetTaskbarWindow | `(&window)` | 获取任务栏窗口 |
| 77 | GetParentWindow | `(window, &parent)` | 获取父窗口 |
| 78 | GetOwnerWindow | `(window, &owner)` | 获取所有者窗口 |
| 79 | GetWindowText | `(window, &text)` | 获取窗口标题 BSTR |
| 80 | GetWindowClassName | `(window, &className)` | 获取窗口类名 BSTR |
| 81 | SetWindowText | `(window, text)` | 设置窗口标题 |
| 82 | IsWindowValid | `(window, &result)` | 判断窗口句柄是否有效 |
| 83 | IsWindowVisible | `(window, &result)` | 判断窗口是否可见 |
| 84 | IsWindowEnabled | `(window, &result)` | 判断窗口是否启用 |
| 85 | IsWindowMinimized | `(window, &result)` | 判断窗口是否最小化 |
| 86 | IsWindowMaximized | `(window, &result)` | 判断窗口是否最大化 |
| 87 | GetWindowRect | `(window, &left, &top, &right, &bottom)` | 获取窗口矩形 |
| 88 | GetClientRect | `(window, &left, &top, &right, &bottom)` | 获取客户区矩形 |
| 89 | ClientToScreen | `(window, &x, &y)` | 客户区坐标转屏幕坐标 |
| 90 | ScreenToClient | `(window, &x, &y)` | 屏幕坐标转客户区坐标 |
| 91 | SetWindowRect | `(window, left, top, width, height)` | 设置窗口位置和尺寸 |
| 92 | ShowWindow | `(window, cmdShow)` | 改变窗口显示状态 |
| 93 | EnableWindow | `(window, enabled)` | 启用或禁用窗口 |
| 94 | SetTopMost | `(window, topMost)` | 设置或取消置顶 |
| 95 | SetForegroundWindow | `(window)` | 请求将窗口置于前台 |
| 96 | GetWindowProcessId | `(window, &processId)` | 获取所属进程 ID |
| 97 | GetWindowThreadId | `(window, &threadId)` | 获取所属线程 ID |
| 98 | CloseWindow | `(window)` | 异步投递 `WM_CLOSE` |
| 99 | RefreshWindow | `(window)` | 重绘窗口及子窗口 |
| 100 | SetCaptureExclusion | `(window, excluded)` | 设置窗口捕获排除属性 |
| 101 | SetCloseEnabled | `(window, enabled)` | 启用或禁用系统菜单关闭项 |
| 102 | RegisterWindowMessage | `(messageName, &messageId)` | 注册进程间窗口消息 |
| 103 | FlashWindow | `(window, flags, count, timeout)` | 配置窗口闪烁 |

---

## 5. IProcessOperation — 进程操作（113 方法）

### 基础进程操作（12 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| EnumProcesses | `(pids, count)` | 枚举所有进程 PID | ✅ |
| OpenProcess | `(pid, desiredAccess, handle)` | 通过 PID 打开进程句柄 | ✅ |
| OpenProcessByPid | `(pid, desiredAccess, handle)` | 通过 PID 打开进程句柄（别名） | ✅ |
| CloseProcess | `(hObject)` | 关闭进程句柄（兼容各种句柄类型） | ✅ |
| CloseProcessHandle | `(hProcess)` | 关闭进程句柄 | ✅ |
| TerminateProcess | `(pid, exitCode)` | 终止进程（通过 PID） | ✅ |
| TerminateProcessByHandle | `(hProcess, exitCode)` | 通过句柄终止进程 | ✅ |
| TerminateProcessById | `(pid, exitCode)` | 通过 PID 终止进程（NtTerminateProcess） | ✅ |
| CreateProcess | `(commandLine, pid)` | 创建进程 | ✅ |
| CreateProcessAsSystem | `(processPath, pid)` | 以 SYSTEM 权限创建进程 | ✅ |
| CreateProcessWithParent | `(parentPID, commandLine, pid, hThread)` | 创建指定父进程的子进程 | ✅ |
| CreateProcessWithParentByHandle | `(hParent, commandLine, pid, hThread)` | 通过句柄创建指定父进程的子进程 | ✅ |

### 进程恢复/暂停/保护/音量（11 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| SuspendProcess | `(pid)` | 暂停进程（NtSuspendProcess） | ✅ |
| SuspendProcessByHandle | `(hProcess)` | 通过句柄暂停进程 | ✅ |
| ResumeProcess | `(pid)` | 恢复进程（NtResumeProcess） | ✅ |
| ResumeProcessByHandle | `(hProcess)` | 通过句柄恢复进程 | ✅ |
| SetPriorityClass | `(pid, priority)` | 设置进程优先级 | ✅ |
| SetPriorityClassByHandle | `(hProcess, priority)` | 通过句柄设置进程优先级 | ✅ |
| GetPriorityClass | `(pid, priority)` | 获取进程优先级 | ✅ |
| GetPriorityClassByHandle | `(hProcess, priority)` | 通过句柄获取进程优先级 | ✅ |
| SetProcessDEP | `(pid)` | 设置进程 DEP 策略 | ✅ |
| SetProcessVolume | `(pid, volume)` | 设置进程主音量（0-100） | ✅ |
| BeginMonitorProcess | `(eventType, callback)` | 开始监视进程事件 | ✅ |
| StopMonitorProcess | `()` | 停止监视进程事件 | ✅ |
| StartProtectProcess | `(pid)` | 启动进程保护（ProcessBreakOnTermination） | ✅ |
| StopProtectProcess | `(pid)` | 停止进程保护 | ✅ |

### 进程令牌/权限（4 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| AdjustPrivilege | `(privilege)` | 调整进程令牌权限（如 SE_DEBUG） | ✅ |
| OpenProcessToken | `(hProcess, desiredAccess, hToken)` | 打开进程令牌 | ✅ |
| GetTokenInformation | `(hProcess, infoClass, info)` | 获取进程令牌信息 | ✅ |
| IsAdmin | `(isAdmin)` | 判断当前进程是否管理员权限 | ✅ |

### 进程名称/路径/命令行（11 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| GetProcessName | `(pid, name)` | 获取进程名称（映像名） | ✅ |
| GetProcessNameByHandle | `(hProcess, name)` | 通过句柄获取进程名称 | ✅ |
| GetProcessPath | `(pid, path)` | 获取进程路径 | ✅ |
| GetProcessPathByHandle | `(hProcess, path)` | 通过句柄获取进程路径 | ✅ |
| GetProcessUserName | `(pid, userName)` | 获取进程所属用户名 | ✅ |
| GetProcessUserNameByHandle | `(hProcess, userName)` | 通过句柄获取进程所属用户名 | ✅ |
| GetCommandLine | `(pid, cmdLine)` | 获取进程命令行 | ✅ |
| GetCommandLineByHandle | `(hProcess, cmdLine)` | 通过句柄获取进程命令行 | ✅ |
| GetCurrentProcessId | `(pid)` | 获取当前进程 ID | ✅ |
| GetCurrentProcessId2 | `(pid)` | 获取当前进程 ID（NtCurrentPeb 方式） | ✅ |
| GetCurrentProcessPseudoHandle | `(hProcess)` | 获取当前进程伪句柄 | ✅ |

### 进程 ID 查询（7 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| GetPidByName | `(processName, pid)` | 通过进程名称查找 PID | ✅ |
| GetPidsByName | `(processName, pids, count, caseSensitive)` | 通过进程名称查找多个 PID | ✅ |
| GetPidByHandle | `(hProcess, pid)` | 通过句柄获取进程 PID | ✅ |
| GetPidByPort | `(port, pid)` | 通过端口号查找进程 PID | ✅ |
| IsPidValid | `(pid, isValid)` | 判断 PID 是否有效（进程是否存在） | ✅ |
| GetCurrentProcessId | `(pid)` | 获取当前进程 ID | ✅ |

### 父进程/子进程查询（10 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| GetParentProcessId | `(pid, parentPID)` | 获取父进程 PID | ✅ |
| GetParentProcessIdByHandle | `(hProcess, parentPID)` | 通过句柄获取父进程 PID | ✅ |
| GetParentProcessName | `(pid, name)` | 获取父进程名称 | ✅ |
| GetParentProcessNameByHandle | `(hProcess, name)` | 通过句柄获取父进程名称 | ✅ |
| GetChildProcessId | `(hProcess, childPID)` | 获取子进程 PID | ✅ |
| GetChildProcessIdByHandle | `(hProcess, childPID)` | 通过句柄获取子进程 PID | ✅ |
| GetChildProcesses | `(hProcess, childPIDs, count)` | 获取所有子进程 PID 列表 | ✅ |
| GetChildProcessesByHandle | `(hProcess, childPIDs, count)` | 通过句柄获取所有子进程 PID 列表 | ✅ |

### PEB / 模块查询（20 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| GetPebBase | `(pid, pebBase)` | 获取进程 PEB 基址 | ✅ |
| GetPebBaseByHandle | `(hProcess, pebBase)` | 通过句柄获取进程 PEB 基址 | ✅ |
| EnumProcessModules | `(pid, modules, count)` | 枚举进程所有模块 | ✅ |
| GetProcessModules | `(pid, modules, count)` | 获取进程模块列表（别名） | ✅ |
| GetProcessModulesByHandle | `(hProcess, modules, count)` | 通过句柄获取进程模块列表 | ✅ |
| GetModuleBaseAddress | `(pid, moduleName, baseAddr)` | 获取指定模块基址（PEB 遍历 → NtQueryVirtualMemory 降级） | ✅ |
| GetModuleBaseAddressByHandle | `(hProcess, moduleName, baseAddr)` | 通过句柄获取指定模块基址 | ✅ |
| GetModuleEntryPoint | `(pid, moduleBase, entryPoint)` | 获取模块入口点 | ✅ |
| GetModuleEntryPointByHandle | `(hProcess, moduleBase, entryPoint)` | 通过句柄获取模块入口点 | ✅ |
| GetModuleInfo | `(pid, moduleName, base, size, path)` | 获取模块信息（基址、大小、完整路径） | ✅ |
| GetModuleInfoByHandle | `(hProcess, moduleName, base, size, path)` | 通过句柄获取模块信息 | ✅ |
| GetModulePeb | `(pid, moduleBase, pebAddr)` | 获取模块的 PEB 地址 | ✅ |
| GetModulePebByHandle | `(hProcess, moduleBase, pebAddr)` | 通过句柄获取模块的 PEB 地址 | ✅ |
| GetExportFunctions | `(pid, moduleName, funcNames, funcAddrs, count)` | 获取模块导出函数列表（解析 PE 导出表） | ✅ |
| GetExportFunctionsByHandle | `(hProcess, moduleName, funcNames, funcAddrs, count)` | 通过句柄获取模块导出函数列表 | ✅ |
| GetExportFunctionsEx | `(pid, moduleName, funcNames, funcAddrs, count)` | 获取模块导出函数列表（扩展，返回实际地址） | ✅ |
| GetExportFunctionsExByHandle | `(hProcess, moduleName, funcNames, funcAddrs, count)` | 通过句柄获取模块导出函数列表（扩展） | ✅ |
| GetFunctionAddress | `(pid, moduleName, funcName, funcAddr)` | 获取模块内指定函数地址（ANSI 名搜索） | ✅ |
| GetFunctionAddressByHandle | `(hProcess, moduleName, funcName, funcAddr)` | 通过句柄获取指定函数地址 | ✅ |
| GetFunctionAddressFast | `(pid, moduleName, funcName, funcAddr)` | 快速获取模块内函数地址（同 GetFunctionAddress） | ✅ |
| GetFunctionAddressFastByHandle | `(hProcess, moduleName, funcName, funcAddr)` | 通过句柄快速获取函数地址 | ✅ |

### 内存/IO/线程查询（7 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| GetProcessMemoryUsage | `(pid, memoryUsage)` | 获取进程内存使用量（工作集） | ✅ |
| GetProcessMemoryUsageByHandle | `(hProcess, memoryUsage)` | 通过句柄获取进程内存使用量 | ✅ |
| GetIoCounters | `(hProcess, readOp, writeOp, otherOp, readXfer, writeXfer, otherXfer)` | 获取进程 IO 计数 | ✅ |
| GetMainThreadId | `(pid, threadId)` | 获取进程主线程 ID | ✅ |
| GetMainThreadIdByHandle | `(hProcess, threadId)` | 通过句柄获取进程主线程 ID | ✅ |
| GetDEPPolicy | `(policy)` | 获取系统 DEP 策略 | ✅ |
| GetProcessWindowHandle | `(pid, hWnd)` | 获取进程主窗口句柄 | ✅ |
| GetProcessWindowHandleByEnum | `(pid, hWnd)` | 枚举窗口获取进程窗口句柄（别名） | ✅ |

### 端口/网络/进程检测（8 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| GetPortByPid | `(pid, port)` | 通过进程 PID 查找端口号（首个） | ✅ |
| GetPidByPort | `(port, pid)` | 通过端口号查找进程 PID | ✅ |
| IsPortOccupied | `(port, pid, occupied)` | 判断端口是否被占用 | ✅ |
| OpenPortByPid | `(pid, ports, count)` | 通过 PID 获取该进程所有端口 | ✅ |
| IsX64Process | `(pid, isX64)` | 判断进程是否为 64 位 | ✅ |
| IsX64ProcessByHandle | `(hProcess, isX64)` | 通过句柄判断进程是否为 64 位 | ✅ |
| IsDebugged | `(isDebugged)` | 判断当前进程是否被调试 | ✅ |
| IsDebuggedEx | `(isDebugged)` | 判断当前进程是否被调试（PEB.BeingDebugged） | ✅ |

### 远程线程 / Hijack（6 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| CreateRemoteThread | `(pid, startAddr, param, threadId, hThread)` | 创建远程线程 | ✅ |
| CreateRemoteThreadByHandle | `(hProcess, startAddr, param, threadId, hThread)` | 通过句柄创建远程线程 | ✅ |
| CreateHijackThread | `(pid, shellCode, shellCodeSize, bindThreadId, delay)` | 创建 Hijack 线程（挂起目标线程，修改上下文执行 shellcode） | ✅ |
| CreateHijackThreadByHandle | `(hProcess, shellCode, shellCodeSize, bindThreadId, delay)` | 通过句柄创建 Hijack 线程 | ✅ |
| DuplicateObjectHandle | `(hSrcProc, hSrcHandle, hTgtProc, hTgtHandle, inherit, options)` | 复制对象句柄 | ✅ |
| InjectDLL | `(pid, dllPath)` | 远程线程注入 DLL | ✅ |

### ShellCode 执行（10 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| ExecuteAssemblyCode | `(pid, shellCode, size, result)` | 远程执行汇编代码（NtCreateThreadEx） | ✅ |
| ExecuteAssemblyCodeByHandle | `(hProcess, shellCode, size, result)` | 通过句柄远程执行汇编代码 | ✅ |
| ExecuteAssemblyCodeByThread | `(pid, shellCode, size, bindThreadId, delay, result)` | 通过绑定线程执行汇编代码（线程劫持） | ✅ |
| ExecuteAssemblyCodeByThread32 | `(hProcess, shellCode, size, bindThreadId, delay)` | 32位绑定线程执行汇编代码 | ✅ |
| ExecuteAssemblyCodeByThreadByHandle | `(hProcess, shellCode, size, bindThreadId, delay, result)` | 通过句柄+绑定线程执行汇编代码 | ✅ |
| ExecuteAssemblyCodeByTimer | `(pid, shellCode, size, delay, result)` | 通过定时器执行汇编代码 | ✅ |
| ExecuteAssemblyCodeByTimerByHandle | `(hProcess, shellCode, size, delay, result)` | 通过句柄+定时器执行汇编代码 | ✅ |
| ExecuteAssemblyCodeEx | `(pid, shellCode, size, result)` | 远程执行汇编代码（扩展） | ✅ |
| ExecuteAssemblyCodeExByHandle | `(hProcess, shellCode, size, result)` | 通过句柄远程执行汇编代码（扩展） | ✅ |
| InvokeProcessFunction | `(pid, shellCode, size, result)` | 在目标进程中执行 shellcode | ✅ |
| InvokeProcessFunctionByHandle | `(hProcess, shellCode, size, result)` | 通过句柄执行 shellcode | ✅ |

### 句柄枚举/防注入（6 方法）
| 方法 | 参数 | 说明 | 状态 |
|------|------|------|:----:|
| EnumProcessHandles | `(pid, handles, types, count)` | 枚举进程所有句柄（NtQuerySystemInformation） | ✅ |
| EnumProcessHandlesByHandle | `(hProcess, handles, types, count)` | 通过句柄枚举进程句柄 | ✅ |
| CloseListHandle | `(pid, hHandle)` | 关闭句柄列表中的指定句柄 | ✅ |
| CloseListHandle2 | `(hProcess, hHandle)` | 通过句柄关闭指定句柄 | ✅ |
| PreventDllInjection | `(pid)` | 禁止 DLL 注入（修改 PEB） | ✅ |
| PreventDllInjectionByHandle | `(hProcess)` | 通过句柄禁止 DLL 注入 | ✅ |

### 状态说明
| 图标 | 含义 |
|:----:|------|
| ✅ | 已实现，测试通过 |
| ⏳ | 未实现（返回 E_NOTIMPL），待后续开发 |

---

## 6. IFileOperation — 文件操作（12 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| CreateFile | `(path)` | 创建空文件 |
| DeleteFile | `(path)` | 删除文件 |
| CopyFile | `(srcPath, dstPath, overwrite)` | 复制文件 |
| MoveFile | `(srcPath, dstPath)` | 移动文件 |
| ReadTextFile | `(path, &content)` | 读取文本文件 |
| WriteTextFile | `(path, content)` | 写入文本文件 |
| ReadBinaryFile | `(path, &data, &size)` | 读取二进制文件 |
| WriteBinaryFile | `(path, data, size)` | 写入二进制文件 |
| FileExists | `(path, &exists)` | 检查文件是否存在 |
| GetFileSize | `(path, &size)` | 获取文件大小（字节） |
| CreateDirectory | `(path)` | 创建目录 |
| EnumFiles | `(dir, pattern, &files, &count)` | 枚举目录中的文件（支持通配符） |

---

## 7. IDiskOperation — 磁盘操作（7 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| EnumDrives | `(&drives)` | 枚举所有驱动器（如 "C:,\0D:,\0"） |
| GetTotalSize | `(drive, &size)` | 获取磁盘总容量 |
| GetFreeSize | `(drive, &size)` | 获取磁盘可用空间 |
| GetFileSystemName | `(drive, &name)` | 获取文件系统（如 "NTFS"） |
| GetVolumeLabel | `(drive, &label)` | 获取卷标 |
| GetVolumeSerial | `(drive, &serial)` | 获取卷序列号 |
| GetDriveType | `(drive, &type)` | 判断驱动器类型（0=未知，1=无根目录，2=可移动，3=固定，4=网络，5=CD-ROM，6=RAM） |

---

## 8. IThreadOperation — 线程操作（7 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| CreateThread | `(func, param, &handle)` | 创建线程 |
| SuspendThread | `(handle)` | 暂停线程 |
| ResumeThread | `(handle)` | 恢复线程 |
| WaitForThread | `(handle, timeout)` | 等待线程结束（INFINITE=无限等待） |
| TerminateThread | `(handle, exitCode)` | 强制终止线程 |
| GetExitCode | `(handle, &exitCode)` | 获取线程退出码（STILL_ACTIVE=仍在运行） |
| GetCurrentThreadId | `(&threadId)` | 获取当前线程 ID |

---

## 9. IMultiThreadOperation — 多线程同步（13 方法）

### 临界区
| 方法 | 参数 | 说明 |
|------|------|------|
| CreateCriticalSection | `(CRITICAL_SECTION* cs)` | 初始化临界区 |
| EnterCriticalSection | `(CRITICAL_SECTION* cs)` | 进入临界区 |
| LeaveCriticalSection | `(CRITICAL_SECTION* cs)` | 离开临界区 |
| DeleteCriticalSection | `(CRITICAL_SECTION* cs)` | 删除临界区 |

### 事件
| 方法 | 参数 | 说明 |
|------|------|------|
| CreateEvent | `(manualReset, initialState, &handle)` | 创建事件对象 |
| SetEvent | `(handle)` | 设置事件为有信号 |
| ResetEvent | `(handle)` | 重置事件为无信号 |

### 等待
| 方法 | 参数 | 说明 |
|------|------|------|
| WaitForSingleObject | `(handle, timeout)` | 等待单个对象 |
| WaitForMultipleObjects | `(count, handles, waitAll, timeout)` | 等待多个对象 |

### 互斥体和信号量
| 方法 | 参数 | 说明 |
|------|------|------|
| CreateMutex | `(initialOwner, name, &handle)` | 创建互斥体 |
| ReleaseMutex | `(handle)` | 释放互斥体 |
| CreateSemaphore | `(initialCount, maximumCount, &handle)` | 创建信号量 |
| ReleaseSemaphore | `(handle, releaseCount)` | 释放信号量 |

---

## 10. ITextEncoding — 编码转换（6 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| AnsiToUtf8 | `(ansiText, &utf8)` | ANSI 转 UTF-8 |
| Utf8ToAnsi | `(utf8Text, &ansi)` | UTF-8 转 ANSI |
| Utf8ToUtf16 | `(utf8Text, &utf16)` | UTF-8 转 UTF-16 |
| Utf16ToUtf8 | `(utf16Text, &utf8)` | UTF-16 转 UTF-8 |
| AnsiToUtf16 | `(ansiText, &utf16)` | ANSI 转 UTF-16 |
| Utf16ToAnsi | `(utf16Text, &ansi)` | UTF-16 转 ANSI |

---

## 11. ICryptoOperation — 加解密（37 方法）

### 摘要算法（11 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| MD5Hash | `(data, dataSize, &hash, &hashSize)` | MD5 哈希（16 字节） |
| MD5HashString | `(text, &hash)` | MD5 哈希（输入字符串，输出十六进制） |
| MD2Hash | `(data, dataSize, &hash, &hashSize)` | MD2 哈希（16 字节） |
| MD4Hash | `(data, dataSize, &hash, &hashSize)` | MD4 哈希（16 字节） |
| SHA1Hash | `(data, dataSize, &hash, &hashSize)` | SHA1 哈希（20 字节） |
| SHA224Hash | `(data, dataSize, &hash, &hashSize)` | SHA224 哈希（28 字节） |
| SHA256Hash | `(data, dataSize, &hash, &hashSize)` | SHA256 哈希（32 字节） |
| SHA384Hash | `(data, dataSize, &hash, &hashSize)` | SHA384 哈希（48 字节） |
| SHA512Hash | `(data, dataSize, &hash, &hashSize)` | SHA512 哈希（64 字节） |
| SHA3Hash | `(data, dataSize, &hash, &hashSize)` | SHA3-256 哈希（32 字节，KeccakF1600） |
| CRC32Hash | `(data, dataSize, &crc)` | CRC32 校验值 |

### HMAC 消息认证码（6 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| HMAC_MD5 | `(data, dataSize, key, keySize, &hash, &hashSize)` | HMAC-MD5 |
| HMAC_SHA1 | `(data, dataSize, key, keySize, &hash, &hashSize)` | HMAC-SHA1 |
| HMAC_SHA224 | `(data, dataSize, key, keySize, &hash, &hashSize)` | HMAC-SHA224 |
| HMAC_SHA256 | `(data, dataSize, key, keySize, &hash, &hashSize)` | HMAC-SHA256 |
| HMAC_SHA384 | `(data, dataSize, key, keySize, &hash, &hashSize)` | HMAC-SHA384 |
| HMAC_SHA512 | `(data, dataSize, key, keySize, &hash, &hashSize)` | HMAC-SHA512 |

### 对称加密（8 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| AESEncrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | AES 加密（128/192/256，CBC/ZeroPadding） |
| AESDecrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | AES 解密 |
| DESEncrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | DES 加密（56-bit，CBC/ZeroPadding） |
| DESDecrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | DES 解密 |
| ThreeDESEncrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | 3DES 加密（112/168-bit，CBC/ZeroPadding） |
| ThreeDESDecrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | 3DES 解密 |
| XOREncrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | XOR 加密/解密 |
| RC4Encrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | RC4 加密/解密 |

### 非对称 RSA（5 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| RSAKeyGenerate | `(keyBits, &publicKey, &publicKeySize, &privateKey, &privateKeySize)` | RSA 密钥对生成（1024~16384） |
| RSAPublicEncrypt | `(data, dataSize, publicKey, publicKeySize, &result, &resultSize)` | RSA 公钥加密（OAEP-SHA1） |
| RSAPrivateDecrypt | `(data, dataSize, privateKey, privateKeySize, &result, &resultSize)` | RSA 私钥解密 |
| RSASign | `(data, dataSize, privateKey, privateKeySize, &signature, &signatureSize)` | RSA 私钥签名（PKCS1-v1.5-SHA256） |
| RSAVerify | `(data, dataSize, publicKey, publicKeySize, signature, signatureSize, &valid)` | RSA 公钥验签 |

### ECC（4 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| ECDSASign | `(data, dataSize, privateKey, privateKeySize, &signature, &signatureSize)` | ECDSA 签名（P-256） |
| ECDSAVerify | `(data, dataSize, publicKey, publicKeySize, signature, signatureSize, &valid)` | ECDSA 验签 |
| ECDHKeyExchange | `(privateKey, privateKeySize, peerPublicKey, peerPublicKeySize, &sharedSecret, &sharedSecretSize)` | ECDH 密钥交换（P-256） |

### DSA 签名（2 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| DSASign | `(data, dataSize, privateKey, privateKeySize, &signature, &signatureSize)` | DSA 签名 |
| DSAVerify | `(data, dataSize, publicKey, publicKeySize, signature, signatureSize, &valid)` | DSA 验签 |

### DH 密钥交换（1 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| DHKeyExchange | `(privateKey, privateKeySize, peerPublicKey, peerPublicKeySize, &sharedSecret, &sharedSecretSize)` | DH 密钥交换 |

---

## 12. IBaseConversion — 进制转换（6 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| DecToBin | `(value, &result)` | 十进制转二进制字符串 |
| DecToOct | `(value, &result)` | 十进制转八进制字符串 |
| DecToHex | `(value, &result)` | 十进制转十六进制字符串 |
| StrToDec | `(value, base, &result)` | 任意进制字符串转十进制 |
| DecToBase | `(value, base, &result)` | 十进制转任意进制字符串 |
| BigStrToDec | `(value, base, &result)` | 大数字符串（最长 256-bit）转十进制 |

---

## 13. ITimeOperation — 时间操作（10 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| GetSystemTime | `(&SYSTEMTIME)` | 获取 UTC 系统时间 |
| GetLocalTime | `(&SYSTEMTIME)` | 获取本地时间 |
| FormatTime | `(SYSTEMTIME*, format, &result)` | 格式化时间（如 "yyyy-MM-dd HH:mm:ss"） |
| GetTimestamp | `(&timestamp)` | 获取 Unix 时间戳（秒） |
| GetHighResTimestamp | `(&timestamp)` | 获取高精度时间戳（微秒） |
| TimestampToSystemTime | `(timestamp, &SYSTEMTIME)` | 时间戳转 SYSTEMTIME |
| SystemTimeToTimestamp | `(SYSTEMTIME*, &timestamp)` | SYSTEMTIME 转时间戳 |
| Sleep | `(milliseconds)` | 睡眠指定毫秒 |
| StartTimer | `(&startTime)` | 启动计时器，返回开始时间 |
| StopTimer | `(startTime, &elapsed)` | 获取经过时间（微秒） |

---

## 14. IMemoryOperation — 内存操作（6 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| ReadProcessMemory | `(handle, address, buffer, size, &bytesRead)` | 读取进程内存 |
| WriteProcessMemory | `(handle, address, buffer, size, &bytesWritten)` | 写入进程内存 |
| VirtualAllocEx | `(handle, address, size, allocationType, protect, &result)` | 申请虚拟内存 |
| VirtualFreeEx | `(handle, address, size, freeType)` | 释放虚拟内存 |
| VirtualQueryEx | `(handle, address, &MBI)` | 查询虚拟内存信息 |
| InjectDLL | `(pid, dllPath)` | 远程线程注入 DLL |

---

## 15. IAssemblyOperation — 汇编操作（18 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| AssembleX86 | `(asmCode, &machineCode, &codeSize)` | X86 汇编指令转机器码 |
| AssembleX64 | `(asmCode, &machineCode, &codeSize)` | X64 汇编指令转机器码 |
| ExecuteAsmX86 | `(machineCode, codeSize, &result)` | 执行 X86 机器码（仅 x86 平台） |
| Push | `(value, &machineCode, &codeSize)` | 生成 Push 指令机器码 |
| Call | `(currentAddr, targetAddr, &machineCode, &codeSize)` | 生成 Call 指令机器码（相对偏移） |
| CallDwordPtr | `(targetAddr, &machineCode, &codeSize)` | 生成 FF 15 Call dword ptr [addr] |
| CallQwordPtr | `(currentAddr, targetAddr, &machineCode, &codeSize)` | 生成 FF 15 Call qword ptr [rip+offset] |
| Jmp | `(currentAddr, targetAddr, &machineCode, &codeSize)` | 生成 Jmp 指令机器码（相对偏移） |
| JmpLong | `(targetAddr, &machineCode, &codeSize)` | 生成 jmp qword ptr [addr] 长跳转 |
| Pushad_X64 | `(&machineCode, &codeSize)` | 生成 x64 pushad 等价序列 |
| Popad_X64 | `(&machineCode, &codeSize)` | 生成 x64 popad 等价序列 |
| AsmEmpty | `()` | 清空汇编指令缓冲区 |
| AsmAdd | `(asmCode)` | 添加汇编指令到缓冲区 |
| AsmTakeShellcode | `(platform, clearResult, outputResult, &machineCode, &codeSize, &asmText)` | 编译并取回 Shellcode |
| AsmGetText | `(&result)` | 取缓冲区汇编文本 |
| AsmPushadX64 | `()` | 缓冲区添加 x64 pushad |
| AsmPopadX64 | `()` | 缓冲区添加 x64 popad |

---

## 16. IDisassemblyOperation — 反汇编操作（4 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| DisassembleX86 | `(machineCode, maxSize, &asmText, &instrLength)` | X86 反汇编单条指令 |
| DisassembleX64 | `(machineCode, maxSize, &asmText, &instrLength)` | X64 反汇编单条指令 |
| DisassembleBatch | `(machineCode, codeSize, is64Bit, &asmLines, &count)` | 批量反汇编 |
| GetInstructionLength | `(machineCode, is64Bit, &length)` | 获取指令长度 |

---

## 17. IHttpOperation — HTTP 操作（6 方法）

### HTTP 请求
```cpp
HttpRequest(url, method, submitText, certPath, certPassword,
    requestCookies, &responseCookies, requestHeaders, &responseHeaders,
    &responseStatusCode, disableRedirect, postData, postDataSize,
    proxyAddress, timeoutSeconds, siteUser, sitePassword,
    proxyUser, proxyPassword, autoMergeCookies, autoFillHeaders,
    normalizeHeaderCase, &responseData, &responseDataSize)
```
- `url`：请求地址
- `method`：请求方法（0=GET, 1=POST, 2=PUT, 3=DELETE, 4=HEAD, 5=PATCH, 6=OPTIONS, 7=TRACE, 8=CONNECT）
- `submitText`：提交文本（POST/PUT 的表单数据）
- `postData` / `postDataSize`：二进制提交数据（与 submitText 二选一）
- `certPath` / `certPassword`：P12 证书路径及密码（可空）
- `requestCookies` / `responseCookies`：请求/响应 Cookie
- `requestHeaders` / `responseHeaders`：请求/响应头（每行一个）
- `proxyAddress` / `proxyUser` / `proxyPassword`：代理设置（可空）
- `timeoutSeconds`：超时秒数
- `disableRedirect`：是否禁止自动重定向
- `autoMergeCookies` / `autoFillHeaders` / `normalizeHeaderCase`：自动处理标记

### 扩展方法
| 方法 | 参数 | 说明 |
|------|------|------|
| HttpRequestCom | 同 HttpRequest | HTTP 请求（COM 组件调用优化版） |
| LoadCert | `(certPath, certPassword, &handle)` | 加载 P12 证书 |
| FreeCert | `(handle)` | 释放证书句柄 |
| MergeCookies | `(requestCookies, responseCookies, &merged)` | 合并 Cookie（同名以服务端为准） |
| GetCookie | `(cookies, cookieName, &value)` | 从 Cookie 字符串中取单条值 |

---

## 18. ITextEncoding — 编码扩展（15 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| AnsiToUtf8 | `(ansiText, &utf8)` | ANSI 转 UTF-8 |
| Utf8ToAnsi | `(utf8Text, &ansi)` | UTF-8 转 ANSI |
| Utf8ToUtf16 | `(utf8Text, &utf16)` | UTF-8 转 UTF-16 |
| Utf16ToUtf8 | `(utf16Text, &utf8)` | UTF-16 转 UTF-8 |
| AnsiToUtf16 | `(ansiText, &utf16)` | ANSI 转 UTF-16 |
| Utf16ToAnsi | `(utf16Text, &ansi)` | UTF-16 转 ANSI |
| UrlEncode | `(text, keepAlphaNum, isUtf8, &result)` | URL 编码 |
| UrlDecode | `(text, isUtf8, &result)` | URL 解码 |
| Utf8ToUnicode | `(utf8Bytes, &result)` | UTF-8 字节转 Unicode 字符串 |
| UnicodeToUtf8 | `(text, &result)` | Unicode 字符串转 UTF-8 字节 |
| Gb2312ToUtf8 | `(gb2312, &result)` | GB2312 转 UTF-8 |
| Utf8ToGb2312 | `(utf8, &result)` | UTF-8 转 GB2312 |
| Base64Encode | `(data, dataSize, &result)` | Base64 编码 |
| Base64Decode | `(base64, &data, &dataSize)` | Base64 解码 |
| TextToUcs2 | `(text, &result)` | 文本转 UCS2 十六进制表示 |

---

## 19. IComInvoke — COM 动态调用（8 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| CreateObject | `(progId, &hObj)` | 通过 ProgID 创建 COM 对象，返回句柄 |
| CreateObjectFromFile | `(dllPath, clsid, &hObj)` | 免注册加载 COM DLL |
| GetProperty | `(hObj, propName, &result)` | 获取 COM 对象属性值 |
| PutProperty | `(hObj, propName, value)` | 设置 COM 对象属性值 |
| Invoke | `(hObj, methodName, argCount, args, &result)` | 调用 COM 方法（有返回值） |
| InvokeNoResult | `(hObj, methodName, argCount, args)` | 调用 COM 方法（无返回值） |
| ReleaseObject | `(hObj)` | 释放 COM 对象 |
| GetInterfaceInfo | `(hObj, &info)` | 获取接口信息（属性/方法列表） |

---

## 20. IFasmOperation — FASM 汇编（1 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| FasmAssemble | `(asmCode, platform, &machineCode, &codeSize)` | FASM 汇编，platform: 0=32位, 1=64位 |

---

## 21. ISmCrypto — SM 国密（8 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| SM3Hash | `(data, dataSize, &hash, &hashSize)` | SM3 哈希（32 字节） |
| SM4Encrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | SM4 加密（128-bit，ECB/PKCS7） |
| SM4Decrypt | `(data, dataSize, key, keySize, &result, &resultSize)` | SM4 解密（128-bit，ECB/PKCS7） |
| SM2GenerateKey | `(&publicKey, &publicKeySize, &privateKey, &privateKeySize)` | SM2 密钥对生成（公钥 64 字节，私钥 32 字节） |
| SM2Encrypt | `(data, dataSize, publicKey, publicKeySize, &result, &resultSize)` | SM2 公钥加密（C1C2C3 模式） |
| SM2Decrypt | `(data, dataSize, privateKey, privateKeySize, &result, &resultSize)` | SM2 私钥解密 |
| SM2Sign | `(data, dataSize, privateKey, privateKeySize, userId, &signature, &sigSize)` | SM2 数字签名（r||s，64 字节） |
| SM2Verify | `(data, dataSize, signature, sigSize, publicKey, publicKeySize, userId, &result)` | SM2 验签 |

---

## 22. IHookOperation — Hook 操作（11 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| InstallDetoursHook | `(targetAddr, callbackAddr, &hookHandle, &originalAddr)` | Detours 风格 API Hook |
| InstallInlineHook | `(hookAddr, hookLength, addCode, addCodeSize, callbackAddr, returnCode, returnCodeSize, coverOriginalCode, fastStackMode, &hookHandle)` | Inline Hook（支持附加字节集、回调返回决策、自定义返回代码） |
| InstallVehHook | `(hookAddr, callbackAddr, &hookHandle)` | VEH 无痕 Hook（硬件断点，不修改目标代码） |
| InstallSuperHook | `(pid, hookAddr, hookLength, shellCode, shellCodeSize, coverOriginalCode, &hookHandle)` | 跨进程 Super Hook |
| InstallIATHook | `(module, importDll, funcName, ordinal, callbackAddr, &hookHandle, &originalAddr)` | IAT Hook |
| InstallEATHook | `(module, funcName, ordinal, callbackAddr, &hookHandle, &originalAddr)` | EAT Hook |
| UninstallHook | `(hookHandle)` | 卸载 Hook |
| EnableHook | `(hookHandle, enable)` | 启用/暂停 Hook |
| GetHookOriginalAddress | `(hookHandle, &originalAddr)` | 获取原函数地址 |
| CallOriginal15 | `(hookHandle, p1..p15, &returnValue)` | 以最多 15 参数调用原函数 |
| RemoveAllHooks | `()` | 卸载所有 Hook |

---

## 23. IDLLOperation — DLL 注入操作（19 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| InjectThreadDLL | `(pid, dllPath)` | 线程注入 DLL（CreateRemoteThread + LoadLibraryW） |
| UnloadThreadDLL | `(pid, dllName)` | 线程卸载 DLL |
| InjectHookDLL | `(hookType, dllPath, hookProcName, threadId, &hook)` | 全局钩子注入 DLL（支持 8 种钩子类型） |
| UnloadHookDLL | `(hook)` | 卸载钩子 |
| InitThreadInject | `(callback, hotKeyMod, hotKeyVk)` | 线程注入初始化（DLL 内部调用） |
| InitHookInject | `(callback, hookType, hotKeyMod, hotKeyVk)` | 钩子注入初始化（DLL 内部调用） |
| SendHookQuit | `(hWnd)` | 发送退出消息 |
| InjectHollowProcess | `(targetPath, dllPath, &processId)` | 镂空注入：创建挂起进程，修改 RIP/EIP 调用 LoadLibraryW |
| InjectReflectiveDLL | `(pid, dllData, dllSize)` | 反射注入：内存映射 DLL 到运行中进程 |
| InjectReflectiveProcess | `(targetPath, dllData, dllSize, &processId)` | 反射注入（创建新进程） |
| LoadPEFromMemory | `(dllData, dllSize, flags, &module)` | 从内存加载 PE |
| FreePEModule | `(module)` | 释放内存加载的 PE |
| GetPEProcAddress | `(module, procName, &procAddr)` | 获取 PE 导出函数地址 |
| RemovePEHeader | `(module)` | 移除 PE 文件头（隐藏模块） |
| UnlinkLdrModule | `(module)` | 从 LDR_MODULE 链表断开（隐藏模块） |
| InjectImeDLLByWindow | `(targetWindow, imeFilePath, &keyboardLayout)` | 通过窗口注入 IME DLL |
| InjectImeDLLByThread | `(threadId, imeFilePath, &keyboardLayout)` | 通过线程注入 IME DLL |
| InjectImePluginByWindow | `(targetWindow, imeFilePath, pluginPath, param, &keyboardLayout)` | 窗口注入 IME 插件 |
| InjectImePluginByThread | `(threadId, imeFilePath, pluginPath, param, &keyboardLayout)` | 线程注入 IME 插件 |

---

## 24. IRegexOperation — 正则表达式操作（11 方法）

| 方法 | 参数 | 说明 |
|------|------|------|
| SetBackend | `(dwBackend)` | 设置正则引擎后端（0=VBScript.RegExp、1=std::regex、2=PCRE2） |
| GetBackend | `(&dwBackend)` | 获取当前后端类型 |
| SetPattern | `(lpszPattern)` | 设置正则表达式模式 |
| SetFlag | `(dwFlag, bValue)` | 设置标志位：0x1=IGNORECASE、0x2=GLOBAL、0x4=MULTILINE、0x8=SINGLELINE |
| GetFlag | `(dwFlag, &bValue)` | 获取标志位状态 |
| Test | `(lpszText, &pbMatch)` | 测试文本是否匹配 |
| Match | `(lpszText, &ppszMatch)` | 获取第一个匹配结果 |
| Matches | `(lpszText, &pppszMatches, &pnCount)` | 获取所有匹配结果 |
| Replace | `(lpszText, lpszReplacement, &ppszResult)` | 替换匹配文本 |
| Split | `(lpszText, &pppszParts, &pnCount)` | 按正则分割文本 |
| RegexEscape | `(lpszText, &ppszResult)` | 转义正则特殊字符 |

### 后端特性
| 后端 | 值 | 说明 | COSE | GLOBAL | MULTILINE | SINGLELINE |
|------|:--:|------|:----:|:------:|:---------:|:----------:|
| VBScript.RegExp | 0 | Windows 自带，通过 IDispatch 调用 | ✓ | ✓ | ✓ | ✓ |
| std::regex | 1 | C++ 标准库，零依赖 | ✓ | ✓ | 部分（内联`(?m)`） | ✗ |
| PCRE2 | 2 | PCRE2 静态库，强兼容 | ✓ | ✓ | ✓ | ✗ |

### 标志位定义
```c
#define REGEX_FLAG_IGNORECASE  0x0001  // 忽略大小写
#define REGEX_FLAG_GLOBAL      0x0002  // 全局匹配
#define REGEX_FLAG_MULTILINE   0x0004  // 多行模式
#define REGEX_FLAG_SINGLELINE  0x0008  // 单行模式（VBScript.RegExp 专用）
```

---

## IJScriptEngine — JavaScript 引擎（13 方法）

基于 ChakraCore（Edge/IE 的 JavaScript 引擎）的 COM 封装，对标 MSScriptControl.ScriptControl 接口风格。

### 代码执行
| 方法 | 参数 | 说明 |
|------|------|------|
| AddCode | `(code)` | 添加并立即执行脚本代码（注入全局变量/函数） |
| Execute | `(code, &result)` | 执行脚本代码块，返回最后一个表达式的值 |
| Eval | `(expression, &result)` | 计算表达式，返回结果 |
| Run | `(&result)` | 执行所有通过 AddCode 添加的代码 |
| CallFunc | `(name, args, &result)` | 调用已定义的 JS 函数，args 为 SAFEARRAY(VT_VARIANT) |

### 对象注入
| 方法 | 参数 | 说明 |
|------|------|------|
| AddObject | `(name, obj)` | 将 COM 对象注入到 JS 全局作用域 |

### 属性
| 属性 | 类型 | 说明 |
|------|------|------|
| Timeout | LONG | 脚本执行超时时间（毫秒），0=不限 |

### 错误信息
| 方法 | 参数 | 说明 |
|------|------|------|
| get_LastError | `(&desc)` | 获取最后错误描述 |
| get_LastErrorLine | `(&line)` | 获取错误行号 |
| get_LastErrorColumn | `(&col)` | 获取错误列号 |

### 生命周期
| 方法 | 参数 | 说明 |
|------|------|------|
| Reset | `()` | 重置引擎，清除所有状态 |
| SetConsoleCallback | `(callback)` | 设置日志回调（预留） |

### 浏览器环境支持
引擎注入以下浏览器全局 API：console（log/warn/error/info）、setTimeout/clearTimeout、setInterval/clearInterval、atob/btoa、performance.now、TextEncoder/TextDecoder、URL/URLSearchParams、fetch、XMLHttpRequest、WebSocket、Buffer、window/self/globalThis 别名。

### 依赖
- **ChakraCore.dll**：运行时动态加载（x86: third_party/ChakraCore/bin/x86/, x64: third_party/ChakraCore/bin/x64/）

---

## 26. ITcpOperation — TCP 操作（13 方法）

基于 HP-Socket v6.0.5 IOCP 框架，提供高性能 TCP Server 和 TCP Client 功能。所有网络 I/O 由 HP-Socket 内部 IOCP 线程池管理，支持数千并发连接。

### TCP Server（8 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| ServerStart | `(bindAddr, port)` | 启动 TCP 服务器，绑定地址和端口 |
| ServerStop | `()` | 停止 TCP 服务器，断开所有连接 |
| ServerSend | `(connId, data, dataSize, &sent)` | 向指定客户端发送数据 |
| ServerBroadcast | `(data, dataSize)` | 向所有客户端广播数据 |
| ServerDisconnect | `(connId, force)` | 断开指定客户端连接 |
| ServerGetConnectionCount | `(&count)` | 获取当前客户端连接数 |
| ServerGetAllConnections | `(&connIds, &count)` | 获取所有客户端连接 ID 列表 |
| ServerGetRemoteAddress | `(connId, &address, &port)` | 获取客户端远程地址和端口 |

### TCP Client（4 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| ClientConnect | `(remoteAddr, port, async)` | 连接远程 TCP 服务器 |
| ClientSend | `(data, dataSize, &sent)` | 发送数据到远程服务器 |
| ClientClose | `()` | 关闭客户端连接 |
| ClientIsConnected | `(&connected)` | 检查客户端是否已连接 |

### 事件回调（1 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| SetEventCallback | `(callback)` | 设置事件回调对象（IDispatch），触发 OnAccept/OnConnect/OnReceive/OnClose 等事件 |

### 事件列表
| 事件 | 触发时机 | 参数 |
|------|----------|------|
| OnAccept | 服务器接受新客户端连接 | connId, address |
| OnConnect | 客户端连接成功 | connId |
| OnReceive | 收到数据 | connId, data, dataSize |
| OnClose | 连接断开 | connId, errorCode |
| OnSend | 数据发送完成 | connId, dataSize |

### 依赖
- **HP-Socket**：静态编译链接（COM_Lib/third_party/HP-Socket/，Apache 2.0 协议）

---

## 27. ITcpOperationEx — TCP Pack 扩展操作（12 方法）

基于 HP-Socket TcpPackServer/TcpPackClient，按 4 字节包头自动拆包，业务层直接按完整消息收发。

### TCP Pack Server（7 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| ServerStartEx | `(bindAddr, port, maxPackSize, packHeaderFlag)` | 启动 TCP Pack 服务器 |
| ServerStopEx | `()` | 停止 TCP Pack 服务器 |
| ServerSendMessage | `(connId, data, dataSize, &sent)` | 向指定连接发送完整业务消息 |
| ServerBroadcastMessage | `(data, dataSize)` | 向所有连接广播完整业务消息 |
| ServerDisconnectEx | `(connId, force)` | 断开指定连接 |
| ServerGetConnectionCountEx | `(&count)` | 获取当前连接数量 |
| ServerGetAllConnectionsEx | `(&connIds, &count)` | 获取全部连接 ID |

### TCP Pack Client（4 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| ClientConnectEx | `(remoteAddr, port, async, maxPackSize, packHeaderFlag)` | 连接 TCP Pack 服务器 |
| ClientSendMessage | `(data, dataSize, &sent)` | 客户端发送完整业务消息 |
| ClientCloseEx | `()` | 关闭 TCP Pack 客户端 |
| ClientIsConnectedEx | `(&connected)` | 检查客户端连接状态 |

### 事件回调（1 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| SetEventCallbackEx | `(callback)` | 设置事件回调对象 |

### 事件列表
| 事件 | 触发时机 | 参数 |
|------|----------|------|
| OnAccept | 服务器接受新客户端连接 | connId, address |
| OnConnect | 客户端连接成功 | connId |
| OnReceive | 收到完整消息 | connId, data, dataSize |
| OnClose | 连接断开 | connId, errorCode |

### 依赖
- **HP-Socket**：TcpPackServer/TcpPackClient 静态编译链接

---

## 28. IHttpOperationEx — HTTP 扩展操作（7 方法）

基于 HP-Socket HttpServer/HttpClient，提供 HTTP 服务端请求响应和客户端请求能力。

### HTTP Server（3 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| HttpServerStart | `(bindAddr, port)` | 启动 HTTP 服务器 |
| HttpServerStop | `()` | 停止 HTTP 服务器 |
| HttpServerSendResponse | `(connId, statusCode, desc, body, bodySize)` | 向客户端发送 HTTP 响应 |

### HTTP Client（3 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| HttpClientConnect | `(remoteAddr, port, async)` | 连接 HTTP 服务器 |
| HttpClientSendRequest | `(method, path, body, bodySize)` | 发送 HTTP 请求 |
| HttpClientClose | `()` | 关闭 HTTP 客户端 |

### 事件回调（1 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| SetEventCallbackEx | `(callback)` | 设置事件回调对象 |

### 事件列表
| 事件 | 触发时机 | 参数 |
|------|----------|------|
| OnRequestLine | 服务器收到请求行 | connId, method, url |
| OnStatusLine | 客户端收到状态行 | connId, statusCode, desc |
| OnBody | 收到请求/响应体 | connId, data, dataSize |
| OnMessageComplete | 请求/响应完成 | connId |

### 依赖
- **HP-Socket**：HttpServer/HttpClient 静态编译链接

---

## 29. IWebSocketOperationEx — WebSocket 扩展操作（10 方法）

基于 HP-Socket HTTP + WebSocket Upgrade 能力，提供 WebSocket 文本/二进制消息收发。

### WebSocket Server（5 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| WsServerStart | `(bindAddr, port)` | 启动 WebSocket 服务器 |
| WsServerStop | `()` | 停止 WebSocket 服务器 |
| WsServerSendText | `(connId, text)` | 服务端发送文本消息 |
| WsServerSendBinary | `(connId, data, dataSize)` | 服务端发送二进制消息 |
| WsServerBroadcastText | `(text)` | 服务端广播文本消息 |

### WebSocket Client（4 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| WsClientConnect | `(remoteAddr, port, path, async)` | 客户端连接 WebSocket 服务端 |
| WsClientSendText | `(text)` | 客户端发送文本消息 |
| WsClientSendBinary | `(data, dataSize)` | 客户端发送二进制消息 |
| WsClientClose | `()` | 关闭 WebSocket 客户端 |

### 事件回调（1 方法）
| 方法 | 参数 | 说明 |
|------|------|------|
| SetEventCallbackEx | `(callback)` | 设置事件回调对象 |

### 事件列表
| 事件 | 触发时机 | 参数 |
|------|----------|------|
| OnUpgrade | WebSocket 握手完成 | connId |
| OnWSMessage | 收到 WebSocket 消息 | connId, data, dataSize |
| OnClose | 连接断开 | connId, errorCode |

### 依赖

---

## 30. ISkinOperation — 皮肤操作（6 方法）

封装 SKIN++ 皮肤引擎，DLL 以 RCDATA 嵌入 COM_Lib.dll，通过 PeLoader 内存加载。

### 核心方法
| 方法 | 参数 | 说明 |
|------|------|------|
| SetSkin | `(skinData, dataSize, password, hue, sat, bri)` | 从内存数据加载皮肤 |
| SetSkinByIndex | `(index, password, hue, sat, bri)` | 按索引加载嵌入的皮肤资源（0~118） |
| UnsetSkin | `()` | 卸载当前皮肤，恢复原始样式 |
| SetTransparent | `(hWnd, alpha)` | 设置窗口透明度（0=全透明，255=不透明） |
| SetAero | `(enable)` | 切换 Aero 玻璃效果 |
| SetTitleMenu | `(hWnd, enable, menuHeight, top, right)` | 设置标题菜单栏 |

### 参数说明
- `skinData`：.she 皮肤文件内容（字节数组）
- `password`：皮肤密码（无密码传空字符串）
- `hue/saturation/brightness`：颜色调整参数（0=不变）
- `index`：皮肤索引 0~118（对应 119 个内置 .she 文件）

### 资源嵌入
- 32位 Skin.dll → RCDATA `SKIN_DLL_X86`
- 64位 Skin.dll → RCDATA `SKIN_DLL_X64`
- 119 个 .she 皮肤文件打包 → RCDATA `SKINS_DATA`

### 依赖
- **Peloader**：内存 PE 加载器（`peloader.h`）
- **Skin.dll**：SKIN++ 皮肤引擎（无源码，仅二进制嵌入）

---

## 31. IWindowAutoSize — 窗口自适应（3 方法）

根据火山代码「类_窗口自适应」移植，提供 11 种约束类型的窗口自适应尺寸处理。

### 核心方法
| 方法 | 参数 | 说明 |
|------|------|------|
| Initialize | `(hMainWnd, nWidth, nHeight, pNoChange, pLeftTop, pWidthHeight, pOnlyLeft, pOnlyTop, pOnlyWidth, pOnlyHeight, pLeftWidth, pLeftHeight, pTopWidth, pTopHeight)` | 初始化窗口自适应，设置主窗口和 11 种约束数组 |
| SizeChanged | `()` | 窗口尺寸改变时调用，按约束类型自适应调整所有子控件 |
| GetControlCount | `(&count)` | 获取当前管理的控件数量 |

### 约束类型说明
| 类型 | 编码 | 说明 |
|------|------|------|
| 默认（无约束） | 0 | left/top/width/height 全部按比例 |
| 禁止变化 | 1 | left/top 按比例，width/height 固定 |
| 左边顶边 | 2 | left/top 按比例，width/height 固定 |
| 宽度高度 | 3 | left/top 按比例，width/height 固定 |
| 仅左边 | 4 | 同默认效果 |
| 仅顶边 | 5 | 同默认效果 |
| 仅宽度 | 6 | left/top/height 按比例，width 固定 |
| 仅高度 | 7 | left/top/width 按比例，height 固定 |
| 左边宽度 | 8 | left/top/height 按比例，width 固定 |
| 左边高度 | 9 | left/top/width 按比例，height 固定 |
| 顶边宽度 | 10 | left/top/height 按比例，width 固定 |
| 顶边高度 | 11 | left/top/width 按比例，height 固定 |

### ILeyBc 转发方法
| 方法 | 说明 |
|------|------|
| AutoSize_Initialize | 转发至 IWindowAutoSize.Initialize |
| AutoSize_SizeChanged | 转发至 IWindowAutoSize.SizeChanged |
| AutoSize_GetControlCount | 转发至 IWindowAutoSize.GetControlCount |

---

## 32. IGuiOperation — GUI 操作（约 240 方法）

封装 Dear ImGui v1.91.8 引擎，提供纯软件光栅化的即时模式 GUI 渲染能力。所有方法均含 SEH 异常保护 + CRITICAL_SECTION 线程安全。

### 生命周期与渲染
| 方法 | 说明 |
|------|------|
| InitializeCu/Initialize | 初始化 ImGui 上下文 |
| SetOutputSize | 设置输出尺寸（宽高） |
| BeginFrame/EndFrame/Render | 帧生命周期控制 |
| GetOutputBuffer | 获取渲染输出像素缓冲区（调用者 CoTaskMemFree） |
| SetMousePos/SetMouseButton/SetMouseWheel | 输入状态设置 |

### 布局与光标（23 方法）
Spacing/Dummy/Indent/Unindent/Group/Separator/SameLine/NewLine/NextColumn/Columns、SetCursorPos/SetCursorPosX/SetCursorPosY/GetCursorPos/GetCursorScreenPos、PushItemWidth/PopItemWidth/SetNextItemWidth/SetColumnOffset/SetColumnWidth、GetContentRegionAvail/GetTextLineHeight/GetFrameHeight

### 窗口操作（19 方法）
SetNextWindowPos/Size/SizeConstraints/Collapsed/Focus/Scroll/ScrollbarSize、SetWindowPos/Size/Collapsed/Focus/Scroll/FontScale、GetWindowPos/Size/Width/Height/IsWindowFocused/IsWindowHovered

### 滚动（10 方法）
GetScrollX/Y、SetScrollX/Y、GetScrollMaxX/Y、SetScrollHereX/Y、SetScrollFromPosX/Y

### ID栈与焦点（5 方法）
PushID/PopID/PushOverrideID/SetItemDefaultFocus/FocusItem

### 文本（7 方法）
TextColored/TextDisabled/TextWrapped/LabelText/BulletText/SeparatorText/Bullet

### 基础控件（32 方法）
SmallButton/ArrowButton/Checkbox/CheckboxFlags/RadioButton/Selectable、ColorEdit3/ColorEdit4/ColorPicker3/ColorPicker4/ColorButton、SliderFloat/Int/Float2/Int2/Float3/Int3/Float4/Int4、SliderAngle、DragFloat/Int/Float2/Int2/Float3/Int3/Float4/Int4、InputFloat/Int/Double/Float2/Int2/Float3/Int3/Float4/Int4、VSliderFloat/VSliderInt

### 列表与组合框（5 方法）
BeginListBox/EndListBox/ListBox/BeginCombo/EndCombo

### 树形（4 方法）
CollapsingHeader/TreeNodeEx/TreePush/TreePop/SetNextItemOpen/GetTreeNodeToLabelSpacing

### 表格（18 方法）
BeginTable/EndTable/TableNextRow/TableNextColumn/TableSetColumnIndex/TableSetupColumn/TableSetupScrollFreeze/TableHeadersRow/TableHeader/TableSetBgColor/TableGetColumnCount/TableGetColumnIndex/TableGetRowIndex/TableGetColumnName/TableGetColumnFlags/TableSetColumnEnabled/TableSetColumnWidth/TableSetColumnSortDirection

### 标签页（6 方法）
BeginTabBar/EndTabBar/GetTabBarHeight/IsTabBarHovered/BeginTabItem/EndTabItem/SetTabItemClosed/TabItemButton

### 菜单（2 方法）
BeginMainMenuBar/EndMainMenuBar、BeginMenuBar/EndMenuBar/BeginMenu/EndMenu/MenuItem

### 状态查询（14 方法）
IsItemHovered/Active/Focused/Clicked/Visible/Edited/Activated/Deactivated、IsAnyItemHovered/Active/Focused/GetItemRectMin/Max/Size/ID/IsRectVisible

### 键盘鼠标（17 方法）
IsKeyDown/Pressed/Released、GetKeyPressedAmount/GetKeyName/SetKeyOwner/TestKeyOwner、IsMouseDown/Clicked/Released/DoubleClicked/Dragging/GetMousePos/GetMousePosOnOpeningCurrentPopup/GetMouseDragDelta/SetMouseCursor/GetFrameCount/GetTime

### 工具提示与弹窗（14 方法）
BeginTooltip/EndTooltip/SetTooltip、OpenPopup/OpenPopupOnItemClick/CloseCurrentPopup/BeginPopup/EndPopup/BeginPopupContextItem/ContextWindow/ContextVoid/PopupModal/IsPopupOpen/GetPopupScrollLeft/Top

### 禁用裁剪文本工具（8 方法）
BeginDisabled/EndDisabled/PushClipRect/PopClipRect/CalcTextSize/GetFontSize/GetColorU32/GetStyleColorVec4

### 调试样式（11 方法）
ShowDemoWindow/ShowMetricsWindow/ShowAboutWindow/ShowDebugLogWindow/ShowStyleEditor/StyleColorsDark/StyleColorsLight/StyleColorsClassic/GetVersion、PushStyleColor/PopStyleColor/PushStyleVar/PopStyleVar

### ILeyBc 转发方法
| 方法 | 说明 |
|------|------|
| Gui_BeginFrame | 转发至 IGuiOperation.BeginFrame |
| Gui_BeginWindow | 转发至 IGuiOperation.BeginWindow |
| Gui_Text | 转发至 IGuiOperation.Text |
| Gui_Button | 转发至 IGuiOperation.Button |
| ...（所有 IGuiOperation 方法均通过 Gui_ 前缀转发） | |

### 依赖
- **Dear ImGui v1.91.8**：源码编译（imgui.cpp/draw/widgets/tables）
- **纯软件光栅化**：Edge function 算法 + 浮点精度 + BGRA 字节序
- **系统字体**：Win32 API 读取 msyh.ttc/simsun.ttc，全字库中文支持

---

## ILeyBc — 统一管理接口

`ILeyBc` 是项目推荐使用的统一接口，一个 CLSID（`CLSID_CLeyBcManager`）管理全部模块和方法。

**方法命名差异**：ILeyBc 中的方法名去掉了匈牙利前缀或做了重命名，例如：
- `ITextOperation::GetPinyin` → `GetPinyin`
- `ICryptoOperation::MD5Hash` → `ComputeMD5`
- `ICryptoOperation::MD5HashString` → `ComputeMD5String`
- `ICryptoOperation::SHA1Hash` → `ComputeSHA1`
- `ICryptoOperation::SHA256Hash` → `ComputeSHA256`
- `ICryptoOperation::MD2Hash` → `ComputeMD2`
- `ICryptoOperation::MD4Hash` → `ComputeMD4`
- `ICryptoOperation::SHA224Hash` → `ComputeSHA224`
- `ICryptoOperation::SHA384Hash` → `ComputeSHA384`
- `ICryptoOperation::SHA512Hash` → `ComputeSHA512`
- `ICryptoOperation::SHA3Hash` → `ComputeSHA3`
- `ICryptoOperation::CRC32Hash` → `ComputeCRC32`
- `ICryptoOperation::HMAC_MD5` → `HMACMD5`
- `ICryptoOperation::HMAC_SHA1` → `HMACSHA1`
- `ICryptoOperation::HMAC_SHA224` → `HMACSHA224`
- `ICryptoOperation::HMAC_SHA256` → `HMACSHA256`
- `ICryptoOperation::HMAC_SHA384` → `HMACSHA384`
- `ICryptoOperation::HMAC_SHA512` → `HMACSHA512`
- 其余方法名保持一致

---

## IDispatch 脚本语言支持（v1.7.0）

当前生成的 736 个 `ILeyBc` 分发入口均支持通过 IDispatch 从脚本语言调用，无需手动构造 DISPPARAMS。

### 调用方式

| 语言 | 示例 |
|------|------|
| **C# (dynamic)** | `string name = comObj.GetComponentName();` |
| **VBScript** | `name = obj.GetComponentName()` |
| **JScript** | `var name = obj.GetComponentName();` |
| **PowerShell** | `$name = $comObj.GetComponentName()` |
| **Python (win32com)** | `name = comObj.GetComponentName()` |

### 实现原理

- 使用原生 `IDispatch` 包装器（`CLeyBcDispatchWrapper`）替代 `CreateStdDispatch`
- 每个方法对应一个独立的 `Invoke_XXX` 包装函数，由编译器自动处理类型转换
- 输出参数（`[out,retval]`）直接作为方法返回值返回
- 输入参数通过 DISPPARAMS 逆序传入，支持类型自动转换（`VariantChangeType`）

### 历史验证结果（v1.7.0，623 个方法）

| 测试项 | 结果 |
|--------|------|
| IDispatch::GetIDsOfNames | ✅ 623 方法名均可查找到 |
| IDispatch::Invoke 单参数 | ✅ GetLength("hello") → 5 |
| IDispatch::Invoke 多参数 | ✅ FindText("hello world","world",0) → 6 |
| IDispatch::Invoke 字符串输出 | ✅ GetComponentName() → "LeyBc" |
| C# 全方法测试 | ✅ 623/623 全部通过 |

---

## 版本记录

| 版本 | 日期 | 说明 |
|------|------|------|
| v1.31.0 | 2026-08-07 | IGuiOperation 接口完整扩展：Dear ImGui v1.91.8 全部 API 封装（52→240 方法），含表格/标签页/列表/颜色/拖拽/树形/调试样式等，纯软件光栅化，全平台编译通过 |

*最新版本文档：v1.31.0*
*最后更新：2026-08-07*