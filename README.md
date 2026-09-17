# LeyBc COM Library

纯静态编译（/MT）的 C++ COM DLL，零第三方运行时依赖（XEDParse/FASM/Skin 内存加载、ChakraCore 静态链接），对外只暴露一个 COM 对象 **`LeyBc.Object`**（1196 方法，IDispatch + dual 类型库，易语言/脚本可直接枚举调用）。

## 快速开始

```bash
# 构建（x86/x64 Release）
cmake -B build/x86-Release -A Win32  -DCMAKE_BUILD_TYPE=Release
cmake --build build/x86-Release --config Release --target COM_Lib
cmake -B build/x64-Release -A x64    -DCMAKE_BUILD_TYPE=Release
cmake --build build/x64-Release --config Release --target COM_Lib

# 注册（位数必须匹配）
C:\Windows\SysWOW64\regsvr32.exe COM_Lib.dll   # 32 位（易语言）
C:\Windows\System32\regsvr32.exe  COM_Lib.dll   # 64 位

# 使用
CreateObject("LeyBc.Object")   # VBS / 易语言
```

## 文档

- `COM_Lib/docs/FUNCTION.md` — 全量功能说明书
- `COM_Lib/docs/ILeyBc_Methods_List.txt` — 1196 方法清单（DISPID + 方法名）
- `AGENTS.md` — 项目开发指令与架构沉淀（含 COM 注册坑、类型库、内存加载等关键约束）

## 说明

- 依赖的 `third_party/ChakraCore`、`HP-Socket`、`openssl` 等源码体积较大，未纳入本仓库；构建前需自行放置（或按 `CMakeLists.txt` 引用路径补全）。
- `build/`、`backup/`、`test/dist/` 等构建/备份产物不入库。