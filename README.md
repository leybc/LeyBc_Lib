# LeyBc COM Library

纯静态编译（/MT）的 C++ COM DLL，零第三方运行时依赖（XEDParse / FASM / Skin 内存加载，ChakraCore 静态链接）。对外只暴露 **一个** COM 对象 **`LeyBc.Object`**，共 **1196 个方法**（IDispatch + dual 类型库），易语言 / VBS / 脚本可直接枚举调用。

## 📖 在线 API 文档

👉 [**完整 API 接口报告（网页版）**](https://leybc.github.io/LeyBc_Lib/docs/api_report.html)

涵盖 38 个接口 / 37 个组件 / 641+ 方法 / 1000+ 功能点，可视化查看全部方法签名与参数。

- [FUNCTION.md 全量功能说明书](docs/FUNCTION.md)
- [ILeyBc_Methods_List.txt 1196 方法清单（DISPID + 方法名）](docs/ILeyBc_Methods_List.txt)
- [api_report.md API 报告（Markdown）](docs/api_report.md)

## ⚡ 快速开始

```bash
# 构建（x86 / x64 Release）
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

## 📋 基本信息

- **ProgID**：`LeyBc.Object`
- **CLSID**：`{1A2B3C4D-5E6F-7890-ABCD-EF1234567890}`
- **方法总数**：1196（DISPID 1–1196 连续）
- **编译方式**：MSVC /MT 纯静态，仅依赖 Windows 系统 DLL

## 📂 仓库结构

```
LeyBc_Lib/
├── README.md                    # 本文件
└── docs/
    ├── api_report.html          # 完整 API 接口报告（网页版，推荐）
    ├── api_report.md            # API 报告 Markdown
    ├── FUNCTION.md              # 全量功能说明书
    └── ILeyBc_Methods_List.txt  # 1196 方法清单
```

## ℹ️ 说明

- 依赖的 `third_party/ChakraCore`、`HP-Socket`、`openssl` 等源码体积较大，未纳入本仓库；构建前需自行放置（或按 `CMakeLists.txt` 引用路径补全）。
- 本仓库仅发布成品与文档；完整源码保持私有。