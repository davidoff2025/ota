# OTA 部署文件生成工具

这是一个基于 Python Tkinter + openpyxl 的 OTA 升级 Excel 导入模板生成工具。

## 主要功能

- 选择 Source Excel 文件并自动识别数据工作表。
- 自动识别源文件字段，包括机型、扩展码、OTA 类型、源版本、目标版本、版本变化、升级包地址、MD5、文件大小、品牌、区域、SHA256、EULA、MAC 等字段。
- 新生成 OTA 导入 Excel 文件，不依赖外部模板文件。
- 从 Source Excel 的“区域”字段复制内容到新生成 Excel 的“区域组”列。
- Source Excel 中的 MAC 字段不会被复制；输出文件的“MAC组”只使用 UI 中输入的 MAC 地址组。
- 根据设备类型代码自动生成 AWS 下载路径：`https://ota-dl.vidaahub.com/ota/{设备类型代码}/{YYYYMMDD}/`
- 根据升级包地址提取 ZIP 文件名，并生成最终 OTA 下载 URL。
- 支持从 FTP 下载 ZIP 文件到本地 `temp` 目录。
- 支持通过 AWS CLI 将 ZIP 文件上传到 S3：`s3://fam-media-andr/ota/{设备类型代码}/{YYYYMMDD}/`
- 支持“一键自动执行”：生成模板 -> 下载 ZIP -> 上传 S3。
- UI 日志窗口会详细打印每一步执行过程。

## 本次更新内容

### 1. 区域字段映射修复

新增 Source Excel 字段识别：

- `区域`
- `区域组`
- `地区`
- `地区组`
- `region`
- `area`
- `country`
- `market`
- `territory`

如果 Source Excel 中识别到“区域”字段，其内容会逐行复制到输出 Excel 的：

```text
区域组
```

如果未识别到该字段，输出文件的“区域组”会留空，并在日志中提示。

### 2. UI 执行日志增强

日志窗口现在会输出更详细的执行信息，包括：

- 操作开始时间
- Source Excel 路径
- Source Excel 文件大小
- 工作表数量和工作表名称
- 每个工作表的行列尺寸
- 字段映射结果
- 关键字段是否识别成功
- 是否检测到“区域”列
- 是否检测到“版本变化”列
- 是否检测到 Source Excel 的 MAC 列，以及不会复制该列的说明
- 每一行生成的关键字段
  - 内部机型信息
  - 设备扩展信息
  - OTA 类型转换结果
  - 源版本
  - 目标完整版本
  - 目标版本
  - 品牌组
  - 区域组
  - MAC组
  - Source 升级包地址
  - ZIP 文件名
  - 最终升级文件地址
  - MD5
  - 文件大小
  - SHA256
- FTP 服务器地址
- FTP 登录用户
- FTP 连接状态
- FTP 登录是否成功
- 远程文件路径
- 本地暂存路径
- 下载进度
- AWS CLI 路径和版本
- AWS 当前身份
- S3 目标路径
- 每个 ZIP 文件上传结果

## 安装依赖

```bash
pip install openpyxl
```

macOS 如果缺少 Tkinter，可安装：

```bash
brew install python-tk
```

## 运行方式

```bash
python appexcelftpv1_updated.py
```

或：

```bash
python3 appexcelftpv1_updated.py
```

## 使用步骤

1. 打开程序。
2. 选择 Source Excel 文件。
3. 输入设备类型代码。
4. 输入设备类型名称。
5. 输入 Feature Code / 特征码。
6. 输入 MAC 地址组。
7. 查看 AWS 文件路径是否正确。
8. 根据需要点击：
   - 下载ZIP文件
   - 上传ZIP文件到AWS
   - 生成模版文件
   - 自动执行

## 字段说明

### Source Excel 关键字段

| 目标字段 | 可识别别名 |
|---|---|
| 机型信息 | 机型、机型信息、内部机型、model、model name |
| 机器扩展码 | 机器扩展码、扩展码、扩展信息、设备扩展信息、extension info、extension code、feature code |
| OTA类型 | OTA类型、ota type、升级类型 |
| 源版本 | 源版本、source version、源文件版本 |
| 目标版本 | 目标版本、target version、完整版本、目标文件版本 |
| 版本变化 | 版本变化、version change、版本差异 |
| 升级包地址 | 升级包地址、升级文件地址、ftp path、upgrade url、file path、文件路径 |
| 升级文件的MD5值 | 升级文件md5值、md5、md5值、md5 checksum |
| 文件大小 | 升级包大小、升级包大小(byte)、升级文件包大小（byte）、file size、size |
| 品牌 | 品牌、brand |
| 区域 | 区域、区域组、地区、地区组、region、area、country、market、territory |
| SHA256 | sha256、sha256值 |
| EULA文件地址 | eula文件地址、eula、eula url |
| MAC | mac、mac地址、mac address |

### 输出 Excel 说明

输出 Excel 中包含以下列：

```text
设备类型代码
设备类型名称
特征码
内部机型信息
设备扩展信息
OTA类型
源版本
目标完整版本
目标版本
品牌组
区域组
MAC组
定向组
升级文件地址
升级文件的MD5值
文件大小
SHA256
EULA文件地址
```

其中：

- “区域组”来自 Source Excel 的“区域”字段。
- “MAC组”只来自 UI 输入，不复制 Source Excel 的 MAC 字段。
- “目标版本”优先使用 Source Excel 的“版本变化”字段；如果没有，则由源版本后 5 位和目标版本后 5 位自动生成。
- “升级文件地址”由 AWS 前缀、设备类型代码、日期和 ZIP 文件名拼接生成。

## AWS / S3 要求

如果需要上传 S3，请确保本机已经安装并配置 AWS CLI：

```bash
aws --version
aws sts get-caller-identity
```

如果未配置，请执行：

```bash
aws configure
```

## 日志复制

程序 UI 提供：

```text
复制日志 / Copy Log
```

按钮，可以将完整执行日志复制到剪贴板，便于排查 FTP、Excel、S3 或字段映射问题。



## 新增功能：升级提示语 Sheet

生成 OTA 模版文件时，将自动创建第二个工作表：

```text
升级提示语
```

包含以下字段：

| 语种 | 描述 |
|------|------|
| 简体中文 | OTA 升级提示语 |
| 繁体中文 | OTA 升级提示语 |
| 法语 | OTA 升级提示语 |
| 西班牙语 | OTA 升级提示语 |
| 德语 | OTA 升级提示语 |
| 葡萄牙语 | OTA 升级提示语 |

该工作表用于维护电视端 OTA 升级弹窗提示信息，并会随模板文件自动生成。
