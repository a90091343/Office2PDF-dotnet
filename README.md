# Office2PDF-dotnet
![office2pdf_v3](https://github.com/evgo2017/Office2PDF-dotnet/raw/master/assets/office2pdf_v3.png)

使用 .net 重新实现 [Office2PDF](https://github.com/evgo2017/Office2PDF) 版本。其它更新、更多信息在[软件主页](https://evgo2017.com/blog/office2pdf)查看。

## 一、下载使用

下载地址 1：[网盘](http://evgo2017.ysepan.com/)（项目密码: evgo2017）（两小时内有下载量的限制，故设置密码）

下载地址 2：[Github Release](https://github.com/evgo2017/Office2PDF-dotnet/releases)

## 二、详细说明

### 1. 转换细节

|            | Word | Excel                                | PPT               |
| ---------- | ---- | ------------------------------------ | ----------------- |
| 文档有内容 | ✅    | ✅（若有多个 Sheet，则生成多个文件）  | ✅（多页）         |
| 文档无内容 | ✅    | ❌（会跳过，不会产生对应的 PDF 文件） | ❌（提示错误跳过转换） |

### 2. 运行要求

#### Office 2007 及以后版本

- [x] 已安装 Office

#### Office 2007 以前版本

- [x] 已安装 Office
- [x] Microsoft Save as PDF  加载项

### 3. 如何使用

① 电脑安装 [.net 9 运行时](https://aka.ms/dotnet-core-applaunch?missing_runtime=true&arch=x64&rid=win-x64&os=win10&apphost_version=9.0.2&gui=true)（点击此链接，会自动打开微软官网下载，安装即可）。

② 运行 `Office2PDF.exe` 即可使用。

## 四、版本更新记录

| 时间       | 内容                                                         | 相关文章                                                     |
| ---------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| 2025.05.29 | v3，开源                                           | [开源地址](https://github.com/evgo2017/Office2PDF-dotnet) |
| 2025.03.31 | v3，使用 .net 重构 后发布                                           | [Office2PDF：Office 批量转为 PDF（v3.0）](https://mp.weixin.qq.com/s/ZKoeyOjXNUtyG8c7GyQc3A) |
| 2020.08.26 | v2，加入 GUI，支持选择类型、子文件夹等功能                   | [Office2PDF 批量转 PDF（第二版）](https://mp.weixin.qq.com/s/VxHxvUUqK2tn0PKNQkXTsQ) |
| 2019.05.13 | 将此项目从自己的 `SomeTools` 项目独立出来，通过 `release` 发布 `exe` |                                                              |
| 2018.11.02 | v1，功能基本实现                                             | [office 转 pdf 技巧及软件](https://mp.weixin.qq.com/s/jZvVXgqcMOIxkKVzJXYEZA) |
