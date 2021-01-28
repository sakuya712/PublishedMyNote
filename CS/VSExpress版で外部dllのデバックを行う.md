---
title: VSExpress版で外部dllのデバックを行う
date: 2020-09-24
tags:
    - CSharp
    - VS
---

目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [VSExpress版で外部dllのデバックを行う](#vsexpress版で外部dllのデバックを行う)

<!-- /code_chunk_output -->

# VSExpress版で外部dllのデバックを行う

1. C#のプロジェクトファイル（.csproj）をテキストエディタで開く
2. 以下の箇所を検索する
```xml
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
```
3. この後ろの行に以下の文字を追加する
```xml
    <StartAction>Program</StartAction>
    <StartProgram>外部アプリケーションのパス</StartProgram>
```
<例> Excelのアドインのデバッグをしたい場合
```xml
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
    <StartAction>Program</StartAction>
    <StartProgram>C:\Program Files\Microsoft Office\Office14\EXCEL.EXE</StartProgram>
    <DebugSymbols>true</DebugSymbols>
    <DebugType>full</DebugType>
    <Optimize>false</Optimize>
    <OutputPath>bin\Debug\</OutputPath>
    <DefineConstants>DEBUG;TRACE</DefineConstants>
    <ErrorReport>prompt</ErrorReport>
    <WarningLevel>4</WarningLevel>
  </PropertyGroup>
  ```