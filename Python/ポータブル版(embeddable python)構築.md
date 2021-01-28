---
title: ポータブル版(embeddable python)構築
date: 2020-07-29
tags:
    - python
---

目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [ポータブル版(embeddable python)構築](#ポータブル版embeddable-python構築)
  - [1. ポータブル版PyhtonをDLする](#1-ポータブル版pyhtonをdlする)
  - [２．pythonXX._pthを修正する](#２pythonxx_pthを修正する)
  - [3. pipをインストールする](#3-pipをインストールする)

<!-- /code_chunk_output -->

# ポータブル版(embeddable python)構築

## 1. ポータブル版PyhtonをDLする

https://www.python.org/downloads/windows/  
上記ページにある`Windows x86-64 embeddable zip file`をダウンロードして解凍。

## ２．pythonXX._pthを修正する
解凍したフォルダに`python38._pth`というファイルが存在している。(バージョンによって数字が異なる)  
このファイルをテキストエディタなどで開いて
```
import site
```
のコメントアウトを外す

## 3. pipをインストールする
この時点ではpipすらないので  
https://bootstrap.pypa.io/get-pip.py  
からダウンロードする。ダウンロードしたファイルをPythonと同じフォルダに置く

インストール
cmdでPythonがあるフォルダまで移動する。そこで
```
python get-pip.py
```
と入力するとpipがインストールされる。

またpipを使用する場合もPythonがあるフォルダまで移動して  
`python -m pip install numpy`のように`-m`のオプションを使用すること。  
これをしないとPathを通してあるPythonの方にインストールされてしまう。  

ちなみにembeddableはライブラリがすっからかんなのでpipで失敗することがわりとある。その場合は、  
`pip download -d ダウンロード先 --no-binary :all: パッケージ`  
これで必要なパッケージも全てダウンロードしたことになる。そして
`python -m pip install --no-deps パッケージ`  
でインストールすると上手くいくかもしれない

どうしてもだめならPathを通しているPythonにインストールさせて、追加されたファイルをembeddable pythonに入れる。ただこれは不具合が起こってもおかしくないので最終手段
