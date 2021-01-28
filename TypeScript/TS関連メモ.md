---
title: TS関連メモ
date: 2020-08-04
tags:
    - 
---

目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [TS関連メモ](#ts関連メモ)
  - [型宣言](#型宣言)
  - [共用体型(型宣言のとき｜がある)](#共用体型型宣言のときがある)
  - [Stringとstringは違う](#stringとstringは違う)
  - [正規表現](#正規表現)
  - [jsonファイルを読み込む](#jsonファイルを読み込む)
  - [文字列操作あれこれ](#文字列操作あれこれ)
    - [特定の文字を消す](#特定の文字を消す)
    - [特定の文字以降を消す](#特定の文字以降を消す)
    - [文字列に特定の文字が含んでいるかどうかを調べる(真偽値を返す)](#文字列に特定の文字が含んでいるかどうかを調べる真偽値を返す)
    - [文字がどこにあるか調べる](#文字がどこにあるか調べる)
  - [配列操作あれこれ](#配列操作あれこれ)
    - [配列の長さを知る](#配列の長さを知る)
    - [配列に要素が入っているか確認する(真偽値を返す)](#配列に要素が入っているか確認する真偽値を返す)
    - [配列に要素が入っているインデックス番号を返す](#配列に要素が入っているインデックス番号を返す)
    - [配列の要素が全て一致しているか確認する(真偽値を返す)](#配列の要素が全て一致しているか確認する真偽値を返す)
    - [連想配列が入った配列から特定の値が入った連想配列を抜き出す](#連想配列が入った配列から特定の値が入った連想配列を抜き出す)

<!-- /code_chunk_output -->

# TS関連メモ

## 型宣言
TSでは  
`const 変数名: 型名 = 値`のように宣言できる

`const`は再宣言、再代入が禁止されているもの  
`let`は再宣言は禁止されているが、再代入は可能  
`var`は再宣言、再代入どちらも可能となっている。  
TSでは`var`は極力使用しない。

```TS
//function 関数名 (引数名:引数の型) : 関数の戻り値の型 {}

function hoge(str:string):string{
    return '私の名前は'+ str + 'です。';
}

```


## 共用体型(型宣言のとき｜がある)

TSには共用体型がある。  
これは複数の型が入る可能性があるものに宣言する。  
例えば以下のように書くと paddingがStringかnumberかの型ということである。
```TS
function padLeft(value: string, padding: string | number) {
    // ...
}
```
C#でもあるヌル許容型としたいなら
```TS
// 左側の型が本来いれる予定の型、何らかの処理でエラーをはいたときnullも入れれるようにしておく
let start: vscode.Position | null = null;
```
のようにかける

## Stringとstringは違う

stringは昔からあるプリミティブ型の文字列  
Stringはオブジェクト型の文字列、そのためnewステートメントも必要  
TSではstringが推奨なのでStringは使わない。しかし他の言語の癖で大文字にしがちなので、注意したい。

## 正規表現

例文1
```TS
var str = "Dim testvar as string"

var buf = str.split(/As/i);
var varName = buf[0].replace(/Dim|\s/g,"")
var typeName = buf[1].trim()
console.log(varName);           //testvar
console.log(typeName);          //string
```
この`/As/i`のiは大文字と小文字は区別しないという意味。そのためDim testvarとstringで分かれる。  
`/Dim|\s/g,""`の`/Dim|\s`はDimかスペースかという意味。但しTSのreplaceは最初の一個しか置換してくないので、どちらとも置換してほしい場合、`g`をつけて文字列全体に対してマッチングさせる。

## jsonファイルを読み込む

現在のTSでは`tsconfig.json`を
```json
{
  "compilerOptions": {
    "moduleResolution": "node",
    "resolveJsonModule": true
  }
}
```
に設定すれば.tsファイルと同じような感覚でインポートできる  
```TS
import dummy from './dummy.json'

console.log(dummy.foo); 
```

## 文字列操作あれこれ

### 特定の文字を消す
replaceを使用する。  
他の言語でもよくある置換を利用したもの
```TS
const str = '夕食はラーメンとカレーです';
const result = str.replace('とカレー', '');
console.log(result); // => 夕食はラーメンです
```

### 特定の文字以降を消す
splitを使用する。  
本来は区切り文字で分割するためのものだが、特定の文字がで分割する場合でも使える。
```TS
const str = '2020年8月10日'
const result = str.split('年')[0];
console.log(result); // -> 2020
//逆に'年'以降が欲しい場合は
const str = '2020年8月10日';
const result = str.split('年')[1];
console.log(result); // => 8月10日
```

### 文字列に特定の文字が含んでいるかどうかを調べる(真偽値を返す)
正規表現が不要な場合はincludesを使用する。
```TS
const str = 'foobarbaz';
const result = str.includes('bar');
console.log(result); // => true
```
正規表現を使うときはRegExpのtestを使う
```TS
const str = 'foobarbaz'
const reg = /\wbar\w/i
const result = reg.test(str)
console.log(result); // => true
```

### 文字がどこにあるか調べる
正規表現を使わないとき => indexOf
正規表現を使うとき => search
```TS
const str = 'foobarbaz';
const result = str.indexOf('bar');
console.log(result); // => 3
```

## 配列操作あれこれ

### 配列の長さを知る
lengthを使う。  
**最大のインデックス番号を返しているわけではない**ので注意
```TS
let List = ['foo','bar','baz'];
let result = List.length
console.log(result); // => 3
```

### 配列に要素が入っているか確認する(真偽値を返す)
someを使う
```TS
let List = ['foo','bar','baz'];
let result = List.some(x => x === 'bar');
console.log(result); // => true
```

### 配列に要素が入っているインデックス番号を返す
indexOfを使う。複数該当するものがあっても最初のものをしか返さないが2つ目引数に開始位置を指定できるので複数回まわせばわかる。  
後ろから数えたい場合はlastIndexOfを使う
```TS
let List = ['foo','bar','baz'];
let result = List.indexOf('foo');
console.log(result); // => 0
```

### 配列の要素が全て一致しているか確認する(真偽値を返す)
everyを使う
```TS
let List = ['foo', 'bar', 'baz'];
let result = List.every(x => /\w/.test(x)); // \wは単語構成文字:[a-zA-Z_0-9]
console.log(result); // => true
```

### 連想配列が入った配列から特定の値が入った連想配列を抜き出す
jsonとかによくある[{key1:value,key2:value},{key1:value,key2:value}]の場合で連想配列を抜き出すにはfindを使う。
```TS
let List = [{key1:'foo',key2:'bar'},{key1:'baz',key2:'qux'}];
let result = List.find(x => x.key1 === 'foo');
console.log(result); // => {key1: 'foo', key2: 'bar'}
```

複数ヒットする可能性があるものはfindではなくfilterを使う

```TS
let List = [{key1:'foo',key2:'bar'},{key1:'baz',key2:'qux'}];
let result = List.filter(x => {
    if (x.key1 === 'foo') { return x };
});
console.log(result); // => [{key1: 'foo', key2: 'bar'}]
```
findと違い戻値も配列になっている。

