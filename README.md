# ArrEx
Simplified ArrayEx.  
This class to extension of using Array, by some functions.

~~~ vba
val1 = ArrEx(val) _
        .RedimPreserve(aexRank2, 1, 4, 1, 5) _
        .WhereBy(2, aexGreaterThan, 2) _
        .OrderByDescending(1) _
        .SelectRows(1, 2) _
        .SelectColumns(2, 3) _
        .LeftJoin(HelloWorlds, 1, 1) _
        .SelectColumns(4) _
        .DebugPrint("," & vbTab)
~~~

## Features
1. Predeclared.
1. Create() is default function, return new instance.
1. By predeclared and Create(), no need to create instance for using.
1. Supporting dimension 0, 1, 2. Not supporting greater than 2.
1. Some functions return new instance of ArrEx. These can be used chain method.
1. Property Value() returns value.
1. Immutable

*Specification may be changed until version 1.0.0. 

### Methods
- RedimPreserve()
- SelectColumns(), SelectRows()
- WhereBy()
- Skip(), Take()
- OrderBy(), OrderByDescending()
- Distinct(), DistinctBy()
- VerticalStack(), HorizontalStack()
- XLookUp()
- InnerJoin(), LeftJoin(), FullOuterJoin(), CrossJoin()
- ...


## Japanese Note
ArrayExをもとにColExのアイディアで作り直したもの。  
こちらは次元でクラスを分けずに一つのクラスで実装。  
PredeclaredでCreate()をデフォルトにしているので短い記述で使えます。  
個人的に配列関係で便利だなと思うメソッドを実装しています。  
Initialize() 以外のSetter が無いImmutable な実装にしています。  


# ArrayEx
(Old design, To be moved to old ver/)

This class to extension of using Array, by some functions.  
Some functions using Lambda

## Dependency classes
- Lambda.cls (<- see [Lambda Repository](https://github.com/yyukki5/Lambda))

## Features
Some functions using Lambda
- WhereBy()
- SelectBy() : Using lambda function for (1 dim -> each element, 2 dim -> each row)


## Sample Code
~~~
Call ArrayEx(array2d) _
    .DebugPrint("," & vbTab) _
    .WhereBy("x => x(1) > 2") _
    .DebugPrint("," & vbTab) _
    .SelectBy("x=>x(2) + x(3)") _
    .DebugPrint("," & vbTab) _
    .OrderByDescending _
    .DebugPrint("," & vbTab)
~~~

## Note
GitHubの練習と、最近勉強したVBAの練習と、作ったものがアウトプットされないと気が済まない気がしてきたというレポジトリ  

C#のメソッドチェーンのように書けると嬉しいなぁというアイディアから。思い立ったので自作。
また、CHOOSEROW, VSTACK, HSTACKなどの関数が使いたいなぁと思ったので自作。
作っているうちに気が付きましたが、インターネットを探すと素晴らしい先人の作品が既にありました...  
ただし、これは趣味のプログラムなので、なにかが優れているということは求めていません...  

- ArrayEx0,1,2,Coreは old ver に移動
- LinqっぽいMethodはいくつかを抜粋して実装しました。
- Lambdaを使っている部分はすこし処理に時間が掛かるので注意。

