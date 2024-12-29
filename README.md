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
(Old design, To be update based on ArrEx)

This class to extension of using Array, by some functions.  
ArrayExCore is predeclared, and can use like samples.

- ArrayEx : ArrayEx
    - ArrayExCore.cls
    - ArrayEx0.cls 
    - ArrayEx1.cls
    - ArrayEx2.cls

## Features
- ArrayEx0, 1, 2 has some functions. These can be used like a chain method.
- ArrayEx1,2 has *Evaluated() methods. These used Evaluate() method instead of delegate.
- ArrayEx1,2 has Extract() and more methods. There use index as array, collection string (like a "1,2,3". "1 to 3", "1:3", ":")
- ArrayEx2 is depend on ArrayEx1,0, Core. 
- ArrayEx1 is depend on ArrayEx0, Core. 
- ArrayEx0 is depend on ArrayExCore.
- ArrayExCore is independent.
- ArrayExCore has some new excel functions like (VSTACK, HSTACK, CHOOSEROW, CHOOSECOLUMN, EXPAND, TEXTSPLIT, ...)


## Sample Code
~~~
Dim rearr
rearr = ArrayExCore.HSTACK([{1,2,3,}], [{4,5,6}])

Dim arr As New ArrayEx2
Call arr.Init(Range("A1:E3").Value) _
    .DebugPrintAll _
    .Extract("1:2", "3:5") _
    .WhereEvaluated("x", 1, "x>1") _
    .DebugPrintAll _
    .GetRow(2) _
    .DebugPrintAll _
    .SetElement(1, 11) _
    .DebugPrintAll _
    .ToRange(Range("A5"))
~~~

## Note
GitHubの練習と、最近勉強したVBAの練習と、作ったものがアウトプットされないと気が済まない気がしてきたというレポジトリ  

C#のメソッドチェーンのように書けると嬉しいなぁというアイディアから。思い立ったので自作。
また、CHOOSEROW, VSTACK, HSTACKなどの関数が使いたいなぁと思ったので自作。
作っているうちに気が付きましたが、インターネットを探すと素晴らしい先人の作品が既にありました...  
ただし、これは趣味のプログラムなので、なにかが優れているということは求めていません...  

- Rangeから値を取得して、配列として処理して、Rangeに値を代入することを主に想定しています。
- Whereなどではインスタンスをコピーして返します。
- 配列の次元ごとにインテリセンスで表示される内容を分けたかったので、クラスも次元毎に分けました。
- ArrayExCoreは単独で使用可能。ArrayEx0,1,2はArrayExCoreに依存します。
- ArrayEx1 はArrayEx0 に依存。ArrayEx2 はArrayEx0, ArrayEx1 に依存します。
- MATLABっぽい感じで要素を指定できると個人的に嬉しいので、":"、また "To" でGet出来るようにしました。
- LinqっぽいMethodはいくつかを抜粋して実装しました。
- 処理が高速というわけではない。

