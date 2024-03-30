# ArrayEx
Extension class of Array. this can use any functions to support using.
ArrayExCore is predeclared, and can use like samples.

- ArrayEx : src\ArrayEx.xlsm
    - ArrayExCore.cls
    - ArrayEx0.cls 
    - ArrayEx1.cls
    - ArrayEx2.cls
- Sample

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


いつも夜中に作っていたので、ミスっていたら申し訳ございません...
I'm sleepy zzZ