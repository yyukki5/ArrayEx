# ArrayEx
GitHubの練習と、最近勉強したVBAの練習と、作ったものがアウトプットされないと気が済まない気がしてきたというレポジトリ  

- ArrayEx (本体)    : vba-files > Class  
    - ArrayEx0.cls 
    - ArrayEx1.cls
    - ArrayEx2.cls
- Test, Sample      : vba-files > Module, and ArrayEx.xlsm 

C#のメソッドチェーンのように書けると嬉しいなぁというアイディアから。思い立ったので自作。  
作っているうちに気が付きましたが、インターネットを探すと素晴らしい先人の作品が既にありました...  
ただし、これは趣味のプログラムなので、なにかが優れているということは求めていません...  

※ 以前のReposから個人情報を削除して再作成


## Features
- Rangeから値を取得して、配列として処理して、Rangeに値を代入することを主に想定しています。
- Getしたときにインスタンスをコピーして返します。
- 配列の次元ごとにインテリセンスで表示される内容を分けたかったので、クラスも次元毎に分けました。
- ArrayEx1 はArrayEx0 に依存。ArrayEx2 はArrayEx0, ArrayEx1 に依存します。
- MATLABっぽい感じで要素を指定できると個人的に嬉しいので、":"　や "To" でGet出来るようにしました。
- LinqっぽいMethodはいくつかを抜粋して実装しました。
- 処理が高速というわけではない。


## Sample Code
~~~
Dim arr As New ArrayEx2
Call arr.Init(Range("A1:E3").Value) _
    .DebugPrintAll _
    .Extract("1:2", "3:5") _
    .DebugPrintAll _
    .GetRow(2) _
    .DebugPrintAll _
    .SetElement(1, 11) _
    .DebugPrintAll _
    .ToRange(Range("A5"))
~~~
