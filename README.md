# PSWebDriver

Selenium WebDriverをPowerShellから使用しやすくするラッパークラス+便利な関数群

## インストール方法
1. 使用したい[ブラウザ用のDriver](http://www.seleniumhq.org/download/#thirdPartyDrivers)をインストールしてPATHを通しておきます（Firefoxは不要）
1. このレポジトリをダウンロードしてPowerShellのモジュールディレクトリに配置します

## 使い方

Google Chromeを使用して[DuckDuckGo](https://duckduckgo.com/)を開き、"PowerShell"と検索、検索結果のスクリーンショットを取得してブラウザを終了する例
```PowerShell
#モジュールの読み込み
using module PSWebDriver
#インスタンスの生成
$Browser = New-Object PSWebDriver('Chrome') #Chrome/Firefox/Edge/IE/HeadlessChrome
#ブラウザを起動
$Browser.Start()
#DuckDuckGoを開く
$Browser.Open('https://duckduckgo.com/')
#検索ボックスに"PowerShell"と入力
$Browser.SendKeys('id=search_form_input_homepage', 'PowerShell')
#検索ボタンをクリック
$Browser.Click('id=search_button_homepage')
#スクリーンショットを保存
$Browser.SaveScreenShot('D:\screenshot.png')
#ブラウザを閉じる
$Browser.Close()
```

----
## リファレンス
### PSWebDriver クラス
Selenium WebDriverをPowerShellから利用するためのラッパークラスです。
メソッド名などSelenium IDEで作成したテストケースをPowerShellスクリプトに置き換えやすくするよう意識しています。

----
#### プロパティ
|名前|型|説明|
----|----|----
|Driver|未定義|WebDriverインスタンス<br>ラップされていないWebDriverネイティブのメソッドを利用したい場合に使えます|

----
#### メソッド
##### 起動/停止
|名前|戻り値型|説明|
----|----|----
|Start()|void|ブラウザを起動します|
|Start([Uri]$URL)|void|ブラウザを起動し、`$URL`で指定されたページを開きます|
|Close()|void|カレントウィンドウを閉じます|
|Quit()|void|ブラウザを閉じます|

##### ブラウザ設定
|名前|戻り値型|説明|
----|----|----
|SetImplicitWait([int]$TimeoutInSeconds)|void|要素検索やページ読込時の暗黙的な待機時間(秒)を指定します|
|GetWindowSize()|System.Drawing.Size|ブラウザのウィンドウサイズを取得します|
|SetWindowSize([System.Drawing.Size]$Size)|void|ブラウザのウィンドウサイズを変更します|
|SetWindowSize([int]$Width,[int]$Height)|void|ブラウザのウィンドウサイズを変更します|

##### 要素検索 / 情報取得
|名前|戻り値型|説明|
----|----|----
|FindElement([string]$SelectorExpression)|Object|`$SelectorExpression`で指定されるページ内の要素を取得します|
|IsElementPresent([string]$SelectorExpression)|bool|`$SelectorExpression`で指定される要素が存在するか確認します|
|GetTitle()|string|現在開いているページタイトルを取得します|
|IsAlertPresent()|bool|アラートが表示されているか確認します|

##### 操作
|名前|戻り値型|説明|
----|----|----
|SendKeys([string]$Target, [string]$Value)|void|`$Target`(SelectorExpression)で指定される要素に対して`$Value`を入力します<br>特殊キーの送信については下部の「特殊キーの入力について」を参照してください|
|ClearAndType([string]$Target, [string]$Value)|void|`$Target`(SelectorExpression)で指定される要素に対して`$Value`を入力します<br>既存の内容をクリアしてから入力する点が`SendKeys()`との違いです<br>(Selenium IDEの`Type`コマンドに相当します)|
|Click([string]$Target)|void|`$Target`(SelectorExpression)で指定される要素をクリックします|
|Select([string]$Target, [string]$Value)|void|`$Target`(SelectorExpression)で指定されるSelect要素から`$Value`をテキストに持つ要素を選択します|
|CloseAlert()|void|アラートを閉じます|
|CloseAlertAndGetText([bool]$Accept)|string|アラートテキストを取得し、アラートを閉じます<br>`$Accept`でアラートに対する`OK` or `Cancel`を指定できます|

##### 待機
|名前|戻り値型|説明|
----|----|----
|WaitForPageToLoad([int]$Timeout)|bool|ページの読み込みが完了するか、`$Timeout`で指定された秒数が経過するまで待機します<br>読み込み完了の場合は`$true`、タイムアウトの場合は`$false`を返します|
|WaitForElementPresent([string]$Target, [int]$Timeout)|bool|`$Target`(SelectorExpression)で指定される要素が見つかるか、`$Timeout`で指定された秒数が経過するまで待機します<br>要素が見つかった場合は`$true`、タイムアウトの場合は`$false`を返します|
|WaitForNotElementPresent([string]$Target, [int]$Timeout)|bool|`$Target`(SelectorExpression)で指定される要素が見つからないか、`$Timeout`で指定された秒数が経過するまで待機します|
|WaitForValue([string]$Target, [string]$Value, [int]$Timeout)|bool|`$Target`(SelectorExpression)で指定される要素の`value`属性が`$Value`で指定された値と一致するか、`$Timeout`で指定された秒数が経過するまで待機します|
|WaitForNotValue([string]$Target, [string]$Value, [int]$Timeout)|bool|`$Target`(SelectorExpression)で指定される要素の`value`属性が`$Value`で指定された値と異なるか、`$Timeout`で指定された秒数が経過するまで待機します|
|WaitForText([string]$Target, [string]$Value, [int]$Timeout)|bool|`$Target`(SelectorExpression)で指定される要素の要素値が`$Value`で指定された値と一致するか、`$Timeout`で指定された秒数が経過するまで待機します|
|WaitForNotText([string]$Target, [string]$Value, [int]$Timeout)|bool|`$Target`(SelectorExpression)で指定される要素の要素値が`$Value`で指定された値と異なるか、`$Timeout`で指定された秒数が経過するまで待機します|
|WaitForVisible([string]$Target, [int]$Timeout)|bool|`$Target`(SelectorExpression)で指定される要素が表示されるか、`$Timeout`で指定された秒数が経過するまで待機します|
|WaitForNotVisible([string]$Target, [int]$Timeout)|bool|`$Target`(SelectorExpression)で指定される要素が表示されなくなるか、`$Timeout`で指定された秒数が経過するまで待機します|
|WaitForTitle([string]$Value, [int]$Timeout)|bool|現在のページタイトルが`$Value`と一致するか、`$Timeout`で指定された秒数が経過するまで待機します|
|WaitForNotTitle([string]$Value, [int]$Timeout)|bool|現在のページタイトルが`$Value`と異なるか、`$Timeout`で指定された秒数が経過するまで待機します|
|Pause([int]$WaitTimeInMilliSeconds)|void|`$WaitTimeInMilliSeconds`で指定された時間(ミリ秒)待機します<br>([System.Threading.Thread]::Sleep()と同等です)|

##### スクリプト実行
|名前|戻り値型|説明|
----|----|----
|ExecuteScript([string]$Script)|string|ページ上でJavaScriptを実行します|
|ExecuteScript([string]$Target, [string]$Script)|string|`$Target`(SelectorExpression)で指定される要素に対してJavaScriptを実行します|

##### スクリーンショット
|名前|戻り値型|説明|
----|----|----
|SaveScreenShot([string]$FileName)|void|スクリーンショットを保存します<br>画像形式はPNGです|
|SaveScreenShot([string]$FileName, [string]$ImageFormat)|void|画像形式を指定してスクリーンショットを保存します<br>`$ImageFormat`に指定可能な値は`Png`,`Jpeg`,`Gif`,`Tiff`,`Bmp`です|
|StartAnimationRecord([int]$Interval)|void|ブラウザ表示の動画記録を開始します<br>`$Interval`(ミリ秒)で指定した間隔で記録します<br>記録間隔の最小値は500msです<br>最大1200フレームまで記録できます|
|StartAnimationRecord()|void|`StartAnimationRecord()`で開始した動画記録を終了します<br>記録された動画は破棄されます|
|StartAnimationRecord([string]$FileName)|void|`StartAnimationRecord()`で開始した動画記録を終了し、ファイルに保存します<br>動画形式はアニメーションGIFです|

----
### SelectorExpressionについて

Webページ上の特定要素を指定するためのパターン文字列です。
Selenium IDEのlocatorに相当します。書式もlocatorとほぼ同等です。
以下の7種類が使用可能です。

* IDパターン
`id`属性値を指定して要素を特定します。
書式は`"id=idvalue"`です。

* Nameパターン
`name`属性値を指定して要素を特定します。
書式は`"name=elementname"`です。

* Tagパターン
DOMタグを指定して要素を特定します。
書式は`"tag=tagname"`です。

* ClassNameパターン
Class名を指定して要素を特定します。
書式は`"classname=classname"`です。

* LinkTextパターン
LinkのTextを指定して要素を特定します。※完全一致
書式は`"link=linktext"`です。

* XPathパターン
XPath構文を使用して要素を特定します。
書式は`"xpath=xpath"`です。
SelectorExpressionが`/`で始まる場合もXPathパターンとみなされます。
（`"xpath=/html/body/h1"`と`"/html/body/h1"`は同等です）

* CSSセレクタパターン
CSSセレクタを使用して要素を特定します。
書式は`"css=selector"`です。

----
#### テキスト検索パターンについて
`WaitForText()`や`WaitForValue()`など要素値を検証する一部のメソッドでは検索対象文字列に特殊な書式を使用することで検索パターンを指定することができます。

* グロビングパターン
いわゆるワイルドカード検索です。PowerShellの`-like`演算子に相当します。
特殊な書式を使用しない場合はデフォルトでグロビングパターンが使用されます。
明示的に指定する場合は検索文字列の前に`glob:`を付けます。
例) `glob:sometext*`

* 正規表現パターン
正規表現を用いて検索します。PowerShellの`-match`演算子に相当します。
検索文字列の前に`regexp:`を付けます。
例) `regexp:^Number[0-9]`

* 完全一致パターン
完全一致検索を行います。PowerShellの`-eq`演算子に相当します。
`*(アスタリスク)`などの特殊文字を検索したい場合に使用します。
検索文字列の前に`exact:`を付けます。
例) `exact:***asterisk****`

----
#### 特殊キーの入力について
`SendKeys()`や`ClearAndType()`でEnterキーや矢印キーなどの特殊キーを送信する場合は、`${KEY_CODE}`という書式を使用します。
使用可能なKEY_CODEの一覧は[こちら](/Static/KEYMAP.txt)

例）`SendKeys()`メソッドを使用してABC[Backspace][Enter]と入力する例
```PowerShell
$Browser.SendKeys('id=target', 'ABC${KEY_BACKSPACE}${KEY_ENTER}')
```

----
## ライセンス
> Copyright (c) 2018 mkht
> PSWebDriver is released under the MIT License
> https://github.com/mkht/PSWebDriver/blob/master/LICENSE
>
> PSWebDriver includes these software / libraries.
> * Selenium.WebDriver
> Copyright (c) 2017 Software Freedom Conservancy
> Licensed under the [Apache 2.0 License](http://www.apache.org/licenses/LICENSE-2.0).
>
> * Selenium.Support
> Copyright (c) 2017 Software Freedom Conservancy
> Licensed under the [Apache 2.0 License](http://www.apache.org/licenses/LICENSE-2.0).
>
> * AnimatedGif
> Copyright (c) 2017 mrousavy
> Licensed under the [MIT License](https://github.com/mrousavy/AnimatedGif/blob/master/LICENSE).
