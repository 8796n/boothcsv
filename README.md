# boothcsv
BOOTHの宛名印刷用CSVを使って帳票を作るやつ

### おしらせ
* これは便利だーって思っていただいたおやさしい方はデバッグに協力すると思って[NyanticLabs.](https://nyantic.booth.pm/)からなんか買ってください。これを使って帳票印刷して送ります。

## 想定する使用対象者は誰？
[BOOTH](https://booth.pm/)で自家通販をしている人で主に「[あんしんBOOTHパック](https://booth.pm/anshin_booth_pack_guides)」の利用者

## おまえはだれだ？
JavaScriptとかなるべく使いたくないBOOTHで自家通販しているあんしんBOOTHパック利用者の猫グッズ屋のおじさん。  
端的に言えば自分で使う用。

## なんで作ったの？
猫グッズを売るときに手数料も安いし手軽に出品できるし匿名配送も使えるのでpixivさんのやってるBOOTHって便利なんですよ。  
そのBOOTHで自家通販をすると、当然ですが売れたら誰かに何かを送ります。  
発送処理をしていて箱に入れる商品どれ？いくつ？みたいなことがよくあったので誤発送を防止するために作りました。 
  
あんしんBOOTHパックで送るときに誰に送るかが分からないため、外箱に注文番号とか印刷したシールがあると便利だよね？それにネコピットで読み込むQRコードがあればもっと便利だよね？ってことでそれもできるようにしました。  
仕組みで防げる誤発送は無くそう！

## なにができるの？
出荷準備のときにあると便利な帳票と「あんしんBOOTHパック」で出荷のときに便利なQRコードをラベルシールに印刷できます。

## 動作環境は？
Windows版のChrome 73で動作確認をしています。  
最近のChromeなら動くと思いますが仕様が変わると動かない機能もあるかもしれません。  
スマートフォンのChromeだとラベルシール印刷がズレると思うのでたぶん駄目です。  
JavaScriptでどうにかしてるので**別途サーバーなどは必要ありませんし外部に情報を送ることもありません。**  

## どうやって使うの？
* [ここ](https://github.com/8796n/boothcsv/archive/master.zip)からzipファイルをダウンロードして展開します。
* boothcsv.html を Chrome から開きます。
* ファイルを選択ボタンを押して[BOOTHの注文一覧](https://manage.booth.pm/orders?state=paid)からダウンロードした宛名印刷用CSVを選択します。
* 実行ボタンを押すと印刷に必要な枚数が表示されます。
* ラベルシールも印刷するときの説明はまたあとで。
* プリンターに用紙をセットして印刷します。
* おしまい。
![サンプル](https://user-images.githubusercontent.com/982314/55877503-17fb0800-5bd5-11e9-9338-9a03d81e67ef.png)


## ラベルシールも印刷したい
これね、ヤマトの営業所持ち込みであんしんBOOTHパック使うときにはすげえ便利なんですよ。[PUDOステーション](https://booth.pixiv.help/hc/ja/articles/360013148033)に放り込むときにもめっちゃ便利です。慣れると1出荷30秒ぐらいでできます。
* A4 44面 四辺余白付のラベルシールを用意します。例えば[これ](https://amzn.to/2KkRXhE)とか[これ](https://amzn.to/2KpdW7k)</a>
* ラベルシールも印刷するにチェックを入れます。
* 途中まで使ったラベルシールを再使用する場合にはスキップする枚数を入れます。あまり再使用しすぎるとプリンターの中で剥がれたりするかもしれないのでほどほどに。
* さっきと一緒でCSVファイルを選んで実行ボタンを押すと、ラベル用紙の枚数も表示されます。
![ラベルシール](https://user-images.githubusercontent.com/982314/55878887-5940e700-5bd8-11e9-963a-20ec106db3ad.png)
* BOOTHの注文詳細で[あんしんBOOTHパックのQRコードを作成](https://booth.pm/anshin_booth_pack_guides/usage)して、表示されたQRコードの画像を右クリックでコピーしてから該当の注文番号のラベルの「Paste QR image here!」を右クリックして貼り付けを選ぶとQRコードが表示されて自動的に匿名配送の受付番号とパスワードが表示されます。自動です。オートマチック！！
![QRコード貼付け後](https://user-images.githubusercontent.com/982314/55879200-fd2a9280-5bd8-11e9-8ffa-5c3ffa7b2254.png)
* なお間違えて貼り付けたときはQRコードをクリックすると再度貼り付けできるようになります。
* 印刷したラベルシールを出荷する箱に貼り付けておけば出荷時にネコピットを箱に貼ったQRコードにかざす→その箱に貼る送り状が印刷されるという流れになるので送り間違いは発生しません。万が一QRコードの読み込みが上手く行かなかった場合でもラベルシールに印刷されている受付番号とパスワードを入力すれば送り状が出力できるので安心です。

## BOOTHで売れてないけど試したい
* sampleにcsvファイルとQRコードのpngファイルがあるので試せます。
* 参考までに手元にあるpngファイルは直接該当箇所にドラッグアンドドロップすれば貼り付けられます。
* 当然ですが、サンプルのデータはダミーの数字が入っているので出荷では使えません。

## 使ってる仕組み
* 印刷に便利なCSS [Paper CSS](https://github.com/cognitom/paper-css)
* QRコード読んでくれるJavaScript [jsQR](https://github.com/cozmo/jsQR)
* CSVをいい感じに読んでくれるJavaScript [PapaParse](https://github.com/mholt/PapaParse)

## 裏側の見どころ
* こだわりの素JavaScriptでHTML5のtemplateタグを使ってます。最近はquerySelectorのおかげでjQueryなしでも要素の選択が割と簡単にできるようになったみたいです。
* Paper CSS強い。templateと合わせると通常の業務で使う帳票はだいたい表現できるぞ！サーバーでPDF作ってダウンロードさせて印刷とかしなくてもいいし、Chrome上での見た目がいきなり印刷プレビュー風で謎の安心感があります。
* BOOTHからダウンロードするCSVのヘッダが日本語なので、classとかも全力で日本語です。可読性高い！

## 動けばOK主義者
* なんで動いてるかよくわからないところはだいたい潰して理解したつもりです。
* 印刷物に関係ないフォームの部分などは動けばOK主義者なので必要最小限です。見た目にこだわる方はプルリクエスト？とかいうやつをしていただけるといいと思います。

## 文字打つの飽きたので続きはまた今度