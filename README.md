# 📝 英文單字小考

這是給手機用的英文單字／片語／克漏字小考網頁。整個工具只有一個網頁，不用安裝 App，打開網址就能用。

## 功能

- **上傳題庫**：可以上傳 .txt / .csv，也可以從 Excel、Google 試算表直接複製貼上
- **四種題型**（可以複選）：
  - 中翻英：看或聽中文，拼出英文（單字、片語、整句都行）
  - 聽音拼字：聽英文發音，拼出單字
  - 克漏字：句子挖空填字
  - 英翻中選擇：四選一
- **題目怎麼出**：文字＋語音、只用語音，或只有文字
- **作答方式**：可以打字，也可以按 🎤 用說的。用說的時候，直接唸單字或一個一個字母拼（W‧O‧R‧D）都行
- **答錯處理**：可以按「再試一次」或「算我對」（語音辨識聽錯時用），答錯的字會記進「錯題本」
- **分享題庫連結**：把題庫做成一個網址傳給別人，對方點開就能考同一份

## 題庫格式

```
# 每行一題：英文 | 中文 | 例句（要考的字用 [ ] 括起來，例句可省略）
apple | 蘋果 | I eat an [apple] every day.
go | 去 | Yesterday she [went] to the park.
look forward to | 期待 | I'm looking forward to the trip.
You can definitely do it. | 你絕對可以的。

# 克漏字：答案用 [ ] 括起來 | 中文提示
He is good [at] playing basketball. | 他很擅長打籃球。

# 同義答案用 / 分開
movie/film | 電影
```

- 例句沒有標 `[ ]` 時，程式會自己找出要考的字並挖空（-s / -ed / -ing 變化都找得到）
- 不規則變化（go → went）要自己用 `[ ]` 標
- 只有考卷照片或 PDF 的話：在題庫頁打開「🤖 手上只有考卷照片或 PDF？」，複製裡面的指令，連同照片一起交給 Claude，再把它回傳的結果貼回來

## 開成網址（GitHub Pages）

1. 到 GitHub 這個 repo 的 **Settings → Pages**
2. Source 選 **Deploy from a branch**，Branch 選 `main`、資料夾選 `/ (root)`，按 Save
3. 等 1～2 分鐘，網址會是 `https://<你的帳號>.github.io/English-Test/`
4. 用手機打開這個網址，可以「加入主畫面」，用起來就像 App

> 語音作答建議用 iPhone 的 Safari 或 Android 的 Chrome。LINE、FB 的內建瀏覽器通常不能用麥克風。第一次按 🎤 時要允許麥克風權限。
