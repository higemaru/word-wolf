# Word Wolf

Word Wolf Sample with Google Sheets and GAS (Only Japanese, so please translate it yourself if you're using another language) .

## Settings

Create the following sheet.

### Config

| Status | StartTime| Limit| Title | WoflName | CitizenName |
| ------ | -------- | ---- | ----- | -------- | ----------- |
|      |      | 300 |ワードウルフ オンライン|人狼|村人|

- Status, StartTime … for system
- Time for discussion. default: 300 sec
- Title, WorlName, CitizenName … Text will be used in the game.

### Words

One prompt per line. Which line appears and which side is the wolf are random.

| ワード1 | ワード2 |
| ------- | ------- |
|         |         |

### Players

Only users registered in the spreadsheet　can play.

- UserEmail … Required: Google Account for login
- Nickname … Required: Nickname
- Facilitator … true or false. If set to true, you'll be the hosts, and the screen layout will change. You might want to use a checkbox for this.
- The other columns are used by the system

| UserEmail | Nickname | Role | Word | Checked | Vote | Facilitator |
| --------- | -------- | ---- | ---- | ------- | ---- | ----------- |
|           |          |      |      |         |      |             |

## Guide for Hosts

Since the host is also play in the game, there will be no info beforehand. Administrative buttons will apper for Hosts.

|      | プレイヤー                 | ファシリテーター             |
| ---- | -------------------------- | ---------------------------- |
| 1    | 「進行役が準備しています」 | 狼の人数を決めて、お題を配布 |
| 2    | ワードを確認する           | ＋ゲーム開始                 |
| 3    | カウントダウン             | ＋投票を開始する             |
| 4    | 投票                       | ＋正解を表示する             |
| 5    | 結果発表                   | ＋次のゲームへ               |

## Others

- The screen is updated via polling, there is a delay of a few seconds between each player's screen.
- The interval is 3 seconds. To change it, modify the "3000" at the very bottom of index.html.

## LICENSE

Zero-Clause BSD