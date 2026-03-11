# Agent Timer Kicker

**Claude Code** と **Codex CLI** の5時間ローリングウィンドウを自動キックし、レート制限のリフレッシュを早めるツールです。

## なぜ必要？

Claude Code と Codex には **5時間のローリングウィンドウ** に基づくレート制限があります。最初のリクエストからカウントが始まるため、何も送らなければタイマーは開始されません：

```
キックなし:
  ──────────────────────┤ 放置 ├───── やっとリクエスト ──▶ ここから5時間
                                                          ↑ リフレッシュが遅い

キックあり:
  ── kick ──────────── 5時間 ────────── リフレッシュ! ── kick ──▶ また5時間開始
  ↑ すぐ開始                            ↑ 必要な時にはもう回復済み
```

最小限のプロンプト (`"just say hi and nothing else."`) を定期送信することで、ウィンドウを常に回転させ、早めに容量を回復させます。

## 特徴

- **Claude / Codex 独立制御** — それぞれ個別の間隔・WSL設定・有効/無効
- **Windows タスクスケジューラ連携** — GUI を閉じても動作し続ける
- **最小モデル** を使用しトークン消費を最小化（Claude: `haiku`, Codex: `gpt-5-codex-mini`）
- **セッション蓄積防止** — Claude は `--no-session-persistence`、Codex は `--ephemeral` を使用
- **ワンクリック** で有効化 / 無効化 / テスト実行
- ログビューア付き
- 単一 `.ps1` ファイル、インストール不要

## スクリーンショット

```
┌─────────────────────────────────────────────┐
│  Agent Timer Kicker                          │
├─── Claude ──────────────────────────────────┤
│  間隔 (分): [30]  ☐ WSL                    │
│  定期キック:  OFF                                 │
│  [ ▶ 有効化 ]  [ ⏹ 無効化 ]  [ 今すぐ実行 ] │
├─── Codex ───────────────────────────────────┤
│  間隔 (分): [30]  ☐ WSL                    │
│  定期キック:  OFF                                 │
│  [ ▶ 有効化 ]  [ ⏹ 無効化 ]  [ 今すぐ実行 ] │
├─── ログ ────────────────────────────────────┤
│  2026-03-11 09:00:00 [Claude] OK            │
│  2026-03-11 09:30:00 [Claude] OK            │
│  [ ログ消去 ]                               │
└─────────────────────────────────────────────┘
```

## 動作要件

| 要件 | 備考 |
|---|---|
| Windows 10 / 11 | タスクスケジューラ + PowerShell 5.1 |
| Claude Code CLI | `claude` が PATH にあり認証済み |
| Codex CLI | `codex` が PATH にあり認証済み（任意） |
| WSL（任意） | 「WSL」オプションをチェックする場合のみ |

## クイックスタート

### 方法A: ダブルクリック

1. このリポジトリをダウンロードまたは clone
2. **`launch.bat`** をダブルクリック
3. 間隔を設定し、必要なら WSL にチェックを入れ、**▶ 有効化** をクリック

> **SmartScreen の警告が出た場合**: 初回実行時に「Windows によって PC が保護されました」と表示されることがあります。**詳細情報** → **実行** をクリックしてください。`launch.bat` は内部で `powershell -ExecutionPolicy Bypass -File claude-timer-kick.ps1` を実行しているだけです。

### 方法B: PowerShell

```powershell
powershell -ExecutionPolicy Bypass -File claude-timer-kick.ps1
```

以上です。スケジュールタスクがバックグラウンドで動き続けるため、GUI は閉じても構いません。

## 仕組み

1. **有効化** で Windows スケジュールタスク (`AgentTimerKick-Claude` / `AgentTimerKick-Codex`) を登録
2. タスクが N 分ごと（デフォルト: 30分）に実行：
   ```
   claude -p --no-session-persistence --model haiku "just say hi and nothing else."
   ```
   Codex の場合：
   ```
   codex exec --ephemeral --model gpt-5-codex-mini "just say hi and nothing else."
   ```
3. **WSL** にチェックがある場合、コマンドの先頭に `wsl` を付与
4. 結果は `data/kick.log` に記録
5. **無効化** でスケジュールタスクを削除 — 残留物なし

## 設定

設定は `data/config.json` に保存され、次回起動時に復元されます。

**モデル** や **プロンプト** を変更するには、`claude-timer-kick.ps1` の先頭を編集：

```powershell
$Script:DefaultPrompt  = "just say hi and nothing else."
$Script:ClaudeModel    = "haiku"
$Script:CodexModel     = "gpt-5-codex-mini"
```

変更後は **無効化** → **有効化** で反映されます。

## アンインストール

1. `launch.bat` で GUI を起動
2. Claude・Codex 両方の **⏹ 無効化** をクリック（タスクスケジューラからタスクが削除されます）
3. リポジトリフォルダごと削除

**注意**: 無効化せずにフォルダを削除すると、タスクスケジューラにタスクが残ります。その場合はタスクスケジューラ (`taskschd.msc`) を開き、`AgentTimerKick-Claude` / `AgentTimerKick-Codex` を手動で削除してください。

レジストリ、サービス、スタートアップへの登録は一切ありません。

## FAQ

**Q: どのような仕組みで動作していますか？**
A: 公式 CLI (`claude -p`, `codex exec`) を通じて通常のリクエストを送信しています。特殊な API や非公式な手段は使用していません。

**Q: WSL 内の cron ではダメ？**
A: WSL の cron は WSL インスタンスの起動が必要です。Windows タスクスケジューラの方が確実で、WSL がアイドル状態でも動作します。cron を使いたい場合は手動設定も可能です：
```bash
*/30 * * * * claude -p --no-session-persistence --model haiku "just say hi and nothing else." >> ~/kick.log 2>&1
```

**Q: Claude と Codex で異なる間隔を設定できますか？**
A: はい。独立セクションになっているので、それぞれの利用パターンに合わせて設定できます。

**Q: `claude` / `codex` が PATH にない場合は？**
A: まず **今すぐ実行** をクリックしてテストしてください。失敗する場合は CLI がインストール・認証済みか確認してください。WSL の場合は WSL ディストロ内に CLI がインストールされている必要があります。

## ライセンス

MIT
