# TODO.md — タスク管理

## 完了済み

### CR実績分析（2026/03/17）
- [x] Excelデータの読み込み・構造把握
- [x] PLAN.md / SPEC.md / TODO.md / KNOWLEDGE.md の作成
- [x] キャンペーン別集計
- [x] クリエイティブタイプ別分析（画像/動画/カルーセル）
- [x] CV獲得CRのCPA順ランキング
- [x] 分析Excelの作成（パンスク広告分析_レポート.xlsx）

### 解約分析（2026/03/19）
- [x] leaveReason分布の集計（価格系が断トツ1位と確認）
- [x] 流入源別解約率の算出（Meta ASC 21.4% / panforyou.jp 12.5%が最良）
- [x] 解約までの平均日数の計算（kepco.jp 48日 / meta ASC 85日）
- [x] 月別解約コホート分析（kepco.jpの月別ブレ・Meta ASC 2025-08最良を確認）
- [x] クーポン即解約パターンの確認（入会後2〜18日で解約するalliance-coupon系を特定）

### データ基盤調査 × LTV分析（2026/03/19）
- [x] サイト構成の整理（WordPress / STUDIO / スクラッチ）
- [x] BigQueryデータセット一覧の確認（analytics_312681459が正解と特定）
- [x] Firestore accountsテーブルのスキーマ確認（gaClientIdの存在を確認）
- [x] GA4 × Firestore JOIN の検証（gaClientId = user_pseudo_id で結合成功）
- [x] Meta広告経由の登録者分析（47名 / register_complete）
- [x] 決済完了者（registerCompleted=true）でのLTV分析（187名）
- [x] 全流入源 × コホート分析（2025-03〜現在）
- [x] Meta ASCのROI検証（12ヶ月コホートで2.1倍黒字を確認）

## 次のアクション（優先順）

- [ ] **代理店への依頼**: Meta広告の全CRに `utm_content={{ad.name}}` を追加 → CR単位LTV追跡の基盤
- [ ] **2024-01の5,435名の異常値の原因確認**（エンジニアまたは過去のキャンペーン記録で調査）
- [ ] **kepco.jpの月別ブレ原因調査**（2025-06/10に解約率50%超。その月の連携条件・訴求内容の変化を確認）
- [ ] **twoMonthsプランへの誘導強化の検討**（継続率94.7%でoneMonthより19pt高い）

- [ ] **alliance-couponキャンペーンの精査**（aeoncard解約率85.7%は即停止検討。goopan500/persona_HPは継続）
- [ ] **オンボーディング改善**（「食べきれない」「冷凍庫に入らない」解約への対処 → 初回量説明・保存方法案内）
- [ ] **価格訴求の見直し**（解約理由1位が「高い」→ 広告・LPでの価値訴求強化）

## 今後の追加分析（任意）

- [ ] Meta ASC月別継続率トレンドのグラフ化
- [x] leaveReason（解約理由）の集計・分析（完了）
- [ ] フリークエンシーとCTRの相関分析
- [ ] 予算配分シミュレーション（上位CRに集中した場合の試算）

## ステータス

最終更新: 2026/03/19
担当: Claude（AI分析）
