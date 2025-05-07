TW Stock MA Breakthrough Analyzer
📈 自動化抓取台灣股市資料、計算移動平均線，並篩選出符合「均線突破策略」的股票，透過圖表與 Email 自動通知。

📌 專案簡介
本專案是一個技術分析自動化工具，整合資料爬蟲、數據處理、均線計算、策略篩選、圖表繪製與 Email 通知等功能，特別針對台灣證券交易所（TWSE）與櫃買中心（TPEx）每日股票資訊。
> 可作為 AI 資料分析、金融科技、資訊管理研究等主題的學術實作作品。

🧠 主要功能
⏱ 自動抓取 TWSE/TPEx 的股票每日收盤資料
📊 計算 5/10/20 日移動平均線（MA）
🎯 篩選符合「收盤價從低於均線 → 突破均線」條件的股票
🧾 匯出 CSV 報告與 PDF 圖表
✉️ 自動 Email 通知（含圖表與資料檔案）
📅 支援排程每日自動執行（透過 `schedule` 套件）

🖥 技術架構
類別	技術
程式語言	Python 3
資料處理	pandas, numpy
爬蟲	requests, re
圖表繪製	matplotlib
自動化與排程	schedule, logging
郵件功能	smtplib, email.mime
檔案設定	`config.json` 統一控制配置

📦 執行方式
```bash
安裝必要套件
pip install -r requirements.txt
一次性執行分析
python 股票均值分析_學術版.py

📌 系統會自動：

1.下載過去數週股市資料

2.計算移動平均線

3.篩選突破股票

4.匯出 CSV 報表 + PDF 圖表

5.Email 通知結果（需正確設定 config.json）

學術/專題延伸建議:

1.加入技術分析指標：MACD、RSI、布林通道等

2.加入回測模組：與隨機策略進行績效比較（使用 backtrader）

3.加入前端介面：如用 Streamlit 做簡易圖形化操作

4.發布成網站或部署至雲端（Heroku / Render）

作者資訊
作者：王承偉(WANG CHENG-WEI)

系所：資訊管理系（應用於研究所考試作品）

聯絡：可於 GitHub Issues 討論

本專案僅用於學術研究與技術展示，未用於任何實際投資建議。請遵守相關資料來源的使用條款。