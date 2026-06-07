# Open Design

此 Repository 包含了 Open Design 本地環境的原始碼，已依據開發者要求完成初始化與環境配置。

## 安裝與啟動說明

### 系統需求
- Node.js v24.x 或以上
- 啟用 Corepack (`corepack enable`)

### 安裝步驟
1. 確保已啟用 Corepack：
   ```bash
   corepack enable
   ```
2. 安裝所有相依套件：
   ```bash
   pnpm install
   ```

### 開發與啟動
啟動所有服務（背景的 Daemon、Web 介面與 Desktop 殼層）：
```bash
pnpm tools-dev
```

如果只想單獨啟動 Web：
```bash
pnpm tools-dev run web
```

### 注意事項
- 初次啟動會自動掃描您的系統 PATH 以尋找支援的 Code Agent CLI (例如 Gemini CLI 等)。
- 建立的原型或 Artifacts 會預設儲存在 `./.od/artifacts/` 中。
