# PMS + 治具管理：整體架構圖

> 本文件包含兩張 Mermaid 圖：**系統總覽** 與 **庫存操作細節**。  
> 可將此檔丟到 Mermaid Live Editor 或直接使用下方 HTML 版本檔案瀏覽。

## 1) 系統總覽 (Overview)

```mermaid
flowchart LR
    %% ===================
    %% PMS + Fixture System Overview
    %% ===================
    classDef title fill:#f6f8fa,stroke:#d0d7de,color:#24292e,stroke-width:1px;
    classDef box fill:#fff,stroke:#d0d7de,stroke-width:1px;
    classDef db fill:#fdf6e3,stroke:#d0d7de,stroke-width:1px;
    classDef svc fill:#eef6ff,stroke:#b6d4fe,stroke-width:1px;
    classDef warn fill:#fff2cc,stroke:#f1c232,stroke-width:1px;
    classDef good fill:#eaffea,stroke:#b7e1b5,stroke-width:1px;

    %% Layers
    subgraph UI["🖥️ Tkinter 主介面 (Notebook 分頁)"]
        direction LR
        UI_LOGIN["登入 / 帳號權限<br/>Admin / Engineer / Leader / Operator"]:::box
        UI_SOP["SOP 管理<br/>上傳/合併/綁定/批量套用/統計"]:::box
        UI_FIX["治具管理<br/>新增治具料號(8或12碼)/入庫/轉倉/扣減<br/>查詢/分倉統計/即時刷新"]:::box
        UI_FBOM["治具 BOM<br/>建置生產所需治具需求清單"]:::box
        UI_DB["DB 管理<br/>切換 Dev/正式、備份/還原、WAL"]:::box
        UI_LOG["操作紀錄 / 版本追溯"]:::box
    end

    subgraph SVC["⚙️ 應用服務層"]
        direction TB
        SVC_AUTH["權限驗證 / Session / 閒置登出 / 單實例啟動"]:::svc
        SVC_SOP["SOP Service<br/>合併PDF / 綁定料號 / 套用 / 統計(各製程唯一料號數)"]:::svc
        SVC_FIX["Fixture Service<br/>insert_fixture()<br/>add_stock() / transfer_stock()<br/>倉別規則/防呆(不可負數)"]:::svc
        SVC_FBOM["Fixture BOM Service"]:::svc
        SVC_LOG["Log Service<br/>log_activity() / ensure_changelog_schema()"]:::svc
        SVC_DBM["DB 管理器<br/>WAL / 備份 / 還原 / 連線切換"]:::svc
        SVC_UPDATE["啟動更新檢查 / 自動更新"]:::svc
    end

    subgraph DATA["🗄️ 資料層 (SQLite)"]
        direction TB
        DB_MAIN["Main Server DB<br/>\\\\192.120.100.177\\工程部\\生產管理\\生產資訊平台\\PMS.db"]:::db
        DB_DEV["Dev 本地 DB (可切換)"]:::db
        T_USERS["users"]:::db
        T_SOP["sop_records"]:::db
        T_WARE["warehouses<br/>預設：虹堡 / 上齊 / 睿均 / 不良品<br/>支援自定新增倉別"]:::db
        T_FIX["fixtures<br/>料號(8或12碼)、品名、規格、單位、安庫、備註等"]:::db
        T_FBOM["fixture_bom<br/>需求結構、料號對應"]:::db
        T_LOG["changelog / activity_log"]:::db
    end

    %% UI -> Service
    UI_LOGIN --> SVC_AUTH
    UI_SOP --> SVC_SOP
    UI_FIX --> SVC_FIX
    UI_FBOM --> SVC_FBOM
    UI_DB --> SVC_DBM
    UI_LOG --> SVC_LOG

    %% Service -> Data
    SVC_AUTH --> T_USERS
    SVC_SOP --> T_SOP
    SVC_SOP --> T_LOG
    SVC_FIX --> T_FIX
    SVC_FIX --> T_WARE
    SVC_FIX --> T_LOG
    SVC_FBOM --> T_FBOM
    SVC_DBM --> DB_MAIN
    SVC_DBM --> DB_DEV
    SVC_LOG --> T_LOG

    %% Business Rules & Flows (notes)
    RULES1["規則：建立治具料號時，同步建立四倉資料；入庫僅限『虹堡』；轉倉可任意倉別但來源不可為負；所有欄位防呆；操作後即時刷新"]:::warn
    RULES2["SOP統計：各製程(DIP/組裝/測試/包裝/OQC)以『非空白的唯一料號數』計算"]:::warn
    RULES3["DB：WAL 模式、備份/還原、正式與開發路徑切換；changelog 由 ensure_changelog_schema() 初始化"]:::warn

    UI_FIX --- RULES1
    UI_SOP --- RULES2
    UI_DB  --- RULES3

    %% Good paths
    SVC_UPDATE:::good --> UI_LOGIN
```

---

## 2) 庫存操作細節 (Stock Operations)

```mermaid
flowchart TD
    %% ===================
    %% Fixture Stock Operations
    %% ===================
    classDef box fill:#fff,stroke:#d0d7de,stroke-width:1px;
    classDef warn fill:#fff2cc,stroke:#f1c232,stroke-width:1px;
    classDef db fill:#fdf6e3,stroke:#d0d7de,stroke-width:1px;

    subgraph Actors["操作端"]
        UI["UI：入庫 / 轉倉 / 扣減 按鈕"]:::box
    end

    subgraph Services["服務邏輯"]
        direction TB
        V1["驗證：料號長度必須 8 或 12 碼"]:::warn
        V2["驗證：數量為正整數、來源倉庫庫存不可為負"]:::warn
        V3["驗證：入庫倉別必須為『虹堡』"]:::warn
        A1["add_stock(part, qty, '虹堡')"]:::box
        T1["transfer_stock(part, qty, src, dst)"]:::box
        LOG["log_activity('入庫' / '轉倉' / '扣減')"]:::box
    end

    subgraph Data["資料層"]
        FIX["fixtures"]:::db
        WARE["warehouses(虹堡/上齊/睿均/不良品/… )"]:::db
        CHG["changelog / activity_log"]:::db
    end

    UI --> V1 --> V2
    UI --> V3
    V3 --> A1 --> WARE
    V2 --> T1 --> WARE
    A1 --> LOG --> CHG
    T1 --> LOG
    LOG --> CHG
    WARE --> FIX
```

---

## 備用：PlantUML 版本

```
@startuml
left to right direction
skinparam componentStyle rectangle
skinparam shadowing false
skinparam linetype ortho

package "UI (Tkinter - Notebook)" {
  [Login/ACL]
  [SOP 管理]
  [治具管理]
  [治具 BOM]
  [DB 管理]
  [操作紀錄視圖]
}

package "Services" {
  [Auth Service]
  [SOP Service]
  [Fixture Service]
  [Fixture BOM Service]
  [Log Service]
  [DB Manager]
  [Updater]
}

database "SQLite (WAL)" {
  [users]
  [sop_records]
  [warehouses]
  [fixtures]
  [fixture_bom]
  [changelog/activity_log]
}

[Login/ACL] --> [Auth Service]
[SOP 管理] --> [SOP Service]
[治具管理] --> [Fixture Service]
[治具 BOM] --> [Fixture BOM Service]
[DB 管理] --> [DB Manager]
[操作紀錄視圖] --> [Log Service]

[Auth Service] --> [users]
[SOP Service] --> [sop_records]
[SOP Service] --> [changelog/activity_log]
[Fixture Service] --> [fixtures]
[Fixture Service] --> [warehouses]
[Fixture Service] --> [changelog/activity_log]
[Fixture BOM Service] --> [fixture_bom]
[DB Manager] --> [users]
[DB Manager] --> [sop_records]
[DB Manager] --> [warehouses]
[DB Manager] --> [fixtures]
[DB Manager] --> [fixture_bom]
[DB Manager] --> [changelog/activity_log]

note right of [Fixture Service]
  - insert_fixture()
  - add_stock()  (only '虹堡')
  - transfer_stock() (any-to-any, non-negative)
  - 防呆：料號 8/12 碼、不可負數
end note

note right of [SOP Service]
  合併/綁定/批量套用；
  統計以各製程(DIP/組裝/測試/包裝/OQC)
  的「非空白唯一料號數」計算
end note

note bottom of [DB Manager]
  DB 切換：Dev <-> Main Server
  Main DB: \\\\192.120.100.177\\工程部\\生產管理\\生產資訊平台\\PMS.db
  ensure_changelog_schema() 初始化
end note
@enduml
```