# World Info Suite

[繁體中文](#繁體中文) | [English](#english)

## 繁體中文

World Info Suite 是 SillyTavern 擴充功能，提供 World Info / Lorebook 的檢視、條目整理、批量處理、排序與關鍵字管理工具。

### 功能總覽

#### 1. 觸發條目檢視器

- 在 AI 訊息上顯示檢視按鈕，查看該回覆觸發的世界書條目。
- 依插入位置分組，並顯示條目來源與狀態。
- 點擊分組標題可收合／展開該區塊的條目。
- 出現兩個以上分組時，左側顯示快速導覽圓點（游標停留顯示分組名稱），點擊可平滑捲動到該區塊。
- 可展開查看關鍵字、次要過濾與內容等明細。
- 預設保留最近 10 則訊息的觸發資料；可調整上限（`0` 為無限制）或一鍵清除快取。

#### 2. 角色世界書快速跳轉

- 在角色管理區顯示角色關聯的世界書。
- 區分主要、額外、聊天與全域來源。
- 點擊即可直接開啟對應世界書。

#### 3. 批量條目編輯器

- 在世界書編輯器工具列提供批量編輯按鈕。
- 可多選條目並套用同一份修改範本。
- 支援移動或複製選中條目到其他世界書、複製多筆內容，以及批量刪除確認。
- 原生條目展開編輯區的 Content 標題列提供單條目複製按鈕。
- 可依目前排序重整條目 UID。

#### 4. 世界書條目收納

- 工具列的資料夾圖示按鈕位於批量編輯按鈕左側；點擊後可建立命名的可摺疊資料夾。
- 資料夾顯示受管理條目數量，並可展開、摺疊及拖曳調整資料夾順序。
- 可將任一世界書條目拖放到資料夾標頭收納；也可使用 `+` 按鈕批次選取條目加入。
- 資料夾提供鉛筆按鈕以重新命名，以及垃圾桶按鈕以刪除資料夾。
- **安全保證：資料夾只保存介面整理資訊，不會影響世界書讀取。刪除資料夾時，所有受管理條目都會回到一般列表，絕不會刪除世界書條目。**
- 若使用「依目前排序重整條目 UID」，資料夾的條目關聯會自動更新。

#### 5. 世界書排序與關鍵字管理

- 管理按鈕位於世界書選擇器左側，點擊可開啟管理彈窗。
- 提供建立時間、名稱及自訂拖拉排序模式。
- 可設定優先關鍵字、角色綁定世界書優先與隱藏關鍵字（隱藏關鍵字只比對世界書名稱）。
- 支援每本世界書的多重 TAG 與快速 TAG 篩選：「全部」會先剔除隱藏書；點選某 TAG 時，擁有該 TAG 的書（含被隱藏者）仍會顯示。

### 設定選項

在 SillyTavern 的「擴充功能」設定面板中，可以啟用或停用各項功能、調整觸發檢視器快取上限、清除快取，以及獨立啟用或停用「世界書條目收納」。

## English

World Info Suite is a SillyTavern extension that provides World Info / Lorebook viewing, entry organization, batch operations, sorting, and keyword-management tools.

### Features

#### 1. Triggered Entry Viewer

- Adds a viewer button to AI messages for inspecting the World Info entries triggered by that response.
- Groups entries by insertion position and displays their source and state.
- Click a group heading to collapse or expand its entries.
- With two or more groups, a column of navigation dots appears on the left (hover for the group name); click one to smooth-scroll to that section.
- Lets you expand entries to inspect keywords, secondary filters, content, and other details.
- Keeps trigger data for the latest 10 messages by default; adjust the limit (`0` is unlimited) or clear the cache in one action.

#### 2. Character Lorebook Quick Access

- Shows a character's associated lorebooks in the character-management area.
- Distinguishes primary, extra, chat, and global sources.
- Opens the selected lorebook directly.

#### 3. Bulk Entry Editor

- Adds a bulk-edit control to the World Info editor toolbar.
- Select multiple entries and apply the same edit template to all of them.
- Supports moving or copying selected entries to another worldbook, copying multiple entry contents, and confirmed bulk deletion.
- Adds a single-entry copy control to the Content title row in the expanded native entry editor.
- Can reindex entry UIDs according to the current sort order.

#### 4. Worldbook Entry Folders

- Use the folder icon to the left of the bulk-edit button to create a named, collapsible folder.
- Folders show their managed entry count and can be expanded, collapsed, or dragged to change folder order.
- Drag any worldbook entry onto a folder header to organize it, or use the `+` button to add multiple selected entries.
- Use the pencil button to rename a folder and the trash button to remove it.
- **Safety guarantee: folders store interface-only organization data and never affect worldbook reading. Removing a folder returns all of its managed entries to the regular list; it never deletes World Info entries.**
- Folder membership is updated automatically after reindexing entry UIDs by the current sort order.

#### 5. Worldbook Sort & Keyword Manager

- Opens from the control to the left of the World Info selector.
- Provides creation-time, name, and custom drag-order sorting modes.
- Supports priority keywords, character-bound worldbook priority, and hidden keywords (matched against worldbook names only).
- Supports multiple TAGs per worldbook and quick TAG filtering: "All" removes hidden books first; picking a TAG still shows every book carrying it, including hidden ones.

### Settings

In SillyTavern's Extensions settings panel, you can enable or disable each feature, adjust the Triggered Entry Viewer cache limit, clear cached trigger data, and independently enable or disable Worldbook Entry Folders.
