# JCSC (Jazz Chord-Scale Calculator)

**The Ultimate Logic-Driven Jazz Theory Engine for Excel**

JCSC is a powerful Excel-based tool designed to analyze Jazz chord progressions instantly. It calculates the theoretically correct Scales, Chord Tones, Guide Tones, and Avoid Notes based on the relationship between the Key and the Chord. It now features an interactive keyboard display to visualize chord and scale tones.

## 🚀 Key Features

### 1. Advanced Theoretical Logic

- **Context-Aware Scale Selection:** Automatically determines the best scale (e.g., Dorian, Mixolydian, Lydian Dominant, Altered) based on the harmonic function (Tonic/Sub-Dominant/Dominant) within the Key.
- **Enharmonic Correctness:** Perfectly handles complex spellings like **Double Sharps (x)** and **Triple Flats (bbb)**.
  - _Example:_ `G#M7` → Major 7th is `F##` (not G).
  - _Example:_ `Fbdim7` → Diminished 7th is `Ebbb` (not D).

### 2. Flexible Input

- **Direct Input:** Type chord symbols directly from the lead sheet (e.g., `Dm7`, `G7`).
- **Relative Input:** Supports transposed input relative to C Major / A Minor (e.g., input `Fm7` in Key C to represent IVm7).
- **Auto-Correction:** Automatically detects inconsistencies between input modes.

### 3. Comprehensive Output

- **Chord Tones:** Root, 3rd, 5th, 7th, etc.
- **Guide Tones:** 3rd & 7th (The core of Jazz lines).
- **Scale Notes:** Full scale display with visual markers for Guide Tones `'` and Avoid Notes `()`.
- **Alternative Scales:** Suggests other usable scales (e.g., Whole Tone, Diminished) for reharmonization.

### 4. Interactive Visualization

- **Dynamic Keyboard Display:** Visualizes the selected chord's tones on a piano keyboard interface.
- **Color-Coded Analysis:** Keys light up to indicate their function:
  - Root Note
  - Avoid Note
  - Guide Tone
  - Chord Tone
  - Scale Tone
- **Instant Updates:** Simply select a row and press **F9** to instantly update the keyboard display for that specific chord.

## 🛠 Requirements

- **Microsoft Excel 2021** or **Microsoft 365**
  - _Note:_ This tool heavily relies on modern dynamic array functions like `LET`, `XLOOKUP`, `TEXTSPLIT`, `LAMBDA`, and `SPILL`. It will **NOT** work on older versions of Excel.

## 📖 Usage

1.  Open `JCSC_v1.0.xlsx`.
2.  Navigate to the main table.
3.  Fill in the columns as follows:

| Column        | Description                                           | Example                             |
| :------------ | :---------------------------------------------------- | :---------------------------------- |
| **KeyIn**     | Input the Key of the song (or section).               | `C`, `Am`, `Gb`, `F#`               |
| **iChord**    | Input the chord symbol directly.                      | `Dm7`, `G7`, `Cmaj7`                |
| **iChordInC** | (Optional) Input the chord relative to Key C (or Am). | `Dm7` (when Key is Bb, meaning Cm7) |

4.  The tool will automatically calculate:

    - **oChord:** The actual sounding chord.
    - **Roman:** Degree name (e.g., `IIm7`, `V7`).
    - **ScaleName:** The recommended scale.
    - **Details:** Chord Tones, Guide Tones, and Scale Notes.

5.  **Visualize on Keyboard:**
    - Click on any row in the main table to select a chord.
    - Press **F9** to update the keyboard visualization.
    - _Note:_ **Do not copy and paste** rows to prevent broken links. **Select a row and press F9 to update.**
    - If you need an image of the keyboard, please **take a screenshot**.

## 🧠 Logical Structure

This tool handles Jazz theory using a relational database approach within Excel.

- **NoteDefinitions:** Defines all 12 notes, including enharmonic spellings (e.g., `C#` vs `Db`, `F##`, `Bbb`).
- **IntervalDefinitions:** Manages musical intervals and their distances.
- **ChordDefinitions:** Defines the structure of chord symbols (e.g., `m7` = R, m3, 5, m7).
- **ScaleDefinitions:** Defines scale intervals and avoid notes.
- **ScaleAssignmentRules:** The logic engine that maps `Key Context` + `Chord Degree` to the correct `Scale`.

## ⚠️ Known Limitations

- **Extreme Theoretical Concepts:** While it handles double/triple accidentals, chords that are theoretically impossible within a specific key context (e.g., `Gbm7` in the key of `G#`) may return logic errors. This is by design to maintain theoretical integrity.

---

---

# JCSC (Jazz Chord-Scale Calculator) - 日本語ドキュメント

**究極の論理駆動型 Excel ジャズ理論エンジン**

JCSC は、ジャズのコード進行を瞬時に解析するために設計された、強力な Excel ツールです。曲のキー（調）とコードの関係性に基づき、理論的に最も正しいスケール、コードトーン、ガイドトーン、アボイドノートを自動的に算出します。さらに、インタラクティブな鍵盤表示機能により、視覚的な理解を助けます。

## 🚀 主な機能

### 1. 高度な理論ロジック

- **文脈に応じたスケール判定:** キー内の和声機能（トニック/サブドミナント/ドミナント）に基づき、最適なスケール（ドリアン、ミクソリディアン、オルタードなど）を自動選択します。
- **異名同音の完全な処理:** **ダブルシャープ (x)** や **トリプルフラット (bbb)** といった複雑な表記も、理論通りに正確に出力します。
  - _例:_ `G#M7` → メジャー 7th は `F##` (G ではない)。
  - _例:_ `Fbdim7` → 減 7 度は `Ebbb` (D ではない)。

### 2. 柔軟な入力システム

- **直接入力:** 譜面に書かれたコードネームをそのまま入力可能（例: `Dm7`, `G7`）。
- **相対入力:** C メジャー（または A マイナー）に移調した状態での入力もサポート（アナライズ学習に最適）。
- **自動整合性チェック:** 入力モード間の矛盾を自動的に検知します。

### 3. 包括的な出力

- **コードトーン:** ルート、3 度、5 度、7 度など。
- **ガイドトーン:** ジャズのアドリブの核となる 3 度と 7 度。
- **スケール構成音:** ガイドトーンを `'` 、アボイドノートを `()` で視覚化したスケール一覧。
- **代替スケール:** リハーモナイズのヒントとなる、その他のスケール候補を提案。

### 4. インタラクティブな視覚化

- **ダイナミックキーボード表示:** 選択したコードの構成音をピアノ鍵盤インターフェース上に視覚化します。
- **カラーコード分析:** 音の機能に応じて鍵盤が色分けされます：
  - ルート音 (Root Note)
  - アボイドノート (Avoid Note)
  - ガイドトーン (Guide Tone)
  - コードトーン (Chord Tone)
  - スケールトーン (Scale Tone)
- **瞬時更新:** 行を選択して **F9** を押すだけで、そのコードに対応した鍵盤表示に瞬時に切り替わります。

## 🛠 動作環境

- **Microsoft Excel 2021** または **Microsoft 365**
  - _注意:_ `LET`, `XLOOKUP`, `TEXTSPLIT`, `SPILL` などの最新の動的配列関数を多用しています。古いバージョンの Excel では動作しません。

## 📖 使い方

1.  `JCSC_v1.0.xlsx` を開きます。
2.  メインのテーブルに移動します。
3.  以下の列を入力します：

| 列名          | 説明                                                         | 入力例                                  |
| :------------ | :----------------------------------------------------------- | :-------------------------------------- |
| **KeyIn**     | 曲（またはそのセクション）のキーを入力します。               | `C`, `Am`, `Gb`, `F#`                   |
| **iChord**    | コードネームを直接入力します。                               | `Dm7`, `G7`, `Cmaj7`                    |
| **iChordInC** | (任意) キーを C(または Am)と仮定した時のコードを入力します。 | `Dm7` (Key が Bb の時の Cm7 を表す場合) |

4.  入力すると、以下の情報が自動計算されます：

    - **oChord:** 実際に演奏される実音コード。
    - **Roman:** ディグリーネーム（度数表記）。
    - **ScaleName:** 推奨されるスケール名。
    - **詳細情報:** コードトーン、ガイドトーン、スケール構成音。

5.  **鍵盤で確認する:**
    - メインテーブルの任意の行（コード）をクリックして選択します。
    - **F9** キーを押すと、鍵盤表示が更新されます。
    - _注意:_ リンク切れを防ぐため、**行のコピー＆ペーストはしないでください。行を選択し、F9 を押して更新してください。**
    - 鍵盤の画像が必要な場合は、**スクリーンショットを撮ってください。**

## 🧠 論理構造

本ツールは、Excel 内部でリレーショナルデータベース的なアプローチを用いてジャズ理論を処理しています。

- **NoteDefinitions:** 12 音の定義（異名同音を含む `C#` vs `Db`, `F##`, `Bbb` など）。
- **IntervalDefinitions:** 音程と距離の管理。
- **ChordDefinitions:** コードシンボルの構造定義 (例: `m7` = R, m3, 5, m7)。
- **ScaleDefinitions:** スケールの音程構成とアボイドノートの定義。
- **ScaleAssignmentRules:** 「キーの文脈」＋「コードの度数」を正しい「スケール」にマッピングするロジックエンジン。

## ⚠️ 制限事項

- **極端な理論的ケース:** ダブル/トリプルシャープ等に対応していますが、特定のキーにおいて理論上定義不能なコード（例: `G#`キーにおける `Gbm7`）は、理論的整合性を保つためにエラーを返します。これは仕様です。

---

---

## License

MIT License

Copyright (c) 2025 GoodRelax

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
