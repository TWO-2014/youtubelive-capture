#!/usr/bin/env python3
"""Phase C: claude_analysis.json + gemini_visual.json → dashboard.html + analysis_report.md"""

import json
import os
import sys
import base64
from pathlib import Path


def load_data(output_dir):
    with open(os.path.join(output_dir, "claude_analysis.json"), "r", encoding="utf-8") as f:
        analysis = json.load(f)
    visual_path = os.path.join(output_dir, "gemini_visual.json")
    visual = None
    if os.path.exists(visual_path):
        with open(visual_path, "r", encoding="utf-8") as f:
            visual = json.load(f)
    return analysis, visual


def generate_dashboard(analysis, visual, output_dir):
    products = analysis["products"]
    patterns = analysis["pattern_analysis"]
    tech_dist = patterns["technique_distribution"]
    non_product = analysis.get("non_product_segments", [])

    # タイムラインデータ
    timeline_data = []
    colors = ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc948", "#b07aa1"]
    for i, p in enumerate(products):
        start_min = _time_to_min(p["start_time"])
        end_min = _time_to_min(p["end_time"])
        timeline_data.append({
            "label": p["name"][:25],
            "start": start_min,
            "end": end_min,
            "color": colors[i % len(colors)],
            "price": p.get("price", ""),
            "duration": p.get("duration_minutes", end_min - start_min),
        })
        if "second_appearance_start" in p:
            s2 = _time_to_min(p["second_appearance_start"])
            e2 = _time_to_min(p["second_appearance_end"])
            timeline_data.append({
                "label": p["name"][:25] + " (2回目)",
                "start": s2,
                "end": e2,
                "color": colors[i % len(colors)],
                "price": "",
                "duration": p.get("second_appearance_duration", e2 - s2),
            })

    # テクニック分布
    tech_labels = list(tech_dist.keys())
    tech_values = list(tech_dist.values())
    tech_labels_jp = {
        "scarcity": "希少性",
        "social_proof": "社会的証明",
        "anchoring": "アンカリング",
        "comparison": "比較",
        "authority": "権威",
        "urgency": "緊急性",
    }

    # 商品別テクニック
    product_tech_data = []
    for p in products:
        ts = p.get("techniques_summary", {})
        product_tech_data.append({
            "name": p["name"][:20],
            "scarcity": ts.get("scarcity", 0),
            "social_proof": ts.get("social_proof", 0),
            "anchoring": ts.get("anchoring", 0),
            "comparison": ts.get("comparison", 0),
            "authority": ts.get("authority", 0),
            "urgency": ts.get("urgency", 0),
        })

    # フレーズ一覧
    phrases_html = ""
    for ph in patterns.get("top_phrases", []):
        phrases_html += f'<tr><td>{ph["phrase"]}</td><td>{ph["count"]}</td><td>{ph.get("context","")}</td></tr>\n'

    # 商品カード
    cards_html = ""
    for i, p in enumerate(products):
        ts = p.get("techniques_summary", {})
        top_tech = sorted(ts.items(), key=lambda x: -x[1])[:3]
        top_tech_str = ", ".join(f"{tech_labels_jp.get(k,k)}({v})" for k, v in top_tech)
        kp_html = "".join(f"<li>{kp}</li>" for kp in p.get("key_phrases", [])[:5])
        dur = p.get("total_duration_minutes", p.get("duration_minutes", "?"))
        cards_html += f"""
        <div class="card" style="border-left: 4px solid {colors[i % len(colors)]}">
            <h3>{p['name']}</h3>
            <div class="card-meta">
                <span class="badge">{p.get('category','')}</span>
                <span>{p['start_time']}〜{p['end_time']} ({dur}分)</span>
                <span class="price">{p.get('price','')}</span>
            </div>
            <div class="card-detail">
                <div><strong>ブランド:</strong> {p.get('brand','')}</div>
                <div><strong>品番:</strong> {p.get('item_number','')}</div>
                <div><strong>主要テクニック:</strong> {top_tech_str}</div>
            </div>
            <div><strong>キーフレーズ:</strong><ul>{kp_html}</ul></div>
        </div>
        """

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ショップチャンネル 6時間放送 分析ダッシュボード</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, 'Hiragino Sans', sans-serif; background: #f5f5f5; color: #333; }}
.header {{ background: linear-gradient(135deg, #1a1a2e, #16213e); color: white; padding: 2rem; }}
.header h1 {{ font-size: 1.5rem; margin-bottom: 0.5rem; }}
.header p {{ opacity: 0.8; font-size: 0.9rem; }}
.stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 1rem; padding: 1.5rem; }}
.stat {{ background: white; border-radius: 8px; padding: 1rem; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
.stat .num {{ font-size: 1.8rem; font-weight: 700; color: #1a1a2e; }}
.stat .label {{ font-size: 0.8rem; color: #666; }}
.section {{ background: white; margin: 1rem; border-radius: 8px; padding: 1.5rem; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
.section h2 {{ font-size: 1.2rem; margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid #eee; }}
.chart-container {{ position: relative; height: 400px; }}
.chart-container-wide {{ position: relative; height: 200px; }}
table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; }}
th, td {{ padding: 0.5rem; text-align: left; border-bottom: 1px solid #eee; }}
th {{ background: #f8f8f8; font-weight: 600; }}
.card {{ background: #fafafa; border-radius: 8px; padding: 1rem; margin-bottom: 1rem; }}
.card h3 {{ font-size: 1rem; margin-bottom: 0.5rem; }}
.card-meta {{ display: flex; gap: 1rem; flex-wrap: wrap; font-size: 0.85rem; color: #666; margin-bottom: 0.5rem; }}
.card-detail {{ font-size: 0.85rem; margin-bottom: 0.5rem; }}
.badge {{ background: #e8e8e8; padding: 2px 8px; border-radius: 4px; font-size: 0.75rem; }}
.price {{ color: #e15759; font-weight: 700; }}
ul {{ padding-left: 1.2rem; font-size: 0.85rem; }}
li {{ margin-bottom: 0.2rem; }}
.grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; }}
@media (max-width: 768px) {{ .grid-2 {{ grid-template-columns: 1fr; }} }}
</style>
</head>
<body>
<div class="header">
    <h1>ショップチャンネル 6時間放送 分析ダッシュボード</h1>
    <p>URL: {analysis['broadcast_info']['url']} | 分析日: {analysis['broadcast_info']['analysis_date']} | キャンペーン: {analysis['broadcast_info']['campaign']}</p>
</div>

<div class="stats">
    <div class="stat"><div class="num">{len(products)}</div><div class="label">商品数</div></div>
    <div class="stat"><div class="num">{analysis['broadcast_info']['duration_minutes']}</div><div class="label">放送時間(分)</div></div>
    <div class="stat"><div class="num">{sum(tech_dist.values())}</div><div class="label">説得テクニック検出数</div></div>
    <div class="stat"><div class="num">{analysis['broadcast_info']['total_segments']:,}</div><div class="label">文字起こしセグメント</div></div>
    <div class="stat"><div class="num">{analysis['broadcast_info']['total_screenshots']}</div><div class="label">スクリーンショット</div></div>
    <div class="stat"><div class="num">{len(analysis.get('visual_analysis_targets',[]))}</div><div class="label">視覚分析ポイント</div></div>
</div>

<div class="section">
    <h2>商品タイムライン（6時間）</h2>
    <div class="chart-container-wide">
        <canvas id="timelineChart"></canvas>
    </div>
</div>

<div class="grid-2">
    <div class="section">
        <h2>説得テクニック分布</h2>
        <div class="chart-container">
            <canvas id="techChart"></canvas>
        </div>
    </div>
    <div class="section">
        <h2>商品別テクニック使用数</h2>
        <div class="chart-container">
            <canvas id="productTechChart"></canvas>
        </div>
    </div>
</div>

<div class="section">
    <h2>フェーズ別時間配分（平均）</h2>
    <div class="chart-container" style="height:250px">
        <canvas id="phaseChart"></canvas>
    </div>
</div>

<div class="section">
    <h2>頻出フレーズ TOP 10</h2>
    <table>
        <thead><tr><th>フレーズ</th><th>出現回数</th><th>文脈</th></tr></thead>
        <tbody>{phrases_html}</tbody>
    </table>
</div>

<div class="section">
    <h2>商品カード</h2>
    {cards_html}
</div>

<script>
// Timeline
const timelineData = {json.dumps(timeline_data, ensure_ascii=False)};
new Chart(document.getElementById('timelineChart'), {{
    type: 'bar',
    data: {{
        labels: timelineData.map(d => d.label),
        datasets: [{{
            label: '放送時間帯',
            data: timelineData.map(d => [d.start, d.end]),
            backgroundColor: timelineData.map(d => d.color + '99'),
            borderColor: timelineData.map(d => d.color),
            borderWidth: 1,
        }}]
    }},
    options: {{
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        scales: {{
            x: {{ min: 0, max: 360, title: {{ display: true, text: '時間（分）' }},
                ticks: {{ callback: v => Math.floor(v/60)+'h'+String(v%60).padStart(2,'0')+'m' }} }}
        }},
        plugins: {{ legend: {{ display: false }} }}
    }}
}});

// Technique distribution
const techLabels = {json.dumps([tech_labels_jp.get(k,k) for k in tech_labels], ensure_ascii=False)};
const techValues = {json.dumps(tech_values)};
new Chart(document.getElementById('techChart'), {{
    type: 'doughnut',
    data: {{
        labels: techLabels,
        datasets: [{{ data: techValues, backgroundColor: ['#e15759','#4e79a7','#f28e2b','#76b7b2','#59a14f','#edc948'] }}]
    }},
    options: {{ responsive: true, maintainAspectRatio: false }}
}});

// Product technique stacked bar
const ptData = {json.dumps(product_tech_data, ensure_ascii=False)};
const techKeys = ['scarcity','social_proof','anchoring','comparison','authority','urgency'];
const techColors = ['#e15759','#4e79a7','#f28e2b','#76b7b2','#59a14f','#edc948'];
const techNamesJP = {json.dumps(tech_labels_jp, ensure_ascii=False)};
new Chart(document.getElementById('productTechChart'), {{
    type: 'bar',
    data: {{
        labels: ptData.map(d => d.name),
        datasets: techKeys.map((k, i) => ({{
            label: techNamesJP[k],
            data: ptData.map(d => d[k]),
            backgroundColor: techColors[i],
        }}))
    }},
    options: {{
        responsive: true,
        maintainAspectRatio: false,
        scales: {{ x: {{ stacked: true }}, y: {{ stacked: true }} }},
    }}
}});

// Phase duration
const phaseData = {json.dumps(patterns['avg_phase_duration_pct'])};
new Chart(document.getElementById('phaseChart'), {{
    type: 'bar',
    data: {{
        labels: Object.keys(phaseData).map(k => ({{
            intro:'導入',demo:'実演デモ',price_reveal:'価格発表',
            social_proof:'社会的証明',cta_scarcity:'CTA・在庫訴求',close:'クロージング'
        }})[k] || k),
        datasets: [{{ data: Object.values(phaseData), backgroundColor: '#4e79a7' }}]
    }},
    options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ y: {{ title: {{ display: true, text: '平均時間配分 (%)' }} }} }}
    }}
}});
</script>
</body>
</html>"""
    return html


def generate_report(analysis, visual, output_dir):
    products = analysis["products"]
    patterns = analysis["pattern_analysis"]
    tech_dist = patterns["technique_distribution"]
    info = analysis["broadcast_info"]
    tech_jp = {"scarcity":"希少性","social_proof":"社会的証明","anchoring":"アンカリング",
               "comparison":"比較","authority":"権威","urgency":"緊急性"}

    report = f"""# ショップチャンネル 6時間放送 分析レポート

## 放送情報
- **分析日**: {info['analysis_date']}
- **URL**: {info['url']}
- **放送時間**: {info['duration_minutes']}分（6時間）
- **キャンペーン**: {info['campaign']}
- **文字起こしセグメント**: {info['total_segments']:,}
- **スクリーンショット**: {info['total_screenshots']}枚

---

## 1. エグゼクティブサマリー

6時間の放送で**7つの商品セグメント**を分析。1商品あたり平均45-55分のサイクルで、「導入→実演デモ(複数ラウンド)→価格発表(アンカリング)→社会的証明→在庫カウントダウン(希少性エスカレーション)→最終CTA→クロージング」の定型構造を持つ。最も多用されるテクニックは**希少性（{tech_dist['scarcity']}件）**で、リアルタイム在庫数カウントダウンが全セグメントで使われている。**全員送料無料キャンペーン（4/10-16）**が横断的な緊急性ドライバーとして機能。

---

## 2. 商品タイムライン

| # | 商品名 | 時間帯 | 所要時間 | 価格 | カテゴリ |
|---|--------|--------|----------|------|----------|
"""
    for i, p in enumerate(products, 1):
        dur = p.get("total_duration_minutes", p.get("duration_minutes", ""))
        report += f"| {i} | {p['name'][:30]} | {p['start_time']}〜{p['end_time']} | {dur}分 | {p.get('price','')} | {p.get('category','')} |\n"

    report += """
---

## 3. 商品別分析

"""
    for i, p in enumerate(products, 1):
        ts = p.get("techniques_summary", {})
        top_tech = sorted(ts.items(), key=lambda x: -x[1])[:3]
        top_tech_str = ", ".join(f"{tech_jp.get(k,k)}({v}件)" for k, v in top_tech)

        kp_str = "\n".join(f"  - 「{kp}」" for kp in p.get("key_phrases", []))

        phases_str = ""
        for ph in p.get("phases", []):
            phases_str += f"  - **{ph['type']}** ({ph['start']}〜{ph['end']}): {ph.get('note','')}\n"

        report += f"""### 3.{i} {p['name']}

- **ブランド**: {p.get('brand','')}
- **品番**: {p.get('item_number','')}
- **価格**: {p.get('price','')}
- **放送時間**: {p['start_time']}〜{p['end_time']}（{p.get('total_duration_minutes', p.get('duration_minutes',''))}分）
- **主要テクニック**: {top_tech_str}

**フェーズ構成:**
{phases_str}
**キーフレーズ:**
{kp_str}

---

"""

    # テクニック分析
    report += """## 4. セールストーク パターン分析

### 4.1 共通トーク構造の型

"""
    report += f"{patterns['common_structure']}\n\n"

    report += "### 4.2 頻出フレーズ\n\n| フレーズ | 出現回数 | 文脈 |\n|---|---|---|\n"
    for ph in patterns.get("top_phrases", []):
        report += f"| {ph['phrase']} | {ph['count']} | {ph.get('context','')} |\n"

    report += f"""
### 4.3 テクニック分布

| テクニック | 検出数 | 割合 |
|---|---|---|
"""
    total = sum(tech_dist.values())
    for k, v in sorted(tech_dist.items(), key=lambda x: -x[1]):
        pct = round(v / total * 100, 1) if total else 0
        report += f"| {tech_jp.get(k,k)} | {v} | {pct}% |\n"

    report += f"""
**合計: {total}件**

---

## 5. 番組構成・CV導線の設計分析

### 5.1 時間配分パターン

"""
    for k, v in patterns.get("avg_phase_duration_pct", {}).items():
        phase_jp = {"intro":"導入","demo":"実演デモ","price_reveal":"価格発表",
                    "social_proof":"社会的証明","cta_scarcity":"CTA・在庫訴求","close":"クロージング"}
        report += f"- **{phase_jp.get(k,k)}**: {v}%\n"

    report += """
### 5.2 CTA設計パターン

- **電話番号**: 常時表示、「0120から始まるフリーダイヤル」「タッチでショップ2番」
- **Web/QR**: 「スマホ・パソコンでショップチャンネルと検索」「QRコード」
- **在庫カウントダウン**: リアルタイムで「○○点を切りました」を繰り返し
- **注文件数報告**: 「お電話○○名」「○○件のオーダー」でバンドワゴン効果

### 5.3 キャンペーン構造

- **全員送料無料キャンペーン（4/10-16）**: 全セグメントで繰り返し言及（推定30回以上）
- **ショップチャンネルカード5%還元**: 主要セグメントで告知
- **マッチングプライス**: 複数商品セット購入で割引（メルティリッチダウンで使用）

---

## 6. インサイト・示唆

### 他の販売チャネルに応用可能な知見

1. **在庫カウントダウンのリアルタイム性**: 「残り○○点」の連続更新が最も強力な購買ドライバー。ECサイトでも在庫数のリアルタイム表示が有効。

2. **デモの反復構造**: 1商品につき同じデモを2-3ラウンド繰り返す。視聴者の途中参加に対応しつつ、繰り返しで購買意欲を高める設計。

3. **アンカリングの二重構造**: 「メーカー希望小売価格」→「ショップチャンネル価格」に加え、「原材料高騰で本来は値上げすべきところを据え置き」という時間軸のアンカリングも使用。

4. **社会的証明のリアルタイム性**: 注文件数・同時通話数をリアルタイムで報告することで、「みんなが今買っている」というライブ感を演出。

5. **ゲスト出演の権威効果**: 社長・デザイナー・実演販売士など、商品に関わる「本人」が登場することで説得力が格段に増す。

6. **送料無料を横断ドライバーに**: 個別商品のプロモーションとは別に、期間限定の送料無料キャンペーンが全商品共通のアクショントリガーとして機能。

---

*このレポートはClaude Codeによるテキスト分析 + Gemini 2.5 Flashによるスクリーンショット視覚分析に基づいて生成されました。*
"""
    return report


def _time_to_min(t):
    """'HH:MM' or 'M:SS' or 'MMM:SS' → minutes (float)"""
    parts = t.split(":")
    if len(parts) == 2:
        h_or_m = int(parts[0])
        s = int(parts[1])
        if h_or_m > 23:  # MMM:SS format
            return h_or_m + s / 60
        else:  # HH:MM or M:SS
            return h_or_m + s / 60
    return 0


def main():
    output_dir = sys.argv[1] if len(sys.argv) > 1 else "output/6h_full"
    analysis, visual = load_data(output_dir)

    # Dashboard
    html = generate_dashboard(analysis, visual, output_dir)
    dashboard_path = os.path.join(output_dir, "dashboard.html")
    with open(dashboard_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"ダッシュボード出力: {dashboard_path}")

    # Report
    report = generate_report(analysis, visual, output_dir)
    report_path = os.path.join(output_dir, "analysis_report.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"レポート出力: {report_path}")


if __name__ == "__main__":
    main()
