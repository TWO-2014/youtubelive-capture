#!/usr/bin/env python3
"""リッチなPDFレポート: 詳細分析データ + Gemini視覚分析を統合（横A4）"""

import json
import os
import sys
import html as html_mod
import subprocess


def esc(text):
    return html_mod.escape(str(text)) if text else ""


def main():
    output_dir = sys.argv[1] if len(sys.argv) > 1 else "output/6h_full"

    with open(os.path.join(output_dir, "claude_analysis.json"), "r", encoding="utf-8") as f:
        summary = json.load(f)
    with open(os.path.join(output_dir, "chunk_details.json"), "r", encoding="utf-8") as f:
        chunks = json.load(f)
    gemini = None
    gpath = os.path.join(output_dir, "gemini_visual.json")
    if os.path.exists(gpath):
        with open(gpath, "r", encoding="utf-8") as f:
            gemini = json.load(f)

    products = summary["products"]
    patterns = summary["pattern_analysis"]
    tech_dist = patterns["technique_distribution"]
    info = summary["broadcast_info"]
    tech_jp = {"scarcity":"希少性","social_proof":"社会的証明","anchoring":"アンカリング",
               "comparison":"比較","authority":"権威","urgency":"緊急性"}
    colors = ["#4e79a7","#f28e2b","#e15759","#76b7b2","#59a14f","#edc948","#b07aa1"]
    tech_colors = {"scarcity":"#e15759","social_proof":"#4e79a7","anchoring":"#f28e2b",
                   "comparison":"#76b7b2","authority":"#59a14f","urgency":"#edc948"}
    total_tech = sum(tech_dist.values())

    pages = []

    # ===== 表紙 =====
    pages.append(f"""<div class="cover">
        <div class="cover-content">
            <h1>ショップチャンネル<br>6時間放送 分析レポート</h1>
            <div class="subtitle">セールストーク構造 &times; 番組構成 &times; CV導線 の統合分析</div>
            <div class="cover-stats">
                <div class="cs"><span class="cs-num">{len(products)}</span><span class="cs-label">商品</span></div>
                <div class="cs"><span class="cs-num">360</span><span class="cs-label">分</span></div>
                <div class="cs"><span class="cs-num">{total_tech}</span><span class="cs-label">テクニック</span></div>
                <div class="cs"><span class="cs-num">{info['total_segments']:,}</span><span class="cs-label">セグメント</span></div>
            </div>
            <div class="cover-meta">
                分析日: {info['analysis_date']} | URL: {info['url']}<br>
                キャンペーン: {info['campaign']}<br>
                分析手法: Claude Code テキスト分析 + Gemini 2.5 Flash 視覚分析
            </div>
        </div>
    </div>""")

    # ===== サマリー + タイムライン =====
    # SVGタイムライン
    bar_h = 22
    svg_bars = ""
    for i, p in enumerate(products):
        c = colors[i % len(colors)]
        s = _time_to_min(p["start_time"])
        e = _time_to_min(p["end_time"])
        x1 = s / 360 * 700; w = max((e - s) / 360 * 700, 2); y = i * (bar_h + 4) + 20
        svg_bars += f'<rect x="{x1}" y="{y}" width="{w}" height="{bar_h}" rx="3" fill="{c}" opacity="0.85"/>'
        svg_bars += f'<text x="{x1+4}" y="{y+15}" font-size="9" fill="white" font-weight="600">{esc(p["name"][:20])}</text>'
        if "second_appearance_start" in p:
            s2 = _time_to_min(p["second_appearance_start"]); e2 = _time_to_min(p["second_appearance_end"])
            x2 = s2/360*700; w2 = max((e2-s2)/360*700, 2)
            svg_bars += f'<rect x="{x2}" y="{y}" width="{w2}" height="{bar_h}" rx="3" fill="{c}" opacity="0.5"/>'
    svg_h = len(products) * (bar_h + 4) + 40
    ticks = ""
    for h in range(7):
        x = h*60/360*700
        ticks += f'<line x1="{x}" y1="10" x2="{x}" y2="{svg_h-5}" stroke="#ddd" stroke-width="0.5"/>'
        ticks += f'<text x="{x}" y="{svg_h}" font-size="8" fill="#999" text-anchor="middle">{h}h</text>'

    timeline_rows = ""
    for i, p in enumerate(products):
        dur = p.get("total_duration_minutes", p.get("duration_minutes", ""))
        timeline_rows += f'<tr><td><span class="dot" style="background:{colors[i%len(colors)]}"></span>{esc(p["name"][:28])}</td><td>{p["start_time"]}〜{p["end_time"]}</td><td style="text-align:right">{dur}分</td><td class="price">{esc(str(p.get("price",""))[:45])}</td><td>{p.get("category","")}</td></tr>'

    # フェーズバー
    phase_jp = {"intro":"導入","demo":"実演デモ","price_reveal":"価格発表","social_proof":"社会的証明","cta_scarcity":"CTA・在庫訴求","close":"クロージング"}
    phase_colors_map = {"intro":"#4e79a7","demo":"#59a14f","price_reveal":"#f28e2b","social_proof":"#76b7b2","cta_scarcity":"#e15759","close":"#edc948"}
    phase_bar = "".join(f'<div style="width:{v}%;background:{phase_colors_map.get(k,"#999")};padding:8px 6px;text-align:center;color:white;font-size:10px;font-weight:600;">{phase_jp.get(k,k)} {v}%</div>' for k, v in patterns.get("avg_phase_duration_pct", {}).items())

    # テクニック分布
    max_tech = max(tech_dist.values())
    tech_bars = ""
    for k, v in sorted(tech_dist.items(), key=lambda x: -x[1]):
        pct = round(v/total_tech*100, 1)
        bw = round(v/max_tech*200)
        tech_bars += f'<div style="display:flex;align-items:center;gap:8px;margin:3px 0;"><span style="width:80px;text-align:right;font-size:10px;font-weight:600;">{tech_jp.get(k,k)}</span><div style="background:{tech_colors.get(k,"#999")};height:16px;width:{bw}px;border-radius:3px;"></div><span style="font-size:10px;">{v} ({pct}%)</span></div>'

    pages.append(f"""<div class="page">
        <h2>エグゼクティブサマリー</h2>
        <p class="body-text">6時間の放送で<strong>7つの商品セグメント</strong>を分析。1商品あたり平均45-55分のサイクルで「導入→実演デモ(複数ラウンド)→価格発表→社会的証明→在庫カウントダウン→最終CTA→クロージング」の定型構造を持つ。最多テクニックは<strong>希少性（{tech_dist['scarcity']}件）</strong>。<strong>全員送料無料キャンペーン（4/10-16）</strong>が全セグメント共通の緊急性ドライバー。</p>
        <h2>商品タイムライン</h2>
        <svg width="720" height="{svg_h+5}" viewBox="0 0 720 {svg_h+5}">{ticks}{svg_bars}</svg>
        <table><thead><tr><th>商品名</th><th>時間帯</th><th>時間</th><th>価格</th><th>カテゴリ</th></tr></thead><tbody>{timeline_rows}</tbody></table>
        <div class="two-col" style="margin-top:10px;">
            <div class="col"><h2>フェーズ別時間配分</h2><div style="display:flex;border-radius:4px;overflow:hidden;margin:8px 0;">{phase_bar}</div></div>
            <div class="col"><h2>テクニック分布（全{total_tech}件）</h2>{tech_bars}</div>
        </div>
    </div>""")

    # ===== 頻出フレーズ + 共通パターン =====
    phrase_rows = "".join(f'<tr><td>「{esc(ph["phrase"])}」</td><td style="text-align:right">{ph["count"]}</td><td>{esc(ph.get("context",""))}</td></tr>' for ph in patterns.get("top_phrases", []))

    pages.append(f"""<div class="page">
        <h2>共通トーク構造</h2>
        <div class="highlight-box">{esc(patterns['common_structure'])}</div>
        <h2>頻出フレーズ TOP 10</h2>
        <table><thead><tr><th>フレーズ</th><th>回数</th><th>文脈</th></tr></thead><tbody>{phrase_rows}</tbody></table>
    </div>""")

    # ===== 商品別詳細ページ（各商品1-2ページ） =====
    # チャンクデータを商品にマッピング
    for pi, prod in enumerate(products):
        color = colors[pi % len(colors)]
        ts = prod.get("techniques_summary", {})
        dur = prod.get("total_duration_minutes", prod.get("duration_minutes", "?"))

        # このproductの時間帯に含まれるチャンクからテクニック実例を収集
        start_min = _time_to_min(prod["start_time"])
        end_min = _time_to_min(prod["end_time"])

        technique_examples = {"scarcity":[], "social_proof":[], "anchoring":[], "comparison":[], "authority":[], "urgency":[]}
        for chunk in chunks:
            techs = chunk.get("techniques", {})
            for tech_key in technique_examples:
                for item in techs.get(tech_key, []):
                    ts_str = item.get("timestamp", "")
                    try:
                        ts_min = int(ts_str.split(":")[0]) if ":" in ts_str else 0
                    except:
                        ts_min = 0
                    if start_min <= ts_min <= end_min + 10:
                        technique_examples[tech_key].append(item)

        # テクニック実例HTML
        tech_detail_html = ""
        for tech_key in ["scarcity", "social_proof", "anchoring", "comparison", "authority", "urgency"]:
            examples = technique_examples[tech_key][:4]  # 最大4例
            if not examples:
                continue
            quotes = "".join(f'<div class="quote"><span class="ts">[{esc(ex.get("timestamp",""))}]</span> {esc(ex.get("quote", ex.get("text","")))}</div>' for ex in examples)
            tech_detail_html += f'<div class="tech-section"><div class="tech-label" style="background:{tech_colors.get(tech_key,"#999")}">{tech_jp[tech_key]} ({len(technique_examples[tech_key])})</div>{quotes}</div>'

        # キーフレーズ
        kp_html = "".join(f"<li>{esc(kp)}</li>" for kp in prod.get("key_phrases", [])[:6])

        # フェーズ
        phases_html = "".join(f'<li><strong>{ph["type"]}</strong> ({ph["start"]}〜{ph["end"]}): {esc(ph.get("note",""))}</li>' for ph in prod.get("phases", []))

        # テクニック分布ミニバー
        max_v = max(ts.values()) if ts else 1
        mini_bars = ""
        for tk, tv in sorted(ts.items(), key=lambda x: -x[1]):
            bw = round(tv / max_v * 140)
            mini_bars += f'<div style="display:flex;align-items:center;gap:4px;margin:2px 0;"><span style="width:60px;text-align:right;font-size:9px;">{tech_jp.get(tk,tk)}</span><div style="background:{tech_colors.get(tk,"#999")};height:13px;width:{bw}px;border-radius:2px;"></div><span style="font-size:9px;">{tv}</span></div>'

        pages.append(f"""<div class="page">
            <div class="product-header" style="border-bottom:3px solid {color};">
                <h2 style="border:none;margin:0;padding:0;">{esc(prod['name'])}</h2>
                <div class="ph-meta">{esc(prod.get('brand',''))} | {prod['start_time']}〜{prod['end_time']} ({dur}分) | <span class="price">{esc(str(prod.get('price','')))}</span> | 品番: {esc(str(prod.get('item_number','N/A')))}</div>
            </div>
            <div class="two-col" style="margin-top:10px;">
                <div class="col" style="flex:2;">
                    <h3>テクニック実例（引用+タイムスタンプ）</h3>
                    {tech_detail_html}
                </div>
                <div class="col" style="flex:1;">
                    <h3>テクニック分布</h3>
                    {mini_bars}
                    <h3 style="margin-top:12px;">キーフレーズ</h3>
                    <ul class="kp-list">{kp_html}</ul>
                </div>
            </div>
            <h3 style="margin-top:10px;">フェーズ構成</h3>
            <ul class="phase-list">{phases_html}</ul>
        </div>""")

    # ===== Gemini 視覚分析ハイライト =====
    if gemini:
        visual_items = ""
        batch_count = 0
        for r in gemini.get("results", []):
            analysis_text = r.get("gemini_analysis", "")
            if len(analysis_text) > 100:  # 実質的な分析があるもののみ
                batch_count += 1
                # 分析テキストを要約的に表示（最初の800文字）
                preview = esc(analysis_text[:800])
                if len(analysis_text) > 800:
                    preview += "..."
                visual_items += f"""<div class="visual-item">
                    <div class="visual-header">
                        <span class="ts-badge">{r['timestamp_display']}</span>
                        <span class="visual-reason">{esc(r['reason'])} — {esc(r['description'])}</span>
                        <span class="visual-file">{r.get('screenshot','')}</span>
                    </div>
                    <div class="visual-body">{preview}</div>
                </div>"""
                if batch_count >= 6:  # 最大6バッチ分
                    break

        pages.append(f"""<div class="page">
            <h2>Gemini 視覚分析ハイライト（スクリーンショット分析結果）</h2>
            <p class="body-text">60枚のスクリーンショットをGemini 2.5 Flashで分析。テロップ・価格表示・QRコード・画面レイアウト等の視覚要素を抽出。以下は主要ポイントの抜粋。</p>
            {visual_items}
        </div>""")

    # ===== インサイト =====
    pages.append("""<div class="page">
        <h2>インサイト・示唆 — 他の販売チャネルに応用可能な知見</h2>
        <div class="two-col">
            <div class="col">
                <div class="insight"><div class="insight-num">1</div><div><strong>在庫カウントダウンのリアルタイム性</strong><p>「残り○○点」の連続更新が最強の購買ドライバー（84件検出）。1セグメントで10-25回更新。ECでも在庫数リアルタイム表示、「残りN点」通知が有効。</p></div></div>
                <div class="insight"><div class="insight-num">2</div><div><strong>デモの反復構造（全体の45%）</strong><p>同じデモを2-3ラウンド繰り返す設計。途中参加者への対応と、繰り返しによる購買意欲醸成の二重効果。動画コマースでも「途中から見ても分かる」構成が鍵。</p></div></div>
                <div class="insight"><div class="insight-num">3</div><div><strong>アンカリングの二重構造</strong><p>「メーカー希望小売価格→SC価格」の空間的アンカリングに加え、「原材料高騰で値上げすべきを据え置き」「30周年特別」の時間的アンカリングを組み合わせ。</p></div></div>
            </div>
            <div class="col">
                <div class="insight"><div class="insight-num">4</div><div><strong>社会的証明のライブ感（73件）</strong><p>「お電話350名」「1万件突破」のリアルタイム報告 + 顧客メッセージの即時読み上げ。バンドワゴン効果とFOMOを同時に発動させる設計。</p></div></div>
                <div class="insight"><div class="insight-num">5</div><div><strong>ゲスト出演の権威効果</strong><p>社長・デザイナー・実演販売士が全商品に登場。開発背景や職人のこだわりを「本人の口」から語る。インフルエンサーマーケティングの原型。</p></div></div>
                <div class="insight"><div class="insight-num">6</div><div><strong>送料無料の横断ドライバー設計</strong><p>期間限定キャンペーン（30回以上言及）が個別商品訴求とは別レイヤーで全セグメントを貫通。「今買う理由」を商品特性に依存せず生成する仕組み。</p></div></div>
            </div>
        </div>
        <div class="footer">Claude Code テキスト分析 + Gemini 2.5 Flash 視覚分析 | 生成日: 2026-04-10</div>
    </div>""")

    # ===== HTML組み立て =====
    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<style>
@page {{ size: A4 landscape; margin: 12mm; }}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: -apple-system, 'Hiragino Sans', 'Noto Sans JP', sans-serif; color: #333; font-size: 10px; line-height: 1.5; }}

.cover {{ background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460); color: white; margin: -12mm; height: 210mm; display: flex; align-items: center; justify-content: center; page-break-after: always; }}
.cover-content {{ text-align: center; }}
.cover h1 {{ font-size: 28px; margin-bottom: 8px; }}
.subtitle {{ font-size: 14px; opacity: 0.7; margin-bottom: 30px; }}
.cover-stats {{ display: flex; gap: 30px; justify-content: center; margin: 30px 0; }}
.cs {{ text-align: center; }}
.cs-num {{ display: block; font-size: 36px; font-weight: 700; }}
.cs-label {{ font-size: 11px; opacity: 0.6; }}
.cover-meta {{ font-size: 10px; opacity: 0.5; line-height: 1.8; margin-top: 30px; }}

.page {{ page-break-before: always; }}
h2 {{ font-size: 13px; color: #1a1a2e; margin: 10px 0 6px 0; padding-bottom: 3px; border-bottom: 2px solid #1a1a2e; }}
h3 {{ font-size: 11px; color: #444; margin: 6px 0 4px 0; }}
.body-text {{ font-size: 10px; line-height: 1.6; margin: 4px 0 8px 0; }}
table {{ width: 100%; border-collapse: collapse; font-size: 9px; margin: 4px 0; }}
th, td {{ padding: 3px 5px; text-align: left; border-bottom: 1px solid #eee; }}
th {{ background: #f0f0f0; font-weight: 600; font-size: 8px; }}
.price {{ color: #e15759; font-weight: 700; }}
.dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-right: 4px; vertical-align: middle; }}
.two-col {{ display: flex; gap: 15px; }}
.col {{ flex: 1; }}

.highlight-box {{ background: #f0f4ff; border-left: 3px solid #4e79a7; padding: 10px 12px; margin: 8px 0; font-size: 10px; line-height: 1.6; border-radius: 0 4px 4px 0; }}

.product-header {{ padding-bottom: 6px; margin-bottom: 4px; }}
.ph-meta {{ font-size: 9px; color: #666; margin-top: 3px; }}

.tech-section {{ margin-bottom: 8px; }}
.tech-label {{ display: inline-block; color: white; font-size: 9px; font-weight: 600; padding: 2px 8px; border-radius: 3px; margin-bottom: 3px; }}
.quote {{ font-size: 9px; color: #444; padding: 2px 0 2px 12px; border-left: 2px solid #ddd; margin: 2px 0; }}
.ts {{ color: #999; font-size: 8px; margin-right: 4px; }}

.kp-list {{ font-size: 9px; padding-left: 14px; margin: 2px 0; }}
.kp-list li {{ margin-bottom: 1px; }}
.phase-list {{ font-size: 8px; padding-left: 14px; margin: 2px 0; color: #555; columns: 2; }}
.phase-list li {{ margin-bottom: 1px; break-inside: avoid; }}

.visual-item {{ margin-bottom: 10px; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden; }}
.visual-header {{ background: #f5f5f5; padding: 5px 8px; font-size: 9px; display: flex; gap: 10px; align-items: center; }}
.ts-badge {{ background: #4e79a7; color: white; padding: 1px 6px; border-radius: 3px; font-size: 8px; font-weight: 600; white-space: nowrap; }}
.visual-reason {{ font-weight: 600; }}
.visual-file {{ color: #999; font-size: 8px; margin-left: auto; }}
.visual-body {{ padding: 6px 8px; font-size: 8px; color: #555; line-height: 1.5; white-space: pre-wrap; max-height: 120px; overflow: hidden; }}

.insight {{ display: flex; gap: 10px; margin-bottom: 10px; padding: 10px; background: #f5f7ff; border-radius: 6px; }}
.insight-num {{ width: 28px; height: 28px; background: #4e79a7; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 14px; flex-shrink: 0; }}
.insight p {{ font-size: 9px; color: #555; margin-top: 2px; }}
.insight strong {{ font-size: 10px; }}

.footer {{ margin-top: 15px; text-align: center; font-size: 8px; color: #999; }}
</style>
</head>
<body>
{"".join(pages)}
</body>
</html>"""

    html_path = os.path.join(output_dir, "report_for_pdf.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    pdf_path = os.path.join(output_dir, "analysis_report.pdf")
    result = subprocess.run(["weasyprint", html_path, pdf_path], capture_output=True, text=True)
    if result.returncode == 0:
        print(f"PDF出力完了: {pdf_path}")
    else:
        print(f"[ERROR] {result.stderr[:500]}")
        sys.exit(1)


def _time_to_min(t):
    parts = t.split(":")
    if len(parts) == 2:
        return int(parts[0]) + int(parts[1]) / 60
    return 0


if __name__ == "__main__":
    main()
