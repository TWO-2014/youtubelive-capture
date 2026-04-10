#!/usr/bin/env python3
"""Phase B: claude_analysis.jsonの視覚分析ターゲットに基づき、Geminiでスクショを分析"""

import json
import os
import sys
import glob
import time

def find_nearest_screenshot(timestamp_sec, ss_dir):
    """タイムスタンプに最も近いスクリーンショットを返す"""
    # 60秒間隔なので最寄りのフレームを選択
    nearest_sec = round(timestamp_sec / 60) * 60
    pattern = os.path.join(ss_dir, f"frame_*_{nearest_sec}s.jpg")
    matches = glob.glob(pattern)
    if matches:
        return matches[0]
    # フォールバック: 前後を探す
    for delta in [60, -60, 120, -120]:
        pattern = os.path.join(ss_dir, f"frame_*_{nearest_sec + delta}s.jpg")
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
    return None


def main():
    output_dir = sys.argv[1] if len(sys.argv) > 1 else "output/6h_full"
    api_key = os.environ.get("GEMINI_API_KEY") or (sys.argv[2] if len(sys.argv) > 2 else None)
    if not api_key:
        print("[ERROR] GEMINI_API_KEY が必要です")
        sys.exit(1)

    analysis_path = os.path.join(output_dir, "claude_analysis.json")
    ss_dir = os.path.join(output_dir, "screenshots")

    with open(analysis_path, "r", encoding="utf-8") as f:
        analysis = json.load(f)

    targets = analysis["visual_analysis_targets"]
    print(f"視覚分析ターゲット: {len(targets)} 箇所")

    # スクリーンショットをマッピング
    target_screenshots = []
    for t in targets:
        ss_path = find_nearest_screenshot(t["timestamp_sec"], ss_dir)
        if ss_path:
            target_screenshots.append({
                "timestamp_sec": t["timestamp_sec"],
                "reason": t["reason"],
                "description": t.get("description", ""),
                "screenshot": ss_path,
            })
    print(f"有効なスクリーンショット: {len(target_screenshots)} 枚")

    # Gemini API
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=api_key)

    prompt_template = """この通販番組（ショップチャンネル）のスクリーンショットから、以下の視覚情報のみを抽出してください。

分析理由: {reason}

1. テロップテキスト（価格、商品名、キャッチコピー、電話番号、品番）
2. 画面レイアウト（テロップの配置、サイズ、色使い）
3. QRコード・電話番号の有無と表示位置
4. 商品の見せ方（実演中/静止画/比較表示/装着デモ）
5. 価格表示の演出（打消し線、赤字、フラッシュ、割引率表示）
6. カウントダウン/残数表示の有無

テキスト内容の分析は不要です。画面に映っている視覚要素のみを構造的に記述してください。
JSON形式で出力してください。"""

    # 5枚ずつバッチ処理
    batch_size = 5
    results = []
    total_batches = (len(target_screenshots) + batch_size - 1) // batch_size

    for batch_idx in range(0, len(target_screenshots), batch_size):
        batch = target_screenshots[batch_idx:batch_idx + batch_size]
        batch_num = batch_idx // batch_size + 1
        print(f"\n  バッチ {batch_num}/{total_batches} ({len(batch)}枚)...")

        contents = []
        for item in batch:
            with open(item["screenshot"], "rb") as f:
                img_data = f.read()
            contents.append(types.Part.from_bytes(data=img_data, mime_type="image/jpeg"))
            ts_min = item["timestamp_sec"] // 60
            ts_sec = item["timestamp_sec"] % 60
            contents.append(f"[{ts_min:02d}:{ts_sec:02d}] {item['reason']}: {item['description']}")

        contents.append(prompt_template.format(reason="上記の各スクリーンショットについて分析"))

        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=contents,
            )
            batch_result = response.text
        except Exception as e:
            print(f"    [ERROR] {e}")
            # レート制限の場合は待機してリトライ
            if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e):
                print("    60秒待機してリトライ...")
                time.sleep(60)
                try:
                    response = client.models.generate_content(
                        model="gemini-2.5-flash",
                        contents=contents,
                    )
                    batch_result = response.text
                except Exception as e2:
                    print(f"    [ERROR] リトライ失敗: {e2}")
                    batch_result = f"ERROR: {e2}"
            else:
                batch_result = f"ERROR: {e}"

        for item in batch:
            ts_min = item["timestamp_sec"] // 60
            results.append({
                "timestamp_sec": item["timestamp_sec"],
                "timestamp_display": f"{ts_min}分{item['timestamp_sec'] % 60}秒",
                "reason": item["reason"],
                "description": item["description"],
                "screenshot": os.path.basename(item["screenshot"]),
            })

        # バッチ結果をまとめて格納
        results[-len(batch)]["gemini_analysis"] = batch_result
        for i, item in enumerate(batch[1:], 1):
            results[-(len(batch) - i)]["gemini_analysis"] = "(上記バッチ分析に含まれる)"

        print(f"    完了")
        # レート制限対策
        if batch_num < total_batches:
            time.sleep(5)

    # 出力
    output_path = os.path.join(output_dir, "gemini_visual.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump({
            "total_analyzed": len(results),
            "total_batches": total_batches,
            "results": results,
        }, f, ensure_ascii=False, indent=2)

    print(f"\n完了: {output_path} ({len(results)}件)")


if __name__ == "__main__":
    main()
