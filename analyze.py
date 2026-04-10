#!/usr/bin/env python3
"""ショップチャンネル YouTube Live 分析パイプライン CLI"""

import argparse
import json
import subprocess
import sys
import os
import glob
import re
from datetime import datetime
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser(description="YouTube Live 分析パイプライン")
    parser.add_argument("--url", required=True, help="YouTubeライブ配信URL")
    parser.add_argument("--duration", type=int, required=True, help="録画時間（分）")
    parser.add_argument("--start", type=int, default=0, help="開始位置（分）")
    parser.add_argument("--screenshot-interval", type=int, default=60, help="スクリーンショット間隔（秒）")
    parser.add_argument("--segment-duration", type=int, default=10, help="セグメント分割単位（分）")
    parser.add_argument("--output", default=None, help="出力ディレクトリ")
    parser.add_argument("--whisper-model", default="large-v3", help="Whisperモデルサイズ")
    parser.add_argument("--skip-analysis", action="store_true", help="文字起こしまでで止める")
    parser.add_argument("--gemini-api-key", default=None, help="Gemini APIキー（未指定時は環境変数GEMINI_API_KEY）")
    parser.add_argument("--resume-from", default=None, help="既存の出力ディレクトリからStep3以降を再実行")
    return parser.parse_args()


def run_cmd(cmd, desc=""):
    """コマンド実行ヘルパー"""
    print(f"  >> {desc or ' '.join(cmd[:3])}...")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  [ERROR] {result.stderr[:500]}")
        return False, result.stderr
    return True, result.stdout


def download_segment(url, start_sec, end_sec, output_path):
    """セグメント単位で映像+音声をダウンロード"""
    section = f"*{start_sec}-{end_sec}"
    cmd = [
        "yt-dlp",
        "-f", "bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/best[height<=720]",
        "--download-sections", section,
        "--merge-output-format", "mp4",
        "-o", output_path,
        "--no-playlist",
        url,
    ]
    ok, _ = run_cmd(cmd, f"yt-dlp で {start_sec//60}分〜{end_sec//60}分 を取得")
    if not ok:
        cmd_fb = [
            "yt-dlp",
            "-f", "best[height<=720]",
            "--download-sections", section,
            "-o", output_path,
            "--no-playlist",
            url,
        ]
        ok, _ = run_cmd(cmd_fb, f"yt-dlp (fallback) で {start_sec//60}分〜{end_sec//60}分 を取得")
    if ok and os.path.exists(output_path):
        return output_path
    # yt-dlp が拡張子を変える場合
    base = os.path.splitext(output_path)[0]
    candidates = glob.glob(f"{base}.*")
    if candidates:
        return candidates[0]
    return None


def step1_process(url, start_min, duration_min, segment_min, screenshot_interval, whisper_model, output_dir):
    """セグメント単位で DL → スクショ → 文字起こし → 削除 を繰り返す"""
    start_sec = start_min * 60
    end_sec = (start_min + duration_min) * 60
    segment_sec = segment_min * 60
    num_segments = (duration_min + segment_min - 1) // segment_min

    ss_dir = os.path.join(output_dir, "screenshots")
    tmp_dir = os.path.join(output_dir, "tmp")
    os.makedirs(ss_dir, exist_ok=True)
    os.makedirs(tmp_dir, exist_ok=True)

    print(f"\n[Step 1/3] セグメント単位で DL + スクショ + 文字起こし")
    print(f"  対象: {start_min}分〜{start_min + duration_min}分（{num_segments}セグメント × {segment_min}分）")

    # Whisperモデルを一度だけロード
    print(f"  Whisper ({whisper_model}) モデルをロード中...")
    import whisper
    model = whisper.load_model(whisper_model)

    all_segments = []
    all_screenshots = []
    transcript_path = os.path.join(output_dir, "transcript.json")
    text_path = os.path.join(output_dir, "transcript.txt")

    for idx in range(num_segments):
        seg_start = start_sec + idx * segment_sec
        seg_end = min(seg_start + segment_sec, end_sec)
        seg_start_min = seg_start // 60
        seg_end_min = seg_end // 60
        seg_label = f"[{idx+1}/{num_segments}] {seg_start_min}分〜{seg_end_min}分"
        print(f"\n  === セグメント {seg_label} ===")

        # (1) DL
        seg_video = os.path.join(tmp_dir, f"seg_{idx:03d}.mp4")
        actual_path = download_segment(url, seg_start, seg_end, seg_video)
        if not actual_path:
            print(f"    [WARN] DL失敗、スキップ")
            continue

        # (2) スクリーンショット抽出
        fps_filter = f"fps=1/{screenshot_interval}"
        cmd_ss = [
            "ffmpeg", "-i", actual_path,
            "-vf", fps_filter,
            os.path.join(ss_dir, f"seg{idx:03d}_frame_%04d.jpg"),
            "-y",
        ]
        ok, _ = run_cmd(cmd_ss, "スクリーンショット抽出")
        if ok:
            frames = sorted(glob.glob(os.path.join(ss_dir, f"seg{idx:03d}_frame_*.jpg")))
            for fi, fpath in enumerate(frames):
                global_sec = seg_start + fi * screenshot_interval
                new_name = os.path.join(ss_dir, f"frame_{len(all_screenshots):04d}_{global_sec}s.jpg")
                os.rename(fpath, new_name)
                all_screenshots.append(new_name)
            print(f"    {len(frames)} 枚のスクリーンショットを取得")

        # (3) 音声抽出
        wav_path = os.path.join(tmp_dir, f"seg_{idx:03d}.wav")
        cmd_wav = [
            "ffmpeg", "-i", actual_path,
            "-ar", "16000", "-ac", "1",
            "-vn", wav_path,
            "-y",
        ]
        ok, _ = run_cmd(cmd_wav, "音声抽出 (WAV 16kHz mono)")
        if not ok:
            print(f"    [WARN] 音声抽出失敗、スキップ")
            os.remove(actual_path)
            continue

        # (4) Whisper 文字起こし
        print(f"    Whisper 文字起こし中...")
        result = model.transcribe(wav_path, language="ja", verbose=False)
        seg_count = 0
        for seg in result["segments"]:
            all_segments.append({
                "start": seg["start"] + seg_start,
                "end": seg["end"] + seg_start,
                "text": seg["text"].strip(),
            })
            seg_count += 1
        print(f"    {seg_count} セグメントの文字起こし完了")

        # (5) 一時ファイル削除
        os.remove(actual_path)
        os.remove(wav_path)

        # (6) トランスクリプトを逐次保存（途中で落ちてもここまでのデータは残る）
        transcript = {"language": "ja", "segments": all_segments}
        with open(transcript_path, "w", encoding="utf-8") as f:
            json.dump(transcript, f, ensure_ascii=False, indent=2)
        with open(text_path, "w", encoding="utf-8") as f:
            for s in all_segments:
                m, sec = divmod(int(s["start"]), 60)
                f.write(f"[{m:02d}:{sec:02d}] {s['text']}\n")
        print(f"    保存済み（累計 {len(all_segments)} セグメント）")

    # tmp_dir クリーンアップ
    import shutil
    shutil.rmtree(tmp_dir, ignore_errors=True)

    total = len(all_segments)
    print(f"\n  処理完了: 全{total}セグメント, スクリーンショット{len(all_screenshots)}枚")
    return transcript, all_screenshots


def step2_analyze(transcript, screenshots, output_dir, args):
    """Gemini API でマルチモーダル構造分析"""
    print("\n[Step 2/3] Gemini API で構造分析中...")

    from google import genai
    from google.genai import types

    api_key = args.gemini_api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("  [ERROR] Gemini APIキーが設定されていません。--gemini-api-key または GEMINI_API_KEY 環境変数を設定してください。")
        sys.exit(1)

    client = genai.Client(api_key=api_key)

    # チャンク分割（5分ごと）
    chunk_duration = 300
    segments = transcript["segments"]
    total_duration = segments[-1]["end"] if segments else 0
    chunks = []
    chunk_start = 0
    while chunk_start < total_duration:
        chunk_end = chunk_start + chunk_duration
        chunk_segs = [s for s in segments if s["start"] >= chunk_start and s["start"] < chunk_end]
        chunk_ss = []
        for ss in screenshots:
            match = re.search(r"_(\d+)s\.jpg$", ss)
            if match:
                ss_time = int(match.group(1))
                if chunk_start <= ss_time < chunk_end:
                    chunk_ss.append(ss)
        chunks.append({
            "start": chunk_start,
            "end": min(chunk_end, total_duration),
            "segments": chunk_segs,
            "screenshots": chunk_ss[:5],
        })
        chunk_start = chunk_end

    print(f"  {len(chunks)} チャンクに分割して分析します")

    chunk_analyses = []
    for i, chunk in enumerate(chunks):
        m_start = int(chunk["start"]) // 60
        m_end = int(chunk["end"]) // 60
        print(f"  チャンク {i+1}/{len(chunks)} ({m_start}分〜{m_end}分) を分析中...")

        chunk_text = "\n".join(
            f"[{int(s['start'])//60:02d}:{int(s['start'])%60:02d}] {s['text']}"
            for s in chunk["segments"]
        )

        contents = []
        for ss_path in chunk["screenshots"]:
            with open(ss_path, "rb") as f:
                img_data = f.read()
            contents.append(types.Part.from_bytes(data=img_data, mime_type="image/jpeg"))

        contents.append(f"""以下は通販番組（ショップチャンネル）のYouTubeライブ配信の{m_start}分〜{m_end}分の区間です。
添付画像はこの区間のスクリーンショットです。

【文字起こしテキスト】
{chunk_text}

以下の観点で分析してください:

## セールストークの話術構造
- トーク構造パターン（問題提起→共感→解決策提示→実演→限定感→CTAなどの流れ）
- キーフレーズ（繰り返し使われる訴求ワード）
- 社会的証明の使い方
- 比較・アンカリング

## 番組構成・CV導線
- 時間配分（商品紹介・デモ・価格発表・クロージングの各フェーズ）
- テロップ・画面演出（スクリーンショットから読み取れる内容）
- CTAの設計（電話番号・QRコード表示のタイミング）
- 繰り返し構造

簡潔かつ構造的に分析結果を出力してください。""")

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=contents,
        )

        chunk_analyses.append({
            "chunk_index": i,
            "time_range": f"{m_start}分〜{m_end}分",
            "analysis": response.text,
        })
        print(f"    完了")

    # 統合分析
    print("  全体統合分析を実行中...")
    all_analyses = "\n\n---\n\n".join(
        f"### {ca['time_range']}\n{ca['analysis']}" for ca in chunk_analyses
    )

    full_text = "\n".join(
        f"[{int(s['start'])//60:02d}:{int(s['start'])%60:02d}] {s['text']}"
        for s in segments
    )

    integration_response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=f"""以下はショップチャンネル（通販番組）のYouTubeライブ配信の各チャンク分析結果です。
これらを統合し、番組全体の構造分析レポートを作成してください。

【チャンク別分析結果】
{all_analyses}

【全文字起こし（先頭8000文字）】
{full_text[:8000]}

以下の構成でレポートを作成してください:

1. **エグゼクティブサマリー**: 番組全体の構造と特徴的なパターンの要約（3-5行）
2. **商品別分析**: 紹介された各商品について、トーク構造と番組構成を分析
3. **セールストーク・パターン分析**: 商品横断で見える共通パターン
   - トーク構造の型
   - よく使われるフレーズ・テクニック
   - 心理テクニック（希少性、社会的証明、アンカリング等）
4. **番組構成・CV導線の設計分析**:
   - 全体の時間配分
   - CTA（購買誘導）の設計パターン
   - 画面演出の特徴
5. **インサイト・示唆**: 他の販売チャネルに応用可能な知見

Markdown形式で出力してください。""",
    )

    return {
        "chunk_analyses": chunk_analyses,
        "integrated_analysis": integration_response.text,
    }


def step3_report(analysis_result, transcript, screenshots, output_dir, args):
    """Markdownレポート出力"""
    print("\n[Step 3/3] レポートを出力中...")

    report_path = os.path.join(output_dir, "analysis.md")
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(f"""# ショップチャンネル分析レポート

## 放送情報
- **分析日時**: {now}
- **URL**: {args.url}
- **分析対象**: {args.duration}分間
- **セグメント分割**: {args.segment_duration}分単位
- **スクリーンショット間隔**: {args.screenshot_interval}秒
- **スクリーンショット枚数**: {len(screenshots)}枚
- **文字起こしセグメント数**: {len(transcript['segments'])}

---

{analysis_result['integrated_analysis']}

---

## 付録: チャンク別詳細分析

""")
        for ca in analysis_result["chunk_analyses"]:
            f.write(f"### {ca['time_range']}\n\n{ca['analysis']}\n\n---\n\n")

        f.write("## 付録: 全文字起こしテキスト\n\n```\n")
        for seg in transcript["segments"]:
            m, s = divmod(int(seg["start"]), 60)
            f.write(f"[{m:02d}:{s:02d}] {seg['text']}\n")
        f.write("```\n")

    # メタデータ保存
    metadata = {
        "url": args.url,
        "duration_min": args.duration,
        "segment_duration_min": args.segment_duration,
        "screenshot_interval_sec": args.screenshot_interval,
        "whisper_model": args.whisper_model,
        "num_screenshots": len(screenshots),
        "num_segments": len(transcript["segments"]),
        "analyzed_at": now,
    }
    meta_path = os.path.join(output_dir, "metadata.json")
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)

    print(f"  レポート出力完了: {report_path}")
    return report_path


def main():
    args = parse_args()
    duration_sec = args.duration * 60

    # 出力ディレクトリ
    if args.resume_from:
        args.output = args.resume_from
        print("=" * 60)
        print("ショップチャンネル YouTube Live 分析パイプライン（再開モード）")
        print("=" * 60)
        print(f"  既存データから Step 2 以降を再実行します")
        print(f"  出力先: {args.output}")
        print("=" * 60)

        # 既存データ読み込み
        with open(os.path.join(args.output, "transcript.json"), "r", encoding="utf-8") as f:
            transcript = json.load(f)
        screenshots = sorted(glob.glob(os.path.join(args.output, "screenshots", "*.jpg")))

        analysis_result = step2_analyze(transcript, screenshots, args.output, args)
        report_path = step3_report(analysis_result, transcript, screenshots, args.output, args)

        print("\n" + "=" * 60)
        print("分析完了!")
        print(f"  レポート: {report_path}")
        print("=" * 60)
        return

    if args.output is None:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        args.output = os.path.join("output", timestamp)
    os.makedirs(args.output, exist_ok=True)

    print("=" * 60)
    print("ショップチャンネル YouTube Live 分析パイプライン")
    print("=" * 60)
    print(f"  URL: {args.url}")
    print(f"  対象区間: {args.start}分〜{args.start + args.duration}分（{args.duration}分間）")
    print(f"  セグメント分割: {args.segment_duration}分単位")
    print(f"  スクショ間隔: {args.screenshot_interval}秒")
    print(f"  Whisperモデル: {args.whisper_model}")
    print(f"  出力先: {args.output}")
    print("=" * 60)

    # Step 1: セグメント単位で DL → スクショ → 文字起こし → 削除
    transcript, screenshots = step1_process(
        args.url, args.start, args.duration, args.segment_duration,
        args.screenshot_interval, args.whisper_model, args.output,
    )
    if transcript is None:
        print("\n[FATAL] 処理に失敗しました。終了します。")
        sys.exit(1)

    if args.skip_analysis:
        print("\n--skip-analysis が指定されたため、ここで終了します。")
        print(f"出力先: {args.output}")
        return

    # Step 2: Gemini構造分析
    analysis_result = step2_analyze(transcript, screenshots, args.output, args)

    # Step 3: レポート出力
    report_path = step3_report(analysis_result, transcript, screenshots, args.output, args)

    print("\n" + "=" * 60)
    print("分析完了!")
    print(f"  レポート: {report_path}")
    print(f"  出力ディレクトリ: {args.output}")
    print("=" * 60)


if __name__ == "__main__":
    main()
