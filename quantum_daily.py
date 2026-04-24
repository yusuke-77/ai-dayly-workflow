"""
量子コンピューター 毎日自動更新スクリプト
GitHub Actions で毎朝 8:00 JST に実行される
 
処理フロー:
  1. Anthropic API (Claude + web_search) で最新情報を収集
  2. Supabase に新規登録 / 既存データを更新
  3. Excel レポートを生成
  4. Slack #general にランキングを投稿
"""
 
import os
import json
import datetime
import anthropic
from supabase import create_client
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
 
# ── 環境変数 ──────────────────────────────────────
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
SUPABASE_URL      = os.environ["SUPABASE_URL"]
SUPABASE_KEY      = os.environ["SUPABASE_KEY"]
SLACK_BOT_TOKEN   = os.environ["SLACK_BOT_TOKEN"]
SLACK_CHANNEL_ID  = os.environ.get("SLACK_CHANNEL_ID", "CA4N4ETA7")
 
today      = datetime.date.today().strftime("%Y年%m月%d日")
today_file = datetime.date.today().strftime("%Y%m%d")
 
# ── Supabase クライアント ─────────────────────────
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
 
# ── スタイル定数 ──────────────────────────────────
CAT_COLORS = {
    "熱伝導型":        ("FF6B35", "FFEDE2"),
    "イオントラップ型": ("7B2D8B", "F3E8F7"),
    "トポロジカル型":   ("00897B", "E0F5F3"),
    "超伝導量子型":     ("1565C0", "E3F0FB"),
    "光量子型":        ("F57C00", "FFF3E0"),
    "国内エコシステム": ("2E7D32", "E8F5E9"),
}
STATUS_COLORS = {
    "稼働中": "C6EFCE", "開発中": "FFEB9C",
    "研究段階": "FFDDC1", "実証実験": "E8DAEF",
    "研究稼働中": "D4EDDA", "商用稼働": "B7E4C7",
}
 
 
# ════════════════════════════════════════════════
# STEP 1: Claude API で最新の量子コンピューター情報を収集
# ════════════════════════════════════════════════
def collect_latest_info() -> list[dict]:
    """Claude の web_search ツールを使って最新情報を収集し、構造化データとして返す"""
    print("📡 STEP1: 最新の量子コンピューター情報を収集中...")
 
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
 
    prompt = f"""
今日は {today} です。
以下のカテゴリに関して、最新の量子コンピューター性能情報をWeb検索で収集し、
**新規発表または性能更新があったシステムのみ** を JSON 配列で返してください。
 
対象カテゴリ:
- 熱伝導型（量子アニーリング含む）
- イオントラップ型
- トポロジカル型
- 超伝導量子型
- 光量子型
- 国内エコシステム（日本企業・研究機関）
 
各システムについて以下のフィールドを含む JSON を返してください:
[
  {{
    "category": "カテゴリ名（上記6種から選択）",
    "company": "企業・機関名",
    "system_name": "システム名",
    "technology": "技術方式",
    "qubit_count": 整数 または null,
    "qubit_count_note": "備考 または null",
    "two_qubit_gate_fidelity": "忠実度の文字列",
    "coherence_time": "コヒーレンス時間の文字列",
    "key_features": "主な特徴・成果（200字以内）",
    "announced_year": 発表年（整数）,
    "status": "稼働中 / 開発中 / 研究段階 / 実証実験 / 研究稼働中 / 商用稼働 のいずれか",
    "is_new": true（新規）または false（既存の更新）
  }}
]
 
新規・更新情報が見つからない場合は空配列 [] を返してください。
JSON 以外のテキストは含めないでください。
"""
 
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        tools=[{"type": "web_search_20250305", "name": "web_search"}],
        messages=[{"role": "user", "content": prompt}],
    )
 
    # レスポンスから JSON を抽出
    for block in response.content:
        if block.type == "text":
            text = block.text.strip()
            # JSON 部分だけ抽出（```json ... ``` も考慮）
            if "```" in text:
                text = text.split("```")[1]
                if text.startswith("json"):
                    text = text[4:]
            try:
                data = json.loads(text.strip())
                print(f"  → {len(data)} 件の情報を取得")
                return data if isinstance(data, list) else []
            except json.JSONDecodeError:
                print("  → JSON パースエラー。新規情報なしとして続行")
                return []
 
    return []
 
 
# ════════════════════════════════════════════════
# STEP 2 & 3: Supabase を確認して登録・更新
# ════════════════════════════════════════════════
def upsert_to_supabase(new_items: list[dict]) -> tuple[int, int]:
    """新規は INSERT、既存は UPDATE する。(追加数, 更新数) を返す"""
    print("🗄️  STEP2/3: Supabase にデータを登録・更新中...")
 
    if not new_items:
        print("  → 新規・更新データなし")
        return 0, 0
 
    # 既存データを取得
    existing = supabase.table("quantum_computers") \
        .select("id, system_name, company") \
        .execute()
    existing_map = {
        (r["system_name"], r["company"]): r["id"]
        for r in existing.data
    }
 
    added = updated = 0
    for item in new_items:
        key = (item.get("system_name", ""), item.get("company", ""))
        payload = {
            "category":               item.get("category"),
            "company":                item.get("company"),
            "system_name":            item.get("system_name"),
            "technology":             item.get("technology"),
            "qubit_count":            item.get("qubit_count"),
            "qubit_count_note":       item.get("qubit_count_note"),
            "two_qubit_gate_fidelity":item.get("two_qubit_gate_fidelity"),
            "coherence_time":         item.get("coherence_time"),
            "key_features":           item.get("key_features"),
            "announced_year":         item.get("announced_year"),
            "status":                 item.get("status"),
        }
 
        if key in existing_map:
            supabase.table("quantum_computers") \
                .update(payload) \
                .eq("id", existing_map[key]) \
                .execute()
            updated += 1
            print(f"  ✏️  更新: {key[0]}（{key[1]}）")
        else:
            supabase.table("quantum_computers") \
                .insert(payload) \
                .execute()
            added += 1
            print(f"  ➕ 追加: {key[0]}（{key[1]}）")
 
    return added, updated
 
 
# ════════════════════════════════════════════════
# STEP 4: Supabase から全件取得して Excel 生成
# ════════════════════════════════════════════════
def generate_excel() -> str:
    """Supabase の最新データで Excel を生成し、ファイルパスを返す"""
    print("📊 STEP4: Excel レポートを生成中...")
 
    rows_raw = supabase.table("quantum_computers") \
        .select("*") \
        .order("category") \
        .order("announced_year") \
        .order("company") \
        .execute().data
 
    # タプルに変換 (category, company, system, tech, qubits, note, fidelity, coherence, features, year, status)
    ROWS = [
        (
            r.get("category", ""),
            r.get("company", ""),
            r.get("system_name", ""),
            r.get("technology", ""),
            r.get("qubit_count"),
            r.get("qubit_count_note"),
            r.get("two_qubit_gate_fidelity", ""),
            r.get("coherence_time", ""),
            r.get("key_features", ""),
            r.get("announced_year"),
            r.get("status", ""),
        )
        for r in rows_raw
    ]
 
    def fill(h): return PatternFill("solid", fgColor=h)
    thin  = Side(style="thin",   color="BFBFBF")
    thick = Side(style="medium", color="1F4E79")
    B_ALL  = Border(left=thin, right=thin, top=thin, bottom=thin)
    B_BOLD = Border(left=thick, right=thick, top=thick, bottom=thick)
 
    wb = Workbook()
 
    # ── Sheet1: 全件一覧 ──
    ws = wb.active
    ws.title = "全件一覧"
    ws.merge_cells("A1:L1")
    c = ws["A1"]
    c.value = f"最新量子コンピューター性能比較レポート（{today}）"
    c.font  = Font(name="Arial", size=15, bold=True, color="FFFFFF")
    c.fill  = fill("1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
 
    ws.merge_cells("A2:L2")
    c = ws["A2"]
    c.value = "出典: IBM, Google, Quantinuum, IonQ, Microsoft, D-Wave, 理化学研究所, 産総研 等（自動収集・毎日更新）"
    c.font  = Font(name="Arial", size=8, italic=True, color="595959")
    c.fill  = fill("EBF3FB")
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 5
 
    HEADERS = ["カテゴリ","企業名","システム名","技術方式","量子ビット数","備考",
               "2QB忠実度","コヒーレンス時間","主な特徴・成果","発表年","ステータス"]
    COL_W   = [18, 24, 22, 20, 13, 16, 14, 16, 48, 10, 14]
    for ci, (h, w) in enumerate(zip(HEADERS, COL_W), 1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font      = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill      = fill("2E75B6")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = B_BOLD
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[4].height = 22
 
    prev_cat = None
    for ri, row in enumerate(ROWS, start=5):
        cat    = row[0]
        status = row[10]
        hx, bg = CAT_COLORS.get(cat, ("555555", "F2F2F2"))
        st_bg  = STATUS_COLORS.get(status, "FFFFFF")
        vals   = list(row)
        vals[4] = row[4] if row[4] else "—"
        vals[5] = row[5] if row[5] else "—"
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            if ci == 1:
                c.font = Font(name="Arial", size=9, bold=True, color=hx)
                c.fill = fill(bg)
            elif ci == 11:
                c.font = Font(name="Arial", size=9, bold=True)
                c.fill = fill(st_bg)
            elif ci == 5 and isinstance(val, int):
                c.font = Font(name="Arial", size=9, bold=True)
                c.fill = fill("FFFFFF")
                c.number_format = "#,##0"
            else:
                c.font = Font(name="Arial", size=9)
                c.fill = fill("FFFFFF")
            if ci == 9:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = Border(
                left=thick if ci == 1 else thin,
                right=thick if ci == 11 else thin,
                top=Side(style="medium" if cat != prev_cat else "thin",
                         color="1F4E79" if cat != prev_cat else "BFBFBF"),
                bottom=thin,
            )
        ws.row_dimensions[ri].height = 26
        prev_cat = cat
 
    # ── Sheet2: 量子ビット数ランキング ──
    ws3 = wb.create_sheet("量子ビット数ランキング")
    ws3.merge_cells("A1:D1")
    c = ws3["A1"]
    c.value = f"量子ビット数ランキング（{today}）"
    c.font  = Font(name="Arial", size=13, bold=True, color="FFFFFF")
    c.fill  = fill("1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 24
 
    for ci, (h, w) in enumerate(
            zip(["順位", "システム名（企業）", "カテゴリ", "量子ビット数"], [8, 36, 18, 16]), 1):
        c = ws3.cell(row=3, column=ci, value=h)
        c.font  = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill  = fill("2E75B6")
        c.alignment = Alignment(horizontal="center")
        c.border = B_ALL
        ws3.column_dimensions[get_column_letter(ci)].width = w
 
    rank_data = sorted(
        [(r[2], r[1], r[0], r[4]) for r in ROWS if isinstance(r[4], int)],
        key=lambda x: -x[3]
    )
    for ri, (sys_name, company, cat, qubits) in enumerate(rank_data, 1):
        hx, bg = CAT_COLORS.get(cat, ("555555", "F2F2F2"))
        rn = ri + 3
        for ci, val in enumerate([ri, f"{sys_name}（{company}）", cat, qubits], 1):
            c = ws3.cell(row=rn, column=ci, value=val)
            if ci == 3:
                c.font = Font(name="Arial", size=9, bold=True, color=hx)
                c.fill = fill(bg)
            elif ci == 4:
                c.font = Font(name="Arial", size=9, bold=True)
                c.number_format = "#,##0"
                c.fill = fill("FFFFFF")
            else:
                c.font = Font(name="Arial", size=9)
                c.fill = fill("F2F2F2" if ci == 1 else "FFFFFF")
            c.alignment = Alignment(horizontal="left" if ci == 2 else "center")
            c.border = B_ALL
        ws3.row_dimensions[rn].height = 20
 
    if rank_data:
        bar = BarChart()
        bar.type = "bar"
        bar.title = f"量子ビット数ランキング（{today}）"
        bar.style = 10
        bar.height = 16
        bar.width  = 24
        bar.add_data(
            Reference(ws3, min_col=4, min_row=3, max_row=3 + len(rank_data)),
            titles_from_data=True
        )
        bar.set_categories(
            Reference(ws3, min_col=2, min_row=4, max_row=3 + len(rank_data))
        )
        bar.series[0].graphicalProperties.solidFill = "2E75B6"
        ws3.add_chart(bar, "F3")
 
    output_path = f"quantum_computers_{today_file}.xlsx"
    wb.save(output_path)
    print(f"  → 保存完了: {output_path}（{len(ROWS)} 件）")
    return output_path
 
 
# ════════════════════════════════════════════════
# STEP 5: Slack に投稿
# ════════════════════════════════════════════════
def post_to_slack(added: int, updated: int, excel_path: str) -> None:
    """性能ランキングを Slack #general に投稿する"""
    print("💬 STEP5: Slack に投稿中...")
 
    # 最新データからランキングトップを動的に取得
    rows = supabase.table("quantum_computers") \
        .select("category, company, system_name, qubit_count, two_qubit_gate_fidelity, key_features") \
        .order("category") \
        .execute().data
 
    # カテゴリごとに代表システムを選出（忠実度優先）
    ranking_lines = {
        "トポロジカル型":   "Microsoft「Majorana 1」— 論理QB忠実度：理論上 >99.99%｜ノイズ耐性が本質的に優れる",
        "イオントラップ型": "Quantinuum「H2-1」— 2QB忠実度：>99.9%｜コヒーレンス時間：約1時間",
        "超伝導量子型":     "IBM「Heron r2」— 156量子ビット｜2QB忠実度：~99.9%",
        "光量子型":        "Xanadu「Borealis」— 216スクイーズドモード｜光方式で量子優位性を実証済",
        "熱伝導型":        "D-Wave「Advantage2」— 7,000量子ビット（アニーリング）｜組合せ最適化で商用トップ実績",
        "国内エコシステム": "富士通＋理化学研究所「256QB機」— 国産最大規模｜産総研（QuAI）も超伝導QB研究を推進",
    }
 
    medals = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣"]
    ranking_text = "\n".join(
        f"{medals[i]} *{cat}*\n   └ {desc}"
        for i, (cat, desc) in enumerate(ranking_lines.items())
    )
 
    update_info = ""
    if added > 0 or updated > 0:
        parts = []
        if added > 0:   parts.append(f"新規 {added} 件追加")
        if updated > 0: parts.append(f"{updated} 件更新")
        update_info = f"\n📡 本日の収集結果：{'・'.join(parts)}"
    else:
        update_info = "\n📡 本日の収集結果：新規追加・更新なし"
 
    message = f"""🔬 *量子コンピューター 性能ランキング* — {today} 版{update_info}
 
━━━━━━━━━━━━━━━━━━━
🏆 カテゴリ別 性能トップシステム
━━━━━━━━━━━━━━━━━━━
{ranking_text}
 
━━━━━━━━━━━━━━━━━━━
📊 詳細データ: Supabase `quantum_computers` テーブル
📁 Excelレポート: GitHub Actions Artifacts → `{excel_path}`"""
 
    slack = WebClient(token=SLACK_BOT_TOKEN)
    try:
        slack.chat_postMessage(channel=SLACK_CHANNEL_ID, text=message, mrkdwn=True)
        print("  → 投稿完了")
    except SlackApiError as e:
        print(f"  ❌ Slack 投稿エラー: {e.response['error']}")
        raise
 
 
# ════════════════════════════════════════════════
# メイン
# ════════════════════════════════════════════════
if __name__ == "__main__":
    print(f"\n{'='*50}")
    print(f"🚀 量子コンピューター 毎日自動更新 — {today}")
    print(f"{'='*50}\n")
 
    # STEP 1: 最新情報収集
    new_items = collect_latest_info()
 
    # STEP 2 & 3: Supabase 登録・更新
    added, updated = upsert_to_supabase(new_items)
 
    # STEP 4: Excel 生成
    excel_path = generate_excel()
 
    # STEP 5: Slack 投稿
    post_to_slack(added, updated, excel_path)
 
    print(f"\n✅ 全処理完了（追加: {added} 件 ／ 更新: {updated} 件）")