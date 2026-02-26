#!/usr/bin/env python3
"""
Create a PowerPoint presentation about Claude Code usage guide.
Built without external dependencies using raw OOXML format.
"""

import zipfile
import io
import os

# ============================================================
# Color palette & constants
# ============================================================
BG_DARK = "1A1A2E"       # dark navy background
BG_MEDIUM = "16213E"     # slightly lighter navy
ACCENT_ORANGE = "E94560"  # accent red-pink
ACCENT_BLUE = "0F3460"   # accent blue
TEXT_WHITE = "FFFFFF"
TEXT_LIGHT = "E0E0E0"
TEXT_GRAY = "B0B0B0"
ACCENT_GREEN = "00D2FF"   # cyan accent
ACCENT_PURPLE = "A855F7"  # purple accent
TITLE_GOLD = "FFD700"     # gold for emphasis

SLIDE_W = 12192000  # EMU (12192000 = 25.4cm * 480000)
SLIDE_H = 6858000   # EMU (6858000 = 14.29cm * 480000)

def emu(cm):
    """Convert cm to EMU."""
    return int(cm * 360000)

def pt_to_emu(pt):
    """Convert points to EMU."""
    return int(pt * 12700)

# ============================================================
# Slide content definitions
# ============================================================
slides_data = [
    # Slide 1: Title
    {
        "type": "title",
        "title": "Claude Code 使い方ガイド",
        "subtitle": "Anthropic公式 エージェント型AIコーディングツール",
        "footer": "2026年版 完全ガイド"
    },
    # Slide 2: What is Claude Code
    {
        "type": "content",
        "title": "Claude Code とは？",
        "bullets": [
            ("概要", "Anthropicが提供するターミナルベースのエージェント型AIコーディングツール"),
            ("自律実行", "コード理解・編集・テスト・Git操作を自然言語の指示だけで自律的に実行"),
            ("透明性", "実行計画を事前に提示し、各ステップの進捗をリアルタイムで表示"),
            ("統合性", "VS Code / JetBrains / ターミナル / GitHub / Web で利用可能"),
        ]
    },
    # Slide 3: Installation
    {
        "type": "content",
        "title": "インストール方法",
        "bullets": [
            ("ネイティブ版", "claude.com/download からダウンロード（推奨・自動更新対応）"),
            ("npm版", "npm install -g @anthropic-ai/claude-code（Node.js v18以上が必要）"),
            ("Web版", "claude.ai/code からブラウザで直接利用可能（インストール不要）"),
            ("起動", "ターミナルでプロジェクトディレクトリに移動し claude と入力"),
        ]
    },
    # Slide 4: Basic Usage
    {
        "type": "content",
        "title": "基本的な使い方",
        "bullets": [
            ("対話モード", "claude と入力して対話的にコーディング指示を出す"),
            ("ワンショット", "claude -p \"指示内容\" でパイプラインや単発実行が可能"),
            ("パイプ入力", "cat file.py | claude -p \"このコードをレビューして\""),
            ("モデル選択", "Opus 4.6（高性能） / Sonnet 4.6（バランス） / Haiku（高速）"),
        ]
    },
    # Slide 5: Slash Commands
    {
        "type": "content",
        "title": "主要なスラッシュコマンド (1/2)",
        "bullets": [
            ("/init", "CLAUDE.mdファイルを自動生成しプロジェクトを初期化"),
            ("/compact", "会話履歴を要約してコンテキストウィンドウを節約"),
            ("/commit", "差分を分析してConventional Commitsメッセージを自動生成"),
            ("/review", "Pull Requestのコードレビューを実行"),
        ]
    },
    # Slide 6: More Slash Commands
    {
        "type": "content",
        "title": "主要なスラッシュコマンド (2/2)",
        "bullets": [
            ("/help", "利用可能なコマンド一覧と使い方を表示"),
            ("/cost", "現在のセッションのトークン使用量・コストを表示"),
            ("/model", "使用するモデルを切り替え（Opus / Sonnet / Haiku）"),
            ("/clear", "会話履歴をクリアして新しいセッションを開始"),
        ]
    },
    # Slide 7: CLAUDE.md
    {
        "type": "content",
        "title": "CLAUDE.md による設定",
        "bullets": [
            ("役割", "プロジェクトのルール・慣習・コンテキストをClaudeに伝える設定ファイル"),
            ("配置場所", "プロジェクトルート / ~/.claude/ / サブディレクトリ（階層的に読み込み）"),
            ("記載内容", "コーディング規約、テストコマンド、アーキテクチャ概要、禁止事項など"),
            ("チーム共有", "Gitにコミットすることでチーム全体でルールを共有"),
        ]
    },
    # Slide 8: Permission & Security
    {
        "type": "content",
        "title": "権限モードとセキュリティ",
        "bullets": [
            ("デフォルト", "すべてのアクションに確認を求める安全なモード"),
            ("acceptEdits", "ファイル編集は自動承認、コマンド実行は確認を求める"),
            ("allowlist", "settings.json で許可/拒否するコマンドを細かく制御"),
            ("サンドボックス", "OS レベルのファイルシステム・ネットワーク分離で安全に実行"),
        ]
    },
    # Slide 9: IDE Integration
    {
        "type": "content",
        "title": "IDE連携",
        "bullets": [
            ("VS Code", "公式拡張機能でエディタ内にClaude Codeパネルを統合"),
            ("JetBrains", "IntelliJ / WebStorm / PyCharm 等のプラグイン（v2025.2+）"),
            ("GitHub Actions", "CI/CDパイプラインでPR自動作成・レビュー・Issue対応"),
            ("デスクトップ", "デスクトップアプリで複数セッションを並行管理"),
        ]
    },
    # Slide 10: MCP Servers
    {
        "type": "content",
        "title": "MCP サーバー連携",
        "bullets": [
            ("MCP とは", "Model Context Protocol - AI と外部ツールの統合オープン規格"),
            ("設定方法", "claude mcp add &lt;name&gt; &lt;command&gt; で簡単に追加"),
            ("活用例", "GitHub / Slack / DB / API / ファイルシステム等との接続"),
            ("スコープ", "プロジェクト単位 / ユーザー単位で設定・共有が可能"),
        ]
    },
    # Slide 11: Agentic Workflow
    {
        "type": "content",
        "title": "エージェント型ワークフロー",
        "bullets": [
            ("3フェーズ", "情報収集 → アクション実行 → 結果検証 の自律的なループ"),
            ("サブエージェント", "最大10の並列ワーカーで複雑なタスクを分担実行"),
            ("自動修正", "テスト失敗時に原因を分析し自動で修正を試行・再実行"),
            ("Git一気通貫", "ブランチ作成・コード変更・テスト・コミット・PR作成まで自動"),
        ]
    },
    # Slide 12: Best Practices
    {
        "type": "content",
        "title": "ベストプラクティス",
        "bullets": [
            ("ゴール指向", "技術的な指示より「何を達成したいか」を伝える"),
            ("CLAUDE.md活用", "/init でプロジェクト固有のルールを設定・チーム共有"),
            ("段階的依頼", "大きなタスクは段階的に分割して依頼・検証する"),
            ("コンテキスト管理", "/compact で会話を圧縮し、長時間セッションを効率化"),
        ]
    },
    # Slide 13: Use Cases
    {
        "type": "content",
        "title": "活用シーン",
        "bullets": [
            ("バグ修正", "エラーログを貼り付けて原因分析から修正まで一括で依頼"),
            ("新機能開発", "要件を伝えてコード生成・テスト作成・PR作成まで自動実行"),
            ("リファクタリング", "既存コードの改善・TypeScript移行・モダン化を自律実行"),
            ("コードレビュー", "/review でPRの品質チェックとレビューコメントを自動生成"),
        ]
    },
    # Slide 14: Summary
    {
        "type": "summary",
        "title": "まとめ",
        "points": [
            "Claude Code はターミナルで動作するエージェント型AIコーディングツール",
            "自然言語で指示するだけで、コードの編集・テスト・Git操作を自律的に実行",
            "CLAUDE.md でプロジェクトルールを設定し、チーム全体で一貫した動作を実現",
            "VS Code / JetBrains / GitHub Actions との連携で開発フロー全体をカバー",
            "サブエージェント・MCP連携で、複雑なタスクも効率的に処理",
        ]
    },
]

# ============================================================
# XML Templates
# ============================================================

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {slide_overrides}
</Types>"""

RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

PRESENTATION_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    {slide_ids}
  </p:sldIdLst>
  <p:sldSz cx="{w}" cy="{h}"/>
  <p:notesSz cx="{h}" cy="{w}"/>
</p:presentation>"""

PRES_PROPS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"""

PRESENTATION_RELS_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdPresProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  {slide_rels}
</Relationships>"""

SLIDE_MASTER_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="{bg}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle/>
    <p:bodyStyle/>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>"""

SLIDE_MASTER_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""

SLIDE_LAYOUT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

SLIDE_LAYOUT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

THEME_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ClaudeTheme">
  <a:themeElements>
    <a:clrScheme name="Claude">
      <a:dk1><a:srgbClr val="1A1A2E"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="16213E"/></a:dk2>
      <a:lt2><a:srgbClr val="E0E0E0"/></a:lt2>
      <a:accent1><a:srgbClr val="E94560"/></a:accent1>
      <a:accent2><a:srgbClr val="0F3460"/></a:accent2>
      <a:accent3><a:srgbClr val="00D2FF"/></a:accent3>
      <a:accent4><a:srgbClr val="A855F7"/></a:accent4>
      <a:accent5><a:srgbClr val="FFD700"/></a:accent5>
      <a:accent6><a:srgbClr val="10B981"/></a:accent6>
      <a:hlink><a:srgbClr val="00D2FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="A855F7"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Claude">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Claude">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""

SLIDE_RELS_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""

# ============================================================
# Shape builders
# ============================================================

def make_sp(sp_id, name, x, y, cx, cy, body_xml):
    """Create a shape XML element."""
    return f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="{name}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  {body_xml}
</p:sp>"""

def make_filled_rect(sp_id, name, x, y, cx, cy, fill_color, alpha=None):
    """Create a filled rectangle shape."""
    alpha_xml = f'<a:alpha val="{alpha}"/>' if alpha else ""
    return f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="{name}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{fill_color}">{alpha_xml}</a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
</p:sp>"""

def make_rounded_rect(sp_id, name, x, y, cx, cy, fill_color, alpha=None):
    """Create a rounded rectangle shape."""
    alpha_xml = f'<a:alpha val="{alpha}"/>' if alpha else ""
    return f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="{name}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="{fill_color}">{alpha_xml}</a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
</p:sp>"""

def make_text_run(text, size_pt, color, bold=False, font="Meiryo UI"):
    """Create a text run XML."""
    bold_attr = ' b="1"' if bold else ""
    return f"""<a:r><a:rPr lang="ja-JP" sz="{size_pt * 100}"{bold_attr} dirty="0">
    <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
    <a:latin typeface="{font}"/><a:ea typeface="{font}"/>
  </a:rPr><a:t>{text}</a:t></a:r>"""

def make_paragraph(runs, align="l", spc_before=0, spc_after=0):
    """Create a paragraph XML with multiple runs."""
    spc_xml = ""
    if spc_before or spc_after:
        parts = []
        if spc_before:
            parts.append(f'<a:spcBef><a:spcPts val="{spc_before}"/></a:spcBef>')
        if spc_after:
            parts.append(f'<a:spcAft><a:spcPts val="{spc_after}"/></a:spcAft>')
        spc_xml = "".join(parts)
    return f'<a:p><a:pPr algn="{align}">{spc_xml}</a:pPr>{"".join(runs)}</a:p>'

def make_textbox(sp_id, name, x, y, cx, cy, paragraphs, anchor="t"):
    """Create a textbox shape with paragraphs."""
    body = f'<p:txBody><a:bodyPr wrap="square" anchor="{anchor}"/><a:lstStyle/>{"".join(paragraphs)}</p:txBody>'
    return make_sp(sp_id, name, x, y, cx, cy, body)

# ============================================================
# Slide Generators
# ============================================================

def generate_title_slide(data):
    """Generate the title slide XML."""
    shapes = []
    sid = 2

    # Background gradient rectangle
    shapes.append(make_filled_rect(sid, "bg", 0, 0, SLIDE_W, SLIDE_H, BG_DARK))
    sid += 1

    # Accent stripe at top
    shapes.append(make_filled_rect(sid, "stripe_top", 0, 0, SLIDE_W, emu(0.6), ACCENT_ORANGE))
    sid += 1

    # Decorative left bar
    shapes.append(make_filled_rect(sid, "left_bar", emu(1.5), emu(3), emu(0.15), emu(5), ACCENT_GREEN))
    sid += 1

    # Terminal icon box
    shapes.append(make_rounded_rect(sid, "terminal_box", emu(3), emu(2.5), emu(3.5), emu(2), "0F3460", "80000"))
    sid += 1

    # Terminal prompt text
    terminal_paras = [
        make_paragraph([make_text_run("$ claude", 16, ACCENT_GREEN, font="Consolas")], "l"),
        make_paragraph([make_text_run("> Ready to assist...", 12, TEXT_GRAY, font="Consolas")], "l"),
    ]
    shapes.append(make_textbox(sid, "terminal_text", emu(3.3), emu(2.8), emu(3), emu(1.5), terminal_paras, "ctr"))
    sid += 1

    # Main title
    title_paras = [
        make_paragraph([make_text_run(data["title"], 36, TEXT_WHITE, bold=True)], "ctr"),
    ]
    shapes.append(make_textbox(sid, "title", emu(5), emu(3.5), emu(20), emu(3), title_paras, "ctr"))
    sid += 1

    # Subtitle
    sub_paras = [
        make_paragraph([make_text_run(data["subtitle"], 18, ACCENT_GREEN)], "ctr"),
    ]
    shapes.append(make_textbox(sid, "subtitle", emu(5), emu(6), emu(20), emu(2), sub_paras, "ctr"))
    sid += 1

    # Footer
    footer_paras = [
        make_paragraph([make_text_run(data["footer"], 12, TEXT_GRAY)], "ctr"),
    ]
    shapes.append(make_textbox(sid, "footer", emu(5), emu(9), emu(20), emu(1.5), footer_paras, "ctr"))
    sid += 1

    # Bottom accent line
    shapes.append(make_filled_rect(sid, "bottom_line", emu(8), emu(11), emu(18), emu(0.1), ACCENT_ORANGE))
    sid += 1

    return shapes

def generate_content_slide(data, slide_num):
    """Generate a content slide with bullets."""
    shapes = []
    sid = 2

    # Background
    shapes.append(make_filled_rect(sid, "bg", 0, 0, SLIDE_W, SLIDE_H, BG_DARK))
    sid += 1

    # Header bar
    shapes.append(make_filled_rect(sid, "header_bar", 0, 0, SLIDE_W, emu(2.5), BG_MEDIUM))
    sid += 1

    # Orange accent line under header
    shapes.append(make_filled_rect(sid, "accent_line", 0, emu(2.5), SLIDE_W, emu(0.08), ACCENT_ORANGE))
    sid += 1

    # Slide number
    num_paras = [make_paragraph([make_text_run(f"{slide_num:02d}", 14, ACCENT_ORANGE, bold=True)], "l")]
    shapes.append(make_textbox(sid, "slide_num", emu(1), emu(0.3), emu(2), emu(1.5), num_paras, "t"))
    sid += 1

    # Title
    title_paras = [make_paragraph([make_text_run(data["title"], 28, TEXT_WHITE, bold=True)], "l")]
    shapes.append(make_textbox(sid, "title", emu(3), emu(0.4), emu(28), emu(2), title_paras, "t"))
    sid += 1

    # Bullet items
    y_start = emu(3.2)
    item_height = emu(2.8)

    for i, (label, desc) in enumerate(data["bullets"]):
        y_pos = y_start + i * item_height

        # Bullet card background
        shapes.append(make_rounded_rect(sid, f"card_{i}", emu(1.5), y_pos, emu(30), item_height - emu(0.3), BG_MEDIUM, "60000"))
        sid += 1

        # Left accent bar for card
        colors = [ACCENT_ORANGE, ACCENT_GREEN, ACCENT_PURPLE, TITLE_GOLD]
        bar_color = colors[i % len(colors)]
        shapes.append(make_filled_rect(sid, f"bar_{i}", emu(1.5), y_pos, emu(0.12), item_height - emu(0.3), bar_color))
        sid += 1

        # Label + description
        bullet_paras = [
            make_paragraph([
                make_text_run(f"{label}    ", 16, bar_color, bold=True),
                make_text_run(desc, 14, TEXT_LIGHT),
            ], "l", spc_before=200),
        ]
        shapes.append(make_textbox(sid, f"text_{i}", emu(2.5), y_pos + emu(0.2), emu(28), item_height - emu(0.5), bullet_paras, "ctr"))
        sid += 1

    # Footer line
    shapes.append(make_filled_rect(sid, "footer_line", emu(1.5), SLIDE_H - emu(1), emu(30), emu(0.05), ACCENT_BLUE))
    sid += 1

    # Page number
    page_paras = [make_paragraph([make_text_run(f"{slide_num} / {len(slides_data)}", 10, TEXT_GRAY)], "r")]
    shapes.append(make_textbox(sid, "page_num", emu(27), SLIDE_H - emu(0.9), emu(5), emu(0.8), page_paras, "t"))
    sid += 1

    return shapes

def generate_summary_slide(data, slide_num):
    """Generate the summary slide."""
    shapes = []
    sid = 2

    # Background
    shapes.append(make_filled_rect(sid, "bg", 0, 0, SLIDE_W, SLIDE_H, BG_DARK))
    sid += 1

    # Top accent
    shapes.append(make_filled_rect(sid, "top_accent", 0, 0, SLIDE_W, emu(0.1), ACCENT_ORANGE))
    sid += 1

    # Title area
    shapes.append(make_filled_rect(sid, "title_bg", 0, emu(0.1), SLIDE_W, emu(2.5), BG_MEDIUM))
    sid += 1

    # Title
    title_paras = [make_paragraph([make_text_run(data["title"], 32, TEXT_WHITE, bold=True)], "ctr")]
    shapes.append(make_textbox(sid, "title", emu(2), emu(0.3), emu(29), emu(2), title_paras, "ctr"))
    sid += 1

    # Accent line
    shapes.append(make_filled_rect(sid, "accent_line", 0, emu(2.6), SLIDE_W, emu(0.08), ACCENT_ORANGE))
    sid += 1

    # Summary points
    y_start = emu(3.5)
    for i, point in enumerate(data["points"]):
        y_pos = y_start + i * emu(2)
        colors = [ACCENT_GREEN, ACCENT_ORANGE, ACCENT_PURPLE, TITLE_GOLD, "00D2FF"]
        dot_color = colors[i % len(colors)]

        # Checkmark / number circle
        shapes.append(make_rounded_rect(sid, f"num_{i}", emu(2), y_pos, emu(0.8), emu(0.8), dot_color))
        sid += 1

        num_paras = [make_paragraph([make_text_run(str(i + 1), 12, BG_DARK, bold=True)], "ctr")]
        shapes.append(make_textbox(sid, f"num_text_{i}", emu(2), y_pos, emu(0.8), emu(0.8), num_paras, "ctr"))
        sid += 1

        # Point text
        point_paras = [make_paragraph([make_text_run(point, 14, TEXT_LIGHT)], "l")]
        shapes.append(make_textbox(sid, f"point_{i}", emu(3.5), y_pos, emu(27), emu(1.2), point_paras, "ctr"))
        sid += 1

    # Bottom bar
    shapes.append(make_filled_rect(sid, "bottom_bar", 0, SLIDE_H - emu(1.2), SLIDE_W, emu(1.2), BG_MEDIUM))
    sid += 1

    footer_paras = [make_paragraph([
        make_text_run("Claude Code", 12, ACCENT_GREEN, bold=True),
        make_text_run("  -  Start coding with AI today", 12, TEXT_GRAY),
    ], "ctr")]
    shapes.append(make_textbox(sid, "footer_text", emu(2), SLIDE_H - emu(1.1), emu(29), emu(1), footer_paras, "ctr"))
    sid += 1

    return shapes

def wrap_slide_xml(shapes):
    """Wrap shapes into a complete slide XML."""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      {"".join(shapes)}
    </p:spTree>
  </p:cSld>
</p:sld>"""

# ============================================================
# Main: build the PPTX
# ============================================================

def build_pptx(output_path):
    """Build the complete PPTX file."""
    num_slides = len(slides_data)

    # Generate slide XMLs
    slide_xmls = []
    for i, sdata in enumerate(slides_data):
        slide_num = i + 1
        if sdata["type"] == "title":
            shapes = generate_title_slide(sdata)
        elif sdata["type"] == "summary":
            shapes = generate_summary_slide(sdata, slide_num)
        else:
            shapes = generate_content_slide(sdata, slide_num)
        slide_xmls.append(wrap_slide_xml(shapes))

    # Build [Content_Types].xml
    slide_overrides = "\n".join(
        f'  <Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(num_slides)
    )
    content_types = CONTENT_TYPES_XML.format(slide_overrides=slide_overrides)

    # Build presentation.xml
    slide_ids = "\n".join(
        f'    <p:sldId id="{256 + i}" r:id="rIdSlide{i+1}"/>'
        for i in range(num_slides)
    )
    presentation = PRESENTATION_XML.format(slide_ids=slide_ids, w=SLIDE_W, h=SLIDE_H)

    # Build presentation.xml.rels
    slide_rels = "\n".join(
        f'  <Relationship Id="rIdSlide{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>'
        for i in range(num_slides)
    )
    pres_rels = PRESENTATION_RELS_TEMPLATE.format(slide_rels=slide_rels)

    # Write ZIP
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", RELS_XML)
        zf.writestr("ppt/presentation.xml", presentation)
        zf.writestr("ppt/presProps.xml", PRES_PROPS_XML)
        zf.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
        zf.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER_XML.format(bg=BG_DARK))
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT_XML)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)

        for i in range(num_slides):
            zf.writestr(f"ppt/slides/slide{i+1}.xml", slide_xmls[i])
            zf.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", SLIDE_RELS_TEMPLATE)

    print(f"Created: {output_path}")
    print(f"Slides: {num_slides}")

if __name__ == "__main__":
    output = os.path.join(os.path.dirname(os.path.abspath(__file__)), "claude_code_guide.pptx")
    build_pptx(output)
