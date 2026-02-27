#!/usr/bin/env python3
"""
Create a PowerPoint presentation about MCP (Model Context Protocol) servers.
Built without external dependencies using raw OOXML format.
"""

import zipfile
import os

# ============================================================
# Color palette & constants
# ============================================================
BG_DARK = "0D1117"       # GitHub-dark background
BG_MEDIUM = "161B22"     # slightly lighter
BG_CARD = "1C2333"       # card background
ACCENT_TEAL = "00BFA6"   # primary teal accent
ACCENT_BLUE = "58A6FF"   # link blue
ACCENT_ORANGE = "F0883E" # warning/highlight orange
ACCENT_PURPLE = "BC8CFF"  # purple accent
ACCENT_GREEN = "3FB950"   # success green
ACCENT_RED = "F85149"     # error/important red
TEXT_WHITE = "F0F6FC"
TEXT_LIGHT = "C9D1D9"
TEXT_GRAY = "8B949E"
TITLE_GOLD = "F7C948"     # gold for emphasis

SLIDE_W = 12192000  # EMU (25.4cm)
SLIDE_H = 6858000   # EMU (14.29cm)


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
        "title": "MCP サーバー 完全ガイド",
        "subtitle": "Model Context Protocol - AI と外部ツールをつなぐオープン規格",
        "footer": "2026年版 | Anthropic MCP Specification"
    },
    # Slide 2: What is MCP
    {
        "type": "content",
        "title": "MCP とは何か？",
        "bullets": [
            ("定義", "Model Context Protocol - AIモデルが外部データ・ツールに安全にアクセスするためのオープン規格"),
            ("目的", "LLMアプリケーションと外部データソース・ツール間の標準的なインターフェースを提供"),
            ("例え", "AI版の「USB-C」- どのAIもどのツールも同じ規格で接続できるユニバーサルな接続方式"),
            ("開発元", "Anthropicが策定しオープンソースで公開。業界全体での採用が進行中"),
        ]
    },
    # Slide 3: Architecture
    {
        "type": "content",
        "title": "MCP のアーキテクチャ",
        "bullets": [
            ("Host", "AIアプリケーション本体（Claude Desktop, IDE, カスタムアプリ等）"),
            ("Client", "Host内でMCPプロトコルを管理するコンポーネント（1サーバーに1クライアント）"),
            ("Server", "外部リソースやツールを公開するプログラム（ローカルまたはリモート実行）"),
            ("Transport", "stdio（ローカル）またはHTTP+SSE（リモート）で通信を実現"),
        ]
    },
    # Slide 4: Core Capabilities
    {
        "type": "content",
        "title": "MCP サーバーの3つの主要機能",
        "bullets": [
            ("Tools", "AIが呼び出せる関数（API呼び出し、DB操作、ファイル処理等）- モデルが制御"),
            ("Resources", "AIが参照できるデータソース（ファイル内容、DB結果等）- アプリが制御"),
            ("Prompts", "事前定義されたテンプレート（コードレビュー手順等）- ユーザーが制御"),
        ]
    },
    # Slide 5: Setting up MCP in Claude Code
    {
        "type": "content",
        "title": "Claude Code での MCP 設定",
        "bullets": [
            ("追加コマンド", "claude mcp add <name> <command> [args] でサーバーを登録"),
            ("スコープ", "-s project（プロジェクト単位）/ -s user（ユーザー単位）を指定"),
            ("一覧確認", "claude mcp list で登録済みサーバーの一覧を表示"),
            ("削除", "claude mcp remove <name> で不要なサーバーを削除"),
        ]
    },
    # Slide 6: Popular MCP Servers
    {
        "type": "content",
        "title": "代表的な MCP サーバー",
        "bullets": [
            ("Filesystem", "ローカルファイルの読み書き・検索を安全なサンドボックスで提供"),
            ("GitHub", "リポジトリ操作・Issue/PR管理・コード検索をAIから直接実行"),
            ("PostgreSQL / SQLite", "データベースへの読み取り専用アクセスやスキーマ解析を提供"),
            ("Slack / Google Drive", "チームコミュニケーションやドキュメント管理との連携"),
        ]
    },
    # Slide 7: Building Your Own MCP Server
    {
        "type": "content",
        "title": "MCP サーバーの自作方法",
        "bullets": [
            ("SDK", "Python SDK / TypeScript SDK が公式提供。pip install mcp または npm install @modelcontextprotocol/sdk"),
            ("ツール定義", "@server.tool() デコレータで関数を定義するだけでAIが呼び出し可能に"),
            ("型安全", "入力パラメータはJSON Schemaで定義。バリデーションが自動で実行される"),
            ("起動", "stdio トランスポートで起動し、Claude Code から即座に利用可能"),
        ]
    },
    # Slide 8: Python Example
    {
        "type": "code",
        "title": "Python での実装例",
        "code_lines": [
            ("from mcp.server.fastmcp import FastMCP", "import"),
            ("", ""),
            ("server = FastMCP('my-server')", "init"),
            ("", ""),
            ("@server.tool()", "decorator"),
            ("async def get_weather(city: str) -> str:", "func"),
            ("    '''指定都市の天気を取得'''", "doc"),
            ("    return f'{city}の天気: 晴れ 25°C'", "return"),
            ("", ""),
            ("server.run(transport='stdio')", "run"),
        ]
    },
    # Slide 9: Transport Types
    {
        "type": "content",
        "title": "トランスポートの種類",
        "bullets": [
            ("stdio", "標準入出力で通信。ローカル実行に最適。セットアップが簡単で低レイテンシ"),
            ("HTTP+SSE", "Server-Sent Events で通信。リモートサーバーに対応。複数クライアント接続可能"),
            ("Streamable HTTP", "最新の推奨方式。HTTP POST/GETベースでステートレスにもステートフルにも対応"),
            ("選択基準", "ローカル→stdio / リモート→HTTP / 本番運用→Streamable HTTP が推奨"),
        ]
    },
    # Slide 10: Security
    {
        "type": "content",
        "title": "セキュリティと権限管理",
        "bullets": [
            ("最小権限", "サーバーごとに必要最小限の権限だけを付与する原則"),
            ("入力検証", "すべてのパラメータをサーバー側で検証。インジェクション攻撃を防止"),
            ("サンドボックス", "ファイルアクセスやネットワーク通信を制限された範囲内で実行"),
            ("ユーザー同意", "ツール実行前にユーザーの明示的な承認を要求（Human-in-the-loop）"),
        ]
    },
    # Slide 11: Integration Patterns
    {
        "type": "content",
        "title": "MCP の活用パターン",
        "bullets": [
            ("開発ワークフロー", "GitHub + CI/CD + DB を連携し、Issue対応からデプロイまでAI駆動で実行"),
            ("データ分析", "DB + スプレッドシート + 可視化ツールを接続し、自然言語で分析指示"),
            ("業務自動化", "Slack + カレンダー + メール + CRM を統合し、日常業務をAIが支援"),
            ("ナレッジ管理", "社内Wiki + ドキュメント + コードベースを横断的にAIが検索・参照"),
        ]
    },
    # Slide 12: Ecosystem
    {
        "type": "content",
        "title": "MCP エコシステム",
        "bullets": [
            ("公式レジストリ", "github.com/modelcontextprotocol/servers に公式・コミュニティサーバー集約"),
            ("対応クライアント", "Claude Desktop / Claude Code / Cursor / Zed / Sourcegraph 等が対応"),
            ("SDK", "Python / TypeScript / Java / Kotlin / C# / Swift / Go 等多言語に対応"),
            ("コミュニティ", "オープンソースで急速に成長。毎月新しいサーバーやツールが公開"),
        ]
    },
    # Slide 13: Best Practices
    {
        "type": "content",
        "title": "MCP サーバー開発のベストプラクティス",
        "bullets": [
            ("ツール設計", "1ツール1機能の原則。明確な名前と詳細な説明文でAIの理解を助ける"),
            ("エラー処理", "is_error フラグ付きの構造化エラー応答。AIが適切にリカバリ可能に"),
            ("テスト", "MCP Inspector ツールで動作確認。単体テストとE2Eテストを組み合わせる"),
            ("ドキュメント", "各ツール・リソースの用途を詳細に記述。AIがいつ使うべきか判断できるように"),
        ]
    },
    # Slide 14: Summary
    {
        "type": "summary",
        "title": "まとめ",
        "points": [
            "MCPはAIと外部ツール・データを接続するオープンな標準プロトコル",
            "Tools / Resources / Prompts の3機能でAIの能力を大幅に拡張",
            "Python / TypeScript SDKで簡単にカスタムサーバーを構築可能",
            "セキュリティを重視した設計：最小権限・入力検証・ユーザー同意",
            "活発なエコシステムが急成長中。開発・分析・業務自動化に幅広く活用",
        ]
    },
]

num_slides = len(slides_data)

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
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="MCPTheme">
  <a:themeElements>
    <a:clrScheme name="MCP">
      <a:dk1><a:srgbClr val="0D1117"/></a:dk1>
      <a:lt1><a:srgbClr val="F0F6FC"/></a:lt1>
      <a:dk2><a:srgbClr val="161B22"/></a:dk2>
      <a:lt2><a:srgbClr val="C9D1D9"/></a:lt2>
      <a:accent1><a:srgbClr val="00BFA6"/></a:accent1>
      <a:accent2><a:srgbClr val="58A6FF"/></a:accent2>
      <a:accent3><a:srgbClr val="BC8CFF"/></a:accent3>
      <a:accent4><a:srgbClr val="F0883E"/></a:accent4>
      <a:accent5><a:srgbClr val="F7C948"/></a:accent5>
      <a:accent6><a:srgbClr val="3FB950"/></a:accent6>
      <a:hlink><a:srgbClr val="58A6FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="BC8CFF"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="MCP">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="MCP">
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

    # Background
    shapes.append(make_filled_rect(sid, "bg", 0, 0, SLIDE_W, SLIDE_H, BG_DARK))
    sid += 1

    # Top accent stripe
    shapes.append(make_filled_rect(sid, "stripe_top", 0, 0, SLIDE_W, emu(0.5), ACCENT_TEAL))
    sid += 1

    # Decorative left bar
    shapes.append(make_filled_rect(sid, "left_bar", emu(1.5), emu(3), emu(0.15), emu(5), ACCENT_BLUE))
    sid += 1

    # Connection diagram box
    shapes.append(make_rounded_rect(sid, "diagram_box", emu(2.5), emu(2.2), emu(5), emu(2.5), "161B22", "80000"))
    sid += 1

    # Diagram text: AI <-> MCP <-> Tools
    diagram_paras = [
        make_paragraph([make_text_run("AI Model", 13, ACCENT_BLUE, bold=True, font="Consolas")], "ctr"),
        make_paragraph([make_text_run("  <-->  MCP  <-->", 13, ACCENT_TEAL, font="Consolas")], "ctr"),
        make_paragraph([make_text_run("External Tools", 13, ACCENT_ORANGE, bold=True, font="Consolas")], "ctr"),
    ]
    shapes.append(make_textbox(sid, "diagram_text", emu(2.7), emu(2.5), emu(4.6), emu(2), diagram_paras, "ctr"))
    sid += 1

    # Main title
    title_paras = [
        make_paragraph([make_text_run(data["title"], 36, TEXT_WHITE, bold=True)], "ctr"),
    ]
    shapes.append(make_textbox(sid, "title", emu(5), emu(3.5), emu(20), emu(3), title_paras, "ctr"))
    sid += 1

    # Subtitle
    sub_paras = [
        make_paragraph([make_text_run(data["subtitle"], 16, ACCENT_TEAL)], "ctr"),
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
    shapes.append(make_filled_rect(sid, "bottom_line", emu(8), emu(11), emu(18), emu(0.1), ACCENT_TEAL))
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

    # Teal accent line under header
    shapes.append(make_filled_rect(sid, "accent_line", 0, emu(2.5), SLIDE_W, emu(0.08), ACCENT_TEAL))
    sid += 1

    # Slide number
    num_paras = [make_paragraph([make_text_run(f"{slide_num:02d}", 14, ACCENT_TEAL, bold=True)], "l")]
    shapes.append(make_textbox(sid, "slide_num", emu(1), emu(0.3), emu(2), emu(1.5), num_paras, "t"))
    sid += 1

    # Title
    title_paras = [make_paragraph([make_text_run(data["title"], 28, TEXT_WHITE, bold=True)], "l")]
    shapes.append(make_textbox(sid, "title", emu(3), emu(0.4), emu(28), emu(2), title_paras, "t"))
    sid += 1

    # Bullet items
    bullet_count = len(data["bullets"])
    y_start = emu(3.2)
    # Adjust item height based on number of bullets
    if bullet_count <= 3:
        item_height = emu(3.2)
    else:
        item_height = emu(2.8)

    for i, (label, desc) in enumerate(data["bullets"]):
        y_pos = y_start + i * item_height

        # Bullet card background
        shapes.append(make_rounded_rect(sid, f"card_{i}", emu(1.5), y_pos, emu(30), item_height - emu(0.3), BG_CARD, "60000"))
        sid += 1

        # Left accent bar for card
        colors = [ACCENT_TEAL, ACCENT_BLUE, ACCENT_PURPLE, ACCENT_ORANGE]
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
    shapes.append(make_filled_rect(sid, "footer_line", emu(1.5), SLIDE_H - emu(1), emu(30), emu(0.05), "161B22"))
    sid += 1

    # Page number
    page_paras = [make_paragraph([make_text_run(f"{slide_num} / {num_slides}", 10, TEXT_GRAY)], "r")]
    shapes.append(make_textbox(sid, "page_num", emu(27), SLIDE_H - emu(0.9), emu(5), emu(0.8), page_paras, "t"))
    sid += 1

    return shapes


def generate_code_slide(data, slide_num):
    """Generate a code example slide."""
    shapes = []
    sid = 2

    # Background
    shapes.append(make_filled_rect(sid, "bg", 0, 0, SLIDE_W, SLIDE_H, BG_DARK))
    sid += 1

    # Header bar
    shapes.append(make_filled_rect(sid, "header_bar", 0, 0, SLIDE_W, emu(2.5), BG_MEDIUM))
    sid += 1

    # Accent line under header
    shapes.append(make_filled_rect(sid, "accent_line", 0, emu(2.5), SLIDE_W, emu(0.08), ACCENT_TEAL))
    sid += 1

    # Slide number
    num_paras = [make_paragraph([make_text_run(f"{slide_num:02d}", 14, ACCENT_TEAL, bold=True)], "l")]
    shapes.append(make_textbox(sid, "slide_num", emu(1), emu(0.3), emu(2), emu(1.5), num_paras, "t"))
    sid += 1

    # Title
    title_paras = [make_paragraph([make_text_run(data["title"], 28, TEXT_WHITE, bold=True)], "l")]
    shapes.append(make_textbox(sid, "title", emu(3), emu(0.4), emu(28), emu(2), title_paras, "t"))
    sid += 1

    # Code block background
    code_y = emu(3.2)
    code_h = emu(9.5)
    shapes.append(make_rounded_rect(sid, "code_bg", emu(2), code_y, emu(29), code_h, "0D1117"))
    sid += 1

    # Code border
    shapes.append(make_filled_rect(sid, "code_border_top", emu(2), code_y, emu(29), emu(0.06), "30363D"))
    sid += 1
    shapes.append(make_filled_rect(sid, "code_border_left", emu(2), code_y, emu(0.06), code_h, "30363D"))
    sid += 1

    # Terminal dots
    dot_paras = [make_paragraph([
        make_text_run("  ", 10, ACCENT_RED, font="Consolas"),
    ], "l")]
    shapes.append(make_textbox(sid, "dots", emu(2.5), code_y + emu(0.2), emu(3), emu(0.6), dot_paras, "t"))
    sid += 1

    # File label
    file_paras = [make_paragraph([make_text_run("server.py", 10, TEXT_GRAY, font="Consolas")], "ctr")]
    shapes.append(make_textbox(sid, "file_label", emu(12), code_y + emu(0.2), emu(8), emu(0.6), file_paras, "t"))
    sid += 1

    # Code content
    color_map = {
        "import": ACCENT_PURPLE,
        "init": ACCENT_BLUE,
        "decorator": ACCENT_ORANGE,
        "func": ACCENT_TEAL,
        "doc": ACCENT_GREEN,
        "return": TEXT_LIGHT,
        "run": ACCENT_BLUE,
        "": TEXT_LIGHT,
    }

    code_paras = []
    for line_text, line_type in data["code_lines"]:
        color = color_map.get(line_type, TEXT_LIGHT)
        if line_text == "":
            code_paras.append(make_paragraph([make_text_run(" ", 13, TEXT_LIGHT, font="Consolas")], "l", spc_before=50))
        else:
            code_paras.append(make_paragraph([make_text_run(line_text, 13, color, font="Consolas")], "l", spc_before=50))

    shapes.append(make_textbox(sid, "code_text", emu(3), code_y + emu(1), emu(27), code_h - emu(1.5), code_paras, "t"))
    sid += 1

    # Footer line
    shapes.append(make_filled_rect(sid, "footer_line", emu(1.5), SLIDE_H - emu(1), emu(30), emu(0.05), "161B22"))
    sid += 1

    # Page number
    page_paras = [make_paragraph([make_text_run(f"{slide_num} / {num_slides}", 10, TEXT_GRAY)], "r")]
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
    shapes.append(make_filled_rect(sid, "top_accent", 0, 0, SLIDE_W, emu(0.1), ACCENT_TEAL))
    sid += 1

    # Title area
    shapes.append(make_filled_rect(sid, "title_bg", 0, emu(0.1), SLIDE_W, emu(2.5), BG_MEDIUM))
    sid += 1

    # Title
    title_paras = [make_paragraph([make_text_run(data["title"], 32, TEXT_WHITE, bold=True)], "ctr")]
    shapes.append(make_textbox(sid, "title", emu(2), emu(0.3), emu(29), emu(2), title_paras, "ctr"))
    sid += 1

    # Accent line
    shapes.append(make_filled_rect(sid, "accent_line", 0, emu(2.6), SLIDE_W, emu(0.08), ACCENT_TEAL))
    sid += 1

    # Summary points
    y_start = emu(3.5)
    for i, point in enumerate(data["points"]):
        y_pos = y_start + i * emu(2)
        colors = [ACCENT_TEAL, ACCENT_BLUE, ACCENT_PURPLE, ACCENT_ORANGE, ACCENT_GREEN]
        dot_color = colors[i % len(colors)]

        # Number circle
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
        make_text_run("MCP", 12, ACCENT_TEAL, bold=True),
        make_text_run("  -  Connecting AI to the world", 12, TEXT_GRAY),
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
    # Generate slide XMLs
    slide_xmls = []
    for i, sdata in enumerate(slides_data):
        slide_num = i + 1
        if sdata["type"] == "title":
            shapes = generate_title_slide(sdata)
        elif sdata["type"] == "summary":
            shapes = generate_summary_slide(sdata, slide_num)
        elif sdata["type"] == "code":
            shapes = generate_code_slide(sdata, slide_num)
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
    output = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mcp_server_guide.pptx")
    build_pptx(output)
