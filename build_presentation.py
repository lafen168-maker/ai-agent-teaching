# -*- coding: utf-8 -*-
"""
Generate professional 16:9 slides for AI Agent workshop:
Target: Teachers and School Administrators
Contexts: Classroom Management, Math Teaching (Nan-I Junior High), School Administration
"""
import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette Definitions
    C_DARK_NAVY = RGBColor(15, 23, 42)      # #0F172A
    C_CARD_BG = RGBColor(255, 255, 255)     # #FFFFFF
    C_PAGE_BG = RGBColor(248, 250, 252)     # #F8FAFC
    C_PRIMARY_BLUE = RGBColor(37, 99, 235)  # #2563EB
    C_INDIGO = RGBColor(79, 70, 229)        # #4F46E5
    C_EMERALD = RGBColor(5, 150, 105)       # #059669
    C_AMBER = RGBColor(217, 119, 6)         # #D97706
    C_PURPLE = RGBColor(124, 58, 237)       # #7C3AED
    C_TEXT_MAIN = RGBColor(30, 41, 59)      # #1E293B
    C_TEXT_MUTED = RGBColor(100, 116, 139)  # #64748B
    C_BORDER = RGBColor(226, 232, 240)      # #E2E8F0
    C_LIGHT_BLUE = RGBColor(239, 246, 255)  # #EFF6FF
    C_LIGHT_GREEN = RGBColor(236, 253, 245) # #ECFDF5
    C_LIGHT_AMBER = RGBColor(254, 243, 199) # #FEF3C7
    C_LIGHT_PURPLE = RGBColor(245, 243, 255)# #F5F3FF

    FONT_HEADING = "Microsoft JhengHei"
    FONT_BODY = "Microsoft JhengHei"

    def set_bg(slide, color=C_PAGE_BG):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.color.rgb = color
        # Send to back if needed, but since it's added first, it's at the back

    def add_header(slide, category, title, subtitle):
        # Category pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(2.8), Inches(0.35))
        pill.fill.solid()
        pill.fill.fore_color.rgb = C_PRIMARY_BLUE
        pill.line.fill.background()
        tf = pill.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = category
        p.font.name = FONT_BODY
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        # Title text box
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = C_DARK_NAVY

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.name = FONT_BODY
            p2.font.size = Pt(13)
            p2.font.color.rgb = C_TEXT_MUTED
            p2.space_before = Pt(4)

    def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(1.5)
        else:
            card.line.fill.background()
        return card

    # ==================== SLIDE 1: COVER ====================
    s1 = prs.slides.add_slide(blank_layout)
    set_bg(s1, C_DARK_NAVY)

    # Accent decorative box
    dec = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.2), Inches(3.2), Inches(0.4))
    dec.fill.solid()
    dec.fill.fore_color.rgb = C_PRIMARY_BLUE
    dec.line.fill.background()
    p = dec.text_frame.paragraphs[0]
    p.text = "AI 時代教師與教育行政專用指南"
    p.font.name = FONT_BODY
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    tb1 = s1.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(3.0))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "AI Agent 核心概念與實踐指南"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(255, 255, 255)

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "從單向提示詞（Prompt）到自主代理人（AI Agent）：班級經營 × 數學教學 × 學校行政全方位賦能"
    p1_sub.font.name = FONT_BODY
    p1_sub.font.size = Pt(17)
    p1_sub.font.color.rgb = RGBColor(148, 163, 184)
    p1_sub.space_before = Pt(12)

    # Bottom cards on cover
    features = [
        ("💡 觀念革命", "超越問答機器人，理解具備感知、記憶、工具與反思的 Agent 架構"),
        ("🧩 核心辨析", "透徹拆解 Prompt、Skill 與 Project 的階層差異與組合威力"),
        ("🏫 三大場景", "實戰落地於「親師溝通與常規」、「數學錯因診斷出題」、「公文行政管考」")
    ]
    for idx, (ft_title, ft_desc) in enumerate(features):
        c_left = Inches(0.8 + idx * 3.9)
        card = add_card(s1, c_left, Inches(4.8), Inches(3.7), Inches(1.8), bg_color=RGBColor(30, 41, 59), border_color=RGBColor(51, 65, 85))
        tb = s1.shapes.add_textbox(c_left + Inches(0.2), Inches(4.9), Inches(3.3), Inches(1.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = ft_title
        p.font.name = FONT_HEADING
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = RGBColor(241, 245, 249)
        p2 = tf.add_paragraph()
        p2.text = ft_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(12)
        p2.font.color.rgb = RGBColor(203, 213, 225)
        p2.space_before = Pt(6)

    # ==================== SLIDE 2: AGENDA ====================
    s2 = prs.slides.add_slide(blank_layout)
    set_bg(s2)
    add_header(s2, "研習導航", "研習架構與四大核心模組", "協助現場教師與行政人員建立系統思維，不只學會下咒語，更學會打造數位分身")

    agenda_items = [
        ("01", "什麼是 AI Agent？", "從「聊天式 LLM」到「具自主行動力的 Agent」，解析大腦、感知、工具與記憶四大基石。", C_PRIMARY_BLUE, C_LIGHT_BLUE),
        ("02", "三劍客差異詳解", "徹底分清 Prompt（便利貼）、Skill（SOP能力包）與 Project（作戰指揮室）的關係。", C_AMBER, C_LIGHT_AMBER),
        ("03", "名詞大解鎖與 AGI", "搞懂 AGI、Multi-Agent、MCP 協議、Tool Calling，掌握邁向未來智慧校園的關鍵演進。", C_PURPLE, C_LIGHT_PURPLE),
        ("04", "三大教育實戰落地", "深入解析「教師班級經營」、「國中數學專業教學」、「學校行政公文管考」的落地工作流。", C_EMERALD, C_LIGHT_GREEN)
    ]
    for idx, (num, title, desc, color, bg_col) in enumerate(agenda_items):
        c_left = Inches(0.8 + idx * 2.95)
        add_card(s2, c_left, Inches(1.9), Inches(2.8), Inches(4.8), bg_color=C_CARD_BG, border_color=color)
        
        # Number header
        num_card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c_left + Inches(0.2), Inches(2.1), Inches(0.9), Inches(0.6))
        num_card.fill.solid()
        num_card.fill.fore_color.rgb = color
        num_card.line.fill.background()
        p = num_card.text_frame.paragraphs[0]
        p.text = num
        p.font.name = FONT_HEADING
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        tb = s2.shapes.add_textbox(c_left + Inches(0.2), Inches(2.9), Inches(2.4), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(17)
        p.font.bold = True
        p.font.color.rgb = C_DARK_NAVY
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(13)
        p2.font.color.rgb = C_TEXT_MUTED
        p2.space_before = Pt(12)

    # ==================== SLIDE 3: WHY AI AGENT? ====================
    s3 = prs.slides.add_slide(blank_layout)
    set_bg(s3)
    add_header(s3, "模組一：觀念突破", "為什麼我們需要 AI Agent？從「問答助手」到「自主代理」", "生成式 AI 正歷經第二波典範轉移：從單一對話框走向多步驟自主閉環工作流")

    # Left card: Traditional Chatbot
    add_card(s3, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0), bg_color=C_CARD_BG)
    tb_left = s3.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_l = tb_left.text_frame
    tf_l.word_wrap = True
    p = tf_l.paragraphs[0]
    p.text = "傳統聊天機器人（Chatbot / LLM）"
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(220, 38, 38)
    
    bullets_left = [
        ("互動模式：", "「一問一答」單次響應，打一下動一下"),
        ("角色定位：", "一本有問必答的百科全書、文案修飾助手"),
        ("嚴重痛點：", "記不住班級脈絡、不會主動看外部檔案、不會使用工具檢驗自己的算式"),
        ("教師感受：", "老師必須自己手動複製貼上、反覆核對、手工調整格式，負擔依然沉重"),
        ("教育實例：", "問「幫我出因數分解題目」，只會給 5 道死板題目，無法核對解答是否精確或符合特定班級程度")
    ]
    for b_title, b_desc in bullets_left:
        p = tf_l.add_paragraph()
        p.text = f"• {b_title} {b_desc}"
        p.font.name = FONT_BODY
        p.font.size = Pt(12.5)
        p.font.color.rgb = C_TEXT_MAIN
        p.space_before = Pt(8)

    # Right card: AI Agent
    add_card(s3, Inches(6.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_LIGHT_BLUE, border_color=C_PRIMARY_BLUE)
    tb_right = s3.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_r = tb_right.text_frame
    tf_r.word_wrap = True
    p = tf_r.paragraphs[0]
    p.text = "智慧代理人（AI Agent）"
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY_BLUE
    
    bullets_right = [
        ("互動模式：", "「目標驅動（Goal-Driven）」，給予目標即可自主拆解並執行多步流程"),
        ("角色定位：", "能獨當一面的數位教學助教、實習行政主任"),
        ("核心特質：", "具備感知（讀取檔案/表格）、規劃（ReAct邏輯）、工具使用（算式驗證/搜尋）、記憶與反思修正"),
        ("教師感受：", "交代任務後，AI 自主翻閱課綱手冊、呼叫程式出題、交叉驗證答案無誤後排版交付"),
        ("教育實例：", "「請根據 701 班段考第 1 章錯題，為落後群設計 3 題階梯式因數分解練習並附解析」，Agent 完整閉環完成！")
    ]
    for b_title, b_desc in bullets_right:
        p = tf_r.add_paragraph()
        p.text = f"✔ {b_title} {b_desc}"
        p.font.name = FONT_BODY
        p.font.size = Pt(12.5)
        p.font.color.rgb = C_DARK_NAVY
        p.space_before = Pt(8)

    # ==================== SLIDE 4: ARCHITECTURE ====================
    s4 = prs.slides.add_slide(blank_layout)
    set_bg(s4)
    add_header(s4, "模組一：運作原理", "AI Agent 的四大核心要素：它是如何運作的？", "Agent = 大腦（LLM推理） + 眼睛（感知輸入） + 雙手雙腳（工具調用） + 筆記本（長期記憶）")

    arch_items = [
        ("1. 感知環境 (Perception)", "接收多元訊號", "學生作業文字、親師聯絡簿訊息、校務系統成績表、掃描 PDF 公文檔、課本目錄等。Agent 能理解並結構化周遭情境。", C_PRIMARY_BLUE, C_LIGHT_BLUE),
        ("2. 規劃推理 (Planning)", "任務拆解與邏輯思維", "運用思維鏈（CoT）或 ReAct 機制：先做什麼、再調用什麼資料、遇到錯誤如何轉向。將大目標拆解為可落地的子任務。", C_AMBER, C_LIGHT_AMBER),
        ("3. 工具與行動 (Tools)", "突破模型能力限制", "調用計算機避免計算幻覺、檢索最新法規、讀寫 Word/PPT/PDF 檔案、發送通知郵件或執行 Python 腳本產生題目排版。", C_EMERALD, C_LIGHT_GREEN),
        ("4. 記憶與反思 (Memory & Reflection)", "跨輪次經驗沉澱", "記住該班學生常犯的錯因、行政長官偏好的簽呈格式；在執行後自我檢驗「答案算對了嗎？語氣是否過於嚴厲？」，並自動修正。", C_PURPLE, C_LIGHT_PURPLE)
    ]
    for idx, (title, sub, body, col, bg_c) in enumerate(arch_items):
        c_left = Inches(0.8 + idx * 2.95)
        add_card(s4, c_left, Inches(1.9), Inches(2.8), Inches(4.9), bg_color=C_CARD_BG, border_color=col)
        
        # Tag inside
        tag = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c_left + Inches(0.2), Inches(2.1), Inches(2.4), Inches(0.4))
        tag.fill.solid()
        tag.fill.fore_color.rgb = bg_c
        tag.line.color.rgb = col
        p = tag.text_frame.paragraphs[0]
        p.text = sub
        p.font.name = FONT_BODY
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = col
        p.alignment = PP_ALIGN.CENTER

        tb = s4.shapes.add_textbox(c_left + Inches(0.2), Inches(2.6), Inches(2.4), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(15.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK_NAVY
        p2 = tf.add_paragraph()
        p2.text = body
        p2.font.name = FONT_BODY
        p2.font.size = Pt(12)
        p2.font.color.rgb = C_TEXT_MAIN
        p2.space_before = Pt(10)

    # ==================== SLIDE 5: TRIAD DIFFERENCE ====================
    s5 = prs.slides.add_slide(blank_layout)
    set_bg(s5)
    add_header(s5, "模組二：概念剖析", "搞懂核心三概念：Prompt vs. Skill vs. Project", "這是現場教師最容易混淆的地方：它們不是競爭關係，而是層層包容的增強關係！")

    triad_data = [
        ("Prompt（提示詞）", "單次執行的即時指令", "方向盤上的每一次轉向 / 口頭吩咐",
         ["生命週期：用完即丟，不留痕跡", "本質：觸發 AI 單一動作的輸入文本", "缺點：每次都要重新交待背景與限制", "學校實例：「幫我把這段文字改成客氣一點的語氣發給家長」"],
         C_AMBER, C_LIGHT_AMBER),
        ("Skill（技能模組）", "標準化可重複調用的 SOP", "副駕駛座的專家教練 / 工具箱",
         ["生命週期：常駐可用，隨調隨用", "本質：打包好的專家工作流程與驗證機制", "優勢：按需動態載入，大幅提升產出穩定度", "學校實例：「親師溝通信生成 Skill」、「國中數學階梯出題 Skill」"],
         C_PRIMARY_BLUE, C_LIGHT_BLUE),
        ("Project（專案空間）", "脈絡上下文與長期知識庫", "整個專屬的作戰指揮室 / 辦公室",
         ["生命週期：長期持續，貫穿整個學期", "本質：固定掛載的資料夾、名冊與歷史記憶", "威力：讓 AI 成為熟悉特定班級或處室的專屬員工", "學校實例：「701班班級經營專案」、「114學年度教務處公文專案」"],
         C_EMERALD, C_LIGHT_GREEN)
    ]

    for idx, (t_name, t_sub, t_metaphor, t_points, col, bg_col) in enumerate(triad_data):
        c_left = Inches(0.8 + idx * 3.9)
        add_card(s5, c_left, Inches(1.9), Inches(3.7), Inches(5.0), bg_color=C_CARD_BG, border_color=col)

        # Header box
        hdr = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_left, Inches(1.9), Inches(3.7), Inches(0.8))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = col
        hdr.line.fill.background()
        p = hdr.text_frame.paragraphs[0]
        p.text = t_name
        p.font.name = FONT_HEADING
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        p_sub = hdr.text_frame.add_paragraph()
        p_sub.text = t_sub
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = RGBColor(241, 245, 249)
        p_sub.alignment = PP_ALIGN.CENTER

        tb = s5.shapes.add_textbox(c_left + Inches(0.2), Inches(2.8), Inches(3.3), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p_meta = tf.paragraphs[0]
        p_meta.text = f"🎯 譬喻：{t_metaphor}"
        p_meta.font.name = FONT_HEADING
        p_meta.font.size = Pt(13)
        p_meta.font.bold = True
        p_meta.font.color.rgb = col

        for pt in t_points:
            p_pt = tf.add_paragraph()
            p_pt.text = f"• {pt}"
            p_pt.font.name = FONT_BODY
            p_pt.font.size = Pt(11.5)
            p_pt.font.color.rgb = C_TEXT_MAIN
            p_pt.space_before = Pt(6)

    # ==================== SLIDE 6: METAPHORS ====================
    s6 = prs.slides.add_slide(blank_layout)
    set_bg(s6)
    add_header(s6, "模組二：直覺比喻", "兩大生動譬喻：讓老師與行政同仁一秒聽懂！", "用生活與校園經驗，徹底打通 Prompt、Skill、Project 與 Agent 的連結")

    # Metaphor 1: The Restaurant / Kitchen
    add_card(s6, Inches(0.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_CARD_BG)
    tb1 = s6.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "🍳 譬喻一：米其林廚房物語"
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY_BLUE

    kitchen_items = [
        ("Prompt（提示詞）", "客人點單時的臨時備註：「今天的牛肉麵不要加酸菜、少放點鹽」。（單次指令，換桌就失效）"),
        ("Skill（技能）", "主廚親撰的「紅燒牛肉麵秘傳 SOP 食譜」。（熬湯火候、香料克數、下麵秒數精確定義，任何學徒照做都能維持水準）"),
        ("Project（專案）", "整間「牛肉麵旗艦餐廳」。（包含食材冷藏庫、進貨廠商資料、過去三個月的營收報表與廚房工作規則）"),
        ("AI Agent（代理人）", "主廚副手（Sous Chef）！（看到客人點單，主動翻開食譜，自己打開冰箱拿肉、開火煮麵、出餐前試喝一口確認鹹淡，有問題立刻重做）")
    ]
    for k_title, k_desc in kitchen_items:
        p = tf1.add_paragraph()
        p.text = f"🔸 {k_title}："
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_DARK_NAVY
        p.space_before = Pt(8)
        p2 = tf1.add_paragraph()
        p2.text = k_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(11.5)
        p2.font.color.rgb = C_TEXT_MUTED
        p2.space_before = Pt(2)

    # Metaphor 2: School Office
    add_card(s6, Inches(6.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_CARD_BG)
    tb2 = s6.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "🏫 譬喻二：學校行政實戰場景"
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = C_EMERALD

    school_items = [
        ("Prompt（提示詞）", "主任在走廊交代組長一句話：「幫我查一下教育部那份公文補助款是多少錢」。（講完就結束，沒附公文編號）"),
        ("Skill（技能）", "教務處的「校外教學標準作業查檢表與簽呈範本」。（審查資格、保險單據、家長同意書格式一應俱全）"),
        ("Project（專案）", "「114學年度畢業旅行專案資料夾」。（內有廠商契約書、家長聯絡名單、歷年問卷、經費預算表）"),
        ("AI Agent（代理人）", "兼具專業與主動性的「行政幹事」！（收到籌辦畢旅目標，主動調出資料夾、套用核銷與公文 Skill、核算遊覽車經費、並列印出完整簽呈草稿給主任批閱）")
    ]
    for s_title, s_desc in school_items:
        p = tf2.add_paragraph()
        p.text = f"🔹 {s_title}："
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_DARK_NAVY
        p.space_before = Pt(8)
        p2 = tf2.add_paragraph()
        p2.text = s_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(11.5)
        p2.font.color.rgb = C_TEXT_MUTED
        p2.space_before = Pt(2)

    # ==================== SLIDE 6B: PROJECT VS WORKSPACE FOLDER ====================
    s6b = prs.slides.add_slide(blank_layout)
    set_bg(s6b)
    add_header(s6b, "進階觀念：Antigravity 配置", "Project（任務團隊）與工作資料夾 Workspace（會議室）", "老師神比喻：團隊是無形的心智靈魂，資料夾是有形的實體辦公開會場地")

    # Left: Concept Cards
    add_card(s6b, Inches(0.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_CARD_BG)
    tb_c = s6b.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    p = tf_c.paragraphs[0]
    p.text = "🏢 核心定義：團隊 vs. 場地"
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = C_INDIGO

    team_vs_room = [
        ("👥 Project（專案）= 任務團隊", "定義了「我們是誰？目標是什麼？遵守什麼守則？」包含系統提示詞、掛載的 Skills 技能包、專案記憶與偏好。（例如：南一數學段考命題小組、校慶籌委會）"),
        ("🚪 Folder（工作資料夾）= 開會教室/會議室", "定義了「團隊今天在哪裡集合開會？桌面上擺放著哪些具體檔案？」代表電腦硬碟中的實體路徑（Workspace），限制 Agent 當下能讀寫的範圍。（例如：國中數學1上課本目錄）"),
        ("🛡️ 門禁哲學：安全沙盒邊界", "指定資料夾就是給 Agent 門禁卡！它只能在被指派的會議室內翻閱檔案，絕不能任意打開隔壁「人事室（考績）」或「註冊組（個資）」的大門。")
    ]
    for h_title, h_desc in team_vs_room:
        p = tf_c.add_paragraph()
        p.text = h_title
        p.font.name = FONT_HEADING
        p.font.size = Pt(12.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK_NAVY
        p.space_before = Pt(8)
        p2 = tf_c.add_paragraph()
        p2.text = h_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(11)
        p2.font.color.rgb = C_TEXT_MAIN
        p2.space_before = Pt(2)

    # Right: Flexibility in School
    add_card(s6b, Inches(6.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_LIGHT_BLUE, border_color=C_PRIMARY_BLUE)
    tb_f = s6b.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_f = tb_f.text_frame
    tf_f.word_wrap = True
    p = tf_f.paragraphs[0]
    p.text = "🏫 校園實務運作兩大靈活性"
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY_BLUE

    flex_cases = [
        ("🔄 靈活性一：同一個團隊，換到不同會議室", "你的「國中數學命題 Project」（已設定好命題心法與防呆驗算 Skill）：\n• 第一次段考：把工作資料夾設為「國中數學1上_因數分解/」\n• 第二次段考：換到「國中數學1上_一元一次方程式/」\n團隊 SOP 完全不用重寫，換個場地立刻開會！"),
        ("🏢 靈活性二：同一個會議室，借給不同團隊進駐", "學校有一個共用資料夾「114活動_校慶資料夾/」：\n• 上午：「學務處活動 Project」進駐，專注於排定競賽時程表\n• 下午：「總務處核銷 Project」進駐，專注於審核經費單據\n不同團隊帶著各自的專業靈魂，處理同一個資料夾中的檔案！")
    ]
    for h_title, h_desc in flex_cases:
        p = tf_f.add_paragraph()
        p.text = h_title
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_PRIMARY_BLUE
        p.space_before = Pt(10)
        p2 = tf_f.add_paragraph()
        p2.text = h_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(11)
        p2.font.color.rgb = C_DARK_NAVY
        p2.space_before = Pt(2)

    # ==================== SLIDE 7: TERMINOLOGY ====================
    s7 = prs.slides.add_slide(blank_layout)
    set_bg(s7)
    add_header(s7, "模組三：專業術語解密", "不再被名詞轟炸！教育人員必懂的 AI 關鍵字", "用最親切的語言解構業界最新技術名詞")

    terms = [
        ("LLM (大語言模型)", "基底語言大腦", "GPT-4、Claude、Gemini 等核心模型。擅長理解文字並接續生成，就像是「文字專家的大腦」，但缺乏身體與外部操作能力。", C_PRIMARY_BLUE),
        ("Tool / Function Calling", "雙手與外部工具", "AI 主動呼叫計算機、搜尋引擎、Python 執行環境、試算表軟體的能力。解決大模型算術差與資訊過期的大問題！", C_EMERALD),
        ("MCP (模型上下文協議)", "通用標準插座", "Model Context Protocol：由 Anthropic 發起的新興開放標準。讓 AI 能以統一規格插入學校資料庫、本地資料夾與 Google 雲端硬碟。", C_AMBER),
        ("Agentic Workflow", "工作流自我迭代", "不是直接噴出一段答案，而是經過「擬定大綱 -> 查證資料 -> 撰寫初稿 -> 自我校對除錯 -> 定稿」的思考迭代循環。", C_PURPLE),
        ("Multi-Agent (多代理人)", "處室分工協作團", "由多個專精不同領域的 Agent 協作，例如「出題 Agent」出完題目後，交給「驗算 Agent」解題校對，再由「排版 Agent」產出試卷。", C_PRIMARY_BLUE),
        ("Human-in-the-Loop", "人在迴圈中把關", "在關鍵節點（如成績確認、公文正式用印送出、敏感親師信發布）必須經由人類教師或行政主管親自確認核可，確保安全合規。", C_DARK_NAVY)
    ]

    for idx, (t_name, t_sub, t_desc, col) in enumerate(terms):
        row = idx // 3
        col_idx = idx % 3
        c_left = Inches(0.8 + col_idx * 3.9)
        c_top = Inches(1.9 + row * 2.5)

        add_card(s7, c_left, c_top, Inches(3.7), Inches(2.3), bg_color=C_CARD_BG, border_color=col)
        tb = s7.shapes.add_textbox(c_left + Inches(0.2), c_top + Inches(0.15), Inches(3.3), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p = tf.paragraphs[0]
        p.text = t_name
        p.font.name = FONT_HEADING
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = col

        p_sub = tf.add_paragraph()
        p_sub.text = f"【{t_sub}】"
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = C_TEXT_MUTED
        p_sub.space_before = Pt(2)

        p_desc = tf.add_paragraph()
        p_desc.text = t_desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(11)
        p_desc.font.color.rgb = C_TEXT_MAIN
        p_desc.space_before = Pt(4)

    # ==================== SLIDE 8: AGI AND THE PATH ====================
    s8 = prs.slides.add_slide(blank_layout)
    set_bg(s8)
    add_header(s8, "模組三：未來視野", "什麼是 AGI？為什麼「AI Agent」是通往 AGI 的關鍵路徑？", "通用人工智慧不再是科幻名詞，而是正在逐步發生的技術躍遷")

    # Left: What is AGI
    add_card(s8, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0), bg_color=C_CARD_BG)
    tb_l = s8.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p = tf_l.paragraphs[0]
    p.text = "🌐 什麼是 AGI（通用人工智慧）？"
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = C_PURPLE

    agi_points = [
        ("定義：", "Artificial General Intelligence，具備相當於人類心智的廣泛認知能力，能在未見過的新環境中自主學習與解決問題。"),
        ("對比 Narrow AI：", "以前的 AI 是「下圍棋就只會下圍棋、辨識貓就只會認貓」的專用弱人工智慧。"),
        ("對比現在的 LLM：", "目前的 LLM 懂非常多語言知識，但缺乏主動在現實世界達成目標的執行力。"),
        ("教育層面的啟示：", "未來 AI 不是單純回答課本問答，而是能像一位全方位的資深教育夥伴，因材施教並協調多元校務資源。")
    ]
    for b_title, b_desc in agi_points:
        p = tf_l.add_paragraph()
        p.text = f"• {b_title} {b_desc}"
        p.font.name = FONT_BODY
        p.font.size = Pt(12)
        p.font.color.rgb = C_TEXT_MAIN
        p.space_before = Pt(8)

    # Right: Evolution Path
    add_card(s8, Inches(6.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_LIGHT_PURPLE, border_color=C_PURPLE)
    tb_r = s8.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    p = tf_r.paragraphs[0]
    p.text = "🚀 AI 演進三階段階梯"
    p.font.name = FONT_HEADING
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = C_DARK_NAVY

    stages = [
        ("階段 1：生成式對話 (GenAI Chatbot)", "2022-2023", "透過提示詞單次問答，被動產生文本、翻譯與大綱，容易產生幻覺。"),
        ("階段 2：智慧代理人體系 (Agentic AI)", "2024-2026 (當前關鍵期)", "具備自主思考、分步規劃、調用工具、讀寫檔案與自我糾錯能力。這是邁向 AGI 最關鍵的基石！"),
        ("階段 3：通用人工智慧 (AGI)", "未來願景", "跨領域全自主通用智慧，能主動發現新科學規律、自主開發軟體與全面重塑教育學習型態。")
    ]
    for s_title, s_time, s_desc in stages:
        p = tf_r.add_paragraph()
        p.text = f"▶ {s_title} 【{s_time}】"
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_PURPLE
        p.space_before = Pt(10)
        p2 = tf_r.add_paragraph()
        p2.text = s_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(11.5)
        p2.font.color.rgb = C_TEXT_MAIN
        p2.space_before = Pt(2)

    # ==================== SLIDE 9: SCENARIO 1 - CLASSROOM MANAGEMENT ====================
    s9 = prs.slides.add_slide(blank_layout)
    set_bg(s9)
    add_header(s9, "模組四：實戰落地", "落地場景一：教師班級經營篇", "告別情緒內耗與繁雜瑣事，讓 AI Agent 成為導師最強大的同理心智囊")

    cm_cards = [
        ("痛點 1：親師溝通情緒高壓", "現場困境：學生在校起衝突或課業落後，導師既要客觀陳述事實，又要顧及家長玻璃心，稍有不慎易引發投訴。\n\nAgent 解方：建立【親師同理心溝通 Skill】，輸入客觀事實與衝突經過，Agent 自動套用非暴力溝通架構，產出兼具教育專業與同理溫度的 LINE/聯絡簿草稿。"),
        ("痛點 2：學生輔導紀錄凌亂", "現場困境：開學後各種個別輔導、晤談、常規違規事件零散記在便利貼或各處，期末做輔導紀錄卡或轉介時痛苦不堪。\n\nAgent 解方：建立【701班班級經營 Project】，長期累積學生個別觀察筆記。當需要輔導轉介時，Agent 一鍵彙整時間軸並生成客觀結構化晤談紀錄。"),
        ("痛點 3：班級常規與正向激勵", "現場困境：常規競賽、打掃評分、加扣分紀錄繁瑣，導師很難即時發現默默做好事的「隱形好學生」。\n\nAgent 解方：Agent 自動統整每週生活常規數據，自動挑選「本週進步之星」，生成給家長的正向表揚小卡，強化親師信任感。")
    ]
    for idx, (ctitle, cdesc) in enumerate(cm_cards):
        c_left = Inches(0.8 + idx * 3.9)
        add_card(s9, c_left, Inches(1.9), Inches(3.7), Inches(5.0), bg_color=C_CARD_BG, border_color=C_PRIMARY_BLUE)
        tb = s9.shapes.add_textbox(c_left + Inches(0.2), Inches(2.1), Inches(3.3), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = ctitle
        p.font.name = FONT_HEADING
        p.font.size = Pt(15.5)
        p.font.bold = True
        p.font.color.rgb = C_PRIMARY_BLUE
        
        parts = cdesc.split("\n\n")
        for part in parts:
            p2 = tf.add_paragraph()
            p2.text = part
            p2.font.name = FONT_BODY
            p2.font.size = Pt(11.5)
            p2.font.color.rgb = C_TEXT_MAIN
            p2.space_before = Pt(8)

    # ==================== SLIDE 10: SCENARIO 1 DEMO ====================
    s10 = prs.slides.add_slide(blank_layout)
    set_bg(s10)
    add_header(s10, "模組四：實戰落地", "班級經營 Agent 示範：從零散衝突到高情商親師信", "示範如何將 Prompt、Skill 與 Project 完美結合於真實校園突發事件")

    # Left: Raw Input & Skill Rule
    add_card(s10, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0), bg_color=C_CARD_BG)
    tb_l = s10.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p = tf_l.paragraphs[0]
    p.text = "📥 教師輸入與 Skill 規範"
    p.font.name = FONT_HEADING
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = C_AMBER

    p1 = tf_l.add_paragraph()
    p1.text = "【老師隨手打下的情境事實（Prompt Input）】"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = C_DARK_NAVY
    p1.space_before = Pt(6)

    p1_box = tf_l.add_paragraph()
    p1_box.text = "「小明今天在體育課跟小華搶籃球起口角，小明推了小華一把導致小華手肘擦傷。小明一開始不認錯，後來在導師室哭了說小華罵他矮。小明媽媽平時防衛心很強，常常覺得老師針對小明。」"
    p1_box.font.name = FONT_BODY
    p1_box.font.size = Pt(11.5)
    p1_box.font.color.rgb = C_TEXT_MAIN
    p1_box.space_before = Pt(3)

    p2 = tf_l.add_paragraph()
    p2.text = "【掛載之 Skill 規範（家長溝通三明治架構）】"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(12)
    p2.font.bold = True
    p2.font.color.rgb = C_PRIMARY_BLUE
    p2.space_before = Pt(10)

    p2_box = tf_l.add_paragraph()
    p2_box.text = "1. 首段同理與優點認可（肯定小明平時的榮譽感）\n2. 客觀還原事件經過（不帶批判字眼，只陳述動作與雙方情緒）\n3. 說明學校當下妥善處置與學生後續反思\n4. 提出親師攜手合作的引導邀請（共同幫助孩子練習情緒管理）"
    p2_box.font.name = FONT_BODY
    p2_box.font.size = Pt(11)
    p2_box.font.color.rgb = C_TEXT_MUTED
    p2_box.space_before = Pt(3)

    # Right: Output
    add_card(s10, Inches(6.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_LIGHT_BLUE, border_color=C_PRIMARY_BLUE)
    tb_r = s10.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    p = tf_r.paragraphs[0]
    p.text = "📤 AI Agent 最終生成的親師通訊草稿"
    p.font.name = FONT_HEADING
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY_BLUE

    output_text = (
        "小明媽媽您好：\n\n"
        "小明在班上一直非常有正義感且熱愛運動，這是非常可貴的特質。今天想跟媽媽分享一件下午在體育課發生的小插曲：\n\n"
        "今天在籃球場上，孩子們因為搶球發生口角，小明一時情急推了小華一下，小華手肘有輕微擦傷，已立刻至健康中心消毒冰敷。老師在導師室陪伴小明平復心情時，小明委屈落淚，主動告訴我對方的話讓他感到自尊受損。\n\n"
        "老師肯定小明願意誠實說出心裡受傷的感受，同時也引導他思考：用肢體推人可能造成危險。小明已主動向小華表達歉意。\n\n"
        "孩子正在學習如何表達憤怒，今晚回家時請媽媽也給小明一個擁抱，我們可以一起引導孩子：「感到生氣時，有哪些更好的表達方式？」感謝媽媽與學校一路以來的並肩支持！"
    )
    p_out = tf_r.add_paragraph()
    p_out.text = output_text
    p_out.font.name = FONT_BODY
    p_out.font.size = Pt(10.8)
    p_out.font.color.rgb = C_DARK_NAVY
    p_out.space_before = Pt(6)

    # ==================== SLIDE 11: SCENARIO 2 - MATH TEACHING ====================
    s11 = prs.slides.add_slide(blank_layout)
    set_bg(s11)
    add_header(s11, "模組四：實戰落地", "落地場景二：國中數學教學專業篇", "參考南一版國中數學 1 上（因數分解與負數運算）：讓 AI 成為精準診斷的數學助教")

    math_cards = [
        ("痛點 1：個別化程度落差極大", "現場困境：國一新生剛進國中，因數分解與質數判別一教，前段學生覺得無聊，後段學生連短除法都卡關。\n\nAgent 解方：建立【南一數學 1 上 Project】，內建教材章節目錄。利用【認知鷹架出題 Skill】，一鍵生成「基礎暖身題」、「精熟挑戰題」與「生活素養題」三種階梯難度，並配備自檢步驟。"),
        ("痛點 2：學生錯題盲點診斷困難", "現場困境：批改 35 本作業，只知道學生算錯，但沒時間一一細究是「負負得正符號法則搞混」還是「質因數分解漏掉了 3」？\n\nAgent 解方：建立【數學錯因診斷 Agent】，輸入學生的錯誤計算過程，Agent 自動標註核心迷思概念（Misconception），並自動推送 1 題專屬的「概念除錯微練習」。"),
        ("痛點 3：素養導向情境題設計耗時", "現場困境：傳統題目太死板，想出符合 108 課綱的素養情境題，老師往往要花整個週末絞盡腦汁。\n\nAgent 解方：結合【生活情境出題 Skill】，指定將「最大公因數」結合「校慶運動會分組」或「烘焙坊切蛋糕」情境，自動生成題幹、圖表描述與標準詳解。")
    ]
    for idx, (ctitle, cdesc) in enumerate(math_cards):
        c_left = Inches(0.8 + idx * 3.9)
        add_card(s11, c_left, Inches(1.9), Inches(3.7), Inches(5.0), bg_color=C_CARD_BG, border_color=C_EMERALD)
        tb = s11.shapes.add_textbox(c_left + Inches(0.2), Inches(2.1), Inches(3.3), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = ctitle
        p.font.name = FONT_HEADING
        p.font.size = Pt(15.5)
        p.font.bold = True
        p.font.color.rgb = C_EMERALD
        
        parts = cdesc.split("\n\n")
        for part in parts:
            p2 = tf.add_paragraph()
            p2.text = part
            p2.font.name = FONT_BODY
            p2.font.size = Pt(11.5)
            p2.font.color.rgb = C_TEXT_MAIN
            p2.space_before = Pt(8)

    # ==================== SLIDE 12: SCENARIO 2 DEMO ====================
    s12 = prs.slides.add_slide(blank_layout)
    set_bg(s12)
    add_header(s12, "模組四：實戰落地", "數學教學實戰示範：因數分解與錯因診斷閉環", "以南一版第 1 章「因數分解與分數運算」為例，展示 Agent 的多步診斷與出題")

    # Left: Student Error & Diagnostic
    add_card(s12, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0), bg_color=C_CARD_BG)
    tb_l = s12.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p = tf_l.paragraphs[0]
    p.text = "🔍 學生作答輸入與 Agent 錯因剖析"
    p.font.name = FONT_HEADING
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = RGBColor(220, 38, 38)

    p1 = tf_l.add_paragraph()
    p1.text = "【題目】：求 72 與 108 的標準分解式及最大公因數 (GCD)"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = C_DARK_NAVY
    p1.space_before = Pt(6)

    p1_box = tf_l.add_paragraph()
    p1_box.text = "【學生計算過程】：\n• 72 = 8 × 9 = 2^3 × 3^2 (正確)\n• 108 = 4 × 27 = 2^2 × 3^3 (正確)\n• 學生寫 GCD(72, 108) = 2^3 × 3^3 = 216 ❌"
    p1_box.font.name = FONT_BODY
    p1_box.font.size = Pt(11.5)
    p1_box.font.color.rgb = C_TEXT_MAIN
    p1_box.space_before = Pt(3)

    p2 = tf_l.add_paragraph()
    p2.text = "【Agent 智能診斷分析結果】：\n1. 基礎乘法與分解正確，運算能力良好。\n2. 核心迷思概念：將「最大公因數 (GCD)」與「最小公倍數 (LCM)」的指數選取規則顛倒！誤選了較高的次方而非共同的最低次方。\n3. 缺乏常識檢驗意識：公因數居然大於原本的數字 (216 > 72)，卻未能自我發覺。"
    p2.font.name = FONT_BODY
    p2.font.size = Pt(11.5)
    p2.font.color.rgb = RGBColor(185, 28, 28)
    p2.space_before = Pt(8)

    # Right: Targeted Remediation Output
    add_card(s12, Inches(6.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_LIGHT_GREEN, border_color=C_EMERALD)
    tb_r = s12.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    p = tf_r.paragraphs[0]
    p.text = "🎯 Agent 自主產生之微補救學習單"
    p.font.name = FONT_HEADING
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = C_EMERALD

    math_remedy = (
        "【第一步：啟發式提問引導 (Socratic Prompt)】\n"
        "「小凱，72 的『因數』有可能比 72 本身還要大嗎？如果算出來是 216，它還能整除 72 嗎？」\n\n"
        "【第二步：觀念口訣與規則對照】\n"
        "• 最大公因數 (GCD) 取【共同質因數】中【次方較小】者！\n"
        "  -> 2 的次方取 min(3, 2) = 2^2\n"
        "  -> 3 的次方取 min(2, 3) = 3^2\n"
        "  -> GCD(72, 108) = 2^2 × 3^2 = 36 ✔\n\n"
        "【第三步：同質結構即時檢核題 (Agent 自動出題)】\n"
        "練習題：已知 A = 2^2 × 3 × 5^2，B = 2 × 3^3 × 5。\n"
        "請圈出 A 與 B 的最大公因數之標準分解式，並用短除法驗證你的答案。"
    )
    p_out = tf_r.add_paragraph()
    p_out.text = math_remedy
    p_out.font.name = FONT_BODY
    p_out.font.size = Pt(11)
    p_out.font.color.rgb = C_DARK_NAVY
    p_out.space_before = Pt(6)

    # ==================== SLIDE 13: SCENARIO 3 - ADMINISTRATIVE WORK ====================
    s13 = prs.slides.add_slide(blank_layout)
    set_bg(s13)
    add_header(s13, "模組四：實戰落地", "落地場景三：學校行政工作篇", "行政同仁的超強外掛：從繁瑣公文、採購核銷到活動專案管考")

    admin_cards = [
        ("痛點 1：長篇公文讀不完、簽呈難寫", "現場困境：教育局動輒轉發 20 頁計畫與繁雜附件，行政組長需提煉精華並寫出合規「擬辦意見」，耗損大量時間。\n\nAgent 解方：建立【校園公文速讀與簽呈 Skill】，自動讀取公文 PDF，10 秒萃取「來文機關、核心要求、受補助金額、承辦處室、截止期限」，並自動產出公文簽核格式。"),
        ("痛點 2：大型活動流程管考容易漏東漏西", "現場困境：舉辦校慶運動會、戶外教育或家長日，要協調場地、保險、車輛、通知單，任何一個環節漏掉都是行政危機。\n\nAgent 解方：建立【校務活動管考 Project】，Agent 根據活動主題自動生成甘特圖工作時程、跨處室分工表與危機預防檢核清單，並自動排程提醒。"),
        ("痛點 3：經費概算與核銷規則複雜", "現場困境：主計室規範嚴格，差旅費、講座鐘點費、膳費各有標準，一旦算錯單據就遭退件重簽。\n\nAgent 解方：建立【經費核銷查核 Skill】，掛載教育部與各縣市政府主計標準，輸入活動預算表，Agent 自動查核單價是否超標、總額是否平衡。")
    ]
    for idx, (ctitle, cdesc) in enumerate(admin_cards):
        c_left = Inches(0.8 + idx * 3.9)
        add_card(s13, c_left, Inches(1.9), Inches(3.7), Inches(5.0), bg_color=C_CARD_BG, border_color=C_AMBER)
        tb = s13.shapes.add_textbox(c_left + Inches(0.2), Inches(2.1), Inches(3.3), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = ctitle
        p.font.name = FONT_HEADING
        p.font.size = Pt(15.5)
        p.font.bold = True
        p.font.color.rgb = C_AMBER
        
        parts = cdesc.split("\n\n")
        for part in parts:
            p2 = tf.add_paragraph()
            p2.text = part
            p2.font.name = FONT_BODY
            p2.font.size = Pt(11.5)
            p2.font.color.rgb = C_TEXT_MAIN
            p2.space_before = Pt(8)

    # ==================== SLIDE 14: SCENARIO 3 DEMO ====================
    s14 = prs.slides.add_slide(blank_layout)
    set_bg(s14)
    add_header(s14, "模組四：實戰落地", "行政工作實戰示範：長篇公文重點萃取與正式簽呈擬辦", "展示 Agent 如何將 15 頁的教育局補助函轉化為標準公文簽呈")

    # Left: Official Letter Input
    add_card(s14, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0), bg_color=C_CARD_BG)
    tb_l = s14.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p = tf_l.paragraphs[0]
    p.text = "📄 教育局原始來文重點 (Agent 解析)"
    p.font.name = FONT_HEADING
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = C_DARK_NAVY

    doc_info = [
        ("發文字號：", "高市教資字第 11430987600 號"),
        ("主旨摘要：", "檢送本市「114學年度推動智慧校園暨 AI 輔助教學實踐計畫」申辦作業須知，請各校於期限內提報。"),
        ("核心規範：", "每校最高補助經常門 15 萬元、資本門 10 萬元；需辦理至少 2 場教師研習及 1 個公開觀議課單元。"),
        ("截止時限：", "114 年 11 月 30 日前至線上系統填報並檢附計畫書。"),
        ("會辦處室：", "會辦總務處（設備採購）、會計室（經費審查）。")
    ]
    for b_title, b_desc in doc_info:
        p = tf_l.add_paragraph()
        p.text = f"• {b_title} {b_desc}"
        p.font.name = FONT_BODY
        p.font.size = Pt(12)
        p.font.color.rgb = C_TEXT_MAIN
        p.space_before = Pt(8)

    # Right: Generated Official Draft
    add_card(s14, Inches(6.8), Inches(1.9), Inches(5.7), Inches(5.0), bg_color=C_LIGHT_AMBER, border_color=C_AMBER)
    tb_r = s14.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    p = tf_r.paragraphs[0]
    p.text = "📑 Agent 自動產出之擬辦簽呈草案"
    p.font.name = FONT_HEADING
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = C_AMBER

    draft_text = (
        "簽 於 教務處資訊組\n\n"
        "主旨：有關教育局函知申辦「114學年度推動智慧校園暨 AI 輔助教學實踐計畫」乙案，簽請 核示。\n\n"
        "說明：\n"
        "一、依教育局 114 年 11 月 1 日高市教資字第 11430987600 號函辦理。\n"
        "二、本案最高可獲經常門 15 萬元、資本門 10 萬元補助，有助於升級本校數學與科技領域平板輔助教學設備。\n"
        "三、經費需配合辦理校內研習 2 場次，擬結合本校領域備課時間辦理。\n\n"
        "擬辦：\n"
        "一、奉核後，由資訊組會同數學領域教師於 11 月 20 日前撰擬計畫書。\n"
        "二、本案會請總務處協助設備訪價、會請會計室核配科目預算後，於期限前陳報局端。\n"
        "三、陳請 校長核示。"
    )
    p_out = tf_r.add_paragraph()
    p_out.text = draft_text
    p_out.font.name = FONT_BODY
    p_out.font.size = Pt(10.5)
    p_out.font.color.rgb = C_DARK_NAVY
    p_out.space_before = Pt(6)

    # ==================== SLIDE 15: MULTI-AGENT ====================
    s15 = prs.slides.add_slide(blank_layout)
    set_bg(s15)
    add_header(s15, "模組三：進階協作", "Multi-Agent（多代理人）：智慧校園的跨處室協作", "從單兵作戰到交響樂團：多個專屬 Agent 扮演不同角色，協同完成大型複雜任務")

    ma_roles = [
        ("教務備課 Agent", "專注於課綱對齊與出題", "讀取南一版教材目錄與 108 課綱，專門負責產出符合認知層次的試題草案與解析。", C_PRIMARY_BLUE),
        ("審題驗算 Agent", "專注於除錯與驗證", "專門執行 Python 程式碼解題，核對是否有無解題、文字悖論或計算錯誤，確保品質無虞。", C_EMERALD),
        ("美編排版 Agent", "專注於版面與格式", "將文字轉換為標準 Word 雙欄排版或 LaTeX 數學符號格式，生成可直接影印的學生卷與解答卷。", C_AMBER),
        ("行政調度 Agent", "專注於進度與溝通", "追蹤出題與審查進度，彙整段考進度表，並發送提醒給各命題召集人，確保如期交卷。", C_PURPLE)
    ]
    for idx, (rname, rsub, rdesc, col) in enumerate(ma_roles):
        c_left = Inches(0.8 + idx * 2.95)
        add_card(s15, c_left, Inches(1.9), Inches(2.8), Inches(4.9), bg_color=C_CARD_BG, border_color=col)
        
        # Circle badge
        badge = s15.shapes.add_shape(MSO_SHAPE.OVAL, c_left + Inches(0.95), Inches(2.1), Inches(0.9), Inches(0.9))
        badge.fill.solid()
        badge.fill.fore_color.rgb = col
        badge.line.fill.background()
        p = badge.text_frame.paragraphs[0]
        p.text = f"0{idx+1}"
        p.font.name = FONT_HEADING
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        tb = s15.shapes.add_textbox(c_left + Inches(0.2), Inches(3.2), Inches(2.4), Inches(3.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = rname
        p.font.name = FONT_HEADING
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = C_DARK_NAVY
        p.alignment = PP_ALIGN.CENTER
        
        p_sub = tf.add_paragraph()
        p_sub.text = rsub
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = col
        p_sub.alignment = PP_ALIGN.CENTER
        p_sub.space_before = Pt(4)

        p_desc = tf.add_paragraph()
        p_desc.text = rdesc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(12)
        p_desc.font.color.rgb = C_TEXT_MAIN
        p_desc.space_before = Pt(10)

    # ==================== SLIDE 16: HUMAN IN THE LOOP ====================
    s16 = prs.slides.add_slide(blank_layout)
    set_bg(s16)
    add_header(s16, "關鍵準則", "人機協同（Human-in-the-Loop）：教師與行政不可替代的智慧", "AI Agent 是強大的發電機與超級執行手，但教育的靈魂與方向盤始終在人類手中")

    principles = [
        ("⚖️ 價值判斷與倫理把關", "AI 不具備真實同理心與道德判斷。在處理學生霸凌事件、偏差行為懲處或家庭特殊狀況時，法律、輔導倫理與人情溫暖必須由教師親自決策。"),
        ("🔒 學生隱私與資料保護 (資安守則)", "切勿將學生真實姓名、身分證號、家庭敏感個資直接丟入未受企業保護的公有 AI 系統中。應使用代號（如：學生甲、學生乙）進行情境諮詢。"),
        ("👁️ 專業驗證與最終審查 (責任機制)", "Agent 生成的教學講義、考題解答與行政公文，教師與行政主管擁有「最終用印與公布權」。AI 是副駕駛，老師才是對教學成效負責的機長。")
    ]
    for idx, (ptitle, pdesc) in enumerate(principles):
        c_top = Inches(1.9 + idx * 1.7)
        add_card(s16, Inches(0.8), c_top, Inches(11.733), Inches(1.5), bg_color=C_CARD_BG, border_color=C_PRIMARY_BLUE)
        tb = s16.shapes.add_textbox(Inches(1.1), c_top + Inches(0.15), Inches(11.1), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = ptitle
        p.font.name = FONT_HEADING
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = C_PRIMARY_BLUE
        p2 = tf.add_paragraph()
        p2.text = pdesc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(12.5)
        p2.font.color.rgb = C_TEXT_MAIN
        p2.space_before = Pt(4)

    # ==================== SLIDE 17: ACTION GUIDE ====================
    s17 = prs.slides.add_slide(blank_layout)
    set_bg(s17)
    add_header(s17, "起步指南", "教學與行政人員的起手式：三步驟打造你的第一個 AI Agent", "不要等到完全學會才開始，今天就能從微型 SOP 啟動你的數位助理！")

    steps = [
        ("Step 1：盤點重複痛點", "找出你的「時間黑洞」", "檢視你每週最繁瑣、機械化且高重複性的工作：是催繳資料？是出隨堂練習卷？還是草擬行政週會宣導詞？挑選最想自動化的一個痛點。", C_AMBER, C_LIGHT_AMBER),
        ("Step 2：將專家智慧模組化", "撰寫你的第一個 Skill", "不要只下單一 Prompt。將你的專業寫成 SOP（角色、情境脈絡、檢查清單、好範例 Few-Shot、禁止事項），固化為可重複呼叫的技能指令手冊。", C_PRIMARY_BLUE, C_LIGHT_BLUE),
        ("Step 3：建立專案知識庫", "搭建你的專屬 Project", "在支援 Project 的 AI 工具中建立資料夾，丟入「南一數學課本PDF」、「學校行事曆」、「公文範本」。從此 AI 回覆時自動對齊校園情境！", C_EMERALD, C_LIGHT_GREEN)
    ]
    for idx, (stitle, ssub, sdesc, col, bg_c) in enumerate(steps):
        c_left = Inches(0.8 + idx * 3.9)
        add_card(s17, c_left, Inches(1.9), Inches(3.7), Inches(5.0), bg_color=C_CARD_BG, border_color=col)
        
        # Pill
        pill = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c_left + Inches(0.2), Inches(2.1), Inches(3.3), Inches(0.45))
        pill.fill.solid()
        pill.fill.fore_color.rgb = col
        pill.line.fill.background()
        p = pill.text_frame.paragraphs[0]
        p.text = stitle
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        tb = s17.shapes.add_textbox(c_left + Inches(0.2), Inches(2.7), Inches(3.3), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = ssub
        p.font.name = FONT_HEADING
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = C_DARK_NAVY
        p2 = tf.add_paragraph()
        p2.text = sdesc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(12)
        p2.font.color.rgb = C_TEXT_MAIN
        p2.space_before = Pt(8)

    # ==================== SLIDE 18: CONCLUSION & QA ====================
    s18 = prs.slides.add_slide(blank_layout)
    set_bg(s18, C_DARK_NAVY)

    tb_end = s18.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10.333), Inches(4.5))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True
    p = tf_end.paragraphs[0]
    p.text = "讓 AI Agent 承擔行政文書的繁瑣，"
    p.font.name = FONT_HEADING
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    p2 = tf_end.add_paragraph()
    p2.text = "把最珍貴的熱情與時間，留給教育現場的每一個孩子！"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(30)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(96, 165, 250)
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(10)

    p3 = tf_end.add_paragraph()
    p3.text = "感謝各位老師與行政夥伴的聆聽與投入\n歡迎開啟互動式網頁教材，體驗自適應出題與行政簽呈模擬器"
    p3.font.name = FONT_BODY
    p3.font.size = Pt(16)
    p3.font.color.rgb = RGBColor(203, 213, 225)
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(25)

    p4 = tf_end.add_paragraph()
    p4.text = "Q & A  交流與討論"
    p4.font.name = FONT_HEADING
    p4.font.size = Pt(24)
    p4.font.bold = True
    p4.font.color.rgb = RGBColor(251, 191, 36)
    p4.alignment = PP_ALIGN.CENTER
    p4.space_before = Pt(25)

    out_path = os.path.join("AI_Agent_研習教材與互動工具", "AI_Agent_基礎概念與教育行政應用.pptx")
    prs.save(out_path)
    print(f"Presentation saved successfully to {out_path} (18 slides, 16:9 widescreen)")

if __name__ == "__main__":
    create_deck()
