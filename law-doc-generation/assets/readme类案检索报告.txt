每次开始此对话框工作前，先阅读本README文件。

【双向钢人论证规则（前置思考流程）】
先别急着回答，也别默认我已经把问题想清楚。请先对这个问题做一次详细思考的"双向钢人论证"：
1. 用最完整、最有力的方式，重述我真正想解决的问题。
2. 使用钢人论证法分别给出支持我当前想法、以及反对它的最强论证。
3. 找出双方真正的分歧，以及最可能改变结论的关键变量。
4. 只问我一个最关键的问题。
等我回答后，再给出明确判断、理由和下一步行动。
说明：本规则与【执行前提问规则】分工不同——本规则要求你在回答前先对问题做深度钢人论证，【执行前提问规则】要求你在不理解或有歧义时先向用户确认。两者并行适用，不冲突。

【本文件用途与分工说明】
- 本README（readme类案检索报告）是"生成与迭代最终类案检索报告Word文档"的专门指引，适用于两类情形：
  1. 新生成：产出完整"类案检索报告"Word（含封面、检索说明、检索结论、案例摘要比对表、案例原文附件）。
  2. 迭代修改：在既有版本报告（如V9/V10）上增删案例、重排案例顺序、重编号、更新页码，其他内容保持不变。
- 与另外两个README的分工：
  - readme案例检索.txt → 仅单独检索案例（不生成报告Word）。
  - readme法律研究报告.txt → 生成"法律研究报告"（初步结论＋法规＋理论观点＋案例，参考"法律研究参考模板.docx"）。
  - 本readme类案检索报告.txt → 生成"类案检索报告"（使用骨架库 `C:\Users\jiajinhua\Desktop\实习VIBE\类案检索报告\骨架库\类案检索报告参考骨架\`，下称"参考骨架"，其为参考产出docx清理后的副本：settings.xml去rsid、document.xml清空body保留命名空间头。删除参考产出docx后仍可据此独立生成）。
- 三者同目录存放，均位于 `C:\Users\jiajinhua\Desktop\实习VIBE\类案检索报告\`。
- 本README定位：从参考产出中学习并固化其【内容结构】与【Word格式规范】（第二部分至第五部分），并固化【迭代操作规范】（第七部分）。生成或修改报告时直接依据本README执行，不必再打开参考产出逐一对照。

【历史版本规则（强制要求）】
- 本README每次更新，必须先保存历史版本，方便回退。
- 操作：更新前，将当前版本文件复制到历史版本目录，文件名追加时间戳（如 `readme类案检索报告_20260811_170000.txt`）。
- 历史版本目录：`C:\Users\jiajinhua\Desktop\实习VIBE\类案检索报告\历史版本\`
- 未保存历史版本前，不得覆盖更新任一文件。

【执行前提问规则（强制要求）】
- 用户告知任何问题、任务，或要求修改任何文件时，凡有不理解、不确定、有歧义之处，必须先向用户确认后再执行，不得自行揣测执行。
- 确认方式：复述我的理解（或列出疑问点），请用户确认无误后再动手。
- 涉及格式、结构、规则、方法等关键事项时，即便看似清楚，也建议先简要复述理解以核对一致，再执行。
- 若用户已明确指示或此前已确认过同一规则，则不必重复提问，直接执行。

【检索工具使用规则】
- 执行本项目法律检索任务时，优先使用北大法宝MCP进行检索。
- 法律法规、司法解释、权威著作等文件优先从北大法宝获取，确保内容准确、官方。
- 检索案例时，优先从北大法宝案例库检索；如北大法宝检索不到或需要补充，再使用其他检索渠道。
- 检索的优先级：北大法宝MCP → 其他可用的官方检索渠道。
- 案例检索的方法流程（多轮检索、要素式审核、关键词同义处理、当事人身份关键词等）与readme案例检索.txt中的【案例检索完整流程】一致，本README不再重复，生成报告前的案例检索环节按该流程执行。

════════════════════════════════════════════════
【一、参考产出与学习方式】
════════════════════════════════════════════════
- 参考产出文件：`C:\Users\jiajinhua\Documents\WXWork\1688857893795344\Cache\File\2026-08\关于夫妻公司股东一方为公司债务提供担保系夫妻共同债务的类案检索报告20260428.docx`
- 该文件是"类案检索报告"的成品范例。本README已将该文件的整体章节结构、每部分内容要求、每部分格式（字体、字号、加粗、居中、缩进、表格列宽、分节与页码等）全部固化于下文。生成同类报告时，直接按本README执行，无需再打开参考产出。
- 若需复核格式细节，可打开参考产出对照；学习规律的方法：纵列对照多个案例，归纳每列的固定写法（如"案例名称、审理法院及案号"列统一为"判决书名称 + 法院全称(案号) + 审结日期"三段；"法院观点及判决摘要"列摘录"本院认为"原文；"案例说明"列用一句话概括裁判规则；"页码"列写附件页码范围如"1-3"）。
- 参考产出技术参数（供生成时核对）：
  - 总段落数：约421段；表格8张（1张5列汇总表＋7张2列附件信息表）；分节3个。
  - 文档页面：A4（pgSz w=11906 h=16838）。
  - document.xml 编码：UTF-8（无BOM）。
- 参考产出文件也可在桌面获取：`C:\Users\jiajinhua\Desktop\关于夫妻公司股东一方为公司债务提供担保系夫妻共同债务的类案检索报告20260428.docx`
- 本README已完整记录所有格式参数（见下方【五C 格式参数速查】），生成时可仅依赖本README及骨架库，无需反复参照参考产出。
- **生成依赖（重要）**：生成或迭代时必须使用骨架库 `C:\Users\jiajinhua\Desktop\实习VIBE\类案检索报告\骨架库\类案检索报告参考骨架\`（settings去rsid、document.xml空body保留命名空间头）。即使删除参考产出docx，也可据此独立生成完整报告。

════════════════════════════════════════════════
【二、整体章节结构与每部分内容要求】
════════════════════════════════════════════════

本类案检索报告的完整结构（自上而下）依次为：

1. 封面页（居中排版，整体带外框线）
   - 第一行：致送法院（带外框的首行，如"致：天津自由贸易试验区人民法院"，Times New Roman加粗、字号28）
   - 空行若干
   - 案件名称（居中，如"信达一汽商业保理有限公司 与 哈密市金鑫德汽车销售有限公司、刘超、任佳丽 有追索权保理纠纷案"，Times New Roman、字号32）
   - 案号（居中，如"（2026）津0319民初4297号"）
   - 空行
   - 大标题"类案检索报告"（居中，Times New Roman、字号28）
   - 空行
   - "提交人：XXX"（居中，如"提交人：信达一汽商业保理有限公司"）
   - 日期（居中，如"二〇二六年四月二十九日"）
   - 封面所有段落带外框线（top=thinThickSmallGap、bottom=thickThinSmallGap），行距固定360
   - 封面为独立分节（第1节），末尾段落（日期段）的 pPr 内嵌 sectPr，含页眉页脚引用（header rId9、footer rId10）、页边距（top/bottom=1440、left/right=1800）、页码从1开始、docGrid linePitch=312。
   - 封面的完整24段结构与格式详见第三部分之"封面页（24段）"。

2. "一、检索情况的说明"（一级标题，pStyle=a，numPr numId=0）
   - 子标题"检索目的及法律问题"（pStyle=a0）
     - 正文段：pStyle=a0 且 numPr numId=0、outlineLvl=9、`<w:b w:val="0"/>`（不加粗），首行缩进（firstLineChars=200/firstLine=480），行距440 exact。内容：写明本案当事人、案号（未知则写"（案件名称及案号待定）"），引用《最高人民法院关于统一法律适用加强类案检索的指导意见（试行）》相关规定，说明检索与本案高度相似案件的目的，并提出待检索的法律问题。
     - 法律问题段：**无pStyle（默认样式）**，不加编号、不加项目符号；spacing beforeLines=100/afterLines=100 line=440 exact；首行缩进（firstLineChars=200/firstLine=482）；jc=both；**整句加粗（w:b）＋下划线（w:u single）**，字号24。写法：一句设问句归纳核心争议（如"债权损失情形下，债权人住所地能否认定为侵权结果发生地，从而由债权人住所地人民法院对本案行使管辖权？"）。
   - 子标题"检索时间及检索平台"（pStyle=a0）
     - "时间：XXXX年X月X日"段、"平台：……"段：**无pStyle（默认样式）**，不加编号；spacing beforeLines=100/afterLines=100 line=440 exact；首行缩进（firstLineChars=200/firstLine=482）；jc=both；**"时间：""平台："标签加粗**，后面内容不加粗，字号24。

3. "二、检索结论"（一级标题，pStyle=a）
   - 子标题"检索结论摘要"（pStyle=a0，带numPr ilvl=1 numId=2）
     - 正文段（pStyle=a1，numPr ilvl=1 numId=2，首行缩进，字号24，不加粗）：概括检索到的主要裁判观点，说明法院认定争议问题的考察重点与分层倾向。摘要应**围绕本方要支持的诉讼观点展开**（如支持"债权人住所地＝侵权结果发生地，债权人住所地法院有管辖权"），列明支持该观点的代表性案例（含入库参考案例、关键裁判原文援引），并简要提示不利案例的存在，供阅者权衡。

4. "案例原文摘要及比对说明"（子标题，pStyle=a0）
   - 下接一张5列表格（详见第四部分之表格一）。
   - 表格说明：按序号罗列检索到的类案，每案一行，列示案例名称/审理法院/案号、法院观点及判决摘要、案例说明、页码。
   - 检索结论正文结束后，紧接本子标题及表格一，**中间不要再加任何子标题（如"支持案例""不利案例"等小标题都不要加）**，表格结束后用一空段过渡，再进入结尾部分。

5. 结尾部分（致送）
   - "此致"（pStyle=KWBodytext）
   - 空段、致送法院段（如"天津自由贸易试验区人民法院"）
   - "提交人：XXX"（右对齐）；"代理人：　"；"XXXX年X月X日"（右对齐）
   - 结尾各段格式：pStyle=KWBodytext、行距400 exact、字号24，详见第三部分之"结尾部分（此致→日期）"。
   - 结尾部分为独立分节（第2节），日期段之后的空一级标题段（pStyle=a）内嵌 sectPr，含页眉页脚引用（header rId11/12/15、footer rId13/14/16）、页边距（top/bottom=1418、left/right=1304）、docGrid linePitch=326（页码延续第1节，不重新开始）。

6. "二、附件：相关判决书或裁定书原文"（一级标题，pStyle=a）
   - 下面依次粘贴每个案例的判决书/裁定书全文。每个案例的附件格式见第三部分之"附件案例部分"及第四部分之表格二。
   - 附件为独立分节（第3节），文档末尾 sectPr（footer rId17）界定该节，页边距（top/right/bottom/left=1418）、页码从1重新开始（pgNumType start=1）、docGrid linePitch=326。因此表格一"页码"列填写的"1-3""4-7"等，均为附件分节内从1起算的页码范围。

════════════════════════════════════════════════
【三、段落与文字格式规范（正文部分）】
════════════════════════════════════════════════

▍封面页（24段，自上而下）
- 每段均：居中（jc=center，首段jc=both）、带外框线（pBdr top=thinThickSmallGap sz24、bottom=thickThinSmallGap sz24，space=1/4/31/4）、行距固定360（line=360 lineRule=exact）。
- 段落序号、字号（sz）与内容：
  1. "致：XXX" —— sz28、加粗（w:b）。
  2-5. 空行 —— sz24。
  6. 原告/申请人名称 —— sz32。
  7. "与" —— sz32。
  8. 被告/被申请人名称 —— sz32。
  9. "XX案"（案由＋"案"）—— sz32。
  10-11. 空行 —— sz32。
  12. 案号（如"（2026）津03民终5187号"）—— sz24、加粗。
  13-15. 空行 —— sz24。
  16. "类案检索报告" —— sz24、加粗。
  17-21. 空行 —— sz28。
  22. "提交人：XXX" —— sz28。
  23. 空行 —— sz28。
  24. 日期（如"二〇二六年八月六日"）—— sz32；本段 pPr 内嵌第1节 sectPr（header rId9/footer rId10、pgMar 1440/1800、pgNumType start=1、docGrid linePitch=312）。
- 空行段的字号须保留（决定空行高度），不可省略或合并。
- 字体：Times New Roman；中文字体继承默认样式。
- 生成占位版时：致送法院、案件名称、案由、案号、提交人处留空（或写"（案件名称）""（案号）"等占位提示），日期可预填检索当日或留空，由用户后续填写。

▍一级标题（"一、检索情况的说明""二、检索结论""二、附件：相关判决书或裁定书原文"）
- pStyle="a"；`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>`（numId=0 在 numbering.xml 中无定义，属"无编号"标记，**不产生自动编号**；"一、""二、"等编号文字**直接写入正文文本**）。
- 行距440 exact（spacing before=240 after=240 line=440 lineRule=exact）。
- 缩进：`<w:ind w:left="851" w:hanging="567"/>`（个别段落含 `<w:tabs><w:tab w:val="clear" w:pos="851"/></w:tabs>`）。
- 字号约24（sz约24），Times New Roman。

▍二级标题/子标题（"检索目的及法律问题""检索时间及检索平台""检索结论摘要""案例原文摘要及比对说明"）
- pStyle="a0"，行距440 exact（spacing before=240 after=240 line=440 lineRule=exact），首行缩进 firstLine=482。
- "检索结论摘要"额外带 numPr（ilvl=1 numId=2）与 `<w:ind w:left="0" w:firstLine="426"/>`。
- 字号约24，Times New Roman。

▍正文段（检索说明正文、检索结论摘要正文）
- 检索说明正文：pStyle="a0" + numPr numId=0 + outlineLvl=9 + `<w:b w:val="0"/>`，首行缩进（firstLineChars=200/firstLine=480），行距440 exact，字号24，Times New Roman。注意：此段**不加粗、不编号**（numId=0无编号）。
- 检索结论摘要正文：pStyle="a1"，numPr **ilvl=0 numId=0（无可见编号）**，首行缩进（firstLineChars=200/firstLine=480），行距440 exact（spacing before=240 afterLines=100 after=240 line=440 lineRule=exact），outlineLvl=9，字号24，Times New Roman，`<w:b w:val="0"/>`（不加粗）。注意：此处**不得**使用 numId=2（numId=2 有定义会渲染"（一）"编号），必须用 numId=0 避免出现多余序号。
- 法律问题句：**无pStyle（默认样式）**，不加编号；spacing beforeLines=100/afterLines=100 line=440 exact；首行缩进（firstLineChars=200/firstLine=482）；jc=both；**加粗（w:b）＋下划线（w:u single）**，字号24。

▍结尾部分（此致→日期，第2节）
- "此致"段：pStyle=KWBodytext，spacing after=0 line=400 exact，缩进 `<w:ind w:firstLineChars="250" w:firstLine="600"/>`，字号24。
- 空段：pStyle=KWBodytext，line=400 exact。
- 致送法院段（如"天津市第三中级人民法院"）：pStyle=KWBodytext，line=400 exact。
- 提交人段："提交人：XXX"，pStyle=KWBodytext，jc=right，spacing afterLines=100 after=240 line=400 exact。
- 代理人段："代理人：　"，pStyle=KWBodytext，缩进 `<w:ind w:right="1680" w:firstLineChars="2300" w:firstLine="5520"/>`。
- 日期段："XXXX年X月X日"，pStyle=KWBodytext，jc=right，缩进 `<w:ind w:firstLineChars="1700" w:firstLine="4080"/>`。
- 最后跟一个空的一级标题段（pStyle=a、numId=0、空文本），其 pPr 内嵌第2节 sectPr（header rId11/12/15、footer rId13/14/16、pgMar top/bottom=1418 left/right=1304、docGrid linePitch=326）。

▍附件案例部分
- 每个案例前有标题段，格式"案例X：判决书/裁定书全称"（如"案例1：黄某、衷某合同、无因管理、不当得利纠纷再审审查与审判监督民事裁定书"）：
  - pStyle="biao"，楷体（eastAsia=楷体_GB2312）加粗（w:b+w:bCs），字号24，首行缩进（firstLineChars=200/firstLine=482），左对齐（jc=left），outlineLvl=1。
- 标题下先放"表格二"（审理法院/案号/裁判日期/案由信息表）。
- 随后是判决书/裁定书原文，逐段粘贴，格式：
  - 正文段：pStyle="p"，行距440 exact，首行缩进（firstLineChars=200/firstLine=480），字号24，Times New Roman，jc=both。
  - 当事人信息段（如"再审申请人（一审被告、二审上诉人）：黄某，女，1965年5月12日出生……"）、委托诉讼代理人段：同上正文段格式，每段一人，独立成段。
  - 判决书首部段（"再审申请人黄某因与被申请人……一案，不服……，向本院申请再审"）：正文格式。
  - "本院认为""本院经审查认为"等裁判理由段：正文格式，完整保留原文。
  - 裁判结果段（如"驳回黄某的再审申请。"）：正文格式。
  - 尾部审判人员段（"审判长：XXX""审判员：XXX"）、日期段（"二Ｏ二Ｏ年六月三十日"）、法官助理/书记员段：正文格式，独立成段。
- 一个案例附件结束后空行，再接下一案例。
- **每个案例标题都要位于页面顶部**：在案例标题段（pStyle=biao）之前设分页，即在上一案例末尾追加分页符段落 `<w:p><w:r><w:br w:type="page" /></w:r></w:p>`（最后一个案例不加分页符），使每个案例标题均从新的一页顶部开始。
- 附件原文获取：**按案号在北大法宝案例库检索，找到判决书/裁定书原文，将"本院认为"等裁判理由段按上述格式填入附件**，不要全部留占位。当事人信息、裁判结果等完整原文尽量一并填入；检索不到全文时，至少填入已获取的"本院认为"原文摘录。
- 若用户明确表示"稍后自行粘贴原文"，则生成框架＋摘要表＋原文占位提示段（如"【本附件请粘贴该案判决书/裁定书原文……】"），并明确告知用户；否则一律填入检索到的原文。

════════════════════════════════════════════════
【四、表格格式规范】
════════════════════════════════════════════════

▍表格一：案例原文摘要及比对表（核心表格，5列）
- 表头（第一行）五列内容：
  序号 | 案例名称、审理法院及案号 | 法院观点及判决摘要 | 案例说明 | 页码
- 每列列宽（tblGrid，单位dxa）：707 | 1617 | 6020 | 1713 | 853
- 各单元格宽度（tcW，type=pct，百分比）：表头行 324 | 741 | 2758 | 785 | 391；数据行同表头行。
- 表格属性：
  - `<w:tblStyle w:val="KWTable"/>`
  - `<w:tblW w:w="5950" w:type="pct"/>`（表格占页面宽百分比）
  - `<w:tblInd w:w="-717" w:type="dxa"/>`
  - `<w:tblLook w:val="04A0" w:firstRow="1" .../>`
- 表头单元格格式：`<w:tcPr><w:tcW w:w="324" w:type="pct"/><w:vAlign w:val="center"/></w:tcPr>`，单元格内段落居中（jc=center）、字号24、Times New Roman；表头行 trPr 含 `<w:cnfStyle w:val="100000000000" w:firstRow="1"/>`。
- 数据行单元格格式：`<w:tcPr><w:tcW w:w="324" w:type="pct"/><w:vAlign w:val="center"/></w:tcPr>`。
- 各数据列单元格内段落格式（spacing beforeLines=100 afterLines=100 line=480 exact，字号24，Times New Roman）：
  - 序号列：居中（jc=center），数字加粗（w:b）。
  - "案例名称、审理法院及案号"列：两端对齐（jc=both）。**单元格内5个段落**，自上而下依次为：
    ① 判决书/裁定书**文书标题全称**（即文书标题，如"黄某、衷某合同、无因管理、不当得利纠纷再审审查与审判监督民事裁定书"，不是"案例X"）；
    ② 空段；
    ③ **审理法院全称＋案号**（如"最高人民法院(2020)最高法民申2755号"，法院与案号之间不加括号包住案号）；
    ④ 空段；
    ⑤ **判决书最底部落款日期，用数字格式 xxxx.xx.xx**（如"2020.06.30"，不是中文"二〇二〇年六月三十日"）。
    即"文书标题 / 法院＋案号 / 落款日期"三段，各用一空段隔开（共5段）。
  - "法院观点及判决摘要"列：两端对齐。**首段为"XX法院认为："（法院名称加粗＋下划线）**，其后摘录判决书"本院认为"等裁判理由原文（须完整准确，可含司法解释条文原文）；**原文中支持本方观点的关键句须加粗＋下划线**（如"债权人住所地即侵权结果发生地""原告住所地即为侵权结果发生地"等）。
  - "案例说明"列：两端对齐。**自写**该案裁判规则的一句话，开头为"XX法院认为，"（**用逗号，不用冒号**；此处**不加粗、不下划线，纯文本**），内容须**支持本方诉讼观点**（如支持"债权人住所地＝侵权结果发生地"），并可附【不利】提示。写法示例："重庆法院认为，清算责任纠纷在性质上系侵权责任纠纷，债权损失造成的债权人财产权益减损发生并显现于债权人住所地，故债权人住所地即侵权结果发生地，债权人住所地法院对本案有管辖权。"
  - "页码"列：居中，**填写该案例对应附件全文（判决书/裁定书原文）所在的页码范围**（如"1-1""2-2""3-4"；页码为附件分节内从1起算的页）。**禁止估算**——生成后用Word COM打开文档，读取每个案例标题段所在页（附件分节第1页=案例1标题页），据此计算各案例页码范围并回填到表格一"页码"列；若打开后版面与计算时一致则页码不变，若因修改内容导致分页变化需重新读取。
- 数据行行数＝检索到的有效案例数，序号从1连续编号。
- XML结构要点：`<w:tbl>` 内按 `<w:tblPr>→<w:tblGrid>→<w:tr>(每行)→<w:tc>(每格)→<w:p>` 层级闭合，每个单元格必须含至少一个 `<w:p>`，文本需XML转义（& < > 等）。

▍表格二：附件案例信息表（每个案例附件各一张，2列，无边框）
- 列宽（tblGrid）：1555 | 7505
- 表格属性：`<w:tblStyle w:val="af0"/>`、`<w:tblW w:w="0" w:type="auto"/>`、六边无边框（tblBorders全部none）、`<w:tblLook .../>`
- 内容为四行两列（共8个单元格），每行是"标签＋值"：
  第1行：审理法院： | 最高人民法院
  第2行：案    号： | (2020)最高法民申2755号
  第3行：裁判日期： | 2020.06.30
  第4行：案    由： | 民事/婚姻家庭、继承纠纷/婚姻家庭纠纷
- 单元格段落：pStyle=p，行距400 exact（spacing before=0 after=0 line=400 lineRule=exact），字号24，Times New Roman，标签列左对齐（标签后跟冒号），值列写具体内容。
- 案由写法：按"民事/大分类/中分类/案由"的层级格式（如"民事/合同、准合同纠纷/合同纠纷/借款合同纠纷/金融借款合同纠纷"）。

════════════════════════════════════════════════
【五、分节与页码（sectPr）规范】
════════════════════════════════════════════════
- 全文档共3节，第1节为封面，第2节为正文（检索说明＋检索结论＋结尾），第3节为附件。
- 第1节（封面）sectPr：位于封面日期段 pPr 内；headerReference default=rId9、footerReference default=rId10；pgSz 11906×16838；pgMar top=1440 right=1800 bottom=1440 left=1800 header=851 footer=992 gutter=0；pgNumType start=1；cols space=425；docGrid type=lines linePitch=312。
- 第2节（正文）sectPr：位于结尾日期段之后的空一级标题段 pPr 内；headerReference even=rId11、default=rId12、first=rId15；footerReference even=rId13、default=rId14、first=rId16；pgMar top=1418 right=1304 bottom=1304 left=1418 header=284 footer=567 gutter=0；cols space=720；docGrid linePitch=326（不设pgNumType，页码延续第1节）。
- 第3节（附件）sectPr：文档末尾、`</w:body>`前；footerReference default=rId17；pgSz 11906×16838；pgMar top=1418 right=1418 bottom=1418 left=1418 header=284 footer=567 gutter=0；pgNumType start=1；cols space=720；docGrid linePitch=326。
- 生成时沿用参考产出既有的 sectPr 与页眉页脚关系（复制骨架库后替换 document.xml，关系ID不变，详见输出要求）。
- **迭代修改时，3个 sectPr 一律保持不动**，只增删改 `<w:body>` 内的段落与表格。

════════════════════════════════════════════════
【五B、docx 包文件结构（骨架清单）】
════════════════════════════════════════════════
- 标准生成方式（不再依赖参考产出docx）：从骨架库复制 `C:\Users\jiajinhua\Desktop\实习VIBE\类案检索报告\骨架库\类案检索报告参考骨架\` → 替换该目录内 word/document.xml 的 <w:body>（骨架库 document.xml 为 `<w:body></w:body>` 空体，保留 `<w:body>` 前头部命名空间与 `</w:body>` 后收尾，整体替换 body 内容）→ 重新压缩为docx。
- 骨架库说明：为参考产出docx清理后的副本（settings.xml 去除全部 rsid 垃圾；document.xml 清空 body 仅保留 30+ 命名空间头与 `<w:body></w:body>`）。删除参考产出docx后仍可据此独立生成，无需再打开模板。
- 以下为手动重建时的最小骨架清单（所有文件位于压缩包根目录，zip条目路径不带前导"/"）：
  必须文件列表：
  ① [Content_Types].xml          — 内容类型映射（每个xml部件的Override）
  ② _rels/.rels                  — 顶级关系（指向word/document.xml）
  ③ word/document.xml            — 主文档体（见下）
  ④ word/_rels/document.xml.rels — 文档级关系（定rId映射）
  ⑤ word/numbering.xml           — 编号定义（本报告numId0=无编号、numId2=带编号）
  ⑥ word/styles.xml              — 样式定义（含pStyle: a/a0/a1/biao/p/KWBodytext/KWTable/af0）
  ⑦ word/settings.xml            — 文档设置
  ⑧ word/fontTable.xml           — 字体表（Times New Roman/宋体/楷体_GB2312）
  ⑨ word/footnotes.xml           — 脚注（可为空骨架）
  ⑩ word/endnotes.xml            — 尾注（可为空骨架）
  ⑪ word/header1~4.xml           — 页眉（对应rId9~rId15）
  ⑫ word/footer1~5.xml           — 页脚（对应rId10~rId17）
  ⑬ word/theme/theme1.xml        — 主题
- document.xml 骨架结构：
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" （其余30+命名空间声明原样保留）>
    <w:body>
      ...（替换内容：封面24段含第1节sectPr → 正文 → 结尾含第2节sectPr → 附件 → 第3节sectPr）
    </w:body>
  </w:document>
- 关键：保留参考产出的全部命名空间（30+ xmlns前缀）与关系ID（rId9~rId17），仅替换 <w:body> 内段落与表格。
- 3 个 sectPr 的位置：第1节在封面日期段 pPr 内；第2节在结尾空一级标题段 pPr 内；第3节在 body 末尾 </w:body> 前。

【五C、格式参数速查（精确 OOXML 值，直接套用不再猜值）】
════════════════════════════════════════════════

以下数值从参考产出模板实测提取，生成时直接使用。

一、封面段（24段）通用 pPr
  所有封面段落：
  <w:jc w:val="center"/>（首段 w:val="both"）
  <w:spacing w:line="360" w:lineRule="exact"/>
  <w:pBdr><w:top w:val="thinThickSmallGap" w:sz="24" w:space="1"/><w:bottom w:val="thickThinSmallGap" w:sz="24" w:space="4"/></w:pBdr>

  各段的 sz / bold：
  1: sz=28, bold=true, jc=both  (致：xxxxx)
  2-5: sz=24 (空行)
  6-9: sz=32 (当事人名称、案由)
  10-11: sz=32 (空行)
  12: sz=24, bold=true (案号)
  13-15: sz=24 (空行)
  16: sz=24, bold=true (类案检索报告)
  17-21: sz=28 (空行)
  22: sz=28 (提交人：xxx)
  23: sz=28 (空行)
  24: sz=32 (日期) — 内嵌第1节 sectPr

二、正文段落格式（除封面、表格、结尾、附件外）
  一级标题（pStyle=a）：
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
    <w:spacing w:before="240" w:after="240" w:line="440" w:lineRule="exact"/>
    <w:ind w:left="851" w:hanging="567"/>

  二级标题（pStyle=a0）：
    <w:spacing w:before="240" w:after="240" w:line="440" w:lineRule="exact"/>
    <w:ind w:firstLine="482"/>
    检索结论摘要：额外 <w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr> + <w:ind w:left="0" w:firstLine="426"/>

  检索目的正文段（pStyle=a0 + numId=0 + outlineLvl=9 + bold=0）：
    <w:pStyle w:val="a0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
    <w:outlineLvl w:val="9"/>
    <w:ind w:firstLineChars="200" w:firstLine="480"/>
    <w:spacing w:line="440" w:lineRule="exact"/>
    <w:rPr><w:b w:val="0"/></w:rPr>

  法律问题句（无 pStyle，加粗+下划线）：
    <w:ind w:firstLineChars="200" w:firstLine="482"/>
    <w:spacing w:beforeLines="100" w:afterLines="100" w:line="440" w:lineRule="exact"/>
    <w:jc w:val="both"/>
    run: <w:rPr><w:rFonts w:hint="eastAsia"/><w:b/><w:u w:val="single"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>

  "时间：/平台："段（无 pStyle，标签 bold）：
    <w:ind w:firstLineChars="200" w:firstLine="482"/>
    <w:spacing w:beforeLines="100" w:afterLines="100" w:line="440" w:lineRule="exact"/>
    <w:jc w:val="both"/>
    run标签: <w:rPr><w:rFonts w:hint="eastAsia"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
    run内容: 不加粗，其余同上

  检索结论摘要正文（pStyle=a1 + numId=0 + outlineLvl=9 + bold=0）：
    <w:pStyle w:val="a1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
    <w:ind w:firstLineChars="200" w:firstLine="480"/>
    <w:spacing w:before="240" w:afterLines="100" w:after="240" w:line="440" w:lineRule="exact"/>
    <w:outlineLvl w:val="9"/>
    <w:rPr><w:b w:val="0"/></w:rPr>

三、表格一（5列，表头+数据行）
  tblPr: <w:tblStyle w:val="KWTable"/><w:tblW w:w="5950" w:type="pct"/><w:tblInd w:w="-717" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1"/>
  tblGrid: <w:gridCol w:w="707"/>|<w:gridCol w:w="1617"/>|<w:gridCol w:w="6020"/>|<w:gridCol w:w="1713"/>|<w:gridCol w:w="853"/>

  表头行：<w:trPr><w:cnfStyle w:val="100000000000" w:firstRow="1"/></w:trPr>
  表头单元格：<w:tcPr><w:tcW w:w="324" w:type="pct"/><w:vAlign w:val="center"/></w:tcPr>
             单元格内段落 jc=center，间距 beforeLines=100 afterLines=100 line=480 exact，字号24

  数据行单元格：同表头单元格 tcPr
  数据行案例名称列（5段）：文书标题 / <w:p/> / 法院+案号 / <w:p/> / 落款日期xxxx.xx.xx

四、结尾段（KWBodytext）
  此致段: <w:spacing w:after="0" w:line="400" w:lineRule="exact"/><w:ind w:firstLineChars="250" w:firstLine="600"/>
  法院段: <w:spacing w:line="400" w:lineRule="exact"/>
  提交人: <w:jc w:val="right"/><w:spacing w:afterLines="100" w:after="240" w:line="400" w:lineRule="exact"/>
  代理人: <w:ind w:right="1680" w:firstLineChars="2300" w:firstLine="5520"/>
  日期段: <w:jc w:val="right"/><w:ind w:firstLineChars="1700" w:firstLine="4080"/>

五、附件案例块
  biao标题: <w:pStyle w:val="biao"/><w:ind w:firstLineChars="200" w:firstLine="482"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/>
           run: <w:rPr><w:rFonts w:eastAsia="楷体_GB2312"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  infotable（表格二，无边框）:
    <w:tblPr><w:tblStyle w:val="af0"/><w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders><全部none/></w:tblBorders></w:tblPr>
    列宽: <w:gridCol w:w="1555"/> | <w:gridCol w:w="7505"/>
    单元格: pStyle=p, 间距 before=0 after=0 line=400 exact, 字号24
  p正文段: <w:pPr><w:pStyle w:val="p"/><w:ind w:firstLineChars="200" w:firstLine="480"/><w:spacing w:line="440" w:lineRule="exact"/></w:pPr>
           run: <w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>

六、分节 sectPr（3节，精确值）
  第1节（封面末段 pPr 内）:
    <w:sectPr><w:headerReference w:type="default" r:id="rId9"/>
    <w:footerReference w:type="default" r:id="rId10"/>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
    <w:pgNumType w:start="1"/>
    <w:cols w:space="425"/>
    <w:docGrid w:linePitch="312"/></w:sectPr>

  第2节（结尾后空一级标题段 pPr 内）:
    <w:sectPr><w:headerReference w:type="even" r:id="rId11"/>
    <w:headerReference w:type="default" r:id="rId12"/>
    <w:headerReference w:type="first" r:id="rId15"/>
    <w:footerReference w:type="even" r:id="rId13"/>
    <w:footerReference w:type="default" r:id="rId14"/>
    <w:footerReference w:type="first" r:id="rId16"/>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1418" w:right="1304" w:bottom="1304" w:left="1418" w:header="284" w:footer="567" w:gutter="0"/>
    <w:cols w:space="720"/>
    <w:docGrid w:linePitch="326"/></w:sectPr>
    （不设 pgNumType，页码延续第1节）

  第3节（body 末尾）:
    <w:sectPr><w:footerReference w:type="default" r:id="rId17"/>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="284" w:footer="567" w:gutter="0"/>
    <w:pgNumType w:start="1"/>
    <w:cols w:space="720"/>
    <w:docGrid w:linePitch="326"/></w:sectPr>

七、字体对照（参考产出默认）
  英文：Times New Roman
  中文宋体：w:rFonts w:ascii="宋体" w:hint="eastAsia"（标题、正文默认）
  中文楷体：w:rFonts w:ascii="楷体_GB2312" w:hint="eastAsia"（biao标题）
  run 通用：<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>


【六、案例检索要点（本报告适用）】
════════════════════════════════════════════════
- 检索对象：与本案争议焦点（如"夫妻一方为公司债务提供担保是否属夫妻共同债务""债权损失情形下债权人住所地能否认定为侵权结果发生地"）高度相似的类案。
- 检索平台：北大法宝案例库优先（按案号检索原文），中国裁判文书网、威科先行（或按用户指定）补充。
- 案例筛选：优先选取最高人民法院及各地高级人民法院、中级人民法院的生效判决/裁定；同一法院可多案并列；兼顾支持与不利观点。
- 案例正文获取：**按案号在北大法宝案例库检索，找到判决书/裁定书原文**（含"本院认为"、当事人、裁判结果、审判人员等完整内容），用于摘要表"法院观点及判决摘要"列和附件粘贴。检索不到全文时，至少填入已获取的"本院认为"原文摘录。
- 摘要素材：法院观点及判决摘要须忠实引用判决书原文（首段"XX法院认为："加粗下划线，原文支持本方观点的关键句加粗下划线）；案例说明须自行提炼该案裁判规则的一句话，内容支持本方诉讼观点。
- 落款日期：案例名称列的落款日期取判决书/裁定书最底部落款时间，**用数字格式 xxxx.xx.xx**（如"2020.06.30"，不用中文大写日期）。
- **排序规则（已确认，强制）**：指导性/入库/参考案例（如最高人民法院参考案例库入库案例）置于最前；其后按法院层级从高到低排列：最高人民法院 → 高级人民法院 → 中级人民法院 → 基层人民法院；同级法院之间保持既有相对顺序（通常按审结日期）。表格一与附件必须按同一顺序排列，序号一致。
- **迭代重排原则**：用户要求调整案例顺序时，仅允许调整案例顺序与随之而来的序号、页码；各案例的正文原文、法院观点摘要、案例说明等**内容一律保持不变**。
- 不利案例处理：不利案例仍需纳入摘要表与附件（与用户确认），标题或名称中标注【不利】，便于用户自行审阅取舍。

════════════════════════════════════════════════
【七、迭代操作规范（在既有报告上增删/重排案例）】
════════════════════════════════════════════════
适用：用户要求在既有版本报告（如V9/V10）基础上"加入新案例""删除案例""重排案例顺序""重编号""更新页码"，其他内容保持不变。此时不重新生成全文，而是在上一版本 docx 上做局部增删改。

7.1 版本与工作目录
- 每次改动生成新版本号（如V10、V11），交付文件名带版本号（如"……类案检索报告V11.docx"）。
- 在临时目录建立工作目录：复制上一版本工作目录为 vNuser（如 v10user→v11user），只改其中的 `word\document.xml`；改好后再打包为新的docx交付。
- 迭代时源文件为上一版本报告docx（非骨架库），可能被Word/WPS占用无法读取，先把源文件复制一份到临时目录（如 src_gs.docx）再解压/解析。
- 保存到临时目录的文本/脚本文件属中间产物，不得混入工作目录（见7.6打包清理）。

7.2 先识别结构再动手（强制）
- 改动前，先用脚本dump当前 `<w:body>` 结构（各子元素索引、pStyle、是否含分页符、文本前若干字），确认案例块边界与表格一位置后，再据此编辑。切勿凭记忆猜索引。
- 每个案例块的标准结构（自上而下）：
  biao标题段（"案例N：……民事裁定书"，pStyle=biao）→ 空段 → 表格二（4行2列信息表）→ 空段 → 正文段若干（pStyle=p）→ 空段 → 分页符段（`<w:br w:type="page"/>`）。
- 案例块边界判定：从某个biao标题段起，到下一biao标题段（不含）为止；最后一个案例块到第3节sectPr（含）之前。
- 案例标题段有两种XML结构，重编号时必须区分处理：
  ① 单w:t：整句"案例N：标题"在一个 `<w:t>` 内 → 用正则替换 `^案例(\d+)?：`；
  ② 多w:t："案例"、"数字"、"：标题"分属多个 `<w:t>` → 改第2个 `<w:t>` 的 InnerText。
- 表格一（索引以dump为准，如 body 第36个子元素）行结构：表头行＋数据行，每行5列（序号/案例名称/法院观点/案例说明/页码），行顺序必须与附件案例顺序一致。

7.3 增删/重排案例的操作
- 新增案例：从源总结/检索文件提取该案正文 → 按既有案例块模板（biao标题段＋表格二＋正文段）构造XML → 插入目标位置 → 在相邻案例块之间补齐分页符段。
- 删除案例：移除整个案例块（biao→分页符段）；检查相邻案例间的分页符是否因此冗余或缺失。
- 重排案例：先收集各案例块节点（biao起、分页符段止），从body移除，再按新顺序重新插入；表格一同步重排到相同顺序。
- 增删/重排后统一执行：重编号案例标题（1..N）→ 重排/重编表格一序号 → 重算并回填表格一页码。

7.4 分页符维护
- 案例块之间必须各有分页符段，保证每个案例标题从新页顶部开始；分页符段总数 = 案例数-1。
- 新增案例插入点前后、删除案例后、重排后，都要核对分页符位置与数量，否则Word会把两个案例排到同一页。

7.5 打包docx（压缩）
- docx是zip：遍历工作目录所有文件，按相对路径（正斜杠、**不带前导"/"**）写入zip条目，[Content_Types].xml、_rels、word/ 等原样保留。
- **打包前必须清理工作目录中的杂项文件**（如 body_structure.txt、tbl1_content.txt、*.ps1、*.xml 中间产物等），只保留docx解压出的原始部件；否则Word打开报"文件可能已经损坏"。
- 临时包先用Word COM验证可正常打开、页码正确，再交付到目标目录（`C:\Users\jiajinhua\Desktop\实习VIBE\类案检索报告\`）。

7.6 页码重算（增删/重排后强制重算，禁止沿用旧页码）
- 新增/删除/重排案例都会改变分页，旧页码一律失效。
- 方法：把无页码（或旧页码）版本打包 → Word COM打开 → 读每个"案例N："标题段所在物理页 → 附件偏移=案例1标题物理页-1 → 各案范围=本案标题页~下一案标题页-1，末案到文档末页（ComputeStatistics(2)）→ 转成附件页码（物理页-偏移）回填表格一"页码"列 → 重新打包验证。
- 若两次打包分页一致则页码不变；改动影响分页的内容后必须重读。

7.7 PowerShell/XML 编码与语法注意（实战教训）
- 写 .ps1 脚本时，文件须以 **UTF-8（带BOM）** 保存后再执行，否则中文乱码或语法解析失败；执行前可用 `[System.IO.File]::WriteAllText($p, $text, [System.Text.UTF8Encoding]::new($true))` 转换后再运行。
- 字符串中变量后紧跟冒号要用 `${}` 括起（如 `"seq${i}"`，不能写 `"seq$i:"`，否则 `$i:` 被解析为驱动器语法）。
- 中文输出前设置 `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8`。
- 解析/保存 document.xml 用 `XmlDocument`＋`XmlNamespaceManager`（w前缀），保存时用 UTF8Encoding($false)（无BOM）。
- 写入的文本必须XML转义（& < > 等）。
- 修改后必须用Word COM验证能正常打开、案例起始页与总页数符合预期。

════════════════════════════════════════════════
【输出要求】
════════════════════════════════════════════════
- 输出的Word内容格式必须与参考产出"关于夫妻公司股东一方为公司债务提供担保系夫妻共同债务的类案检索报告20260428.docx"完全一致（包括封面、章节、表格、附件格式）。
- 生成Word时：从骨架库复制 `C:\Users\jiajinhua\Desktop\实习VIBE\类案检索报告\骨架库\类案检索报告参考骨架\` 为工作目录（其 document.xml 已清空 body、保留 30+ 命名空间头；settings.xml 已去rsid），替换其 `word/document.xml` 的 `<w:body>...</w:body>` 整体内容（正文按本README【三、段落与文字格式规范】【四、表格格式规范】【五C、格式参数速查】构造，含封面24段＋第1节sectPr＋第2节sectPr＋第3节sectPr），重新压缩为docx。**不再依赖参考产出docx，删除后仍可独立生成。**
- 技术要点：
  - document.xml 头部（`<w:document ...>` 及全部xmlns命名空间声明）原样保留，`<w:body>` 及其内容替换。
  - 封面块（24段，含第1节sectPr）与末尾第3节sectPr原样保留；如需保留第2节分节，在结尾日期段之后补空一级标题段（pStyle=a、空文本）并内嵌第2节sectPr（rId11-16）。
  - 文档编码：document.xml 以 UTF-8（无BOM）写入。
  - 重压缩docx时保持zip条目相对路径正确（不得带前导"/"），[Content_Types].xml、word/_rels/*.rels 等原样保留。
  - **迭代修改（第七部分）时：不替换body，只增删改body内段落与表格；3个sectPr与页眉页脚关系保持不变。**
- 文件名：按用户要求命名（如"关于XX的类案检索报告YYYYMMDD.docx"）；迭代版本在文件名末尾加版本号（如V9/V10/V11）。
- **页码计算方法（Word COM精确读取）**：
  1. 先按不含页码的版本生成docx，用Word COM打开（`New-Object -ComObject Word.Application; $doc=$w.Documents.Open(路径,$false,$true)`）。
  2. 遍历`$doc.Paragraphs`，对文本匹配`^案例\d+：`的段落，用`$p.Range.Information(3)`读取其所在物理页码。
  3. 附件分节（第3节）页码从1起算：案例1标题所在物理页即附件第1页，偏移量=案例1物理页码-1。
  4. 每个案例的页码范围 = 从本案例标题页 到（下一案例标题页-1）；最后一个案例到文档末页（`$doc.ComputeStatistics(2)`）。
  5. 将各案例页码范围回填表格一"页码"列后重新生成；若版面未变则页码不变。
  - 注意：必须在无其他改动时一次性读取，读取后不要改动影响分页的内容，否则需重读。

【生成后校验】
- 生成后用Word打开验证：封面外框线、表格显示、附件原文完整、页码对应关系正确。
- 若Word无法打开（常见于表格XML结构错误），检查：`<w:tbl>...<w:tblPr>...<w:tblGrid>...<w:tr>...<w:tc>...<w:p>...` 层级闭合是否完整、每个单元格必须含至少一个 `<w:p>`、文本需XML转义（& < > 等）。
- 可用脚本校验：XmlDocument.Load 成功加载、段落数、表格数（表格一1张＋表格二n张）、sectPr数（3）、n个案例标题齐全、zip条目完整（无杂项文件）。
- 格式内容核对要点（对照参考产出）：
  - 表格一之前的正文：一级标题（pStyle=a＋numId=0＋ind left=851 hanging=567）不应出现多余编号或项目符号；"检索目的及法律问题"正文段（a0＋numId=0＋outlineLvl=9）不加粗；法律问题句加粗＋下划线；"时间：/平台："段标签加粗、默认样式；"案例原文摘要及比对说明"与表格一之间**不得**插入"支持案例""不利案例"等任何小标题。
  - 检索结论部分："检索结论摘要"标题（a0＋numId=2）可渲染"（一）"；**其下摘要正文段必须用 numId=0（无可见编号）＋outlineLvl=9＋afterLines=100**，不得用 numId=2，否则会出现多余"（二）"序号。
  - 表格一名称列：每个数据行名称列须为5段（文书标题/空/法院＋案号/空/落款日期）；落款日期须为**数字格式 xxxx.xx.xx**（如2020.06.30），不得用中文大写日期。
  - 法院观点列：首段"XX法院认为："加粗下划线，原文关键句加粗下划线。
  - 案例说明列：纯文本（**不加粗、不下划线**），去除任何加粗/下划线/【B】标记；开头"XX法院认为，"**用逗号不用冒号**。
  - 页码列：须填写每个案例对应附件原文的页码范围（**用Word COM读取真实分页后计算，禁止按字数估算**），连续编号，不得留空。
  - 附件：应填入按案号检索到的判决书/裁定书原文（"本院认为"等），不得仅留占位提示（除非用户明确表示稍后自行粘贴）。
  - 分页：每个案例标题前应有分页（案例间分页符段落数＝案例数-1，`<w:br w:type="page"/>`），确保每个案例标题位于页面顶部。
- **迭代修改后的额外校验**：
  - Word COM确认可打开、无"文件损坏"提示。
  - 案例标题齐全且连续编号1..N，顺序与表格一一致。
  - 表格一每行页码与Word实测分页一致；3个sectPr数量与位置未变；总页数符合预期。
