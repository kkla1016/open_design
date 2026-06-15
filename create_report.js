const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, BorderStyle, WidthType, AlignmentType, LevelFormat, ShadingType, VerticalAlign } = require('docx');

const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const cellBorders = { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder };

function createCell(textList, width, bold = false, isHeader = false) {
    return new TableCell({
        borders: cellBorders,
        width: { size: width, type: WidthType.DXA },
        shading: { fill: isHeader ? "D5E8F0" : "FFFFFF", type: ShadingType.CLEAR },
        verticalAlign: VerticalAlign.CENTER,
        children: [
            new Paragraph({
                alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                children: textList.map(text => new TextRun({ text: text, bold: bold }))
            })
        ]
    });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "微軟正黑體", size: 22 } } }, // 11pt default
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal",
        run: { size: 36, bold: true, color: "000000", font: "微軟正黑體" },
        paragraph: { spacing: { before: 240, after: 120 }, alignment: AlignmentType.CENTER } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: "000000", font: "微軟正黑體" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: "000000", font: "微軟正黑體" },
        paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: "000000", font: "微軟正黑體" },
        paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 2 } }
    ]
  },
  numbering: {
    config: [
      { reference: "bullet-list-1",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullet-list-2",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullet-list-3",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullet-list-4",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullet-list-5",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "num-list-1",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
    ]
  },
  sections: [{
    properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    children: [
        new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("【台股盤後強勢族群監控】2026/06/11")] }),
        
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("大盤簡述")] }),
        new Paragraph({ numbering: { reference: "bullet-list-1", level: 0 }, children: [
            new TextRun({ text: "今日加權指數收盤點位：", bold: true }), new TextRun("43,149.46 點")
        ]}),
        new Paragraph({ numbering: { reference: "bullet-list-1", level: 0 }, children: [
            new TextRun({ text: "漲跌幅：", bold: true }), new TextRun("下跌 76.08 點（-0.18%）")
        ]}),
        new Paragraph({ numbering: { reference: "bullet-list-1", level: 0 }, children: [
            new TextRun({ text: "成交量值：", bold: true }), new TextRun("約 1 兆 2,579 億元")
        ]}),
        new Paragraph({ numbering: { reference: "bullet-list-1", level: 0 }, children: [
            new TextRun({ text: "盤後解析：", bold: true }), new TextRun("今日台積電（2330）除息並於盤中完成填息。受國際地緣政治與美股科技股回檔影響，大盤全日震盪劇烈，盤中高低振幅達 1,456 點（下探 42,006 點後強力拉回），終場拉出逾千點長下影線，顯示低檔承接買盤強勁。資金明顯避開漲多且乖離大的權值股，轉向具備漲價題材、AI 規格升級與低基期優勢的「石英元件」、「FOPLP 先進封測」及「高階被動元件」族群。")
        ]}),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("強勢族群彙整表")] }),
        new Table({
            columnWidths: [1200, 1000, 800, 1800, 1800, 2760],
            margins: { top: 100, bottom: 100, left: 180, right: 180 },
            rows: [
                new TableRow({
                    tableHeader: true,
                    children: [
                        createCell(["產業別"], 1200, true, true),
                        createCell(["台股公司名稱"], 1000, true, true),
                        createCell(["台股代號"], 800, true, true),
                        createCell(["相對應/連動之美股代號"], 1800, true, true),
                        createCell(["今日盤後表現 (漲幅% / 成交量變化)"], 1800, true, true),
                        createCell(["上漲主因與資金動態簡述"], 2760, true, true)
                    ]
                }),
                new TableRow({ children: [
                    createCell(["石英元件"], 1200, true), createCell(["晶技"], 1000), createCell(["3042"], 800), createCell(["NVDA（AI 伺服器需求）"], 1800), createCell(["漲停 +10% / 量增突破"], 1800), createCell(["高階 OCXO 供不應求，受惠 AI 光收發模組成長，外資大買逾 3,300 張，資金集中族群發動。"], 2760)
                ]}),
                new TableRow({ children: [
                    createCell(["石英元件"], 1200, true), createCell(["安碁"], 1000), createCell(["6174"], 800), createCell(["TXN（車用與工控需求）"], 1800), createCell(["漲停 +10% / 成交量放大"], 1800), createCell(["產品漲價題材與低基期優勢，跟隨龍頭比價上攻，籌碼安定。"], 2760)
                ]}),
                new TableRow({ children: [
                    createCell(["石英元件"], 1200, true), createCell(["加高"], 1000), createCell(["8182"], 800), createCell(["TXN（車用與網通需求）"], 1800), createCell(["漲停 +10% / 成交量活絡"], 1800), createCell(["低基期補漲，技術面站回短均線，籌碼換手積極。"], 2760)
                ]}),
                new TableRow({ children: [
                    createCell(["封測與FOPLP"], 1200, true), createCell(["南茂"], 1000), createCell(["8150"], 800), createCell(["AMD / NVDA（先進封裝）"], 1800), createCell(["漲停 +10% (收103元) / 爆量"], 1800), createCell(["FOPLP 技術加速落地，投信單日大買 4,769 張鎖碼，AI 高階封裝需求升溫。"], 2760)
                ]}),
                new TableRow({ children: [
                    createCell(["面板級封裝"], 1200, true), createCell(["群創"], 1000), createCell(["3481"], 800), createCell(["AMAT（設備與封裝轉型）"], 1800), createCell(["強勢上漲 / 爆大量"], 1800), createCell(["跨足面板級封裝 (FOPLP) 具轉機題材，三大法人單日狂掃逾 3 萬張。"], 2760)
                ]}),
                new TableRow({ children: [
                    createCell(["被動元件"], 1200, true), createCell(["信昌電"], 1000), createCell(["6173"], 800), createCell(["TDK（AI 基建需求）"], 1800), createCell(["漲停 +10% (收251.5元) / 爆量2.6萬張"], 1800), createCell(["創歷史新高，AI 伺服器帶動高階產品出貨，盤中劇烈洗盤換手，投信連買支撐。"], 2760)
                ]}),
                new TableRow({ children: [
                    createCell(["光學鏡頭"], 1200, true), createCell(["聯一光"], 1000), createCell(["3441"], 800), createCell(["AAPL（消費電子與AI視覺）"], 1800), createCell(["漲停 +10% (收71.5元) / 爆量2.2萬張"], 1800), createCell(["產業鏈復甦與 AI 視覺應用帶動，外資與主力連續回補，技術結構轉強。"], 2760)
                ]})
            ]
        }),
        new Paragraph({ children: [new TextRun({ text: "(註：盤後資料受結算時間影響，實際籌碼數據以證交所最終公告為準)", italics: true, size: 20 })] }),

        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("【同族群比價與落後補漲清單】")] }),
        new Paragraph({ numbering: { reference: "bullet-list-2", level: 0 }, children: [
            new TextRun({ text: "希華（2484）", bold: true }), new TextRun(" — 石英元件同業。同具網通與車用題材，基期相對較低，具備跟隨晶技與安碁的比價效應。")
        ]}),
        new Paragraph({ numbering: { reference: "bullet-list-2", level: 0 }, children: [
            new TextRun({ text: "台嘉碩（3221）", bold: true }), new TextRun(" — 石英元件與表面聲波元件廠。受惠網通基礎建設復甦，籌碼安定，具落後補漲機會。")
        ]}),
        new Paragraph({ numbering: { reference: "bullet-list-2", level: 0 }, children: [
            new TextRun({ text: "力成（6239）", bold: true }), new TextRun(" — 記憶體與邏輯封測大廠。FOPLP 與高階封裝指標股，若南茂持續走強，力成具備強大基本面與落後補漲動能。")
        ]}),
        new Paragraph({ numbering: { reference: "bullet-list-2", level: 0 }, children: [
            new TextRun({ text: "國巨（2327）", bold: true }), new TextRun(" — 被動元件龍頭。AI 伺服器與車用高階 MLCC 需求大增，為信昌電比價向上之保護傘。")
        ]}),
        new Paragraph({ numbering: { reference: "bullet-list-2", level: 0 }, children: [
            new TextRun({ text: "華通（2313）", bold: true }), new TextRun(" — HDI 板大廠。受惠 SpaceX 即將 IPO 與低軌衛星題材，今日獲法人大買 1.1 萬張，同屬泛通訊與 AI 基建落後補漲題材。")
        ]}),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("核心股進場策略建議報告")] }),
        
        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("【晶技（3042）】操盤經理人評估報告")] }),
        new Paragraph({ children: [new TextRun({ text: "【強勢動能評估】", bold: true })] }),
        new Paragraph({ children: [new TextRun({ text: "上漲持續性：產業底部起漲與波段突破型（非一日行情）。", bold: true })] }),
        new Paragraph({ children: [new TextRun("晶技這波上漲非單純資金輪動，而是具備結構性需求。受惠 AI 伺服器光收發模組、網通及車用需求爆發，高階 OCXO（恆溫控制石英振盪器）呈現供不應求。今日帶量強勢突破整理區間，顯示波段資金已實質進駐，核心利多具備長線延續性。")] }),
        new Paragraph({ children: [new TextRun({ text: "【技術與籌碼建議】", bold: true })] }),
        new Paragraph({ numbering: { reference: "bullet-list-3", level: 0 }, children: [new TextRun({ text: "均線結構：", bold: true }), new TextRun("5、10、20、60 日均線重返多頭排列，技術面強勢。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-3", level: 0 }, children: [new TextRun({ text: "籌碼面：", bold: true }), new TextRun("外資大舉買超逾 3,300 張，投信同步加碼，主力籌碼集中度顯著提升。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-3", level: 0 }, children: [new TextRun({ text: "進場位階：", bold: true }), new TextRun("強勢表態後，建議於量縮回測 5 日線或突破前高頸線支撐區低接，避免盤中急拉追高。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-3", level: 0 }, children: [new TextRun({ text: "停損點：", bold: true }), new TextRun("嚴格設定於跌破 20 日線（月線）或近期爆量長紅 K 棒下緣。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-3", level: 0 }, children: [new TextRun({ text: "風險報酬比：", bold: true }), new TextRun("下檔支撐強韌，預估風險報酬比至少 1 : 2.5 以上。")]}),

        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("【南茂（8150）】操盤經理人評估報告")] }),
        new Paragraph({ children: [new TextRun({ text: "【強勢動能評估】", bold: true })] }),
        new Paragraph({ children: [new TextRun({ text: "上漲持續性：波段突破與題材發酵型。", bold: true })] }),
        new Paragraph({ children: [new TextRun("FOPLP（扇出型面板級封裝）被視為繼 CoWoS 之後的下世代封裝救星。南茂具備實質產能與技術積累，伴隨 AI 邊緣運算與高階封裝需求外溢，今日爆量強攻漲停（103元），確立了產業題材發酵與內資法人的高度認同，屬中長線波段起漲。")] }),
        new Paragraph({ children: [new TextRun({ text: "【技術與籌碼建議】", bold: true })] }),
        new Paragraph({ numbering: { reference: "bullet-list-4", level: 0 }, children: [new TextRun({ text: "均線結構：", bold: true }), new TextRun("帶量突破整理區間，短期均線呈多頭發散。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-4", level: 0 }, children: [new TextRun({ text: "籌碼面：", bold: true }), new TextRun("投信單日大舉買超 4,769 張，顯示內資法人強力看好其波段行情，主力籌碼流入內資大戶手中。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-4", level: 0 }, children: [new TextRun({ text: "進場位階：", bold: true }), new TextRun("強勢漲停隔日若遇大盤震盪開高走低，可待量縮回測 5 日均線或 100 元整數關卡支撐時分批佈局。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-4", level: 0 }, children: [new TextRun({ text: "停損點：", bold: true }), new TextRun("跌破 10 日均線或今日跳空缺口下緣。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-4", level: 0 }, children: [new TextRun({ text: "風險報酬比：", bold: true }), new TextRun("波段空間打開，預估風險報酬比達 1 : 3。")]}),

        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("【信昌電（6173）】操盤經理人評估報告")] }),
        new Paragraph({ children: [new TextRun({ text: "【強勢動能評估】", bold: true })] }),
        new Paragraph({ children: [new TextRun({ text: "上漲持續性：強勢波段延續型。", bold: true })] }),
        new Paragraph({ children: [new TextRun("AI 基礎建設帶動高階被動元件用量暴增。信昌電今日強勢攻上 251.5 元創下歷史新高。盤中經歷劇烈洗盤（爆量 2.6 萬張）後再度鎖死漲停，顯示換手成功，洗出浮額，屬於實質長線題材帶動的強勢飆股。")] }),
        new Paragraph({ children: [new TextRun({ text: "【技術與籌碼建議】", bold: true })] }),
        new Paragraph({ numbering: { reference: "bullet-list-5", level: 0 }, children: [new TextRun({ text: "均線結構：", bold: true }), new TextRun("均線極度強勢多頭排列，技術指標進入高檔強勢區。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-5", level: 0 }, children: [new TextRun({ text: "籌碼面：", bold: true }), new TextRun("投信連續 5 日買超鎖碼，盤中主力大單敲進換手，化解外資逢高獲利了結賣壓。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-5", level: 0 }, children: [new TextRun({ text: "進場位階：", bold: true }), new TextRun("因已創歷史新高且乖離率大，強烈不建議盲目追高。應耐心等待量縮回測 10 日線，或確認爆量低點不破時再行伺機介入。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-5", level: 0 }, children: [new TextRun({ text: "停損點：", bold: true }), new TextRun("跌破今日爆量換手區間低點或 20 日線。")]}),
        new Paragraph({ numbering: { reference: "bullet-list-5", level: 0 }, children: [new TextRun({ text: "風險報酬比：", bold: true }), new TextRun("高檔震盪風險增，預估風險報酬比約 1 : 1.5。")]}),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("法人與 ETF 籌碼異動追蹤")] }),
        
        new Paragraph({ children: [new TextRun({ text: "【晶技（3042）法人籌碼追蹤】", bold: true })] }),
        new Paragraph({ children: [
            new TextRun({ text: "異動狀況：", bold: true }), new TextRun("外資近 1 日大舉買進逾 3,300 張，投信亦站在買方。")
        ]}),
        new Paragraph({ children: [
            new TextRun({ text: "ETF 預期：", bold: true }), new TextRun("具備優良配息紀錄，穩居多檔台股高股息 ETF 及中型 100 成分股，被動資金持續提供活水。")
        ]}),
        new Paragraph({ children: [
            new TextRun({ text: "影響：", bold: true }), new TextRun("外資認錯回補與 ETF 被動資金提供中期穩固支撐，拉回時下檔買盤極具韌性。")
        ]}),

        new Paragraph({ children: [new TextRun({ text: "【南茂（8150）法人籌碼追蹤】", bold: true })] }),
        new Paragraph({ children: [
            new TextRun({ text: "異動狀況：", bold: true }), new TextRun("投信單日大買 4,769 張，為近期內資重點鎖碼標的。")
        ]}),
        new Paragraph({ children: [
            new TextRun({ text: "ETF 預期：", bold: true }), new TextRun("市場預期後續有望受惠於半導體主題或先進封裝 ETF 的持股權重調整。")
        ]}),
        new Paragraph({ children: [
            new TextRun({ text: "影響：", bold: true }), new TextRun("投信積極大額建倉顯示極度認同 FOPLP 長線趨勢，籌碼安定度大幅提升，有利抵禦大盤震盪並推升波段。")
        ]}),

        new Paragraph({ children: [new TextRun({ text: "【群創（3481）法人籌碼追蹤】", bold: true })] }),
        new Paragraph({ children: [
            new TextRun({ text: "異動狀況：", bold: true }), new TextRun("三大法人單日狂掃逾 3 萬張（斥資約 17.9 億元）。")
        ]}),
        new Paragraph({ children: [
            new TextRun({ text: "ETF 預期：", bold: true }), new TextRun("頻繁進出於大型權值與高息 ETF 成分股調整名單，轉型題材吸引法人買盤。")
        ]}),
        new Paragraph({ children: [
            new TextRun({ text: "影響：", bold: true }), new TextRun("法人資金大舉湧入，為其 FOPLP 轉型題材提供強大流動性與底部翻揚支撐。")
        ]}),

        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("剛起漲 vs 已經高檔 判斷")] }),
        new Table({
            columnWidths: [1800, 2500, 2500, 2560],
            margins: { top: 100, bottom: 100, left: 180, right: 180 },
            rows: [
                new TableRow({
                    tableHeader: true,
                    children: [
                        createCell(["判斷維度"], 1800, true, true),
                        createCell(["剛起漲（優質）"], 2500, true, true),
                        createCell(["已經高檔（風險）"], 2500, true, true),
                        createCell(["本族群判定 (石英/FOPLP)"], 2560, true, true)
                    ]
                }),
                new TableRow({ children: [
                    createCell(["時間週期"], 1800), createCell(["近 1 週類股漲幅前 5 大，但近 1 月未完全噴出"], 2500), createCell(["近 1 個月已漲 30% 以上且加速乖離"], 2500), createCell(["✅ 符合剛起漲"], 2560, true)
                ]}),
                new TableRow({ children: [
                    createCell(["量能變化"], 1800), createCell(["底部爆量、量增 150% 至 300%"], 2500), createCell(["量縮不創高或高檔爆量長上影"], 2500), createCell(["✅ 底部/頸線初爆量"], 2560, true)
                ]}),
                new TableRow({ children: [
                    createCell(["技術線型"], 1800), createCell(["多頭排列、20MA 翻揚、突破整理區"], 2500), createCell(["技術指標過熱、乖離過大"], 2500), createCell(["✅ 剛突破整理區"], 2560, true)
                ]}),
                new TableRow({ children: [
                    createCell(["籌碼面"], 1800), createCell(["大戶持股上升、散戶追價不明顯"], 2500), createCell(["散戶指標快速增加"], 2500), createCell(["✅ 主力/大戶持股升"], 2560, true)
                ]}),
                new TableRow({ children: [
                    createCell(["法人動作"], 1800), createCell(["外資連三買或投信布局"], 2500), createCell(["外資轉賣、投信不接"], 2500), createCell(["✅ 投信大買/外資回補"], 2560, true)
                ]})
            ]
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("最終結論：")] }),
        new Paragraph({ numbering: { reference: "num-list-1", level: 0 }, children: [
            new TextRun({ text: "石英元件（以晶技為首）與 FOPLP 封測族群（以南茂為首）：", bold: true }), new TextRun("經評估屬於"), new TextRun({ text: "「剛起漲的優質投資機會」", bold: true }), new TextRun("。這兩個族群皆在今日大盤重挫千點的壓力測試下，展現出帶量突破整理區間的極強韌性。伴隨法人（特別是投信）實質資金建倉與 AI 題材進一步擴散，籌碼與技術面皆具優勢，建議列為本週優先佈局的核心強勢族群，"), new TextRun({ text: "拉回 5 日線皆是絕佳買點", bold: true }), new TextRun("。")
        ]}),
        new Paragraph({ numbering: { reference: "num-list-1", level: 0 }, children: [
            new TextRun({ text: "被動元件（如信昌電）：", bold: true }), new TextRun("雖基本面強勁且創新高，但短期乖離過大，屬於「已經高檔 / 強勢延續」。建議針對此類標的改採"), new TextRun({ text: "「等拉回量縮」", bold: true }), new TextRun("策略，切勿於盤中急拉時追高，以嚴格控管資金風險。")
        ]})
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('台股盤後強勢族群監控_20260611.docx', buffer);
  console.log('Document created successfully');
});
