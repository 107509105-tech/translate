# PDF Translation Tool - 中文轉英文

使用PyMuPDF自動將中文PDF翻譯成英文，同時保持原有格式。

## 📋 功能特點

- ✅ 使用PyMuPDF提取和處理PDF
- ✅ 自動翻譯中文到英文（使用Google Translate）
- ✅ 保持原有格式（字體大小、顏色、位置）
- ✅ 翻譯覆蓋率：74%+

## 🚀 使用方法

### 1. 安裝依賴

```bash
python3 -m venv venv
source venv/bin/activate
pip install pymupdf googletrans==4.0.0rc1
```

### 2. 執行翻譯

```bash
./venv/bin/python translate_pdf_auto.py
```

### 3. 查看結果

翻譯後的PDF將保存為：`113_chartN_english.pdf`

## 📁 文件說明

- `translate_pdf_auto.py` - **推薦使用** 自動翻譯腳本（使用Google Translate）
- `translate_pdf_complete.py` - 基於字典的翻譯腳本
- `translate_pdf_v2.py` - 第二版翻譯腳本
- `translate_pdf.py` - 第一版翻譯腳本

## 🎯 翻譯結果

當前翻譯統計：
- 總文本片段：123個
- 已翻譯片段：91個
- 翻譯覆蓋率：**74.0%**

## ⚙️ 自定義設置

### 修改字體大小

在`translate_pdf_auto.py`中，找到以下代碼並調整倍率：

```python
# 根據文本長度動態調整字體大小
if text_length > 100:
    fontsize = repl["fontsize"] * 0.5  # 調整此處
elif text_length > 50:
    fontsize = repl["fontsize"] * 0.65  # 調整此處
else:
    fontsize = repl["fontsize"] * 0.75  # 調整此處
```

### 修改翻譯語言

```python
# 在translate_text函數中
result = translator.translate(text_key, src='zh-tw', dest='en')
# 可以改為其他語言代碼，如：src='zh-cn'（簡體中文）
```

## 📝 已知限制

1. **流程圖框內的文字**：左側流程圖框中的中文可能是圖像的一部分，無法自動翻譯
2. **複雜格式**：某些複雜的表格或圖形可能需要手動調整
3. **翻譯質量**：使用Google Translate，翻譯質量可能不如人工翻譯

## 🔧 進階選項

### 處理流程圖框內的文字

如果需要翻譯流程圖框內的文字，可以：

1. 使用PDF編輯器（如Adobe Acrobat）手動編輯
2. 或使用OCR技術先識別圖像中的文字

### 提高翻譯質量

可以修改`TRANSLATIONS`字典來提供更準確的翻譯：

```python
TRANSLATIONS = {
    "碩士班研究生修業流程圖": "Master's Program Graduate Study Flowchart",
    # 添加更多自定義翻譯...
}
```

## 💡 使用建議

1. **首次運行**：建議先在測試PDF上運行，檢查效果
2. **備份原文件**：翻譯前請備份原始PDF
3. **檢查結果**：翻譯完成後請仔細檢查翻譯質量和格式
4. **手動調整**：對於關鍵內容，建議人工審核和修正

## 🐛 常見問題

### Q: 為什麼有些中文沒有被翻譯？
A: 可能原因：
- 文字是圖像的一部分
- 文字在特殊圖層中
- 文字使用了特殊編碼

### Q: 翻譯後格式跑掉了怎麼辦？
A: 可以調整字體大小倍率，或使用更專業的PDF編輯工具進一步調整

### Q: 可以翻譯成其他語言嗎？
A: 可以，修改`translate_text`函數中的`dest`參數即可

## 📞 技術支持

如有問題，請檢查：
1. Python版本（需要3.7+）
2. 依賴是否正確安裝
3. 網絡連接（Google Translate需要網絡）

## 📄 License

此工具僅供學習和研究使用。
