using System;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace AITranslateCore
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class AITranslateCore : IRibbonExtensibility
    {
        protected dynamic app;
        protected string hostType;
        protected IRibbonUI ribbon;

        // Lưu trữ ngôn ngữ đã chọn trực tiếp trên Ribbon
        protected static string selectedSource = "Auto Detect";
        protected static string selectedTarget = "Vietnamese";

        protected string serverUrl = "https://localhost:3000";
        protected string clientToken = "secure-token-123";

        // Stacks lưu lịch sử dịch để Undo/Redo cho Excel
        protected static System.Collections.Generic.Stack<ExcelUndoState> excelUndoStack = new System.Collections.Generic.Stack<ExcelUndoState>();
        protected static System.Collections.Generic.Stack<ExcelUndoState> excelRedoStack = new System.Collections.Generic.Stack<ExcelUndoState>();

        public AITranslateCore(object Application, string host)
        {
            app = Application;
            hostType = host;
            LoadSettings();
        }

        // Tự động tải cài đặt Token và URL cục bộ
        private void LoadSettings()
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string settingsPath = Path.Combine(appData, "AITranslateAddin", "settings.txt");
                if (File.Exists(settingsPath))
                {
                    string[] lines = File.ReadAllLines(settingsPath);
                    foreach (string line in lines)
                    {
                        if (line.StartsWith("API_URL=")) serverUrl = line.Substring(8).Trim();
                        if (line.StartsWith("TOKEN=")) clientToken = line.Substring(6).Trim();
                    }
                }
            }
            catch { }
        }

        // Lưu đối tượng Ribbon khi load
        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        // Trả về XML cấu trúc của Custom Ribbon Menu với 2 Dropdown chọn ngôn ngữ trực tiếp
        public string GetCustomUI(string RibbonID)
        {
            string activeLabel = "Translate Page";
            if (hostType == "Excel") activeLabel = "Translate Sheet";
            else if (hostType == "PPT") activeLabel = "Translate Slide";

            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
                <ribbon>
                    <tabs>
                        <tab id='TabTranslate' label='Translate'>
                            <group id='GroupTranslate' label='AI Translate'>
                                <dropDown id='dropSource' label='From' getSelectedItemIndex='GetSelectedSourceIndex' onAction='OnSourceSelected'>
                                    <item id='srcAuto' label='Auto Detect' />
                                    <item id='srcVi' label='Vietnamese' />
                                    <item id='srcEn' label='English' />
                                    <item id='srcZh' label='Chinese' />
                                    <item id='srcJa' label='Japanese' />
                                    <item id='srcKo' label='Korean' />
                                    <item id='srcFr' label='French' />
                                    <item id='srcDe' label='German' />
                                    <item id='srcEs' label='Spanish' />
                                    <item id='srcRu' label='Russian' />
                                </dropDown>
                                <dropDown id='dropTarget' label='To' getSelectedItemIndex='GetSelectedTargetIndex' onAction='OnTargetSelected' getItemCount='GetTargetItemCount' getItemLabel='GetTargetItemLabel' getItemID='GetTargetItemID'>
                                </dropDown>
                                <button id='btnSelection' label='Translate Selection' size='large' onAction='OnAction' tag='selection' imageMso='Translate' />
                                <button id='btnActive' label='" + activeLabel + @"' size='large' onAction='OnAction' tag='active' imageMso='PageSetupPageDialog' />
                                <button id='btnAll' label='Translate All' size='large' onAction='OnAction' tag='all' imageMso='WebPagePreview' />
                                <button id='btnUndo' label='Undo' size='large' onAction='OnUndoAction' imageMso='Undo' getEnabled='GetUndoEnabled' getVisible='GetUndoVisible' />
                                <button id='btnRedo' label='Redo' size='large' onAction='OnRedoAction' imageMso='Redo' getEnabled='GetRedoEnabled' getVisible='GetRedoVisible' />
                            </group>
                        </tab>
                    </tabs>
                </ribbon>
            </customUI>";
        }

        // Đặt index mặc định khi load Ribbon
        public int GetSelectedSourceIndex(IRibbonControl control)
        {
            return 0; // "Auto Detect"
        }

        // Lọc danh sách ngôn ngữ đích (không bao gồm ngôn ngữ nguồn đã chọn nếu khác Auto Detect)
        private string[] GetFilteredTargetLangs()
        {
            string[] allTargets = new string[] {
                "Vietnamese", "English", "Chinese", "Japanese", "Korean", "French", "German", "Spanish", "Russian"
            };
            if (string.IsNullOrEmpty(selectedSource) || selectedSource == "Auto Detect")
            {
                return allTargets;
            }
            var list = new System.Collections.Generic.List<string>();
            foreach (var lang in allTargets)
            {
                if (lang != selectedSource)
                {
                    list.Add(lang);
                }
            }
            return list.ToArray();
        }

        public int GetTargetItemCount(IRibbonControl control)
        {
            return GetFilteredTargetLangs().Length;
        }

        public string GetTargetItemLabel(IRibbonControl control, int index)
        {
            var langs = GetFilteredTargetLangs();
            if (index >= 0 && index < langs.Length)
            {
                return langs[index];
            }
            return "";
        }

        public string GetTargetItemID(IRibbonControl control, int index)
        {
            return "tgt_" + index;
        }

        public int GetSelectedTargetIndex(IRibbonControl control)
        {
            var langs = GetFilteredTargetLangs();
            for (int i = 0; i < langs.Length; i++)
            {
                if (langs[i] == selectedTarget)
                {
                    return i;
                }
            }
            if (langs.Length > 0)
            {
                selectedTarget = langs[0];
                return 0;
            }
            return -1;
        }

        // Sự kiện khi chọn ngôn ngữ nguồn trên Ribbon
        public void OnSourceSelected(IRibbonControl control, string selectedId, int selectedIndex)
        {
            string[] sourceLangs = new string[] {
                "Auto Detect", "Vietnamese", "English", "Chinese", "Japanese", "Korean", "French", "German", "Spanish", "Russian"
            };
            if (selectedIndex >= 0 && selectedIndex < sourceLangs.Length)
            {
                selectedSource = sourceLangs[selectedIndex];
            }

            // Cập nhật lại dropdown đích để loại trừ ngôn ngữ nguồn vừa chọn
            if (ribbon != null)
            {
                ribbon.InvalidateControl("dropTarget");
            }
        }

        // Sự kiện khi chọn ngôn ngữ đích trên Ribbon
        public void OnTargetSelected(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var langs = GetFilteredTargetLangs();
            if (selectedIndex >= 0 && selectedIndex < langs.Length)
            {
                selectedTarget = langs[selectedIndex];
            }
        }

        // Callbacks cho nút Undo trong Excel
        public bool GetUndoEnabled(IRibbonControl control)
        {
            return excelUndoStack.Count > 0;
        }

        public bool GetUndoVisible(IRibbonControl control)
        {
            return hostType == "Excel";
        }

        public void OnUndoAction(IRibbonControl control)
        {
            try
            {
                if (excelUndoStack.Count > 0)
                {
                    var state = excelUndoStack.Pop();
                    dynamic targetRange = GetRangeFromState(state);

                    if (targetRange != null)
                    {
                        // Lưu giá trị công thức hiện tại (đã dịch) cho Redo trước khi khôi phục
                        excelRedoStack.Push(new ExcelUndoState { 
                            WorkbookName = state.WorkbookName,
                            WorksheetName = state.WorksheetName,
                            Address = state.Address,
                            OriginalValues = targetRange.Formula 
                        });

                        targetRange.Formula = state.OriginalValues;
                    }

                    if (ribbon != null)
                    {
                        ribbon.InvalidateControl("btnUndo");
                        ribbon.InvalidateControl("btnRedo");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Undo failed:\n" + ex.Message, "AI Translate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Callbacks cho nút Redo trong Excel
        public bool GetRedoEnabled(IRibbonControl control)
        {
            return excelRedoStack.Count > 0;
        }

        public bool GetRedoVisible(IRibbonControl control)
        {
            return hostType == "Excel";
        }

        public void OnRedoAction(IRibbonControl control)
        {
            try
            {
                if (excelRedoStack.Count > 0)
                {
                    var state = excelRedoStack.Pop();
                    dynamic targetRange = GetRangeFromState(state);

                    if (targetRange != null)
                    {
                        // Lưu giá trị công thức hiện tại (chưa dịch) cho Undo trước khi thực hiện Redo
                        excelUndoStack.Push(new ExcelUndoState { 
                            WorkbookName = state.WorkbookName,
                            WorksheetName = state.WorksheetName,
                            Address = state.Address,
                            OriginalValues = targetRange.Formula 
                        });

                        targetRange.Formula = state.OriginalValues;
                    }

                    if (ribbon != null)
                    {
                        ribbon.InvalidateControl("btnUndo");
                        ribbon.InvalidateControl("btnRedo");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Redo failed:\n" + ex.Message, "AI Translate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Hàm phân giải Range từ tên Workbook, Worksheet và Address để tránh mất tham chiếu COM proxy
        protected dynamic GetRangeFromState(ExcelUndoState state)
        {
            try
            {
                dynamic workbook = app.Workbooks[state.WorkbookName];
                dynamic worksheet = workbook.Worksheets[state.WorksheetName];
                return worksheet.Range[state.Address];
            }
            catch
            {
                return null;
            }
        }

        // Sự kiện click nút bấm trên Ribbon (Dịch thuật)
        public void OnAction(object control)
        {
            dynamic ctrl = control;
            string actionType = ctrl.Tag;

            // Load cấu hình mới nhất trước khi chạy
            LoadSettings();

            // Hiển thị trạng thái dịch trên Status Bar của Office
            SetStatusBar("AI Translate: Translating " + actionType + " contents to [" + selectedTarget + "]...");

            try
            {
                if (actionType == "selection")
                {
                    TranslateSelection();
                }
                else if (actionType == "active")
                {
                    TranslateActive();
                }
                else if (actionType == "all")
                {
                    TranslateAll();
                }

                SetStatusBar(false);
                MessageBox.Show("Translation completed successfully!", "AI Translate", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                SetStatusBar(false);
                MessageBox.Show("Translation failed:\n" + ex.Message, "AI Translate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetStatusBar(object status)
        {
            try
            {
                app.StatusBar = status;
            }
            catch { }
        }

        // Xử lý dịch Selection trực tiếp
        private void TranslateSelection()
        {
            if (hostType == "Word")
            {
                dynamic selection = app.Selection;
                dynamic range = selection.Range;
                string text = range.Text;
                if (string.IsNullOrEmpty(text) || text == "\r")
                    throw new Exception("Please select some text in Word first.");

                TranslateWordRange(range, selectedTarget, selectedSource);
            }
            else if (hostType == "Excel")
            {
                dynamic range = app.Selection;
                TranslateExcelRange(range, selectedTarget, selectedSource);
            }
            else if (hostType == "PPT")
            {
                dynamic selection = app.ActiveWindow.Selection;
                if (selection.Type == 2 || selection.Type == 3) // Text hoặc Shape
                {
                    dynamic textRange = selection.TextRange;
                    string text = textRange.Text;
                    if (string.IsNullOrEmpty(text))
                        throw new Exception("Please select text in PowerPoint first.");

                    string translated = AsyncTranslateText(text, selectedTarget, selectedSource);
                    textRange.Text = translated;
                }
                else
                {
                    throw new Exception("Please select text inside a shape on the slide first.");
                }
            }
        }

        // Xử lý dịch Active Sheet / Section / Slide
        private void TranslateActive()
        {
            if (hostType == "Word")
            {
                // Sử dụng Bookmark ẩn "\Page" để lấy phạm vi (Range) của trang hiện tại
                dynamic range = app.ActiveDocument.Bookmarks.Item("\\Page").Range;
                TranslateWordRange(range, selectedTarget, selectedSource);
            }
            else if (hostType == "Excel")
            {
                dynamic sheet = app.ActiveSheet;
                dynamic usedRange = sheet.UsedRange;
                TranslateExcelRange(usedRange, selectedTarget, selectedSource);
            }
            else if (hostType == "PPT")
            {
                dynamic slide = app.ActiveWindow.View.Slide;
                TranslatePPTSlide(slide, selectedTarget, selectedSource);
            }
        }

        // Xử lý dịch toàn bộ tài liệu
        private void TranslateAll()
        {
            if (hostType == "Word")
            {
                dynamic doc = app.ActiveDocument;
                dynamic range = doc.Content;
                TranslateWordRange(range, selectedTarget, selectedSource);
            }
            else if (hostType == "Excel")
            {
                dynamic workbook = app.ActiveWorkbook;
                foreach (dynamic sheet in workbook.Worksheets)
                {
                    dynamic usedRange = sheet.UsedRange;
                    TranslateExcelRange(usedRange, selectedTarget, selectedSource);
                }
            }
            else if (hostType == "PPT")
            {
                dynamic pres = app.ActivePresentation;
                foreach (dynamic slide in pres.Slides)
                {
                    TranslatePPTSlide(slide, selectedTarget, selectedSource);
                }
            }
        }

        // Dịch và ghi đè nội dung ô Excel (bao gồm cả dịch text bên trong các công thức/function)
        private void TranslateExcelRange(dynamic range, string targetLang, string sourceLang)
        {
            dynamic values = range.Value2;
            dynamic formulas = range.Formula;
            if (values == null) return;

            string wbkName = range.Worksheet.Parent.Name;
            string wshName = range.Worksheet.Name;
            string addr = range.Address;

            if (values is string)
            {
                string val = (string)values;
                string formulaVal = formulas as string;

                if (formulaVal != null && formulaVal.StartsWith("="))
                {
                    System.Collections.Generic.List<string> literals;
                    System.Collections.Generic.List<int[]> ranges;
                    ExtractFormulaLiterals(formulaVal, out literals, out ranges);

                    if (literals.Count > 0)
                    {
                        // Lưu trạng thái trước khi dịch
                        excelUndoStack.Push(new ExcelUndoState { 
                            WorkbookName = wbkName,
                            WorksheetName = wshName,
                            Address = addr,
                            OriginalValues = formulas 
                        });
                        excelRedoStack.Clear();
                        if (ribbon != null)
                        {
                            ribbon.InvalidateControl("btnUndo");
                            ribbon.InvalidateControl("btnRedo");
                        }

                        string[] translatedLiterals = HttpTranslateArray(literals.ToArray(), targetLang, sourceLang);
                        string newFormula = RebuildFormula(formulaVal, ranges, translatedLiterals);
                        range.Formula = newFormula;
                    }
                }
                else if (!string.IsNullOrEmpty(val))
                {
                    // Lưu trạng thái trước khi dịch
                    excelUndoStack.Push(new ExcelUndoState { 
                        WorkbookName = wbkName,
                        WorksheetName = wshName,
                        Address = addr,
                        OriginalValues = formulas 
                    });
                    excelRedoStack.Clear();
                    if (ribbon != null)
                    {
                        ribbon.InvalidateControl("btnUndo");
                        ribbon.InvalidateControl("btnRedo");
                    }

                    range.Formula = HttpTranslateText(val, targetLang, sourceLang);
                }
                TranslateExcelShapesAndComments(range.Worksheet, targetLang, sourceLang);
                return;
            }

            object[,] valArray = (object[,])values;
            object[,] formulaArray = (object[,])formulas;
            int rows = valArray.GetLength(0);
            int cols = valArray.GetLength(1);

            System.Collections.Generic.List<string> textList = new System.Collections.Generic.List<string>();
            System.Collections.Generic.List<int[]> staticCoords = new System.Collections.Generic.List<int[]>(); // { r, c, textListIndex }
            System.Collections.Generic.List<FormulaTranslationInfo> formulaInfos = new System.Collections.Generic.List<FormulaTranslationInfo>();

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    object cellVal = valArray[r, c];
                    object cellFormula = formulaArray[r, c];

                    if (cellFormula != null && cellFormula is string && ((string)cellFormula).StartsWith("="))
                    {
                        string formulaStr = (string)cellFormula;
                        System.Collections.Generic.List<string> literals;
                        System.Collections.Generic.List<int[]> ranges;
                        ExtractFormulaLiterals(formulaStr, out literals, out ranges);

                        if (literals.Count > 0)
                        {
                            FormulaTranslationInfo info = new FormulaTranslationInfo
                            {
                                Row = r,
                                Col = c,
                                OriginalFormula = formulaStr,
                                Literals = literals,
                                Ranges = ranges,
                                TextListStartIndex = textList.Count
                            };
                            textList.AddRange(literals);
                            formulaInfos.Add(info);
                        }
                    }
                    else if (cellVal != null && cellVal is string)
                    {
                        string str = (string)cellVal;
                        if (!string.IsNullOrEmpty(str))
                        {
                            textList.Add(str);
                            staticCoords.Add(new int[] { r, c, textList.Count - 1 });
                        }
                    }
                }
            }

            if (textList.Count == 0) return;

            string[] translated = new string[textList.Count];
            Exception error = null;

            ProgressForm progress = new ProgressForm("AI Translate: Translating cells & formulas...");

            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                try
                {
                    int batchSize = 500;
                    for (int i = 0; i < textList.Count; i += batchSize)
                    {
                        int count = Math.Min(batchSize, textList.Count - i);
                        string[] batchTexts = new string[count];
                        textList.CopyTo(i, batchTexts, 0, count);

                        progress.UpdateMessage("Translating " + (i + 1) + " to " + (i + count) + " of " + textList.Count + "...");
                        string[] batchResult = HttpTranslateArray(batchTexts, targetLang, sourceLang);
                        Array.Copy(batchResult, 0, translated, i, count);
                    }
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                finally
                {
                    progress.Invoke(new Action(() => progress.Close()));
                }
            });

            thread.Start();
            progress.ShowDialog();

            if (error != null) throw error;

            // Ghi nhận các ô static đã dịch
            for (int i = 0; i < staticCoords.Count; i++)
            {
                int[] coord = staticCoords[i];
                int r = coord[0];
                int c = coord[1];
                int index = coord[2];
                formulaArray[r, c] = translated[index];
            }

            // Ghi nhận các công thức đã dịch
            foreach (var info in formulaInfos)
            {
                string[] subTranslated = new string[info.Ranges.Count];
                Array.Copy(translated, info.TextListStartIndex, subTranslated, 0, info.Ranges.Count);
                string newFormula = RebuildFormula(info.OriginalFormula, info.Ranges, subTranslated);
                formulaArray[info.Row, info.Col] = newFormula;
            }

            // Lưu trạng thái trước khi thay đổi giá trị của mảng ô
            excelUndoStack.Push(new ExcelUndoState { 
                WorkbookName = wbkName,
                WorksheetName = wshName,
                Address = addr,
                OriginalValues = formulas 
            });
            excelRedoStack.Clear();
            if (ribbon != null)
            {
                ribbon.InvalidateControl("btnUndo");
                ribbon.InvalidateControl("btnRedo");
            }

            range.Formula = formulaArray;
            TranslateExcelShapesAndComments(range.Worksheet, targetLang, sourceLang);
        }

        // Hàm trích xuất các chuỗi literal trong công thức Excel
        private void ExtractFormulaLiterals(string formula, out System.Collections.Generic.List<string> literals, out System.Collections.Generic.List<int[]> ranges)
        {
            literals = new System.Collections.Generic.List<string>();
            ranges = new System.Collections.Generic.List<int[]>();

            if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
            {
                return;
            }

            int len = formula.Length;
            bool inString = false;
            System.Text.StringBuilder currentString = new System.Text.StringBuilder();
            int stringStart = -1;

            for (int i = 0; i < len; i++)
            {
                char c = formula[i];
                if (inString)
                {
                    if (c == '"')
                    {
                        // Kiểm tra nếu là dấu nháy kép được escape ("")
                        if (i + 1 < len && formula[i + 1] == '"')
                        {
                            currentString.Append('"');
                            i++; // Bỏ qua dấu nháy tiếp theo
                        }
                        else
                        {
                            // Kết thúc chuỗi literal
                            inString = false;
                            literals.Add(currentString.ToString());
                            ranges.Add(new int[] { stringStart, i - stringStart + 1 });
                            currentString.Clear();
                        }
                    }
                    else
                    {
                        currentString.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inString = true;
                        stringStart = i;
                    }
                }
            }
        }

        // Hàm tái dựng lại công thức với các chuỗi literal đã được dịch
        private string RebuildFormula(string originalFormula, System.Collections.Generic.List<int[]> ranges, string[] translatedValues)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder(originalFormula);
            for (int i = ranges.Count - 1; i >= 0; i--)
            {
                int[] range = ranges[i];
                int start = range[0];
                int length = range[1];
                string transVal = translatedValues[i];

                // Escape các dấu nháy kép trong chuỗi đã dịch
                string escapedTrans = transVal.Replace("\"", "\"\"");
                string replacement = "\"" + escapedTrans + "\"";

                sb.Remove(start, length);
                sb.Insert(start, replacement);
            }
            return sb.ToString();
        }

        // Dịch slide PowerPoint
        private void TranslatePPTSlide(dynamic slide, string targetLang, string sourceLang)
        {
            System.Collections.Generic.List<dynamic> textRanges = new System.Collections.Generic.List<dynamic>();
            System.Collections.Generic.List<string> texts = new System.Collections.Generic.List<string>();

            // Duyệt đệ quy tất cả các Shapes (bao gồm Group, SmartArt, Text Box...)
            foreach (dynamic shape in slide.Shapes)
            {
                ExtractPPTShapeTexts(shape, textRanges, texts);
            }

            // Dịch bình luận (Comments) của Slide
            try
            {
                foreach (dynamic comment in slide.Comments)
                {
                    string text = comment.Text;
                    if (!string.IsNullOrEmpty(text))
                    {
                        comment.Text = AsyncTranslateText(text, targetLang, sourceLang);
                    }
                }
            }
            catch { }

            if (texts.Count == 0) return;

            string[] translated = new string[texts.Count];
            Exception error = null;

            ProgressForm progress = new ProgressForm("AI Translate: Translating slide shapes...");

            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                try
                {
                    int batchSize = 100;
                    for (int i = 0; i < texts.Count; i += batchSize)
                    {
                        int count = Math.Min(batchSize, texts.Count - i);
                        string[] batchTexts = new string[count];
                        texts.CopyTo(i, batchTexts, 0, count);

                        progress.UpdateMessage("Translating shape " + (i + 1) + " to " + (i + count) + " of " + texts.Count + "...");
                        string[] batchResult = HttpTranslateArray(batchTexts, targetLang, sourceLang);
                        Array.Copy(batchResult, 0, translated, i, count);
                    }
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                finally
                {
                    progress.Invoke(new Action(() => progress.Close()));
                }
            });

            thread.Start();
            progress.ShowDialog();

            if (error != null) throw error;

            for (int i = 0; i < textRanges.Count; i++)
            {
                textRanges[i].Text = translated[i];
            }
        }

        // Dịch vùng tài liệu Word giữ nguyên định dạng qua trung gian HTML
        private void TranslateWordRange(dynamic range, string targetLang, string sourceLang)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "AITranslateWord");
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);

            string tempSourceHtml = Path.Combine(tempDir, Guid.NewGuid().ToString() + "_source.html");
            string tempTargetHtml = Path.Combine(tempDir, Guid.NewGuid().ToString() + "_target.html");

            System.Collections.Generic.List<ShapeBackup> shapesBackup = new System.Collections.Generic.List<ShapeBackup>();
            System.Collections.Generic.List<CommentBackup> commentsBackup = new System.Collections.Generic.List<CommentBackup>();
            dynamic doc = app.ActiveDocument;
            dynamic backupDoc = null;

            try
            {
                // Lưu lại vị trí biên ban đầu của Range để khôi phục chính xác sau này
                int originalStart = range.Start;
                int originalEnd = range.End;

                // 1. Sao chép nội dung range sang một tài liệu tạm ẩn TRƯỚC (khi Selection gốc chưa bị thay đổi)
                range.Copy();
                dynamic tempDoc = app.Documents.Add(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false);
                tempDoc.Content.Paste();

                // 2. Sao lưu và xóa comments nằm trong phạm vi dịch ở tài liệu chính
                try
                {
                    System.Collections.Generic.List<dynamic> commentsToDelete = new System.Collections.Generic.List<dynamic>();
                    foreach (dynamic comment in doc.Comments)
                    {
                        try
                        {
                            int commentStart = comment.Range.Start;
                            if (commentStart >= originalStart && commentStart <= originalEnd)
                            {
                                commentsBackup.Add(new CommentBackup
                                {
                                    ParagraphIndex = GetParagraphIndex(comment.Range),
                                    Text = comment.Range.Text,
                                    Author = comment.Author
                                });
                                commentsToDelete.Add(comment);
                            }
                        }
                        catch { }
                    }
                    foreach (dynamic comment in commentsToDelete)
                    {
                        try { comment.Delete(); } catch { }
                    }
                }
                catch { }

                // 3. Sao lưu và xóa shapes nằm trong phạm vi dịch ở tài liệu chính
                try
                {
                    System.Collections.Generic.List<dynamic> shapesToDelete = new System.Collections.Generic.List<dynamic>();
                    foreach (dynamic shape in doc.Shapes)
                    {
                        try
                        {
                            int anchorStart = shape.Anchor.Start;
                            if (anchorStart >= originalStart && anchorStart <= originalEnd)
                            {
                                if (backupDoc == null)
                                {
                                    backupDoc = app.Documents.Add(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false);
                                }

                                int paraIndex = GetParagraphIndex(shape.Anchor);
                                shape.Select();
                                app.Selection.Copy();
                                
                                backupDoc.Content.InsertParagraphAfter();
                                dynamic pasteRange = backupDoc.Paragraphs[backupDoc.Paragraphs.Count].Range;
                                pasteRange.Select();
                                app.Selection.Paste();

                                if (backupDoc.Shapes.Count > 0)
                                {
                                    shapesBackup.Add(new ShapeBackup
                                    {
                                        ParagraphIndex = paraIndex,
                                        Name = shape.Name,
                                        Type = (int)shape.Type,
                                        Left = (float)shape.Left,
                                        Top = (float)shape.Top,
                                        Width = (float)shape.Width,
                                        Height = (float)shape.Height,
                                        BackupShapeIndex = backupDoc.Shapes.Count
                                    });

                                    shapesToDelete.Add(shape);
                                }
                            }
                        }
                        catch { }
                    }
                    foreach (dynamic shape in shapesToDelete)
                    {
                        try { shape.Delete(); } catch { }
                    }
                }
                catch { }

                // 4. Xóa các shapes và comments trong tempDoc để không bị xuất ra HTML
                while (tempDoc.Shapes.Count > 0)
                {
                    try { tempDoc.Shapes[1].Delete(); } catch { break; }
                }
                while (tempDoc.InlineShapes.Count > 0)
                {
                    try { tempDoc.InlineShapes[1].Delete(); } catch { break; }
                }
                while (tempDoc.Comments.Count > 0)
                {
                    try { tempDoc.Comments[1].Delete(); } catch { break; }
                }

                // Cấu hình mã hóa UTF-8
                tempDoc.WebOptions.Encoding = 65001;

                // 5. Lưu tài liệu tạm dưới dạng Filtered HTML
                tempDoc.SaveAs2(tempSourceHtml, 10);
                tempDoc.Close(false);

                // 6. Đọc nội dung HTML từ file nguồn dạng UTF-8
                string htmlContent = File.ReadAllText(tempSourceHtml, Encoding.UTF8);

                // 7. Gửi lên API dịch HTML
                string translatedHtml = null;
                Exception error = null;

                ProgressForm progress = new ProgressForm("AI Translate: Translating page content...");

                System.Threading.Thread thread = new System.Threading.Thread(() =>
                {
                    try
                    {
                        translatedHtml = HttpTranslateHtml(htmlContent, targetLang, sourceLang);
                    }
                    catch (Exception ex)
                    {
                        error = ex;
                    }
                    finally
                    {
                        progress.Invoke(new Action(() => progress.Close()));
                    }
                });

                thread.Start();
                progress.ShowDialog();

                if (error != null) throw error;

                // 8. Ghi nội dung HTML đã dịch ra file đích dưới dạng UTF-8
                File.WriteAllText(tempTargetHtml, translatedHtml, Encoding.UTF8);

                // 9. Nhập file HTML đã dịch vào Range nguồn của tài liệu chính (thay thế nội dung cũ)
                // Tái tạo lại Range dựa trên vị trí tuyệt đối ban đầu để tránh ảnh hưởng của việc thay đổi Selection
                int currentEnd = doc.Content.End;
                int safeEnd = Math.Min(originalEnd, currentEnd);
                int safeStart = Math.Min(originalStart, safeEnd);
                dynamic targetRange = doc.Range(safeStart, safeEnd);
                targetRange.InsertFile(tempTargetHtml);

                // 10. Khôi phục các shapes gốc từ backupDoc
                if (backupDoc != null && shapesBackup.Count > 0)
                {
                    foreach (var backup in shapesBackup)
                    {
                        try
                        {
                            dynamic bShape = backupDoc.Shapes[backup.BackupShapeIndex];
                            bShape.Select();
                            app.Selection.Copy();

                            int targetParaIdx = Math.Min(backup.ParagraphIndex, doc.Paragraphs.Count);
                            if (targetParaIdx < 1) targetParaIdx = 1;
                            dynamic restoreTargetRange = doc.Paragraphs[targetParaIdx].Range;
                            restoreTargetRange.Select();
                            app.Selection.Paste();

                            dynamic restoredShape = doc.Shapes[doc.Shapes.Count];
                            restoredShape.Name = backup.Name;
                            restoredShape.Left = backup.Left;
                            restoredShape.Top = backup.Top;
                            restoredShape.Width = backup.Width;
                            restoredShape.Height = backup.Height;
                        }
                        catch { }
                    }
                }

                // 11. Khôi phục các comments gốc từ commentsBackup
                if (commentsBackup.Count > 0)
                {
                    foreach (var backup in commentsBackup)
                    {
                        try
                        {
                            int targetParaIdx = Math.Min(backup.ParagraphIndex, doc.Paragraphs.Count);
                            if (targetParaIdx < 1) targetParaIdx = 1;
                            dynamic restoreTargetRange = doc.Paragraphs[targetParaIdx].Range;
                            doc.Comments.Add(restoreTargetRange, backup.Text);
                        }
                        catch { }
                    }
                }

                // 12. Tiến hành dịch nội dung chữ trong các Shapes, SmartArts, TextBoxes và Comments
                TranslateWordShapesAndComments(doc, targetLang, sourceLang);
            }
            finally
            {
                try { if (backupDoc != null) backupDoc.Close(false); } catch { }
                try { if (File.Exists(tempSourceHtml)) File.Delete(tempSourceHtml); } catch { }
                try { if (File.Exists(tempTargetHtml)) File.Delete(tempTargetHtml); } catch { }
            }
        }

        private string HttpTranslateHtml(string htmlContent, string targetLang, string sourceLang)
        {
            string srcPart = sourceLang == "Auto Detect" ? "" : "\"sourceLanguage\":\"" + sourceLang + "\",";
            string json = "{" + srcPart + "\"htmlContent\":" + SimpleJsonEncode(htmlContent) + ",\"targetLanguage\":\"" + targetLang + "\"}";
            string response = MakePostRequest(json);

            string result = ParseJsonValue(response, "translatedHtml");
            if (result != null) return result;
            throw new Exception("Unexpected HTML server response format: " + response);
        }

        private string ParseJsonValue(string json, string key)
        {
            string keyPattern = "\"" + key + "\":\"";
            int idx = json.IndexOf(keyPattern);
            if (idx == -1) return null;

            int start = idx + keyPattern.Length;
            StringBuilder sb = new StringBuilder();
            for (int i = start; i < json.Length; i++)
            {
                char c = json[i];
                if (c == '\"')
                {
                    if (i > 0 && json[i - 1] != '\\')
                    {
                        break;
                    }
                }
                sb.Append(c);
            }
            return SimpleJsonDecode(sb.ToString());
        }

        private string AsyncTranslateText(string text, string targetLang, string sourceLang)
        {
            string result = null;
            Exception error = null;

            ProgressForm progress = new ProgressForm("AI Translate: Translating selection...");

            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                try
                {
                    result = HttpTranslateText(text, targetLang, sourceLang);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                finally
                {
                    progress.Invoke(new Action(() => progress.Close()));
                }
            });

            thread.Start();
            progress.ShowDialog();

            if (error != null) throw error;
            return result;
        }

        // Thực hiện HTTP POST đơn lẻ
        private string HttpTranslateText(string text, string targetLang, string sourceLang)
        {
            string srcPart = sourceLang == "Auto Detect" ? "" : "\"sourceLanguage\":\"" + sourceLang + "\",";
            string json = "{" + srcPart + "\"texts\":[" + SimpleJsonEncode(text) + "],\"targetLanguage\":\"" + targetLang + "\"}";
            string response = MakePostRequest(json);

            string[] arr = ParseJsonArray(response);
            if (arr != null && arr.Length > 0) return arr[0];
            throw new Exception("Unexpected server response format: " + response);
        }

        // Thực hiện HTTP POST mảng
        private string[] HttpTranslateArray(string[] texts, string targetLang, string sourceLang)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("{");
            if (sourceLang != "Auto Detect")
            {
                sb.Append("\"sourceLanguage\":\"" + sourceLang + "\",");
            }
            sb.Append("\"texts\":[");
            for (int i = 0; i < texts.Length; i++)
            {
                sb.Append(SimpleJsonEncode(texts[i]));
                if (i < texts.Length - 1) sb.Append(",");
            }
            sb.Append("],\"targetLanguage\":\"" + targetLang + "\"}");

            string response = MakePostRequest(sb.ToString());
            string[] result = ParseJsonArray(response);
            if (result == null || result.Length != texts.Length)
            {
                throw new Exception("Translation count mismatch. Response: " + response);
            }
            return result;
        }

        private string MakePostRequest(string jsonPayload)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072; // TLS 1.2
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(serverUrl + "/api/translate");
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Headers.Add("Authorization", "Bearer " + clientToken);

            byte[] bytes = Encoding.UTF8.GetBytes(jsonPayload);
            request.ContentLength = bytes.Length;

            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream responseStream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(responseStream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private string SimpleJsonEncode(string text)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("\"");
            foreach (char c in text)
            {
                switch (c)
                {
                    case '\"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        int codepoint = Convert.ToInt32(c);
                        if (codepoint >= 32 && codepoint <= 126)
                        {
                            sb.Append(c);
                        }
                        else
                        {
                            sb.Append("\\u" + codepoint.ToString("x4"));
                        }
                        break;
                }
            }
            sb.Append("\"");
            return sb.ToString();
        }

        private string SimpleJsonDecode(string text)
        {
            return text.Replace("\\n", "\n")
                       .Replace("\\r", "\r")
                       .Replace("\\t", "\t")
                       .Replace("\\\"", "\"")
                       .Replace("\\\\", "\\");
        }

        private string[] ParseJsonArray(string json)
        {
            int startIdx = json.IndexOf("[");
            int endIdx = json.LastIndexOf("]");
            if (startIdx == -1 || endIdx == -1) return null;

            string content = json.Substring(startIdx + 1, endIdx - startIdx - 1);
            System.Collections.Generic.List<string> list = new System.Collections.Generic.List<string>();

            bool inQuote = false;
            StringBuilder current = new StringBuilder();
            for (int i = 0; i < content.Length; i++)
            {
                char c = content[i];
                if (c == '\"')
                {
                    if (i > 0 && content[i - 1] == '\\')
                    {
                        current.Append(c);
                    }
                    else
                    {
                        inQuote = !inQuote;
                        if (!inQuote)
                        {
                            list.Add(SimpleJsonDecode(current.ToString()));
                            current.Length = 0;
                        }
                    }
                }
                else if (inQuote)
                {
                    current.Append(c);
                }
            }

            return list.ToArray();
        }

        // Trích xuất index của Paragraph chứa Range
        private int GetParagraphIndex(dynamic anchor)
        {
            try
            {
                dynamic doc = anchor.Document;
                dynamic testRange = doc.Range(0, anchor.End);
                return testRange.Paragraphs.Count;
            }
            catch { return 1; }
        }

        // PowerPoint Shape Text Extractor Helper (Recursive for Groups & SmartArt)
        private void ExtractPPTShapeTexts(dynamic shape, System.Collections.Generic.List<dynamic> textRanges, System.Collections.Generic.List<string> texts)
        {
            try
            {
                // Group (Type == 6 / Group)
                if (shape.Type == 6)
                {
                    foreach (dynamic subShape in shape.GroupItems)
                    {
                        ExtractPPTShapeTexts(subShape, textRanges, texts);
                    }
                }
                // SmartArt (HasSmartArt == -1 / msoTrue)
                else if (shape.HasSmartArt == -1)
                {
                    foreach (dynamic node in shape.SmartArt.AllNodes)
                    {
                        dynamic tr = node.TextFrame2.TextRange;
                        string text = tr.Text;
                        if (!string.IsNullOrEmpty(text))
                        {
                            textRanges.Add(tr);
                            texts.Add(text);
                        }
                    }
                }
                // Normal shape with TextFrame
                else if (shape.HasTextFrame == -1)
                {
                    dynamic tf = shape.TextFrame;
                    if (tf.HasText == -1)
                    {
                        dynamic tr = tf.TextRange;
                        string text = tr.Text;
                        if (!string.IsNullOrEmpty(text))
                        {
                            textRanges.Add(tr);
                            texts.Add(text);
                        }
                    }
                }
            }
            catch { }
        }

        // Excel Shapes, SmartArts and Comments Translator
        private void TranslateExcelShapesAndComments(dynamic sheet, string targetLang, string sourceLang)
        {
            // 1. Dịch các Shapes, SmartArts, TextBoxes, Grouped Items trong Excel
            try
            {
                System.Collections.Generic.List<dynamic> textFrames = new System.Collections.Generic.List<dynamic>();
                System.Collections.Generic.List<string> texts = new System.Collections.Generic.List<string>();

                foreach (dynamic shape in sheet.Shapes)
                {
                    ExtractExcelShapeTexts(shape, textFrames, texts, targetLang, sourceLang);
                }

                if (texts.Count > 0)
                {
                    string[] translated = HttpTranslateArray(texts.ToArray(), targetLang, sourceLang);
                    for (int i = 0; i < textFrames.Count; i++)
                    {
                        textFrames[i].Characters().Text = translated[i];
                    }
                }
            }
            catch { }

            // 2. Dịch các Comments trong Excel (Classic)
            try
            {
                foreach (dynamic comment in sheet.Comments)
                {
                    string text = comment.Text();
                    if (!string.IsNullOrEmpty(text))
                    {
                        comment.Text(HttpTranslateText(text, targetLang, sourceLang));
                    }
                }
            }
            catch { }

            // 3. Dịch các Threaded Comments trong Excel (Modern)
            try
            {
                foreach (dynamic comment in sheet.CommentsThreaded)
                {
                    string text = comment.Text;
                    if (!string.IsNullOrEmpty(text))
                    {
                        comment.Text = HttpTranslateText(text, targetLang, sourceLang);
                    }
                    foreach (dynamic reply in comment.Replies)
                    {
                        string replyText = reply.Text;
                        if (!string.IsNullOrEmpty(replyText))
                        {
                            reply.Text = HttpTranslateText(replyText, targetLang, sourceLang);
                        }
                    }
                }
            }
            catch { }
        }

        // Excel Shape Text Extractor Helper
        private void ExtractExcelShapeTexts(dynamic shape, System.Collections.Generic.List<dynamic> textFrames, System.Collections.Generic.List<string> texts, string targetLang, string sourceLang)
        {
            try
            {
                // Group (Type == 6 / Group)
                if (shape.Type == 6)
                {
                    foreach (dynamic subShape in shape.GroupItems)
                    {
                        ExtractExcelShapeTexts(subShape, textFrames, texts, targetLang, sourceLang);
                    }
                }
                // SmartArt (HasSmartArt == -1)
                else if (shape.HasSmartArt == -1)
                {
                    foreach (dynamic node in shape.SmartArt.AllNodes)
                    {
                        dynamic tf = node.TextFrame2;
                        if (tf.HasText == -1)
                        {
                            dynamic tr = tf.TextRange;
                            string text = tr.Text;
                            if (!string.IsNullOrEmpty(text))
                            {
                                tr.Text = HttpTranslateText(text, targetLang, sourceLang);
                            }
                        }
                    }
                }
                // Normal Shape with TextFrame
                else
                {
                    dynamic tf = shape.TextFrame;
                    string text = tf.Characters().Text;
                    if (!string.IsNullOrEmpty(text))
                    {
                        textFrames.Add(tf);
                        texts.Add(text);
                    }
                }
            }
            catch { }
        }

        // Word Shapes, SmartArts and Comments Translator
        private void TranslateWordShapesAndComments(dynamic doc, string targetLang, string sourceLang)
        {
            // 1. Dịch các Comments trong Word
            try
            {
                foreach (dynamic comment in doc.Comments)
                {
                    dynamic range = comment.Range;
                    string text = range.Text;
                    if (!string.IsNullOrEmpty(text))
                    {
                        range.Text = HttpTranslateText(text, targetLang, sourceLang);
                    }
                }
            }
            catch { }

            // 2. Dịch các Floating Shapes trong Word
            try
            {
                foreach (dynamic shape in doc.Shapes)
                {
                    TranslateWordShape(shape, targetLang, sourceLang);
                }
            }
            catch { }

            // 3. Dịch các Inline Shapes trong Word
            try
            {
                foreach (dynamic inlineShape in doc.InlineShapes)
                {
                    if (inlineShape.HasTextFrame == -1)
                    {
                        TranslateWordShape(inlineShape, targetLang, sourceLang);
                    }
                }
            }
            catch { }
        }

        // Word Shape Text Translator Helper
        private void TranslateWordShape(dynamic shape, string targetLang, string sourceLang)
        {
            try
            {
                // Group (Type == 6 / Group)
                if (shape.Type == 6)
                {
                    foreach (dynamic subShape in shape.GroupItems)
                    {
                        TranslateWordShape(subShape, targetLang, sourceLang);
                    }
                }
                // SmartArt (HasSmartArt == -1)
                else if (shape.HasSmartArt == -1)
                {
                    foreach (dynamic node in shape.SmartArt.AllNodes)
                    {
                        dynamic tf = node.TextFrame2;
                        if (tf.HasText == -1)
                        {
                            dynamic tr = tf.TextRange;
                            string text = tr.Text;
                            if (!string.IsNullOrEmpty(text))
                            {
                                tr.Text = HttpTranslateText(text, targetLang, sourceLang);
                            }
                        }
                    }
                }
                // Normal Shape with TextFrame (như Text Box)
                else if (shape.HasTextFrame == -1)
                {
                    dynamic tf = shape.TextFrame;
                    if (tf.HasText == -1)
                    {
                        dynamic tr = tf.TextRange;
                        string text = tr.Text;
                        if (!string.IsNullOrEmpty(text))
                        {
                            tr.Text = HttpTranslateText(text, targetLang, sourceLang);
                        }
                    }
                }
            }
            catch { }
        }
    }

    // Lớp giao diện tiến trình xử lý (Progress Dialog)
    public class ProgressForm : Form
    {
        private Label lblMessage;
        private ProgressBar progressBar;

        public ProgressForm(string message)
        {
            this.Text = "AI Translate";
            this.Size = new System.Drawing.Size(320, 110);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.Color.FromArgb(243, 242, 241);

            lblMessage = new Label();
            lblMessage.Text = message;
            lblMessage.Location = new System.Drawing.Point(20, 20);
            lblMessage.Size = new System.Drawing.Size(280, 20);
            lblMessage.Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular);

            progressBar = new ProgressBar();
            progressBar.Location = new System.Drawing.Point(20, 45);
            progressBar.Size = new System.Drawing.Size(280, 15);
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;

            this.Controls.Add(lblMessage);
            this.Controls.Add(progressBar);
        }

        public void UpdateMessage(string msg)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(UpdateMessage), msg);
                return;
            }
            lblMessage.Text = msg;
        }
    }

    // Lớp lưu trữ trạng thái để phục vụ tính năng Undo trong Excel
    public class ExcelUndoState
    {
        public string WorkbookName { get; set; }
        public string WorksheetName { get; set; }
        public string Address { get; set; }
        public object OriginalValues { get; set; }
    }

    // Lớp bổ trợ thông tin dịch thuật công thức Excel
    public class FormulaTranslationInfo
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public string OriginalFormula { get; set; }
        public System.Collections.Generic.List<string> Literals { get; set; }
        public System.Collections.Generic.List<int[]> Ranges { get; set; }
        public int TextListStartIndex { get; set; }
    }

    // Lớp sao lưu Shape để khôi phục sau khi dịch Word HTML
    public class ShapeBackup
    {
        public int ParagraphIndex { get; set; }
        public string Name { get; set; }
        public int Type { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public int BackupShapeIndex { get; set; }
    }

    // Lớp sao lưu Comment để khôi phục sau khi dịch Word HTML
    public class CommentBackup
    {
        public int ParagraphIndex { get; set; }
        public string Text { get; set; }
        public string Author { get; set; }
    }
}
