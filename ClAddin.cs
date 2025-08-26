using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NavisPlugun
{
    [Plugin("NavisPlugin", "CONN", DisplayName = "Navis Plugin :")]
    public class ClAddin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                // 获取当前文档
                Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

                if (doc.Models.Count == 0)
                {
                    MessageBox.Show("没有打开任何模型", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                }

                // 弹出选择对话框
                var result = MessageBox.Show("请选择操作：\n\n是(Y) - 导出结构树到CSV\n否(N) - 从CSV文件修改模型\n取消 - 退出程序",
                    "选择操作", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // 导出结构树
                    StructureTreeExporter exporter = new StructureTreeExporter();
                    exporter.ExportStructureTree(doc);
                }
                else if (result == DialogResult.No)
                {
                    // 从CSV文件修改模型
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                    openFileDialog.Title = "选择CSV文件";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ModelModifier modifier = new ModelModifier();
                        modifier.ModifyModelFromCSV(doc, openFileDialog.FileName);//修改项等
                        SelectionSetProcessor selectionSetProcessor = new SelectionSetProcessor();
                        selectionSetProcessor.CreateSelectionSetsFromCSV(doc, openFileDialog.FileName);//创建集合
                    }
                }
                else
                {
                    // 取消操作
                    MessageBox.Show("操作已取消", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行过程中发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0;
        }
    }

    /// <summary>
    /// 结构树导出类
    /// </summary>
    public class StructureTreeExporter
    {
        private DataTable _structureTable;
        private const string TargetParentName = "48\"-ATM-2111-117-001-A1A-N";

        /// <summary>
        /// 导出结构树到CSV
        /// </summary>
        public void ExportStructureTree(Document doc)
        {
            try
            {
                // 创建DataTable来存储结构树数据
                _structureTable = CreateStructureDataTable();

                // 加载模型结构到DataTable
                LoadModelStructure(doc);

                // 导出DataTable到CSV
                ExportDataTableToCSV();

                MessageBox.Show($"成功导出结构树数据，共 {_structureTable.Rows.Count} 个节点", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出结构树时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 创建结构树DataTable
        /// </summary>
        private DataTable CreateStructureDataTable()
        {
            DataTable table = new DataTable("ModelStructure");

            // 添加固定列
            table.Columns.Add("NodeName", typeof(string));
            table.Columns.Add("NodeLevel", typeof(int));
            table.Columns.Add("ClassName", typeof(string));
            table.Columns.Add("IsHidden", typeof(bool));
            table.Columns.Add("ChildrenCount", typeof(int));
            table.Columns.Add("ModelSource", typeof(string));

            // 添加处理状态列（用于后续修改）
            table.Columns.Add("隐藏", typeof(string));
            table.Columns.Add("颜色", typeof(string));
            table.Columns.Add("透明", typeof(string));

            // 添加默认的三个层级列
            table.Columns.Add("Level1", typeof(string));
            table.Columns.Add("Level2", typeof(string));
            table.Columns.Add("Level3", typeof(string));

            return table;
        }

        /// <summary>
        /// 加载模型结构到DataTable
        /// </summary>
        private void LoadModelStructure(Document doc)
        {
            // 清空DataTable
            _structureTable.Clear();

            // 循环现有模型
            foreach (var documentModel in doc.Models)
            {
                var modelItemList = documentModel.RootItem.Descendants;
                Model model = documentModel;

                // 获取直接子节点
                var modelItems = modelItemList.Where(o => o.Parent == model.RootItem);

                if (modelItems.Any())
                {
                    foreach (var quItem in modelItems)
                    {
                        // 添加节点到DataTable（从层级0开始）
                        AddNodeToDataTable(quItem, 0, model.RootItem.DisplayName, model.SourceFileName, new List<string>());

                        // 递归处理子节点
                        if (quItem.Children.Any())
                        {
                            LoadChildToDataTable(quItem.Children, quItem, 1, model.SourceFileName, new List<string> { quItem.DisplayName });
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 递归处理子节点到DataTable
        /// </summary>
        private void LoadChildToDataTable(IEnumerable<ModelItem> modelItemEnumerableCollection, ModelItem parentItem,
            int level, string modelSource, List<string> parentHierarchy)
        {
            // 使用LINQ筛选直接子节点
            var query = modelItemEnumerableCollection.Where(o => o.Parent == parentItem);

            if (query.Any())
            {
                foreach (var quItem in query)
                {
                    // 创建当前层级路径
                    var currentHierarchy = new List<string>(parentHierarchy) { quItem.DisplayName };

                    // 添加节点到DataTable
                    AddNodeToDataTable(quItem, level, parentItem.DisplayName, modelSource, currentHierarchy);

                    // 递归处理子节点
                    if (quItem.Children.Any())
                    {
                        LoadChildToDataTable(quItem.Children, quItem, level + 1, modelSource, currentHierarchy);
                    }
                }
            }
        }

        /// <summary>
        /// 添加节点到DataTable
        /// </summary>
        private void AddNodeToDataTable(ModelItem item, int level, string parentName, string modelSource, List<string> hierarchy)
        {
            // 确保有足够的层级列
            EnsureHierarchyColumns(hierarchy.Count);

            DataRow row = _structureTable.NewRow();

            // 设置节点名称
            row["NodeName"] = item.DisplayName;

            // 设置层级列
            for (int i = 0; i < hierarchy.Count; i++)
            {
                string columnName = GetHierarchyColumnName(i + 1);
                if (_structureTable.Columns.Contains(columnName))
                {
                    row[columnName] = hierarchy[i];
                }
            }

            // 设置其他列
            row["NodeLevel"] = level;
            row["ClassName"] = item.ClassDisplayName;
            row["IsHidden"] = item.IsHidden;
            row["ChildrenCount"] = item.Children.Any();
            row["ModelSource"] = modelSource;

            // 处理状态列初始为空
            row["隐藏"] = "";
            row["颜色"] = "";
            row["透明"] = "";

            _structureTable.Rows.Add(row);
        }

        /// <summary>
        /// 确保有足够的层级列
        /// </summary>
        private void EnsureHierarchyColumns(int requiredLevels)
        {
            int currentLevels = GetCurrentHierarchyColumnCount();

            // 如果需要更多层级列，则添加
            for (int i = currentLevels + 1; i <= requiredLevels; i++)
            {
                string columnName = GetHierarchyColumnName(i);
                if (!_structureTable.Columns.Contains(columnName))
                {
                    _structureTable.Columns.Add(columnName, typeof(string));
                }
            }
        }

        /// <summary>
        /// 获取层级列名称
        /// </summary>
        private string GetHierarchyColumnName(int level)
        {
            return $"Level{level}";
        }

        /// <summary>
        /// 获取当前层级列数量
        /// </summary>
        private int GetCurrentHierarchyColumnCount()
        {
            int count = 0;
            foreach (DataColumn column in _structureTable.Columns)
            {
                if (column.ColumnName.StartsWith("Level") && column.ColumnName != "NodeLevel")
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// 导出DataTable到CSV文件
        /// </summary>
        private void ExportDataTableToCSV()
        {
            try
            {
                // 保存文件到桌面
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"模型结构树导出_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                string filePath = Path.Combine(desktopPath, fileName);

                using (var writer = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
                {
                    // 写入表头
                    var headers = _structureTable.Columns.Cast<DataColumn>()
                        .Select(column => EscapeCsvField(column.ColumnName));
                    writer.WriteLine(string.Join(",", headers));

                    // 写入数据行
                    foreach (DataRow row in _structureTable.Rows)
                    {
                        var fields = row.ItemArray.Select(field => EscapeCsvField(field?.ToString() ?? ""));
                        writer.WriteLine(string.Join(",", fields));
                    }
                }

                MessageBox.Show($"CSV文件已保存到：{filePath}", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出CSV时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 转义CSV字段
        /// </summary>
        private string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "\"\"";

            if (field.Contains(",") || field.Contains("\"") || field.Contains("\r") || field.Contains("\n"))
            {
                field = field.Replace("\"", "\"\"");
                return $"\"{field}\"";
            }

            return field;
        }
    }

    /// <summary>
    /// 模型修改类
    /// </summary>
    public class ModelModifier
    {
        private Dictionary<string, List<ModelItem>> _modelItemCache;
        /// <summary>
        /// 从CSV文件修改模型
        /// </summary>
        public void ModifyModelFromCSV(Document doc, string csvFilePath)
        {
            try
            {
                DataTable csvData = ReadCSVFile(csvFilePath);
                if (csvData == null || csvData.Rows.Count == 0)
                    return;

                // 初始化缓存
                InitializeModelItemCache(doc);

                using (Transaction transaction = doc.BeginTransaction("ModifyModelFromCSV"))
                {
                    int processedCount = 0;
                    int notFoundCount = 0;
                    // 用于存储选择集分组
                    Dictionary<string, ModelItemCollection> selectionSets = new Dictionary<string, ModelItemCollection>();

                    foreach (DataRow row in csvData.Rows)
                    {
                        string nodeName = row["NodeName"]?.ToString() ?? "";
                        string hideValue = row["隐藏"]?.ToString() ?? "";
                        string colorValue = row["颜色"]?.ToString() ?? "";
                        string transparencyValue = row["透明"]?.ToString() ?? "";
                        string selectionSetValue = row["集合"]?.ToString() ?? ""; // 新增的选择集列

                        // 获取层级信息
                        int nodeLevel = Convert.ToInt32(row["NodeLevel"] ?? 0);
                        List<string> hierarchy = ExtractHierarchyFromRow(row);

                        // 使用缓存快速查找
                        ModelItem item = FindModelItemByCache(nodeName, hierarchy, nodeLevel);
                        if (item == null)
                        {
                            notFoundCount++;
                            continue;
                        }

                        // 处理选择集
                        //ProcessSelectionSetOperations(doc, item, selectionSetValue, selectionSets);

                        // 处理各项操作
                        ProcessItemOperations(doc, item, hideValue, colorValue, transparencyValue);
                        processedCount++;
                    }
                    // 创建选择集
                    //CreateSelectionSets(doc, selectionSets);
                    transaction.Commit();

                    MessageBox.Show($"处理完成！\n找到项目: {processedCount}\n未找到项目: {notFoundCount}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"修改模型时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 清理缓存
                _modelItemCache?.Clear();
                _modelItemCache = null;
            }
        }

        
       

        /// <summary>
        /// 处理单个项目的各项操作
        /// </summary>
        private void ProcessItemOperations(Document doc, ModelItem item, string hideValue, string colorValue, string transparencyValue)
        {
            // 处理隐藏
            if (!string.IsNullOrEmpty(hideValue) && (hideValue.ToUpper() == "TRUE" || hideValue.ToUpper() == "Y"))
            {
                doc.State.SetOverrideHideRender(item, true);
            }

            // 处理颜色
            if (!string.IsNullOrEmpty(colorValue))
            {
                Color color = ParseColor(colorValue);
                if (color != null)
                {
                    ModelItemCollection selection = new ModelItemCollection();
                    selection.Add(item);
                    doc.State.OverrideColor(selection, color);
                }
            }

            // 处理透明度
            if (!string.IsNullOrEmpty(transparencyValue) &&
                double.TryParse(transparencyValue, out double transparency))
            {
                ModelItemCollection selection = new ModelItemCollection();
                selection.Add(item);
                doc.State.OverrideTransparency(selection, transparency);
            }
        }

        /// <summary>
        /// 初始化模型项缓存
        /// </summary>
        private void InitializeModelItemCache(Document doc)
        {
            _modelItemCache = new Dictionary<string, List<ModelItem>>();

            foreach (Model model in doc.Models)
            {
                foreach (ModelItem item in model.RootItem.DescendantsAndSelf)
                {
                    string key = item.DisplayName;
                    if (!_modelItemCache.ContainsKey(key))
                    {
                        _modelItemCache[key] = new List<ModelItem>();
                    }
                    _modelItemCache[key].Add(item);
                }
            }
        }

        /// <summary>
        /// 使用缓存快速查找模型项
        /// </summary>
        private ModelItem FindModelItemByCache(string nodeName, List<string> hierarchy, int expectedLevel)
        {
            if (_modelItemCache == null || !_modelItemCache.ContainsKey(nodeName))
                return null;

            var candidates = _modelItemCache[nodeName];

            if (candidates.Count == 1)
                return candidates[0];

            // 多个候选项时使用层级筛选
            foreach (var item in candidates)
            {
                if (CheckItemHierarchy(item, hierarchy, expectedLevel))
                {
                    return item;
                }
            }

            return null;
        }
        /// <summary>
        /// 检查项的层级是否匹配（为缓存方法添加）
        /// </summary>
        private bool CheckItemHierarchy(ModelItem item, List<string> expectedHierarchy, int expectedLevel)
        {
            if (expectedHierarchy == null || expectedHierarchy.Count == 0)
                return true;

            // 检查层级深度
            int actualLevel = GetItemLevel(item);
            if (actualLevel != expectedLevel)
                return false;

            // 检查层级路径
            List<string> actualHierarchy = GetItemHierarchy(item);
            if (actualHierarchy.Count != expectedHierarchy.Count)
                return false;

            for (int i = 0; i < expectedHierarchy.Count; i++)
            {
                if (actualHierarchy[i] != expectedHierarchy[i])
                    return false;
            }

            return true;
        }

        /// <summary>
        /// 获取项的层级深度
        /// </summary>
        private int GetItemLevel(ModelItem item)
        {
            int level = 0;
            ModelItem current = item;
            while (current.Parent != null && current.Parent.Parent != null) // 排除根节点
            {
                level++;
                current = current.Parent;
            }
            return level;
        }

        /// <summary>
        /// 获取项的完整层级路径
        /// </summary>
        private List<string> GetItemHierarchy(ModelItem item)
        {
            var hierarchy = new List<string>();
            ModelItem current = item;

            // 向上遍历直到根节点的直接子节点
            while (current != null && current.Parent != null && current.Parent.Parent != null)
            {
                hierarchy.Insert(0, current.DisplayName);
                current = current.Parent;
            }

            return hierarchy;
        }

        /// <summary>
        /// 从数据行提取层级信息
        /// </summary>
        private List<string> ExtractHierarchyFromRow(DataRow row)
        {
            var hierarchy = new List<string>();

            for (int i = 1; i <= 20; i++) // 假设最多20个层级
            {
                string levelColumnName = $"Level{i}";
                if (row.Table.Columns.Contains(levelColumnName))
                {
                    string levelValue = row[levelColumnName]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(levelValue))
                    {
                        hierarchy.Add(levelValue);
                    }
                }
                else
                {
                    break;
                }
            }

            return hierarchy;
        }
        /// <summary>
        /// 读取CSV文件
        /// </summary>
        private DataTable ReadCSVFile(string filePath)
        {
            DataTable table = new DataTable();

            using (var reader = new StreamReader(filePath, System.Text.Encoding.UTF8))
            {
                // 读取表头
                string headerLine = reader.ReadLine();
                if (headerLine == null) return table;

                var headers = ParseCsvLine(headerLine);
                foreach (var header in headers)
                {
                    table.Columns.Add(header, typeof(string));
                }

                // 读取数据行
                while (!reader.EndOfStream)
                {
                    string dataLine = reader.ReadLine();
                    if (string.IsNullOrEmpty(dataLine)) continue;

                    var fields = ParseCsvLine(dataLine);
                    DataRow row = table.NewRow();

                    for (int i = 0; i < Math.Min(fields.Length, table.Columns.Count); i++)
                    {
                        row[i] = fields[i];
                    }

                    table.Rows.Add(row);
                }
            }

            return table;
        }

        /// <summary>
        /// 解析CSV行
        /// </summary>
        private string[] ParseCsvLine(string line)
        {
            var fields = new List<string>();
            bool inQuotes = false;
            string currentField = "";

            foreach (char c in line)
            {
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField);
                    currentField = "";
                }
                else
                {
                    currentField += c;
                }
            }

            fields.Add(currentField);
            return fields.ToArray();
        }

        /// <summary>
        /// 根据名称查找模型项
        /// </summary>
        private ModelItem FindModelItemByName(Document doc, string nodeName)
        {
            foreach (Model model in doc.Models)
            {
                foreach (ModelItem item in model.RootItem.DescendantsAndSelf)
                {
                    if (item.DisplayName == nodeName)
                    {
                        return item;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 解析颜色字符串
        /// </summary>
        private Color ParseColor(string colorString)
        {
            try
            {
                var parts = colorString.Split(',');
                if (parts.Length == 3)
                {
                    if (byte.TryParse(parts[0].Trim(), out byte r) &&
                        byte.TryParse(parts[1].Trim(), out byte g) &&
                        byte.TryParse(parts[2].Trim(), out byte b))
                    {
                        return Color.FromByteRGB(r, g, b);
                    }
                }
            }
            catch
            {
                // 解析失败
            }
            return null;
        }
    }

    public class SelectionSetProcessor
    {
        /// <summary>
        /// 从CSV数据创建选择集
        /// </summary>
        public void CreateSelectionSetsFromCSV(Document doc, string csvFilePath)
        {
            try
            {
                DataTable csvData = ReadCSVFile(csvFilePath);
                SelectionSetProcessor selectionProcessor = new SelectionSetProcessor();

                if (csvData == null || csvData.Rows.Count == 0)
                {
                    MessageBox.Show("CSV文件为空或无效", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!csvData.Columns.Contains("集合"))
                {
                    MessageBox.Show("CSV文件中没有找到'集合'列", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 按集合名称分组
                var selectionSetGroups = csvData.AsEnumerable()
                    .Where(row => !string.IsNullOrEmpty(row.Field<string>("集合")))
                    .GroupBy(row => row.Field<string>("集合").Trim());

                if (!selectionSetGroups.Any())
                {
                    MessageBox.Show("CSV文件中没有找到有效的集合数据", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 初始化模型项缓存
                Dictionary<string, List<ModelItem>> modelItemCache = InitializeModelItemCache(doc);

                // 在事务内执行所有操作
                using (Transaction transaction = doc.BeginTransaction("CreateSelectionSets"))
                {
                    int setsCreated = 0;
                    int totalItems = 0;

                    foreach (var group in selectionSetGroups)
                    {
                        string selectionSetName = group.Key;
                        ModelItemCollection selectionItems = new ModelItemCollection();

                        foreach (DataRow row in group)
                        {
                            string nodeName = row["NodeName"]?.ToString() ?? "";
                            int nodeLevel = Convert.ToInt32(row["NodeLevel"] ?? 0);
                            List<string> hierarchy = ExtractHierarchyFromRow(row);

                            // 查找模型项
                            ModelItem item = FindModelItemByCache(modelItemCache, nodeName, hierarchy, nodeLevel);
                            if (item != null)
                            {
                                selectionItems.Add(item);
                            }
                        }

                        if (selectionItems.Count > 0)
                        {
                            // 创建选择集
                            if (CreateSelectionSet(doc, selectionSetName, selectionItems))
                            {
                                setsCreated++;
                                totalItems += selectionItems.Count;
                            }
                        }
                    }

                    transaction.Commit();

                    MessageBox.Show($"成功创建 {setsCreated} 个选择集，共包含 {totalItems} 个模型项", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建选择集时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable ReadCSVFile(string filePath)
        {
            DataTable table = new DataTable();

            try
            {
                using (var reader = new StreamReader(filePath, System.Text.Encoding.UTF8))
                {
                    // 读取表头
                    string headerLine = reader.ReadLine();
                    if (headerLine == null) return table;

                    var headers = ParseCsvLine(headerLine);
                    foreach (var header in headers)
                    {
                        table.Columns.Add(header, typeof(string));
                    }

                    // 读取数据行
                    while (!reader.EndOfStream)
                    {
                        string dataLine = reader.ReadLine();
                        if (string.IsNullOrEmpty(dataLine)) continue;

                        var fields = ParseCsvLine(dataLine);
                        DataRow row = table.NewRow();

                        for (int i = 0; i < Math.Min(fields.Length, table.Columns.Count); i++)
                        {
                            row[i] = fields[i];
                        }

                        table.Rows.Add(row);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取CSV文件时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return table;
        }

        private string[] ParseCsvLine(string line)
        {
            var fields = new List<string>();
            bool inQuotes = false;
            string currentField = "";

            foreach (char c in line)
            {
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField);
                    currentField = "";
                }
                else
                {
                    currentField += c;
                }
            }

            fields.Add(currentField);
            return fields.ToArray();
        }

        /// <summary>
        /// 创建选择集（安全版本）
        /// </summary>
        private bool CreateSelectionSet(Document doc, string selectionSetName, ModelItemCollection items)
        {
            try
            {
                // 创建新的选择集
                SelectionSet newSet = new SelectionSet();
                newSet.DisplayName = selectionSetName;

                // 复制模型项到选择集
                ModelItemCollection setItems = new ModelItemCollection();
                foreach (ModelItem item in items)
                {
                    setItems.Add(item);
                }
                newSet.CopyFrom(setItems);

                // 使用文档的SelectionSets.Add方法
                doc.SelectionSets.AddCopy(newSet);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建选择集 '{selectionSetName}' 时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 初始化模型项缓存
        /// </summary>
        private Dictionary<string, List<ModelItem>> InitializeModelItemCache(Document doc)
        {
            var cache = new Dictionary<string, List<ModelItem>>();

            try
            {
                foreach (Model model in doc.Models)
                {
                    foreach (ModelItem item in model.RootItem.DescendantsAndSelf)
                    {
                        string key = item.DisplayName;
                        if (!cache.ContainsKey(key))
                        {
                            cache[key] = new List<ModelItem>();
                        }
                        cache[key].Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化模型缓存时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return cache;
        }

        /// <summary>
        /// 使用缓存查找模型项
        /// </summary>
        private ModelItem FindModelItemByCache(Dictionary<string, List<ModelItem>> cache, string nodeName,
            List<string> hierarchy, int expectedLevel)
        {
            try
            {
                if (cache == null || !cache.ContainsKey(nodeName))
                    return null;

                var candidates = cache[nodeName];

                if (candidates.Count == 1)
                    return candidates[0];

                // 多个候选项时使用层级筛选
                foreach (var item in candidates)
                {
                    if (CheckItemHierarchy(item, hierarchy, expectedLevel))
                    {
                        return item;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"查找模型项 '{nodeName}' 时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// 检查项的层级是否匹配
        /// </summary>
        private bool CheckItemHierarchy(ModelItem item, List<string> expectedHierarchy, int expectedLevel)
        {
            try
            {
                if (expectedHierarchy == null || expectedHierarchy.Count == 0)
                    return true;

                // 检查层级深度
                int actualLevel = GetItemLevel(item);
                if (actualLevel != expectedLevel)
                    return false;

                // 检查层级路径
                List<string> actualHierarchy = GetItemHierarchy(item);
                if (actualHierarchy.Count != expectedHierarchy.Count)
                    return false;

                for (int i = 0; i < expectedHierarchy.Count; i++)
                {
                    if (actualHierarchy[i] != expectedHierarchy[i])
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取项的层级深度
        /// </summary>
        private int GetItemLevel(ModelItem item)
        {
            try
            {
                int level = 0;
                ModelItem current = item;
                while (current.Parent != null && current.Parent.Parent != null) // 排除根节点
                {
                    level++;
                    current = current.Parent;
                }
                return level;
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// 获取项的完整层级路径
        /// </summary>
        private List<string> GetItemHierarchy(ModelItem item)
        {
            var hierarchy = new List<string>();

            try
            {
                ModelItem current = item;

                // 向上遍历直到根节点的直接子节点
                while (current != null && current.Parent != null && current.Parent.Parent != null)
                {
                    hierarchy.Insert(0, current.DisplayName);
                    current = current.Parent;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取层级路径时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return hierarchy;
        }

        /// <summary>
        /// 从数据行提取层级信息
        /// </summary>
        private List<string> ExtractHierarchyFromRow(DataRow row)
        {
            var hierarchy = new List<string>();

            try
            {
                for (int i = 1; i <= 20; i++) // 假设最多20个层级
                {
                    string levelColumnName = $"Level{i}";
                    if (row.Table.Columns.Contains(levelColumnName))
                    {
                        string levelValue = row[levelColumnName]?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(levelValue))
                        {
                            hierarchy.Add(levelValue);
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"提取层级信息时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return hierarchy;
        }
    }
}
1111