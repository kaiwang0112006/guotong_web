import pandas as pd

# 读取Excel文件
excel_file = 'processed_knowledge_base_20250919_155852.xlsx'
df = pd.read_excel(excel_file)

# 处理第二列中的换行符，转换为HTML的<br>标签
if len(df.columns) >= 2:
    second_column = df.columns[1]
    df[second_column] = df[second_column].apply(lambda x: str(x).replace('\n', '<br>') if pd.notna(x) else x)

# 生成HTML表格，设置escape=False以保留HTML标签
html_table = df.to_html(index=False, justify='left', classes='dataframe', escape=False)

# 添加CSS样式和JavaScript搜索功能
# 使用三引号字符串，修复JavaScript正则表达式中的转义问题
html_content = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel数据表格</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            color: #333;
            text-align: center;
        }
        .search-container {
            text-align: center;
            margin: 20px 0;
        }
        .search-box {
            padding: 10px;
            width: 300px;
            font-size: 16px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        .search-btn {
            padding: 10px 20px;
            font-size: 16px;
            background-color: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        .search-btn:hover {
            background-color: #45a049;
        }
        .table-container {
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            overflow-x: auto;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 0 auto;
        }
        th, td {
            text-align: left;
            padding: 12px;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .no-results {
            text-align: center;
            padding: 20px;
            color: #666;
            font-style: italic;
        }
        .highlight {
            background-color: #ffff00;
            font-weight: bold;
        }
    </style>
    <script>
        // 保存表格的原始内容，用于重置
        var originalTableContent = null;

        // 页面加载完成后保存原始表格内容
        document.addEventListener('DOMContentLoaded', function() {
            var table = document.querySelector("table");
            originalTableContent = table.innerHTML;

            // 为搜索框添加回车键事件
            var searchInput = document.getElementById("searchInput");
            searchInput.addEventListener("keypress", function(event) {
                if (event.key === "Enter") {
                    event.preventDefault();
                    document.getElementById("searchBtn").click();
                }
            });
        });

        function searchTable() {
            // 获取搜索关键词
            var input = document.getElementById("searchInput");
            var filter = input.value.toUpperCase();
            var originalFilter = input.value;

            // 获取表格
            var table = document.querySelector("table");

            // 如果搜索框为空，恢复原始表格内容
            if (filter === "") {
                if (originalTableContent) {
                    table.innerHTML = originalTableContent;
                }

                // 显示所有行
                var tr = table.getElementsByTagName("tr");
                for (var i = 1; i < tr.length; i++) {
                    tr[i].style.display = "";
                }

                // 隐藏无结果提示
                var noResultsDiv = document.getElementById("noResults");
                if (noResultsDiv) {
                    noResultsDiv.style.display = "none";
                }

                return;
            }

            // 恢复原始表格内容，以便重新应用高亮
            if (originalTableContent) {
                table.innerHTML = originalTableContent;
            }

            // 获取行
            var tr = table.getElementsByTagName("tr");

            // 用于跟踪是否有匹配结果
            var hasResults = false;

            // 遍历所有行，隐藏不匹配的行
            for (var i = 1; i < tr.length; i++) {
                var td = tr[i].getElementsByTagName("td");
                var match = false;

                // 检查该行的所有单元格
                for (var j = 0; j < td.length; j++) {
                    var cell = td[j];
                    var txtValue = cell.textContent || cell.innerText;

                    if (txtValue.toUpperCase().indexOf(filter) > -1) {
                        match = true;

                        // 高亮显示关键词
                        var highlightedText = highlightKeywords(txtValue, originalFilter);
                        cell.innerHTML = highlightedText;
                    }
                }

                // 显示匹配的行，隐藏不匹配的行
                if (match) {
                    tr[i].style.display = "";
                    hasResults = true;
                } else {
                    tr[i].style.display = "none";
                }
            }

            // 检查是否需要显示无结果提示
            var noResultsDiv = document.getElementById("noResults");
            if (!hasResults && filter !== "") {
                if (!noResultsDiv) {
                    noResultsDiv = document.createElement("div");
                    noResultsDiv.id = "noResults";
                    noResultsDiv.className = "no-results";
                    noResultsDiv.textContent = "没有找到匹配的结果";
                    table.parentNode.insertBefore(noResultsDiv, table.nextSibling);
                } else {
                    noResultsDiv.style.display = "block";
                }
            } else if (noResultsDiv) {
                noResultsDiv.style.display = "none";
            }
        }

        // 高亮显示关键词的函数
        function highlightKeywords(text, keyword) {
            if (!keyword) return text;

            // 创建正则表达式，使用不区分大小写的标志
            var regex = new RegExp('(' + escapeRegExp(keyword) + ')', 'gi');

            // 替换关键词为带有高亮样式的HTML
            return text.replace(regex, '<span class=\"highlight\">$1</span>');
        }

        // 转义正则表达式特殊字符
        function escapeRegExp(string) {
            return string.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
        }
    </script>
</head>
<body>
    <h1>Excel数据表格</h1>

    <div class="search-container">
        <input type="text" id="searchInput" class="search-box" placeholder="输入关键词搜索...">
        <button type="button" id="searchBtn" class="search-btn" onclick="searchTable()">搜索</button>
    </div>

    <div class="table-container">
        ''' + html_table + '''
    </div>
</body>
</html>
'''

# 保存为HTML文件
html_file = 'processed_knowledge_base.html'
with open(html_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"Excel数据已成功转换为HTML表格并保存到 {html_file}")
print(f"表格包含 {len(df)} 行数据")
print(f"表格包含以下列: {list(df.columns)}")