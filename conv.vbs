' CreateObject Method
'     语法：CreateObject(appname.objectType, [servename])
'     解释：appname, 必要， Variant(字符串)。提供该对象的应用程序名。
'           objecttype, 必要，Variant。带创建对象的类型或是类。
'           servename，可选，Variant。要在其上创建对象的网络服务器名称。
'
'     说明：要创建ActiveX对象，只需将CreateObject返回的对象赋给一个对象变量：
'     例子：Set oExcel = CreateObject("Excel.Application")

' 声明一个对象变量，并使用动态创建方法创建该对象
Dim oExcel
Set oExcel = CreateObject("Excel.Application")

' 1) 使Excel可见
oExcel.Visible = true

' 2) 更改Excel标题栏
'oExcel.caption = "qyx's vbs"

' 3) 添加一个新的工作薄
Set newWb = oExcel.Workbooks.Add
Set newWbWs = newWb.Sheets

' 4) 打开已存在的工作薄
Set wb = oExcel.Workbooks.Open("D:\Desktop\222\123.xls")
Set ws = wb.Sheets
'ws.Add(null, ws(ws.Count)).Name = "Sheet3"
'ws("Sheet3").Range("A1").Value = "ASD"
'ws("Sheet2").Copy null,newWbWs("Sheet1")
ws(8).Copy null,newWbWs(1)
newWbWs(ws(8).Name).Name = "123"


'wb.Save
wb.Close

'newWb.SaveAs "D:\Desktop\222\456.xlsx",51
newWb.SaveAs "D:\Desktop\222\456.xlsx"
newWb.Close



' 5) 设置活动工作表
' oExcel.Worksheets(1).Activate
' 或者
' oExcel.worksheets("Sheet1").activate


' 6) 给单元格赋值
'oExcel.cells(1,1).value = "This is column A, row 1"

' 7) 设置指定行的高度（单位：磅, 0.035cm）
'oExcel.activeSheet.rows(2).rowHeight = 1/0.035 ' 1cm

' 8) 设置指定列的宽度（单位：字符个数）
'oExcel.activeSheet.columns(1).columnWidth = 5

' 9) 在第8行之前插入分页符
'oExcel.worksheets(1).rows(8).pagebreak = 1

' 10) 在第8列之前删除分页符
'oExcel.worksheets(1).columns(8).pagebreak = 0

' 11) 指定边框线宽度
'     说明：1-左 2-右 3-顶 4-底 5-\ 6-/
'oExcel.activeSheet.range("B3:D4").borders(5).weight = 3

' 12) 清除第1行第4列单元格公式
'oExcel.activeSheet.cells(1,4).clearcontents
' oExcel.activeSheet.cells(1,4).value = ""

' 13) 设置第一行字体属性
'oExcel.activeSheet.rows(1).font.name = "黑体"
'oExcel.activesheet.rows(1).font.color = vbRed
'oExcel.activeSheet.rows(1).font.bold = true
'oExcel.activesheet.rows(1).font.underLine = true
' 14) 页面设置
' a) 页眉
'oExcel.activeSheet.pageSetup.centerHeader = "报表演示"

' b) 页脚
'oExcel.activeSheet.pageSetup.centerFooter = "第&P页"

' c) 页眉到顶端边距2cm
'oExcel.activeSheet.pageSetup.headerMargin = 2/0.035 

' d) 页脚到底端边距3cm
'oExcel.activeSheet.pageSetup.footerMargin = 3/0.035

' e) 顶边距2cm
'oExcel.activeSheet.pageSetup.topMargin = 2/0.035

' f) 底边距2cm
'oExcel.activeSheet.pageSetup.bottomMargin = 2/0.035

' g) 左边距2cm
'oExcel.activeSheet.pageSetup.leftMargin = 2/0.035

' h) 右边距2cm
'oExcel.activeSheet.pageSetup.rightMargin = 2/0.035

' i) 页眉水平居中
'oExcel.activeSheet.pageSetup.centerVertically = 2/0.035

' k) 打印单元格网线
'oExcel.activeSheet.pageSetup.printGridLines = true

' 15) 拷贝与粘贴操作
' a) 拷贝整个工作表
' oExcel.activeSheet.copy    ' 未测试

' b) 拷贝指定区域
'oExcel.activeSheet.range("A1:E2").copy

' c) 从A1位置开始粘贴
'oExcel.activeSheet.range("A1").pasteSpecial

' d) 从文件尾部开始粘贴
' oExcel.activeSheet.range.pasteSpecial '未测试

' 16) 插入一行或一列
'oExcel.activeSheet.rows(2).insert
'oExcel.activeSheet.columns(1).insert

' 17) 删除一行或一列
'oExcel.activeSheet.rows(2).delete
'oExcel.activeSheet.columns(1).delete

' 18) 打印预览工作表
'oExcel.activeSheet.printPreview

' 19) 打印输出工作表
'oExcel.activeSheet.printOut
' 20) 工作表保存
'oExcel.activeWorkBook.saveAs "D:\Desktop\222\te.xls", 56
'oExcel.activeWorkBook.Save

' 21) 关闭退出
' 关闭工作薄
'oExcel.activeWorkBook.Close

' 使用应用程序对象的quit方法关闭Excel
oExcel.Quit

' 释放该对象变量
Set oExcel = Nothing

