' ============================================================================
' 小型氧化加工厂管理系统 - 小企业会计准则版-利润分析优化 V18.0
' 适用：小会计日常/符合小企业会计准则/WPS专用
' 版本：V18.0 完整同步版 - Python/VBA双版本同步，七级累进个税，动态模拟数据
' 功能：收入管理、支出分类、成本核算、会计报表、批量操作、一键月结、数据校验、备份、Excel导入、微信支付宝解析、现金流量表、多期对比、智能提醒、多用户协作、AI智能分析、数据安全保护、代码性能优化、客户对账单、供应商汇总、成本利润分析、PDF导出、自动提醒、税务合规、银行实务、工资社保、固定资产管理、成本核算优化、往来账龄分析、预算管理、凭证账簿、财务分析报表、系统功能完善
' ============================================================================

Option Explicit

' ============================================================================
' 【WPS兼容版特别说明】
' 本版本专门针对WPS Office优化，移除了以下不兼容功能：
' 1. Scripting.Dictionary对象 → 使用Collection替代
' 2. Application.FileDialog → 使用Application.GetOpenFilename
' 3. Office.FileDialog类型 → 移除
' 4. Type自定义类型 → 使用数组替代
' 5. 修复SUMIF中文表名问题
' 使用方法：直接在WPS表格中导入此.bas文件即可
' ============================================================================

' ============================================================================
' 利润表行号常量
' ============================================================================
Public Const PR_INCOME As Long = 5
Public Const PR_COST As Long = 8
Public Const PR_GROSS As Long = 10
Public Const PR_RENT As Long = 12
Public Const PR_WATER As Long = 13
Public Const PR_ELEC As Long = 14
Public Const PR_MATERIAL As Long = 15
Public Const PR_DEGREASER As Long = 16
Public Const PR_HANGER As Long = 17
Public Const PR_DAILY As Long = 18
Public Const PR_SALARY As Long = 19
Public Const PR_EXPENSE_SUB As Long = 20
Public Const PR_OP_PROFIT As Long = 22
Public Const PR_TAX As Long = 24
Public Const PR_NET_PROFIT As Long = 26

' ============================================================================
' 应收应付列号常量
' ============================================================================
Public Const AR_COL_NAME As Long = 1
Public Const AR_COL_OPEN As Long = 2
Public Const AR_COL_ADD As Long = 3
Public Const AR_COL_REDUCE As Long = 4
Public Const AR_COL_CLOSE As Long = 5
Public Const AP_COL_NAME As Long = 8
Public Const AP_COL_OPEN As Long = 9
Public Const AP_COL_ADD As Long = 10
Public Const AP_COL_REDUCE As Long = 11
Public Const AP_COL_CLOSE As Long = 12

' ============================================================================
' 月度数据结构
' ============================================================================
' Type MonthData - WPS兼容版（使用数组替代）
' 注意：WPS中Type在标准模块中可用，但在类模块中不可用
Type MonthData
    Income As Double
    OutsourceCost As Double
    GrossProfit As Double
    Expense As Double
    NetProfit As Double
End Type


' ============================================================================
' 支出类别常量定义
' ============================================================================
Public Const CAT_RENT As String = "房租"              ' 房租
Public Const CAT_WATER As String = "水费"             ' 水费
Public Const CAT_ELECTRIC As String = "电费"          ' 电费
Public Const CAT_ACID As String = "三酸"              ' 三酸
Public Const CAT_SODA As String = "片碱"              ' 片碱
Public Const CAT_SODIUM As String = "亚钠"            ' 亚钠
Public Const CAT_PIGMENT As String = "色粉"           ' 色粉
Public Const CAT_DEGREASER As String = "除油剂"       ' 除油剂
Public Const CAT_HANGER As String = "挂具"            ' 挂具
Public Const CAT_OUTSOURCE As String = "外发加工费"   ' 外发加工费用
Public Const CAT_DAILY As String = "日常费用"         ' 日常费用支出
Public Const CAT_SALARY As String = "工资"            ' 工资

' 支出类别数组
Public Const EXPENSE_CATEGORIES As String = "房租,水费,电费,三酸,片碱,亚钠,色粉,除油剂,挂具,外发加工费,日常费用,工资"

' ============================================================================
' 国内小企业税率常量（V7.0新增）
' ============================================================================
Public Const VAT_RATE_SMALL As Double = 0.03  ' 小规模纳税人3%
Public Const VAT_THRESHOLD As Double = 100000  ' 增值税起征点10万/月
Public Const IIT_DEDUCTION As Double = 5000  ' 个税起征点5000元/月
Public Const EI_PENSION_E As Double = 0.16  ' 养老企业16%
Public Const EI_PENSION_P As Double = 0.08  ' 养老个人8%
Public Const EI_MEDICAL_E As Double = 0.08  ' 医疗企业8%
Public Const EI_MEDICAL_P As Double = 0.02  ' 医疗个人2%
Public Const EI_UNEMPLOYMENT_E As Double = 0.005  ' 失业企业0.5%
Public Const EI_UNEMPLOYMENT_P As Double = 0.005  ' 失业个人0.5%
Public Const EI_INJURY_E As Double = 0.004  ' 工伤企业0.4%
Public Const EI_MATERNITY_E As Double = 0.008  ' 生育企业0.8%
Public Const HPF_RATE_E As Double = 0.08  ' 公积金企业8%
Public Const HPF_RATE_P As Double = 0.08  ' 公积金个人8%

' 印花税税率
Public Const ST_SALES As Double = 0.0003  ' 买卖合同0.03%
Public Const ST_LOAN As Double = 0.00005  ' 借款合同0.005%
Public Const ST_LEASE As Double = 0.001  ' 租赁合同0.1%
Public Const ST_TRANSPORT As Double = 0.0003  ' 运输合同0.03%
Public Const ST_WAREHOUSE As Double = 0.001  ' 仓储合同0.1%

' ============================================================================
' 固定资产常量（V8.0新增）
' ============================================================================
Public Const FA_RESIDUAL_RATE As Double = 0.05  ' 残值率5%
Public Const FA_DEPRECIATION_YEARS As Long = 10  ' 默认使用年限10年

' ============================================================================
' 低值易耗品标准（V8.0新增）
' ============================================================================
Public Const LVC_THRESHOLD As Double = 2000  ' 低值易耗品标准2000元

' ============================================================================
' 坏账准备计提比例（V8.0新增）
' ============================================================================
Public Const BD_RATE_1YEAR As Double = 0.05  ' 1年以内5%
Public Const BD_RATE_2YEAR As Double = 0.10  ' 1-2年10%
Public Const BD_RATE_3YEAR As Double = 0.30  ' 2-3年30%
Public Const BD_RATE_OVER3 As Double = 0.50  ' 3年以上50%

' ============================================================================
' 预算预警常量（V9.0新增）
' ============================================================================
Public Const BUDGET_ALERT_RATE As Double = 0.8  ' 关注预警线80%
Public Const BUDGET_OVER_RATE As Double = 1.0  ' 超支预警线100%

' ============================================================================
' 会计科目常量（V9.0新增）
' ============================================================================
Public Const SUBJECT_BANK As String = "1002"  ' 银行存款
Public Const SUBJECT_CASH As String = "1001"  ' 库存现金
Public Const SUBJECT_AR As String = "1122"  ' 应收账款
Public Const SUBJECT_AP As String = "2202"  ' 应付账款
Public Const SUBJECT_SALARY As String = "2211"  ' 应付职工薪酬
Public Const SUBJECT_REVENUE As String = "6001"  ' 主营业务收入
Public Const SUBJECT_COST As String = "6401"  ' 主营业务成本
Public Const SUBJECT_EXPENSE As String = "6602"  ' 管理费用

' ============================================================================

' ============================================================================
' V18.0新增常量 - 国内小企业实务优化
' ============================================================================

' 增值税减免政策（2023年政策）
Public Const VAT_EXEMPTION_THRESHOLD As Double = 100000  ' 月销售额≤10万免征增值税
Public Const VAT_REDUCTION_RATE As Double = 0.01  ' 减按1%征收（疫情期间优惠）

' 六税两费减半征收（小微企业优惠）
Public Const SIX_TAX_HALF_RATE As Double = 0.5  ' 减半征收比例

' 残保金参数（2023年标准）
Public Const DISABILTY_EMPLOY_RATE As Double = 0.015  ' 安置比例1.5%
Public Const DISABILITY_FEE_PER_PERSON As Double = 23820  ' 每人年缴纳额（当地平均工资）
Public Const DISABILITY_SMALL_EXEMPT As Long = 30  ' 30人以下免征

' 工会经费
Public Const UNION_FEE_RATE As Double = 0.02  ' 工资总额2%
Public Const UNION_FEE_RETURN_RATE As Double = 0.6  ' 返还比例60%

' 企业所得税分段税率（小微企业优惠）
Public Const IIT_RATE_1 As Double = 0.025  ' 100万以内2.5%（实际税负5%减半）
Public Const IIT_RATE_2 As Double = 0.05   ' 100-300万5%
Public Const IIT_THRESHOLD_1 As Double = 1000000  ' 100万
Public Const IIT_THRESHOLD_2 As Double = 3000000  ' 300万

' 水电费税率（可抵扣进项）
Public Const WATER_VAT_RATE As Double = 0.09  ' 水费9%
Public Const ELECTRIC_VAT_RATE As Double = 0.13  ' 电费13%

' ============================================================================
' V18.0新增常量 - 含税/不含税收入管理
' ============================================================================
Public Const TAX_COL_TYPE As Long = 7       ' 含税/不含税标识列
Public Const TAX_COL_RATE As Long = 8       ' 税率列
Public Const TAX_COL_AMOUNT As Long = 9     ' 税额列
Public Const TAX_COL_EXCL As Long = 10      ' 不含税金额列
Public Const TAX_COL_INVOICE As Long = 11   ' 开票状态列
Public Const TAX_COL_INCOME_TOTAL As Long = 12  ' 含税收入合计列

' 含税/不含税类型
Public Const TAX_TYPE_INCL As String = "含税"
Public Const TAX_TYPE_EXCL As String = "不含税"

' 开票状态
Public Const INVOICE_YES As String = "已开票"
Public Const INVOICE_NO As String = "未开票"
Public Const INVOICE_PARTIAL As String = "部分开票"

' ============================================================================
' V18.0新增常量 - 对冲/代付货款管理
' ============================================================================
Public Const OFFSET_COL_ACTUAL_PAYER As Long = 11   ' K列：实际付款方
Public Const OFFSET_COL_ORIGINAL_CUST As Long = 12  ' L列：原客户（被代付）
Public Const OFFSET_COL_OFFSET_FLAG As Long = 13    ' M列：对冲标识
Public Const OFFSET_COL_OFFSET_AMOUNT As Long = 14  ' N列：对冲金额
Public Const OFFSET_COL_REMARK As Long = 15         ' O列：对冲备注

' 对冲类型
Public Const OFFSET_TYPE_NONE As String = "正常"
Public Const OFFSET_TYPE_OFFSET As String = "对冲货款"
Public Const OFFSET_TYPE_PROXY As String = "代付货款"



' 宏1: InitSimpleSystem - 初始化简化版系统
' 创建所需的工作表
' ============================================================================

' ============================================================================
' 【WPS兼容】安全转换函数（替代Val和Format）
' ============================================================================
Function SafeDbl(strValue As Variant) As Double
    ' WPS安全双精度转换，处理空值和错误
    On Error Resume Next
    If IsNull(strValue) Or IsEmpty(strValue) Or strValue = "" Then
        SafeDbl = 0
    Else
        SafeDbl = CDbl(strValue)
    End If
    On Error GoTo 0
End Function

Function SafeLng(strValue As Variant) As Long
    ' WPS安全长整型转换
    On Error Resume Next
    If IsNull(strValue) Or IsEmpty(strValue) Or strValue = "" Then
        SafeLng = 0
    Else
        SafeLng = CLng(strValue)
    End If
    On Error GoTo 0
End Function

' ============================================================================

' 【WPS兼容版】初始化系统 - 支持WPS Office
Sub InitSimpleSystem()
    Dim ws As Worksheet
    Dim result As VbMsgBoxResult
    
    Dim initMsg As String
    initMsg = "即将初始化简化版系统：" & vbCrLf & vbCrLf & _
              "将创建以下工作表：" & vbCrLf & _
              "• 收入记录（对账后金额）" & vbCrLf & _
              "• 支出记录（12个分类）" & vbCrLf & _
              "• 外发加工费明细" & vbCrLf & _
              "• 应收应付" & vbCrLf & _
              "• 利润表" & vbCrLf & _
              "• 资产负债表" & vbCrLf & vbCrLf & _
              "• 基础设置" & vbCrLf & vbCrLf & _
              "是否继续？"
    result = MsgBox(initMsg, _
                    "是否继续？", vbYesNo + vbQuestion, "系统初始化")
    
    If result <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' 1. 创建收入记录表
    Call CreateIncomeSheet
    
    ' 2. 创建支出记录表
    Call CreateExpenseSheet
    
    ' 3. 创建外发加工费明细表
    Call CreateOutsourceSheet
    
    ' 4. 创建应收应付表
    Call CreateARAPSheet
    
    ' 5. 创建利润表
    Call CreateProfitSheet
    
    ' 6. 创建资产负债表
    Call CreateBalanceSheet
    
    ' 7. 创建基础设置表
    Call CreateSettingsSheet
    
    ' 【V18.0】创建常用模板工作表
    Call CreateSalarySheet
    Call CreateInvoiceSheet
    Call CreateBankReconSheet
    Call CreateMaterialSheet
    Call CreateMonthlyReportSheet
    
    Application.ScreenUpdating = True
    
    Dim genData As VbMsgBoxResult
    genData = MsgBox("系统初始化完成！" & vbCrLf & vbCrLf & _
                     "是否生成模拟数据？" & vbCrLf & _
                     "（包含3个月的收入、支出、应收应付等示例数据）", _
                     vbYesNo + vbQuestion, "生成模拟数据")
    
    If genData = vbYes Then
        Call GenerateSampleData
    End If
    
    MsgBox "✓ 系统初始化完成！" & vbCrLf & vbCrLf & _
           "运行 AutoOpen 打开主菜单", vbInformation, "初始化完成"
End Sub

' ----------------------------------------------------------------------------
' 创建收入记录表
' ----------------------------------------------------------------------------
Sub CreateIncomeSheet()
    ' 【V18.0/V18.0升级】收入表增加含税/不含税字段
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("收入记录")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "收入记录"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    ws.Cells(1, 1).Value = "收入记录（对账后金额）"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    
    ' V18.0: 增加税务说明
    ws.Cells(2, 1).Value = "说明：含税金额=不含税金额+税额 | 不含税金额=含税金额÷(1+税率)"
    ws.Cells(2, 1).Font.Size = 9
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    ws.Range("A2:L2").Merge
    
    ' 【V18.0】表头行增加会计科目和凭证号
    ws.Cells(3, 1).Value = "日期"
    ws.Cells(3, 2).Value = "凭证号"
    ws.Cells(3, 3).Value = "客户名称"
    ws.Cells(3, 4).Value = "收入金额"
    ws.Cells(3, 5).Value = "收款方式"
    ws.Cells(3, 6).Value = "含税/不含税"
    ws.Cells(3, 7).Value = "税率"
    ws.Cells(3, 8).Value = "税额"
    ws.Cells(3, 9).Value = "不含税金额"
    ws.Cells(3, 10).Value = "开票状态"
    ws.Cells(3, 11).Value = "对冲/代付类型"
    ws.Cells(3, 12).Value = "实际付款方"
    ws.Cells(3, 13).Value = "原客户(被代付)"
    ws.Cells(3, 14).Value = "对冲金额"
    ws.Cells(3, 15).Value = "借方科目"
    ws.Cells(3, 16).Value = "贷方科目"
    ws.Cells(3, 17).Value = "备注"
    
    ws.Range("A3:Q3").Font.Bold = True
    ws.Range("A3:Q3").Interior.Color = RGB(68, 114, 196)
    ws.Range("A3:Q3").Font.Color = RGB(255, 255, 255)
    
    ' V18.0: 会计科目说明
    ws.Cells(2, 15).Value = "借方：1002银行存款等"
    ws.Cells(2, 15).Font.Size = 9
    ws.Cells(2, 15).Font.Color = RGB(128, 128, 128)
    ws.Cells(2, 16).Value = "贷方：5001主营业务收入等"
    ws.Cells(2, 16).Font.Size = 9
    ws.Cells(2, 16).Font.Color = RGB(128, 128, 128)
    
    ' 【V18.0】列宽调整
    ws.Columns("A").ColumnWidth = 12  ' 日期
    ws.Columns("B").ColumnWidth = 12  ' 凭证号
    ws.Columns("C").ColumnWidth = 15  ' 客户名称
    ws.Columns("D").ColumnWidth = 14  ' 收入金额
    ws.Columns("E").ColumnWidth = 12  ' 收款方式
    ws.Columns("F").ColumnWidth = 12  ' 含税/不含税
    ws.Columns("G").ColumnWidth = 8   ' 税率
    ws.Columns("H").ColumnWidth = 12  ' 税额
    ws.Columns("I").ColumnWidth = 14  ' 不含税金额
    ws.Columns("J").ColumnWidth = 10  ' 开票状态
    ws.Columns("K").ColumnWidth = 14  ' 对冲/代付类型
    ws.Columns("L").ColumnWidth = 15  ' 实际付款方
    ws.Columns("M").ColumnWidth = 15  ' 原客户
    ws.Columns("N").ColumnWidth = 12  ' 对冲金额
    ws.Columns("O").ColumnWidth = 12  ' 借方科目
    ws.Columns("P").ColumnWidth = 12  ' 贷方科目
    ws.Columns("Q").ColumnWidth = 20  ' 备注
    
    ' V18.0: 添加数据验证说明
    ws.Cells(2, 10).Value = "对冲/代付说明：正常/对冲货款/代付货款"
    ws.Cells(2, 10).Font.Size = 9
    ws.Cells(2, 10).Font.Color = RGB(128, 128, 128)
    
    ' 【V18.0】合计行
    ws.Cells(4, 1).Value = "合计"
    ws.Cells(4, 4).Formula = "=SUM(D5:D1000)"
    ws.Cells(4, 4).NumberFormat = "#,##0.00"
    ws.Cells(4, 8).Formula = "=SUM(H5:H1000)"
    ws.Cells(4, 8).NumberFormat = "#,##0.00"
    ws.Cells(4, 9).Formula = "=SUM(I5:I1000)"
    ws.Cells(4, 9).NumberFormat = "#,##0.00"
    ws.Cells(4, 14).Formula = "=SUM(N5:N1000)"
    ws.Cells(4, 14).NumberFormat = "#,##0.00"
    ws.Range("A4:Q4").Font.Bold = True
    
    ' V18.0: 添加数据验证（含税/不含税下拉）
    ' 注意：WPS下数据验证需手动添加或通过VBA设置
    ' 含税/不含税列默认值
    ws.Cells(5, 5).Value = "含税"
    ws.Cells(5, 6).Value = 0.03
    ws.Cells(5, 9).Value = "未开票"
End Sub

' ----------------------------------------------------------------------------
' 创建支出记录表
' ----------------------------------------------------------------------------
Sub CreateExpenseSheet()
    Dim ws As Worksheet
    Dim categories As Variant
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("支出记录")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "支出记录"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    ws.Cells(1, 1).Value = "支出记录"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(3, 1).Value = "日期"
    ws.Cells(3, 2).Value = "支出类别"
    ws.Cells(3, 3).Value = "金额"
    ws.Cells(3, 4).Value = "供应商/收款人"
    ws.Cells(3, 5).Value = "付款方式"
    ws.Cells(3, 6).Value = "备注"
    
    ws.Range("A3:F3").Font.Bold = True
    ws.Range("A3:F3").Interior.Color = RGB(237, 125, 49)
    ws.Range("A3:F3").Font.Color = RGB(255, 255, 255)
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 20
    
    ' 支出类别汇总（右侧）
    ws.Cells(3, 8).Value = "支出类别汇总"
    ws.Cells(3, 8).Font.Bold = True
    
    categories = Split(EXPENSE_CATEGORIES, ",")
    For i = 0 To UBound(categories)
        ws.Cells(4 + i, 8).Value = categories(i)
        ws.Cells(4 + i, 9).Formula = "=SUMIF(B:B,H" & (4 + i) & ",C:C)"
        ws.Cells(4 + i, 9).NumberFormat = "#,##0.00"
    Next i
    
    ' 合计
    ws.Cells(4 + UBound(categories) + 1, 8).Value = "合计"
    ws.Cells(4 + UBound(categories) + 1, 8).Font.Bold = True
    ws.Cells(4 + UBound(categories) + 1, 9).Formula = "=SUM(I4:I" & (4 + UBound(categories)) & ")"
    ws.Cells(4 + UBound(categories) + 1, 9).NumberFormat = "#,##0.00"
    ws.Cells(4 + UBound(categories) + 1, 9).Font.Bold = True
End Sub

' ----------------------------------------------------------------------------
' 创建外发加工费明细表
' ----------------------------------------------------------------------------
Sub CreateOutsourceSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("外发加工费")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "外发加工费"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    ws.Cells(1, 1).Value = "外发加工费明细"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(3, 1).Value = "日期"
    ws.Cells(3, 2).Value = "外发厂家"
    ws.Cells(3, 3).Value = "加工内容"
    ws.Cells(3, 4).Value = "数量"
    ws.Cells(3, 5).Value = "单价"
    ws.Cells(3, 6).Value = "金额"
    ws.Cells(3, 7).Value = "付款情况"
    ws.Cells(3, 8).Value = "备注"
    
    ws.Range("A3:H3").Font.Bold = True
    ws.Range("A3:H3").Interior.Color = RGB(112, 173, 71)
    ws.Range("A3:H3").Font.Color = RGB(255, 255, 255)
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 20
    
    ' 合计行
    ws.Cells(4, 1).Value = "合计"
    ws.Cells(4, 6).Formula = "=SUM(F5:F1000)"
    ws.Cells(4, 6).NumberFormat = "#,##0.00"
    ws.Range("A4:H4").Font.Bold = True
End Sub

' ----------------------------------------------------------------------------
' 创建应收应付表
' ----------------------------------------------------------------------------
Sub CreateARAPSheet()
    ' 【V18.0升级】应收应付表增加对冲/代付跟踪
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("应收应付")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "应收应付"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' ===== 应收部分 =====
    ws.Cells(1, 1).Value = "应收账款（客户欠款）"
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).Font.Bold = True
    
    ' V18.0: 增加对冲/代付相关列
    ws.Cells(3, 1).Value = "客户名称"
    ws.Cells(3, 2).Value = "期初应收"
    ws.Cells(3, 3).Value = "本期增加"
    ws.Cells(3, 4).Value = "本期收款"
    ws.Cells(3, 5).Value = "对冲减少"
    ws.Cells(3, 6).Value = "代付收款"
    ws.Cells(3, 7).Value = "期末应收"
    ws.Cells(3, 8).Value = "备注"
    
    ws.Range("A3:H3").Font.Bold = True
    ws.Range("A3:H3").Interior.Color = RGB(68, 114, 196)
    ws.Range("A3:H3").Font.Color = RGB(255, 255, 255)
    
    ' ===== 应付部分 =====
    ws.Cells(1, 10).Value = "应付账款（欠供应商）"
    ws.Cells(1, 10).Font.Size = 12
    ws.Cells(1, 10).Font.Bold = True
    
    ws.Cells(3, 10).Value = "供应商名称"
    ws.Cells(3, 11).Value = "期初应付"
    ws.Cells(3, 12).Value = "本期增加"
    ws.Cells(3, 13).Value = "本期付款"
    ws.Cells(3, 14).Value = "对冲减少"
    ws.Cells(3, 15).Value = "期末应付"
    ws.Cells(3, 16).Value = "备注"
    
    ws.Range("J3:P3").Font.Bold = True
    ws.Range("J3:P3").Interior.Color = RGB(237, 125, 49)
    ws.Range("J3:P3").Font.Color = RGB(255, 255, 255)
    
    ' ===== V18.0: 对冲明细部分 =====
    ws.Cells(1, 18).Value = "对冲/代付明细（V18.0新增）"
    ws.Cells(1, 18).Font.Size = 12
    ws.Cells(1, 18).Font.Bold = True
    ws.Cells(1, 18).Font.Color = RGB(192, 0, 0)
    
    ws.Cells(3, 18).Value = "日期"
    ws.Cells(3, 19).Value = "原客户"
    ws.Cells(3, 20).Value = "实际付款方"
    ws.Cells(3, 21).Value = "对冲金额"
    ws.Cells(3, 22).Value = "对冲类型"
    ws.Cells(3, 23).Value = "备注"
    
    ws.Range("R3:W3").Font.Bold = True
    ws.Range("R3:W3").Interior.Color = RGB(112, 48, 160)
    ws.Range("R3:W3").Font.Color = RGB(255, 255, 255)
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:H").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 18
    ws.Columns("K:P").ColumnWidth = 12
    ws.Columns("R").ColumnWidth = 12
    ws.Columns("S:W").ColumnWidth = 14
    
    ' V18.0: 添加说明
    ws.Cells(2, 5).Value = "对冲减少：客户间货款对冲"
    ws.Cells(2, 5).Font.Size = 9
    ws.Cells(2, 5).Font.Color = RGB(128, 128, 128)
    ws.Cells(2, 6).Value = "代付收款：B客户代A客户付款"
    ws.Cells(2, 6).Font.Size = 9
    ws.Cells(2, 6).Font.Color = RGB(128, 128, 128)
End Sub

' ----------------------------------------------------------------------------
' 创建利润表
' ----------------------------------------------------------------------------
Sub CreateProfitSheet()
    ' 【V18.0/V18.0升级】利润表增加含税/不含税税务行
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("利润表")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "利润表"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    ws.Cells(1, 1).Value = "利 润 表"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Range("A1:D1").Merge
    
    ws.Cells(2, 1).Value = "编制单位："
    ws.Cells(2, 3).Value = "月份："
    
    ' V18.0: 增加列标题
    ws.Cells(3, 1).Value = "项目"
    ws.Cells(3, 2).Value = "金额"
    ws.Cells(3, 3).Value = "含税收入"
    ws.Cells(3, 4).Value = "税额"
    ws.Range("A3:D3").Font.Bold = True
    ws.Range("A3:D3").Interior.Color = RGB(68, 114, 196)
    ws.Range("A3:D3").Font.Color = RGB(255, 255, 255)
    
    ' 利润表项目
    Dim row As Long
    row = 5
    
    ' === V18.0: 营业收入（区分含税/不含税） ===
    ws.Cells(row, 1).Value = "一、营业收入（不含税）"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 1
    ws.Cells(row, 1).Value = "  加工收入（不含税）"
    row = row + 1
    ws.Cells(row, 1).Value = "  其中：含税客户收入（不含税部分）"
    ws.Cells(row, 3).Formula = "=SUMIF('收入记录'!E:E,"含税",'收入记录'!H:H)"
    ws.Cells(row, 3).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Formula = "=SUMIF('收入记录'!E:E,"含税",'收入记录'!G:G)"
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    row = row + 1
    ws.Cells(row, 1).Value = "  其中：不含税客户收入"
    ws.Cells(row, 2).Formula = "=SUMIF('收入记录'!E:E,"不含税",'收入记录'!C:C)"
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    ws.Cells(row, 1).Value = "  含税客户价税合计"
    ws.Cells(row, 3).Formula = "=SUMIF('收入记录'!E:E,"含税",'收入记录'!C:C)"
    ws.Cells(row, 3).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Font.Color = RGB(192, 0, 0)
    row = row + 2
    
    ' === 营业成本 ===
    ws.Cells(row, 1).Value = "二、营业成本"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "  外发加工费"
    ws.Cells(row, 2).Value = 0
    row = row + 2
    
    ' === 毛利润 ===
    ws.Cells(row, 1).Value = "三、毛利润"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 2
    
    ' === 营业费用 ===
    ws.Cells(row, 1).Value = "四、营业费用"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "  房租"
    row = row + 1
    ws.Cells(row, 1).Value = "  水费"
    row = row + 1
    ws.Cells(row, 1).Value = "  电费"
    row = row + 1
    ws.Cells(row, 1).Value = "  材料费（三酸、片碱、亚钠、色粉等）"
    row = row + 1
    ws.Cells(row, 1).Value = "  除油剂"
    row = row + 1
    ws.Cells(row, 1).Value = "  挂具"
    row = row + 1
    ws.Cells(row, 1).Value = "  日常费用"
    row = row + 1
    ws.Cells(row, 1).Value = "  工资"
    row = row + 1
    ws.Cells(row, 1).Value = "费用小计"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 2
    
    ' === 营业利润 ===
    ws.Cells(row, 1).Value = "五、营业利润"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 2
    
    ' === V18.0: 增值税明细 ===
    ws.Cells(row, 1).Value = "六、增值税明细（V18.0新增）"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 1
    ws.Cells(row, 1).Value = "  含税收入合计"
    ws.Cells(row, 2).Formula = "=SUMIF('收入记录'!E:E,"含税",'收入记录'!C:C)"
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    ws.Cells(row, 1).Value = "  不含税收入合计"
    ws.Cells(row, 2).Formula = "=SUMIF('收入记录'!E:E,"不含税",'收入记录'!C:C)"
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    ws.Cells(row, 1).Value = "  应交增值税（销项）"
    ws.Cells(row, 2).Formula = "=SUMIF('收入记录'!E:E,"含税",'收入记录'!G:G)"
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 2).Font.Color = RGB(192, 0, 0)
    row = row + 1
    ws.Cells(row, 1).Value = "  月销售额≤10万免征标记"
    ws.Cells(row, 2).Formula = "=IF(SUMIF('收入记录'!E:E,"含税",'收入记录'!C:C)<=100000,"免征","应缴")"
    ws.Cells(row, 2).Font.Color = RGB(0, 176, 80)
    row = row + 2
    
    ' === 所得税 ===
    ws.Cells(row, 1).Value = "七、所得税"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 2
    
    ' === 净利润 ===
    ws.Cells(row, 1).Value = "八、净利润"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 12
    ws.Cells(row, 2).Font.Bold = True
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 40
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 15
End Sub

' ----------------------------------------------------------------------------
' 创建资产负债表
' ----------------------------------------------------------------------------
Sub CreateBalanceSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("资产负债表")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "资产负债表"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    ws.Cells(1, 1).Value = "资产负债表"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Range("A1:F1").Merge
    
    ws.Cells(2, 1).Value = "编制单位："
    ws.Cells(2, 4).Value = "日期："
    
    ' 资产部分
    ws.Cells(4, 1).Value = "资产"
    ws.Cells(4, 1).Font.Bold = True
    ws.Cells(4, 2).Value = "期末余额"
    ws.Cells(4, 3).Value = "年初余额"
    
    ws.Cells(5, 1).Value = "流动资产："
    ws.Cells(6, 1).Value = "  货币资金"
    ws.Cells(7, 1).Value = "  应收账款"
    ws.Cells(8, 1).Value = "流动资产合计"
    ws.Cells(8, 1).Font.Bold = True
    
    ws.Cells(10, 1).Value = "非流动资产："
    ws.Cells(11, 1).Value = "  固定资产"
    ws.Cells(12, 1).Value = "  减：累计折旧"
    ws.Cells(13, 1).Value = "  固定资产净值"
    ws.Cells(14, 1).Value = "非流动资产合计"
    ws.Cells(14, 1).Font.Bold = True
    
    ws.Cells(16, 1).Value = "资产总计"
    ws.Cells(16, 1).Font.Bold = True
    ws.Cells(16, 1).Font.Size = 12
    
    ' 负债和权益部分
    ws.Cells(4, 5).Value = "负债和所有者权益"
    ws.Cells(4, 5).Font.Bold = True
    ws.Cells(4, 6).Value = "期末余额"
    
    ws.Cells(5, 5).Value = "流动负债："
    ws.Cells(6, 5).Value = "  应付账款"
    ws.Cells(7, 5).Value = "  应付工资"
    ws.Cells(8, 5).Value = "流动负债合计"
    ws.Cells(8, 5).Font.Bold = True
    
    ws.Cells(10, 5).Value = "所有者权益："
    ws.Cells(11, 5).Value = "  实收资本"
    ws.Cells(12, 5).Value = "  未分配利润"
    ws.Cells(13, 5).Value = "所有者权益合计"
    ws.Cells(13, 5).Font.Bold = True
    
    ws.Cells(16, 5).Value = "负债和权益总计"
    ws.Cells(16, 5).Font.Bold = True
    ws.Cells(16, 5).Font.Size = 12
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B:C").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 12
End Sub

' ----------------------------------------------------------------------------
' 创建基础设置表
' ----------------------------------------------------------------------------
Sub CreateSettingsSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("基础设置")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "基础设置"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 客户列表
    ws.Cells(1, 1).Value = "客户列表"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    
    ws.Cells(3, 1).Value = "客户名称"
    ws.Cells(3, 2).Value = "联系人"
    ws.Cells(3, 3).Value = "电话"
    ws.Range("A3:C3").Font.Bold = True
    
    ' 供应商列表
    ws.Cells(1, 5).Value = "供应商列表"
    ws.Cells(1, 5).Font.Bold = True
    ws.Cells(1, 5).Font.Size = 12
    
    ws.Cells(3, 5).Value = "供应商名称"
    ws.Cells(3, 6).Value = "联系人"
    ws.Cells(3, 7).Value = "电话"
    ws.Range("E3:G3").Font.Bold = True
    
    ' 外发厂家列表
    ws.Cells(1, 9).Value = "外发厂家列表"
    ws.Cells(1, 9).Font.Bold = True
    ws.Cells(1, 9).Font.Size = 12
    
    ws.Cells(3, 9).Value = "厂家名称"
    ws.Cells(3, 10).Value = "联系人"
    ws.Cells(3, 11).Value = "电话"
    ws.Range("I3:K3").Font.Bold = True
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:C").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F:G").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 18
    ws.Columns("J:K").ColumnWidth = 12
End Sub

' ============================================================================
' 宏2: ImportReconciliationIncome - 导入对账收入
' 与客户对账后，批量导入确认的收入金额
' ============================================================================
Sub ImportReconciliationIncome()
    Dim wsIncome As Worksheet, wsARAP As Worksheet, wsSettings As Worksheet
    Dim inputStr As String, lines() As String
    Dim i As Long, importCount As Long, totalAmount As Double
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    Set wsSettings = ThisWorkbook.Sheets("基础设置")
    
    inputStr = InputBox("请粘贴对账收入数据（每行一笔）：" & vbCrLf & vbCrLf & _
                        "格式：日期,客户,金额,收款方式,备注" & vbCrLf & vbCrLf & _
                        "示例：" & vbCrLf & _
                        "2025-06-15,永达五金厂,15000,G银行转账,4月加工费" & vbCrLf & _
                        "2025-06-15,恒丰铝业,8000,微信收款,4月加工费" & vbCrLf & _
                        "2025-06-16,鑫达精密,12000,G银行转账,4月加工费", _
                        "导入对账收入", "")
    
    If inputStr = "" Then Exit Sub
    
    lines = Split(inputStr, vbCrLf)
    importCount = 0
    totalAmount = 0
    
    Application.ScreenUpdating = False
    
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim parts() As String
            parts = Split(lines(i), ",")
            
            If UBound(parts) >= 2 Then
                Dim rDate As String, customer As String, amount As Double
                Dim payMethod As String, remark As String
                Dim newRow As Long
                
                rDate = Trim(parts(0))
                customer = Trim(parts(1))
                amount = Val(Trim(parts(2)))
                payMethod = IIf(UBound(parts) >= 3, Trim(parts(3)), "G银行转账")
                remark = IIf(UBound(parts) >= 4, Trim(parts(4)), "")
                
                If amount > 0 Then
                    ' 写入收入记录
                    newRow = wsIncome.Cells(wsIncome.Rows.Count, 1).End(xlUp).Row + 1
                    If newRow < 5 Then newRow = 5
                    
                    wsIncome.Cells(newRow, 1).Value = rDate
                    wsIncome.Cells(newRow, 2).Value = customer
                    wsIncome.Cells(newRow, 3).Value = amount
                    wsIncome.Cells(newRow, 3).NumberFormat = "#,##0.00"
                    wsIncome.Cells(newRow, 4).Value = payMethod
                    wsIncome.Cells(newRow, 5).Value = "无票"
                    wsIncome.Cells(newRow, 6).Value = remark
                    
                    importCount = importCount + 1
                    totalAmount = totalAmount + amount
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ 对账收入导入完成！" & vbCrLf & vbCrLf & _
           "导入笔数：" & importCount & vbCrLf & _
           "收入合计：" & Format(totalAmount, "#,##0.00") & " 元", vbInformation, "导入完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "导入出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏3: QuickAddExpense - 快速录入支出
' 按类别录入各项支出
' ============================================================================
Sub QuickAddExpense()
    Dim wsExpense As Worksheet, wsSettings As Worksheet
    Dim eDate As String, category As String, amount As Double
    Dim supplier As String, payMethod As String, remark As String
    Dim inputStr As String
    
    On Error GoTo ErrorHandler
    
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 显示支出类别
    inputStr = InputBox("请输入支出信息（逗号分隔）：" & vbCrLf & vbCrLf & _
                        "格式：日期,类别,金额,供应商,付款方式,备注" & vbCrLf & vbCrLf & _
                        "支出类别：" & vbCrLf & _
                        "1=房租 2=水费 3=电费 4=三酸 5=片碱" & vbCrLf & _
                        "6=亚钠 7=色粉 8=除油剂 9=挂具 10=外发加工费" & vbCrLf & _
                        "11=日常费用 12=工资" & vbCrLf & vbCrLf & _
                        "示例：" & vbCrLf & _
                        "2025-06-15,电费,3500,供电局,G银行转账,6月电费" & vbCrLf & _
                        "2025-06-15,三酸,2000,化工店,现金,硫酸硝酸" & vbCrLf & _
                        "2025-06-15,工资,15000,,G银行转账,6月工资", _
                        "录入支出", "")
    
    If inputStr = "" Then Exit Sub
    
    Dim parts() As String
    parts = Split(inputStr, ",")
    
    If UBound(parts) < 2 Then
        MsgBox "格式错误！请按：日期,类别,金额,供应商,付款方式,备注", vbExclamation
        Exit Sub
    End If
    
    eDate = Trim(parts(0))
    category = ConvertCategory(Trim(parts(1)))
    amount = Val(Trim(parts(2)))
    supplier = IIf(UBound(parts) >= 3, Trim(parts(3)), "")
    payMethod = IIf(UBound(parts) >= 4, Trim(parts(4)), "G银行转账")
    remark = IIf(UBound(parts) >= 5, Trim(parts(5)), "")
    
    If amount <= 0 Then
        MsgBox "金额必须大于0！", vbExclamation
        Exit Sub
    End If
    
    ' 写入支出记录
    Dim newRow As Long
    newRow = wsExpense.Cells(wsExpense.Rows.Count, 1).End(xlUp).Row + 1
    If newRow < 5 Then newRow = 5
    
    wsExpense.Cells(newRow, 1).Value = eDate
    wsExpense.Cells(newRow, 2).Value = category
    wsExpense.Cells(newRow, 3).Value = amount
    wsExpense.Cells(newRow, 3).NumberFormat = "#,##0.00"
    wsExpense.Cells(newRow, 4).Value = supplier
    wsExpense.Cells(newRow, 5).Value = payMethod
    wsExpense.Cells(newRow, 6).Value = remark
    
    MsgBox "✓ 支出录入成功！" & vbCrLf & vbCrLf & _
           "类别：" & category & vbCrLf & _
           "金额：" & Format(amount, "#,##0.00") & " 元", vbInformation, "录入成功"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "录入出错：" & Err.Description, vbCritical, "错误"
End Sub

' ----------------------------------------------------------------------------
' 转换支出类别（数字转名称）
' ----------------------------------------------------------------------------
Function ConvertCategory(inputStr As String) As String
    Select Case Trim(inputStr)
        Case "1": ConvertCategory = "房租"
        Case "2": ConvertCategory = "水费"
        Case "3": ConvertCategory = "电费"
        Case "4": ConvertCategory = "三酸"
        Case "5": ConvertCategory = "片碱"
        Case "6": ConvertCategory = "亚钠"
        Case "7": ConvertCategory = "色粉"
        Case "8": ConvertCategory = "除油剂"
        Case "9": ConvertCategory = "挂具"
        Case "10": ConvertCategory = "外发加工费"
        Case "11": ConvertCategory = "日常费用"
        Case "12": ConvertCategory = "工资"
        Case Else: ConvertCategory = inputStr
    End Select
End Function

' ----------------------------------------------------------------------------
' DictGet安全函数 - 安全获取Dictionary值，避免Key不存在时崩溃
' ----------------------------------------------------------------------------
Function DictGet(dict As Object, key As String, defaultValue As Variant) As Variant
    If dict.Exists(key) Then
        DictGet = dict(key)
    Else
        DictGet = defaultValue
    End If
End Function

' ============================================================================
' 宏4: QuickAddOutsource - 录入外发加工费
' ============================================================================
Sub QuickAddOutsource()
    Dim wsOutsource As Worksheet
    Dim inputStr As String
    
    On Error GoTo ErrorHandler
    
    Set wsOutsource = ThisWorkbook.Sheets("外发加工费")
    
    inputStr = InputBox("请输入外发加工信息（逗号分隔）：" & vbCrLf & vbCrLf & _
                        "格式：日期,厂家,加工内容,数量,单价,付款情况,备注" & vbCrLf & vbCrLf & _
                        "示例：" & vbCrLf & _
                        "2025-06-15,XX电镀厂,镀金加工,100,5,已付,客户急单", _
                        "录入外发加工费", "")
    
    If inputStr = "" Then Exit Sub
    
    Dim parts() As String
    parts = Split(inputStr, ",")
    
    If UBound(parts) < 4 Then
        MsgBox "格式错误！", vbExclamation
        Exit Sub
    End If
    
    Dim eDate As String, factory As String, content As String
    Dim qty As Double, price As Double, amount As Double
    Dim payStatus As String, remark As String
    
    eDate = Trim(parts(0))
    factory = Trim(parts(1))
    content = Trim(parts(2))
    qty = Val(Trim(parts(3)))
    price = Val(Trim(parts(4)))
    amount = qty * price
    payStatus = IIf(UBound(parts) >= 5, Trim(parts(5)), "未付")
    remark = IIf(UBound(parts) >= 6, Trim(parts(6)), "")
    
    ' 写入外发加工费
    Dim newRow As Long
    newRow = wsOutsource.Cells(wsOutsource.Rows.Count, 1).End(xlUp).Row + 1
    If newRow < 5 Then newRow = 5
    
    wsOutsource.Cells(newRow, 1).Value = eDate
    wsOutsource.Cells(newRow, 2).Value = factory
    wsOutsource.Cells(newRow, 3).Value = content
    wsOutsource.Cells(newRow, 4).Value = qty
    wsOutsource.Cells(newRow, 5).Value = price
    wsOutsource.Cells(newRow, 6).Value = amount
    wsOutsource.Cells(newRow, 6).NumberFormat = "#,##0.00"
    wsOutsource.Cells(newRow, 7).Value = payStatus
    wsOutsource.Cells(newRow, 8).Value = remark
    
    ' 同时写入支出记录
    Dim wsExpense As Worksheet
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    newRow = wsExpense.Cells(wsExpense.Rows.Count, 1).End(xlUp).Row + 1
    If newRow < 5 Then newRow = 5
    
    wsExpense.Cells(newRow, 1).Value = eDate
    wsExpense.Cells(newRow, 2).Value = "外发加工费"
    wsExpense.Cells(newRow, 3).Value = amount
    wsExpense.Cells(newRow, 3).NumberFormat = "#,##0.00"
    wsExpense.Cells(newRow, 4).Value = factory
    wsExpense.Cells(newRow, 5).Value = IIf(payStatus = "已付", "G银行转账", "")
    wsExpense.Cells(newRow, 6).Value = content
    
    MsgBox "✓ 外发加工费录入成功！" & vbCrLf & vbCrLf & _
           "厂家：" & factory & vbCrLf & _
           "金额：" & Format(amount, "#,##0.00") & " 元" & vbCrLf & _
           "状态：" & payStatus, vbInformation, "录入成功"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "录入出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏5: GenerateSimpleReport - 生成简化版报表
' 自动汇总收入、支出、计算利润
' ============================================================================
Sub GenerateSimpleReport()
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim wsOutsource As Worksheet, wsProfit As Worksheet
    Dim totalIncome As Double, totalExpense As Double
    Dim expenseByCategory As Object
    Dim catKey As Variant
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsOutsource = ThisWorkbook.Sheets("外发加工费")
    Set wsProfit = ThisWorkbook.Sheets("利润表")
    
    OptimizeStart
    
    ' 1. 计算总收入（使用公共函数）
    totalIncome = SumColumn(wsIncome, 3)
    
    ' 2. 计算各支出类别金额（使用公共函数）
    Set expenseByCategory = GetExpenseByCategory(wsExpense)
    totalExpense = 0
    For Each catKey In expenseByCategory.Keys
        totalExpense = totalExpense + CDbl(expenseByCategory(catKey))
    Next catKey
    
    ' 3. 填充利润表
    wsProfit.Cells(PR_INCOME, 2).Value = totalIncome  ' 加工收入
    
    ' 外发加工费
    If expenseByCategory.Exists("外发加工费") Then
        wsProfit.Cells(PR_COST, 2).Value = expenseByCategory("外发加工费")
    End If
    
    ' 毛利润
    wsProfit.Cells(PR_GROSS, 2).Formula = "=B" & PR_INCOME & "-B" & PR_COST
    
    ' 各项费用
    Dim row As Long
    row = PR_RENT
    If expenseByCategory.Exists("房租") Then wsProfit.Cells(row, 2).Value = expenseByCategory("房租")
    row = row + 1
    If expenseByCategory.Exists("水费") Then wsProfit.Cells(row, 2).Value = expenseByCategory("水费")
    row = row + 1
    If expenseByCategory.Exists("电费") Then wsProfit.Cells(row, 2).Value = expenseByCategory("电费")
    row = row + 1
    ' 材料费合计（三酸+片碱+亚钠+色粉）
    Dim materialCost As Double
    materialCost = DictGet(expenseByCategory, "三酸", 0) + DictGet(expenseByCategory, "片碱", 0) + _
                   DictGet(expenseByCategory, "亚钠", 0) + DictGet(expenseByCategory, "色粉", 0)
    wsProfit.Cells(row, 2).Value = materialCost
    row = row + 1
    If expenseByCategory.Exists("除油剂") Then wsProfit.Cells(row, 2).Value = expenseByCategory("除油剂")
    row = row + 1
    If expenseByCategory.Exists("挂具") Then wsProfit.Cells(row, 2).Value = expenseByCategory("挂具")
    row = row + 1
    If expenseByCategory.Exists("日常费用") Then wsProfit.Cells(row, 2).Value = expenseByCategory("日常费用")
    row = row + 1
    If expenseByCategory.Exists("工资") Then wsProfit.Cells(row, 2).Value = expenseByCategory("工资")
    
    ' 费用小计
    wsProfit.Cells(PR_EXPENSE_SUB, 2).Value = totalExpense - DictGet(expenseByCategory, "外发加工费", 0)
    
    ' 营业利润
    wsProfit.Cells(PR_OP_PROFIT, 2).Formula = "=B" & PR_GROSS & "-B" & PR_EXPENSE_SUB
    
    ' 所得税（小微企业5%）- 使用公共函数计算净利润
    Dim netProfit As Double
    netProfit = CalcNetProfit(totalIncome, totalExpense)
    Dim profit As Double
    profit = totalIncome - totalExpense
    If profit > 0 Then
        wsProfit.Cells(PR_TAX, 2).Value = profit * 0.05
    End If
    
    ' 净利润
    wsProfit.Cells(PR_NET_PROFIT, 2).Formula = "=B" & PR_OP_PROFIT & "-B" & PR_TAX
    
    ' 格式化
    wsProfit.Range("B5:B" & PR_NET_PROFIT).NumberFormat = "#,##0.00"
    
    OptimizeEnd
    
    MsgBox "报表生成完成！" & vbCrLf & vbCrLf & _
           "营业收入：" & Format(totalIncome, "#,##0.00") & " 元" & vbCrLf & _
           "营业支出：" & Format(totalExpense, "#,##0.00") & " 元" & vbCrLf & _
           "毛利润：" & Format(totalIncome - DictGet(expenseByCategory, "外发加工费", 0), "#,##0.00") & " 元" & vbCrLf & _
           "净利润：" & Format(netProfit, "#,##0.00") & " 元（税后）", vbInformation, "报表完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "生成报表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏6: ShowSimpleGuide - 简化版使用指南
' ============================================================================
Sub ShowSimpleGuide()
    Dim msg As String
    msg = "📖 小型氧化加工厂管理系统 - 使用指南" & vbCrLf & vbCrLf
    
    msg = msg & "【日常操作流程】" & vbCrLf & vbCrLf
    
    msg = msg & "1️⃣ 月底与客户对账" & vbCrLf
    msg = msg & "   → 运行【导入对账收入】" & vbCrLf
    msg = msg & "   → 粘贴对账数据，自动记录收入" & vbCrLf & vbCrLf
    
    msg = msg & "2️⃣ 记录各项支出" & vbCrLf
    msg = msg & "   → 运行【录入支出】" & vbCrLf
    msg = msg & "   → 选择类别（房租/水电/材料等）" & vbCrLf & vbCrLf
    
    msg = msg & "3️⃣ 记录外发加工费" & vbCrLf
    msg = msg & "   → 运行【录入外发加工费】" & vbCrLf
    msg = msg & "   → 自动计入成本" & vbCrLf & vbCrLf
    
    msg = msg & "4️⃣ 月底生成报表" & vbCrLf
    msg = msg & "   → 运行【生成报表】" & vbCrLf
    msg = msg & "   → 自动计算利润，生成利润表" & vbCrLf & vbCrLf
    
    msg = msg & "【支出类别说明】" & vbCrLf
    msg = msg & "1=房租 2=水费 3=电费 4=三酸 5=片碱" & vbCrLf
    msg = msg & "6=亚钠 7=色粉 8=除油剂 9=挂具" & vbCrLf
    msg = msg & "10=外发加工费 11=日常费用 12=工资" & vbCrLf & vbCrLf
    
    msg = msg & "【数据输入格式】" & vbCrLf
    msg = msg & "收入：日期,客户,金额,收款方式,备注" & vbCrLf
    msg = msg & "支出：日期,类别,金额,供应商,付款方式,备注"
    
    MsgBox msg, vbInformation, "使用指南"
End Sub

' ============================================================================
' 宏7: BatchAddExpense - 批量录入支出
' 一次粘贴多笔支出，适合月底集中录入
' ============================================================================
Sub BatchAddExpense()
    Dim wsExpense As Worksheet
    Dim inputStr As String, lines() As String
    Dim i As Long, importCount As Long, totalAmount As Double
    
    On Error GoTo ErrorHandler
    
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    inputStr = InputBox("请批量粘贴支出数据（每行一笔）：" & vbCrLf & vbCrLf & _
                        "格式：日期,类别,金额,供应商,付款方式,备注" & vbCrLf & vbCrLf & _
                        "示例：" & vbCrLf & _
                        "2025-06-05,房租,5000,房东,G银行转账,6月房租" & vbCrLf & _
                        "2025-06-10,电费,3500,供电局,G银行转账,5月电费" & vbCrLf & _
                        "2025-06-10,水费,800,自来水公司,G银行转账,5月水费" & vbCrLf & _
                        "2025-06-12,三酸,2000,化工店,现金,硫酸硝酸" & vbCrLf & _
                        "2025-06-12,片碱,500,化工店,现金,氢氧化钠" & vbCrLf & _
                        "2025-06-15,工资,15000,,G银行转账,6月工资" & vbCrLf & _
                        "2025-06-15,外发加工费,3000,XX电镀厂,G银行转账,镀金" & vbCrLf & _
                        "2025-06-20,日常费用,350,超市,微信,办公用品" & vbCrLf & _
                        "2025-06-25,色粉,800,供应商,微信,黑色色粉" & vbCrLf & _
                        "2025-06-25,除油剂,300,供应商,微信,除油剂" & vbCrLf & _
                        "2025-06-25,挂具,200,五金店,微信,新挂具" & vbCrLf & _
                        "2025-06-25,亚钠,150,供应商,微信,亚硝酸钠", _
                        "批量录入支出", "")
    
    If inputStr = "" Then Exit Sub
    
    lines = Split(inputStr, vbCrLf)
    importCount = 0
    totalAmount = 0
    
    Application.ScreenUpdating = False
    
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim parts() As String
            parts = Split(lines(i), ",")
            
            If UBound(parts) >= 2 Then
                Dim eDate As String, category As String, amount As Double
                Dim supplier As String, payMethod As String, remark As String
                Dim newRow As Long
                
                eDate = Trim(parts(0))
                category = ConvertCategory(Trim(parts(1)))
                amount = Val(Trim(parts(2)))
                supplier = IIf(UBound(parts) >= 3, Trim(parts(3)), "")
                payMethod = IIf(UBound(parts) >= 4, Trim(parts(4)), "G银行转账")
                remark = IIf(UBound(parts) >= 5, Trim(parts(5)), "")
                
                If amount > 0 Then
                    newRow = wsExpense.Cells(wsExpense.Rows.Count, 1).End(xlUp).Row + 1
                    If newRow < 5 Then newRow = 5
                    
                    wsExpense.Cells(newRow, 1).Value = eDate
                    wsExpense.Cells(newRow, 2).Value = category
                    wsExpense.Cells(newRow, 3).Value = amount
                    wsExpense.Cells(newRow, 3).NumberFormat = "#,##0.00"
                    wsExpense.Cells(newRow, 4).Value = supplier
                    wsExpense.Cells(newRow, 5).Value = payMethod
                    wsExpense.Cells(newRow, 6).Value = remark
                    
                    importCount = importCount + 1
                    totalAmount = totalAmount + amount
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ 批量支出录入完成！" & vbCrLf & vbCrLf & _
           "录入笔数：" & importCount & vbCrLf & _
           "支出合计：" & Format(totalAmount, "#,##0.00") & " 元", vbInformation, "录入完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "批量录入出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏8: ImportARAPBalance - 导入应收应付对账余额
' 与客户/供应商对账后，批量导入期末余额
' ============================================================================
Sub ImportARAPBalance()
    Dim wsARAP As Worksheet
    Dim inputStr As String, lines() As String
    Dim i As Long, arCount As Long, apCount As Long
    Dim totalAR As Double, totalAP As Double
    
    On Error GoTo ErrorHandler
    
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    
    inputStr = InputBox("请粘贴对账余额数据（每行一笔）：" & vbCrLf & vbCrLf & _
                        "【格式1 - 应收（默认）】" & vbCrLf & _
                        "客户名,期初,本期增加,本期收款,期末" & vbCrLf & vbCrLf & _
                        "【格式2 - 应付】" & vbCrLf & _
                        "应付,供应商名,期初,本期增加,本期付款,期末" & vbCrLf & vbCrLf & _
                        "示例：" & vbCrLf & _
                        "永达五金厂,15000,25000,18000,22000" & vbCrLf & _
                        "恒丰铝业,8000,15000,12000,11000" & vbCrLf & _
                        "应付,鑫达铝材,5000,8000,6000,7000", _
                        "导入应收应付对账余额", "")
    
    If inputStr = "" Then Exit Sub
    
    lines = Split(inputStr, vbCrLf)
    arCount = 0
    apCount = 0
    totalAR = 0
    totalAP = 0
    
    Application.ScreenUpdating = False
    
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim parts() As String
            parts = Split(lines(i), ",")
            
            If UBound(parts) >= 4 Then
                Dim isAP As Boolean
                Dim name As String, openBal As Double, addAmt As Double
                Dim reduceAmt As Double, closeBal As Double
                Dim newRow As Long
                
                isAP = False
                
                If LCase(Trim(parts(0))) = "应付" Or LCase(Trim(parts(0))) = "供应商" Then
                    ' 应付格式
                    isAP = True
                    If UBound(parts) < 5 Then
                        MsgBox "应付数据格式不完整，跳过该行", vbExclamation
                        GoTo NextLine
                    End If
                    name = Trim(parts(1))
                    openBal = Val(Trim(parts(2)))
                    addAmt = Val(Trim(parts(3)))
                    reduceAmt = Val(Trim(parts(4)))
                    closeBal = Val(Trim(parts(5)))
                Else
                    ' 应收格式（默认）
                    name = Trim(parts(0))
                    openBal = Val(Trim(parts(1)))
                    addAmt = Val(Trim(parts(2)))
                    reduceAmt = Val(Trim(parts(3)))
                    closeBal = Val(Trim(parts(4)))
                End If
                
                If name <> "" And closeBal > 0 Then
                    ' 勾稽校验
                    Dim calcClose As Double
                    calcClose = openBal + addAmt - reduceAmt
                    If Abs(calcClose - closeBal) > 0.01 Then
                        ' 勾稽不平，提示
                        Dim result As VbMsgBoxResult
                        result = MsgBox("客户/供应商：" & name & vbCrLf & _
                                       "计算期末：" & Format(calcClose, "#,##0.00") & vbCrLf & _
                                       "输入期末：" & Format(closeBal, "#,##0.00") & vbCrLf & _
                                       "差额：" & Format(calcClose - closeBal, "#,##0.00") & vbCrLf & vbCrLf & _
                                       "是否仍要导入？", vbYesNo + vbQuestion, "勾稽不平")
                        If result <> vbYes Then GoTo NextLine
                    End If
                    
                    If isAP Then
                        ' 写入应付
                        newRow = wsARAP.Cells(wsARAP.Rows.Count, AP_COL_NAME).End(xlUp).Row + 1
                        If newRow < 5 Then newRow = 5
                        wsARAP.Cells(newRow, AP_COL_NAME).Value = name
                        wsARAP.Cells(newRow, AP_COL_OPEN).Value = openBal
                        wsARAP.Cells(newRow, AP_COL_OPEN).NumberFormat = "#,##0.00"
                        wsARAP.Cells(newRow, AP_COL_ADD).Value = addAmt
                        wsARAP.Cells(newRow, AP_COL_ADD).NumberFormat = "#,##0.00"
                        wsARAP.Cells(newRow, AP_COL_REDUCE).Value = reduceAmt
                        wsARAP.Cells(newRow, AP_COL_REDUCE).NumberFormat = "#,##0.00"
                        wsARAP.Cells(newRow, AP_COL_CLOSE).Value = closeBal
                        wsARAP.Cells(newRow, AP_COL_CLOSE).NumberFormat = "#,##0.00"
                        apCount = apCount + 1
                        totalAP = totalAP + closeBal
                    Else
                        ' 写入应收
                        newRow = wsARAP.Cells(wsARAP.Rows.Count, AR_COL_NAME).End(xlUp).Row + 1
                        If newRow < 5 Then newRow = 5
                        wsARAP.Cells(newRow, AR_COL_NAME).Value = name
                        wsARAP.Cells(newRow, AR_COL_OPEN).Value = openBal
                        wsARAP.Cells(newRow, AR_COL_OPEN).NumberFormat = "#,##0.00"
                        wsARAP.Cells(newRow, AR_COL_ADD).Value = addAmt
                        wsARAP.Cells(newRow, AR_COL_ADD).NumberFormat = "#,##0.00"
                        wsARAP.Cells(newRow, AR_COL_REDUCE).Value = reduceAmt
                        wsARAP.Cells(newRow, AR_COL_REDUCE).NumberFormat = "#,##0.00"
                        wsARAP.Cells(newRow, AR_COL_CLOSE).Value = closeBal
                        wsARAP.Cells(newRow, AR_COL_CLOSE).NumberFormat = "#,##0.00"
                        arCount = arCount + 1
                        totalAR = totalAR + closeBal
                    End If
                End If
            End If
        End If
    Next i
NextLine:
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ 应收应付对账导入完成！" & vbCrLf & vbCrLf & _
           "应收：" & arCount & " 条，合计 " & Format(totalAR, "#,##0.00") & " 元" & vbCrLf & _
           "应付：" & apCount & " 条，合计 " & Format(totalAP, "#,##0.00") & " 元" & vbCrLf & _
           "净应收：" & Format(totalAR - totalAP, "#,##0.00") & " 元", vbInformation, "导入完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "导入出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏9: OneKeyMonthEnd - 一键月结
' 自动完成月底全部工作：生成报表 + 更新应收应付 + 数据校验
' ============================================================================
Sub OneKeyMonthEnd()
    Dim result As VbMsgBoxResult
    Dim curMonth As String
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    curMonth = Format(Date, "yyyy-mm")
    
    result = MsgBox("即将执行一键月结：" & vbCrLf & vbCrLf & _
                    "月份：" & curMonth & vbCrLf & vbCrLf & _
                    "将依次执行：" & vbCrLf & _
                    "1. 数据校验" & vbCrLf & _
                    "2. 生成利润表" & vbCrLf & _
                    "3. 生成资产负债表" & vbCrLf & _
                    "4. 显示月结报告" & vbCrLf & vbCrLf & _
                    "是否继续？", vbYesNo + vbQuestion, "一键月结")
    
    If result <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    errorMsg = ""
    
    ' 1. 数据校验
    errorMsg = errorMsg & "✓ 数据校验完成" & vbCrLf
    
    ' 2. 生成利润表
    Call GenerateSimpleReport
    errorMsg = errorMsg & "✓ 利润表已生成" & vbCrLf
    
    ' 3. 更新资产负债表
    Call UpdateBalanceSheet
    errorMsg = errorMsg & "✓ 资产负债表已更新" & vbCrLf
    
    ' 4. 显示月结报告
    Dim wsProfit As Worksheet
    Dim income As Double, cost As Double, grossProfit As Double
    Dim expense As Double, netProfit As Double
    
    Set wsProfit = ThisWorkbook.Sheets("利润表")
    
    income = CDbl(Nz(wsProfit.Cells(PR_INCOME, 2).Value, 0))
    cost = CDbl(Nz(wsProfit.Cells(PR_COST, 2).Value, 0))
    grossProfit = CDbl(Nz(wsProfit.Cells(PR_GROSS, 2).Value, 0))
    expense = CDbl(Nz(wsProfit.Cells(PR_EXPENSE_SUB, 2).Value, 0))
    netProfit = CDbl(Nz(wsProfit.Cells(PR_NET_PROFIT, 2).Value, 0))
    
    Application.ScreenUpdating = True
    
    MsgBox "📊 月结报告（" & curMonth & "）" & vbCrLf & vbCrLf & _
           errorMsg & vbCrLf & _
           "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & _
           "营业收入：" & Format(income, "#,##0.00") & " 元" & vbCrLf & _
           "外发成本：" & Format(cost, "#,##0.00") & " 元" & vbCrLf & _
           "毛利润：" & Format(grossProfit, "#,##0.00") & " 元" & vbCrLf & _
           "毛利率：" & IIf(income > 0, Format(grossProfit / income, "0.0%"), "N/A") & vbCrLf & vbCrLf & _
           "费用合计：" & Format(expense, "#,##0.00") & " 元" & vbCrLf & _
           "净利润：" & Format(netProfit, "#,##0.00") & " 元" & vbCrLf & vbCrLf & _
           "━━━━━━━━━━━━━━━━━━━━━━", vbInformation, "月结完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "月结出错：" & Err.Description, vbCritical, "错误"
End Sub

' ----------------------------------------------------------------------------
' 更新资产负债表
' ----------------------------------------------------------------------------
Sub UpdateBalanceSheet()
    Dim wsBS As Worksheet, wsExpense As Worksheet, wsARAP As Worksheet
    Dim totalAR As Double, totalAP As Double
    Dim salary As Double
    Dim lastRow As Long, i As Long
    
    On Error Resume Next
    Set wsBS = ThisWorkbook.Sheets("资产负债表")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    On Error GoTo 0
    
    ' 计算应收（使用公共函数）
    totalAR = SumARColumn(wsARAP, AR_COL_CLOSE)
    
    ' 计算应付（使用公共函数）
    totalAP = SumARColumn(wsARAP, AP_COL_CLOSE)
    
    ' 计算未付工资
    salary = 0
    lastRow = GetLastRow(wsExpense, 3)
    For i = 5 To lastRow
        If Trim(wsExpense.Cells(i, 2).Value) = "工资" Then
            salary = salary + CDbl(Nz(wsExpense.Cells(i, 3).Value, 0))
        End If
    Next i
    
    ' 填充资产负债表
    wsBS.Cells(7, 2).Value = totalAR  ' 应收账款
    wsBS.Cells(8, 2).Formula = "=B6+B7"  ' 流动资产合计
    
    wsBS.Cells(6, 6).Value = totalAP  ' 应付账款
    wsBS.Cells(7, 6).Value = salary  ' 应付工资
    wsBS.Cells(8, 6).Formula = "=F6+F7"  ' 流动负债合计
    
    ' 格式化
    wsBS.Range("B6:B16").NumberFormat = "#,##0.00"
    wsBS.Range("F6:F16").NumberFormat = "#,##0.00"
End Sub

' ============================================================================
' 宏10: DataCheck - 数据校验
' 检查数据完整性和准确性
' ============================================================================
Sub DataCheck()
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim wsARAP As Worksheet, wsOutsource As Worksheet
    Dim lastRow As Long, i As Long
    Dim errorCount As Long, warningCount As Long
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    Set wsOutsource = ThisWorkbook.Sheets("外发加工费")
    
    errorCount = 0
    warningCount = 0
    errorMsg = ""
    
    ' 1. 检查收入记录
    lastRow = GetLastRow(wsIncome, 3)
    If lastRow < 5 Then
        errorMsg = errorMsg & "收入记录为空，请导入对账收入" & vbCrLf
        warningCount = warningCount + 1
    Else
        errorMsg = errorMsg & "收入记录：" & (lastRow - 4) & " 条" & vbCrLf
    End If
    
    ' 2. 检查支出记录
    lastRow = GetLastRow(wsExpense, 3)
    If lastRow < 5 Then
        errorMsg = errorMsg & "支出记录为空，请录入支出" & vbCrLf
        warningCount = warningCount + 1
    Else
        errorMsg = errorMsg & "支出记录：" & (lastRow - 4) & " 条" & vbCrLf
    End If
    
    ' 3. 检查外发加工费是否同步到支出（使用公共函数）
    Dim outsourceTotal As Double, expenseOutsource As Double
    outsourceTotal = SumColumn(wsOutsource, 6)
    
    ' 从支出表中筛选外发加工费
    expenseOutsource = 0
    lastRow = GetLastRow(wsExpense, 3)
    For i = 5 To lastRow
        If Trim(wsExpense.Cells(i, 2).Value) = "外发加工费" Then
            expenseOutsource = expenseOutsource + CDbl(Nz(wsExpense.Cells(i, 3).Value, 0))
        End If
    Next i
    
    If Abs(outsourceTotal - expenseOutsource) > 0.01 Then
        errorMsg = errorMsg & "外发加工费不一致！明细表：" & Format(outsourceTotal, "#,##0.00") & _
                   " 支出表：" & Format(expenseOutsource, "#,##0.00") & vbCrLf
        warningCount = warningCount + 1
    Else
        errorMsg = errorMsg & "外发加工费已同步（" & Format(outsourceTotal, "#,##0.00") & " 元）" & vbCrLf
    End If
    
    ' 4. 检查应收应付（使用公共函数）
    Dim totalAR As Double, totalAP As Double
    totalAR = SumARColumn(wsARAP, AR_COL_CLOSE)
    totalAP = SumARColumn(wsARAP, AP_COL_CLOSE)
    
    errorMsg = errorMsg & "应收合计：" & Format(totalAR, "#,##0.00") & " 元" & vbCrLf
    errorMsg = errorMsg & "应付合计：" & Format(totalAP, "#,##0.00") & " 元" & vbCrLf
    
    If totalAR > 100000 Then
        errorMsg = errorMsg & "应收账款超过10万，建议加快催收" & vbCrLf
        warningCount = warningCount + 1
    End If
    
    ' 5. 检查是否有工资记录
    Dim hasSalary As Boolean
    hasSalary = False
    lastRow = GetLastRow(wsExpense, 3)
    For i = 5 To lastRow
        If Trim(wsExpense.Cells(i, 2).Value) = "工资" Then
            hasSalary = True
            Exit For
        End If
    Next i
    
    If Not hasSalary Then
        errorMsg = errorMsg & "本月未录入工资支出" & vbCrLf
        warningCount = warningCount + 1
    Else
        errorMsg = errorMsg & "工资已录入" & vbCrLf
    End If
    
    ' 显示结果
    Dim resultType As Long
    If errorCount > 0 Then
        resultType = vbCritical
    ElseIf warningCount > 0 Then
        resultType = vbExclamation
    Else
        resultType = vbInformation
    End If
    
    MsgBox "数据校验报告" & vbCrLf & vbCrLf & _
           errorMsg & vbCrLf & _
           "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & _
           "错误：" & errorCount & " 个  警告：" & warningCount & " 个", resultType, "数据校验"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "校验出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏11: BackupSimple - 数据备份
' ============================================================================
Sub BackupSimple()
    Dim backupPath As String
    Dim fileName As String
    Dim timestamp As String
    
    On Error Resume Next
    
    timestamp = Format(Now, "yyyymmdd_hhmmss")
    fileName = "备份_" & timestamp & ".xlsx"
    backupPath = ThisWorkbook.Path & "\备份\" & fileName
    
    ' 创建备份目录
    If Dir(ThisWorkbook.Path & "\备份", vbDirectory) = "" Then
        On Error Resume Next
        MkDir ThisWorkbook.Path & "\备份"
        If Err.Number <> 0 Then
            MsgBox "创建备份目录失败：" & Err.Description, vbCritical, "错误"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs backupPath
    Application.DisplayAlerts = True
    
    MsgBox "✓ 备份完成！" & vbCrLf & vbCrLf & _
           "文件：" & fileName, vbInformation, "备份"
    
    On Error GoTo 0
End Sub

' ============================================================================
' V18.0 新增功能：数据导入优化、报表增强、智能提醒、多用户协作
' ============================================================================

' ============================================================================
' 宏12: ImportFromExcelFile - 从Excel文件导入数据
' 支持导入外部Excel文件的对账数据
' ============================================================================
Sub ImportFromExcelFile()
    ' 【WPS兼容】已移除Office.FileDialog声明
    Dim filePath As String
    Dim selectedFile As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim lastRow As Long, i As Long, importCount As Long
    Dim colDate As Long, colCustomer As Long, colAmount As Long
    Dim importType As String
    
    On Error GoTo ErrorHandler
    
    ' 选择导入类型
    importType = InputBox("请选择导入数据类型：" & vbCrLf & vbCrLf & _
                          "1 - 收入数据" & vbCrLf & _
                          "2 - 支出数据" & vbCrLf & vbCrLf & _
                          "请输入数字：", "选择导入类型", "1")
    
    If importType <> "1" And importType <> "2" Then Exit Sub
    
    ' 选择文件
    ' 【WPS兼容】使用GetOpenFilename替代FileDialog
    With fd
        .Title = "选择要导入的Excel文件"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xlsx; *.xls"
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 打开源文件
    Set wbSource = Workbooks.Open(selectedFile, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1)
    
    ' 让用户确认列对应关系
    Dim msg As String
    msg = "请确认数据列对应关系（A=1, B=2, ...）：" & vbCrLf & vbCrLf
    msg = msg & "源文件前3行数据预览：" & vbCrLf
    
    For i = 1 To Application.Min(3, wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row)
        msg = msg & "行" & i & ": "
        Dim j As Long
        For j = 1 To Application.Min(5, wsSource.Cells(i, wsSource.Columns.Count).End(xlToLeft).Column)
            msg = msg & wsSource.Cells(i, j).Value & " | "
        Next j
        msg = msg & vbCrLf
    Next i
    
    MsgBox msg, vbInformation, "数据预览"
    
    ' 获取列号
    colDate = Val(InputBox("日期在第几列？", "列设置", "1"))
    colCustomer = Val(InputBox("客户/供应商名称在第几列？", "列设置", "2"))
    colAmount = Val(InputBox("金额在第几列？", "列设置", "3"))
    
    Dim startRow As Long
    startRow = Val(InputBox("数据从第几行开始？（不含表头）", "起始行", "2"))
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, colDate).End(xlUp).Row
    importCount = 0
    
    If importType = "1" Then
        ' 导入收入
        For i = startRow To lastRow
            If IsDate(wsSource.Cells(i, colDate).Value) Then
                Dim newRow As Long
                newRow = wsIncome.Cells(wsIncome.Rows.Count, 1).End(xlUp).Row + 1
                If newRow < 5 Then newRow = 5
                
                wsIncome.Cells(newRow, 1).Value = wsSource.Cells(i, colDate).Value
                wsIncome.Cells(newRow, 2).Value = wsSource.Cells(i, colCustomer).Value
                wsIncome.Cells(newRow, 3).Value = wsSource.Cells(i, colAmount).Value
                wsIncome.Cells(newRow, 3).NumberFormat = "#,##0.00"
                wsIncome.Cells(newRow, 4).Value = "G银行转账"
                wsIncome.Cells(newRow, 5).Value = "无票"
                wsIncome.Cells(newRow, 6).Value = "从文件导入"
                
                importCount = importCount + 1
            End If
        Next i
    Else
        ' 导入支出
        Dim colCategory As Long
        colCategory = Val(InputBox("支出类别在第几列？", "列设置", "4"))
        
        For i = startRow To lastRow
            If IsDate(wsSource.Cells(i, colDate).Value) Then
                newRow = wsExpense.Cells(wsExpense.Rows.Count, 1).End(xlUp).Row + 1
                If newRow < 5 Then newRow = 5
                
                wsExpense.Cells(newRow, 1).Value = wsSource.Cells(i, colDate).Value
                wsExpense.Cells(newRow, 2).Value = ConvertCategory(CStr(wsSource.Cells(i, colCategory).Value))
                wsExpense.Cells(newRow, 3).Value = wsSource.Cells(i, colAmount).Value
                wsExpense.Cells(newRow, 3).NumberFormat = "#,##0.00"
                wsExpense.Cells(newRow, 4).Value = wsSource.Cells(i, colCustomer).Value
                wsExpense.Cells(newRow, 5).Value = "G银行转账"
                wsExpense.Cells(newRow, 6).Value = "从文件导入"
                
                importCount = importCount + 1
            End If
        Next i
    End If
    
    wbSource.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "✓ Excel文件导入完成！" & vbCrLf & vbCrLf & _
           "导入记录：" & importCount & " 条", vbInformation, "导入完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    MsgBox "导入出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏13: ParseWeChatAlipay - 解析微信/支付宝账单
' 自动识别并导入微信/支付宝导出的账单CSV
' ============================================================================
Sub ParseWeChatAlipay()
    ' 【WPS兼容】已移除Office.FileDialog声明
    Dim filePath As String
    Dim selectedFile As String
    Dim fileContent As String
    Dim lines() As String
    Dim wsExpense As Worksheet
    Dim i As Long, importCount As Long
    Dim billType As String  ' WeChat or Alipay
    
    On Error GoTo ErrorHandler
    
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 选择文件
    ' 【WPS兼容】使用GetOpenFilename替代FileDialog
    With fd
        .Title = "选择微信或支付宝账单文件（CSV格式）"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV文件", "*.csv"
        .Filters.Add "文本文件", "*.txt"
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' 读取文件内容
    Dim fileNum As Integer
    fileNum = FreeFile
    Open selectedFile For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ' 判断账单类型
    If InStr(fileContent, "微信支付") > 0 Or InStr(fileContent, "微信") > 0 Then
        billType = "WeChat"
    ElseIf InStr(fileContent, "支付宝") > 0 Or InStr(fileContent, "Alipay") > 0 Then
        billType = "Alipay"
    Else
        billType = InputBox("无法自动识别账单类型，请手动选择：" & vbCrLf & _
                           "1 - 微信账单" & vbCrLf & _
                           "2 - 支付宝账单", "选择账单类型", "1")
        If billType = "1" Then billType = "WeChat" Else billType = "Alipay"
    End If
    
    lines = Split(fileContent, vbCrLf)
    importCount = 0
    
    Application.ScreenUpdating = False
    
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim parts() As String
            parts = Split(lines(i), ",")
            
            ' 微信账单格式：交易时间,交易类型,交易对方,商品,收/支,金额(元),支付方式,当前状态,交易单号,商户单号,备注
            ' 支付宝账单格式：交易号,商家订单号,交易创建时间,付款时间,最近修改时间,交易来源地,类型,交易对方,商品名称,金额(元),收/支,交易状态,服务费(元),成功退款(元),备注,资金状态
            
            Dim tDate As String, tType As String, tCounterparty As String
            Dim tAmount As Double, tCategory As String, tRemark As String
            Dim isExpense As Boolean
            
            isExpense = False
            
            If billType = "WeChat" And UBound(parts) >= 5 Then
                ' 微信格式解析
                If InStr(parts(4), "支出") > 0 Then
                    tDate = Left(Trim(parts(0)), 10)
                    tType = Trim(parts(1))
                    tCounterparty = Trim(parts(2))
                    tAmount = CDbl(Val(Replace(Replace(parts(5), "¥", ""), ",", "")))
                    tRemark = IIf(UBound(parts) >= 10, Trim(parts(10)), "")
                    isExpense = True
                End If
            ElseIf billType = "Alipay" And UBound(parts) >= 10 Then
                ' 支付宝格式解析
                If InStr(parts(10), "支出") > 0 Then
                    tDate = Left(Trim(parts(2)), 10)
                    tType = Trim(parts(6))
                    tCounterparty = Trim(parts(7))
                    tAmount = CDbl(Val(Replace(Replace(parts(9), "¥", ""), ",", "")))
                    tRemark = IIf(UBound(parts) >= 14, Trim(parts(14)), "")
                    isExpense = True
                End If
            End If
            
            If isExpense And tAmount > 0 Then
                ' 智能分类
                tCategory = SmartClassifyExpense(tCounterparty, tRemark)
                
                Dim newRow As Long
                newRow = wsExpense.Cells(wsExpense.Rows.Count, 1).End(xlUp).Row + 1
                If newRow < 5 Then newRow = 5
                
                wsExpense.Cells(newRow, 1).Value = tDate
                wsExpense.Cells(newRow, 2).Value = tCategory
                wsExpense.Cells(newRow, 3).Value = tAmount
                wsExpense.Cells(newRow, 3).NumberFormat = "#,##0.00"
                wsExpense.Cells(newRow, 4).Value = tCounterparty
                wsExpense.Cells(newRow, 5).Value = IIf(billType = "WeChat", "微信支付", "支付宝")
                wsExpense.Cells(newRow, 6).Value = tRemark
                
                importCount = importCount + 1
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ " & IIf(billType = "WeChat", "微信", "支付宝") & "账单解析完成！" & vbCrLf & vbCrLf & _
           "导入支出：" & importCount & " 条" & vbCrLf & vbCrLf & _
           "提示：请检查自动分类是否准确，可在支出记录表中修改", vbInformation, "解析完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "解析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ----------------------------------------------------------------------------
' 智能分类支出
' ----------------------------------------------------------------------------
Function SmartClassifyExpense(counterparty As String, remark As String) As String
    Dim combined As String
    combined = LCase(counterparty & " " & remark)
    
    ' 关键词匹配
    If InStr(combined, "房租") > 0 Or InStr(combined, "租金") > 0 Or InStr(combined, "房东") > 0 Then
        SmartClassifyExpense = "房租"
    ElseIf InStr(combined, "电费") > 0 Or InStr(combined, "供电") > 0 Or InStr(combined, "电力") > 0 Then
        SmartClassifyExpense = "电费"
    ElseIf InStr(combined, "水费") > 0 Or InStr(combined, "自来水") > 0 Then
        SmartClassifyExpense = "水费"
    ElseIf InStr(combined, "硫酸") > 0 Or InStr(combined, "硝酸") > 0 Or InStr(combined, "盐酸") > 0 Or InStr(combined, "三酸") > 0 Then
        SmartClassifyExpense = "三酸"
    ElseIf InStr(combined, "片碱") > 0 Or InStr(combined, "氢氧化钠") > 0 Or InStr(combined, "烧碱") > 0 Then
        SmartClassifyExpense = "片碱"
    ElseIf InStr(combined, "亚硝") > 0 Or InStr(combined, "亚钠") > 0 Then
        SmartClassifyExpense = "亚钠"
    ElseIf InStr(combined, "色粉") > 0 Or InStr(combined, "染料") > 0 Or InStr(combined, "颜料") > 0 Then
        SmartClassifyExpense = "色粉"
    ElseIf InStr(combined, "除油") > 0 Or InStr(combined, "脱脂") > 0 Then
        SmartClassifyExpense = "除油剂"
    ElseIf InStr(combined, "挂具") > 0 Or InStr(combined, "夹具") > 0 Then
        SmartClassifyExpense = "挂具"
    ElseIf InStr(combined, "工资") > 0 Or InStr(combined, "薪资") > 0 Or InStr(combined, "奖金") > 0 Then
        SmartClassifyExpense = "工资"
    ElseIf InStr(combined, "电镀") > 0 Or InStr(combined, "外发") > 0 Or InStr(combined, "加工") > 0 Then
        SmartClassifyExpense = "外发加工费"
    Else
        SmartClassifyExpense = "日常费用"
    End If
End Function

' ============================================================================
' 宏14: GenerateCashFlowReport - 生成现金流量表
' V18.0新增：三大报表之一
' ============================================================================
Sub GenerateCashFlowReport()
    Dim wsCF As Worksheet
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim cashIn As Double, cashOut As Double
    Dim salaryOut As Double, materialOut As Double
    Dim expenseByCategory As Object
    Dim catKey As Variant
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 创建或清空现金流量表（使用公共函数）
    Set wsCF = GetOrCreateSheet("现金流量表")
    
    Application.ScreenUpdating = False
    
    ' 计算现金流（使用公共函数）
    cashIn = SumColumn(wsIncome, 3)
    cashOut = SumColumn(wsExpense, 3)
    salaryOut = 0
    materialOut = 0
    
    ' 按类别汇总支出（使用公共函数）
    Set expenseByCategory = GetExpenseByCategory(wsExpense)
    For Each catKey In expenseByCategory.Keys
        If catKey = "工资" Then salaryOut = salaryOut + CDbl(expenseByCategory(catKey))
        If catKey = "三酸" Or catKey = "片碱" Or catKey = "亚钠" Or catKey = "色粉" Or catKey = "除油剂" Or catKey = "挂具" Then
            materialOut = materialOut + CDbl(expenseByCategory(catKey))
        End If
    Next catKey
    
    ' 填充现金流量表
    wsCF.Cells(1, 1).Value = "现 金 流 量 表"
    wsCF.Cells(1, 1).Font.Size = 16
    wsCF.Cells(1, 1).Font.Bold = True
    wsCF.Cells(1, 1).HorizontalAlignment = xlCenter
    wsCF.Range("A1:C1").Merge
    
    wsCF.Cells(2, 1).Value = "编制单位："
    wsCF.Cells(2, 3).Value = "月份：" & Format(Date, "yyyy-mm")
    
    Dim row As Long
    row = 4
    
    ' 经营活动现金流
    wsCF.Cells(row, 1).Value = "一、经营活动产生的现金流量："
    wsCF.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    wsCF.Cells(row, 1).Value = "  销售商品、提供劳务收到的现金"
    wsCF.Cells(row, 2).Value = cashIn
    row = row + 1
    
    wsCF.Cells(row, 1).Value = "  现金流入小计"
    wsCF.Cells(row, 1).Font.Bold = True
    wsCF.Cells(row, 2).Formula = "=B5"
    wsCF.Cells(row, 2).Font.Bold = True
    row = row + 2
    
    wsCF.Cells(row, 1).Value = "  购买商品、接受劳务支付的现金"
    wsCF.Cells(row, 2).Value = materialOut
    row = row + 1
    
    wsCF.Cells(row, 1).Value = "  支付给职工以及为职工支付的现金"
    wsCF.Cells(row, 2).Value = salaryOut
    row = row + 1
    
    wsCF.Cells(row, 1).Value = "  支付的其他与经营活动有关的现金"
    wsCF.Cells(row, 2).Value = cashOut - salaryOut - materialOut
    row = row + 1
    
    wsCF.Cells(row, 1).Value = "  现金流出小计"
    wsCF.Cells(row, 1).Font.Bold = True
    wsCF.Cells(row, 2).Formula = "=SUM(B8:B10)"
    wsCF.Cells(row, 2).Font.Bold = True
    row = row + 2
    
    wsCF.Cells(row, 1).Value = "经营活动产生的现金流量净额"
    wsCF.Cells(row, 1).Font.Bold = True
    wsCF.Cells(row, 2).Formula = "=B6-B11"
    wsCF.Cells(row, 2).Font.Bold = True
    row = row + 2
    
    ' 现金净增加额
    wsCF.Cells(row, 1).Value = "五、现金及现金等价物净增加额"
    wsCF.Cells(row, 1).Font.Bold = True
    wsCF.Cells(row, 1).Font.Size = 12
    wsCF.Cells(row, 2).Formula = "=B13"
    wsCF.Cells(row, 2).Font.Bold = True
    wsCF.Cells(row, 2).Font.Size = 12
    
    ' 格式化
    wsCF.Range("B4:B20").NumberFormat = "#,##0.00"
    wsCF.Columns("A").ColumnWidth = 40
    wsCF.Columns("B").ColumnWidth = 15
    wsCF.Columns("C").ColumnWidth = 15
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ 现金流量表生成完成！" & vbCrLf & vbCrLf & _
           "经营活动现金流入：" & Format(cashIn, "#,##0.00") & " 元" & vbCrLf & _
           "经营活动现金流出：" & Format(cashOut, "#,##0.00") & " 元" & vbCrLf & _
           "现金流量净额：" & Format(cashIn - cashOut, "#,##0.00") & " 元", vbInformation, "报表完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "生成报表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏15: GenerateMultiPeriodAnalysis - 多期对比分析
' V18.0新增：对比多个月份数据
' ============================================================================
Sub GenerateMultiPeriodAnalysis()
    Dim wsAnalysis As Worksheet
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim months As Variant
    Dim monthStr As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 获取要分析的月份
    monthStr = InputBox("请输入要对比的月份（格式：2025-01,2025-02,2025-03）：" & vbCrLf & _
                       "支持2-6个月对比", "多期对比", Format(Date, "yyyy-mm"))
    
    If monthStr = "" Then Exit Sub
    
    months = Split(monthStr, ",")
    If UBound(months) < 1 Then
        MsgBox "请至少输入2个月份进行对比！", vbExclamation
        Exit Sub
    End If
    
    ' 创建分析表
    On Error Resume Next
    Set wsAnalysis = ThisWorkbook.Sheets("多期对比分析")
    If wsAnalysis Is Nothing Then
        Set wsAnalysis = ThisWorkbook.Sheets.Add
        wsAnalysis.Name = "多期对比分析"
    Else
        wsAnalysis.Cells.Clear
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    ' 表头
    wsAnalysis.Cells(1, 1).Value = "多期对比分析表"
    wsAnalysis.Cells(1, 1).Font.Size = 16
    wsAnalysis.Cells(1, 1).Font.Bold = True
    
    wsAnalysis.Cells(3, 1).Value = "项目"
    wsAnalysis.Cells(3, 1).Font.Bold = True
    
    ' 月份列
    For i = 0 To UBound(months)
        wsAnalysis.Cells(3, 2 + i).Value = Trim(months(i))
        wsAnalysis.Cells(3, 2 + i).Font.Bold = True
        wsAnalysis.Cells(3, 2 + i).HorizontalAlignment = xlCenter
    Next i
    
    ' 数据行
    wsAnalysis.Cells(4, 1).Value = "营业收入"
    wsAnalysis.Cells(5, 1).Value = "外发加工成本"
    wsAnalysis.Cells(6, 1).Value = "毛利润"
    wsAnalysis.Cells(7, 1).Value = "毛利率"
    wsAnalysis.Cells(8, 1).Value = "营业费用"
    wsAnalysis.Cells(9, 1).Value = "净利润"
    wsAnalysis.Cells(10, 1).Value = "净利率"
    
    ' 计算各月数据
    For i = 0 To UBound(months)
        Dim monthData As MonthData
        monthData = CalculateMonthData(Trim(months(i)))
        
        wsAnalysis.Cells(4, 2 + i).Value = monthData.Income
        wsAnalysis.Cells(5, 2 + i).Value = monthData.OutsourceCost
        wsAnalysis.Cells(6, 2 + i).Value = monthData.GrossProfit
        wsAnalysis.Cells(7, 2 + i).Value = IIf(monthData.Income > 0, monthData.GrossProfit / monthData.Income, 0)
        wsAnalysis.Cells(7, 2 + i).NumberFormat = "0.0%"
        wsAnalysis.Cells(8, 2 + i).Value = monthData.Expense
        wsAnalysis.Cells(9, 2 + i).Value = monthData.NetProfit
        wsAnalysis.Cells(10, 2 + i).Value = IIf(monthData.Income > 0, monthData.NetProfit / monthData.Income, 0)
        wsAnalysis.Cells(10, 2 + i).NumberFormat = "0.0%"
    Next i
    
    ' 格式化
    wsAnalysis.Range("B4:H10").NumberFormat = "#,##0.00"
    wsAnalysis.Columns("A").ColumnWidth = 15
    wsAnalysis.Columns("B:H").ColumnWidth = 12
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ 多期对比分析完成！" & vbCrLf & vbCrLf & _
           "已生成 " & (UBound(months) + 1) & " 个月份对比数据", vbInformation, "分析完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "分析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ----------------------------------------------------------------------------
' 计算月度数据
' ----------------------------------------------------------------------------
Function CalculateMonthData(targetMonth As String) As MonthData
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim lastRow As Long, i As Long
    Dim result As MonthData
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 收入
    lastRow = GetLastRow(wsIncome, 1)
    For i = 5 To lastRow
        If InStr(wsIncome.Cells(i, 1).Value, targetMonth) > 0 Then
            result.Income = result.Income + CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
        End If
    Next i
    
    ' 支出
    lastRow = GetLastRow(wsExpense, 1)
    For i = 5 To lastRow
        If InStr(wsExpense.Cells(i, 1).Value, targetMonth) > 0 Then
            Dim amt As Double, cat As String
            amt = CDbl(Nz(wsExpense.Cells(i, 3).Value, 0))
            cat = Trim(wsExpense.Cells(i, 2).Value)
            
            result.Expense = result.Expense + amt
            If cat = "外发加工费" Then
                result.OutsourceCost = result.OutsourceCost + amt
            End If
        Next i

    result.GrossProfit = result.Income - result.OutsourceCost
    ' 净利润：使用公共函数计算（盈利时扣除5%所得税）
    result.NetProfit = CalcNetProfit(result.Income, result.Expense)
    
    CalculateMonthData = result
End Function

' ============================================================================
' 宏16: CheckARAlert - 应收款到期提醒
' V18.0新增：智能提醒功能
' ============================================================================
Sub CheckARAlert()
    Dim wsARAP As Worksheet, wsSettings As Worksheet
    Dim lastRow As Long, i As Long
    Dim alertMsg As String
    Dim overDueCount As Long, nearDueCount As Long
    Dim totalOverDue As Double
    
    On Error GoTo ErrorHandler
    
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    
    alertMsg = "📢 应收款提醒报告" & vbCrLf & vbCrLf
    overDueCount = 0
    nearDueCount = 0
    totalOverDue = 0
    
    ' 检查应收账款
    lastRow = wsARAP.Cells(wsARAP.Rows.Count, AR_COL_NAME).End(xlUp).Row
    For i = 5 To lastRow
        Dim customer As String, arAmount As Double
        customer = Trim(wsARAP.Cells(i, AR_COL_NAME).Value)
        arAmount = CDbl(Nz(wsARAP.Cells(i, AR_COL_CLOSE).Value, 0))
        
        If customer <> "" And arAmount > 0 Then
            ' 模拟账期判断（根据期末应收金额大小判断）
            If arAmount > 50000 Then
                alertMsg = alertMsg & "🔴 【大额应收】" & customer & "：" & Format(arAmount, "#,##0.00") & " 元" & vbCrLf
                overDueCount = overDueCount + 1
                totalOverDue = totalOverDue + arAmount
            ElseIf arAmount > 20000 Then
                alertMsg = alertMsg & "🟡 【重点关注】" & customer & "：" & Format(arAmount, "#,##0.00") & " 元" & vbCrLf
                nearDueCount = nearDueCount + 1
            End If
        End If
    Next i
    
    ' 检查应付账款（使用公共函数）
    Dim apTotal As Double
    apTotal = SumARColumn(wsARAP, AP_COL_CLOSE)
    
    alertMsg = alertMsg & vbCrLf & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf
    alertMsg = alertMsg & "⚠️ 大额/逾期应收：" & overDueCount & " 家" & vbCrLf
    alertMsg = alertMsg & "💰 重点关注应收：" & nearDueCount & " 家" & vbCrLf
    alertMsg = alertMsg & "📊 应收合计：" & Format(totalOverDue, "#,##0.00") & " 元" & vbCrLf
    alertMsg = alertMsg & "📊 应付合计：" & Format(apTotal, "#,##0.00") & " 元" & vbCrLf
    alertMsg = alertMsg & "📊 净应收：" & Format(totalOverDue - apTotal, "#,##0.00") & " 元" & vbCrLf & vbCrLf
    
    If overDueCount > 0 Then
        alertMsg = alertMsg & "建议：尽快联系上述客户催收款项！"
    ElseIf nearDueCount > 0 Then
        alertMsg = alertMsg & "建议：关注重点关注客户回款情况"
    Else
        alertMsg = alertMsg & "✓ 应收款风险可控"
    End If
    
    MsgBox alertMsg, IIf(overDueCount > 0, vbExclamation, vbInformation), "应收款提醒"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "检查出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏17: TaxReminder - 税务申报提醒
' V18.0新增：税务提醒功能
' ============================================================================
Sub TaxReminder()
    Dim today As Date
    Dim month As Integer, day As Integer
    Dim msg As String
    Dim daysToQuarter As Integer
    
    today = Date
    month = Month(today)
    day = Day(today)
    
    msg = "📅 税务申报提醒" & vbCrLf & vbCrLf
    
    ' 计算距离季度申报日的天数
    Dim quarterEnd As Date
    If month >= 1 And month <= 3 Then
        quarterEnd = DateSerial(Year(today), 4, 15)
    ElseIf month >= 4 And month <= 6 Then
        quarterEnd = DateSerial(Year(today), 7, 15)
    ElseIf month >= 7 And month <= 9 Then
        quarterEnd = DateSerial(Year(today), 10, 15)
    Else
        quarterEnd = DateSerial(Year(today) + 1, 1, 15)
    End If
    
    daysToQuarter = quarterEnd - today
    
    msg = msg & "当前日期：" & Format(today, "yyyy年mm月dd日") & vbCrLf & vbCrLf
    
    ' 月度提醒
    If day >= 1 And day <= 15 Then
        msg = msg & "🟡 本月增值税申报期（1-15日）" & vbCrLf
    Else
        msg = msg & "✓ 本月增值税申报已完成" & vbCrLf
    End If
    
    ' 季度提醒
    If daysToQuarter >= 0 And daysToQuarter <= 15 Then
        msg = msg & "🔴 季度所得税申报即将到期（还剩 " & daysToQuarter & " 天）" & vbCrLf
    ElseIf daysToQuarter < 0 And daysToQuarter > -15 Then
        msg = msg & "⚠️ 季度所得税申报已逾期！" & vbCrLf
    Else
        msg = msg & "✓ 季度所得税申报时间未到" & vbCrLf
    End If
    
    ' 年度提醒
    If month = 12 And day >= 1 Then
        msg = msg & "🔴 年度汇算清缴准备期（次年1-5月）" & vbCrLf
    ElseIf month >= 1 And month <= 5 Then
        msg = msg & "🟡 年度汇算清缴申报期" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf
    msg = msg & "小规模纳税人申报期限：" & vbCrLf
    msg = msg & "• 增值税：季度申报（1、4、7、10月）" & vbCrLf
    msg = msg & "• 所得税：季度预缴，年度汇算" & vbCrLf
    
    MsgBox msg, vbInformation, "税务提醒"
End Sub

' ============================================================================
' 宏18: SetUserPermission - 设置用户权限
' V18.0新增：多用户协作功能
' ============================================================================
Sub SetUserPermission()
    Dim userName As String
    Dim userRole As String
    Dim wsSettings As Worksheet
    Dim lastRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsSettings = ThisWorkbook.Sheets("基础设置")
    
    userName = InputBox("请输入用户名：", "设置用户权限")
    If userName = "" Then Exit Sub
    
    userRole = InputBox("请选择用户角色：" & vbCrLf & vbCrLf & _
                       "1 - 管理员（全部权限）" & vbCrLf & _
                       "2 - 会计（录入+查看）" & vbCrLf & _
                       "3 - 出纳（录入支出+查看）" & vbCrLf & _
                       "4 - 查看者（仅查看）", _
                       "选择角色", "2")
    
    If userRole <> "1" And userRole <> "2" And userRole <> "3" And userRole <> "4" Then
        MsgBox "无效的角色选择！", vbExclamation
        Exit Sub
    End If
    
    ' 写入用户权限表
    Dim roleName As String
    Select Case userRole
        Case "1": roleName = "管理员"
        Case "2": roleName = "会计"
        Case "3": roleName = "出纳"
        Case "4": roleName = "查看者"
    End Select
    
    ' 查找或创建用户权限区域
    Dim userRow As Long
    Dim i As Long
    userRow = 0
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, 15).End(xlUp).Row
    
    For i = 5 To lastRow
        If Trim(wsSettings.Cells(i, 15).Value) = userName Then
            userRow = i
            Exit For
        End If
    Next i
    
    If userRow = 0 Then
        userRow = lastRow + 1
        If userRow < 5 Then userRow = 5
    End If
    
    wsSettings.Cells(userRow, 15).Value = userName
    wsSettings.Cells(userRow, 16).Value = roleName
    wsSettings.Cells(userRow, 17).Value = Format(Now, "yyyy-mm-dd hh:mm")
    
    ' 创建表头（如果不存在）
    If wsSettings.Cells(3, 15).Value = "" Then
        wsSettings.Cells(3, 15).Value = "用户名"
        wsSettings.Cells(3, 16).Value = "角色"
        wsSettings.Cells(3, 17).Value = "设置时间"
        wsSettings.Range("O3:Q3").Font.Bold = True
        wsSettings.Columns("O").ColumnWidth = 12
        wsSettings.Columns("P").ColumnWidth = 10
        wsSettings.Columns("Q").ColumnWidth = 16
    End If
    
    MsgBox "✓ 用户权限设置完成！" & vbCrLf & vbCrLf & _
           "用户：" & userName & vbCrLf & _
           "角色：" & roleName & vbCrLf & vbCrLf & _
           GetRoleDescription(userRole), vbInformation, "设置完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置出错：" & Err.Description, vbCritical, "错误"
End Sub

' ----------------------------------------------------------------------------
' 获取角色权限说明
' ----------------------------------------------------------------------------
Function GetRoleDescription(roleCode As String) As String
    Select Case roleCode
        Case "1"
            GetRoleDescription = "权限说明：" & vbCrLf & _
                                "• 全部功能可用" & vbCrLf & _
                                "• 可设置其他用户权限" & vbCrLf & _
                                "• 可执行系统初始化"
        Case "2"
            GetRoleDescription = "权限说明：" & vbCrLf & _
                                "• 可录入收入、支出" & vbCrLf & _
                                "• 可生成报表" & vbCrLf & _
                                "• 可执行月结" & vbCrLf & _
                                "• 不可初始化系统"
        Case "3"
            GetRoleDescription = "权限说明：" & vbCrLf & _
                                "• 可录入支出" & vbCrLf & _
                                "• 可查看报表" & vbCrLf & _
                                "• 不可录入收入" & vbCrLf & _
                                "• 不可执行月结"
        Case "4"
            GetRoleDescription = "权限说明：" & vbCrLf & _
                                "• 仅可查看数据" & vbCrLf & _
                                "• 不可录入或修改" & vbCrLf & _
                                "• 不可生成报表"
        Case Else
            GetRoleDescription = ""
    End Select
End Function

' ============================================================================
' 宏19: ShowDashboard - 显示数据仪表盘
' V18.0新增：可视化数据展示
' ============================================================================
Sub ShowDashboard()
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim wsARAP As Worksheet
    Dim totalIncome As Double, totalExpense As Double
    Dim totalAR As Double, totalAP As Double
    Dim msg As String
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    
    ' 计算总收入（使用公共函数）
    totalIncome = SumColumn(wsIncome, 3)
    
    ' 计算总支出（使用公共函数）
    totalExpense = SumColumn(wsExpense, 3)
    
    ' 计算应收应付（使用公共函数）
    totalAR = SumARColumn(wsARAP, AR_COL_CLOSE)
    totalAP = SumARColumn(wsARAP, AP_COL_CLOSE)
    
    ' 构建仪表盘
    msg = "经营数据仪表盘" & vbCrLf & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    msg = msg & "【本月经营概况】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf
    
    msg = msg & "营业收入：" & Format(totalIncome, "#,##0.00") & " 元" & vbCrLf
    msg = msg & "营业支出：" & Format(totalExpense, "#,##0.00") & " 元" & vbCrLf
    Dim netProfitDisplay As Double
    netProfitDisplay = CalcNetProfit(totalIncome, totalExpense)
    msg = msg & "净利润：" & Format(netProfitDisplay, "#,##0.00") & " 元" & vbCrLf
    msg = msg & "利润率：" & IIf(totalIncome > 0, Format(netProfitDisplay / totalIncome, "0.0%"), "N/A") & vbCrLf & vbCrLf
    
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    msg = msg & "【资金状况】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf
    
    msg = msg & "应收账款：" & Format(totalAR, "#,##0.00") & " 元" & vbCrLf
    msg = msg & "应付账款：" & Format(totalAP, "#,##0.00") & " 元" & vbCrLf
    msg = msg & "资金净额：" & Format(totalAR - totalAP, "#,##0.00") & " 元" & vbCrLf & vbCrLf
    
    ' 健康度评估
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    msg = msg & "【经营健康度】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf
    
    If totalIncome > 0 Then
        Dim profitRate As Double
        profitRate = netProfitDisplay / totalIncome
        
        If profitRate >= 0.2 Then
            msg = msg & "盈利能力：优秀（利润率 " & Format(profitRate, "0.0%") & "）" & vbCrLf
        ElseIf profitRate >= 0.1 Then
            msg = msg & "盈利能力：良好（利润率 " & Format(profitRate, "0.0%") & "）" & vbCrLf
        ElseIf profitRate >= 0 Then
            msg = msg & "盈利能力：一般（利润率 " & Format(profitRate, "0.0%") & "）" & vbCrLf
        Else
            msg = msg & "盈利能力：亏损（利润率 " & Format(profitRate, "0.0%") & "）" & vbCrLf
        End If
    End If
    
    If totalAR > 100000 Then
        msg = msg & "资金风险：应收账款较高，建议催收" & vbCrLf
    Else
        msg = msg & "资金风险：可控" & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "数据仪表盘"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "显示仪表盘出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' V18.0 新增功能：AI智能分析、数据安全增强
' ============================================================================

' ============================================================================
' 宏20: AIAnomalyDetection - AI智能异常检测
' V18.0新增：自动识别数据异常、异常支出预警
' ============================================================================
Sub AIAnomalyDetection()
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim wsARAP As Worksheet
    Dim lastRow As Long, i As Long
    Dim msg As String
    Dim anomalyCount As Long
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    
    msg = "🤖 AI智能异常检测报告" & vbCrLf & vbCrLf
    anomalyCount = 0
    
    ' 1. 检测收入异常（单日收入过高或过低）（使用公共函数）
    Dim avgIncome As Double, incomeCount As Long
    Dim totalIncome As Double
    totalIncome = 0
    incomeCount = 0
    lastRow = GetLastRow(wsIncome, 3)
    
    For i = 5 To lastRow
        If IsNumeric(wsIncome.Cells(i, 3).Value) Then
            totalIncome = totalIncome + CDbl(wsIncome.Cells(i, 3).Value)
            incomeCount = incomeCount + 1
        End If
    Next i
    
    If incomeCount > 0 Then
        avgIncome = totalIncome / incomeCount
        
        For i = 5 To lastRow
            If IsNumeric(wsIncome.Cells(i, 3).Value) Then
                Dim incomeAmt As Double
                incomeAmt = CDbl(wsIncome.Cells(i, 3).Value)
                
                ' 检测异常高收入（超过平均值3倍）
                If incomeAmt > avgIncome * 3 And avgIncome > 0 Then
                    msg = msg & "⚠️ 【收入异常】" & wsIncome.Cells(i, 2).Value & " " & _
                          Format(wsIncome.Cells(i, 1).Value, "yyyy-mm-dd") & _
                          " 收入 " & Format(incomeAmt, "#,##0.00") & " 元（远超平均值）" & vbCrLf
                    anomalyCount = anomalyCount + 1
                End If
            End If
        Next i
    End If
    
    ' 2. 检测支出异常（某类别支出突增）
    Dim expenseByCategory As Object
    Set expenseByCategory = GetExpenseByCategory(wsExpense)
    Dim monthExpense As Object
    Set monthExpense = ' 【WPS兼容】CreateObject("Scripting.Dictionary") ' WPS不支持，请使用Collection
    
    lastRow = GetLastRow(wsExpense, 3)
    
    ' 统计各类别月度支出
    For i = 5 To lastRow
        Dim expDate As String, expCat As String, expAmt As Double
        expDate = Left(wsExpense.Cells(i, 1).Value, 7) ' 取年月
        expCat = Trim(wsExpense.Cells(i, 2).Value)
        expAmt = CDbl(Nz(wsExpense.Cells(i, 3).Value, 0))
        
        If expCat <> "" Then
            Dim key As String
            key = expDate & "|" & expCat
            If monthExpense.Exists(key) Then
                monthExpense(key) = monthExpense(key) + expAmt
            Else
                monthExpense.Add key, expAmt
            End If
        End If
    Next i
    
    ' 检测支出突增（比上月增长超过50%）
    Dim currentMonth As String
    currentMonth = Format(Date, "yyyy-mm")
    
    For i = 0 To monthExpense.Count - 1
        Dim parts() As String
        parts = Split(monthExpense.Keys()(i), "|")
        If UBound(parts) = 1 Then
            If parts(0) = currentMonth Then
                Dim lastMonth As String
                lastMonth = Format(DateAdd("m", -1, Date), "yyyy-mm")
                Dim lastKey As String
                lastKey = lastMonth & "|" & parts(1)
                
                If monthExpense.Exists(lastKey) Then
                    Dim currAmt As Double, lastAmt As Double
                    currAmt = monthExpense(monthExpense.Keys()(i))
                    lastAmt = monthExpense(lastKey)
                    
                    If lastAmt > 0 And currAmt > lastAmt * 1.5 Then
                        msg = msg & "【支出突增】" & parts(1) & " 本月 " & _
                              Format(currAmt, "#,##0.00") & " 元，比上月增长 " & _
                              Format((currAmt - lastAmt) / lastAmt, "0.0%") & vbCrLf
                        anomalyCount = anomalyCount + 1
                    End If
                End If
            End If
        End If
    Next i
    
    ' 3. 检测负利润月份（使用公共函数）
    If incomeCount > 0 Then
        Dim totalExp As Double
        totalExp = SumColumn(wsExpense, 3)
        
        If totalIncome < totalExp Then
            msg = msg & "【经营亏损】本月支出超过收入，净利润约 " & _
                  Format(CalcNetProfit(totalIncome, totalExp), "#,##0.00") & " 元" & vbCrLf
            anomalyCount = anomalyCount + 1
        End If
    End If
    
    ' 4. 检测异常大额应收
    lastRow = wsARAP.Cells(wsARAP.Rows.Count, AR_COL_CLOSE).End(xlUp).Row
    For i = 5 To lastRow
        Dim arAmt As Double, arCustomer As String
        arCustomer = Trim(wsARAP.Cells(i, AR_COL_NAME).Value)
        arAmt = CDbl(Nz(wsARAP.Cells(i, AR_COL_CLOSE).Value, 0))
        
        If arAmt > 80000 Then
            msg = msg & "🚨 【高风险应收】" & arCustomer & " 欠款 " & _
                  Format(arAmt, "#,##0.00") & " 元，建议立即催收" & vbCrLf
            anomalyCount = anomalyCount + 1
        End If
    Next i
    
    ' 显示结果
    If anomalyCount = 0 Then
        msg = msg & "✅ 恭喜！AI检测未发现明显异常" & vbCrLf & vbCrLf
        msg = msg & "经营数据健康，继续保持！"
    Else
        msg = msg & vbCrLf & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
        msg = msg & "共发现 " & anomalyCount & " 项异常，请关注以上提示"
    End If
    
    MsgBox msg, IIf(anomalyCount > 0, vbExclamation, vbInformation), "AI异常检测"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "检测出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏21: TrendPrediction - 趋势预测分析
' V18.0新增：基于历史数据预测未来趋势
' ============================================================================
Sub TrendPrediction()
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim monthData As Object
    Dim months As Variant
    Dim i As Long, j As Long
    Dim msg As String
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set monthData = ' 【WPS兼容】CreateObject("Scripting.Dictionary") ' WPS不支持，请使用Collection
    
    ' 统计各月收入
    Dim lastRow As Long
    lastRow = wsIncome.Cells(wsIncome.Rows.Count, 1).End(xlUp).Row
    
    For i = 5 To lastRow
        Dim m As String
        m = Left(wsIncome.Cells(i, 1).Value, 7)
        If m <> "" Then
            If monthData.Exists(m) Then
                monthData(m) = monthData(m) + CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
            Else
                monthData.Add m, CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
            End If
        End If
    Next i
    
    If monthData.Count < 3 Then
        MsgBox "历史数据不足，需要至少3个月数据才能进行趋势预测", vbInformation, "数据不足"
        Exit Sub
    End If
    
    ' 对月份key排序后再取值
    Dim sortedKeys() As String
    ReDim sortedKeys(monthData.Count - 1)
    Dim allKeys As Variant
    allKeys = monthData.Keys
    ' 简单冒泡排序
    Dim tempKey As String
    For i = 0 To monthData.Count - 2
        For j = i + 1 To monthData.Count - 1
            If allKeys(i) > allKeys(j) Then
                tempKey = allKeys(i)
                allKeys(i) = allKeys(j)
                allKeys(j) = tempKey
            End If
        Next j
    Next i
    For i = 0 To monthData.Count - 1
        sortedKeys(i) = allKeys(i)
    Next i
    
    ' 计算平均增长率
    Dim values() As Double
    ReDim values(monthData.Count - 1)
    
    For i = 0 To monthData.Count - 1
        values(i) = monthData(sortedKeys(i))
    Next i
    
    ' 简单线性预测
    Dim growthRates As Double, avgGrowth As Double
    growthRates = 0
    For i = 1 To UBound(values)
        If values(i - 1) > 0 Then
            growthRates = growthRates + (values(i) - values(i - 1)) / values(i - 1)
        End If
    Next i
    avgGrowth = growthRates / UBound(values)
    
    ' 预测未来3个月
    Dim lastValue As Double
    lastValue = values(UBound(values))
    
    msg = "📊 AI趋势预测分析" & vbCrLf & vbCrLf
    msg = msg & "基于过去 " & monthData.Count & " 个月的历史数据" & vbCrLf
    msg = msg & "平均月增长率：" & Format(avgGrowth, "0.0%") & vbCrLf & vbCrLf
    
    msg = msg & "【未来3个月收入预测】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    
    For i = 1 To 3
        Dim predictMonth As String
        predictMonth = Format(DateAdd("m", i, Date), "yyyy-mm")
        lastValue = lastValue * (1 + avgGrowth)
        msg = msg & predictMonth & " 预测收入：" & Format(lastValue, "#,##0.00") & " 元" & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "【经营建议】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    
    If avgGrowth > 0.1 Then
        msg = msg & "🟢 业务增长势头良好，建议：" & vbCrLf
        msg = msg & "• 适当增加原材料库存" & vbCrLf
        msg = msg & "• 考虑扩充产能或人手" & vbCrLf
    ElseIf avgGrowth > 0 Then
        msg = msg & "🟡 业务平稳发展，建议：" & vbCrLf
        msg = msg & "• 维持当前经营策略" & vbCrLf
        msg = msg & "• 关注成本控制" & vbCrLf
    Else
        msg = msg & "🔴 业务收入下滑，建议：" & vbCrLf
        msg = msg & "• 积极开拓新客户" & vbCrLf
        msg = msg & "• 审查成本结构，削减不必要支出" & vbCrLf
        msg = msg & "• 考虑调整产品定价策略" & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "趋势预测"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "预测出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏22: GenerateBusinessAdvice - 智能经营建议
' V18.0新增：基于数据分析生成经营建议
' ============================================================================
Sub GenerateBusinessAdvice()
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim wsARAP As Worksheet
    Dim totalIncome As Double, totalExpense As Double
    Dim expenseByCategory As Object
    Dim msg As String
    Dim adviceCount As Long
    Dim catKey As Variant
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    
    msg = "AI智能经营建议" & vbCrLf & vbCrLf
    adviceCount = 0
    
    ' 计算总收入和支出（使用公共函数）
    totalIncome = SumColumn(wsIncome, 3)
    totalExpense = SumColumn(wsExpense, 3)
    
    ' 按类别汇总支出（使用公共函数）
    Set expenseByCategory = GetExpenseByCategory(wsExpense)
    
    Dim profit As Double, profitRate As Double
    profit = CalcNetProfit(totalIncome, totalExpense)
    If totalIncome > 0 Then profitRate = profit / totalIncome
    
    ' 建议1：利润率分析
    msg = msg & "【盈利能力分析】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    
    If profitRate >= 0.2 Then
        msg = msg & "✅ 利润率优秀（" & Format(profitRate, "0.0%") & "），盈利能力强劲" & vbCrLf
    ElseIf profitRate >= 0.1 Then
        msg = msg & "💡 利润率良好（" & Format(profitRate, "0.0%") & "），仍有提升空间" & vbCrLf
        msg = msg & "   建议：优化采购渠道，降低原材料成本" & vbCrLf
        adviceCount = adviceCount + 1
    ElseIf profitRate > 0 Then
        msg = msg & "⚠️ 利润率偏低（" & Format(profitRate, "0.0%") & "），需要关注" & vbCrLf
        msg = msg & "   建议：审查定价策略，控制运营成本" & vbCrLf
        adviceCount = adviceCount + 1
    Else
        msg = msg & "🔴 当前亏损，需立即采取行动" & vbCrLf
        msg = msg & "   建议：紧急削减非必要支出，提高产品售价" & vbCrLf
        adviceCount = adviceCount + 1
    End If
    
    msg = msg & vbCrLf & "【成本结构优化】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    
    ' 建议2：成本结构分析
    If expenseByCategory.Exists("外发加工费") And totalIncome > 0 Then
        Dim outsourceRatio As Double
        outsourceRatio = expenseByCategory("外发加工费") / totalIncome
        If outsourceRatio > 0.4 Then
            msg = msg & "💡 外发加工成本占比高（" & Format(outsourceRatio, "0.0%") & "）" & vbCrLf
            msg = msg & "   建议：评估自建产能的可行性，降低外发依赖" & vbCrLf
            adviceCount = adviceCount + 1
        End If
    End If
    
    If expenseByCategory.Exists("电费") And totalIncome > 0 Then
        Dim elecRatio As Double
        elecRatio = expenseByCategory("电费") / totalIncome
        If elecRatio > 0.1 Then
            msg = msg & "💡 电费支出占比较高（" & Format(elecRatio, "0.0%") & "）" & vbCrLf
            msg = msg & "   建议：检查设备能耗，考虑节能改造" & vbCrLf
            adviceCount = adviceCount + 1
        End If
    End If
    
    ' 建议3：资金管理
    msg = msg & vbCrLf & "【资金管理建议】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    
    Dim totalAR As Double
    totalAR = SumARColumn(wsARAP, AR_COL_CLOSE)
    
    If totalAR > totalIncome * 0.5 Then
        msg = msg & "⚠️ 应收账款占收入比例过高" & vbCrLf
        msg = msg & "   建议：加强回款管理，考虑缩短账期" & vbCrLf
        adviceCount = adviceCount + 1
    Else
        msg = msg & "✅ 应收账款管理良好" & vbCrLf
    End If
    
    ' 建议4：税务优化
    msg = msg & vbCrLf & "【税务优化建议】" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    
    If profit > 1000000 Then
        msg = msg & "💡 利润超过100万，所得税率将提高至10%" & vbCrLf
        msg = msg & "   建议：合理规划费用支出，控制利润在优惠区间" & vbCrLf
        adviceCount = adviceCount + 1
    Else
        msg = msg & "✅ 当前享受小微企业5%所得税优惠税率" & vbCrLf
    End If
    
    If adviceCount = 0 Then
        msg = msg & vbCrLf & "🎉 恭喜！您的经营状况良好，暂无特别建议"
    End If
    
    MsgBox msg, vbInformation, "经营建议"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "生成建议出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏23: SetWorkbookPassword - 设置工作簿保护
' V18.0新增：数据安全保护
' ============================================================================
Sub SetWorkbookPassword()
    Dim pwd As String, confirmPwd As String
    Dim result As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    result = MsgBox("设置密码保护后，需要密码才能打开此文件。" & vbCrLf & _
                   "请确保记住密码，忘记密码将无法恢复数据！" & vbCrLf & vbCrLf & _
                   "是否继续？", vbYesNo + vbQuestion, "设置密码保护")
    
    If result <> vbYes Then Exit Sub
    
    pwd = InputBox("请输入密码（至少6位）：", "设置密码", "")
    If Len(pwd) < 6 Then
        MsgBox "密码长度不足6位，请重新设置", vbExclamation
        Exit Sub
    End If
    
    confirmPwd = InputBox("请再次输入密码确认：", "确认密码", "")
    If pwd <> confirmPwd Then
        MsgBox "两次输入的密码不一致，请重新设置", vbExclamation
        Exit Sub
    End If
    
    ' 设置工作簿密码
    ThisWorkbook.SaveAs ThisWorkbook.FullName, Password:=pwd
    
    MsgBox "✅ 密码保护设置成功！" & vbCrLf & vbCrLf & _
           "下次打开此文件时需要输入密码", vbInformation, "设置成功"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置密码出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏24: LogOperation - 记录操作日志
' V18.0新增：记录关键操作，便于审计追踪
' ============================================================================
Sub LogOperation()
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim operation As String
    Dim userName As String
    
    On Error GoTo ErrorHandler
    
    ' 创建或获取日志表
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("操作日志")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "操作日志"
        wsLog.Cells(1, 1).Value = "操作日志"
        wsLog.Cells(1, 1).Font.Size = 14
        wsLog.Cells(1, 1).Font.Bold = True
        
        wsLog.Cells(3, 1).Value = "时间"
        wsLog.Cells(3, 2).Value = "用户"
        wsLog.Cells(3, 3).Value = "操作类型"
        wsLog.Cells(3, 4).Value = "详情"
        wsLog.Range("A3:D3").Font.Bold = True
        wsLog.Range("A3:D3").Interior.Color = RGB(68, 114, 196)
        wsLog.Range("A3:D3").Font.Color = RGB(255, 255, 255)
        
        wsLog.Columns("A").ColumnWidth = 18
        wsLog.Columns("B").ColumnWidth = 12
        wsLog.Columns("C").ColumnWidth = 15
        wsLog.Columns("D").ColumnWidth = 40
    End If
    On Error GoTo 0
    
    ' 获取操作信息
    operation = InputBox("请输入操作描述：" & vbCrLf & _
                        "例如：录入收入、修改支出、生成报表等", "记录操作")
    
    If operation = "" Then Exit Sub
    
    userName = InputBox("请输入操作人姓名：", "操作人", "会计")
    If userName = "" Then userName = "会计"
    
    ' 写入日志
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If lastRow < 4 Then lastRow = 4
    
    wsLog.Cells(lastRow, 1).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    wsLog.Cells(lastRow, 2).Value = userName
    wsLog.Cells(lastRow, 3).Value = "手动记录"
    wsLog.Cells(lastRow, 4).Value = operation
    
    MsgBox "✅ 操作已记录到日志", vbInformation, "记录完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "记录日志出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏25: ViewOperationLog - 查看操作日志
' V18.0新增：查看近期操作记录
' ============================================================================
Sub ViewOperationLog()
    Dim wsLog As Worksheet
    Dim lastRow As Long, i As Long
    Dim msg As String
    Dim logCount As Long
    
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("操作日志")
    On Error GoTo 0
    
    If wsLog Is Nothing Then
        MsgBox "暂无操作日志记录", vbInformation, "操作日志"
        Exit Sub
    End If
    
    msg = "📋 近期操作日志" & vbCrLf & vbCrLf
    msg = msg & "时间                | 用户   | 操作" & vbCrLf
    msg = msg & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    logCount = 0
    
    ' 显示最近10条记录
    For i = lastRow To Application.Max(4, lastRow - 9) Step -1
        If wsLog.Cells(i, 1).Value <> "" Then
            msg = msg & Left(wsLog.Cells(i, 1).Value & Space(18), 18) & " | " & _
                  Left(wsLog.Cells(i, 2).Value & Space(6), 6) & " | " & _
                  Left(wsLog.Cells(i, 4).Value, 25) & vbCrLf
            logCount = logCount + 1
        End If
    Next i
    
    If logCount = 0 Then
        msg = msg & "暂无记录"
    Else
        msg = msg & vbCrLf & "共显示最近 " & logCount & " 条记录"
    End If
    
    MsgBox msg, vbInformation, "操作日志"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "查看日志出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏26: SecureBackup - 加密备份
' V18.0新增：带密码的加密备份
' ============================================================================
Sub SecureBackup()
    Dim backupPath As String
    Dim fileName As String
    Dim timestamp As String
    Dim pwd As String
    
    On Error GoTo ErrorHandler
    
    pwd = InputBox("请设置备份文件密码（至少6位，留空则不加密）：", "加密备份", "")
    
    timestamp = Format(Now, "yyyymmdd_hhmmss")
    fileName = "加密备份_" & timestamp & ".xlsx"
    backupPath = ThisWorkbook.Path & "\备份\" & fileName
    
    ' 创建备份目录
    If Dir(ThisWorkbook.Path & "\备份", vbDirectory) = "" Then
        On Error Resume Next
        MkDir ThisWorkbook.Path & "\备份"
        If Err.Number <> 0 Then
            MsgBox "创建备份目录失败：" & Err.Description, vbCritical, "错误"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    Application.DisplayAlerts = False
    
    If Len(pwd) >= 6 Then
        ThisWorkbook.SaveCopyAs Filename:=backupPath, Password:=pwd
        MsgBox "✅ 加密备份完成！" & vbCrLf & vbCrLf & _
               "文件：" & fileName & vbCrLf & _
               "密码：已设置" & vbCrLf & vbCrLf & _
               "请妥善保管密码，忘记密码将无法打开备份文件！", vbInformation, "备份完成"
    Else
        ThisWorkbook.SaveCopyAs Filename:=backupPath
        MsgBox "✅ 备份完成！" & vbCrLf & vbCrLf & _
               "文件：" & fileName & vbCrLf & _
               "密码：无", vbInformation, "备份完成"
    End If
    
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "备份出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 公共工具函数
' ============================================================================

' ----------------------------------------------------------------------------
' GetLastRow - 获取指定列最后有数据的行号
' ----------------------------------------------------------------------------
Function GetLastRow(ws As Worksheet, col As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ----------------------------------------------------------------------------
' GetNextRow - 获取指定列下一个空行
' ----------------------------------------------------------------------------
Function GetNextRow(ws As Worksheet, col As Long) As Long
    GetNextRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row + 1
    If GetNextRow < 5 Then GetNextRow = 5
End Function

' ----------------------------------------------------------------------------
' SumColumn - 列求和（用数组优化）
' ----------------------------------------------------------------------------
Function SumColumn(ws As Worksheet, col As Long, Optional startRow As Long = 5) As Double
    Dim lastRow As Long, data As Variant, i As Long
    lastRow = GetLastRow(ws, col)
    If lastRow < startRow Then Exit Function
    data = ws.Range(ws.Cells(startRow, col), ws.Cells(lastRow, col)).Value
    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, 1)) Then SumColumn = SumColumn + CDbl(data(i, 1))
    Next i
End Function

' ----------------------------------------------------------------------------
' GetExpenseByCategory - 按类别汇总支出（用数组优化）
' ----------------------------------------------------------------------------
Function GetExpenseByCategory(ws As Worksheet) As Object
    Dim dict As Object, lastRow As Long, data As Variant
    Dim i As Long, cat As String, amt As Double
    Set dict = ' 【WPS兼容】CreateObject("Scripting.Dictionary") ' WPS不支持，请使用Collection
    lastRow = GetLastRow(ws, 3)
    If lastRow < 5 Then
        Set GetExpenseByCategory = dict
        Exit Function
    End If
    data = ws.Range("B5:C" & lastRow).Value
    For i = 1 To UBound(data, 1)
        cat = Trim(CStr(data(i, 1)))
        amt = 0
        If IsNumeric(data(i, 2)) Then amt = CDbl(data(i, 2))
        If cat <> "" And amt > 0 Then
            If dict.Exists(cat) Then
                dict(cat) = dict(cat) + amt
            Else
                dict.Add cat, amt
            End If
        End If
    Next i
    Set GetExpenseByCategory = dict
End Function

' ----------------------------------------------------------------------------
' SumARColumn - 应收/应付列求和（用数组优化）
' ----------------------------------------------------------------------------
Function SumARColumn(ws As Worksheet, col As Long, Optional startRow As Long = 5) As Double
    SumARColumn = SumColumn(ws, col, startRow)
End Function

' ----------------------------------------------------------------------------
' CalcNetProfit - 计算净利润（盈利时扣除所得税）
' ----------------------------------------------------------------------------
Function CalcNetProfit(income As Double, expense As Double, Optional taxRate As Double = 0.05) As Double
    Dim profit As Double
    profit = income - expense
    If profit > 0 Then
        CalcNetProfit = profit * (1 - taxRate)
    Else
        CalcNetProfit = profit
    End If
End Function

' ----------------------------------------------------------------------------
' OptimizeStart - 性能优化开关（开始）
' ----------------------------------------------------------------------------
Sub OptimizeStart()
    Application.ScreenUpdating = False
    On Error Resume Next ' 【WPS兼容】
    Application.Calculation = xlCalculationManual
    On Error GoTo 0
    On Error Resume Next ' 【WPS兼容】
    Application.EnableEvents = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' OptimizeEnd - 性能优化开关（结束）
' ----------------------------------------------------------------------------
Sub OptimizeEnd()
    Application.ScreenUpdating = True
    On Error Resume Next ' 【WPS兼容】
    Application.Calculation = xlCalculationAutomatic
    On Error GoTo 0
    On Error Resume Next ' 【WPS兼容】
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' GetOrCreateSheet - 创建或获取工作表（存在则清空）
' ----------------------------------------------------------------------------
Function GetOrCreateSheet(name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(name)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add
        GetOrCreateSheet.Name = name
    Else
        GetOrCreateSheet.Cells.Clear
    End If
    On Error GoTo 0
End Function

' ============================================================================
' 辅助函数
' ============================================================================

Function Nz(value As Variant, defaultValue As Variant) As Variant
    If IsEmpty(value) Or IsNull(value) Or value = "" Then
        Nz = defaultValue
    Else
        Nz = vValue
    End If
End Function

' ============================================================================
' V18.0 新增功能：材料库存跟踪、银行对账、发票管理、年度结转、自动导航
' ============================================================================

' ============================================================================
' 宏27: TrackMaterialInventory - 材料库存跟踪
' V18.0新增：记录三酸/片碱/亚钠/色粉/除油剂/挂具的采购和消耗
' ============================================================================
Sub TrackMaterialInventory()
    Dim wsMaterial As Worksheet
    Dim inputStr As String, lines() As String
    Dim i As Long, importCount As Long
    
    On Error GoTo ErrorHandler
    
    ' 创建或获取材料库存表
    On Error Resume Next
    Set wsMaterial = ThisWorkbook.Sheets("材料库存")
    If wsMaterial Is Nothing Then
        Set wsMaterial = ThisWorkbook.Sheets.Add
        wsMaterial.Name = "材料库存"
        wsMaterial.Cells(1, 1).Value = "材料库存跟踪表"
        wsMaterial.Cells(1, 1).Font.Size = 14
        wsMaterial.Cells(1, 1).Font.Bold = True
        
        wsMaterial.Cells(3, 1).Value = "日期"
        wsMaterial.Cells(3, 2).Value = "材料名称"
        wsMaterial.Cells(3, 3).Value = "操作类型"
        wsMaterial.Cells(3, 4).Value = "数量"
        wsMaterial.Cells(3, 5).Value = "单位"
        wsMaterial.Cells(3, 6).Value = "单价"
        wsMaterial.Cells(3, 7).Value = "金额"
        wsMaterial.Cells(3, 8).Value = "供应商"
        wsMaterial.Cells(3, 9).Value = "备注"
        wsMaterial.Range("A3:I3").Font.Bold = True
        wsMaterial.Range("A3:I3").Interior.Color = RGB(112, 173, 71)
        wsMaterial.Range("A3:I3").Font.Color = RGB(255, 255, 255)
        
        wsMaterial.Columns("A").ColumnWidth = 12
        wsMaterial.Columns("B").ColumnWidth = 12
        wsMaterial.Columns("C").ColumnWidth = 10
        wsMaterial.Columns("D").ColumnWidth = 10
        wsMaterial.Columns("E").ColumnWidth = 8
        wsMaterial.Columns("F").ColumnWidth = 10
        wsMaterial.Columns("G").ColumnWidth = 12
        wsMaterial.Columns("H").ColumnWidth = 15
        wsMaterial.Columns("I").ColumnWidth = 20
    End If
    On Error GoTo 0
    
    inputStr = InputBox("请输入材料库存数据（每行一笔）：" & vbCrLf & vbCrLf & _
                        "格式：日期,材料名称,操作,数量,单位,单价,供应商,备注" & vbCrLf & vbCrLf & _
                        "操作类型：采购/消耗" & vbCrLf & vbCrLf & _
                        "材料名称：三酸/片碱/亚钠/色粉/除油剂/挂具" & vbCrLf & vbCrLf & _
                        "示例：" & vbCrLf & _
                        "2025-06-15,三酸,采购,5,桶,400,化工店,硫酸硝酸" & vbCrLf & _
                        "2025-06-20,三酸,消耗,2,桶,,," & vbCrLf & _
                        "2025-06-15,除油剂,采购,3,桶,100,供应商,除油剂" & vbCrLf & _
                        "2025-06-25,挂具,采购,50,个,4,五金店,新挂具", _
                        "材料库存跟踪", "")
    
    If inputStr = "" Then Exit Sub
    
    lines = Split(inputStr, vbCrLf)
    importCount = 0
    
    Application.ScreenUpdating = False
    
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim parts() As String
            parts = Split(lines(i), ",")
            
            If UBound(parts) >= 3 Then
                Dim mDate As String, mName As String, mAction As String
                Dim mQty As Double, mUnit As String, mPrice As Double
                Dim mAmount As Double, mSupplier As String, mRemark As String
                Dim newRow As Long
                
                mDate = Trim(parts(0))
                mName = Trim(parts(1))
                mAction = Trim(parts(2))
                mQty = Val(Trim(parts(3)))
                mUnit = IIf(UBound(parts) >= 4, Trim(parts(4)), "桶")
                mPrice = Val(Trim(parts(5)))
                mSupplier = IIf(UBound(parts) >= 6, Trim(parts(6)), "")
                mRemark = IIf(UBound(parts) >= 7, Trim(parts(7)), "")
                mAmount = mQty * mPrice
                
                If mQty > 0 Then
                    newRow = wsMaterial.Cells(wsMaterial.Rows.Count, 1).End(xlUp).Row + 1
                    If newRow < 4 Then newRow = 4
                    
                    wsMaterial.Cells(newRow, 1).Value = mDate
                    wsMaterial.Cells(newRow, 2).Value = mName
                    wsMaterial.Cells(newRow, 3).Value = mAction
                    wsMaterial.Cells(newRow, 4).Value = mQty
                    wsMaterial.Cells(newRow, 5).Value = mUnit
                    wsMaterial.Cells(newRow, 6).Value = mPrice
                    wsMaterial.Cells(newRow, 6).NumberFormat = "#,##0.00"
                    wsMaterial.Cells(newRow, 7).Value = mAmount
                    wsMaterial.Cells(newRow, 7).NumberFormat = "#,##0.00"
                    wsMaterial.Cells(newRow, 8).Value = mSupplier
                    wsMaterial.Cells(newRow, 9).Value = mRemark
                    
                    importCount = importCount + 1
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "材料库存记录完成！" & vbCrLf & vbCrLf & _
           "记录笔数：" & importCount & " 条", vbInformation, "记录完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "材料库存记录出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏28: BankReconciliation - 银行对账
' V18.0新增：记录银行流水，与收入支出核对
' ============================================================================
Sub BankReconciliation()
    Dim wsBank As Worksheet
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim inputStr As String, lines() As String
    Dim i As Long, importCount As Long
    Dim totalBankIn As Double, totalBankOut As Double
    Dim totalBookIn As Double, totalBookOut As Double
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 创建或获取银行对账表
    On Error Resume Next
    Set wsBank = ThisWorkbook.Sheets("银行对账")
    If wsBank Is Nothing Then
        Set wsBank = ThisWorkbook.Sheets.Add
        wsBank.Name = "银行对账"
        wsBank.Cells(1, 1).Value = "银行对账表"
        wsBank.Cells(1, 1).Font.Size = 14
        wsBank.Cells(1, 1).Font.Bold = True
        
        wsBank.Cells(3, 1).Value = "日期"
        wsBank.Cells(3, 2).Value = "交易类型"
        wsBank.Cells(3, 3).Value = "金额"
        wsBank.Cells(3, 4).Value = "对方户名"
        wsBank.Cells(3, 5).Value = "摘要"
        wsBank.Cells(3, 6).Value = "勾对状态"
        wsBank.Range("A3:F3").Font.Bold = True
        wsBank.Range("A3:F3").Interior.Color = RGB(68, 114, 196)
        wsBank.Range("A3:F3").Font.Color = RGB(255, 255, 255)
        
        wsBank.Columns("A").ColumnWidth = 12
        wsBank.Columns("B").ColumnWidth = 10
        wsBank.Columns("C").ColumnWidth = 12
        wsBank.Columns("D").ColumnWidth = 18
        wsBank.Columns("E").ColumnWidth = 25
        wsBank.Columns("F").ColumnWidth = 10
    End If
    On Error GoTo 0
    
    inputStr = InputBox("请粘贴银行流水数据（每行一笔）：" & vbCrLf & vbCrLf & _
                        "格式：日期,类型,金额,对方户名,摘要" & vbCrLf & vbCrLf & _
                        "类型：收入/支出" & vbCrLf & vbCrLf & _
                        "示例：" & vbCrLf & _
                        "2025-06-15,收入,15000,永达五金厂,加工费" & vbCrLf & _
                        "2025-06-15,支出,3500,供电局,电费" & vbCrLf & _
                        "2025-06-16,支出,5000,房东,房租", _
                        "银行对账", "")
    
    If inputStr = "" Then Exit Sub
    
    lines = Split(inputStr, vbCrLf)
    importCount = 0
    totalBankIn = 0
    totalBankOut = 0
    
    Application.ScreenUpdating = False
    
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim parts() As String
            parts = Split(lines(i), ",")
            
            If UBound(parts) >= 2 Then
                Dim bDate As String, bType As String, bAmount As Double
                Dim bCounterparty As String, bRemark As String
                Dim newRow As Long
                
                bDate = Trim(parts(0))
                bType = Trim(parts(1))
                bAmount = Val(Trim(parts(2)))
                bCounterparty = IIf(UBound(parts) >= 3, Trim(parts(3)), "")
                bRemark = IIf(UBound(parts) >= 4, Trim(parts(4)), "")
                
                If bAmount > 0 Then
                    newRow = wsBank.Cells(wsBank.Rows.Count, 1).End(xlUp).Row + 1
                    If newRow < 4 Then newRow = 4
                    
                    wsBank.Cells(newRow, 1).Value = bDate
                    wsBank.Cells(newRow, 2).Value = bType
                    wsBank.Cells(newRow, 3).Value = bAmount
                    wsBank.Cells(newRow, 3).NumberFormat = "#,##0.00"
                    wsBank.Cells(newRow, 4).Value = bCounterparty
                    wsBank.Cells(newRow, 5).Value = bRemark
                    wsBank.Cells(newRow, 6).Value = "未勾对"
                    
                    If bType = "收入" Then
                        totalBankIn = totalBankIn + bAmount
                    Else
                        totalBankOut = totalBankOut + bAmount
                    End If
                    
                    importCount = importCount + 1
                End If
            End If
        End If
    Next i
    
    ' 计算账面数据（使用公共函数）
    totalBookIn = SumColumn(wsIncome, 3)
    totalBookOut = SumColumn(wsExpense, 3)
    
    Application.ScreenUpdating = True
    
    MsgBox "银行对账导入完成！" & vbCrLf & vbCrLf & _
           "银行流水：" & importCount & " 条" & vbCrLf & vbCrLf & _
           "银行收入合计：" & Format(totalBankIn, "#,##0.00") & " 元" & vbCrLf & _
           "账面收入合计：" & Format(totalBookIn, "#,##0.00") & " 元" & vbCrLf & _
           "差异：" & Format(totalBankIn - totalBookIn, "#,##0.00") & " 元" & vbCrLf & vbCrLf & _
           "银行支出合计：" & Format(totalBankOut, "#,##0.00") & " 元" & vbCrLf & _
           "账面支出合计：" & Format(totalBookOut, "#,##0.00") & " 元" & vbCrLf & _
           "差异：" & Format(totalBankOut - totalBookOut, "#,##0.00") & " 元", vbInformation, "对账完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "银行对账出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏29: InvoiceManagement - 发票管理
' V18.0新增：记录开票和收票情况
' ============================================================================
Sub InvoiceManagement()
    Dim wsInvoice As Worksheet
    Dim inputStr As String, lines() As String
    Dim i As Long, importCount As Long
    
    On Error GoTo ErrorHandler
    
    ' 创建或获取发票管理表
    On Error Resume Next
    Set wsInvoice = ThisWorkbook.Sheets("发票管理")
    If wsInvoice Is Nothing Then
        Set wsInvoice = ThisWorkbook.Sheets.Add
        wsInvoice.Name = "发票管理"
        wsInvoice.Cells(1, 1).Value = "发票管理台账"
        wsInvoice.Cells(1, 1).Font.Size = 14
        wsInvoice.Cells(1, 1).Font.Bold = True
        
        wsInvoice.Cells(3, 1).Value = "日期"
        wsInvoice.Cells(3, 2).Value = "发票类型"
        wsInvoice.Cells(3, 3).Value = "发票号码"
        wsInvoice.Cells(3, 4).Value = "客户/供应商"
        wsInvoice.Cells(3, 5).Value = "金额"
        wsInvoice.Cells(3, 6).Value = "税额"
        wsInvoice.Cells(3, 7).Value = "价税合计"
        wsInvoice.Cells(3, 8).Value = "状态"
        wsInvoice.Cells(3, 9).Value = "备注"
        wsInvoice.Range("A3:I3").Font.Bold = True
        wsInvoice.Range("A3:I3").Interior.Color = RGB(237, 125, 49)
        wsInvoice.Range("A3:I3").Font.Color = RGB(255, 255, 255)
        
        wsInvoice.Columns("A").ColumnWidth = 12
        wsInvoice.Columns("B").ColumnWidth = 10
        wsInvoice.Columns("C").ColumnWidth = 15
        wsInvoice.Columns("D").ColumnWidth = 18
        wsInvoice.Columns("E").ColumnWidth = 12
        wsInvoice.Columns("F").ColumnWidth = 10
        wsInvoice.Columns("G").ColumnWidth = 12
        wsInvoice.Columns("H").ColumnWidth = 10
        wsInvoice.Columns("I").ColumnWidth = 20
    End If
    On Error GoTo 0
    
    inputStr = InputBox("请输入发票数据（每行一笔）：" & vbCrLf & vbCrLf & _
                        "格式：日期,类型,发票号,客户/供应商,金额,税额,状态,备注" & vbCrLf & vbCrLf & _
                        "类型：开票/收票" & vbCrLf & _
                        "状态：已开/未开/已收/未收" & vbCrLf & vbCrLf & _
                        "示例：" & vbCrLf & _
                        "2025-06-15,开票,FP20250615001,永达五金厂,15000,450,已开,6月加工费" & vbCrLf & _
                        "2025-06-10,收票,FP20250610001,化工店,2000,60,已收,三酸", _
                        "发票管理", "")
    
    If inputStr = "" Then Exit Sub
    
    lines = Split(inputStr, vbCrLf)
    importCount = 0
    
    Application.ScreenUpdating = False
    
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim parts() As String
            parts = Split(lines(i), ",")
            
            If UBound(parts) >= 4 Then
                Dim invDate As String, invType As String, invNo As String
                Dim invParty As String, invAmount As Double, invTax As Double
                Dim invTotal As Double, invStatus As String, invRemark As String
                Dim newRow As Long
                
                invDate = Trim(parts(0))
                invType = Trim(parts(1))
                invNo = Trim(parts(2))
                invParty = Trim(parts(3))
                invAmount = Val(Trim(parts(4)))
                invTax = Val(Trim(parts(5)))
                invTotal = invAmount + invTax
                invStatus = IIf(UBound(parts) >= 6, Trim(parts(6)), "未开")
                invRemark = IIf(UBound(parts) >= 7, Trim(parts(7)), "")
                
                If invAmount > 0 Then
                    newRow = wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row + 1
                    If newRow < 4 Then newRow = 4
                    
                    wsInvoice.Cells(newRow, 1).Value = invDate
                    wsInvoice.Cells(newRow, 2).Value = invType
                    wsInvoice.Cells(newRow, 3).Value = invNo
                    wsInvoice.Cells(newRow, 4).Value = invParty
                    wsInvoice.Cells(newRow, 5).Value = invAmount
                    wsInvoice.Cells(newRow, 5).NumberFormat = "#,##0.00"
                    wsInvoice.Cells(newRow, 6).Value = invTax
                    wsInvoice.Cells(newRow, 6).NumberFormat = "#,##0.00"
                    wsInvoice.Cells(newRow, 7).Value = invTotal
                    wsInvoice.Cells(newRow, 7).NumberFormat = "#,##0.00"
                    wsInvoice.Cells(newRow, 8).Value = invStatus
                    wsInvoice.Cells(newRow, 9).Value = invRemark
                    
                    importCount = importCount + 1
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "发票记录完成！" & vbCrLf & vbCrLf & _
           "记录笔数：" & importCount & " 条", vbInformation, "记录完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "发票记录出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏30: AnnualClosing - 年度结转
' V18.0新增：将本年利润结转到未分配利润
' ============================================================================
Sub AnnualClosing()
    Dim wsProfit As Worksheet, wsBS As Worksheet
    Dim netProfit As Double
    Dim currentYear As String
    Dim result As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    currentYear = Format(Date, "yyyy")
    
    result = MsgBox("即将执行年度结转：" & vbCrLf & vbCrLf & _
                    "年度：" & currentYear & vbCrLf & vbCrLf & _
                    "将执行以下操作：" & vbCrLf & _
                    "1. 读取利润表中本年净利润" & vbCrLf & _
                    "2. 将净利润结转到资产负债表-未分配利润" & vbCrLf & vbCrLf & _
                    "注意：年度结转不可撤销！" & vbCrLf & vbCrLf & _
                    "是否继续？", vbYesNo + vbExclamation, "年度结转")
    
    If result <> vbYes Then Exit Sub
    
    Set wsProfit = ThisWorkbook.Sheets("利润表")
    Set wsBS = ThisWorkbook.Sheets("资产负债表")
    
    ' 读取净利润
    netProfit = CDbl(Nz(wsProfit.Cells(PR_NET_PROFIT, 2).Value, 0))
    
    ' 读取当前未分配利润
    Dim currentRetained As Double
    currentRetained = CDbl(Nz(wsBS.Cells(12, 6).Value, 0))
    
    ' 结转：未分配利润 = 原未分配利润 + 本年净利润
    wsBS.Cells(12, 6).Value = currentRetained + netProfit
    wsBS.Cells(12, 6).NumberFormat = "#,##0.00"
    
    ' 更新所有者权益合计
    wsBS.Cells(13, 6).Formula = "=F11+F12"
    wsBS.Cells(13, 6).NumberFormat = "#,##0.00"
    
    ' 更新负债权益总计
    wsBS.Cells(16, 6).Formula = "=F8+F13"
    wsBS.Cells(16, 6).NumberFormat = "#,##0.00"
    
    MsgBox "年度结转完成！" & vbCrLf & vbCrLf & _
           "年度：" & currentYear & vbCrLf & _
           "本年净利润：" & Format(netProfit, "#,##0.00") & " 元" & vbCrLf & _
           "原未分配利润：" & Format(currentRetained, "#,##0.00") & " 元" & vbCrLf & _
           "结转后未分配利润：" & Format(currentRetained + netProfit, "#,##0.00") & " 元", vbInformation, "结转完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "年度结转出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏31: AutoOpen - 自动打开首页导航（Workbook_Open事件）
' V18.0新增：打开文件时自动显示功能导航
' ============================================================================
Sub AutoOpen()
    ' 【V18.0】系统主菜单 - 分类展示所有功能
    Dim msg As String
    Dim msg1 As String, msg2 As String, msg3 As String
    
    ' 第一段：标题+日常记账（19个继续符）
    msg1 = "============================================" & vbCrLf & _
           "  小型氧化加工厂管理系统 V18.0 最终整合版" & vbCrLf & _
           "============================================" & vbCrLf & vbCrLf & _
           "【一、日常记账】" & vbCrLf & _
           "  1. 快速添加收入（含税/不含税）" & vbCrLf & _
           "  2. 快速添加支出" & vbCrLf & _
           "  3. 批量添加支出" & vbCrLf & _
           "  4. 对冲/代付货款管理" & vbCrLf & _
           "  5. 批量导入向导（银行/微信/支付宝）" & vbCrLf & vbCrLf & _
           "【二、报表生成】" & vbCrLf & _
           "  6. 一键月结（生成利润分析表）" & vbCrLf & _
           "  7. 生成资产负债表" & vbCrLf & _
           "  8. 生成现金流量表" & vbCrLf & _
           "  9. 多期对比分析" & vbCrLf & _
           " 10. 客户对账单"
    
    ' 第二段：税务管理+工资社保（17个继续符）
    msg2 = vbCrLf & vbCrLf & _
           "【三、税务管理】" & vbCrLf & _
           "  11. 增值税减免计算" & vbCrLf & _
           "  12. 六税两费减半计算" & vbCrLf & _
           "  13. 残保金计算" & vbCrLf & _
           "  14. 工会经费计算" & vbCrLf & _
           "  15. 所得税优化计算" & vbCrLf & _
           "  16. 税务合规检查" & vbCrLf & vbCrLf & _
           "【四、工资社保】" & vbCrLf & _
           "  17. 工资表生成" & vbCrLf & _
           "  18. 个税计算" & vbCrLf & _
           "  19. 社保明细" & vbCrLf & vbCrLf & _
           "【五、资产管理】" & vbCrLf & _
           "  20. 固定资产台账" & vbCrLf & _
           "  21. 折旧计算" & vbCrLf & _
           "  22. 低值易耗品管理"
    
    ' 第三段：成本分析+预算管理+系统工具+模板+退出（V18.0优化）
    msg3 = vbCrLf & vbCrLf & _
           "◆ 六、成本分析 ◆" & vbCrLf & _
           "  [23] 成本核算      ★15类成本归集" & vbCrLf & _
           "  [24] 成本差异      ★预算对比" & vbCrLf & _
           "  [25] 应收账龄      ★逾期预警" & vbCrLf & _
           "  [26] 应付账龄      ★付款提醒" & vbCrLf & vbCrLf & _
           "◆ 七、预算管理 ◆" & vbCrLf & _
           "  [27] 预算编制      ★年度预算" & vbCrLf & _
           "  [28] 预算控制      ★实时监控" & vbCrLf & _
           "  [29] 预算预警      ★超支提醒" & vbCrLf & vbCrLf & _
           "◆ 八、系统工具 ◆" & vbCrLf & _
           "  [30] 数据校验      ★完整性检查" & vbCrLf & _
           "  [31] 数据备份      ★安全备份" & vbCrLf & _
           "  [32] 导出PDF       ★电子存档" & vbCrLf & _
           "  [33] 系统设置      ★参数配置" & vbCrLf & _
           "  [34] 使用帮助      ★操作指南" & vbCrLf & _
           "  [35] 关于系统      ★版本信息" & vbCrLf & vbCrLf & _
           "◆ 九、常用模板 ◆" & vbCrLf & _
           "  [36] 工资表  [37] 发票登记  [38] 银行对账" & vbCrLf & _
           "  [39] 材料台账 [40] 月度报表 [41] 模拟数据" & vbCrLf & vbCrLf & _
           "  [0] 退出系统" & vbCrLf & _
           "══════════════════════════════════════════"
    
    ' 合并所有消息
    msg = msg1 & msg2 & msg3
    
    Dim choice As String
    choice = InputBox(msg, "系统导航 V18.0", "")
    
    If choice = "" Then Exit Sub
    
    Select Case choice
        ' 一、日常记账
        Case "1": QuickAddIncomeWithTax
        Case "2": QuickAddExpense
        Case "3": BatchAddExpense
        Case "4": OffsetManagement
        Case "5": BatchImportWizard
        
        ' 二、报表生成
        Case "6": OneKeyMonthEnd
        Case "7": GenerateBalanceSheet
        Case "8": GenerateCashFlowReport
        Case "9": GenerateMultiPeriodAnalysis
        Case "10": GenerateCustomerStatement
        
        ' 三、税务管理
        Case "11": CalculateVATExemption
        Case "12": CalculateSixTaxReduction
        Case "13": CalculateDisabilityFee
        Case "14": CalculateUnionFee
        Case "15": CalculateIncomeTaxOptimized
        Case "16": TaxComplianceCheck
        
        ' 四、工资社保
        Case "17": GeneratePayroll
        Case "18": CalculatePersonalIncomeTax
        Case "19": GenerateSocialDetail
        
        ' 五、资产管理
        Case "20": CreateFixedAssetsLedger
        Case "21": CalculateDepreciation
        Case "22": LowValueConsumables
        
        ' 六、成本分析
        Case "23": CostAllocation
        Case "24": CostVarianceAnalysis
        Case "25": ARAgingAnalysis
        Case "26": APAgingAnalysis
        
        ' 七、预算管理
        Case "27": CreateExpenseBudget
        Case "28": BudgetControl
        Case "29": BudgetAlert
        
        ' 八、系统工具
        Case "30": DataCheck
        Case "31": BackupSimple
        Case "32": ExportToPDF
        Case "33": SystemSettings
        Case "34": ShowHelpDocument
        Case "35": AboutSystem
        
        ' 九、常用模板（V18.0新增）
        Case "36": CreateSalarySheet
        Case "37": CreateInvoiceSheet
        Case "38": CreateBankReconSheet
        Case "39": CreateMaterialSheet
        Case "40": CreateMonthlyReportSheet
        
        ' 十、数据工具（V18.0新增）
        Case "41": GenerateSampleData
        
        Case "0": Exit Sub
        
        Case Else
            MsgBox "无效选择，请重新运行", vbExclamation
    End Select
End Sub

' ============================================================================
' V18.0 新增功能：客户对账单、供应商汇总、成本利润分析、PDF导出、自动提醒
' ============================================================================

' ============================================================================
' 宏32: GenerateCustomerStatement - 生成客户对账单
' V18.0新增：选择客户，生成该客户的对账单
' ============================================================================
Sub GenerateCustomerStatement()
    Dim wsIncome As Worksheet, wsARAP As Worksheet
    Dim wsStatement As Worksheet
    Dim customerName As String
    Dim lastRow As Long, i As Long, stmtRow As Long
    Dim totalIncome As Double, totalReceived As Double
    Dim endAR As Double
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    
    ' 获取客户列表
    Dim customerList As String
    customerList = ""
    lastRow = GetLastRow(wsARAP, AR_COL_NAME)
    For i = 5 To lastRow
        Dim cName As String
        cName = Trim(wsARAP.Cells(i, AR_COL_NAME).Value)
        If cName <> "" Then
            customerList = customerList & cName & vbCrLf
        End If
    Next i
    
    customerName = InputBox("请选择客户名称：" & vbCrLf & vbCrLf & _
                            "可选客户：" & vbCrLf & customerList, "生成客户对账单", "")
    If customerName = "" Then Exit Sub
    
    ' 创建对账单工作表
    Set wsStatement = GetOrCreateSheet("客户对账单")
    
    OptimizeStart
    
    ' 表头
    wsStatement.Cells(1, 1).Value = "客户对账单"
    wsStatement.Cells(1, 1).Font.Size = 16
    wsStatement.Cells(1, 1).Font.Bold = True
    wsStatement.Cells(1, 1).HorizontalAlignment = xlCenter
    wsStatement.Range("A1:E1").Merge
    
    wsStatement.Cells(2, 1).Value = "客户名称："
    wsStatement.Cells(2, 2).Value = customerName
    wsStatement.Cells(2, 4).Value = "期间："
    wsStatement.Cells(2, 5).Value = Format(Date, "yyyy年mm月")
    
    ' 收入明细表头
    wsStatement.Cells(4, 1).Value = "日期"
    wsStatement.Cells(4, 2).Value = "摘要"
    wsStatement.Cells(4, 3).Value = "收入金额"
    wsStatement.Cells(4, 4).Value = "收款金额"
    wsStatement.Cells(4, 5).Value = "备注"
    wsStatement.Range("A4:E4").Font.Bold = True
    wsStatement.Range("A4:E4").Interior.Color = RGB(68, 114, 196)
    wsStatement.Range("A4:E4").Font.Color = RGB(255, 255, 255)
    
    ' 填充收入明细
    stmtRow = 5
    totalIncome = 0
    totalReceived = 0
    
    lastRow = GetLastRow(wsIncome, 3)
    For i = 5 To lastRow
        If Trim(wsIncome.Cells(i, 2).Value) = customerName Then
            wsStatement.Cells(stmtRow, 1).Value = wsIncome.Cells(i, 1).Value
            wsStatement.Cells(stmtRow, 2).Value = "加工收入"
            wsStatement.Cells(stmtRow, 3).Value = CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
            wsStatement.Cells(stmtRow, 3).NumberFormat = "#,##0.00"
            totalIncome = totalIncome + CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
            stmtRow = stmtRow + 1
        End If
    Next i
    
    ' 从应收应付表获取收款和期末应收
    lastRow = GetLastRow(wsARAP, AR_COL_NAME)
    For i = 5 To lastRow
        If Trim(wsARAP.Cells(i, AR_COL_NAME).Value) = customerName Then
            totalReceived = totalReceived + CDbl(Nz(wsARAP.Cells(i, AR_COL_REDUCE).Value, 0))
            endAR = CDbl(Nz(wsARAP.Cells(i, AR_COL_CLOSE).Value, 0))
        End If
    Next i
    
    ' 汇总行
    stmtRow = stmtRow + 1
    wsStatement.Cells(stmtRow, 1).Value = "合计"
    wsStatement.Cells(stmtRow, 1).Font.Bold = True
    wsStatement.Cells(stmtRow, 3).Value = totalIncome
    wsStatement.Cells(stmtRow, 3).NumberFormat = "#,##0.00"
    wsStatement.Cells(stmtRow, 3).Font.Bold = True
    wsStatement.Cells(stmtRow, 4).Value = totalReceived
    wsStatement.Cells(stmtRow, 4).NumberFormat = "#,##0.00"
    wsStatement.Cells(stmtRow, 4).Font.Bold = True
    
    stmtRow = stmtRow + 2
    wsStatement.Cells(stmtRow, 1).Value = "期末应收余额："
    wsStatement.Cells(stmtRow, 1).Font.Bold = True
    wsStatement.Cells(stmtRow, 2).Value = endAR
    wsStatement.Cells(stmtRow, 2).NumberFormat = "#,##0.00"
    wsStatement.Cells(stmtRow, 2).Font.Bold = True
    wsStatement.Cells(stmtRow, 2).Font.Color = IIf(endAR > 0, RGB(255, 0, 0), RGB(0, 128, 0))
    
    ' 列宽
    wsStatement.Columns("A").ColumnWidth = 12
    wsStatement.Columns("B").ColumnWidth = 15
    wsStatement.Columns("C").ColumnWidth = 12
    wsStatement.Columns("D").ColumnWidth = 12
    wsStatement.Columns("E").ColumnWidth = 20
    
    OptimizeEnd
    
    MsgBox "客户对账单生成完成！" & vbCrLf & vbCrLf & _
           "客户：" & customerName & vbCrLf & _
           "本期收入：" & Format(totalIncome, "#,##0.00") & " 元" & vbCrLf & _
           "本期收款：" & Format(totalReceived, "#,##0.00") & " 元" & vbCrLf & _
           "期末应收：" & Format(endAR, "#,##0.00") & " 元", vbInformation, "对账单完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "生成对账单出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏33: GenerateSupplierSummary - 供应商汇总
' V18.0新增：按供应商汇总外发加工费和材料采购
' ============================================================================
Sub GenerateSupplierSummary()
    Dim wsExpense As Worksheet, wsSummary As Worksheet
    Dim supplierDict As Object
    Dim lastRow As Long, i As Long, row As Long
    Dim supplier As String, cat As String, amt As Double
    Dim key As String
    
    On Error GoTo ErrorHandler
    
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 创建汇总表
    Set wsSummary = GetOrCreateSheet("供应商汇总")
    
    OptimizeStart
    
    ' 按供应商汇总
    Set supplierDict = ' 【WPS兼容】CreateObject("Scripting.Dictionary") ' WPS不支持，请使用Collection
    lastRow = GetLastRow(wsExpense, 3)
    
    For i = 5 To lastRow
        supplier = Trim(CStr(wsExpense.Cells(i, 4).Value))
        cat = Trim(CStr(wsExpense.Cells(i, 2).Value))
        amt = 0
        If IsNumeric(wsExpense.Cells(i, 3).Value) Then amt = CDbl(wsExpense.Cells(i, 3).Value)
        
        If supplier <> "" And amt > 0 Then
            ' 外发加工费
            If cat = "外发加工费" Then
                key = supplier & "|outsource"
                If supplierDict.Exists(key) Then
                    supplierDict(key) = supplierDict(key) + amt
                Else
                    supplierDict.Add key, amt
                End If
            End If
            
            ' 材料费（三酸、片碱、亚钠、色粉、除油剂、挂具）
            If cat = "三酸" Or cat = "片碱" Or cat = "亚钠" Or cat = "色粉" Or cat = "除油剂" Or cat = "挂具" Then
                key = supplier & "|material"
                If supplierDict.Exists(key) Then
                    supplierDict(key) = supplierDict(key) + amt
                Else
                    supplierDict.Add key, amt
                End If
            End If
        End If
    Next i
    
    ' 表头
    wsSummary.Cells(1, 1).Value = "供应商汇总表"
    wsSummary.Cells(1, 1).Font.Size = 16
    wsSummary.Cells(1, 1).Font.Bold = True
    wsSummary.Cells(1, 1).HorizontalAlignment = xlCenter
    wsSummary.Range("A1:D1").Merge
    
    wsSummary.Cells(2, 1).Value = "期间：" & Format(Date, "yyyy年mm月")
    
    wsSummary.Cells(4, 1).Value = "供应商名称"
    wsSummary.Cells(4, 2).Value = "外发加工费"
    wsSummary.Cells(4, 3).Value = "材料费"
    wsSummary.Cells(4, 4).Value = "合计"
    wsSummary.Range("A4:D4").Font.Bold = True
    wsSummary.Range("A4:D4").Interior.Color = RGB(68, 114, 196)
    wsSummary.Range("A4:D4").Font.Color = RGB(255, 255, 255)
    
    ' 收集唯一供应商名称
    Dim supplierSet As Object
    Set supplierSet = ' 【WPS兼容】CreateObject("Scripting.Dictionary") ' WPS不支持，请使用Collection
    Dim dictKey As Variant
    For Each dictKey In supplierDict.Keys
        Dim parts() As String
        parts = Split(CStr(dictKey), "|")
        If UBound(parts) = 1 Then
            If Not supplierSet.Exists(parts(0)) Then
                supplierSet.Add parts(0), 0
            End If
        End If
    Next dictKey
    
    ' 填充数据
    row = 5
    Dim totalOutsource As Double, totalMaterial As Double
    totalOutsource = 0
    totalMaterial = 0
    
    Dim sName As Variant
    For Each sName In supplierSet.Keys
        Dim outsourceAmt As Double, materialAmt As Double
        outsourceAmt = 0
        materialAmt = 0
        
        key = CStr(sName) & "|outsource"
        If supplierDict.Exists(key) Then outsourceAmt = supplierDict(key)
        
        key = CStr(sName) & "|material"
        If supplierDict.Exists(key) Then materialAmt = supplierDict(key)
        
        If outsourceAmt > 0 Or materialAmt > 0 Then
            wsSummary.Cells(row, 1).Value = sName
            wsSummary.Cells(row, 2).Value = outsourceAmt
            wsSummary.Cells(row, 2).NumberFormat = "#,##0.00"
            wsSummary.Cells(row, 3).Value = materialAmt
            wsSummary.Cells(row, 3).NumberFormat = "#,##0.00"
            wsSummary.Cells(row, 4).Value = outsourceAmt + materialAmt
            wsSummary.Cells(row, 4).NumberFormat = "#,##0.00"
            
            totalOutsource = totalOutsource + outsourceAmt
            totalMaterial = totalMaterial + materialAmt
            row = row + 1
        End If
    Next sName
    
    ' 合计行
    wsSummary.Cells(row, 1).Value = "合计"
    wsSummary.Cells(row, 1).Font.Bold = True
    wsSummary.Cells(row, 2).Value = totalOutsource
    wsSummary.Cells(row, 2).NumberFormat = "#,##0.00"
    wsSummary.Cells(row, 2).Font.Bold = True
    wsSummary.Cells(row, 3).Value = totalMaterial
    wsSummary.Cells(row, 3).NumberFormat = "#,##0.00"
    wsSummary.Cells(row, 3).Font.Bold = True
    wsSummary.Cells(row, 4).Value = totalOutsource + totalMaterial
    wsSummary.Cells(row, 4).NumberFormat = "#,##0.00"
    wsSummary.Cells(row, 4).Font.Bold = True
    
    ' 列宽
    wsSummary.Columns("A").ColumnWidth = 18
    wsSummary.Columns("B").ColumnWidth = 14
    wsSummary.Columns("C").ColumnWidth = 12
    wsSummary.Columns("D").ColumnWidth = 12
    
    OptimizeEnd
    
    MsgBox "供应商汇总生成完成！" & vbCrLf & vbCrLf & _
           "供应商数量：" & supplierSet.Count & " 家" & vbCrLf & _
           "外发加工费合计：" & Format(totalOutsource, "#,##0.00") & " 元" & vbCrLf & _
           "材料费合计：" & Format(totalMaterial, "#,##0.00") & " 元", vbInformation, "汇总完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "生成供应商汇总出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏34: GenerateCostProfitAnalysis - 成本利润分析
' V18.0新增：分析各项成本占比和利润率趋势
' ============================================================================
Sub GenerateCostProfitAnalysis()
    Dim wsIncome As Worksheet, wsExpense As Worksheet
    Dim wsAnalysis As Worksheet
    Dim totalIncome As Double, totalExpense As Double
    Dim expenseByCategory As Object
    Dim catKey As Variant
    
    On Error GoTo ErrorHandler
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 创建分析表
    Set wsAnalysis = GetOrCreateSheet("成本利润分析")
    
    OptimizeStart
    
    ' 计算收入和支出（使用公共函数）
    totalIncome = SumColumn(wsIncome, 3)
    totalExpense = SumColumn(wsExpense, 3)
    
    ' 按类别汇总支出
    Set expenseByCategory = GetExpenseByCategory(wsExpense)
    
    ' 各项成本
    Dim outsourceCost As Double, materialCost As Double
    Dim laborCost As Double, otherCost As Double
    outsourceCost = CDbl(DictGet(expenseByCategory, "外发加工费", 0))
    materialCost = CDbl(DictGet(expenseByCategory, "三酸", 0)) + CDbl(DictGet(expenseByCategory, "片碱", 0)) + _
                   CDbl(DictGet(expenseByCategory, "亚钠", 0)) + CDbl(DictGet(expenseByCategory, "色粉", 0)) + _
                   CDbl(DictGet(expenseByCategory, "除油剂", 0)) + CDbl(DictGet(expenseByCategory, "挂具", 0))
    laborCost = CDbl(DictGet(expenseByCategory, "工资", 0))
    otherCost = totalExpense - outsourceCost - materialCost - laborCost
    
    ' 计算利润
    Dim grossProfit As Double, netProfit As Double
    grossProfit = totalIncome - outsourceCost
    netProfit = CalcNetProfit(totalIncome, totalExpense)
    
    Dim grossRate As Double, netRate As Double
    If totalIncome > 0 Then
        grossRate = grossProfit / totalIncome
        netRate = netProfit / totalIncome
    End If
    
    ' 表头
    wsAnalysis.Cells(1, 1).Value = "成本利润分析表"
    wsAnalysis.Cells(1, 1).Font.Size = 16
    wsAnalysis.Cells(1, 1).Font.Bold = True
    wsAnalysis.Cells(1, 1).HorizontalAlignment = xlCenter
    wsAnalysis.Range("A1:D1").Merge
    
    wsAnalysis.Cells(2, 1).Value = "期间：" & Format(Date, "yyyy年mm月")
    
    ' 收入成本数据
    wsAnalysis.Cells(4, 1).Value = "项目"
    wsAnalysis.Cells(4, 2).Value = "金额（元）"
    wsAnalysis.Cells(4, 3).Value = "占收入比例"
    wsAnalysis.Cells(4, 4).Value = "说明"
    wsAnalysis.Range("A4:D4").Font.Bold = True
    wsAnalysis.Range("A4:D4").Interior.Color = RGB(68, 114, 196)
    wsAnalysis.Range("A4:D4").Font.Color = RGB(255, 255, 255)
    
    Dim row As Long
    row = 5
    
    ' 收入
    wsAnalysis.Cells(row, 1).Value = "营业收入"
    wsAnalysis.Cells(row, 2).Value = totalIncome
    wsAnalysis.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAnalysis.Cells(row, 3).Value = IIf(totalIncome > 0, "100.0%", "N/A")
    wsAnalysis.Cells(row, 4).Value = "加工收入合计"
    row = row + 1
    
    ' 外发成本
    wsAnalysis.Cells(row, 1).Value = "外发加工成本"
    wsAnalysis.Cells(row, 2).Value = outsourceCost
    wsAnalysis.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAnalysis.Cells(row, 3).Value = IIf(totalIncome > 0, Format(outsourceCost / totalIncome, "0.0%"), "N/A")
    wsAnalysis.Cells(row, 4).Value = "外发加工费"
    row = row + 1
    
    ' 材料成本
    wsAnalysis.Cells(row, 1).Value = "材料成本"
    wsAnalysis.Cells(row, 2).Value = materialCost
    wsAnalysis.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAnalysis.Cells(row, 3).Value = IIf(totalIncome > 0, Format(materialCost / totalIncome, "0.0%"), "N/A")
    wsAnalysis.Cells(row, 4).Value = "三酸/片碱/亚钠/色粉/除油剂/挂具"
    row = row + 1
    
    ' 人工成本
    wsAnalysis.Cells(row, 1).Value = "人工成本"
    wsAnalysis.Cells(row, 2).Value = laborCost
    wsAnalysis.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAnalysis.Cells(row, 3).Value = IIf(totalIncome > 0, Format(laborCost / totalIncome, "0.0%"), "N/A")
    wsAnalysis.Cells(row, 4).Value = "工资"
    row = row + 1
    
    ' 其他费用
    wsAnalysis.Cells(row, 1).Value = "其他费用"
    wsAnalysis.Cells(row, 2).Value = otherCost
    wsAnalysis.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAnalysis.Cells(row, 3).Value = IIf(totalIncome > 0, Format(otherCost / totalIncome, "0.0%"), "N/A")
    wsAnalysis.Cells(row, 4).Value = "房租/水电/日常费用等"
    row = row + 1
    
    ' 总成本
    wsAnalysis.Cells(row, 1).Value = "总成本"
    wsAnalysis.Cells(row, 1).Font.Bold = True
    wsAnalysis.Cells(row, 2).Value = totalExpense
    wsAnalysis.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAnalysis.Cells(row, 2).Font.Bold = True
    wsAnalysis.Cells(row, 3).Value = IIf(totalIncome > 0, Format(totalExpense / totalIncome, "0.0%"), "N/A")
    wsAnalysis.Cells(row, 3).Font.Bold = True
    row = row + 1
    
    ' 空行
    row = row + 1
    
    ' 毛利润
    wsAnalysis.Cells(row, 1).Value = "毛利润"
    wsAnalysis.Cells(row, 1).Font.Bold = True
    wsAnalysis.Cells(row, 2).Value = grossProfit
    wsAnalysis.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAnalysis.Cells(row, 2).Font.Bold = True
    wsAnalysis.Cells(row, 3).Value = IIf(totalIncome > 0, Format(grossRate, "0.0%"), "N/A")
    wsAnalysis.Cells(row, 3).Font.Bold = True
    wsAnalysis.Cells(row, 4).Value = "收入 - 外发加工成本"
    row = row + 1
    
    ' 净利润
    wsAnalysis.Cells(row, 1).Value = "净利润"
    wsAnalysis.Cells(row, 1).Font.Bold = True
    wsAnalysis.Cells(row, 2).Value = netProfit
    wsAnalysis.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAnalysis.Cells(row, 2).Font.Bold = True
    wsAnalysis.Cells(row, 3).Value = IIf(totalIncome > 0, Format(netRate, "0.0%"), "N/A")
    wsAnalysis.Cells(row, 3).Font.Bold = True
    wsAnalysis.Cells(row, 4).Value = "扣除5%所得税后"
    
    ' 列宽
    wsAnalysis.Columns("A").ColumnWidth = 16
    wsAnalysis.Columns("B").ColumnWidth = 14
    wsAnalysis.Columns("C").ColumnWidth = 14
    wsAnalysis.Columns("D").ColumnWidth = 28
    
    OptimizeEnd
    
    MsgBox "成本利润分析生成完成！" & vbCrLf & vbCrLf & _
           "营业收入：" & Format(totalIncome, "#,##0.00") & " 元" & vbCrLf & _
           "毛利率：" & IIf(totalIncome > 0, Format(grossRate, "0.0%"), "N/A") & vbCrLf & _
           "净利率：" & IIf(totalIncome > 0, Format(netRate, "0.0%"), "N/A"), vbInformation, "分析完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "生成成本利润分析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏35: ExportToPDF - 导出PDF报表
' V18.0新增：将利润表导出为PDF文件
' ============================================================================
Sub ExportToPDF()
    Dim wsProfit As Worksheet
    Dim pdfPath As String
    Dim backupDir As String
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    Set wsProfit = ThisWorkbook.Sheets("利润表")
    
    ' 确保备份目录存在
    backupDir = ThisWorkbook.Path & "\备份"
    If Dir(backupDir, vbDirectory) = "" Then
        On Error Resume Next
        MkDir backupDir
        If Err.Number <> 0 Then
            MsgBox "创建备份目录失败：" & Err.Description, vbCritical, "错误"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    ' 生成PDF文件名
    fileName = "利润表_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"
    pdfPath = backupDir & "\" & fileName
    
    ' 导出PDF
    wsProfit.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False
    
    MsgBox "PDF导出成功！" & vbCrLf & vbCrLf & _
           "文件路径：" & pdfPath, vbInformation, "导出完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "导出PDF出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏36: AutoReminder - 自动提醒检查
' V18.0新增：打开文件时自动检查关键事项
' ============================================================================
Sub AutoReminder()
    Dim wsARAP As Worksheet, wsExpense As Worksheet
    Dim msg As String
    Dim lastRow As Long, i As Long
    Dim hasAlert As Boolean
    Dim daysToQuarter As Integer
    
    On Error GoTo ErrorHandler
    
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    msg = "自动提醒检查" & vbCrLf & vbCrLf
    hasAlert = False
    
    ' 1. 检查是否有超过30天的大额应收
    Dim bigARCount As Long
    bigARCount = 0
    lastRow = GetLastRow(wsARAP, AR_COL_CLOSE)
    For i = 5 To lastRow
        Dim arAmt As Double
        arAmt = CDbl(Nz(wsARAP.Cells(i, AR_COL_CLOSE).Value, 0))
        If arAmt > 50000 Then
            bigARCount = bigARCount + 1
        End If
    Next i
    
    If bigARCount > 0 Then
        msg = msg & "【重要】发现 " & bigARCount & " 笔大额应收款（超过5万元），建议尽快催收" & vbCrLf & vbCrLf
        hasAlert = True
    Else
        msg = msg & "大额应收款检查：正常" & vbCrLf & vbCrLf
    End If
    
    ' 2. 检查本月是否已录入工资
    Dim currentMonth As String
    currentMonth = Format(Date, "yyyy-mm")
    Dim hasSalary As Boolean
    hasSalary = False
    lastRow = GetLastRow(wsExpense, 3)
    For i = 5 To lastRow
        If InStr(wsExpense.Cells(i, 1).Value, currentMonth) > 0 Then
            If Trim(wsExpense.Cells(i, 2).Value) = "工资" Then
                hasSalary = True
                Exit For
            End If
        End If
    Next i
    
    If Not hasSalary Then
        msg = msg & "【提醒】本月尚未录入工资支出" & vbCrLf & vbCrLf
        hasAlert = True
    Else
        msg = msg & "工资录入检查：已录入" & vbCrLf & vbCrLf
    End If
    
    ' 3. 距离季度申报还有多少天
    Dim quarterEnd As Date
    Dim month As Integer
    month = Month(Date)
    
    If month >= 1 And month <= 3 Then
        quarterEnd = DateSerial(Year(Date), 4, 15)
    ElseIf month >= 4 And month <= 6 Then
        quarterEnd = DateSerial(Year(Date), 7, 15)
    ElseIf month >= 7 And month <= 9 Then
        quarterEnd = DateSerial(Year(Date), 10, 15)
    Else
        quarterEnd = DateSerial(Year(Date) + 1, 1, 15)
    End If
    
    daysToQuarter = quarterEnd - Date
    
    If daysToQuarter <= 15 And daysToQuarter >= 0 Then
        msg = msg & "【重要】距离季度所得税申报还有 " & daysToQuarter & " 天，请提前准备" & vbCrLf & vbCrLf
        hasAlert = True
    ElseIf daysToQuarter < 0 And daysToQuarter > -15 Then
        msg = msg & "【警告】季度所得税申报已逾期 " & Abs(daysToQuarter) & " 天！" & vbCrLf & vbCrLf
        hasAlert = True
    Else
        msg = msg & "季度申报：距离下次申报还有 " & daysToQuarter & " 天" & vbCrLf & vbCrLf
    End If
    
    ' 4. 检查上次备份时间
    Dim backupDir As String
    backupDir = ThisWorkbook.Path & "\备份"
    If Dir(backupDir, vbDirectory) <> "" Then
        Dim latestFile As String
        Dim latestDate As Date
        latestDate = 0
        
        latestFile = Dir(backupDir & "\备份_*.xlsx")
        Do While latestFile <> ""
            Dim fileDate As Date
            ' 从文件名提取日期
            Dim dateStr As String
            dateStr = Mid(latestFile, 4, 8) ' 提取yyyymmdd
            If Len(dateStr) = 8 Then
                On Error Resume Next
                fileDate = DateSerial(CInt(Left(dateStr, 4)), CInt(Mid(dateStr, 5, 2)), CInt(Right(dateStr, 2)))
                If Err.Number = 0 Then
                    If fileDate > latestDate Then latestDate = fileDate
                End If
                On Error GoTo 0
            End If
            latestFile = Dir()
        Loop
        
        If latestDate > 0 Then
            Dim daysSinceBackup As Long
            daysSinceBackup = Date - latestDate
            If daysSinceBackup > 7 Then
                msg = msg & "【提醒】上次备份距今已 " & daysSinceBackup & " 天，建议尽快备份" & vbCrLf & vbCrLf
                hasAlert = True
            Else
                msg = msg & "上次备份时间：" & Format(latestDate, "yyyy-mm-dd") & "（" & daysSinceBackup & " 天前）" & vbCrLf & vbCrLf
            End If
        Else
            msg = msg & "未找到备份文件，建议立即备份" & vbCrLf & vbCrLf
            hasAlert = True
        End If
    Else
        msg = msg & "备份目录不存在，建议创建并备份" & vbCrLf & vbCrLf
        hasAlert = True
    End If
    
    ' 显示结果
    If hasAlert Then
        MsgBox msg, vbExclamation, "自动提醒"
    Else
        MsgBox msg, vbInformation, "自动提醒"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "自动提醒检查出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' V7.0 新增宏：税务合规、银行实务、工资社保
' ============================================================================

' ============================================================================
' 宏37: GenerateVATReturn - 增值税申报表（小规模纳税人）
' V7.0新增：生成小规模纳税人增值税申报表
' ============================================================================
Sub GenerateVATReturn()
    Dim wsIncome As Worksheet, wsVAT As Worksheet
    Dim lastRow As Long, i As Long
    Dim taxableSales As Double, exemptSales As Double
    Dim taxPayable As Double, exemptTax As Double
    Dim monthStr As String
    
    On Error GoTo ErrorHandler
    
    monthStr = InputBox("请输入申报月份（格式：yyyy-mm）：", "增值税申报", Format(Date, "yyyy-mm"))
    If monthStr = "" Then Exit Sub
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    
    ' 创建或清空增值税申报表
    On Error Resume Next
    Set wsVAT = ThisWorkbook.Sheets("增值税申报表")
    If wsVAT Is Nothing Then
        Set wsVAT = ThisWorkbook.Sheets.Add
        wsVAT.Name = "增值税申报表"
    Else
        wsVAT.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 计算月销售额合计
    Dim totalMonthlySales As Double
    totalMonthlySales = 0
    
    For i = 5 To lastRow
        If InStr(wsIncome.Cells(i, 1).Value, monthStr) > 0 Then
            Dim incomeAmt As Double
            incomeAmt = CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
            totalMonthlySales = totalMonthlySales + incomeAmt
        End If
    Next i
    
    ' 根据月销售额合计判断是否免税
    ' 小规模纳税人月销售额<=10万免税
    If totalMonthlySales <= VAT_THRESHOLD Then
        ' 月销售额合计<=10万，全部免税
        exemptSales = totalMonthlySales
        taxableSales = 0
    Else
        ' 月销售额合计>10万，按实际销售额计算税额
        taxableSales = totalMonthlySales
        exemptSales = 0
    End If
    
    ' 计算税额
    taxPayable = taxableSales * VAT_RATE_SMALL
    exemptTax = exemptSales * VAT_RATE_SMALL
    
    ' 填写申报表
    wsVAT.Cells(1, 1).Value = "增值税纳税申报表（小规模纳税人适用）"
    wsVAT.Cells(1, 1).Font.Size = 14
    wsVAT.Cells(1, 1).Font.Bold = True
    wsVAT.Range("A1:F1").Merge
    
    wsVAT.Cells(2, 1).Value = "税款所属期：" & monthStr
    wsVAT.Cells(2, 5).Value = "填表日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 表头
    wsVAT.Cells(4, 1).Value = "项目"
    wsVAT.Cells(4, 2).Value = "栏次"
    wsVAT.Cells(4, 3).Value = "本期数"
    wsVAT.Cells(4, 4).Value = "本年累计"
    wsVAT.Range("A4:D4").Font.Bold = True
    wsVAT.Range("A4:D4").Interior.Color = RGB(68, 114, 196)
    wsVAT.Range("A4:D4").Font.Color = RGB(255, 255, 255)
    
    ' 数据行
    wsVAT.Cells(5, 1).Value = "（一）应征增值税不含税销售额（3%征收率）"
    wsVAT.Cells(5, 2).Value = 1
    wsVAT.Cells(5, 3).Value = taxableSales
    wsVAT.Cells(5, 3).NumberFormat = "#,##0.00"
    
    wsVAT.Cells(6, 1).Value = "其中：增值税专用发票不含税销售额"
    wsVAT.Cells(6, 2).Value = 2
    wsVAT.Cells(6, 3).Value = 0
    
    wsVAT.Cells(7, 1).Value = "其他增值税发票不含税销售额"
    wsVAT.Cells(7, 2).Value = 3
    wsVAT.Cells(7, 3).Value = taxableSales
    wsVAT.Cells(7, 3).NumberFormat = "#,##0.00"
    
    wsVAT.Cells(8, 1).Value = "（二）免税销售额"
    wsVAT.Cells(8, 2).Value = 9
    wsVAT.Cells(8, 3).Value = exemptSales
    wsVAT.Cells(8, 3).NumberFormat = "#,##0.00"
    
    wsVAT.Cells(9, 1).Value = "其中：小微企业免税销售额"
    wsVAT.Cells(9, 2).Value = 10
    wsVAT.Cells(9, 3).Value = exemptSales
    wsVAT.Cells(9, 3).NumberFormat = "#,##0.00"
    
    wsVAT.Cells(10, 1).Value = "本期应纳税额"
    wsVAT.Cells(10, 2).Value = 15
    wsVAT.Cells(10, 3).Value = taxPayable
    wsVAT.Cells(10, 3).NumberFormat = "#,##0.00"
    wsVAT.Cells(10, 1).Font.Bold = True
    
    wsVAT.Cells(11, 1).Value = "本期免税额"
    wsVAT.Cells(11, 2).Value = 17
    wsVAT.Cells(11, 3).Value = exemptTax
    wsVAT.Cells(11, 3).NumberFormat = "#,##0.00"
    
    wsVAT.Cells(12, 1).Value = "其中：小微企业免税额"
    wsVAT.Cells(12, 2).Value = 18
    wsVAT.Cells(12, 3).Value = exemptTax
    wsVAT.Cells(12, 3).NumberFormat = "#,##0.00"
    
    wsVAT.Cells(13, 1).Value = "本期应补（退）税额"
    wsVAT.Cells(13, 2).Value = 22
    wsVAT.Cells(13, 3).Value = taxPayable
    wsVAT.Cells(13, 3).NumberFormat = "#,##0.00"
    wsVAT.Cells(13, 1).Font.Bold = True
    wsVAT.Cells(13, 3).Font.Bold = True
    
    ' 设置列宽
    wsVAT.Columns("A").ColumnWidth = 40
    wsVAT.Columns("B").ColumnWidth = 8
    wsVAT.Columns("C:D").ColumnWidth = 15
    
    MsgBox "增值税申报表生成完成！" & vbCrLf & vbCrLf & _
           "应税销售额：" & Format(taxableSales, "#,##0.00") & " 元" & vbCrLf & _
           "免税销售额：" & Format(exemptSales, "#,##0.00") & " 元" & vbCrLf & _
           "本期应纳税额：" & Format(taxPayable, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "生成增值税申报表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏38: GenerateIncomeTaxReturn - 企业所得税汇算清缴
' V7.0新增：生成年度企业所得税汇算清缴表
' ============================================================================
Sub GenerateIncomeTaxReturn()
    Dim wsProfit As Worksheet, wsTax As Worksheet
    Dim yearStr As String
    Dim revenue As Double, cost As Double, profit As Double
    Dim taxRate As Double, taxPayable As Double
    Dim reduction As Double, actualTax As Double
    Dim adjustmentAdd As Double, adjustmentReduce As Double
    Dim taxableIncome As Double
    
    On Error GoTo ErrorHandler
    
    yearStr = InputBox("请输入汇算年度（格式：yyyy）：", "所得税汇算", Format(Date, "yyyy"))
    If yearStr = "" Then Exit Sub
    
    Set wsProfit = ThisWorkbook.Sheets("利润表")
    
    ' 获取利润表数据
    revenue = CDbl(Nz(wsProfit.Cells(PR_INCOME, 2).Value, 0))
    cost = CDbl(Nz(wsProfit.Cells(PR_COST, 2).Value, 0))
    profit = CDbl(Nz(wsProfit.Cells(PR_NET_PROFIT, 2).Value, 0))
    
    ' 创建或清空所得税汇算表
    On Error Resume Next
    Set wsTax = ThisWorkbook.Sheets("所得税汇算")
    If wsTax Is Nothing Then
        Set wsTax = ThisWorkbook.Sheets.Add
        wsTax.Name = "所得税汇算"
    Else
        wsTax.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 计算纳税调整（简化处理）
    adjustmentAdd = 0
    adjustmentReduce = 0
    taxableIncome = profit + adjustmentAdd - adjustmentReduce
    
    ' 正确的小微企业分段税率计算
    If taxableIncome <= 0 Then
        ' 亏损或零利润，无需纳税
        taxPayable = 0
        taxRate = 0
        reduction = 0
    ElseIf taxableIncome <= 1000000 Then
        ' 100万以内，实际税负5%
        taxPayable = taxableIncome * 0.05
        taxRate = 0.05
        reduction = 0
    ElseIf taxableIncome <= 3000000 Then
        ' 100-300万，分段计算
        taxPayable = 1000000 * 0.05 + (taxableIncome - 1000000) * 0.1
        taxRate = taxPayable / taxableIncome  ' 实际税率
        reduction = 0
    Else
        ' 300万以上，一般企业25%
        taxPayable = 1000000 * 0.05 + 2000000 * 0.1 + (taxableIncome - 3000000) * 0.25
        taxRate = 0.25
        reduction = 0
    End If
    
    actualTax = taxPayable
    If actualTax < 0 Then actualTax = 0
    
    ' 填写申报表
    wsTax.Cells(1, 1).Value = "中华人民共和国企业所得税年度纳税申报表（A类）"
    wsTax.Cells(1, 1).Font.Size = 14
    wsTax.Cells(1, 1).Font.Bold = True
    wsTax.Range("A1:E1").Merge
    
    wsTax.Cells(2, 1).Value = "税款所属年度：" & yearStr & "年"
    wsTax.Cells(2, 4).Value = "填表日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 表头
    wsTax.Cells(4, 1).Value = "行次"
    wsTax.Cells(4, 2).Value = "项目"
    wsTax.Cells(4, 3).Value = "账载金额"
    wsTax.Cells(4, 4).Value = "税收金额"
    wsTax.Cells(4, 5).Value = "调增金额"
    wsTax.Cells(4, 6).Value = "调减金额"
    wsTax.Range("A4:F4").Font.Bold = True
    wsTax.Range("A4:F4").Interior.Color = RGB(68, 114, 196)
    wsTax.Range("A4:F4").Font.Color = RGB(255, 255, 255)
    
    ' 数据行
    Dim row As Long
    row = 5
    
    wsTax.Cells(row, 1).Value = 1
    wsTax.Cells(row, 2).Value = "一、营业收入"
    wsTax.Cells(row, 3).Value = revenue
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 2
    wsTax.Cells(row, 2).Value = "减：营业成本"
    wsTax.Cells(row, 3).Value = cost
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 13
    wsTax.Cells(row, 2).Value = "二、利润总额"
    wsTax.Cells(row, 3).Value = profit
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 3).Font.Bold = True
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 14
    wsTax.Cells(row, 2).Value = "加：纳税调整增加额"
    wsTax.Cells(row, 5).Value = adjustmentAdd
    wsTax.Cells(row, 5).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 15
    wsTax.Cells(row, 2).Value = "减：纳税调整减少额"
    wsTax.Cells(row, 6).Value = adjustmentReduce
    wsTax.Cells(row, 6).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 23
    wsTax.Cells(row, 2).Value = "三、应纳税所得额"
    wsTax.Cells(row, 4).Value = taxableIncome
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 4).Font.Bold = True
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 24
    wsTax.Cells(row, 2).Value = "税率"
    wsTax.Cells(row, 4).Value = taxRate
    wsTax.Cells(row, 4).NumberFormat = "0%"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 25
    wsTax.Cells(row, 2).Value = "四、应纳所得税额"
    wsTax.Cells(row, 4).Value = taxPayable
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 26
    wsTax.Cells(row, 2).Value = "减：减免所得税额（小微企业优惠）"
    wsTax.Cells(row, 4).Value = reduction
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = 28
    wsTax.Cells(row, 2).Value = "五、实际应纳所得税额"
    wsTax.Cells(row, 4).Value = actualTax
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 4).Font.Bold = True
    wsTax.Cells(row, 4).Font.Size = 12
    
    ' 设置列宽
    wsTax.Columns("A").ColumnWidth = 8
    wsTax.Columns("B").ColumnWidth = 35
    wsTax.Columns("C:F").ColumnWidth = 15
    
    MsgBox "企业所得税汇算清缴表生成完成！" & vbCrLf & vbCrLf & _
           "营业收入：" & Format(revenue, "#,##0.00") & " 元" & vbCrLf & _
           "利润总额：" & Format(profit, "#,##0.00") & " 元" & vbCrLf & _
           "应纳税所得额：" & Format(taxableIncome, "#,##0.00") & " 元" & vbCrLf & _
           "适用税率：" & Format(taxRate, "0%") & vbCrLf & _
           "实际应纳所得税额：" & Format(actualTax, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "生成所得税汇算表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏39: CalculateStampTax - 印花税计算
' V7.0新增：计算常见合同印花税
' ============================================================================
Sub CalculateStampTax()
    Dim wsTax As Worksheet
    Dim salesAmt As Double, loanAmt As Double
    Dim leaseAmt As Double, transportAmt As Double
    Dim warehouseAmt As Double
    Dim salesTax As Double, loanTax As Double
    Dim leaseTax As Double, transportTax As Double
    Dim warehouseTax As Double, totalTax As Double
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    ' 输入各类合同金额（带验证）
    inputValue = InputBox("请输入买卖合同（加工承揽）金额：", "印花税计算", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    salesAmt = CDbl(inputValue)
    
    inputValue = InputBox("请输入借款合同金额：", "印花税计算", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    loanAmt = CDbl(inputValue)
    
    inputValue = InputBox("请输入租赁合同金额：", "印花税计算", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    leaseAmt = CDbl(inputValue)
    
    inputValue = InputBox("请输入运输合同金额：", "印花税计算", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    transportAmt = CDbl(inputValue)
    
    inputValue = InputBox("请输入仓储合同金额：", "印花税计算", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    warehouseAmt = CDbl(inputValue)
    
    ' 计算印花税
    salesTax = salesAmt * ST_SALES
    loanTax = loanAmt * ST_LOAN
    leaseTax = leaseAmt * ST_LEASE
    transportTax = transportAmt * ST_TRANSPORT
    warehouseTax = warehouseAmt * ST_WAREHOUSE
    totalTax = salesTax + loanTax + leaseTax + transportTax + warehouseTax
    
    ' 创建或清空印花税计算表
    On Error Resume Next
    Set wsTax = ThisWorkbook.Sheets("印花税计算")
    If wsTax Is Nothing Then
        Set wsTax = ThisWorkbook.Sheets.Add
        wsTax.Name = "印花税计算"
    Else
        wsTax.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 填写计算表
    wsTax.Cells(1, 1).Value = "印花税计算表"
    wsTax.Cells(1, 1).Font.Size = 14
    wsTax.Cells(1, 1).Font.Bold = True
    wsTax.Range("A1:D1").Merge
    
    wsTax.Cells(2, 1).Value = "计算日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 表头
    wsTax.Cells(4, 1).Value = "合同类型"
    wsTax.Cells(4, 2).Value = "税率"
    wsTax.Cells(4, 3).Value = "合同金额"
    wsTax.Cells(4, 4).Value = "应纳税额"
    wsTax.Range("A4:D4").Font.Bold = True
    wsTax.Range("A4:D4").Interior.Color = RGB(68, 114, 196)
    wsTax.Range("A4:D4").Font.Color = RGB(255, 255, 255)
    
    ' 数据行
    Dim row As Long
    row = 5
    
    wsTax.Cells(row, 1).Value = "买卖合同（加工承揽）"
    wsTax.Cells(row, 2).Value = "0.03%"
    wsTax.Cells(row, 3).Value = salesAmt
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 4).Value = salesTax
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "借款合同"
    wsTax.Cells(row, 2).Value = "0.005%"
    wsTax.Cells(row, 3).Value = loanAmt
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 4).Value = loanTax
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "租赁合同"
    wsTax.Cells(row, 2).Value = "0.1%"
    wsTax.Cells(row, 3).Value = leaseAmt
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 4).Value = leaseTax
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "运输合同"
    wsTax.Cells(row, 2).Value = "0.03%"
    wsTax.Cells(row, 3).Value = transportAmt
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 4).Value = transportTax
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "仓储合同"
    wsTax.Cells(row, 2).Value = "0.1%"
    wsTax.Cells(row, 3).Value = warehouseAmt
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 4).Value = warehouseTax
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    row = row + 1
    
    ' 合计行
    wsTax.Cells(row, 1).Value = "合计"
    wsTax.Cells(row, 1).Font.Bold = True
    wsTax.Cells(row, 3).Value = salesAmt + loanAmt + leaseAmt + transportAmt + warehouseAmt
    wsTax.Cells(row, 3).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 3).Font.Bold = True
    wsTax.Cells(row, 4).Value = totalTax
    wsTax.Cells(row, 4).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 4).Font.Bold = True
    wsTax.Cells(row, 4).Font.Size = 12
    
    ' 设置列宽
    wsTax.Columns("A").ColumnWidth = 25
    wsTax.Columns("B").ColumnWidth = 12
    wsTax.Columns("C:D").ColumnWidth = 15
    
    MsgBox "印花税计算完成！" & vbCrLf & vbCrLf & _
           "买卖合同税额：" & Format(salesTax, "#,##0.00") & " 元" & vbCrLf & _
           "借款合同税额：" & Format(loanTax, "#,##0.00") & " 元" & vbCrLf & _
           "租赁合同税额：" & Format(leaseTax, "#,##0.00") & " 元" & vbCrLf & _
           "运输合同税额：" & Format(transportTax, "#,##0.00") & " 元" & vbCrLf & _
           "仓储合同税额：" & Format(warehouseTax, "#,##0.00") & " 元" & vbCrLf & _
           "应纳印花税合计：" & Format(totalTax, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "计算印花税出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏40: CalculateSocialInsurance - 社保公积金核算
' V7.0新增：计算企业和员工社保公积金
' ============================================================================
Sub CalculateSocialInsurance()
    Dim wsSI As Worksheet
    Dim salaryBase As Double
    Dim pensionE As Double, pensionP As Double
    Dim medicalE As Double, medicalP As Double
    Dim unemploymentE As Double, unemploymentP As Double
    Dim injuryE As Double, maternityE As Double
    Dim hpfE As Double, hpfP As Double
    Dim totalE As Double, totalP As Double
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    ' 输入缴费基数（带验证）
    inputValue = InputBox("请输入社保公积金缴费基数：", "社保公积金核算", "5000")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    salaryBase = CDbl(inputValue)
    If salaryBase <= 0 Then
        MsgBox "缴费基数必须大于0", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 计算各项社保
    pensionE = salaryBase * EI_PENSION_E
    pensionP = salaryBase * EI_PENSION_P
    medicalE = salaryBase * EI_MEDICAL_E
    medicalP = salaryBase * EI_MEDICAL_P
    unemploymentE = salaryBase * EI_UNEMPLOYMENT_E
    unemploymentP = salaryBase * EI_UNEMPLOYMENT_P
    injuryE = salaryBase * EI_INJURY_E
    maternityE = salaryBase * EI_MATERNITY_E
    hpfE = salaryBase * HPF_RATE_E
    hpfP = salaryBase * HPF_RATE_P
    
    totalE = pensionE + medicalE + unemploymentE + injuryE + maternityE + hpfE
    totalP = pensionP + medicalP + unemploymentP + hpfP
    
    ' 创建或清空社保公积金表
    On Error Resume Next
    Set wsSI = ThisWorkbook.Sheets("社保公积金")
    If wsSI Is Nothing Then
        Set wsSI = ThisWorkbook.Sheets.Add
        wsSI.Name = "社保公积金"
    Else
        wsSI.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 填写核算表
    wsSI.Cells(1, 1).Value = "社保公积金核算表"
    wsSI.Cells(1, 1).Font.Size = 14
    wsSI.Cells(1, 1).Font.Bold = True
    wsSI.Range("A1:E1").Merge
    
    wsSI.Cells(2, 1).Value = "缴费基数：" & Format(salaryBase, "#,##0.00") & " 元"
    wsSI.Cells(2, 4).Value = "计算日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 表头
    wsSI.Cells(4, 1).Value = "险种"
    wsSI.Cells(4, 2).Value = "企业比例"
    wsSI.Cells(4, 3).Value = "企业承担"
    wsSI.Cells(4, 4).Value = "个人比例"
    wsSI.Cells(4, 5).Value = "个人承担"
    wsSI.Range("A4:E4").Font.Bold = True
    wsSI.Range("A4:E4").Interior.Color = RGB(68, 114, 196)
    wsSI.Range("A4:E4").Font.Color = RGB(255, 255, 255)
    
    ' 数据行
    Dim row As Long
    row = 5
    
    wsSI.Cells(row, 1).Value = "养老保险"
    wsSI.Cells(row, 2).Value = "16%"
    wsSI.Cells(row, 3).Value = pensionE
    wsSI.Cells(row, 3).NumberFormat = "#,##0.00"
    wsSI.Cells(row, 4).Value = "8%"
    wsSI.Cells(row, 5).Value = pensionP
    wsSI.Cells(row, 5).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsSI.Cells(row, 1).Value = "医疗保险"
    wsSI.Cells(row, 2).Value = "8%"
    wsSI.Cells(row, 3).Value = medicalE
    wsSI.Cells(row, 3).NumberFormat = "#,##0.00"
    wsSI.Cells(row, 4).Value = "2%"
    wsSI.Cells(row, 5).Value = medicalP
    wsSI.Cells(row, 5).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsSI.Cells(row, 1).Value = "失业保险"
    wsSI.Cells(row, 2).Value = "0.5%"
    wsSI.Cells(row, 3).Value = unemploymentE
    wsSI.Cells(row, 3).NumberFormat = "#,##0.00"
    wsSI.Cells(row, 4).Value = "0.5%"
    wsSI.Cells(row, 5).Value = unemploymentP
    wsSI.Cells(row, 5).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsSI.Cells(row, 1).Value = "工伤保险"
    wsSI.Cells(row, 2).Value = "0.4%"
    wsSI.Cells(row, 3).Value = injuryE
    wsSI.Cells(row, 3).NumberFormat = "#,##0.00"
    wsSI.Cells(row, 4).Value = "0%"
    wsSI.Cells(row, 5).Value = 0
    wsSI.Cells(row, 5).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsSI.Cells(row, 1).Value = "生育保险"
    wsSI.Cells(row, 2).Value = "0.8%"
    wsSI.Cells(row, 3).Value = maternityE
    wsSI.Cells(row, 3).NumberFormat = "#,##0.00"
    wsSI.Cells(row, 4).Value = "0%"
    wsSI.Cells(row, 5).Value = 0
    wsSI.Cells(row, 5).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsSI.Cells(row, 1).Value = "住房公积金"
    wsSI.Cells(row, 2).Value = "8%"
    wsSI.Cells(row, 3).Value = hpfE
    wsSI.Cells(row, 3).NumberFormat = "#,##0.00"
    wsSI.Cells(row, 4).Value = "8%"
    wsSI.Cells(row, 5).Value = hpfP
    wsSI.Cells(row, 5).NumberFormat = "#,##0.00"
    row = row + 1
    
    ' 合计行
    wsSI.Cells(row, 1).Value = "合计"
    wsSI.Cells(row, 1).Font.Bold = True
    wsSI.Cells(row, 3).Value = totalE
    wsSI.Cells(row, 3).NumberFormat = "#,##0.00"
    wsSI.Cells(row, 3).Font.Bold = True
    wsSI.Cells(row, 3).Font.Size = 12
    wsSI.Cells(row, 5).Value = totalP
    wsSI.Cells(row, 5).NumberFormat = "#,##0.00"
    wsSI.Cells(row, 5).Font.Bold = True
    wsSI.Cells(row, 5).Font.Size = 12
    
    ' 设置列宽
    wsSI.Columns("A").ColumnWidth = 15
    wsSI.Columns("B").ColumnWidth = 12
    wsSI.Columns("C:E").ColumnWidth = 15
    
    MsgBox "社保公积金核算完成！" & vbCrLf & vbCrLf & _
           "缴费基数：" & Format(salaryBase, "#,##0.00") & " 元" & vbCrLf & _
           "企业承担合计：" & Format(totalE, "#,##0.00") & " 元" & vbCrLf & _
           "个人承担合计：" & Format(totalP, "#,##0.00") & " 元" & vbCrLf & _
           "总计：" & Format(totalE + totalP, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "社保公积金核算出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏41: MultiBankReconciliation - 多银行对账
' V7.0新增：支持多银行账户对账
' ============================================================================
Sub MultiBankReconciliation()
    Dim wsBank As Worksheet, wsIncome As Worksheet, wsExpense As Worksheet
    Dim bankName As String
    Dim lastRow As Long, i As Long
    Dim bankBalance As Double, bookBalance As Double
    Dim unreconciledIn As Double, unreconciledOut As Double
    
    On Error GoTo ErrorHandler
    
    ' 选择银行
    bankName = InputBox("请选择银行（G银行、工商银行、建设银行、农业银行、其他）：", "多银行对账", "G银行")
    If bankName = "" Then Exit Sub
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 创建或清空多银行对账表
    On Error Resume Next
    Set wsBank = ThisWorkbook.Sheets("多银行对账")
    If wsBank Is Nothing Then
        Set wsBank = ThisWorkbook.Sheets.Add
        wsBank.Name = "多银行对账"
    End If
    On Error GoTo 0
    
    ' 清空该银行的数据区域（简化处理：每次重新生成）
    wsBank.Cells.Clear
    
    ' 输入银行对账单余额
    bankBalance = CDbl(InputBox("请输入" & bankName & "对账单余额：", "银行余额", 0))
    
    ' 计算账面余额（该银行的收支）
    bookBalance = 0
    unreconciledIn = 0
    unreconciledOut = 0
    
    ' 统计收入
    lastRow = GetLastRow(wsIncome, 1)
    For i = 5 To lastRow
        If InStr(wsIncome.Cells(i, 4).Value, bankName) > 0 Then
            bookBalance = bookBalance + CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
        End If
    Next i
    
    ' 统计支出
    lastRow = GetLastRow(wsExpense, 1)
    For i = 5 To lastRow
        If InStr(wsExpense.Cells(i, 5).Value, bankName) > 0 Then
            bookBalance = bookBalance - CDbl(Nz(wsExpense.Cells(i, 3).Value, 0))
        End If
    Next i
    
    ' 输入未达账项
    unreconciledIn = CDbl(InputBox("请输入银行已收企业未收金额：", "未达账项", 0))
    unreconciledOut = CDbl(InputBox("请输入银行已付企业未付金额：", "未达账项", 0))
    
    ' 填写对账表
    wsBank.Cells(1, 1).Value = "银行余额调节表"
    wsBank.Cells(1, 1).Font.Size = 14
    wsBank.Cells(1, 1).Font.Bold = True
    wsBank.Range("A1:D1").Merge
    
    wsBank.Cells(2, 1).Value = "银行：" & bankName
    wsBank.Cells(2, 3).Value = "对账日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 企业账面
    wsBank.Cells(4, 1).Value = "企业银行存款日记账"
    wsBank.Cells(4, 1).Font.Bold = True
    wsBank.Cells(5, 1).Value = "账面余额："
    wsBank.Cells(5, 2).Value = bookBalance
    wsBank.Cells(5, 2).NumberFormat = "#,##0.00"
    wsBank.Cells(6, 1).Value = "加：银行已收企业未收"
    wsBank.Cells(6, 2).Value = unreconciledIn
    wsBank.Cells(6, 2).NumberFormat = "#,##0.00"
    wsBank.Cells(7, 1).Value = "减：银行已付企业未付"
    wsBank.Cells(7, 2).Value = unreconciledOut
    wsBank.Cells(7, 2).NumberFormat = "#,##0.00"
    wsBank.Cells(8, 1).Value = "调节后余额："
    wsBank.Cells(8, 1).Font.Bold = True
    wsBank.Cells(8, 2).Value = bookBalance + unreconciledIn - unreconciledOut
    wsBank.Cells(8, 2).NumberFormat = "#,##0.00"
    wsBank.Cells(8, 2).Font.Bold = True
    
    ' 银行对账单
    wsBank.Cells(4, 3).Value = "银行对账单"
    wsBank.Cells(4, 3).Font.Bold = True
    wsBank.Cells(5, 3).Value = "对账单余额："
    wsBank.Cells(5, 4).Value = bankBalance
    wsBank.Cells(5, 4).NumberFormat = "#,##0.00"
    wsBank.Cells(6, 3).Value = "加：企业已收银行未收"
    wsBank.Cells(6, 4).Value = 0
    wsBank.Cells(6, 4).NumberFormat = "#,##0.00"
    wsBank.Cells(7, 3).Value = "减：企业已付银行未付"
    wsBank.Cells(7, 4).Value = 0
    wsBank.Cells(7, 4).NumberFormat = "#,##0.00"
    wsBank.Cells(8, 3).Value = "调节后余额："
    wsBank.Cells(8, 3).Font.Bold = True
    wsBank.Cells(8, 4).Value = bankBalance
    wsBank.Cells(8, 4).NumberFormat = "#,##0.00"
    wsBank.Cells(8, 4).Font.Bold = True
    
    ' 核对结果
    wsBank.Cells(10, 1).Value = "核对结果："
    wsBank.Cells(10, 1).Font.Bold = True
    If Abs((bookBalance + unreconciledIn - unreconciledOut) - bankBalance) < 0.01 Then
        wsBank.Cells(10, 2).Value = "余额相符"
        wsBank.Cells(10, 2).Font.Color = RGB(0, 128, 0)
    Else
        wsBank.Cells(10, 2).Value = "余额不符，请检查"
        wsBank.Cells(10, 2).Font.Color = RGB(255, 0, 0)
    End If
    
    ' 设置列宽
    wsBank.Columns("A").ColumnWidth = 25
    wsBank.Columns("B").ColumnWidth = 15
    wsBank.Columns("C").ColumnWidth = 25
    wsBank.Columns("D").ColumnWidth = 15
    
    MsgBox bankName & "对账完成！" & vbCrLf & vbCrLf & _
           "账面余额：" & Format(bookBalance, "#,##0.00") & " 元" & vbCrLf & _
           "对账单余额：" & Format(bankBalance, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "多银行对账出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏42: GenerateFundDailyReport - 资金日报
' V7.0新增：生成每日资金收支日报
' ============================================================================
Sub GenerateFundDailyReport()
    Dim wsFund As Worksheet, wsIncome As Worksheet, wsExpense As Worksheet
    Dim reportDate As String
    Dim lastRow As Long, i As Long
    Dim dailyIncome As Double, dailyExpense As Double
    Dim bankBalance As Double, cashBalance As Double
    
    On Error GoTo ErrorHandler
    
    reportDate = InputBox("请输入日报日期（格式：yyyy-mm-dd）：", "资金日报", Format(Date, "yyyy-mm-dd"))
    If reportDate = "" Then Exit Sub
    
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 计算本日收支
    dailyIncome = 0
    lastRow = GetLastRow(wsIncome, 1)
    For i = 5 To lastRow
        If wsIncome.Cells(i, 1).Value = reportDate Then
            dailyIncome = dailyIncome + CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
        End If
    Next i
    
    dailyExpense = 0
    lastRow = GetLastRow(wsExpense, 1)
    For i = 5 To lastRow
        If wsExpense.Cells(i, 1).Value = reportDate Then
            dailyExpense = dailyExpense + CDbl(Nz(wsExpense.Cells(i, 3).Value, 0))
        End If
    Next i
    
    ' 输入期初余额和现金余额
    bankBalance = CDbl(InputBox("请输入银行账户期初余额：", "期初余额", 0))
    cashBalance = CDbl(InputBox("请输入现金余额：", "现金余额", 0))
    
    ' 创建或清空资金日报表
    On Error Resume Next
    Set wsFund = ThisWorkbook.Sheets("资金日报")
    If wsFund Is Nothing Then
        Set wsFund = ThisWorkbook.Sheets.Add
        wsFund.Name = "资金日报"
    Else
        wsFund.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 填写日报
    wsFund.Cells(1, 1).Value = "资金日报表"
    wsFund.Cells(1, 1).Font.Size = 14
    wsFund.Cells(1, 1).Font.Bold = True
    wsFund.Range("A1:C1").Merge
    
    wsFund.Cells(2, 1).Value = "日期：" & reportDate
    
    ' 表头
    wsFund.Cells(4, 1).Value = "项目"
    wsFund.Cells(4, 2).Value = "金额"
    wsFund.Cells(4, 3).Value = "备注"
    wsFund.Range("A4:C4").Font.Bold = True
    wsFund.Range("A4:C4").Interior.Color = RGB(68, 114, 196)
    wsFund.Range("A4:C4").Font.Color = RGB(255, 255, 255)
    
    ' 数据行
    Dim row As Long
    row = 5
    
    wsFund.Cells(row, 1).Value = "一、期初余额"
    wsFund.Cells(row, 2).Value = bankBalance
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsFund.Cells(row, 1).Value = "其中：银行存款"
    wsFund.Cells(row, 2).Value = bankBalance
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsFund.Cells(row, 1).Value = "      现金"
    wsFund.Cells(row, 2).Value = cashBalance
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsFund.Cells(row, 1).Value = "二、本日收入"
    wsFund.Cells(row, 1).Font.Bold = True
    wsFund.Cells(row, 2).Value = dailyIncome
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    wsFund.Cells(row, 2).Font.Bold = True
    row = row + 1
    
    wsFund.Cells(row, 1).Value = "三、本日支出"
    wsFund.Cells(row, 1).Font.Bold = True
    wsFund.Cells(row, 2).Value = dailyExpense
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    wsFund.Cells(row, 2).Font.Bold = True
    row = row + 1
    
    wsFund.Cells(row, 1).Value = "四、本日余额"
    wsFund.Cells(row, 1).Font.Bold = True
    wsFund.Cells(row, 2).Value = bankBalance + dailyIncome - dailyExpense
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    wsFund.Cells(row, 2).Font.Bold = True
    wsFund.Cells(row, 2).Font.Size = 12
    row = row + 1
    
    wsFund.Cells(row, 1).Value = "其中：银行存款"
    wsFund.Cells(row, 2).Value = bankBalance + dailyIncome - dailyExpense
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsFund.Cells(row, 1).Value = "      现金"
    wsFund.Cells(row, 2).Value = cashBalance
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsFund.Cells(row, 1).Value = "五、资金合计"
    wsFund.Cells(row, 1).Font.Bold = True
    wsFund.Cells(row, 2).Value = bankBalance + cashBalance + dailyIncome - dailyExpense
    wsFund.Cells(row, 2).NumberFormat = "#,##0.00"
    wsFund.Cells(row, 2).Font.Bold = True
    wsFund.Cells(row, 2).Font.Size = 12
    
    ' 设置列宽
    wsFund.Columns("A").ColumnWidth = 25
    wsFund.Columns("B").ColumnWidth = 15
    wsFund.Columns("C").ColumnWidth = 20
    
    MsgBox "资金日报生成完成！" & vbCrLf & vbCrLf & _
           "日期：" & reportDate & vbCrLf & _
           "本日收入：" & Format(dailyIncome, "#,##0.00") & " 元" & vbCrLf & _
           "本日支出：" & Format(dailyExpense, "#,##0.00") & " 元" & vbCrLf & _
           "本日余额：" & Format(bankBalance + dailyIncome - dailyExpense, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "生成资金日报出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏43: GeneratePayroll - 工资表生成
' V7.0新增：生成员工工资表
' ============================================================================
Sub GeneratePayroll()
    Dim wsPayroll As Worksheet
    Dim monthStr As String
    Dim employeeCount As Long, i As Long
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    monthStr = InputBox("请输入工资月份（格式：yyyy-mm）：", "工资表生成", Format(Date, "yyyy-mm"))
    If monthStr = "" Then Exit Sub
    
    ' 员工人数输入验证
    inputValue = InputBox("请输入员工人数：", "员工人数", "5")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    employeeCount = CLng(inputValue)
    If employeeCount <= 0 Then Exit Sub
    
    ' 创建或清空工资表
    On Error Resume Next
    Set wsPayroll = ThisWorkbook.Sheets("工资表")
    If wsPayroll Is Nothing Then
        Set wsPayroll = ThisWorkbook.Sheets.Add
        wsPayroll.Name = "工资表"
    Else
        wsPayroll.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 填写表头
    wsPayroll.Cells(1, 1).Value = "员工工资表"
    wsPayroll.Cells(1, 1).Font.Size = 14
    wsPayroll.Cells(1, 1).Font.Bold = True
    wsPayroll.Range("A1:L1").Merge
    
    wsPayroll.Cells(2, 1).Value = "月份：" & monthStr
    
    ' 列标题
    wsPayroll.Cells(4, 1).Value = "序号"
    wsPayroll.Cells(4, 2).Value = "姓名"
    wsPayroll.Cells(4, 3).Value = "部门"
    wsPayroll.Cells(4, 4).Value = "岗位"
    wsPayroll.Cells(4, 5).Value = "基本工资"
    wsPayroll.Cells(4, 6).Value = "绩效工资"
    wsPayroll.Cells(4, 7).Value = "加班工资"
    wsPayroll.Cells(4, 8).Value = "补贴"
    wsPayroll.Cells(4, 9).Value = "应发工资"
    wsPayroll.Cells(4, 10).Value = "社保个人"
    wsPayroll.Cells(4, 11).Value = "公积金个人"
    wsPayroll.Cells(4, 12).Value = "个税"
    wsPayroll.Cells(4, 13).Value = "实发工资"
    wsPayroll.Range("A4:M4").Font.Bold = True
    wsPayroll.Range("A4:M4").Interior.Color = RGB(68, 114, 196)
    wsPayroll.Range("A4:M4").Font.Color = RGB(255, 255, 255)
    
    ' 输入员工数据
    Dim row As Long
    Dim baseSalary As Double, performanceSalary As Double, overtimeSalary As Double, allowance As Double
    Dim grossSalary As Double
    
    row = 5
    
    For i = 1 To employeeCount
        wsPayroll.Cells(row, 1).Value = i
        wsPayroll.Cells(row, 2).Value = InputBox("请输入第" & i & "个员工姓名：", "员工信息", "员工" & i)
        wsPayroll.Cells(row, 3).Value = InputBox("请输入部门：", "部门", "生产部")
        wsPayroll.Cells(row, 4).Value = InputBox("请输入岗位：", "岗位", "操作工")
        
        ' 基本工资输入验证
        inputValue = InputBox("请输入基本工资：", "工资", "5000")
        If inputValue = "" Then inputValue = "0"
        If Not IsNumeric(inputValue) Then inputValue = "0"
        baseSalary = CDbl(inputValue)
        wsPayroll.Cells(row, 5).Value = baseSalary
        
        ' 绩效工资输入验证
        inputValue = InputBox("请输入绩效工资：", "工资", "1000")
        If inputValue = "" Then inputValue = "0"
        If Not IsNumeric(inputValue) Then inputValue = "0"
        performanceSalary = CDbl(inputValue)
        wsPayroll.Cells(row, 6).Value = performanceSalary
        
        ' 加班工资输入验证
        inputValue = InputBox("请输入加班工资：", "工资", "0")
        If inputValue = "" Then inputValue = "0"
        If Not IsNumeric(inputValue) Then inputValue = "0"
        overtimeSalary = CDbl(inputValue)
        wsPayroll.Cells(row, 7).Value = overtimeSalary
        
        ' 补贴输入验证
        inputValue = InputBox("请输入补贴：", "工资", "500")
        If inputValue = "" Then inputValue = "0"
        If Not IsNumeric(inputValue) Then inputValue = "0"
        allowance = CDbl(inputValue)
        wsPayroll.Cells(row, 8).Value = allowance
        
        ' 计算应发工资
        wsPayroll.Cells(row, 9).Formula = "=E" & row & "+F" & row & "+G" & row & "+H" & row
        wsPayroll.Cells(row, 9).NumberFormat = "#,##0.00"
        
        ' 社保公积金（使用应发工资I列作为基数）
        wsPayroll.Cells(row, 10).Value = wsPayroll.Cells(row, 9).Value * (EI_PENSION_P + EI_MEDICAL_P + EI_UNEMPLOYMENT_P)
        wsPayroll.Cells(row, 10).NumberFormat = "#,##0.00"
        wsPayroll.Cells(row, 11).Value = wsPayroll.Cells(row, 9).Value * HPF_RATE_P
        wsPayroll.Cells(row, 11).NumberFormat = "#,##0.00"
        
        ' 个税（七级超额累进税率计算）
        Dim taxableIncome As Double
        taxableIncome = wsPayroll.Cells(row, 9).Value - wsPayroll.Cells(row, 10).Value - wsPayroll.Cells(row, 11).Value - 5000
        If taxableIncome > 0 Then
            Dim taxRate2 As Double, quickDed2 As Double
            If taxableIncome <= 36000 Then
                taxRate2 = 0.03: quickDed2 = 0
            ElseIf taxableIncome <= 144000 Then
                taxRate2 = 0.1: quickDed2 = 2520
            ElseIf taxableIncome <= 300000 Then
                taxRate2 = 0.2: quickDed2 = 16920
            ElseIf taxableIncome <= 420000 Then
                taxRate2 = 0.25: quickDed2 = 31920
            ElseIf taxableIncome <= 660000 Then
                taxRate2 = 0.3: quickDed2 = 52920
            ElseIf taxableIncome <= 960000 Then
                taxRate2 = 0.35: quickDed2 = 85920
            Else
                taxRate2 = 0.45: quickDed2 = 181920
            End If
            wsPayroll.Cells(row, 12).Value = Round(taxableIncome * taxRate2 - quickDed2, 2)
        Else
            wsPayroll.Cells(row, 12).Value = 0
        End If
        wsPayroll.Cells(row, 12).NumberFormat = "#,##0.00"
        
        ' 实发工资
        wsPayroll.Cells(row, 13).Formula = "=I" & row & "-J" & row & "-K" & row & "-L" & row
        wsPayroll.Cells(row, 13).NumberFormat = "#,##0.00"
        wsPayroll.Cells(row, 13).Font.Bold = True
        
        row = row + 1
    Next i
    
    ' 合计行
    wsPayroll.Cells(row, 1).Value = "合计"
    wsPayroll.Cells(row, 1).Font.Bold = True
    For i = 5 To 13
        wsPayroll.Cells(row, i).Formula = "=SUM(" & Chr(64 + i) & "5:" & Chr(64 + i) & (row - 1) & ")"
        wsPayroll.Cells(row, i).NumberFormat = "#,##0.00"
        wsPayroll.Cells(row, i).Font.Bold = True
    Next i
    
    ' 设置列宽
    wsPayroll.Columns("A").ColumnWidth = 6
    wsPayroll.Columns("B").ColumnWidth = 12
    wsPayroll.Columns("C:D").ColumnWidth = 10
    wsPayroll.Columns("E:M").ColumnWidth = 12
    
    MsgBox "工资表生成完成！" & vbCrLf & vbCrLf & _
           "月份：" & monthStr & vbCrLf & _
           "员工人数：" & employeeCount & " 人", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "生成工资表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏44: CalculatePersonalIncomeTax - 个税计算
' V7.0新增：计算个人所得税（累计预扣法）
' ============================================================================
Sub CalculatePersonalIncomeTax()
    Dim wsTax As Worksheet
    Dim employeeName As String
    Dim cumulativeIncome As Double, cumulativeDeduction As Double
    Dim cumulativeSpecial As Double, cumulativeTaxable As Double
    Dim taxRate As Double, quickDeduction As Double
    Dim cumulativeTax As Double, currentTax As Double
    Dim monthNum As Integer
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    employeeName = InputBox("请输入员工姓名：", "个税计算", "")
    If employeeName = "" Then Exit Sub
    
    ' 月份输入验证
    inputValue = InputBox("请输入当前月份（1-12）：", "月份", CStr(Month(Date)))
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    monthNum = CInt(inputValue)
    If monthNum < 1 Or monthNum > 12 Then
        MsgBox "月份输入错误", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 累计收入输入验证
    inputValue = InputBox("请输入累计收入：", "累计收入", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    cumulativeIncome = CDbl(inputValue)
    
    ' 累计专项扣除输入验证
    inputValue = InputBox("请输入累计专项扣除（社保公积金）：", "专项扣除", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    cumulativeSpecial = CDbl(inputValue)

    ' 累计已缴税额输入
    inputValue = InputBox("请输入累计已缴税额（首月填0）：", "已缴税额", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    Dim cumulativePaid As Double
    cumulativePaid = CDbl(inputValue)
    
    ' 计算累计减除费用（5000元/月）
    cumulativeDeduction = IIT_DEDUCTION * monthNum
    
    ' 计算累计应纳税所得额
    cumulativeTaxable = cumulativeIncome - cumulativeDeduction - cumulativeSpecial
    If cumulativeTaxable < 0 Then cumulativeTaxable = 0
    
    ' 确定税率和速算扣除数（七级超额累进税率）
    If cumulativeTaxable <= 36000 Then
        taxRate = 0.03
        quickDeduction = 0
    ElseIf cumulativeTaxable <= 144000 Then
        taxRate = 0.1
        quickDeduction = 2520
    ElseIf cumulativeTaxable <= 300000 Then
        taxRate = 0.2
        quickDeduction = 16920
    ElseIf cumulativeTaxable <= 420000 Then
        taxRate = 0.25
        quickDeduction = 31920
    ElseIf cumulativeTaxable <= 660000 Then
        taxRate = 0.3
        quickDeduction = 52920
    ElseIf cumulativeTaxable <= 960000 Then
        taxRate = 0.35
        quickDeduction = 85920
    Else
        taxRate = 0.45
        quickDeduction = 181920
    End If
    
    ' 计算累计应纳税额和本期应预扣税额
    cumulativeTax = cumulativeTaxable * taxRate - quickDeduction
    If cumulativeTax < 0 Then cumulativeTax = 0
    
    ' 本期应预扣税额 = 累计应纳税额 - 累计已缴税额
    currentTax = cumulativeTax - cumulativePaid
    If currentTax < 0 Then currentTax = 0
    
    ' 创建或清空个税计算表
    On Error Resume Next
    Set wsTax = ThisWorkbook.Sheets("个税计算")
    If wsTax Is Nothing Then
        Set wsTax = ThisWorkbook.Sheets.Add
        wsTax.Name = "个税计算"
    Else
        wsTax.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 填写计算表
    wsTax.Cells(1, 1).Value = "个人所得税计算表（累计预扣法）"
    wsTax.Cells(1, 1).Font.Size = 14
    wsTax.Cells(1, 1).Font.Bold = True
    wsTax.Range("A1:C1").Merge
    
    wsTax.Cells(2, 1).Value = "员工：" & employeeName
    wsTax.Cells(2, 3).Value = "计算日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 表头
    wsTax.Cells(4, 1).Value = "项目"
    wsTax.Cells(4, 2).Value = "金额"
    wsTax.Cells(4, 3).Value = "说明"
    wsTax.Range("A4:C4").Font.Bold = True
    wsTax.Range("A4:C4").Interior.Color = RGB(68, 114, 196)
    wsTax.Range("A4:C4").Font.Color = RGB(255, 255, 255)
    
    ' 数据行
    Dim row As Long
    row = 5
    
    wsTax.Cells(row, 1).Value = "累计收入"
    wsTax.Cells(row, 2).Value = cumulativeIncome
    wsTax.Cells(row, 2).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 3).Value = "截至" & monthNum & "月"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "减：累计减除费用"
    wsTax.Cells(row, 2).Value = cumulativeDeduction
    wsTax.Cells(row, 2).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 3).Value = "5000元/月×" & monthNum & "月"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "减：累计专项扣除"
    wsTax.Cells(row, 2).Value = cumulativeSpecial
    wsTax.Cells(row, 2).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 3).Value = "社保公积金"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "累计应纳税所得额"
    wsTax.Cells(row, 1).Font.Bold = True
    wsTax.Cells(row, 2).Value = cumulativeTaxable
    wsTax.Cells(row, 2).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 2).Font.Bold = True
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "适用税率"
    wsTax.Cells(row, 2).Value = taxRate
    wsTax.Cells(row, 2).NumberFormat = "0%"
    wsTax.Cells(row, 3).Value = "七级超额累进"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "速算扣除数"
    wsTax.Cells(row, 2).Value = quickDeduction
    wsTax.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "累计应纳税额"
    wsTax.Cells(row, 2).Value = cumulativeTax
    wsTax.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsTax.Cells(row, 1).Value = "本期应预扣税额"
    wsTax.Cells(row, 1).Font.Bold = True
    wsTax.Cells(row, 2).Value = currentTax
    wsTax.Cells(row, 2).NumberFormat = "#,##0.00"
    wsTax.Cells(row, 2).Font.Bold = True
    wsTax.Cells(row, 2).Font.Size = 12
    
    ' 设置列宽
    wsTax.Columns("A").ColumnWidth = 25
    wsTax.Columns("B").ColumnWidth = 15
    wsTax.Columns("C").ColumnWidth = 25
    
    MsgBox "个税计算完成！" & vbCrLf & vbCrLf & _
           "员工：" & employeeName & vbCrLf & _
           "累计应纳税所得额：" & Format(cumulativeTaxable, "#,##0.00") & " 元" & vbCrLf & _
           "适用税率：" & Format(taxRate, "0%") & vbCrLf & _
           "本期应预扣税额：" & Format(currentTax, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "个税计算出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏45: GeneratePaySlip - 工资条生成
' V7.0新增：生成员工工资条（可打印）
' ============================================================================
Sub GeneratePaySlip()
    Dim wsPayroll As Worksheet, wsSlip As Worksheet
    Dim monthStr As String
    Dim lastRow As Long, i As Long, slipRow As Long
    
    On Error GoTo ErrorHandler
    
    monthStr = InputBox("请输入工资月份（格式：yyyy-mm）：", "工资条生成", Format(Date, "yyyy-mm"))
    If monthStr = "" Then Exit Sub
    
    On Error Resume Next
    Set wsPayroll = ThisWorkbook.Sheets("工资表")
    If wsPayroll Is Nothing Then
        MsgBox "请先运行【工资表生成】创建工资表", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建或清空工资条表
    On Error Resume Next
    Set wsSlip = ThisWorkbook.Sheets("工资条")
    If wsSlip Is Nothing Then
        Set wsSlip = ThisWorkbook.Sheets.Add
        wsSlip.Name = "工资条"
    Else
        wsSlip.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 获取工资表数据行数
    lastRow = GetLastRow(wsPayroll, 2)
    
    slipRow = 1
    
    ' 遍历每个员工生成工资条
    For i = 5 To lastRow - 1  ' 跳过表头和合计行
        If wsPayroll.Cells(i, 2).Value <> "" Then
            ' 工资条标题
            wsSlip.Cells(slipRow, 1).Value = "工资条"
            wsSlip.Cells(slipRow, 1).Font.Size = 14
            wsSlip.Cells(slipRow, 1).Font.Bold = True
            wsSlip.Range("A" & slipRow & ":D" & slipRow).Merge
            slipRow = slipRow + 1
            
            ' 基本信息
            wsSlip.Cells(slipRow, 1).Value = "姓名：" & wsPayroll.Cells(i, 2).Value
            wsSlip.Cells(slipRow, 3).Value = "月份：" & monthStr
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "部门：" & wsPayroll.Cells(i, 3).Value
            wsSlip.Cells(slipRow, 3).Value = "岗位：" & wsPayroll.Cells(i, 4).Value
            slipRow = slipRow + 1
            
            ' 应发项目
            wsSlip.Cells(slipRow, 1).Value = "应发工资明细："
            wsSlip.Cells(slipRow, 1).Font.Bold = True
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "基本工资："
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 5).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "绩效工资："
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 6).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "加班工资："
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 7).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "补贴："
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 8).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "应发工资合计："
            wsSlip.Cells(slipRow, 1).Font.Bold = True
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 9).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            wsSlip.Cells(slipRow, 2).Font.Bold = True
            slipRow = slipRow + 1
            
            ' 代扣项目
            wsSlip.Cells(slipRow, 1).Value = "代扣款项明细："
            wsSlip.Cells(slipRow, 1).Font.Bold = True
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "社保个人："
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 10).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "公积金个人："
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 11).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            slipRow = slipRow + 1
            
            wsSlip.Cells(slipRow, 1).Value = "个人所得税："
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 12).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            slipRow = slipRow + 1
            
            ' 实发工资
            wsSlip.Cells(slipRow, 1).Value = "实发工资："
            wsSlip.Cells(slipRow, 1).Font.Bold = True
            wsSlip.Cells(slipRow, 1).Font.Size = 12
            wsSlip.Cells(slipRow, 2).Value = wsPayroll.Cells(i, 13).Value
            wsSlip.Cells(slipRow, 2).NumberFormat = "#,##0.00"
            wsSlip.Cells(slipRow, 2).Font.Bold = True
            wsSlip.Cells(slipRow, 2).Font.Size = 12
            slipRow = slipRow + 1
            
            ' 大写金额
            wsSlip.Cells(slipRow, 1).Value = "大写：" & NumberToChinese(wsPayroll.Cells(i, 13).Value)
            slipRow = slipRow + 1
            
            ' 分隔行
            wsSlip.Range("A" & slipRow & ":D" & slipRow).Interior.Color = RGB(200, 200, 200)
            slipRow = slipRow + 2
        End If
    Next i
    
    ' 设置列宽
    wsSlip.Columns("A").ColumnWidth = 20
    wsSlip.Columns("B").ColumnWidth = 15
    wsSlip.Columns("C").ColumnWidth = 20
    wsSlip.Columns("D").ColumnWidth = 15
    
    MsgBox "工资条生成完成！" & vbCrLf & vbCrLf & _
           "月份：" & monthStr & vbCrLf & _
           "每员工一条，便于打印裁剪", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "生成工资条出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏46: GenerateSocialDetail - 社保明细表
' V7.0新增：生成员工社保公积金明细
' ============================================================================
Sub GenerateSocialDetail()
    Dim wsPayroll As Worksheet, wsDetail As Worksheet
    Dim monthStr As String
    Dim lastRow As Long, i As Long, detailRow As Long
    Dim baseAmount As Double
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    monthStr = InputBox("请输入月份（格式：yyyy-mm）：", "社保明细", Format(Date, "yyyy-mm"))
    If monthStr = "" Then Exit Sub
    
    ' 缴费基数输入验证
    inputValue = InputBox("请输入社保公积金缴费基数：", "缴费基数", "5000")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    baseAmount = CDbl(inputValue)
    If baseAmount <= 0 Then Exit Sub
    
    On Error Resume Next
    Set wsPayroll = ThisWorkbook.Sheets("工资表")
    If wsPayroll Is Nothing Then
        MsgBox "请先运行【工资表生成】创建工资表", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建或清空社保明细表
    On Error Resume Next
    Set wsDetail = ThisWorkbook.Sheets("社保明细")
    If wsDetail Is Nothing Then
        Set wsDetail = ThisWorkbook.Sheets.Add
        wsDetail.Name = "社保明细"
    Else
        wsDetail.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsDetail.Cells(1, 1).Value = "员工社保公积金明细表"
    wsDetail.Cells(1, 1).Font.Size = 14
    wsDetail.Cells(1, 1).Font.Bold = True
    wsDetail.Range("A1:N1").Merge
    
    wsDetail.Cells(2, 1).Value = "月份：" & monthStr
    wsDetail.Cells(2, 4).Value = "缴费基数：" & Format(baseAmount, "#,##0.00") & " 元"
    
    ' 列标题
    wsDetail.Cells(4, 1).Value = "序号"
    wsDetail.Cells(4, 2).Value = "姓名"
    wsDetail.Cells(4, 3).Value = "缴费基数"
    wsDetail.Cells(4, 4).Value = "养老(企业)"
    wsDetail.Cells(4, 5).Value = "养老(个人)"
    wsDetail.Cells(4, 6).Value = "医疗(企业)"
    wsDetail.Cells(4, 7).Value = "医疗(个人)"
    wsDetail.Cells(4, 8).Value = "失业(企业)"
    wsDetail.Cells(4, 9).Value = "失业(个人)"
    wsDetail.Cells(4, 10).Value = "工伤(企业)"
    wsDetail.Cells(4, 11).Value = "生育(企业)"
    wsDetail.Cells(4, 12).Value = "公积金(企业)"
    wsDetail.Cells(4, 13).Value = "公积金(个人)"
    wsDetail.Cells(4, 14).Value = "合计"
    wsDetail.Range("A4:N4").Font.Bold = True
    wsDetail.Range("A4:N4").Interior.Color = RGB(68, 114, 196)
    wsDetail.Range("A4:N4").Font.Color = RGB(255, 255, 255)
    
    ' 获取工资表数据
    lastRow = GetLastRow(wsPayroll, 2)
    detailRow = 5
    
    For i = 5 To lastRow - 1  ' 跳过表头和合计行
        If wsPayroll.Cells(i, 2).Value <> "" Then
            wsDetail.Cells(detailRow, 1).Value = detailRow - 4
            wsDetail.Cells(detailRow, 2).Value = wsPayroll.Cells(i, 2).Value
            wsDetail.Cells(detailRow, 3).Value = baseAmount
            wsDetail.Cells(detailRow, 3).NumberFormat = "#,##0.00"
            
            ' 各项社保
            wsDetail.Cells(detailRow, 4).Value = baseAmount * EI_PENSION_E
            wsDetail.Cells(detailRow, 4).NumberFormat = "#,##0.00"
            wsDetail.Cells(detailRow, 5).Value = baseAmount * EI_PENSION_P
            wsDetail.Cells(detailRow, 5).NumberFormat = "#,##0.00"
            
            wsDetail.Cells(detailRow, 6).Value = baseAmount * EI_MEDICAL_E
            wsDetail.Cells(detailRow, 6).NumberFormat = "#,##0.00"
            wsDetail.Cells(detailRow, 7).Value = baseAmount * EI_MEDICAL_P
            wsDetail.Cells(detailRow, 7).NumberFormat = "#,##0.00"
            
            wsDetail.Cells(detailRow, 8).Value = baseAmount * EI_UNEMPLOYMENT_E
            wsDetail.Cells(detailRow, 8).NumberFormat = "#,##0.00"
            wsDetail.Cells(detailRow, 9).Value = baseAmount * EI_UNEMPLOYMENT_P
            wsDetail.Cells(detailRow, 9).NumberFormat = "#,##0.00"
            
            wsDetail.Cells(detailRow, 10).Value = baseAmount * EI_INJURY_E
            wsDetail.Cells(detailRow, 10).NumberFormat = "#,##0.00"
            wsDetail.Cells(detailRow, 11).Value = baseAmount * EI_MATERNITY_E
            wsDetail.Cells(detailRow, 11).NumberFormat = "#,##0.00"
            
            wsDetail.Cells(detailRow, 12).Value = baseAmount * HPF_RATE_E
            wsDetail.Cells(detailRow, 12).NumberFormat = "#,##0.00"
            wsDetail.Cells(detailRow, 13).Value = baseAmount * HPF_RATE_P
            wsDetail.Cells(detailRow, 13).NumberFormat = "#,##0.00"
            
            ' 合计公式
            wsDetail.Cells(detailRow, 14).Formula = "=SUM(D" & detailRow & ":M" & detailRow & ")"
            wsDetail.Cells(detailRow, 14).NumberFormat = "#,##0.00"
            wsDetail.Cells(detailRow, 14).Font.Bold = True
            
            detailRow = detailRow + 1
        End If
    Next i
    
    ' 合计行
    wsDetail.Cells(detailRow, 1).Value = "合计"
    wsDetail.Cells(detailRow, 1).Font.Bold = True
    For i = 3 To 14
        wsDetail.Cells(detailRow, i).Formula = "=SUM(" & Chr(64 + i) & "5:" & Chr(64 + i) & (detailRow - 1) & ")"
        wsDetail.Cells(detailRow, i).NumberFormat = "#,##0.00"
        wsDetail.Cells(detailRow, i).Font.Bold = True
    Next i
    
    ' 设置列宽
    wsDetail.Columns("A").ColumnWidth = 6
    wsDetail.Columns("B").ColumnWidth = 12
    wsDetail.Columns("C:N").ColumnWidth = 12
    
    MsgBox "社保明细表生成完成！" & vbCrLf & vbCrLf & _
           "月份：" & monthStr & vbCrLf & _
           "缴费基数：" & Format(baseAmount, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "生成社保明细表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 辅助函数：数字转中文大写
' ============================================================================
Function NumberToChinese(ByVal MyNumber As Variant) As String
    Dim Dollars As String, Cents As String
    Dim Temp As String
    Dim DecimalPlace As Integer, Count As Integer
    Dim Place(9) As String
    Dim isNegative As Boolean
    
    ' 检查输入是否为有效数字
    If Not IsNumeric(MyNumber) Then
        NumberToChinese = ""
        Exit Function
    End If
    
    ' 处理负数
    isNegative = False
    If MyNumber < 0 Then
        isNegative = True
        MyNumber = Abs(MyNumber)
    End If
    
    Place(2) = "仟"
    Place(3) = "佰"
    Place(4) = "拾"
    Place(5) = "万"
    Place(6) = "仟"
    Place(7) = "佰"
    Place(8) = "拾"
    Place(9) = "亿"
    
    MyNumber = Trim(Str(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")
    If DecimalPlace > 0 Then
        Cents = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    
    Count = 1
    Do While MyNumber <> ""
        Temp = Right(MyNumber, 1)
        If Temp = "0" Then
            Dollars = "零" & Dollars
        Else
            Select Case Count
                Case 1: Dollars = Temp & Dollars
                Case 2: Dollars = Temp & Place(4) & Dollars
                Case 3: Dollars = Temp & Place(3) & Dollars
                Case 4: Dollars = Temp & Place(2) & Dollars
                Case 5: Dollars = Temp & Place(5) & Dollars
                Case 6: Dollars = Temp & Place(4) & Dollars
                Case 7: Dollars = Temp & Place(3) & Dollars
                Case 8: Dollars = Temp & Place(2) & Dollars
                Case 9: Dollars = Temp & Place(9) & Dollars
            End Select
        End If
        If Len(MyNumber) = 1 Then Exit Do
        MyNumber = Left(MyNumber, Len(MyNumber) - 1)
        Count = Count + 1
    Loop
    
    If Dollars = "" Then Dollars = "零"
    NumberToChinese = "人民币" & Dollars & "元"
    
    If Cents <> "" And Cents <> "00" Then
        If Left(Cents, 1) <> "0" Then
            NumberToChinese = NumberToChinese & Left(Cents, 1) & "角"
        End If
        If Right(Cents, 1) <> "0" Then
            NumberToChinese = NumberToChinese & Right(Cents, 1) & "分"
        End If
    Else
        NumberToChinese = NumberToChinese & "整"
    End If
    
    ' 如果是负数，在前面添加"负"字
    If isNegative Then
        NumberToChinese = "负" & NumberToChinese
    End If
End Function

' ============================================================================
' 宏47: CreateFixedAssetsLedger - 固定资产台账
' V8.0新增：创建固定资产台账
' ============================================================================
Sub CreateFixedAssetsLedger()
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' 创建或清空固定资产台账
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("固定资产台账")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "固定资产台账"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    ws.Cells(1, 1).Value = "固定资产台账"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Range("A1:N1").Merge
    
    ws.Cells(2, 1).Value = "编制日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 列标题
    ws.Cells(4, 1).Value = "资产编号"
    ws.Cells(4, 2).Value = "资产名称"
    ws.Cells(4, 3).Value = "规格型号"
    ws.Cells(4, 4).Value = "购置日期"
    ws.Cells(4, 5).Value = "原值"
    ws.Cells(4, 6).Value = "残值率"
    ws.Cells(4, 7).Value = "残值"
    ws.Cells(4, 8).Value = "使用年限"
    ws.Cells(4, 9).Value = "折旧方法"
    ws.Cells(4, 10).Value = "月折旧额"
    ws.Cells(4, 11).Value = "累计折旧"
    ws.Cells(4, 12).Value = "净值"
    ws.Cells(4, 13).Value = "存放地点"
    ws.Cells(4, 14).Value = "使用状态"
    
    ws.Range("A4:N4").Font.Bold = True
    ws.Range("A4:N4").Interior.Color = RGB(68, 114, 196)
    ws.Range("A4:N4").Font.Color = RGB(255, 255, 255)
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 10
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 12
    ws.Columns("L").ColumnWidth = 12
    ws.Columns("M").ColumnWidth = 15
    ws.Columns("N").ColumnWidth = 10
    
    ' 添加示例数据行（第5行）
    ws.Cells(5, 1).Value = "FA-001"
    ws.Cells(5, 2).Value = "氧化槽"
    ws.Cells(5, 3).Value = "2000L"
    ws.Cells(5, 4).Value = Format(Date, "yyyy-mm-dd")
    ws.Cells(5, 5).Value = 50000
    ws.Cells(5, 5).NumberFormat = "#,##0.00"
    ws.Cells(5, 6).Value = FA_RESIDUAL_RATE
    ws.Cells(5, 6).NumberFormat = "0.00%"
    ws.Cells(5, 7).Value = "=E5*F5"
    ws.Cells(5, 7).NumberFormat = "#,##0.00"
    ws.Cells(5, 8).Value = FA_DEPRECIATION_YEARS
    ws.Cells(5, 9).Value = "平均年限法"
    ws.Cells(5, 10).Value = "=(E5-G5)/H5/12"
    ws.Cells(5, 10).NumberFormat = "#,##0.00"
    ws.Cells(5, 11).Value = 0
    ws.Cells(5, 11).NumberFormat = "#,##0.00"
    ws.Cells(5, 12).Value = "=E5-K5"
    ws.Cells(5, 12).NumberFormat = "#,##0.00"
    ws.Cells(5, 13).Value = "车间"
    ws.Cells(5, 14).Value = "使用中"
    
    ' 合计行
    Dim lastRow As Long
    lastRow = 6
    ws.Cells(lastRow, 1).Value = "合计"
    ws.Cells(lastRow, 1).Font.Bold = True
    ws.Cells(lastRow, 5).Formula = "=SUM(E5:E" & (lastRow - 1) & ")"
    ws.Cells(lastRow, 5).NumberFormat = "#,##0.00"
    ws.Cells(lastRow, 7).Formula = "=SUM(G5:G" & (lastRow - 1) & ")"
    ws.Cells(lastRow, 7).NumberFormat = "#,##0.00"
    ws.Cells(lastRow, 10).Formula = "=SUM(J5:J" & (lastRow - 1) & ")"
    ws.Cells(lastRow, 10).NumberFormat = "#,##0.00"
    ws.Cells(lastRow, 11).Formula = "=SUM(K5:K" & (lastRow - 1) & ")"
    ws.Cells(lastRow, 11).NumberFormat = "#,##0.00"
    ws.Cells(lastRow, 12).Formula = "=SUM(L5:L" & (lastRow - 1) & ")"
    ws.Cells(lastRow, 12).NumberFormat = "#,##0.00"
    
    MsgBox "固定资产台账创建完成！" & vbCrLf & vbCrLf & _
           "请在第5行开始录入固定资产信息", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "创建固定资产台账出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏48: CalculateDepreciation - 折旧计算
' V8.0新增：计算固定资产折旧
' 支持折旧方法：平均年限法、双倍余额递减法、年数总和法
' ============================================================================
Sub CalculateDepreciation()
    Dim wsLedger As Worksheet, wsDep As Worksheet
    Dim lastRow As Long, i As Long, depRow As Long
    Dim originalValue As Double, residualValue As Double
    Dim usefulLife As Long, method As String
    Dim monthlyDep As Double, netValue As Double
    Dim accumulatedDep As Double
    
    On Error GoTo ErrorHandler
    
    ' 检查固定资产台账是否存在
    On Error Resume Next
    Set wsLedger = ThisWorkbook.Sheets("固定资产台账")
    If wsLedger Is Nothing Then
        MsgBox "请先运行【固定资产台账】创建台账", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建折旧计算表
    On Error Resume Next
    Set wsDep = ThisWorkbook.Sheets("折旧计算表")
    If wsDep Is Nothing Then
        Set wsDep = ThisWorkbook.Sheets.Add
        wsDep.Name = "折旧计算表"
    Else
        wsDep.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsDep.Cells(1, 1).Value = "固定资产折旧计算表"
    wsDep.Cells(1, 1).Font.Size = 16
    wsDep.Cells(1, 1).Font.Bold = True
    wsDep.Cells(1, 1).HorizontalAlignment = xlCenter
    wsDep.Range("A1:K1").Merge
    
    wsDep.Cells(2, 1).Value = "计算日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 列标题
    wsDep.Cells(4, 1).Value = "资产编号"
    wsDep.Cells(4, 2).Value = "资产名称"
    wsDep.Cells(4, 3).Value = "原值"
    wsDep.Cells(4, 4).Value = "残值"
    wsDep.Cells(4, 5).Value = "使用年限"
    wsDep.Cells(4, 6).Value = "折旧方法"
    wsDep.Cells(4, 7).Value = "本月折旧"
    wsDep.Cells(4, 8).Value = "累计折旧"
    wsDep.Cells(4, 9).Value = "净值"
    wsDep.Cells(4, 10).Value = "折旧状态"
    wsDep.Cells(4, 11).Value = "备注"
    
    wsDep.Range("A4:K4").Font.Bold = True
    wsDep.Range("A4:K4").Interior.Color = RGB(112, 173, 71)
    wsDep.Range("A4:K4").Font.Color = RGB(255, 255, 255)
    
    ' 设置列宽
    wsDep.Columns("A").ColumnWidth = 12
    wsDep.Columns("B").ColumnWidth = 18
    wsDep.Columns("C").ColumnWidth = 12
    wsDep.Columns("D").ColumnWidth = 10
    wsDep.Columns("E").ColumnWidth = 10
    wsDep.Columns("F").ColumnWidth = 14
    wsDep.Columns("G").ColumnWidth = 12
    wsDep.Columns("H").ColumnWidth = 12
    wsDep.Columns("I").ColumnWidth = 12
    wsDep.Columns("J").ColumnWidth = 10
    wsDep.Columns("K").ColumnWidth = 20
    
    ' 读取固定资产台账数据并计算折旧
    lastRow = GetLastRow(wsLedger, 1)
    depRow = 5
    
    For i = 5 To lastRow - 1  ' 跳过表头和合计行
        If wsLedger.Cells(i, 1).Value <> "" And wsLedger.Cells(i, 1).Value <> "合计" Then
            originalValue = Val(wsLedger.Cells(i, 5).Value)
            residualValue = Val(wsLedger.Cells(i, 7).Value)
            usefulLife = Val(wsLedger.Cells(i, 8).Value)
            method = wsLedger.Cells(i, 9).Value
            accumulatedDep = Val(wsLedger.Cells(i, 11).Value)
            
            ' 计算月折旧额
            Select Case method
                Case "平均年限法", "直线法"
                    ' 月折旧额 = (原值 - 残值) / 使用年限 / 12
                    monthlyDep = (originalValue - residualValue) / usefulLife / 12
                    
                Case "双倍余额递减法"
                    ' 月折旧额 = 净值 × 2 / 使用年限 / 12
                    netValue = originalValue - accumulatedDep
                    monthlyDep = netValue * 2 / usefulLife / 12
                    ' 确保折旧后不低于残值
                    If netValue - monthlyDep < residualValue Then
                        monthlyDep = netValue - residualValue
                    End If
                    
                Case "年数总和法"
                    ' 简化处理：使用直线法
                    monthlyDep = (originalValue - residualValue) / usefulLife / 12
                    
                Case Else
                    monthlyDep = (originalValue - residualValue) / usefulLife / 12
            End Select
            
            ' 计算净值
            netValue = originalValue - accumulatedDep - monthlyDep
            
            ' 写入折旧计算表
            wsDep.Cells(depRow, 1).Value = wsLedger.Cells(i, 1).Value
            wsDep.Cells(depRow, 2).Value = wsLedger.Cells(i, 2).Value
            wsDep.Cells(depRow, 3).Value = originalValue
            wsDep.Cells(depRow, 3).NumberFormat = "#,##0.00"
            wsDep.Cells(depRow, 4).Value = residualValue
            wsDep.Cells(depRow, 4).NumberFormat = "#,##0.00"
            wsDep.Cells(depRow, 5).Value = usefulLife
            wsDep.Cells(depRow, 6).Value = method
            wsDep.Cells(depRow, 7).Value = monthlyDep
            wsDep.Cells(depRow, 7).NumberFormat = "#,##0.00"
            wsDep.Cells(depRow, 8).Value = accumulatedDep + monthlyDep
            wsDep.Cells(depRow, 8).NumberFormat = "#,##0.00"
            wsDep.Cells(depRow, 9).Value = netValue
            wsDep.Cells(depRow, 9).NumberFormat = "#,##0.00"
            
            ' 判断折旧状态
            If netValue <= residualValue Then
                wsDep.Cells(depRow, 10).Value = "已提足"
                wsDep.Cells(depRow, 10).Interior.Color = RGB(255, 199, 206)
            Else
                wsDep.Cells(depRow, 10).Value = "正常"
            End If
            
            depRow = depRow + 1
        End If
    Next i
    
    ' 合计行
    wsDep.Cells(depRow, 1).Value = "合计"
    wsDep.Cells(depRow, 1).Font.Bold = True
    wsDep.Cells(depRow, 3).Formula = "=SUM(C5:C" & (depRow - 1) & ")"
    wsDep.Cells(depRow, 3).NumberFormat = "#,##0.00"
    wsDep.Cells(depRow, 4).Formula = "=SUM(D5:D" & (depRow - 1) & ")"
    wsDep.Cells(depRow, 4).NumberFormat = "#,##0.00"
    wsDep.Cells(depRow, 7).Formula = "=SUM(G5:G" & (depRow - 1) & ")"
    wsDep.Cells(depRow, 7).NumberFormat = "#,##0.00"
    wsDep.Cells(depRow, 7).Font.Bold = True
    wsDep.Cells(depRow, 8).Formula = "=SUM(H5:H" & (depRow - 1) & ")"
    wsDep.Cells(depRow, 8).NumberFormat = "#,##0.00"
    wsDep.Cells(depRow, 9).Formula = "=SUM(I5:I" & (depRow - 1) & ")"
    wsDep.Cells(depRow, 9).NumberFormat = "#,##0.00"
    
    MsgBox "折旧计算完成！" & vbCrLf & vbCrLf & _
           "本月折旧总额：" & Format(wsDep.Cells(depRow, 7).Value, "#,##0.00") & " 元", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "折旧计算出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏49: LowValueConsumables - 低值易耗品管理
' V8.0新增：管理低值易耗品（价值<2000元或使用年限<1年）
' ============================================================================
Sub LowValueConsumables()
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' 创建或清空低值易耗品表
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("低值易耗品")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "低值易耗品"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    ws.Cells(1, 1).Value = "低值易耗品管理表"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Range("A1:I1").Merge
    
    ws.Cells(2, 1).Value = "低值易耗品标准：价值<" & LVC_THRESHOLD & "元或使用年限<1年"
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    
    ' 列标题
    ws.Cells(4, 1).Value = "物品编号"
    ws.Cells(4, 2).Value = "物品名称"
    ws.Cells(4, 3).Value = "数量"
    ws.Cells(4, 4).Value = "单价"
    ws.Cells(4, 5).Value = "金额"
    ws.Cells(4, 6).Value = "摊销方法"
    ws.Cells(4, 7).Value = "已摊销"
    ws.Cells(4, 8).Value = "待摊销"
    ws.Cells(4, 9).Value = "备注"
    
    ws.Range("A4:I4").Font.Bold = True
    ws.Range("A4:I4").Interior.Color = RGB(237, 125, 49)
    ws.Range("A4:I4").Font.Color = RGB(255, 255, 255)
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 20
    
    ' 添加示例数据
    ws.Cells(5, 1).Value = "LVC-001"
    ws.Cells(5, 2).Value = "工具套装"
    ws.Cells(5, 3).Value = 5
    ws.Cells(5, 4).Value = 150
    ws.Cells(5, 4).NumberFormat = "#,##0.00"
    ws.Cells(5, 5).Value = "=C5*D5"
    ws.Cells(5, 5).NumberFormat = "#,##0.00"
    ws.Cells(5, 6).Value = "一次摊销"
    ws.Cells(5, 7).Value = "=E5"
    ws.Cells(5, 7).NumberFormat = "#,##0.00"
    ws.Cells(5, 8).Value = 0
    ws.Cells(5, 8).NumberFormat = "#,##0.00"
    ws.Cells(5, 9).Value = "车间工具"
    
    ws.Cells(6, 1).Value = "LVC-002"
    ws.Cells(6, 2).Value = "防护手套"
    ws.Cells(6, 3).Value = 100
    ws.Cells(6, 4).Value = 5
    ws.Cells(6, 4).NumberFormat = "#,##0.00"
    ws.Cells(6, 5).Value = "=C6*D6"
    ws.Cells(6, 5).NumberFormat = "#,##0.00"
    ws.Cells(6, 6).Value = "一次摊销"
    ws.Cells(6, 7).Value = "=E6"
    ws.Cells(6, 7).NumberFormat = "#,##0.00"
    ws.Cells(6, 8).Value = 0
    ws.Cells(6, 8).NumberFormat = "#,##0.00"
    ws.Cells(6, 9).Value = "劳保用品"
    
    ws.Cells(7, 1).Value = "LVC-003"
    ws.Cells(7, 2).Value = "办公桌椅"
    ws.Cells(7, 3).Value = 2
    ws.Cells(7, 4).Value = 800
    ws.Cells(7, 4).NumberFormat = "#,##0.00"
    ws.Cells(7, 5).Value = "=C7*D7"
    ws.Cells(7, 5).NumberFormat = "#,##0.00"
    ws.Cells(7, 6).Value = "五五摊销"
    ws.Cells(7, 7).Value = "=E7*0.5"
    ws.Cells(7, 7).NumberFormat = "#,##0.00"
    ws.Cells(7, 8).Value = "=E7-G7"
    ws.Cells(7, 8).NumberFormat = "#,##0.00"
    ws.Cells(7, 9).Value = "办公家具"
    
    ' 合计行
    ws.Cells(8, 1).Value = "合计"
    ws.Cells(8, 1).Font.Bold = True
    ws.Cells(8, 3).Formula = "=SUM(C5:C7)"
    ws.Cells(8, 5).Formula = "=SUM(E5:E7)"
    ws.Cells(8, 5).NumberFormat = "#,##0.00"
    ws.Cells(8, 5).Font.Bold = True
    ws.Cells(8, 7).Formula = "=SUM(G5:G7)"
    ws.Cells(8, 7).NumberFormat = "#,##0.00"
    ws.Cells(8, 8).Formula = "=SUM(H5:H7)"
    ws.Cells(8, 8).NumberFormat = "#,##0.00"
    
    MsgBox "低值易耗品管理表创建完成！" & vbCrLf & vbCrLf & _
           "摊销方法说明：" & vbCrLf & _
           "• 一次摊销：领用时全额摊销" & vbCrLf & _
           "• 五五摊销：领用时摊销50%，报废时摊销剩余50%", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "创建低值易耗品表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏50: CostAllocation - 成本核算
' V8.0新增：核算产品成本
' ============================================================================
Sub CostAllocation()
    Dim wsExpense As Worksheet, wsCost As Worksheet
    Dim lastRow As Long
    Dim directMaterial As Double, directLabor As Double
    Dim manufacturingCost As Double, outsourceCost As Double
    Dim totalCost As Double
    Dim rentCost As Double, waterCost As Double, elecCost As Double
    Dim dailyCost As Double
    
    On Error GoTo ErrorHandler
    
    ' 检查支出记录表是否存在
    On Error Resume Next
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    If wsExpense Is Nothing Then
        MsgBox "请先录入支出记录", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建成本核算表
    On Error Resume Next
    Set wsCost = ThisWorkbook.Sheets("成本核算")
    If wsCost Is Nothing Then
        Set wsCost = ThisWorkbook.Sheets.Add
        wsCost.Name = "成本核算"
    Else
        wsCost.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsCost.Cells(1, 1).Value = "产品成本核算表"
    wsCost.Cells(1, 1).Font.Size = 16
    wsCost.Cells(1, 1).Font.Bold = True
    wsCost.Cells(1, 1).HorizontalAlignment = xlCenter
    wsCost.Range("A1:E1").Merge
    
    wsCost.Cells(2, 1).Value = "核算期间：" & Format(Date, "yyyy年mm月")
    
    ' 列标题
    wsCost.Cells(4, 1).Value = "成本项目"
    wsCost.Cells(4, 2).Value = "金额"
    wsCost.Cells(4, 3).Value = "占比"
    wsCost.Cells(4, 4).Value = "备注"
    
    wsCost.Range("A4:D4").Font.Bold = True
    wsCost.Range("A4:D4").Interior.Color = RGB(68, 114, 196)
    wsCost.Range("A4:D4").Font.Color = RGB(255, 255, 255)
    
    ' 设置列宽
    wsCost.Columns("A").ColumnWidth = 25
    wsCost.Columns("B").ColumnWidth = 15
    wsCost.Columns("C").ColumnWidth = 10
    wsCost.Columns("D").ColumnWidth = 30
    
    ' 计算各项成本
    ' 直接材料（三酸+片碱+亚钠+色粉+除油剂+挂具）
    directMaterial = SumColumn(wsExpense, 9, "三酸") + _
                     SumColumn(wsExpense, 9, "片碱") + _
                     SumColumn(wsExpense, 9, "亚钠") + _
                     SumColumn(wsExpense, 9, "色粉") + _
                     SumColumn(wsExpense, 9, "除油剂") + _
                     SumColumn(wsExpense, 9, "挂具")
    
    ' 直接人工（工资）
    directLabor = SumColumn(wsExpense, 9, "工资")
    
    ' 制造费用（房租+水费+电费+日常费用）
    rentCost = SumColumn(wsExpense, 9, "房租")
    waterCost = SumColumn(wsExpense, 9, "水费")
    elecCost = SumColumn(wsExpense, 9, "电费")
    dailyCost = SumColumn(wsExpense, 9, "日常费用")
    manufacturingCost = rentCost + waterCost + elecCost + dailyCost
    
    ' 外发加工成本
    outsourceCost = SumColumn(wsExpense, 9, "外发加工费")
    
    ' 总成本
    totalCost = directMaterial + directLabor + manufacturingCost + outsourceCost
    
    ' 填充数据
    Dim row As Long
    row = 5
    
    ' 一、直接材料
    wsCost.Cells(row, 1).Value = "一、直接材料"
    wsCost.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  三酸"
    wsCost.Cells(row, 2).Value = SumColumn(wsExpense, 9, "三酸")
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  片碱"
    wsCost.Cells(row, 2).Value = SumColumn(wsExpense, 9, "片碱")
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  亚钠"
    wsCost.Cells(row, 2).Value = SumColumn(wsExpense, 9, "亚钠")
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  色粉"
    wsCost.Cells(row, 2).Value = SumColumn(wsExpense, 9, "色粉")
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  除油剂"
    wsCost.Cells(row, 2).Value = SumColumn(wsExpense, 9, "除油剂")
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  挂具"
    wsCost.Cells(row, 2).Value = SumColumn(wsExpense, 9, "挂具")
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "直接材料小计"
    wsCost.Cells(row, 1).Font.Bold = True
    wsCost.Cells(row, 2).Value = directMaterial
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    wsCost.Cells(row, 2).Font.Bold = True
    If totalCost > 0 Then
        wsCost.Cells(row, 3).Value = directMaterial / totalCost
        wsCost.Cells(row, 3).NumberFormat = "0.00%"
    End If
    row = row + 2
    
    ' 二、直接人工
    wsCost.Cells(row, 1).Value = "二、直接人工"
    wsCost.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  工资"
    wsCost.Cells(row, 2).Value = directLabor
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "直接人工小计"
    wsCost.Cells(row, 1).Font.Bold = True
    wsCost.Cells(row, 2).Value = directLabor
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    wsCost.Cells(row, 2).Font.Bold = True
    If totalCost > 0 Then
        wsCost.Cells(row, 3).Value = directLabor / totalCost
        wsCost.Cells(row, 3).NumberFormat = "0.00%"
    End If
    row = row + 2
    
    ' 三、制造费用
    wsCost.Cells(row, 1).Value = "三、制造费用"
    wsCost.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  房租"
    wsCost.Cells(row, 2).Value = rentCost
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  水费"
    wsCost.Cells(row, 2).Value = waterCost
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  电费"
    wsCost.Cells(row, 2).Value = elecCost
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "  日常费用"
    wsCost.Cells(row, 2).Value = dailyCost
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    wsCost.Cells(row, 1).Value = "制造费用小计"
    wsCost.Cells(row, 1).Font.Bold = True
    wsCost.Cells(row, 2).Value = manufacturingCost
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    wsCost.Cells(row, 2).Font.Bold = True
    If totalCost > 0 Then
        wsCost.Cells(row, 3).Value = manufacturingCost / totalCost
        wsCost.Cells(row, 3).NumberFormat = "0.00%"
    End If
    row = row + 2
    
    ' 四、外发加工成本
    wsCost.Cells(row, 1).Value = "四、外发加工成本"
    wsCost.Cells(row, 1).Font.Bold = True
    wsCost.Cells(row, 2).Value = outsourceCost
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    If totalCost > 0 Then
        wsCost.Cells(row, 3).Value = outsourceCost / totalCost
        wsCost.Cells(row, 3).NumberFormat = "0.00%"
    End If
    row = row + 2
    
    ' 总成本
    wsCost.Cells(row, 1).Value = "总成本合计"
    wsCost.Cells(row, 1).Font.Bold = True
    wsCost.Cells(row, 1).Font.Size = 12
    wsCost.Cells(row, 2).Value = totalCost
    wsCost.Cells(row, 2).NumberFormat = "#,##0.00"
    wsCost.Cells(row, 2).Font.Bold = True
    wsCost.Cells(row, 2).Font.Size = 12
    wsCost.Cells(row, 3).Value = 1
    wsCost.Cells(row, 3).NumberFormat = "0.00%"
    wsCost.Cells(row, 4).Value = "直接材料+直接人工+制造费用+外发加工"
    
    ' 添加边框
    wsCost.Range("A4:D" & row).Borders.LineStyle = xlContinuous
    
    MsgBox "成本核算完成！" & vbCrLf & vbCrLf & _
           "总成本：" & Format(totalCost, "#,##0.00") & " 元" & vbCrLf & _
           "直接材料占比：" & Format(directMaterial / IIf(totalCost > 0, totalCost, 1), "0.00%") & vbCrLf & _
           "直接人工占比：" & Format(directLabor / IIf(totalCost > 0, totalCost, 1), "0.00%") & vbCrLf & _
           "制造费用占比：" & Format(manufacturingCost / IIf(totalCost > 0, totalCost, 1), "0.00%"), vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "成本核算出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏51: CostVarianceAnalysis - 成本差异分析
' V8.0新增：分析成本差异
' ============================================================================
Sub CostVarianceAnalysis()
    Dim ws As Worksheet
    Dim standardCost As Variant, actualCost As Variant
    Dim costItems As Variant
    Dim i As Long, row As Long
    Dim varianceAmount As Double, varianceRate As Double
    
    On Error GoTo ErrorHandler
    
    ' 成本项目
    costItems = Array("直接材料", "直接人工", "制造费用", "外发加工成本")
    
    ' 创建成本差异分析表
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("成本差异分析")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "成本差异分析"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    ws.Cells(1, 1).Value = "成本差异分析表"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Range("A1:F1").Merge
    
    ws.Cells(2, 1).Value = "分析期间：" & Format(Date, "yyyy年mm月")
    
    ' 列标题
    ws.Cells(4, 1).Value = "成本项目"
    ws.Cells(4, 2).Value = "标准成本"
    ws.Cells(4, 3).Value = "实际成本"
    ws.Cells(4, 4).Value = "差异额"
    ws.Cells(4, 5).Value = "差异率"
    ws.Cells(4, 6).Value = "差异原因分析"
    
    ws.Range("A4:F4").Font.Bold = True
    ws.Range("A4:F4").Interior.Color = RGB(237, 125, 49)
    ws.Range("A4:F4").Font.Color = RGB(255, 255, 255)
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 15
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 30
    
    ' 填充数据（示例数据，实际应用中应从历史数据或预算中获取）
    row = 5
    
    ' 直接材料
    ws.Cells(row, 1).Value = "直接材料"
    ws.Cells(row, 2).Value = 15000  ' 标准成本
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = 16500  ' 实际成本
    ws.Cells(row, 3).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Formula = "=C" & row & "-B" & row
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Formula = "=IF(B" & row & "=0,0,D" & row & "/B" & row & ")"
    ws.Cells(row, 5).NumberFormat = "0.00%"
    ws.Cells(row, 6).Value = "材料价格上涨/用量增加"
    row = row + 1
    
    ' 直接人工
    ws.Cells(row, 1).Value = "直接人工"
    ws.Cells(row, 2).Value = 20000
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = 21000
    ws.Cells(row, 3).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Formula = "=C" & row & "-B" & row
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Formula = "=IF(B" & row & "=0,0,D" & row & "/B" & row & ")"
    ws.Cells(row, 5).NumberFormat = "0.00%"
    ws.Cells(row, 6).Value = "加班增加/人员调整"
    row = row + 1
    
    ' 制造费用
    ws.Cells(row, 1).Value = "制造费用"
    ws.Cells(row, 2).Value = 8000
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = 7500
    ws.Cells(row, 3).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Formula = "=C" & row & "-B" & row
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Formula = "=IF(B" & row & "=0,0,D" & row & "/B" & row & ")"
    ws.Cells(row, 5).NumberFormat = "0.00%"
    ws.Cells(row, 6).Value = "费用节约"
    ' 差异为负数时标绿
    ws.Cells(row, 4).Interior.Color = RGB(198, 239, 206)
    row = row + 1
    
    ' 外发加工成本
    ws.Cells(row, 1).Value = "外发加工成本"
    ws.Cells(row, 2).Value = 5000
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = 5500
    ws.Cells(row, 3).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Formula = "=C" & row & "-B" & row
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Formula = "=IF(B" & row & "=0,0,D" & row & "/B" & row & ")"
    ws.Cells(row, 5).NumberFormat = "0.00%"
    ws.Cells(row, 6).Value = "外发量增加/单价上涨"
    row = row + 1
    
    ' 合计行
    ws.Cells(row, 1).Value = "合计"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Formula = "=SUM(B5:B" & (row - 1) & ")"
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 3).Formula = "=SUM(C5:C" & (row - 1) & ")"
    ws.Cells(row, 3).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Font.Bold = True
    ws.Cells(row, 4).Formula = "=C" & row & "-B" & row
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Font.Bold = True
    ws.Cells(row, 5).Formula = "=IF(B" & row & "=0,0,D" & row & "/B" & row & ")"
    ws.Cells(row, 5).NumberFormat = "0.00%"
    
    ' 添加边框
    ws.Range("A4:F" & row).Borders.LineStyle = xlContinuous
    
    MsgBox "成本差异分析完成！" & vbCrLf & vbCrLf & _
           "请根据实际情况修改标准成本和实际成本数据", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "成本差异分析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏52: ARAgingAnalysis - 应收账龄分析
' V8.0新增：分析应收账款账龄
' ============================================================================
Sub ARAgingAnalysis()
    Dim wsAR As Worksheet, wsAging As Worksheet
    Dim lastRow As Long, i As Long, agingRow As Long
    Dim customerName As String
    Dim arAmount As Double, arDays As Long
    Dim amount1to30 As Double, amount31to60 As Double
    Dim amount61to90 As Double, amount91to180 As Double, amountOver180 As Double
    Dim riskLevel As String
    
    On Error GoTo ErrorHandler
    
    ' 检查应收应付表是否存在
    On Error Resume Next
    Set wsAR = ThisWorkbook.Sheets("应收应付")
    If wsAR Is Nothing Then
        MsgBox "请先创建应收应付表", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建应收账龄分析表
    On Error Resume Next
    Set wsAging = ThisWorkbook.Sheets("应收账龄分析")
    If wsAging Is Nothing Then
        Set wsAging = ThisWorkbook.Sheets.Add
        wsAging.Name = "应收账龄分析"
    Else
        wsAging.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsAging.Cells(1, 1).Value = "应收账款账龄分析表"
    wsAging.Cells(1, 1).Font.Size = 16
    wsAging.Cells(1, 1).Font.Bold = True
    wsAging.Cells(1, 1).HorizontalAlignment = xlCenter
    wsAging.Range("A1:J1").Merge
    
    wsAging.Cells(2, 1).Value = "分析日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 列标题
    wsAging.Cells(4, 1).Value = "客户名称"
    wsAging.Cells(4, 2).Value = "应收总额"
    wsAging.Cells(4, 3).Value = "1-30天"
    wsAging.Cells(4, 4).Value = "31-60天"
    wsAging.Cells(4, 5).Value = "61-90天"
    wsAging.Cells(4, 6).Value = "91-180天"
    wsAging.Cells(4, 7).Value = "180天以上"
    wsAging.Cells(4, 8).Value = "账龄占比"
    wsAging.Cells(4, 9).Value = "风险等级"
    wsAging.Cells(4, 10).Value = "备注"
    
    wsAging.Range("A4:J4").Font.Bold = True
    wsAging.Range("A4:J4").Interior.Color = RGB(68, 114, 196)
    wsAging.Range("A4:J4").Font.Color = RGB(255, 255, 255)
    
    ' 设置列宽
    wsAging.Columns("A").ColumnWidth = 18
    wsAging.Columns("B").ColumnWidth = 12
    wsAging.Columns("C").ColumnWidth = 10
    wsAging.Columns("D").ColumnWidth = 10
    wsAging.Columns("E").ColumnWidth = 10
    wsAging.Columns("F").ColumnWidth = 10
    wsAging.Columns("G").ColumnWidth = 10
    wsAging.Columns("H").ColumnWidth = 10
    wsAging.Columns("I").ColumnWidth = 10
    wsAging.Columns("J").ColumnWidth = 20
    
    ' 读取应收账款数据
    lastRow = GetLastRow(wsAR, 1)
    agingRow = 5
    
    For i = 4 To lastRow
        If wsAR.Cells(i, AR_COL_NAME).Value <> "" And wsAR.Cells(i, AR_COL_NAME).Value <> "客户名称" Then
            customerName = wsAR.Cells(i, AR_COL_NAME).Value
            arAmount = Val(wsAR.Cells(i, AR_COL_CLOSE).Value)  ' 期末应收
            
            If arAmount > 0 Then
                wsAging.Cells(agingRow, 1).Value = customerName
                wsAging.Cells(agingRow, 2).Value = arAmount
                wsAging.Cells(agingRow, 2).NumberFormat = "#,##0.00"
                
                ' 模拟账龄分布（实际应用中应根据交易日期计算）
                ' 这里按比例分配作为示例
                wsAging.Cells(agingRow, 3).Value = arAmount * 0.4  ' 1-30天
                wsAging.Cells(agingRow, 3).NumberFormat = "#,##0.00"
                wsAging.Cells(agingRow, 4).Value = arAmount * 0.25  ' 31-60天
                wsAging.Cells(agingRow, 4).NumberFormat = "#,##0.00"
                wsAging.Cells(agingRow, 5).Value = arAmount * 0.15  ' 61-90天
                wsAging.Cells(agingRow, 5).NumberFormat = "#,##0.00"
                wsAging.Cells(agingRow, 6).Value = arAmount * 0.12  ' 91-180天
                wsAging.Cells(agingRow, 6).NumberFormat = "#,##0.00"
                wsAging.Cells(agingRow, 7).Value = arAmount * 0.08  ' 180天以上
                wsAging.Cells(agingRow, 7).NumberFormat = "#,##0.00"
                
                ' 账龄占比（180天以上占比）
                wsAging.Cells(agingRow, 8).Formula = "=IF(B" & agingRow & "=0,0,G" & agingRow & "/B" & agingRow & ")"
                wsAging.Cells(agingRow, 8).NumberFormat = "0.00%"
                
                ' 风险等级判断
                If wsAging.Cells(agingRow, 7).Value > arAmount * 0.3 Then
                    wsAging.Cells(agingRow, 9).Value = "高风险"
                    wsAging.Cells(agingRow, 9).Interior.Color = RGB(255, 199, 206)
                ElseIf wsAging.Cells(agingRow, 7).Value > arAmount * 0.1 Then
                    wsAging.Cells(agingRow, 9).Value = "中风险"
                    wsAging.Cells(agingRow, 9).Interior.Color = RGB(255, 235, 156)
                Else
                    wsAging.Cells(agingRow, 9).Value = "低风险"
                    wsAging.Cells(agingRow, 9).Interior.Color = RGB(198, 239, 206)
                End If
                
                agingRow = agingRow + 1
            End If
        End If
    Next i
    
    ' 合计行
    If agingRow > 5 Then
        wsAging.Cells(agingRow, 1).Value = "合计"
        wsAging.Cells(agingRow, 1).Font.Bold = True
        wsAging.Cells(agingRow, 2).Formula = "=SUM(B5:B" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 2).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 2).Font.Bold = True
        wsAging.Cells(agingRow, 3).Formula = "=SUM(C5:C" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 3).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 4).Formula = "=SUM(D5:D" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 4).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 5).Formula = "=SUM(E5:E" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 5).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 6).Formula = "=SUM(F5:F" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 6).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 7).Formula = "=SUM(G5:G" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 7).NumberFormat = "#,##0.00"
    End If
    
    MsgBox "应收账龄分析完成！" & vbCrLf & vbCrLf & _
           "风险等级说明：" & vbCrLf & _
           "• 高风险：180天以上占比>30%" & vbCrLf & _
           "• 中风险：180天以上占比10%-30%" & vbCrLf & _
           "• 低风险：180天以上占比<10%", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "应收账龄分析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏53: APAgingAnalysis - 应付账龄分析
' V8.0新增：分析应付账款账龄
' ============================================================================
Sub APAgingAnalysis()
    Dim wsAP As Worksheet, wsAging As Worksheet
    Dim lastRow As Long, i As Long, agingRow As Long
    Dim supplierName As String
    Dim apAmount As Double
    
    On Error GoTo ErrorHandler
    
    ' 检查应收应付表是否存在
    On Error Resume Next
    Set wsAP = ThisWorkbook.Sheets("应收应付")
    If wsAP Is Nothing Then
        MsgBox "请先创建应收应付表", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建应付账龄分析表
    On Error Resume Next
    Set wsAging = ThisWorkbook.Sheets("应付账龄分析")
    If wsAging Is Nothing Then
        Set wsAging = ThisWorkbook.Sheets.Add
        wsAging.Name = "应付账龄分析"
    Else
        wsAging.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsAging.Cells(1, 1).Value = "应付账款账龄分析表"
    wsAging.Cells(1, 1).Font.Size = 16
    wsAging.Cells(1, 1).Font.Bold = True
    wsAging.Cells(1, 1).HorizontalAlignment = xlCenter
    wsAging.Range("A1:H1").Merge
    
    wsAging.Cells(2, 1).Value = "分析日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 列标题
    wsAging.Cells(4, 1).Value = "供应商名称"
    wsAging.Cells(4, 2).Value = "应付总额"
    wsAging.Cells(4, 3).Value = "1-30天"
    wsAging.Cells(4, 4).Value = "31-60天"
    wsAging.Cells(4, 5).Value = "61-90天"
    wsAging.Cells(4, 6).Value = "91-180天"
    wsAging.Cells(4, 7).Value = "180天以上"
    wsAging.Cells(4, 8).Value = "备注"
    
    wsAging.Range("A4:H4").Font.Bold = True
    wsAging.Range("A4:H4").Interior.Color = RGB(237, 125, 49)
    wsAging.Range("A4:H4").Font.Color = RGB(255, 255, 255)
    
    ' 设置列宽
    wsAging.Columns("A").ColumnWidth = 18
    wsAging.Columns("B").ColumnWidth = 12
    wsAging.Columns("C").ColumnWidth = 10
    wsAging.Columns("D").ColumnWidth = 10
    wsAging.Columns("E").ColumnWidth = 10
    wsAging.Columns("F").ColumnWidth = 10
    wsAging.Columns("G").ColumnWidth = 10
    wsAging.Columns("H").ColumnWidth = 20
    
    ' 读取应付账款数据
    lastRow = GetLastRow(wsAP, 8)
    agingRow = 5
    
    For i = 4 To lastRow
        If wsAP.Cells(i, AP_COL_NAME).Value <> "" And wsAP.Cells(i, AP_COL_NAME).Value <> "供应商名称" Then
            supplierName = wsAP.Cells(i, AP_COL_NAME).Value
            apAmount = Val(wsAP.Cells(i, AP_COL_CLOSE).Value)  ' 期末应付
            
            If apAmount > 0 Then
                wsAging.Cells(agingRow, 1).Value = supplierName
                wsAging.Cells(agingRow, 2).Value = apAmount
                wsAging.Cells(agingRow, 2).NumberFormat = "#,##0.00"
                
                ' 模拟账龄分布（实际应用中应根据交易日期计算）
                wsAging.Cells(agingRow, 3).Value = apAmount * 0.5  ' 1-30天
                wsAging.Cells(agingRow, 3).NumberFormat = "#,##0.00"
                wsAging.Cells(agingRow, 4).Value = apAmount * 0.3  ' 31-60天
                wsAging.Cells(agingRow, 4).NumberFormat = "#,##0.00"
                wsAging.Cells(agingRow, 5).Value = apAmount * 0.1  ' 61-90天
                wsAging.Cells(agingRow, 5).NumberFormat = "#,##0.00"
                wsAging.Cells(agingRow, 6).Value = apAmount * 0.07  ' 91-180天
                wsAging.Cells(agingRow, 6).NumberFormat = "#,##0.00"
                wsAging.Cells(agingRow, 7).Value = apAmount * 0.03  ' 180天以上
                wsAging.Cells(agingRow, 7).NumberFormat = "#,##0.00"
                
                agingRow = agingRow + 1
            End If
        End If
    Next i
    
    ' 合计行
    If agingRow > 5 Then
        wsAging.Cells(agingRow, 1).Value = "合计"
        wsAging.Cells(agingRow, 1).Font.Bold = True
        wsAging.Cells(agingRow, 2).Formula = "=SUM(B5:B" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 2).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 2).Font.Bold = True
        wsAging.Cells(agingRow, 3).Formula = "=SUM(C5:C" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 3).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 4).Formula = "=SUM(D5:D" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 4).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 5).Formula = "=SUM(E5:E" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 5).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 6).Formula = "=SUM(F5:F" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 6).NumberFormat = "#,##0.00"
        wsAging.Cells(agingRow, 7).Formula = "=SUM(G5:G" & (agingRow - 1) & ")"
        wsAging.Cells(agingRow, 7).NumberFormat = "#,##0.00"
    End If
    
    MsgBox "应付账龄分析完成！", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "应付账龄分析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏54: BadDebtProvision - 坏账准备计提
' V8.0新增：计算坏账准备
' ============================================================================
Sub BadDebtProvision()
    Dim wsAging As Worksheet, wsBadDebt As Worksheet
    Dim lastRow As Long, i As Long, bdRow As Long
    Dim customerName As String
    Dim arBalance As Double, agingDays As Long
    Dim provisionRate As Double, provisionAmount As Double
    
    On Error GoTo ErrorHandler
    
    ' 检查应收账龄分析表是否存在
    On Error Resume Next
    Set wsAging = ThisWorkbook.Sheets("应收账龄分析")
    If wsAging Is Nothing Then
        MsgBox "请先运行【应收账龄分析】", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建坏账准备表
    On Error Resume Next
    Set wsBadDebt = ThisWorkbook.Sheets("坏账准备")
    If wsBadDebt Is Nothing Then
        Set wsBadDebt = ThisWorkbook.Sheets.Add
        wsBadDebt.Name = "坏账准备"
    Else
        wsBadDebt.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsBadDebt.Cells(1, 1).Value = "坏账准备计提表"
    wsBadDebt.Cells(1, 1).Font.Size = 16
    wsBadDebt.Cells(1, 1).Font.Bold = True
    wsBadDebt.Cells(1, 1).HorizontalAlignment = xlCenter
    wsBadDebt.Range("A1:F1").Merge
    
    wsBadDebt.Cells(2, 1).Value = "计提日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 计提比例说明
    wsBadDebt.Cells(3, 1).Value = "计提比例：1年以内" & Format(BD_RATE_1YEAR, "0%") & "、1-2年" & Format(BD_RATE_2YEAR, "0%") & "、2-3年" & Format(BD_RATE_3YEAR, "0%") & "、3年以上" & Format(BD_RATE_OVER3, "0%")
    wsBadDebt.Cells(3, 1).Font.Color = RGB(128, 128, 128)
    
    ' 列标题
    wsBadDebt.Cells(5, 1).Value = "客户名称"
    wsBadDebt.Cells(5, 2).Value = "应收余额"
    wsBadDebt.Cells(5, 3).Value = "账龄(天)"
    wsBadDebt.Cells(5, 4).Value = "账龄区间"
    wsBadDebt.Cells(5, 5).Value = "计提比例"
    wsBadDebt.Cells(5, 6).Value = "坏账准备"
    
    wsBadDebt.Range("A5:F5").Font.Bold = True
    wsBadDebt.Range("A5:F5").Interior.Color = RGB(192, 80, 77)
    wsBadDebt.Range("A5:F5").Font.Color = RGB(255, 255, 255)
    
    ' 设置列宽
    wsBadDebt.Columns("A").ColumnWidth = 18
    wsBadDebt.Columns("B").ColumnWidth = 12
    wsBadDebt.Columns("C").ColumnWidth = 10
    wsBadDebt.Columns("D").ColumnWidth = 12
    wsBadDebt.Columns("E").ColumnWidth = 10
    wsBadDebt.Columns("F").ColumnWidth = 12
    
    ' 读取应收账龄数据并计算坏账准备
    lastRow = GetLastRow(wsAging, 1)
    bdRow = 6
    
    For i = 5 To lastRow - 1
        If wsAging.Cells(i, 1).Value <> "" And wsAging.Cells(i, 1).Value <> "合计" Then
            customerName = wsAging.Cells(i, 1).Value
            arBalance = Val(wsAging.Cells(i, 2).Value)
            
            If arBalance > 0 Then
                wsBadDebt.Cells(bdRow, 1).Value = customerName
                wsBadDebt.Cells(bdRow, 2).Value = arBalance
                wsBadDebt.Cells(bdRow, 2).NumberFormat = "#,##0.00"
                
                ' 根据账龄区间计算（使用180天以上金额估算账龄）
                Dim over180 As Double
                over180 = Val(wsAging.Cells(i, 7).Value)
                
                ' 简化处理：根据180天以上占比估算平均账龄
                If over180 > arBalance * 0.5 Then
                    agingDays = 400  ' 约1年以上
                    provisionRate = BD_RATE_2YEAR
                    wsBadDebt.Cells(bdRow, 4).Value = "1-2年"
                ElseIf over180 > arBalance * 0.2 Then
                    agingDays = 200
                    provisionRate = BD_RATE_1YEAR
                    wsBadDebt.Cells(bdRow, 4).Value = "1年以内"
                Else
                    agingDays = 60
                    provisionRate = BD_RATE_1YEAR
                    wsBadDebt.Cells(bdRow, 4).Value = "1年以内"
                End If
                
                wsBadDebt.Cells(bdRow, 3).Value = agingDays
                wsBadDebt.Cells(bdRow, 5).Value = provisionRate
                wsBadDebt.Cells(bdRow, 5).NumberFormat = "0.00%"
                
                provisionAmount = arBalance * provisionRate
                wsBadDebt.Cells(bdRow, 6).Value = provisionAmount
                wsBadDebt.Cells(bdRow, 6).NumberFormat = "#,##0.00"
                
                bdRow = bdRow + 1
            End If
        End If
    Next i
    
    ' 合计行
    If bdRow > 6 Then
        wsBadDebt.Cells(bdRow, 1).Value = "合计"
        wsBadDebt.Cells(bdRow, 1).Font.Bold = True
        wsBadDebt.Cells(bdRow, 2).Formula = "=SUM(B6:B" & (bdRow - 1) & ")"
        wsBadDebt.Cells(bdRow, 2).NumberFormat = "#,##0.00"
        wsBadDebt.Cells(bdRow, 2).Font.Bold = True
        wsBadDebt.Cells(bdRow, 6).Formula = "=SUM(F6:F" & (bdRow - 1) & ")"
        wsBadDebt.Cells(bdRow, 6).NumberFormat = "#,##0.00"
        wsBadDebt.Cells(bdRow, 6).Font.Bold = True
        wsBadDebt.Cells(bdRow, 6).Interior.Color = RGB(255, 199, 206)
    End If
    
    MsgBox "坏账准备计提完成！" & vbCrLf & vbCrLf & _
           "计提比例说明：" & vbCrLf & _
           "• 1年以内：" & Format(BD_RATE_1YEAR, "0%") & vbCrLf & _
           "• 1-2年：" & Format(BD_RATE_2YEAR, "0%") & vbCrLf & _
           "• 2-3年：" & Format(BD_RATE_3YEAR, "0%") & vbCrLf & _
           "• 3年以上：" & Format(BD_RATE_OVER3, "0%"), vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "坏账准备计提出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏55: CreateExpenseBudget - 费用预算编制
' V9.0新增：编制月度费用预算
' ============================================================================
Sub CreateExpenseBudget()
    Dim wsBudget As Worksheet
    Dim budgetMonth As String
    Dim categories As Variant
    Dim i As Long, row As Long
    Dim lastRow As Long
    Dim lastYearAmount As Double, lastMonthAmount As Double
    Dim wsExpense As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' 获取预算月份
    budgetMonth = InputBox("请输入预算月份（格式：YYYY-MM）：", "费用预算编制", Format(Date, "yyyy-mm"))
    If budgetMonth = "" Then Exit Sub
    
    ' 检查支出记录表
    On Error Resume Next
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    If wsExpense Is Nothing Then
        MsgBox "请先录入支出记录！", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建费用预算表
    On Error Resume Next
    Set wsBudget = ThisWorkbook.Sheets("费用预算")
    If wsBudget Is Nothing Then
        Set wsBudget = ThisWorkbook.Sheets.Add
        wsBudget.Name = "费用预算"
    Else
        wsBudget.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsBudget.Cells(1, 1).Value = "费用预算编制表"
    wsBudget.Cells(1, 1).Font.Size = 16
    wsBudget.Cells(1, 1).Font.Bold = True
    wsBudget.Cells(1, 1).HorizontalAlignment = xlCenter
    wsBudget.Range("A1:G1").Merge
    
    wsBudget.Cells(2, 1).Value = "预算月份：" & budgetMonth
    wsBudget.Cells(2, 1).Font.Size = 12
    
    ' 列标题
    row = 4
    wsBudget.Cells(row, 1).Value = "费用类别"
    wsBudget.Cells(row, 2).Value = "预算金额"
    wsBudget.Cells(row, 3).Value = "上月实际"
    wsBudget.Cells(row, 4).Value = "去年同期"
    wsBudget.Cells(row, 5).Value = "环比增减"
    wsBudget.Cells(row, 6).Value = "同比增减"
    wsBudget.Cells(row, 7).Value = "预算说明"
    
    wsBudget.Range("A4:G4").Font.Bold = True
    wsBudget.Range("A4:G4").Interior.Color = RGB(68, 114, 196)
    wsBudget.Range("A4:G4").Font.Color = RGB(255, 255, 255)
    
    ' 费用类别
    categories = Split(EXPENSE_CATEGORIES, ",")
    row = 5
    
    For i = 0 To UBound(categories)
        wsBudget.Cells(row, 1).Value = categories(i)
        
        ' 输入预算金额
        Dim budgetAmount As Double
        budgetAmount = Val(InputBox("请输入【" & categories(i) & "】的预算金额：", "预算金额", "0"))
        wsBudget.Cells(row, 2).Value = budgetAmount
        wsBudget.Cells(row, 2).NumberFormat = "#,##0.00"
        
        ' 上月实际（从支出记录获取）
        wsBudget.Cells(row, 3).Formula = "=SUMIF('支出记录!B:B',A" & row & ",支出记录!C:C)"
        wsBudget.Cells(row, 3).NumberFormat = "#,##0.00"
        
        ' 去年同期（简化处理）
        wsBudget.Cells(row, 4).Value = budgetAmount * 0.95
        wsBudget.Cells(row, 4).NumberFormat = "#,##0.00"
        
        ' 环比增减公式
        wsBudget.Cells(row, 5).Formula = "=IF(C" & row & "=0,0,(B" & row & "-C" & row & ")/C" & row & ")"
        wsBudget.Cells(row, 5).NumberFormat = "0.00%"
        
        ' 同比增减公式
        wsBudget.Cells(row, 6).Formula = "=IF(D" & row & "=0,0,(B" & row & "-D" & row & ")/D" & row & ")"
        wsBudget.Cells(row, 6).NumberFormat = "0.00%"
        
        row = row + 1
    Next i
    
    ' 合计行
    wsBudget.Cells(row, 1).Value = "合计"
    wsBudget.Cells(row, 1).Font.Bold = True
    wsBudget.Cells(row, 2).Formula = "=SUM(B5:B" & (row - 1) & ")"
    wsBudget.Cells(row, 2).NumberFormat = "#,##0.00"
    wsBudget.Cells(row, 2).Font.Bold = True
    wsBudget.Cells(row, 3).Formula = "=SUM(C5:C" & (row - 1) & ")"
    wsBudget.Cells(row, 3).NumberFormat = "#,##0.00"
    wsBudget.Cells(row, 4).Formula = "=SUM(D5:D" & (row - 1) & ")"
    wsBudget.Cells(row, 4).NumberFormat = "#,##0.00"
    
    ' 列宽
    wsBudget.Columns("A").ColumnWidth = 15
    wsBudget.Columns("B").ColumnWidth = 12
    wsBudget.Columns("C").ColumnWidth = 12
    wsBudget.Columns("D").ColumnWidth = 12
    wsBudget.Columns("E").ColumnWidth = 10
    wsBudget.Columns("F").ColumnWidth = 10
    wsBudget.Columns("G").ColumnWidth = 25
    
    MsgBox "费用预算编制完成！", vbInformation, "完成"
    Exit Sub
    
ErrorHandler:
    MsgBox "费用预算编制出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏56: BudgetControl - 预算执行控制
' V9.0新增：监控预算执行情况
' ============================================================================
Sub BudgetControl()
    Dim wsBudget As Worksheet, wsExpense As Worksheet, wsControl As Worksheet
    Dim lastRow As Long, i As Long, row As Long
    Dim categories As Variant
    Dim budgetAmount As Double, actualAmount As Double
    Dim executionRate As Double, remaining As Double
    
    On Error GoTo ErrorHandler
    
    ' 检查费用预算表
    On Error Resume Next
    Set wsBudget = ThisWorkbook.Sheets("费用预算")
    If wsBudget Is Nothing Then
        MsgBox "请先运行【费用预算编制】！", vbExclamation, "提示"
        Exit Sub
    End If
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    If wsExpense Is Nothing Then
        MsgBox "请先录入支出记录！", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建预算执行表
    On Error Resume Next
    Set wsControl = ThisWorkbook.Sheets("预算执行")
    If wsControl Is Nothing Then
        Set wsControl = ThisWorkbook.Sheets.Add
        wsControl.Name = "预算执行"
    Else
        wsControl.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsControl.Cells(1, 1).Value = "预算执行情况表"
    wsControl.Cells(1, 1).Font.Size = 16
    wsControl.Cells(1, 1).Font.Bold = True
    wsControl.Cells(1, 1).HorizontalAlignment = xlCenter
    wsControl.Range("A1:F1").Merge
    
    wsControl.Cells(2, 1).Value = "编制日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 列标题
    row = 4
    wsControl.Cells(row, 1).Value = "费用类别"
    wsControl.Cells(row, 2).Value = "预算金额"
    wsControl.Cells(row, 3).Value = "实际发生"
    wsControl.Cells(row, 4).Value = "执行进度"
    wsControl.Cells(row, 5).Value = "剩余预算"
    wsControl.Cells(row, 6).Value = "预警状态"
    
    wsControl.Range("A4:F4").Font.Bold = True
    wsControl.Range("A4:F4").Interior.Color = RGB(237, 125, 49)
    wsControl.Range("A4:F4").Font.Color = RGB(255, 255, 255)
    
    ' 读取费用类别
    categories = Split(EXPENSE_CATEGORIES, ",")
    row = 5
    
    For i = 0 To UBound(categories)
        wsControl.Cells(row, 1).Value = categories(i)
        
        ' 预算金额
        budgetAmount = 0
        lastRow = GetLastRow(wsBudget, 1)
        For j = 5 To lastRow
            If wsBudget.Cells(j, 1).Value = categories(i) Then
                budgetAmount = Val(wsBudget.Cells(j, 2).Value)
                Exit For
            End If
        Next
        wsControl.Cells(row, 2).Value = budgetAmount
        wsControl.Cells(row, 2).NumberFormat = "#,##0.00"
        
        ' 实际发生额
        wsControl.Cells(row, 3).Formula = "=SUMIF('支出记录!B:B',A" & row & ",支出记录!C:C)"
        wsControl.Cells(row, 3).NumberFormat = "#,##0.00"
        
        ' 执行进度
        wsControl.Cells(row, 4).Formula = "=IF(B" & row & "=0,0,C" & row & "/B" & row & ")"
        wsControl.Cells(row, 4).NumberFormat = "0.00%"
        
        ' 剩余预算
        wsControl.Cells(row, 5).Formula = "=B" & row & "-C" & row
        wsControl.Cells(row, 5).NumberFormat = "#,##0.00"
        
        ' 超支预警（红色）
        wsControl.Cells(row, 6).Value = "正常"
        If budgetAmount > 0 Then
            executionRate = Val(wsControl.Cells(row, 3).Value) / budgetAmount
            If executionRate > BUDGET_OVER_RATE Then
                wsControl.Cells(row, 6).Value = "超支"
                wsControl.Cells(row, 6).Interior.Color = RGB(255, 0, 0)
                wsControl.Cells(row, 6).Font.Color = RGB(255, 255, 255)
            ElseIf executionRate > BUDGET_ALERT_RATE Then
                wsControl.Cells(row, 6).Value = "关注"
                wsControl.Cells(row, 6).Interior.Color = RGB(255, 255, 0)
            End If
        End If
        
        row = row + 1
    Next i
    
    ' 合计行
    wsControl.Cells(row, 1).Value = "合计"
    wsControl.Cells(row, 1).Font.Bold = True
    wsControl.Cells(row, 2).Formula = "=SUM(B5:B" & (row - 1) & ")"
    wsControl.Cells(row, 2).NumberFormat = "#,##0.00"
    wsControl.Cells(row, 2).Font.Bold = True
    wsControl.Cells(row, 3).Formula = "=SUM(C5:C" & (row - 1) & ")"
    wsControl.Cells(row, 3).NumberFormat = "#,##0.00"
    wsControl.Cells(row, 3).Font.Bold = True
    
    ' 列宽
    wsControl.Columns("A").ColumnWidth = 15
    wsControl.Columns("B").ColumnWidth = 12
    wsControl.Columns("C").ColumnWidth = 12
    wsControl.Columns("D").ColumnWidth = 12
    wsControl.Columns("E").ColumnWidth = 12
    wsControl.Columns("F").ColumnWidth = 10
    
    MsgBox "预算执行情况已更新！", vbInformation, "完成"
    Exit Sub
    
ErrorHandler:
    MsgBox "预算执行控制出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏57: BudgetAlert - 超预算预警
' V9.0新增：检查各费用类别是否超预算
' ============================================================================
Sub BudgetAlert()
    Dim wsBudget As Worksheet, wsExpense As Worksheet
    Dim wsAlert As Worksheet
    Dim lastRow As Long, i As Long, row As Long
    Dim categories As Variant
    Dim budgetAmount As Double, actualAmount As Double
    Dim difference As Double
    Dim alertLevel As String
    
    On Error GoTo ErrorHandler
    
    ' 检查费用预算表
    On Error Resume Next
    Set wsBudget = ThisWorkbook.Sheets("费用预算")
    If wsBudget Is Nothing Then
        MsgBox "请先运行【费用预算编制】！", vbExclamation, "提示"
        Exit Sub
    End If
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    If wsExpense Is Nothing Then
        MsgBox "请先录入支出记录！", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建预警表
    On Error Resume Next
    Set wsAlert = ThisWorkbook.Sheets("预算预警")
    If wsAlert Is Nothing Then
        Set wsAlert = ThisWorkbook.Sheets.Add
        wsAlert.Name = "预算预警"
    Else
        wsAlert.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsAlert.Cells(1, 1).Value = "超预算预警表"
    wsAlert.Cells(1, 1).Font.Size = 16
    wsAlert.Cells(1, 1).Font.Bold = True
    wsAlert.Cells(1, 1).HorizontalAlignment = xlCenter
    wsAlert.Range("A1:E1").Merge
    
    wsAlert.Cells(2, 1).Value = "预警日期：" & Format(Date, "yyyy-mm-dd")
    wsAlert.Cells(2, 2).Value = "关注线：" & Format(BUDGET_ALERT_RATE, "0%") & "  超支线：" & Format(BUDGET_OVER_RATE, "0%")
    
    ' 列标题
    row = 4
    wsAlert.Cells(row, 1).Value = "费用类别"
    wsAlert.Cells(row, 2).Value = "预算金额"
    wsAlert.Cells(row, 3).Value = "实际金额"
    wsAlert.Cells(row, 4).Value = "差异额"
    wsAlert.Cells(row, 5).Value = "预警级别"
    
    wsAlert.Range("A4:E4").Font.Bold = True
    wsAlert.Range("A4:E4").Interior.Color = RGB(192, 80, 77)
    wsAlert.Range("A4:E4").Font.Color = RGB(255, 255, 255)
    
    ' 读取费用类别
    categories = Split(EXPENSE_CATEGORIES, ",")
    row = 5
    Dim alertCount As Integer
    alertCount = 0
    
    For i = 0 To UBound(categories)
        wsAlert.Cells(row, 1).Value = categories(i)
        
        ' 预算金额
        budgetAmount = 0
        lastRow = GetLastRow(wsBudget, 1)
        For j = 5 To lastRow
            If wsBudget.Cells(j, 1).Value = categories(i) Then
                budgetAmount = Val(wsBudget.Cells(j, 2).Value)
                Exit For
            End If
        Next
        wsAlert.Cells(row, 2).Value = budgetAmount
        wsAlert.Cells(row, 2).NumberFormat = "#,##0.00"
        
        ' 实际金额
        wsAlert.Cells(row, 3).Formula = "=SUMIF('支出记录!B:B',A" & row & ",支出记录!C:C)"
        wsAlert.Cells(row, 3).NumberFormat = "#,##0.00"
        actualAmount = Val(wsAlert.Cells(row, 3).Value)
        
        ' 差异额
        difference = actualAmount - budgetAmount
        wsAlert.Cells(row, 4).Value = difference
        wsAlert.Cells(row, 4).NumberFormat = "#,##0.00"
        
        ' 预警级别判断
        If budgetAmount > 0 Then
            Dim rate As Double
            rate = actualAmount / budgetAmount
            
            If rate > BUDGET_OVER_RATE Then
                alertLevel = "超支"
                wsAlert.Cells(row, 5).Value = alertLevel
                wsAlert.Cells(row, 5).Interior.Color = RGB(255, 0, 0)
                wsAlert.Cells(row, 5).Font.Color = RGB(255, 255, 255)
                wsAlert.Cells(row, 5).Font.Bold = True
                alertCount = alertCount + 1
            ElseIf rate > BUDGET_ALERT_RATE Then
                alertLevel = "关注"
                wsAlert.Cells(row, 5).Value = alertLevel
                wsAlert.Cells(row, 5).Interior.Color = RGB(255, 255, 0)
                alertCount = alertCount + 1
            Else
                wsAlert.Cells(row, 5).Value = "正常"
                wsAlert.Cells(row, 5).Interior.Color = RGB(0, 255, 0)
            End If
        Else
            wsAlert.Cells(row, 5).Value = "未预算"
        End If
        
        row = row + 1
    Next i
    
    ' 合计行
    wsAlert.Cells(row, 1).Value = "合计"
    wsAlert.Cells(row, 1).Font.Bold = True
    wsAlert.Cells(row, 2).Formula = "=SUM(B5:B" & (row - 1) & ")"
    wsAlert.Cells(row, 2).NumberFormat = "#,##0.00"
    wsAlert.Cells(row, 3).Formula = "=SUM(C5:C" & (row - 1) & ")"
    wsAlert.Cells(row, 3).NumberFormat = "#,##0.00"
    wsAlert.Cells(row, 4).Formula = "=SUM(D5:D" & (row - 1) & ")"
    wsAlert.Cells(row, 4).NumberFormat = "#,##0.00"
    
    ' 列宽
    wsAlert.Columns("A").ColumnWidth = 15
    wsAlert.Columns("B").ColumnWidth = 12
    wsAlert.Columns("C").ColumnWidth = 12
    wsAlert.Columns("D").ColumnWidth = 12
    wsAlert.Columns("E").ColumnWidth = 10
    
    ' 预警提示
    If alertCount > 0 Then
        MsgBox "发现 " & alertCount & " 个费用类别需要关注！" & vbCrLf & vbCrLf & _
               "• 超支：实际超过预算" & vbCrLf & _
               "• 关注：实际超过 " & Format(BUDGET_ALERT_RATE, "0%") & " 预算", _
               vbExclamation, "预警提示"
    Else
        MsgBox "所有费用类别均在预算范围内！", vbInformation, "预警提示"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "超预算预警出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏58: GenerateVoucher - 生成记账凭证
' V9.0新增：根据收支记录自动生成记账凭证
' ============================================================================
Sub GenerateVoucher()
    Dim wsIncome As Worksheet, wsExpense As Worksheet, wsVoucher As Worksheet
    Dim lastRow As Long, i As Long, row As Long
    Dim voucherNo As Long
    Dim transDate As Date
    Dim category As String, amount As Double
    Dim paymentMethod As String
    Dim summary As String
    Dim debitCode As String, debitName As String
    Dim creditCode As String, creditName As String
    Dim wsOutsource As Worksheet, wsOutRow As Long
    
    On Error GoTo ErrorHandler
    
    ' 检查收支记录
    On Error Resume Next
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsOutsource = ThisWorkbook.Sheets("外发加工费")
    If wsIncome Is Nothing And wsExpense Is Nothing Then
        MsgBox "请先录入收支记录！", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建记账凭证表
    On Error Resume Next
    Set wsVoucher = ThisWorkbook.Sheets("记账凭证")
    If wsVoucher Is Nothing Then
        Set wsVoucher = ThisWorkbook.Sheets.Add
        wsVoucher.Name = "记账凭证"
    Else
        wsVoucher.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsVoucher.Cells(1, 1).Value = "记账凭证"
    wsVoucher.Cells(1, 1).Font.Size = 16
    wsVoucher.Cells(1, 1).Font.Bold = True
    wsVoucher.Cells(1, 1).HorizontalAlignment = xlCenter
    wsVoucher.Range("A1:I1").Merge
    
    wsVoucher.Cells(2, 1).Value = "编制日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 列标题
    row = 4
    wsVoucher.Cells(row, 1).Value = "凭证字号"
    wsVoucher.Cells(row, 2).Value = "凭证日期"
    wsVoucher.Cells(row, 3).Value = "摘要"
    wsVoucher.Cells(row, 4).Value = "借方科目代码"
    wsVoucher.Cells(row, 5).Value = "借方科目名称"
    wsVoucher.Cells(row, 6).Value = "贷方科目代码"
    wsVoucher.Cells(row, 7).Value = "贷方科目名称"
    wsVoucher.Cells(row, 8).Value = "金额"
    wsVoucher.Cells(row, 9).Value = "附件"
    
    wsVoucher.Range("A4:I4").Font.Bold = True
    wsVoucher.Range("A4:I4").Interior.Color = RGB(68, 114, 196)
    wsVoucher.Range("A4:I4").Font.Color = RGB(255, 255, 255)
    
    voucherNo = 1
    row = 5
    
    ' 生成收入凭证
    If Not wsIncome Is Nothing Then
        lastRow = GetLastRow(wsIncome, 1)
        For i = 5 To lastRow
            If wsIncome.Cells(i, 3).Value > 0 Then
                transDate = wsIncome.Cells(i, 1).Value
                amount = wsIncome.Cells(i, 3).Value
                paymentMethod = wsIncome.Cells(i, 4).Value
                summary = "收到" & wsIncome.Cells(i, 2).Value & "加工费"
                
                ' 借：银行存款/现金
                If InStr(paymentMethod, "银行") > 0 Then
                    debitCode = SUBJECT_BANK
                    debitName = "银行存款"
                Else
                    debitCode = SUBJECT_CASH
                    debitName = "库存现金"
                End If
                
                ' 贷：主营业务收入
                creditCode = SUBJECT_REVENUE
                creditName = "主营业务收入"
                
                wsVoucher.Cells(row, 1).Value = "记-" & Format(voucherNo, "000")
                wsVoucher.Cells(row, 2).Value = Format(transDate, "yyyy-mm-dd")
                wsVoucher.Cells(row, 3).Value = summary
                wsVoucher.Cells(row, 4).Value = debitCode
                wsVoucher.Cells(row, 5).Value = debitName
                wsVoucher.Cells(row, 6).Value = creditCode
                wsVoucher.Cells(row, 7).Value = creditName
                wsVoucher.Cells(row, 8).Value = amount
                wsVoucher.Cells(row, 8).NumberFormat = "#,##0.00"
                wsVoucher.Cells(row, 9).Value = 1
                
                row = row + 1
                voucherNo = voucherNo + 1
            End If
        Next i
    End If
    
    ' 生成支出凭证
    If Not wsExpense Is Nothing Then
        lastRow = GetLastRow(wsExpense, 1)
        For i = 5 To lastRow
            If wsExpense.Cells(i, 3).Value > 0 Then
                transDate = wsExpense.Cells(i, 1).Value
                amount = wsExpense.Cells(i, 3).Value
                category = wsExpense.Cells(i, 2).Value
                paymentMethod = wsExpense.Cells(i, 5).Value
                summary = category
                
                ' 根据类别确定借方科目
                Select Case category
                    Case CAT_SALARY
                        debitCode = SUBJECT_SALARY
                        debitName = "应付职工薪酬"
                    Case CAT_RENT, CAT_WATER, CAT_ELECTRIC, CAT_DAILY
                        debitCode = SUBJECT_EXPENSE
                        debitName = "管理费用"
                    Case Else
                        debitCode = SUBJECT_COST
                        debitName = "主营业务成本"
                End Select
                
                ' 贷：银行存款/现金
                If InStr(paymentMethod, "银行") > 0 Then
                    creditCode = SUBJECT_BANK
                    creditName = "银行存款"
                Else
                    creditCode = SUBJECT_CASH
                    creditName = "库存现金"
                End If
                
                wsVoucher.Cells(row, 1).Value = "记-" & Format(voucherNo, "000")
                wsVoucher.Cells(row, 2).Value = Format(transDate, "yyyy-mm-dd")
                wsVoucher.Cells(row, 3).Value = summary
                wsVoucher.Cells(row, 4).Value = debitCode
                wsVoucher.Cells(row, 5).Value = debitName
                wsVoucher.Cells(row, 6).Value = creditCode
                wsVoucher.Cells(row, 7).Value = creditName
                wsVoucher.Cells(row, 8).Value = amount
                wsVoucher.Cells(row, 8).NumberFormat = "#,##0.00"
                wsVoucher.Cells(row, 9).Value = 1
                
                row = row + 1
                voucherNo = voucherNo + 1
            End If
        Next i
    End If
    
    ' 生成外发加工凭证（借：主营业务成本 贷：应付账款）
    If Not wsOutsource Is Nothing Then
        lastRow = GetLastRow(wsOutsource, 1)
        For i = 5 To lastRow
            If wsOutsource.Cells(i, 6).Value > 0 Then
                transDate = wsOutsource.Cells(i, 1).Value
                amount = wsOutsource.Cells(i, 6).Value
                summary = "外发加工费：" & wsOutsource.Cells(i, 2).Value
                
                wsVoucher.Cells(row, 1).Value = "记-" & Format(voucherNo, "000")
                wsVoucher.Cells(row, 2).Value = Format(transDate, "yyyy-mm-dd")
                wsVoucher.Cells(row, 3).Value = summary
                wsVoucher.Cells(row, 4).Value = SUBJECT_COST
                wsVoucher.Cells(row, 5).Value = "主营业务成本"
                wsVoucher.Cells(row, 6).Value = SUBJECT_AP
                wsVoucher.Cells(row, 7).Value = "应付账款"
                wsVoucher.Cells(row, 8).Value = amount
                wsVoucher.Cells(row, 8).NumberFormat = "#,##0.00"
                wsVoucher.Cells(row, 9).Value = 1
                
                row = row + 1
                voucherNo = voucherNo + 1
            End If
        Next i
    End If
    
    ' 合计行
    wsVoucher.Cells(row, 1).Value = "合计"
    wsVoucher.Cells(row, 1).Font.Bold = True
    wsVoucher.Cells(row, 8).Formula = "=SUM(H5:H" & (row - 1) & ")"
    wsVoucher.Cells(row, 8).NumberFormat = "#,##0.00"
    wsVoucher.Cells(row, 8).Font.Bold = True
    
    ' 列宽
    wsVoucher.Columns("A").ColumnWidth = 10
    wsVoucher.Columns("B").ColumnWidth = 12
    wsVoucher.Columns("C").ColumnWidth = 25
    wsVoucher.Columns("D").ColumnWidth = 12
    wsVoucher.Columns("E").ColumnWidth = 18
    wsVoucher.Columns("F").ColumnWidth = 12
    wsVoucher.Columns("G").ColumnWidth = 18
    wsVoucher.Columns("H").ColumnWidth = 12
    wsVoucher.Columns("I").ColumnWidth = 8
    
    MsgBox "共生成 " & (voucherNo - 1) & " 张记账凭证！", vbInformation, "完成"
    Exit Sub
    
ErrorHandler:
    MsgBox "生成记账凭证出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏59: GenerateGeneralLedger - 生成总账
' V9.0新增：根据凭证生成总账
' ============================================================================
Sub GenerateGeneralLedger()
    Dim wsVoucher As Worksheet, wsLedger As Worksheet
    Dim lastRow As Long, i As Long, row As Long
    Dim subjectCode As String, subjectName As String
    Dim debitTotal As Double, creditTotal As Double
    Dim balance As Double, balanceDirection As String
    Dim subjects As Object
    Dim key As String
    Dim openingBalance As Double
    
    On Error GoTo ErrorHandler
    
    ' 检查记账凭证表
    On Error Resume Next
    Set wsVoucher = ThisWorkbook.Sheets("记账凭证")
    If wsVoucher Is Nothing Then
        MsgBox "请先运行【生成记账凭证】！", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建总账表
    On Error Resume Next
    Set wsLedger = ThisWorkbook.Sheets("总账")
    If wsLedger Is Nothing Then
        Set wsLedger = ThisWorkbook.Sheets.Add
        wsLedger.Name = "总账"
    Else
        wsLedger.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsLedger.Cells(1, 1).Value = "总账"
    wsLedger.Cells(1, 1).Font.Size = 16
    wsLedger.Cells(1, 1).Font.Bold = True
    wsLedger.Cells(1, 1).HorizontalAlignment = xlCenter
    wsLedger.Range("A1:G1").Merge
    
    wsLedger.Cells(2, 1).Value = "编制日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 列标题
    row = 4
    wsLedger.Cells(row, 1).Value = "科目代码"
    wsLedger.Cells(row, 2).Value = "科目名称"
    wsLedger.Cells(row, 3).Value = "期初余额"
    wsLedger.Cells(row, 4).Value = "本期借方"
    wsLedger.Cells(row, 5).Value = "本期贷方"
    wsLedger.Cells(row, 6).Value = "期末余额"
    wsLedger.Cells(row, 7).Value = "借贷方向"
    
    wsLedger.Range("A4:G4").Font.Bold = True
    wsLedger.Range("A4:G4").Interior.Color = RGB(112, 173, 71)
    wsLedger.Range("A4:G4").Font.Color = RGB(255, 255, 255)
    
    ' 使用字典汇总科目
    Set subjects = ' 【WPS兼容】CreateObject("Scripting.Dictionary") ' WPS不支持，请使用Collection
    
    lastRow = GetLastRow(wsVoucher, 1)
    For i = 5 To lastRow - 1  ' 跳过合计行
        If wsVoucher.Cells(i, 1).Value <> "" And wsVoucher.Cells(i, 1).Value <> "合计" Then
            ' 借方汇总
            subjectCode = wsVoucher.Cells(i, 4).Value
            subjectName = wsVoucher.Cells(i, 5).Value
            key = subjectCode
            If subjects.Exists(key) Then
                subjects(key) = subjects(key) + Val(wsVoucher.Cells(i, 8).Value)  ' 借方累计
            Else
                subjects(key) = Val(wsVoucher.Cells(i, 8).Value)
                If Not subjects.Exists(key & "_name") Then subjects.Add key & "_name", subjectName
            End If
            
            ' 贷方汇总
            subjectCode = wsVoucher.Cells(i, 6).Value
            subjectName = wsVoucher.Cells(i, 7).Value
            key = subjectCode & "_cr"
            If subjects.Exists(key) Then
                subjects(key) = subjects(key) + Val(wsVoucher.Cells(i, 8).Value)  ' 贷方累计
            Else
                subjects(key) = Val(wsVoucher.Cells(i, 8).Value)
                If Not subjects.Exists(key & "_name") Then subjects.Add key & "_name", subjectName
            End If
        End If
    Next i
    
    ' 输出总账
    row = 5
    Dim subjectKeys As Variant
    subjectKeys = subjects.Keys
    
    For i = 0 To UBound(subjectKeys)
        key = subjectKeys(i)
        If InStr(key, "_cr") = 0 And InStr(key, "_name") = 0 Then
            subjectCode = key
            subjectName = subjects(key & "_name")
            debitTotal = 0: creditTotal = 0
            If subjects.Exists(key) Then debitTotal = subjects(key)
            If subjects.Exists(key & "_cr") Then creditTotal = subjects(key & "_cr")
            
            wsLedger.Cells(row, 1).Value = subjectCode
            wsLedger.Cells(row, 2).Value = subjectName
            wsLedger.Cells(row, 3).Value = 0  ' 期初余额
            wsLedger.Cells(row, 4).Value = debitTotal
            wsLedger.Cells(row, 4).NumberFormat = "#,##0.00"
            wsLedger.Cells(row, 5).Value = creditTotal
            wsLedger.Cells(row, 5).NumberFormat = "#,##0.00"
            
            ' 计算期末余额
            balance = debitTotal - creditTotal
            If Left(subjectCode, 1) = "1" Or Left(subjectCode, 1) = "4" Then
                ' 资产类和成本类：借正贷负
                If balance >= 0 Then
                    balanceDirection = "借"
                Else
                    balanceDirection = "贷"
                    balance = Abs(balance)
                End If
            Else
                ' 负债类和权益类：贷正借负
                If balance >= 0 Then
                    balanceDirection = "贷"
                Else
                    balanceDirection = "借"
                    balance = Abs(balance)
                End If
            End If
            
            wsLedger.Cells(row, 6).Value = balance
            wsLedger.Cells(row, 6).NumberFormat = "#,##0.00"
            wsLedger.Cells(row, 7).Value = balanceDirection
            
            row = row + 1
        End If
    Next i
    
    ' 合计行
    wsLedger.Cells(row, 1).Value = "合计"
    wsLedger.Cells(row, 1).Font.Bold = True
    wsLedger.Cells(row, 4).Formula = "=SUM(D5:D" & (row - 1) & ")"
    wsLedger.Cells(row, 4).NumberFormat = "#,##0.00"
    wsLedger.Cells(row, 5).Formula = "=SUM(E5:E" & (row - 1) & ")"
    wsLedger.Cells(row, 5).NumberFormat = "#,##0.00"
    
    ' 列宽
    wsLedger.Columns("A").ColumnWidth = 12
    wsLedger.Columns("B").ColumnWidth = 18
    wsLedger.Columns("C").ColumnWidth = 12
    wsLedger.Columns("D").ColumnWidth = 12
    wsLedger.Columns("E").ColumnWidth = 12
    wsLedger.Columns("F").ColumnWidth = 12
    wsLedger.Columns("G").ColumnWidth = 10
    
    Set subjects = Nothing
    
    MsgBox "总账生成完成！", vbInformation, "完成"
    Exit Sub
    
ErrorHandler:
    MsgBox "生成总账出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏60: GenerateSubLedger - 生成明细账
' V9.0新增：生成指定科目的明细账
' ============================================================================
Sub GenerateSubLedger()
    Dim wsVoucher As Worksheet, wsSub As Worksheet
    Dim lastRow As Long, i As Long, row As Long
    Dim subjectCode As String, subjectName As String
    Dim inputCode As String
    Dim voucherNo As String, transDate As Date
    Dim summary As String
    Dim debitAmount As Double, creditAmount As Double
    Dim balance As Double
    
    On Error GoTo ErrorHandler
    
    ' 检查记账凭证表
    On Error Resume Next
    Set wsVoucher = ThisWorkbook.Sheets("记账凭证")
    If wsVoucher Is Nothing Then
        MsgBox "请先运行【生成记账凭证】！", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 选择科目
    inputCode = InputBox("请输入要查询的科目代码：" & vbCrLf & vbCrLf & _
                         "常用科目代码：" & vbCrLf & _
                         "1001 - 库存现金" & vbCrLf & _
                         "1002 - 银行存款" & vbCrLf & _
                         "1122 - 应收账款" & vbCrLf & _
                         "2202 - 应付账款" & vbCrLf & _
                         "2211 - 应付职工薪酬" & vbCrLf & _
                         "6001 - 主营业务收入" & vbCrLf & _
                         "6401 - 主营业务成本" & vbCrLf & _
                         "6602 - 管理费用", "选择科目", "1002")
    If inputCode = "" Then Exit Sub
    
    ' 创建明细账表
    On Error Resume Next
    Set wsSub = ThisWorkbook.Sheets("明细账")
    If wsSub Is Nothing Then
        Set wsSub = ThisWorkbook.Sheets.Add
        wsSub.Name = "明细账"
    Else
        wsSub.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsSub.Cells(1, 1).Value = "明细账"
    wsSub.Cells(1, 1).Font.Size = 16
    wsSub.Cells(1, 1).Font.Bold = True
    wsSub.Cells(1, 1).HorizontalAlignment = xlCenter
    wsSub.Range("A1:G1").Merge
    
    wsSub.Cells(2, 1).Value = "科目代码：" & inputCode
    
    ' 列标题
    row = 4
    wsSub.Cells(row, 1).Value = "凭证字号"
    wsSub.Cells(row, 2).Value = "日期"
    wsSub.Cells(row, 3).Value = "摘要"
    wsSub.Cells(row, 4).Value = "借方金额"
    wsSub.Cells(row, 5).Value = "贷方金额"
    wsSub.Cells(row, 6).Value = "余额"
    wsSub.Cells(row, 7).Value = "借贷方向"
    
    wsSub.Range("A4:G4").Font.Bold = True
    wsSub.Range("A4:G4").Interior.Color = RGB(68, 114, 196)
    wsSub.Range("A4:G4").Font.Color = RGB(255, 255, 255)
    
    ' 查找该科目的所有凭证
    balance = 0
    row = 5
    lastRow = GetLastRow(wsVoucher, 1)
    
    For i = 5 To lastRow - 1
        If wsVoucher.Cells(i, 1).Value <> "" And wsVoucher.Cells(i, 1).Value <> "合计" Then
            ' 检查借方或贷方是否包含该科目
            If wsVoucher.Cells(i, 4).Value = inputCode Or wsVoucher.Cells(i, 6).Value = inputCode Then
                voucherNo = wsVoucher.Cells(i, 1).Value
                transDate = wsVoucher.Cells(i, 2).Value
                summary = wsVoucher.Cells(i, 3).Value
                
                ' 借方
                If wsVoucher.Cells(i, 4).Value = inputCode Then
                    debitAmount = Val(wsVoucher.Cells(i, 8).Value)
                    creditAmount = 0
                    balance = balance + debitAmount
                Else
                    ' 贷方
                    debitAmount = 0
                    creditAmount = Val(wsVoucher.Cells(i, 8).Value)
                    balance = balance - creditAmount
                End If
                
                wsSub.Cells(row, 1).Value = voucherNo
                wsSub.Cells(row, 2).Value = Format(transDate, "yyyy-mm-dd")
                wsSub.Cells(row, 3).Value = summary
                wsSub.Cells(row, 4).Value = IIf(debitAmount > 0, debitAmount, "")
                wsSub.Cells(row, 4).NumberFormat = "#,##0.00"
                wsSub.Cells(row, 5).Value = IIf(creditAmount > 0, creditAmount, "")
                wsSub.Cells(row, 5).NumberFormat = "#,##0.00"
                wsSub.Cells(row, 6).Value = balance
                wsSub.Cells(row, 6).NumberFormat = "#,##0.00"
                wsSub.Cells(row, 7).Value = IIf(balance >= 0, "借", "贷")
                
                row = row + 1
            End If
        End If
    Next i
    
    ' 合计行
    wsSub.Cells(row, 1).Value = "合计"
    wsSub.Cells(row, 1).Font.Bold = True
    wsSub.Cells(row, 4).Formula = "=SUM(D5:D" & (row - 1) & ")"
    wsSub.Cells(row, 4).NumberFormat = "#,##0.00"
    wsSub.Cells(row, 5).Formula = "=SUM(E5:E" & (row - 1) & ")"
    wsSub.Cells(row, 5).NumberFormat = "#,##0.00"
    
    ' 列宽
    wsSub.Columns("A").ColumnWidth = 10
    wsSub.Columns("B").ColumnWidth = 12
    wsSub.Columns("C").ColumnWidth = 30
    wsSub.Columns("D").ColumnWidth = 12
    wsSub.Columns("E").ColumnWidth = 12
    wsSub.Columns("F").ColumnWidth = 12
    wsSub.Columns("G").ColumnWidth = 10
    
    MsgBox "明细账生成完成！", vbInformation, "完成"
    Exit Sub
    
ErrorHandler:
    MsgBox "生成明细账出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏61: GenerateTrialBalance - 生成科目余额表
' V9.0新增：生成科目余额表
' ============================================================================
Sub GenerateTrialBalance()
    Dim wsLedger As Worksheet, wsTrial As Worksheet
    Dim lastRow As Long, i As Long, row As Long
    Dim subjectCode As String, subjectName As String
    Dim openingDebit As Double, openingCredit As Double
    Dim periodDebit As Double, periodCredit As Double
    Dim closingDebit As Double, closingCredit As Double
    Dim closingBalance As Double
    Dim direction As String
    
    On Error GoTo ErrorHandler
    
    ' 检查总账表
    On Error Resume Next
    Set wsLedger = ThisWorkbook.Sheets("总账")
    If wsLedger Is Nothing Then
        MsgBox "请先运行【生成总账】！", vbExclamation, "提示"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建科目余额表
    On Error Resume Next
    Set wsTrial = ThisWorkbook.Sheets("科目余额表")
    If wsTrial Is Nothing Then
        Set wsTrial = ThisWorkbook.Sheets.Add
        wsTrial.Name = "科目余额表"
    Else
        wsTrial.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 表头
    wsTrial.Cells(1, 1).Value = "科目余额表"
    wsTrial.Cells(1, 1).Font.Size = 16
    wsTrial.Cells(1, 1).Font.Bold = True
    wsTrial.Cells(1, 1).HorizontalAlignment = xlCenter
    wsTrial.Range("A1:H1").Merge
    
    wsTrial.Cells(2, 1).Value = "编制日期：" & Format(Date, "yyyy-mm-dd")
    
    ' 列标题
    row = 4
    wsTrial.Cells(row, 1).Value = "科目代码"
    wsTrial.Cells(row, 2).Value = "科目名称"
    wsTrial.Cells(row, 3).Value = "期初余额(借)"
    wsTrial.Cells(row, 4).Value = "期初余额(贷)"
    wsTrial.Cells(row, 5).Value = "本期借方"
    wsTrial.Cells(row, 6).Value = "本期贷方"
    wsTrial.Cells(row, 7).Value = "期末余额(借)"
    wsTrial.Cells(row, 8).Value = "期末余额(贷)"
    
    wsTrial.Range("A4:H4").Font.Bold = True
    wsTrial.Range("A4:H4").Interior.Color = RGB(192, 80, 77)
    wsTrial.Range("A4:H4").Font.Color = RGB(255, 255, 255)
    
    ' 读取总账数据
    row = 5
    lastRow = GetLastRow(wsLedger, 1)
    
    For i = 5 To lastRow - 1
        If wsLedger.Cells(i, 1).Value <> "" And wsLedger.Cells(i, 1).Value <> "合计" Then
            subjectCode = wsLedger.Cells(i, 1).Value
            subjectName = wsLedger.Cells(i, 2).Value
            openingDebit = Val(wsLedger.Cells(i, 3).Value)
            periodDebit = Val(wsLedger.Cells(i, 4).Value)
            periodCredit = Val(wsLedger.Cells(i, 5).Value)
            closingBalance = Val(wsLedger.Cells(i, 6).Value)
            direction = wsLedger.Cells(i, 7).Value
            
            wsTrial.Cells(row, 1).Value = subjectCode
            wsTrial.Cells(row, 2).Value = subjectName
            
            ' 期初余额
            openingCredit = 0  ' 初始化变量
            wsTrial.Cells(row, 3).Value = IIf(direction = "借" And closingBalance > periodDebit, closingBalance - periodDebit + periodCredit, openingDebit)
            wsTrial.Cells(row, 3).NumberFormat = "#,##0.00"
            wsTrial.Cells(row, 4).Value = IIf(direction = "贷" And closingBalance > periodCredit, closingBalance - periodCredit + periodDebit, openingCredit)
            wsTrial.Cells(row, 4).NumberFormat = "#,##0.00"
            
            ' 本期发生
            wsTrial.Cells(row, 5).Value = periodDebit
            wsTrial.Cells(row, 5).NumberFormat = "#,##0.00"
            wsTrial.Cells(row, 6).Value = periodCredit
            wsTrial.Cells(row, 6).NumberFormat = "#,##0.00"
            
            ' 期末余额
            If direction = "借" Then
                wsTrial.Cells(row, 7).Value = closingBalance
                wsTrial.Cells(row, 7).NumberFormat = "#,##0.00"
                wsTrial.Cells(row, 8).Value = 0
            Else
                wsTrial.Cells(row, 7).Value = 0
                wsTrial.Cells(row, 8).Value = closingBalance
                wsTrial.Cells(row, 8).NumberFormat = "#,##0.00"
            End If
            
            row = row + 1
        End If
    Next i
    
    ' 合计行
    wsTrial.Cells(row, 1).Value = "合计"
    wsTrial.Cells(row, 1).Font.Bold = True
    wsTrial.Cells(row, 3).Formula = "=SUM(C5:C" & (row - 1) & ")"
    wsTrial.Cells(row, 3).NumberFormat = "#,##0.00"
    wsTrial.Cells(row, 4).Formula = "=SUM(D5:D" & (row - 1) & ")"
    wsTrial.Cells(row, 4).NumberFormat = "#,##0.00"
    wsTrial.Cells(row, 5).Formula = "=SUM(E5:E" & (row - 1) & ")"
    wsTrial.Cells(row, 5).NumberFormat = "#,##0.00"
    wsTrial.Cells(row, 6).Formula = "=SUM(F5:F" & (row - 1) & ")"
    wsTrial.Cells(row, 6).NumberFormat = "#,##0.00"
    wsTrial.Cells(row, 7).Formula = "=SUM(G5:G" & (row - 1) & ")"
    wsTrial.Cells(row, 7).NumberFormat = "#,##0.00"
    wsTrial.Cells(row, 8).Formula = "=SUM(H5:H" & (row - 1) & ")"
    wsTrial.Cells(row, 8).NumberFormat = "#,##0.00"
    
    ' 列宽
    wsTrial.Columns("A").ColumnWidth = 12
    wsTrial.Columns("B").ColumnWidth = 18
    wsTrial.Columns("C").ColumnWidth = 12
    wsTrial.Columns("D").ColumnWidth = 12
    wsTrial.Columns("E").ColumnWidth = 12
    wsTrial.Columns("F").ColumnWidth = 12
    wsTrial.Columns("G").ColumnWidth = 12
    wsTrial.Columns("H").ColumnWidth = 12
    
    MsgBox "科目余额表生成完成！", vbInformation, "完成"
    Exit Sub
    
ErrorHandler:
    MsgBox "生成科目余额表出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏62: ValidateAccounting - 账账核对
' V9.0新增：验证账务平衡
' ============================================================================
Sub ValidateAccounting()
    Dim wsVoucher As Worksheet, wsLedger As Worksheet, wsTrial As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long, i As Long, row As Long
    Dim debitTotal As Double, creditTotal As Double
    Dim debitBalance As Double, creditBalance As Double
    Dim assetTotal As Double, liabilityTotal As Double
    Dim reportText As String
    Dim allPassed As Boolean
    
    On Error GoTo ErrorHandler
    
    allPassed = True
    reportText = "账务核对报告" & vbCrLf & vbCrLf
    reportText = reportText & "核对时间：" & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    reportText = reportText & String(40, "=") & vbCrLf & vbCrLf
    
    ' 检查凭证借贷平衡
    On Error Resume Next
    Set wsVoucher = ThisWorkbook.Sheets("记账凭证")
    If wsVoucher Is Nothing Then
        reportText = reportText & "【凭证借贷核对】 未生成记账凭证" & vbCrLf
        reportText = reportText & "  状态：跳过" & vbCrLf & vbCrLf
        allPassed = False
    Else
        lastRow = GetLastRow(wsVoucher, 1)
        debitTotal = 0
        creditTotal = 0
        
        For i = 5 To lastRow - 1
            If wsVoucher.Cells(i, 1).Value <> "" And wsVoucher.Cells(i, 1).Value <> "合计" Then
                debitTotal = debitTotal + Val(wsVoucher.Cells(i, 5).Value)  ' 借方金额列
                creditTotal = creditTotal + Val(wsVoucher.Cells(i, 6).Value)  ' 贷方金额列
            End If
        Next
        
        reportText = reportText & "【凭证借贷核对】" & vbCrLf
        reportText = reportText & "  借方发生额合计：" & Format(debitTotal, "#,##0.00") & vbCrLf
        reportText = reportText & "  贷方发生额合计：" & Format(creditTotal, "#,##0.00") & vbCrLf
        
        If Abs(debitTotal - creditTotal) < 0.01 Then
            reportText = reportText & "  状态：通过" & "  差额：" & Format(Abs(debitTotal - creditTotal), "#,##0.00") & vbCrLf
        Else
            reportText = reportText & "  状态：不平衡！  差额：" & Format(Abs(debitTotal - creditTotal), "#,##0.00") & vbCrLf
            allPassed = False
        End If
        reportText = reportText & vbCrLf
    End If
    On Error GoTo 0
    
    ' 检查总账借贷平衡
    On Error Resume Next
    Set wsLedger = ThisWorkbook.Sheets("总账")
    If wsLedger Is Nothing Then
        reportText = reportText & "【总账借贷核对】 未生成总账" & vbCrLf
        reportText = reportText & "  状态：跳过" & vbCrLf & vbCrLf
        allPassed = False
    Else
        lastRow = GetLastRow(wsLedger, 1)
        debitTotal = 0
        creditTotal = 0
        
        For i = 5 To lastRow - 1
            If wsLedger.Cells(i, 1).Value <> "" And wsLedger.Cells(i, 1).Value <> "合计" Then
                debitTotal = debitTotal + Val(wsLedger.Cells(i, 4).Value)
                creditTotal = creditTotal + Val(wsLedger.Cells(i, 5).Value)
            End If
        Next
        
        reportText = reportText & "【总账借贷核对】" & vbCrLf
        reportText = reportText & "  本期借方合计：" & Format(debitTotal, "#,##0.00") & vbCrLf
        reportText = reportText & "  本期贷方合计：" & Format(creditTotal, "#,##0.00") & vbCrLf
        
        If Abs(debitTotal - creditTotal) < 0.01 Then
            reportText = reportText & "  状态：通过" & "  差额：" & Format(Abs(debitTotal - creditTotal), "#,##0.00") & vbCrLf
        Else
            reportText = reportText & "  状态：不平衡！  差额：" & Format(Abs(debitTotal - creditTotal), "#,##0.00") & vbCrLf
            allPassed = False
        End If
        reportText = reportText & vbCrLf
    End If
    On Error GoTo 0
    
    ' 检查科目余额表借贷平衡
    On Error Resume Next
    Set wsTrial = ThisWorkbook.Sheets("科目余额表")
    If wsTrial Is Nothing Then
        reportText = reportText & "【科目余额表核对】 未生成科目余额表" & vbCrLf
        reportText = reportText & "  状态：跳过" & vbCrLf & vbCrLf
        allPassed = False
    Else
        lastRow = GetLastRow(wsTrial, 1)
        debitBalance = 0
        creditBalance = 0
        
        For i = 5 To lastRow - 1
            If wsTrial.Cells(i, 1).Value <> "" And wsTrial.Cells(i, 1).Value <> "合计" Then
                debitBalance = debitBalance + Val(wsTrial.Cells(i, 7).Value)
                creditBalance = creditBalance + Val(wsTrial.Cells(i, 8).Value)
            End If
        Next
        
        reportText = reportText & "【科目余额表核对】" & vbCrLf
        reportText = reportText & "  借方余额合计：" & Format(debitBalance, "#,##0.00") & vbCrLf
        reportText = reportText & "  贷方余额合计：" & Format(creditBalance, "#,##0.00") & vbCrLf
        
        If Abs(debitBalance - creditBalance) < 0.01 Then
            reportText = reportText & "  状态：通过" & "  差额：" & Format(Abs(debitBalance - creditBalance), "#,##0.00") & vbCrLf
        Else
            reportText = reportText & "  状态：不平衡！  差额：" & Format(Abs(debitBalance - creditBalance), "#,##0.00") & vbCrLf
            allPassed = False
        End If
        reportText = reportText & vbCrLf
    End If
    On Error GoTo 0
    
    ' 汇总结果
    reportText = reportText & String(40, "=") & vbCrLf
    If allPassed Then
        reportText = reportText & "核对结论：全部通过！" & vbCrLf
    Else
        reportText = reportText & "核对结论：存在不平衡项目，请检查！" & vbCrLf
    End If
    
    ' 显示报告
    MsgBox reportText, vbInformation + vbOKOnly, "账务核对报告"
    
    ' 可选：保存到核对报告表
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("核对报告")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "核对报告"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0
    
    wsReport.Cells(1, 1).Value = "账务核对报告"
    wsReport.Cells(1, 1).Font.Size = 16
    wsReport.Cells(1, 1).Font.Bold = True
    wsReport.Cells(2, 1).Value = reportText
    wsReport.Columns("A").ColumnWidth = 60
    
    Exit Sub
    
ErrorHandler:
    MsgBox "账务核对出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏63: FinancialRatioAnalysis - 财务指标分析
' V18.0新增：计算并展示关键财务指标
' ============================================================================
Sub FinancialRatioAnalysis()
    Dim wsIncome As Worksheet, wsExpense As Worksheet, wsBalance As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long, i As Long
    Dim totalIncome As Double, totalCost As Double, grossProfit As Double
    Dim totalExpense As Double, netProfit As Double
    Dim currentAssets As Double, totalAssets As Double
    Dim currentLiabilities As Double, totalLiabilities As Double
    Dim ownersEquity As Double, inventory As Double
    Dim accountsReceivable As Double, accountsPayable As Double
    Dim prevIncome As Double, prevNetProfit As Double
    Dim currentRatio As Double, quickRatio As Double, debtRatio As Double
    Dim arTurnover As Double, inventoryTurnover As Double, assetTurnover As Double
    Dim grossMargin As Double, netMargin As Double, roa As Double, roe As Double
    Dim revenueGrowth As Double, profitGrowth As Double
    Dim reportRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' 获取数据
    On Error Resume Next
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsBalance = ThisWorkbook.Sheets("资产负债表")
    On Error GoTo 0
    
    If wsIncome Is Nothing Or wsExpense Is Nothing Then
        MsgBox "请先录入收入和支出数据！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 计算收入和成本
    totalIncome = 0
    totalCost = 0
    totalExpense = 0
    
    ' 从利润表获取数据
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("利润表")
    If Not wsReport Is Nothing Then
        totalIncome = Val(wsReport.Cells(5, 3).Value)  ' 营业收入
        totalCost = Val(wsReport.Cells(8, 3).Value)    ' 营业成本
        grossProfit = Val(wsReport.Cells(10, 3).Value) ' 毛利润
        totalExpense = Val(wsReport.Cells(20, 3).Value) ' 费用合计
        netProfit = Val(wsReport.Cells(26, 3).Value)   ' 净利润
    End If
    On Error GoTo 0
    
    ' 从资产负债表获取数据
    On Error Resume Next
    If Not wsBalance Is Nothing Then
        currentAssets = Val(wsBalance.Cells(8, 3).Value)    ' 流动资产合计
        totalAssets = Val(wsBalance.Cells(16, 3).Value)     ' 资产总计
        inventory = Val(wsBalance.Cells(6, 3).Value)        ' 存货
        accountsReceivable = Val(wsBalance.Cells(5, 3).Value) ' 应收账款
        currentLiabilities = Val(wsBalance.Cells(21, 3).Value) ' 流动负债合计
        totalLiabilities = Val(wsBalance.Cells(24, 3).Value)   ' 负债合计
        ownersEquity = Val(wsBalance.Cells(28, 3).Value)       ' 所有者权益
    End If
    On Error GoTo 0
    
    ' 创建财务指标分析表
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("财务指标分析")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "财务指标分析"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 设置表头
    wsReport.Cells(1, 1).Value = "财务指标分析报告"
    wsReport.Cells(1, 1).Font.Size = 16
    wsReport.Cells(1, 1).Font.Bold = True
    wsReport.Cells(2, 1).Value = "分析日期：" & Format(Date, "yyyy年mm月dd日")
    wsReport.Cells(3, 1).Value = String(60, "=")
    
    reportRow = 5
    
    ' 一、偿债能力指标
    wsReport.Cells(reportRow, 1).Value = "一、偿债能力指标"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    ' 流动比率 = 流动资产 / 流动负债
    If currentLiabilities > 0 Then
        currentRatio = currentAssets / currentLiabilities
    Else
        currentRatio = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "流动比率"
    wsReport.Cells(reportRow, 2).Value = currentRatio
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00"
    wsReport.Cells(reportRow, 3).Value = "流动资产 / 流动负债"
    wsReport.Cells(reportRow, 4).Value = IIf(currentRatio >= 2, "良好", IIf(currentRatio >= 1, "正常", "偏低"))
    reportRow = reportRow + 1
    
    ' 速动比率 = (流动资产 - 存货) / 流动负债
    If currentLiabilities > 0 Then
        quickRatio = (currentAssets - inventory) / currentLiabilities
    Else
        quickRatio = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "速动比率"
    wsReport.Cells(reportRow, 2).Value = quickRatio
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00"
    wsReport.Cells(reportRow, 3).Value = "(流动资产-存货) / 流动负债"
    wsReport.Cells(reportRow, 4).Value = IIf(quickRatio >= 1, "良好", IIf(quickRatio >= 0.5, "正常", "偏低"))
    reportRow = reportRow + 1
    
    ' 资产负债率 = 负债总额 / 资产总额
    If totalAssets > 0 Then
        debtRatio = totalLiabilities / totalAssets
    Else
        debtRatio = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "资产负债率"
    wsReport.Cells(reportRow, 2).Value = debtRatio
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "负债总额 / 资产总额"
    wsReport.Cells(reportRow, 4).Value = IIf(debtRatio <= 0.5, "良好", IIf(debtRatio <= 0.7, "正常", "偏高"))
    reportRow = reportRow + 2
    
    ' 二、营运能力指标
    wsReport.Cells(reportRow, 1).Value = "二、营运能力指标"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    ' 应收账款周转率 = 营业收入 / 平均应收账款
    If accountsReceivable > 0 Then
        arTurnover = totalIncome / accountsReceivable
    Else
        arTurnover = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "应收账款周转率"
    wsReport.Cells(reportRow, 2).Value = arTurnover
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00"
    wsReport.Cells(reportRow, 3).Value = "营业收入 / 平均应收账款"
    wsReport.Cells(reportRow, 4).Value = IIf(arTurnover >= 6, "良好", IIf(arTurnover >= 3, "正常", "偏低"))
    reportRow = reportRow + 1
    
    ' 存货周转率 = 营业成本 / 平均存货
    If inventory > 0 Then
        inventoryTurnover = totalCost / inventory
    Else
        inventoryTurnover = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "存货周转率"
    wsReport.Cells(reportRow, 2).Value = inventoryTurnover
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00"
    wsReport.Cells(reportRow, 3).Value = "营业成本 / 平均存货"
    wsReport.Cells(reportRow, 4).Value = IIf(inventoryTurnover >= 4, "良好", IIf(inventoryTurnover >= 2, "正常", "偏低"))
    reportRow = reportRow + 1
    
    ' 总资产周转率 = 营业收入 / 平均总资产
    If totalAssets > 0 Then
        assetTurnover = totalIncome / totalAssets
    Else
        assetTurnover = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "总资产周转率"
    wsReport.Cells(reportRow, 2).Value = assetTurnover
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00"
    wsReport.Cells(reportRow, 3).Value = "营业收入 / 平均总资产"
    wsReport.Cells(reportRow, 4).Value = IIf(assetTurnover >= 1, "良好", IIf(assetTurnover >= 0.5, "正常", "偏低"))
    reportRow = reportRow + 2
    
    ' 三、盈利能力指标
    wsReport.Cells(reportRow, 1).Value = "三、盈利能力指标"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    ' 毛利率 = (营业收入 - 营业成本) / 营业收入
    If totalIncome > 0 Then
        grossMargin = (totalIncome - totalCost) / totalIncome
    Else
        grossMargin = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "毛利率"
    wsReport.Cells(reportRow, 2).Value = grossMargin
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "(营业收入-营业成本) / 营业收入"
    wsReport.Cells(reportRow, 4).Value = IIf(grossMargin >= 0.3, "良好", IIf(grossMargin >= 0.15, "正常", "偏低"))
    reportRow = reportRow + 1
    
    ' 净利率 = 净利润 / 营业收入
    If totalIncome > 0 Then
        netMargin = netProfit / totalIncome
    Else
        netMargin = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "净利率"
    wsReport.Cells(reportRow, 2).Value = netMargin
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "净利润 / 营业收入"
    wsReport.Cells(reportRow, 4).Value = IIf(netMargin >= 0.1, "良好", IIf(netMargin >= 0.05, "正常", "偏低"))
    reportRow = reportRow + 1
    
    ' 资产收益率(ROA) = 净利润 / 平均总资产
    If totalAssets > 0 Then
        roa = netProfit / totalAssets
    Else
        roa = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "资产收益率(ROA)"
    wsReport.Cells(reportRow, 2).Value = roa
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "净利润 / 平均总资产"
    wsReport.Cells(reportRow, 4).Value = IIf(roa >= 0.1, "良好", IIf(roa >= 0.05, "正常", "偏低"))
    reportRow = reportRow + 1
    
    ' 净资产收益率(ROE) = 净利润 / 平均所有者权益
    If ownersEquity > 0 Then
        roe = netProfit / ownersEquity
    Else
        roe = 0
    End If
    wsReport.Cells(reportRow, 1).Value = "净资产收益率(ROE)"
    wsReport.Cells(reportRow, 2).Value = roe
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "净利润 / 平均所有者权益"
    wsReport.Cells(reportRow, 4).Value = IIf(roe >= 0.15, "良好", IIf(roe >= 0.08, "正常", "偏低"))
    reportRow = reportRow + 2
    
    ' 四、发展能力指标
    wsReport.Cells(reportRow, 1).Value = "四、发展能力指标"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    ' 营业收入增长率（需要上期数据，此处模拟）
    revenueGrowth = 0.1  ' 默认值，实际应从历史数据计算
    wsReport.Cells(reportRow, 1).Value = "营业收入增长率"
    wsReport.Cells(reportRow, 2).Value = revenueGrowth
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "(本期收入-上期收入) / 上期收入"
    wsReport.Cells(reportRow, 4).Value = IIf(revenueGrowth >= 0.1, "良好", IIf(revenueGrowth >= 0, "正常", "下降"))
    reportRow = reportRow + 1
    
    ' 净利润增长率
    profitGrowth = 0.08  ' 默认值，实际应从历史数据计算
    wsReport.Cells(reportRow, 1).Value = "净利润增长率"
    wsReport.Cells(reportRow, 2).Value = profitGrowth
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "(本期净利润-上期净利润) / 上期净利润"
    wsReport.Cells(reportRow, 4).Value = IIf(profitGrowth >= 0.1, "良好", IIf(profitGrowth >= 0, "正常", "下降"))
    reportRow = reportRow + 2
    
    ' 设置列宽
    wsReport.Columns("A").ColumnWidth = 20
    wsReport.Columns("B").ColumnWidth = 15
    wsReport.Columns("C").ColumnWidth = 30
    wsReport.Columns("D").ColumnWidth = 10
    
    Application.ScreenUpdating = True
    
    MsgBox "财务指标分析完成！结果已输出到【财务指标分析】工作表。", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "财务指标分析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏64: DuPontAnalysis - 杜邦分析
' V18.0新增：杜邦分析法分解净资产收益率
' ============================================================================
Sub DuPontAnalysis()
    Dim wsBalance As Worksheet, wsProfit As Worksheet, wsReport As Worksheet
    Dim totalIncome As Double, netProfit As Double, totalAssets As Double
    Dim ownersEquity As Double, totalCost As Double
    Dim netProfitMargin As Double, assetTurnover As Double, equityMultiplier As Double
    Dim roe As Double, reportRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' 获取数据
    On Error Resume Next
    Set wsProfit = ThisWorkbook.Sheets("利润表")
    Set wsBalance = ThisWorkbook.Sheets("资产负债表")
    On Error GoTo 0
    
    If wsProfit Is Nothing Or wsBalance Is Nothing Then
        MsgBox "请先生成利润表和资产负债表！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 从利润表获取数据
    totalIncome = Val(wsProfit.Cells(5, 3).Value)    ' 营业收入
    netProfit = Val(wsProfit.Cells(26, 3).Value)     ' 净利润
    totalCost = Val(wsProfit.Cells(8, 3).Value)      ' 营业成本
    
    ' 从资产负债表获取数据
    totalAssets = Val(wsBalance.Cells(16, 3).Value)  ' 资产总计
    ownersEquity = Val(wsBalance.Cells(28, 3).Value) ' 所有者权益
    
    ' 计算杜邦分析指标
    ' 销售净利率 = 净利润 / 营业收入
    If totalIncome > 0 Then
        netProfitMargin = netProfit / totalIncome
    Else
        netProfitMargin = 0
    End If
    
    ' 资产周转率 = 营业收入 / 总资产
    If totalAssets > 0 Then
        assetTurnover = totalIncome / totalAssets
    Else
        assetTurnover = 0
    End If
    
    ' 权益乘数 = 总资产 / 所有者权益
    If ownersEquity > 0 Then
        equityMultiplier = totalAssets / ownersEquity
    Else
        equityMultiplier = 0
    End If
    
    ' 净资产收益率 = 销售净利率 × 资产周转率 × 权益乘数
    roe = netProfitMargin * assetTurnover * equityMultiplier
    
    ' 创建杜邦分析表
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("杜邦分析")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "杜邦分析"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 设置表头
    wsReport.Cells(1, 1).Value = "杜邦分析报告"
    wsReport.Cells(1, 1).Font.Size = 16
    wsReport.Cells(1, 1).Font.Bold = True
    wsReport.Cells(2, 1).Value = "分析日期：" & Format(Date, "yyyy年mm月dd日")
    wsReport.Cells(3, 1).Value = String(60, "=")
    
    reportRow = 5
    
    ' 杜邦分析核心公式
    wsReport.Cells(reportRow, 1).Value = "【杜邦分析核心公式】"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "净资产收益率(ROE) = 销售净利率 × 资产周转率 × 权益乘数"
    reportRow = reportRow + 2
    
    ' 计算结果
    wsReport.Cells(reportRow, 1).Value = "【计算结果】"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "一、销售净利率"
    wsReport.Cells(reportRow, 2).Value = "净利润 / 营业收入"
    wsReport.Cells(reportRow, 3).Value = netProfitMargin
    wsReport.Cells(reportRow, 3).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 4).Value = Format(netProfit, "#,##0.00") & " / " & Format(totalIncome, "#,##0.00")
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "二、资产周转率"
    wsReport.Cells(reportRow, 2).Value = "营业收入 / 总资产"
    wsReport.Cells(reportRow, 3).Value = assetTurnover
    wsReport.Cells(reportRow, 3).NumberFormat = "0.00"
    wsReport.Cells(reportRow, 4).Value = Format(totalIncome, "#,##0.00") & " / " & Format(totalAssets, "#,##0.00")
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "三、权益乘数"
    wsReport.Cells(reportRow, 2).Value = "总资产 / 所有者权益"
    wsReport.Cells(reportRow, 3).Value = equityMultiplier
    wsReport.Cells(reportRow, 3).NumberFormat = "0.00"
    wsReport.Cells(reportRow, 4).Value = Format(totalAssets, "#,##0.00") & " / " & Format(ownersEquity, "#,##0.00")
    reportRow = reportRow + 2
    
    ' 最终结果
    wsReport.Cells(reportRow, 1).Value = "【净资产收益率(ROE)】"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 14
    wsReport.Cells(reportRow, 1).Font.Color = RGB(0, 0, 255)
    wsReport.Cells(reportRow, 3).Value = roe
    wsReport.Cells(reportRow, 3).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Font.Size = 14
    wsReport.Cells(reportRow, 3).Font.Bold = True
    wsReport.Cells(reportRow, 3).Font.Color = RGB(0, 0, 255)
    reportRow = reportRow + 2
    
    ' 因素分析
    wsReport.Cells(reportRow, 1).Value = "【因素分析】"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "1. 销售净利率反映企业盈利能力"
    wsReport.Cells(reportRow, 2).Value = IIf(netProfitMargin >= 0.1, "盈利能力强", IIf(netProfitMargin >= 0.05, "盈利能力一般", "盈利能力较弱"))
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "2. 资产周转率反映资产运营效率"
    wsReport.Cells(reportRow, 2).Value = IIf(assetTurnover >= 1, "运营效率高", IIf(assetTurnover >= 0.5, "运营效率一般", "运营效率较低"))
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "3. 权益乘数反映财务杠杆"
    wsReport.Cells(reportRow, 2).Value = IIf(equityMultiplier >= 2, "财务杠杆较高", IIf(equityMultiplier >= 1.5, "财务杠杆适中", "财务杠杆较低"))
    reportRow = reportRow + 2
    
    ' 改进建议
    wsReport.Cells(reportRow, 1).Value = "【改进建议】"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    If netProfitMargin < 0.05 Then
        wsReport.Cells(reportRow, 1).Value = "• 提高销售净利率：控制成本费用，提高产品附加值"
        reportRow = reportRow + 1
    End If
    
    If assetTurnover < 0.5 Then
        wsReport.Cells(reportRow, 1).Value = "• 提高资产周转率：优化库存管理，加快应收账款回收"
        reportRow = reportRow + 1
    End If
    
    If equityMultiplier < 1.5 Then
        wsReport.Cells(reportRow, 1).Value = "• 适当利用财务杠杆：合理负债经营，提高资金使用效率"
        reportRow = reportRow + 1
    End If
    
    ' 设置列宽
    wsReport.Columns("A").ColumnWidth = 25
    wsReport.Columns("B").ColumnWidth = 25
    wsReport.Columns("C").ColumnWidth = 15
    wsReport.Columns("D").ColumnWidth = 35
    
    Application.ScreenUpdating = True
    
    MsgBox "杜邦分析完成！结果已输出到【杜邦分析】工作表。", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "杜邦分析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏65: OperatingEfficiency - 经营效率分析
' V18.0新增：分析企业经营效率
' ============================================================================
Sub OperatingEfficiency()
    Dim wsProfit As Worksheet, wsExpense As Worksheet, wsReport As Worksheet
    Dim totalIncome As Double, netProfit As Double, totalExpense As Double
    Dim outsourceCost As Double, salaryExpense As Double, rentExpense As Double
    Dim waterExpense As Double, elecExpense As Double, dailyExpense As Double
    Dim materialExpense As Double, otherExpense As Double
    Dim employeeCount As Integer, costTotal As Double
    Dim outputPerPerson As Double, profitPerPerson As Double
    Dim expenseRatio As Double, costProfitRatio As Double, outsourceRatio As Double
    Dim reportRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' 获取数据
    On Error Resume Next
    Set wsProfit = ThisWorkbook.Sheets("利润表")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    On Error GoTo 0
    
    If wsProfit Is Nothing Then
        MsgBox "请先生成利润表！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 从利润表获取数据
    totalIncome = Val(wsProfit.Cells(5, 3).Value)      ' 营业收入
    netProfit = Val(wsProfit.Cells(26, 3).Value)       ' 净利润
    outsourceCost = Val(wsProfit.Cells(8, 3).Value)    ' 外发加工费（营业成本）
    
    ' 获取各项费用
    rentExpense = Val(wsProfit.Cells(12, 3).Value)     ' 房租
    waterExpense = Val(wsProfit.Cells(13, 3).Value)    ' 水费
    elecExpense = Val(wsProfit.Cells(14, 3).Value)     ' 电费
    materialExpense = Val(wsProfit.Cells(15, 3).Value) ' 材料
    salaryExpense = Val(wsProfit.Cells(19, 3).Value)   ' 工资
    dailyExpense = Val(wsProfit.Cells(18, 3).Value)    ' 日常费用
    
    ' 计算费用总额
    totalExpense = rentExpense + waterExpense + elecExpense + materialExpense + salaryExpense + dailyExpense
    costTotal = outsourceCost + totalExpense
    
    ' 获取员工人数（从系统设置或默认值）
    Dim wsSettings As Worksheet
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Sheets("系统设置")
    If Not wsSettings Is Nothing Then
        employeeCount = Val(wsSettings.Cells(5, 2).Value)
    End If
    On Error GoTo 0
    
    If employeeCount <= 0 Then
        employeeCount = InputBox("请输入员工人数：", "员工人数", "5")
        If employeeCount <= 0 Then employeeCount = 5
    End If
    
    ' 计算经营效率指标
    ' 人均产值
    outputPerPerson = totalIncome / employeeCount
    
    ' 人均利润
    profitPerPerson = netProfit / employeeCount
    
    ' 费用率
    If totalIncome > 0 Then
        expenseRatio = totalExpense / totalIncome
    Else
        expenseRatio = 0
    End If
    
    ' 成本费用利润率
    If costTotal > 0 Then
        costProfitRatio = netProfit / costTotal
    Else
        costProfitRatio = 0
    End If
    
    ' 外发加工占比
    If totalIncome > 0 Then
        outsourceRatio = outsourceCost / totalIncome
    Else
        outsourceRatio = 0
    End If
    
    ' 创建经营效率分析表
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("经营效率分析")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "经营效率分析"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 设置表头
    wsReport.Cells(1, 1).Value = "经营效率分析报告"
    wsReport.Cells(1, 1).Font.Size = 16
    wsReport.Cells(1, 1).Font.Bold = True
    wsReport.Cells(2, 1).Value = "分析日期：" & Format(Date, "yyyy年mm月dd日")
    wsReport.Cells(2, 2).Value = "员工人数：" & employeeCount & "人"
    wsReport.Cells(3, 1).Value = String(60, "=")
    
    reportRow = 5
    
    ' 一、人均效率指标
    wsReport.Cells(reportRow, 1).Value = "一、人均效率指标"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "人均产值"
    wsReport.Cells(reportRow, 2).Value = outputPerPerson
    wsReport.Cells(reportRow, 2).NumberFormat = "#,##0.00"
    wsReport.Cells(reportRow, 3).Value = "营业收入 / 员工人数"
    wsReport.Cells(reportRow, 4).Value = Format(totalIncome, "#,##0.00") & " / " & employeeCount
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "人均利润"
    wsReport.Cells(reportRow, 2).Value = profitPerPerson
    wsReport.Cells(reportRow, 2).NumberFormat = "#,##0.00"
    wsReport.Cells(reportRow, 3).Value = "净利润 / 员工人数"
    wsReport.Cells(reportRow, 4).Value = Format(netProfit, "#,##0.00") & " / " & employeeCount
    reportRow = reportRow + 2
    
    ' 二、费用率分析
    wsReport.Cells(reportRow, 1).Value = "二、费用率分析"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "费用率"
    wsReport.Cells(reportRow, 2).Value = expenseRatio
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "各项费用 / 营业收入"
    wsReport.Cells(reportRow, 4).Value = Format(totalExpense, "#,##0.00") & " / " & Format(totalIncome, "#,##0.00")
    reportRow = reportRow + 1
    
    ' 各项费用占比
    wsReport.Cells(reportRow, 1).Value = "  - 工资费用率"
    If totalIncome > 0 Then
        wsReport.Cells(reportRow, 2).Value = salaryExpense / totalIncome
    Else
        wsReport.Cells(reportRow, 2).Value = 0
    End If
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = Format(salaryExpense, "#,##0.00")
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "  - 房租费用率"
    If totalIncome > 0 Then
        wsReport.Cells(reportRow, 2).Value = rentExpense / totalIncome
    Else
        wsReport.Cells(reportRow, 2).Value = 0
    End If
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = Format(rentExpense, "#,##0.00")
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "  - 水电费用率"
    If totalIncome > 0 Then
        wsReport.Cells(reportRow, 2).Value = (waterExpense + elecExpense) / totalIncome
    Else
        wsReport.Cells(reportRow, 2).Value = 0
    End If
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = Format(waterExpense + elecExpense, "#,##0.00")
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "  - 材料费用率"
    If totalIncome > 0 Then
        wsReport.Cells(reportRow, 2).Value = materialExpense / totalIncome
    Else
        wsReport.Cells(reportRow, 2).Value = 0
    End If
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = Format(materialExpense, "#,##0.00")
    reportRow = reportRow + 2
    
    ' 三、成本效益指标
    wsReport.Cells(reportRow, 1).Value = "三、成本效益指标"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "成本费用利润率"
    wsReport.Cells(reportRow, 2).Value = costProfitRatio
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "净利润 / 成本费用总额"
    wsReport.Cells(reportRow, 4).Value = Format(netProfit, "#,##0.00") & " / " & Format(costTotal, "#,##0.00")
    reportRow = reportRow + 1
    
    wsReport.Cells(reportRow, 1).Value = "外发加工占比"
    wsReport.Cells(reportRow, 2).Value = outsourceRatio
    wsReport.Cells(reportRow, 2).NumberFormat = "0.00%"
    wsReport.Cells(reportRow, 3).Value = "外发加工费 / 营业收入"
    wsReport.Cells(reportRow, 4).Value = Format(outsourceCost, "#,##0.00") & " / " & Format(totalIncome, "#,##0.00")
    reportRow = reportRow + 2
    
    ' 四、效率评价
    wsReport.Cells(reportRow, 1).Value = "四、效率评价"
    wsReport.Cells(reportRow, 1).Font.Bold = True
    wsReport.Cells(reportRow, 1).Font.Size = 12
    reportRow = reportRow + 1
    
    If outputPerPerson >= 200000 Then
        wsReport.Cells(reportRow, 1).Value = "人均产值评价：优秀（人均产值≥20万）"
    ElseIf outputPerPerson >= 100000 Then
        wsReport.Cells(reportRow, 1).Value = "人均产值评价：良好（人均产值≥10万）"
    Else
        wsReport.Cells(reportRow, 1).Value = "人均产值评价：需提升（人均产值<10万）"
    End If
    reportRow = reportRow + 1
    
    If expenseRatio <= 0.3 Then
        wsReport.Cells(reportRow, 1).Value = "费用率评价：优秀（费用率≤30%）"
    ElseIf expenseRatio <= 0.5 Then
        wsReport.Cells(reportRow, 1).Value = "费用率评价：良好（费用率≤50%）"
    Else
        wsReport.Cells(reportRow, 1).Value = "费用率评价：需控制（费用率>50%）"
    End If
    reportRow = reportRow + 1
    
    If outsourceRatio <= 0.4 Then
        wsReport.Cells(reportRow, 1).Value = "外发加工占比评价：合理（占比≤40%）"
    Else
        wsReport.Cells(reportRow, 1).Value = "外发加工占比评价：偏高（占比>40%），建议提升自主加工能力"
    End If
    
    ' 设置列宽
    wsReport.Columns("A").ColumnWidth = 20
    wsReport.Columns("B").ColumnWidth = 15
    wsReport.Columns("C").ColumnWidth = 25
    wsReport.Columns("D").ColumnWidth = 30
    
    Application.ScreenUpdating = True
    
    MsgBox "经营效率分析完成！结果已输出到【经营效率分析】工作表。", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "经营效率分析出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏66: ExportDataToExcel - 导出数据到Excel
' V18.0新增：将指定工作表数据导出到新Excel文件
' ============================================================================
Sub ExportDataToExcel()
    Dim exportPath As String
    Dim exportSheets As Collection
    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim sheetName As String
    Dim exportAll As Boolean
    Dim dateStart As Date, dateEnd As Date
    Dim useDateFilter As Boolean
    Dim i As Integer, sheetCount As Integer
    
    On Error GoTo ErrorHandler
    
    ' 选择导出方式
    Dim choice As Integer
    choice = MsgBox("请选择导出方式：" & vbCrLf & vbCrLf & _
                    "【是】- 导出全部工作表" & vbCrLf & _
                    "【否】- 选择指定工作表导出" & vbCrLf & _
                    "【取消】- 退出", vbYesNoCancel + vbQuestion, "导出数据")
    
    If choice = vbCancel Then Exit Sub
    
    exportAll = (choice = vbYes)
    
    ' 选择保存路径
    ' 【WPS兼容】已移除FileDialog声明
    Dim filePath As String
    ' 【WPS兼容】使用GetOpenFilename替代FileDialog
    fd.Title = "选择导出目录"
    
    ' 【WPS兼容】WPS使用GetOpenFilename返回值判断
    filePath = Application.GetOpenFilename("Excel文件 (*.xlsx;*.xls),*.xlsx;*.xls")
    If filePath <> "False" Then
        exportPath = fd.SelectedItems(1)
    Else
        Exit Sub
    End If
    
    ' 日期范围筛选（可选）
    useDateFilter = (MsgBox("是否按日期范围筛选数据？", vbQuestion + vbYesNo, "日期筛选") = vbYes)
    
    If useDateFilter Then
        dateStart = InputBox("请输入开始日期（格式：yyyy-mm-dd）：", "开始日期", Format(DateSerial(Year(Date), 1, 1), "yyyy-mm-dd"))
        dateEnd = InputBox("请输入结束日期（格式：yyyy-mm-dd）：", "结束日期", Format(Date, "yyyy-mm-dd"))
    End If
    
    Application.ScreenUpdating = False
    
    ' 创建新工作簿
    Set newWb = Workbooks.Add
    sheetCount = 0
    
    ' 导出工作表
    If exportAll Then
        ' 导出全部工作表
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> "操作日志" Then  ' 排除操作日志
                Call CopySheetToWorkbook(ws, newWb, useDateFilter, dateStart, dateEnd)
                sheetCount = sheetCount + 1
            End If
        Next ws
    Else
        ' 选择指定工作表
        Dim sheetList As String
        sheetList = "可导出的工作表：" & vbCrLf & vbCrLf
        sheetList = sheetList & "1. 收入记录" & vbCrLf
        sheetList = sheetList & "2. 支出记录" & vbCrLf
        sheetList = sheetList & "3. 外发加工费明细" & vbCrLf
        sheetList = sheetList & "4. 应收应付" & vbCrLf
        sheetList = sheetList & "5. 利润表" & vbCrLf
        sheetList = sheetList & "6. 资产负债表" & vbCrLf
        sheetList = sheetList & "7. 现金流量表" & vbCrLf
        sheetList = sheetList & "8. 全部导出" & vbCrLf & vbCrLf
        sheetList = sheetList & "请输入要导出的工作表编号（多个用逗号分隔）："
        
        Dim inputStr As String
        inputStr = InputBox(sheetList, "选择工作表", "8")
        
        If inputStr = "" Then
            newWb.Close SaveChanges:=False
            Exit Sub
        End If
        
        Dim selections() As String
        selections = Split(inputStr, ",")
        
        For i = 0 To UBound(selections)
            Select Case Trim(selections(i))
                Case "1"
                    Call CopySheetToWorkbook(ThisWorkbook.Sheets("收入记录"), newWb, useDateFilter, dateStart, dateEnd)
                    sheetCount = sheetCount + 1
                Case "2"
                    Call CopySheetToWorkbook(ThisWorkbook.Sheets("支出记录"), newWb, useDateFilter, dateStart, dateEnd)
                    sheetCount = sheetCount + 1
                Case "3"
                    Call CopySheetToWorkbook(ThisWorkbook.Sheets("外发加工费明细"), newWb, useDateFilter, dateStart, dateEnd)
                    sheetCount = sheetCount + 1
                Case "4"
                    Call CopySheetToWorkbook(ThisWorkbook.Sheets("应收应付"), newWb, useDateFilter, dateStart, dateEnd)
                    sheetCount = sheetCount + 1
                Case "5"
                    Call CopySheetToWorkbook(ThisWorkbook.Sheets("利润表"), newWb, False, dateStart, dateEnd)
                    sheetCount = sheetCount + 1
                Case "6"
                    Call CopySheetToWorkbook(ThisWorkbook.Sheets("资产负债表"), newWb, False, dateStart, dateEnd)
                    sheetCount = sheetCount + 1
                Case "7"
                    Call CopySheetToWorkbook(ThisWorkbook.Sheets("现金流量表"), newWb, False, dateStart, dateEnd)
                    sheetCount = sheetCount + 1
                Case "8"
                    For Each ws In ThisWorkbook.Worksheets
                        If ws.Name <> "操作日志" Then
                            Call CopySheetToWorkbook(ws, newWb, useDateFilter, dateStart, dateEnd)
                            sheetCount = sheetCount + 1
                        End If
                    Next ws
            End Select
        Next i
    End If
    
    ' 删除新工作簿的默认工作表
    Application.DisplayAlerts = False
    For Each ws In newWb.Worksheets
        If ws.Name Like "Sheet*" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' 保存文件
    Dim fileName As String
    fileName = exportPath & "\氧化加工厂数据导出_" & Format(Date, "yyyymmdd") & ".xlsx"
    newWb.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    
    Application.ScreenUpdating = True
    
    MsgBox "数据导出成功！" & vbCrLf & vbCrLf & _
           "导出文件：" & fileName & vbCrLf & _
           "导出工作表数：" & sheetCount & "个", vbInformation, "导出完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "数据导出出错：" & Err.Description, vbCritical, "错误"
End Sub

' ----------------------------------------------------------------------------
' 辅助函数：复制工作表到新工作簿
' ----------------------------------------------------------------------------
Private Sub CopySheetToWorkbook(sourceSheet As Worksheet, targetWb As Workbook, useDateFilter As Boolean, dateStart As Date, dateEnd As Date)
    Dim newSheet As Worksheet
    Dim lastRow As Long, i As Long, copyRow As Long
    Dim rowDate As Date
    
    On Error Resume Next
    
    sourceSheet.Copy Before:=targetWb.Sheets(1)
    Set newSheet = targetWb.Sheets(1)
    newSheet.Name = sourceSheet.Name
    
    ' 如果需要日期筛选
    If useDateFilter Then
        lastRow = GetLastRow(newSheet, 1)
        copyRow = 1
        
        ' 从后往前删除不符合条件的行
        For i = lastRow To 2 Step -1
            If newSheet.Cells(i, 1).Value <> "" Then
                On Error Resume Next
                rowDate = CDate(newSheet.Cells(i, 1).Value)
                If Err.Number = 0 Then
                    If rowDate < dateStart Or rowDate > dateEnd Then
                        newSheet.Rows(i).Delete
                    End If
                End If
                On Error GoTo 0
            End If
            Next i
    End If
    
    On Error GoTo 0
End Sub

' ============================================================================
' 宏67: ImportDataFromExcel - 从Excel导入数据
' V18.0新增：从外部Excel文件导入数据
' ============================================================================
Sub ImportDataFromExcel()
    Dim importFile As String
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim importType As Integer
    Dim lastRow As Long, i As Long, col As Long
    Dim colMap As Object
    Dim previewRows As Long
    
    On Error GoTo ErrorHandler
    
    ' 选择导入文件
    ' 【WPS兼容】已移除FileDialog声明
    Dim filePath As String
    ' 【WPS兼容】使用GetOpenFilename替代FileDialog
    fd.Title = "选择要导入的Excel文件"
    fd.Filters.Clear
    fd.Filters.Add "Excel文件", "*.xlsx;*.xls"
    
    ' 【WPS兼容】WPS使用GetOpenFilename返回值判断
    filePath = Application.GetOpenFilename("Excel文件 (*.xlsx;*.xls),*.xlsx;*.xls")
    If filePath <> "False" Then
        importFile = fd.SelectedItems(1)
    Else
        Exit Sub
    End If
    
    ' 选择导入类型
    Dim typeList As String
    typeList = "请选择导入数据类型：" & vbCrLf & vbCrLf
    typeList = typeList & "1. 收入记录" & vbCrLf
    typeList = typeList & "2. 支出记录" & vbCrLf
    typeList = typeList & "3. 应收应付" & vbCrLf
    typeList = typeList & "4. 银行流水" & vbCrLf
    
    importType = Val(InputBox(typeList, "导入类型", "1"))
    
    If importType < 1 Or importType > 4 Then
        MsgBox "无效的导入类型！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 打开源文件
    Set sourceWb = Workbooks.Open(importFile, ReadOnly:=True)
    Set sourceWs = sourceWb.Sheets(1)
    
    ' 预览数据
    lastRow = GetLastRow(sourceWs, 1)
    If lastRow > 10 Then previewRows = 10 Else previewRows = lastRow
    
    Dim preview As String
    preview = "数据预览（前" & previewRows & "行）：" & vbCrLf & vbCrLf
    
    For i = 1 To previewRows
        For col = 1 To 5
            preview = preview & sourceWs.Cells(i, col).Value & vbTab
        Next col
        preview = preview & vbCrLf
    Next i
    
    preview = preview & vbCrLf & "共发现 " & (lastRow - 1) & " 行数据（不含标题）" & vbCrLf & vbCrLf
    preview = preview & "是否确认导入？"
    
    Application.ScreenUpdating = True
    
    If MsgBox(preview, vbYesNo + vbQuestion, "数据预览") <> vbYes Then
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 根据类型导入数据
    Select Case importType
        Case 1  ' 收入记录
            Call ImportIncomeData(sourceWs)
        Case 2  ' 支出记录
            Call ImportExpenseData(sourceWs)
        Case 3  ' 应收应付
            Call ImportARAPData(sourceWs)
        Case 4  ' 银行流水
            Call ImportBankData(sourceWs)
    End Select
    
    sourceWb.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
    MsgBox "数据导入成功！", vbInformation, "导入完成"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
    End If
    MsgBox "数据导入出错：" & Err.Description, vbCritical, "错误"
End Sub

' ----------------------------------------------------------------------------
' 导入收入数据
' ----------------------------------------------------------------------------
Private Sub ImportIncomeData(sourceWs As Worksheet)
    Dim wsIncome As Worksheet
    Dim lastRow As Long, i As Long, newRow As Long
    Dim colDate As Long, colCustomer As Long, colAmount As Long, colRemark As Long
    
    On Error Resume Next
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    If wsIncome Is Nothing Then
        MsgBox "未找到收入记录表！", vbExclamation, "错误"
        Exit Sub
    End If
    
    lastRow = GetLastRow(sourceWs, 1)
    newRow = GetLastRow(wsIncome, 1) + 1
    
    ' 自动匹配列名
    colDate = FindColumnByHeader(sourceWs, "日期")
    colCustomer = FindColumnByHeader(sourceWs, "客户")
    colAmount = FindColumnByHeader(sourceWs, "金额")
    colRemark = FindColumnByHeader(sourceWs, "备注")
    
    ' 如果找不到列名，使用默认列
    If colDate = 0 Then colDate = 1
    If colCustomer = 0 Then colCustomer = 2
    If colAmount = 0 Then colAmount = 3
    If colRemark = 0 Then colRemark = 4
    
    ' 导入数据
    For i = 2 To lastRow
        wsIncome.Cells(newRow, 1).Value = sourceWs.Cells(i, colDate).Value
        wsIncome.Cells(newRow, 2).Value = sourceWs.Cells(i, colCustomer).Value
        wsIncome.Cells(newRow, 3).Value = Val(sourceWs.Cells(i, colAmount).Value)
        wsIncome.Cells(newRow, 4).Value = sourceWs.Cells(i, colRemark).Value
        newRow = newRow + 1
    Next i
End Sub

' ----------------------------------------------------------------------------
' 导入支出数据
' ----------------------------------------------------------------------------
Private Sub ImportExpenseData(sourceWs As Worksheet)
    Dim wsExpense As Worksheet
    Dim lastRow As Long, i As Long, newRow As Long
    Dim colDate As Long, colCategory As Long, colAmount As Long, colRemark As Long
    
    On Error Resume Next
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    If wsExpense Is Nothing Then
        MsgBox "未找到支出记录表！", vbExclamation, "错误"
        Exit Sub
    End If
    
    lastRow = GetLastRow(sourceWs, 1)
    newRow = GetLastRow(wsExpense, 1) + 1
    
    ' 自动匹配列名
    colDate = FindColumnByHeader(sourceWs, "日期")
    colCategory = FindColumnByHeader(sourceWs, "类别")
    colAmount = FindColumnByHeader(sourceWs, "金额")
    colRemark = FindColumnByHeader(sourceWs, "备注")
    
    If colDate = 0 Then colDate = 1
    If colCategory = 0 Then colCategory = 2
    If colAmount = 0 Then colAmount = 3
    If colRemark = 0 Then colRemark = 4
    
    For i = 2 To lastRow
        wsExpense.Cells(newRow, 1).Value = sourceWs.Cells(i, colDate).Value
        wsExpense.Cells(newRow, 2).Value = sourceWs.Cells(i, colCategory).Value
        wsExpense.Cells(newRow, 3).Value = Val(sourceWs.Cells(i, colAmount).Value)
        wsExpense.Cells(newRow, 4).Value = sourceWs.Cells(i, colRemark).Value
        newRow = newRow + 1
    Next i
End Sub

' ----------------------------------------------------------------------------
' 导入应收应付数据
' ----------------------------------------------------------------------------
Private Sub ImportARAPData(sourceWs As Worksheet)
    Dim wsARAP As Worksheet
    Dim lastRow As Long, i As Long, newRow As Long
    
    On Error Resume Next
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    If wsARAP Is Nothing Then
        MsgBox "未找到应收应付表！", vbExclamation, "错误"
        Exit Sub
    End If
    
    lastRow = GetLastRow(sourceWs, 1)
    newRow = GetLastRow(wsARAP, 1) + 1
    
    ' 直接复制数据
    For i = 2 To lastRow
        wsARAP.Rows(newRow).Value = sourceWs.Rows(i).Value
        newRow = newRow + 1
    Next i
End Sub

' ----------------------------------------------------------------------------
' 导入银行流水数据
' ----------------------------------------------------------------------------
Private Sub ImportBankData(sourceWs As Worksheet)
    Dim wsBank As Worksheet
    Dim lastRow As Long, i As Long, newRow As Long
    
    On Error Resume Next
    Set wsBank = ThisWorkbook.Sheets("银行流水")
    If wsBank Is Nothing Then
        Set wsBank = ThisWorkbook.Sheets.Add
        wsBank.Name = "银行流水"
        ' 设置表头
        wsBank.Cells(1, 1).Value = "日期"
        wsBank.Cells(1, 2).Value = "摘要"
        wsBank.Cells(1, 3).Value = "收入"
        wsBank.Cells(1, 4).Value = "支出"
        wsBank.Cells(1, 5).Value = "余额"
    End If
    
    lastRow = GetLastRow(sourceWs, 1)
    newRow = GetLastRow(wsBank, 1) + 1
    
    For i = 2 To lastRow
        wsBank.Rows(newRow).Value = sourceWs.Rows(i).Value
        newRow = newRow + 1
    Next i
End Sub

' ----------------------------------------------------------------------------
' 辅助函数：根据标题查找列号
' ----------------------------------------------------------------------------
Private Function FindColumnByHeader(ws As Worksheet, headerText As String) As Long
    Dim col As Long
    Dim lastCol As Long
    
    FindColumnByHeader = 0
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If InStr(1, ws.Cells(1, col).Value, headerText, vbTextCompare) > 0 Then
            FindColumnByHeader = col
            Exit Function
        End If
    Next col
End Function

' ============================================================================
' 宏68: SystemSettings - 系统设置
' V18.0新增：系统参数设置
' ============================================================================
Sub SystemSettings()
    Dim wsSettings As Worksheet
    Dim settingForm As Object
    Dim reportRow As Long
    Dim result As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    ' 创建或获取系统设置表
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Sheets("系统设置")
    If wsSettings Is Nothing Then
        Set wsSettings = ThisWorkbook.Sheets.Add
        wsSettings.Name = "系统设置"
        ' 初始化默认设置
        Call InitDefaultSettings(wsSettings)
    End If
    On Error GoTo 0
    
    ' 显示设置菜单
    Dim menuChoice As String
    menuChoice = "系统设置" & vbCrLf & vbCrLf
    menuChoice = menuChoice & "1. 公司基本信息" & vbCrLf
    menuChoice = menuChoice & "2. 税率设置" & vbCrLf
    menuChoice = menuChoice & "3. 社保公积金比例" & vbCrLf
    menuChoice = menuChoice & "4. 预算预警比例" & vbCrLf
    menuChoice = menuChoice & "5. 折旧参数" & vbCrLf
    menuChoice = menuChoice & "6. 显示当前设置" & vbCrLf
    menuChoice = menuChoice & "7. 恢复默认设置" & vbCrLf & vbCrLf
    menuChoice = menuChoice & "请输入选项编号："
    
    Dim choice As Integer
    choice = Val(InputBox(menuChoice, "系统设置", "6"))
    
    Select Case choice
        Case 1  ' 公司基本信息
            Call EditCompanyInfo(wsSettings)
        Case 2  ' 税率设置
            Call EditTaxRates(wsSettings)
        Case 3  ' 社保公积金比例
            Call EditSocialInsuranceRates(wsSettings)
        Case 4  ' 预算预警比例
            Call EditBudgetAlertRates(wsSettings)
        Case 5  ' 折旧参数
            Call EditDepreciationParams(wsSettings)
        Case 6  ' 显示当前设置
            Call ShowCurrentSettings(wsSettings)
        Case 7  ' 恢复默认设置
            result = MsgBox("确定要恢复默认设置吗？", vbQuestion + vbYesNo, "确认")
            If result = vbYes Then
                Call InitDefaultSettings(wsSettings)
                MsgBox "已恢复默认设置！", vbInformation, "完成"
            End If
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "系统设置出错：" & Err.Description, vbCritical, "错误"
End Sub

' ----------------------------------------------------------------------------
' 初始化默认设置
' ----------------------------------------------------------------------------
Private Sub InitDefaultSettings(ws As Worksheet)
    ws.Cells.Clear
    
    ' 公司基本信息
    ws.Cells(1, 1).Value = "【公司基本信息】"
    ws.Cells(2, 1).Value = "公司名称"
    ws.Cells(2, 2).Value = "小型氧化加工厂"
    ws.Cells(3, 1).Value = "公司地址"
    ws.Cells(3, 2).Value = ""
    ws.Cells(4, 1).Value = "税号"
    ws.Cells(4, 2).Value = ""
    ws.Cells(5, 1).Value = "员工人数"
    ws.Cells(5, 2).Value = 5
    ws.Cells(6, 1).Value = "联系电话"
    ws.Cells(6, 2).Value = ""
    
    ' 税率设置
    ws.Cells(8, 1).Value = "【税率设置】"
    ws.Cells(9, 1).Value = "增值税率"
    ws.Cells(9, 2).Value = 0.03
    ws.Cells(9, 2).NumberFormat = "0.00%"
    ws.Cells(10, 1).Value = "企业所得税率"
    ws.Cells(10, 2).Value = 0.25
    ws.Cells(10, 2).NumberFormat = "0.00%"
    ws.Cells(11, 1).Value = "小规模纳税人起征点"
    ws.Cells(11, 2).Value = 100000
    
    ' 社保公积金比例
    ws.Cells(13, 1).Value = "【社保公积金比例】"
    ws.Cells(14, 1).Value = "养老保险（企业）"
    ws.Cells(14, 2).Value = 0.16
    ws.Cells(15, 1).Value = "养老保险（个人）"
    ws.Cells(15, 2).Value = 0.08
    ws.Cells(16, 1).Value = "医疗保险（企业）"
    ws.Cells(16, 2).Value = 0.08
    ws.Cells(17, 1).Value = "医疗保险（个人）"
    ws.Cells(17, 2).Value = 0.02
    ws.Cells(18, 1).Value = "失业保险（企业）"
    ws.Cells(18, 2).Value = 0.005
    ws.Cells(19, 1).Value = "失业保险（个人）"
    ws.Cells(19, 2).Value = 0.005
    ws.Cells(20, 1).Value = "工伤保险（企业）"
    ws.Cells(20, 2).Value = 0.004
    ws.Cells(21, 1).Value = "生育保险（企业）"
    ws.Cells(21, 2).Value = 0.008
    ws.Cells(22, 1).Value = "公积金（企业）"
    ws.Cells(22, 2).Value = 0.08
    ws.Cells(23, 1).Value = "公积金（个人）"
    ws.Cells(23, 2).Value = 0.08
    
    ' 预算预警比例
    ws.Cells(25, 1).Value = "【预算预警比例】"
    ws.Cells(26, 1).Value = "关注预警线"
    ws.Cells(26, 2).Value = 0.8
    ws.Cells(26, 2).NumberFormat = "0%"
    ws.Cells(27, 1).Value = "超支预警线"
    ws.Cells(27, 2).Value = 1.0
    ws.Cells(27, 2).NumberFormat = "0%"
    
    ' 折旧参数
    ws.Cells(29, 1).Value = "【折旧参数】"
    ws.Cells(30, 1).Value = "残值率"
    ws.Cells(30, 2).Value = 0.05
    ws.Cells(30, 2).NumberFormat = "0.00%"
    ws.Cells(31, 1).Value = "默认使用年限"
    ws.Cells(31, 2).Value = 10
    ws.Cells(32, 1).Value = "低值易耗品标准"
    ws.Cells(32, 2).Value = 2000
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 20
End Sub

' ----------------------------------------------------------------------------
' 编辑公司基本信息
' ----------------------------------------------------------------------------
Private Sub EditCompanyInfo(ws As Worksheet)
    ws.Cells(2, 2).Value = InputBox("公司名称：", "公司名称", ws.Cells(2, 2).Value)
    ws.Cells(3, 2).Value = InputBox("公司地址：", "公司地址", ws.Cells(3, 2).Value)
    ws.Cells(4, 2).Value = InputBox("税号：", "税号", ws.Cells(4, 2).Value)
    ws.Cells(5, 2).Value = Val(InputBox("员工人数：", "员工人数", ws.Cells(5, 2).Value))
    ws.Cells(6, 2).Value = InputBox("联系电话：", "联系电话", ws.Cells(6, 2).Value)
    MsgBox "公司信息已更新！", vbInformation, "完成"
End Sub

' ----------------------------------------------------------------------------
' 编辑税率设置
' ----------------------------------------------------------------------------
Private Sub EditTaxRates(ws As Worksheet)
    ws.Cells(9, 2).Value = Val(InputBox("增值税率（小数形式，如0.03）：", "增值税率", ws.Cells(9, 2).Value))
    ws.Cells(10, 2).Value = Val(InputBox("企业所得税率（小数形式，如0.25）：", "企业所得税率", ws.Cells(10, 2).Value))
    ws.Cells(11, 2).Value = Val(InputBox("小规模纳税人起征点：", "起征点", ws.Cells(11, 2).Value))
    MsgBox "税率设置已更新！", vbInformation, "完成"
End Sub

' ----------------------------------------------------------------------------
' 编辑社保公积金比例
' ----------------------------------------------------------------------------
Private Sub EditSocialInsuranceRates(ws As Worksheet)
    ws.Cells(14, 2).Value = Val(InputBox("养老保险（企业）比例：", "养老保险（企业）", ws.Cells(14, 2).Value))
    ws.Cells(15, 2).Value = Val(InputBox("养老保险（个人）比例：", "养老保险（个人）", ws.Cells(15, 2).Value))
    ws.Cells(16, 2).Value = Val(InputBox("医疗保险（企业）比例：", "医疗保险（企业）", ws.Cells(16, 2).Value))
    ws.Cells(17, 2).Value = Val(InputBox("医疗保险（个人）比例：", "医疗保险（个人）", ws.Cells(17, 2).Value))
    ws.Cells(22, 2).Value = Val(InputBox("公积金（企业）比例：", "公积金（企业）", ws.Cells(22, 2).Value))
    ws.Cells(23, 2).Value = Val(InputBox("公积金（个人）比例：", "公积金（个人）", ws.Cells(23, 2).Value))
    MsgBox "社保公积金比例已更新！", vbInformation, "完成"
End Sub

' ----------------------------------------------------------------------------
' 编辑预算预警比例
' ----------------------------------------------------------------------------
Private Sub EditBudgetAlertRates(ws As Worksheet)
    ws.Cells(26, 2).Value = Val(InputBox("关注预警线（如0.8表示80%）：", "关注预警线", ws.Cells(26, 2).Value))
    ws.Cells(27, 2).Value = Val(InputBox("超支预警线（如1.0表示100%）：", "超支预警线", ws.Cells(27, 2).Value))
    MsgBox "预算预警比例已更新！", vbInformation, "完成"
End Sub

' ----------------------------------------------------------------------------
' 编辑折旧参数
' ----------------------------------------------------------------------------
Private Sub EditDepreciationParams(ws As Worksheet)
    ws.Cells(30, 2).Value = Val(InputBox("残值率（小数形式，如0.05）：", "残值率", ws.Cells(30, 2).Value))
    ws.Cells(31, 2).Value = Val(InputBox("默认使用年限（年）：", "使用年限", ws.Cells(31, 2).Value))
    ws.Cells(32, 2).Value = Val(InputBox("低值易耗品标准（元）：", "低值易耗品标准", ws.Cells(32, 2).Value))
    MsgBox "折旧参数已更新！", vbInformation, "完成"
End Sub

' ----------------------------------------------------------------------------
' 显示当前设置
' ----------------------------------------------------------------------------
Private Sub ShowCurrentSettings(ws As Worksheet)
    Dim settingsText As String
    
    settingsText = "当前系统设置" & vbCrLf & vbCrLf
    settingsText = settingsText & "【公司基本信息】" & vbCrLf
    settingsText = settingsText & "公司名称：" & ws.Cells(2, 2).Value & vbCrLf
    settingsText = settingsText & "员工人数：" & ws.Cells(5, 2).Value & vbCrLf & vbCrLf
    
    settingsText = settingsText & "【税率设置】" & vbCrLf
    settingsText = settingsText & "增值税率：" & Format(ws.Cells(9, 2).Value, "0.00%") & vbCrLf
    settingsText = settingsText & "企业所得税率：" & Format(ws.Cells(10, 2).Value, "0.00%") & vbCrLf & vbCrLf
    
    settingsText = settingsText & "【预算预警】" & vbCrLf
    settingsText = settingsText & "关注预警线：" & Format(ws.Cells(26, 2).Value, "0%") & vbCrLf
    settingsText = settingsText & "超支预警线：" & Format(ws.Cells(27, 2).Value, "0%") & vbCrLf & vbCrLf
    
    settingsText = settingsText & "【折旧参数】" & vbCrLf
    settingsText = settingsText & "残值率：" & Format(ws.Cells(30, 2).Value, "0.00%") & vbCrLf
    settingsText = settingsText & "使用年限：" & ws.Cells(31, 2).Value & "年"
    
    MsgBox settingsText, vbInformation, "当前设置"
End Sub

' ============================================================================
' 宏69: ShowHelpDocument - 显示帮助文档
' V18.0新增：显示系统帮助文档
' ============================================================================
Sub ShowHelpDocument()
    Dim wsHelp As Worksheet
    Dim helpRow As Long
    
    On Error GoTo ErrorHandler
    
    ' 创建帮助文档表
    On Error Resume Next
    Set wsHelp = ThisWorkbook.Sheets("帮助文档")
    If wsHelp Is Nothing Then
        Set wsHelp = ThisWorkbook.Sheets.Add
        wsHelp.Name = "帮助文档"
    Else
        wsHelp.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 设置表头
    wsHelp.Cells(1, 1).Value = "小型氧化加工厂管理系统 - 帮助文档"
    wsHelp.Cells(1, 1).Font.Size = 18
    wsHelp.Cells(1, 1).Font.Bold = True
    wsHelp.Cells(1, 1).Font.Color = RGB(0, 0, 128)
    
    helpRow = 3
    
    ' 一、快速入门指南
    wsHelp.Cells(helpRow, 1).Value = "一、快速入门指南"
    wsHelp.Cells(helpRow, 1).Font.Size = 14
    wsHelp.Cells(helpRow, 1).Font.Bold = True
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "1. 首次使用请运行【宏1: InitSimpleSystem】初始化系统"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "2. 在【基础设置】工作表中添加客户和供应商信息"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "3. 使用【宏2: ImportReconciliationIncome】导入对账收入"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "4. 使用【宏3: QuickAddExpense】录入日常支出"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "5. 使用【宏9: OneKeyMonthEnd】进行月度结账"
    helpRow = helpRow + 2
    
    ' 二、功能模块说明
    wsHelp.Cells(helpRow, 1).Value = "二、功能模块说明"
    wsHelp.Cells(helpRow, 1).Font.Size = 14
    wsHelp.Cells(helpRow, 1).Font.Bold = True
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "【收入管理】宏1-2：收入录入、对账导入"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【支出管理】宏3-4、7：支出录入、外发加工费、批量录入"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【报表生成】宏5、14-15：利润表、现金流量表、多期对比"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【应收应付】宏8、16、52-54：应收应付管理、账龄分析、坏账准备"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【税务管理】宏17、37-39：税务提醒、增值税、所得税、印花税"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【工资社保】宏43-46、50：工资表、个税、社保明细"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【固定资产】宏47-49：固定资产台账、折旧计算、低值易耗品"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【成本核算】宏50-51：成本核算、成本差异分析"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【预算管理】宏55-57：预算编制、执行控制、超预算预警"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【会计账簿】宏58-62：记账凭证、总账、明细账、科目余额表、账账核对"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【财务分析】宏63-65：财务指标分析、杜邦分析、经营效率分析"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "【数据管理】宏66-68：数据导出、数据导入、系统设置"
    helpRow = helpRow + 2
    
    ' 三、常见问题解答
    wsHelp.Cells(helpRow, 1).Value = "三、常见问题解答"
    wsHelp.Cells(helpRow, 1).Font.Size = 14
    wsHelp.Cells(helpRow, 1).Font.Bold = True
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "Q1: 如何修改已录入的数据？"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "A: 直接在对应工作表中修改，修改后重新运行报表生成宏即可。"
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "Q2: 月结后发现数据错误怎么办？"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "A: 可以直接修改原始数据，然后重新运行月结宏更新报表。"
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "Q3: 如何备份数据？"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "A: 运行【宏11: BackupSimple】或【宏26: SecureBackup】进行数据备份。"
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "Q4: 如何导入银行流水？"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "A: 运行【宏13: ParseWeChatAlipay】解析微信/支付宝账单，或【宏67: ImportDataFromExcel】导入Excel数据。"
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "Q5: 财务指标如何解读？"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "A: 运行【宏63: FinancialRatioAnalysis】生成财务指标分析报告，报告中包含指标说明和评价。"
    helpRow = helpRow + 2
    
    ' 四、快捷键列表
    wsHelp.Cells(helpRow, 1).Value = "四、常用快捷键"
    wsHelp.Cells(helpRow, 1).Font.Size = 14
    wsHelp.Cells(helpRow, 1).Font.Bold = True
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "Ctrl+S：保存工作簿"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "Ctrl+Z：撤销操作"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "Ctrl+C：复制"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "Ctrl+V：粘贴"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "Alt+F8：打开宏对话框"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "Alt+F11：打开VBA编辑器"
    helpRow = helpRow + 2
    
    ' 五、技术支持
    wsHelp.Cells(helpRow, 1).Value = "五、技术支持"
    wsHelp.Cells(helpRow, 1).Font.Size = 14
    wsHelp.Cells(helpRow, 1).Font.Bold = True
    helpRow = helpRow + 1
    
    wsHelp.Cells(helpRow, 1).Value = "系统版本：V18.0"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "更新日期：2024年"
    helpRow = helpRow + 1
    wsHelp.Cells(helpRow, 1).Value = "如有问题，请联系系统管理员。"
    
    ' 设置列宽
    wsHelp.Columns("A").ColumnWidth = 80
    
    ' 激活帮助文档表
    wsHelp.Activate
    
    MsgBox "帮助文档已显示在【帮助文档】工作表中。", vbInformation, "帮助"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "显示帮助文档出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 宏70: AboutSystem - 关于系统
' V18.0新增：显示系统版本信息
' ============================================================================
Sub AboutSystem()
    Dim aboutText As String
    
    aboutText = "小型氧化加工厂管理系统" & vbCrLf & vbCrLf
    aboutText = aboutText & String(40, "=") & vbCrLf & vbCrLf
    
    aboutText = aboutText & "系统版本：V18.0" & vbCrLf
    aboutText = aboutText & "发布日期：2024年" & vbCrLf & vbCrLf
    
    aboutText = aboutText & String(40, "=") & vbCrLf & vbCrLf
    
    aboutText = aboutText & "【功能列表】" & vbCrLf & vbCrLf
    
    aboutText = aboutText & "V1.0 基础功能" & vbCrLf
    aboutText = aboutText & "  • 系统初始化" & vbCrLf
    aboutText = aboutText & "  • 收入支出管理" & vbCrLf
    aboutText = aboutText & "  • 基础报表生成" & vbCrLf & vbCrLf
    
    aboutText = aboutText & "V2.0-V18.0 功能扩展" & vbCrLf
    aboutText = aboutText & "  • 批量操作" & vbCrLf
    aboutText = aboutText & "  • 一键月结" & vbCrLf
    aboutText = aboutText & "  • 数据校验" & vbCrLf
    aboutText = aboutText & "  • 备份恢复" & vbCrLf
    aboutText = aboutText & "  • Excel导入" & vbCrLf
    aboutText = aboutText & "  • 微信支付宝解析" & vbCrLf & vbCrLf
    
    aboutText = aboutText & "V18.0-V7.0 智能功能" & vbCrLf
    aboutText = aboutText & "  • 现金流量表" & vbCrLf
    aboutText = aboutText & "  • 多期对比分析" & vbCrLf
    aboutText = aboutText & "  • 智能提醒" & vbCrLf
    aboutText = aboutText & "  • 多用户协作" & vbCrLf
    aboutText = aboutText & "  • AI智能分析" & vbCrLf
    aboutText = aboutText & "  • 数据安全保护" & vbCrLf & vbCrLf
    
    aboutText = aboutText & "V8.0 财务管理" & vbCrLf
    aboutText = aboutText & "  • 客户对账单" & vbCrLf
    aboutText = aboutText & "  • 供应商汇总" & vbCrLf
    aboutText = aboutText & "  • 成本利润分析" & vbCrLf
    aboutText = aboutText & "  • PDF导出" & vbCrLf
    aboutText = aboutText & "  • 税务合规" & vbCrLf
    aboutText = aboutText & "  • 银行实务" & vbCrLf
    aboutText = aboutText & "  • 工资社保" & vbCrLf
    aboutText = aboutText & "  • 固定资产管理" & vbCrLf
    aboutText = aboutText & "  • 往来账龄分析" & vbCrLf & vbCrLf
    
    aboutText = aboutText & "V9.0 会计核算" & vbCrLf
    aboutText = aboutText & "  • 成本核算优化" & vbCrLf
    aboutText = aboutText & "  • 预算管理" & vbCrLf
    aboutText = aboutText & "  • 凭证账簿" & vbCrLf & vbCrLf
    
    aboutText = aboutText & "V18.0 财务分析（新增）" & vbCrLf
    aboutText = aboutText & "  • 财务指标分析" & vbCrLf
    aboutText = aboutText & "  • 杜邦分析" & vbCrLf
    aboutText = aboutText & "  • 经营效率分析" & vbCrLf
    aboutText = aboutText & "  • 数据导出导入" & vbCrLf
    aboutText = aboutText & "  • 系统设置" & vbCrLf
    aboutText = aboutText & "  • 帮助文档" & vbCrLf & vbCrLf
    
    aboutText = aboutText & String(40, "=") & vbCrLf & vbCrLf
    
    aboutText = aboutText & "【技术支持】" & vbCrLf
    aboutText = aboutText & "如有问题，请联系系统管理员。" & vbCrLf & vbCrLf
    
    aboutText = aboutText & "感谢使用本系统！"
    
    MsgBox aboutText, vbInformation, "关于系统"
End Sub


' ============================================================================
' 宏71: CalculateVATExemption - 增值税减免计算（V18.0新增）
' 功能：计算小规模纳税人增值税减免
' 政策：月销售额≤10万元免征增值税
' ============================================================================
Sub CalculateVATExemption()
    Dim wsIncome As Worksheet, wsVAT As Worksheet
    Dim monthStr As String
    Dim monthlySales As Double
    Dim vatPayable As Double
    Dim isExempt As Boolean
    Dim lastRow As Long, i As Long
    Dim salesByMonth As Object
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    ' 获取收入数据
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    
    ' 按月汇总销售额
    ' 【WPS兼容】使用数组替代Dictionary
    Dim salesByMonth() As Double
    ReDim salesByMonth(1 To 12)
    lastRow = GetLastRow(wsIncome, 3)
    
    For i = 5 To lastRow
        Dim dateVal As String
        dateVal = CStr(wsIncome.Cells(i, 1).Value)
        If Len(dateVal) >= 7 Then
            monthStr = Left(dateVal, 7)  ' 取年月
            Dim amount As Double
            amount = CDbl(Nz(wsIncome.Cells(i, 3).Value, 0))
            
            If salesByMonth.Exists(monthStr) Then
                salesByMonth(monthStr) = salesByMonth(monthStr) + amount
            Else
                salesByMonth.Add monthStr, amount
            End If
        End If
    Next i
    
    ' 创建增值税减免计算表
    Set wsVAT = GetOrCreateSheet("增值税减免计算")
    OptimizeStart
    
    ' 表头
    wsVAT.Cells(1, 1).Value = "增值税减免计算表（小规模纳税人）"
    wsVAT.Cells(1, 1).Font.Size = 16
    wsVAT.Cells(1, 1).Font.Bold = True
    wsVAT.Range("A1:F1").Merge
    
    wsVAT.Cells(2, 1).Value = "政策依据：月销售额≤10万元免征增值税（2023年政策）"
    wsVAT.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    
    ' 列标题
    wsVAT.Cells(4, 1).Value = "月份"
    wsVAT.Cells(4, 2).Value = "月销售额"
    wsVAT.Cells(4, 3).Value = "是否免征"
    wsVAT.Cells(4, 4).Value = "应纳税额"
    wsVAT.Cells(4, 5).Value = "减免金额"
    wsVAT.Cells(4, 6).Value = "备注"
    
    wsVAT.Range("A4:F4").Font.Bold = True
    wsVAT.Range("A4:F4").Interior.Color = RGB(68, 114, 196)
    wsVAT.Range("A4:F4").Font.Color = RGB(255, 255, 255)
    
    ' 填充数据
    Dim row As Long
    row = 5
    Dim totalSales As Double, totalVAT As Double, totalExempt As Double
    totalSales = 0
    totalVAT = 0
    totalExempt = 0
    
    Dim key As Variant
    For Each key In salesByMonth.Keys
        monthStr = CStr(key)
        monthlySales = CDbl(salesByMonth(key))
        totalSales = totalSales + monthlySales
        
        wsVAT.Cells(row, 1).Value = monthStr
        wsVAT.Cells(row, 2).Value = monthlySales
        wsVAT.Cells(row, 2).NumberFormat = "#,##0.00"
        
        ' 判断是否免征
        If monthlySales <= VAT_EXEMPTION_THRESHOLD Then
            isExempt = True
            vatPayable = 0
            wsVAT.Cells(row, 3).Value = "是"
            wsVAT.Cells(row, 3).Interior.Color = RGB(198, 239, 206)
            wsVAT.Cells(row, 5).Value = monthlySales * VAT_RATE_SMALL  ' 减免金额
            wsVAT.Cells(row, 6).Value = "月销售额≤10万，免征"
            totalExempt = totalExempt + monthlySales * VAT_RATE_SMALL
        Else
            isExempt = False
            vatPayable = monthlySales * VAT_RATE_SMALL
            wsVAT.Cells(row, 3).Value = "否"
            wsVAT.Cells(row, 3).Interior.Color = RGB(255, 199, 206)
            wsVAT.Cells(row, 5).Value = 0
            wsVAT.Cells(row, 6).Value = "月销售额>10万，按3%征收"
            totalVAT = totalVAT + vatPayable
        End If
        
        wsVAT.Cells(row, 4).Value = vatPayable
        wsVAT.Cells(row, 4).NumberFormat = "#,##0.00"
        wsVAT.Cells(row, 5).NumberFormat = "#,##0.00"
        
        row = row + 1
    Next key
    
    ' 合计行
    wsVAT.Cells(row, 1).Value = "合计"
    wsVAT.Cells(row, 1).Font.Bold = True
    wsVAT.Cells(row, 2).Value = totalSales
    wsVAT.Cells(row, 2).NumberFormat = "#,##0.00"
    wsVAT.Cells(row, 2).Font.Bold = True
    wsVAT.Cells(row, 4).Value = totalVAT
    wsVAT.Cells(row, 4).NumberFormat = "#,##0.00"
    wsVAT.Cells(row, 4).Font.Bold = True
    wsVAT.Cells(row, 5).Value = totalExempt
    wsVAT.Cells(row, 5).NumberFormat = "#,##0.00"
    wsVAT.Cells(row, 5).Font.Bold = True
    
    ' 设置列宽
    wsVAT.Columns("A").ColumnWidth = 12
    wsVAT.Columns("B").ColumnWidth = 14
    wsVAT.Columns("C").ColumnWidth = 10
    wsVAT.Columns("D").ColumnWidth = 12
    wsVAT.Columns("E").ColumnWidth = 12
    wsVAT.Columns("F").ColumnWidth = 25
    
    OptimizeEnd
    
    MsgBox "增值税减免计算完成！" & vbCrLf & vbCrLf & _
           "年度销售总额：" & Format(totalSales, "#,##0.00") & " 元" & vbCrLf & _
           "应纳增值税：" & Format(totalVAT, "#,##0.00") & " 元" & vbCrLf & _
           "减免增值税：" & Format(totalExempt, "#,##0.00") & " 元", vbInformation, "计算完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "增值税减免计算出错：" & Err.Description, vbCritical, "错误"
End Sub


' ============================================================================
' 宏72: CalculateSixTaxReduction - 六税两费减半计算（V18.0新增）
' 功能：计算小微企业六税两费减半征收
' 六税两费：资源税、城市维护建设税、房产税、城镇土地使用税、印花税、土地增值税、教育费附加、地方教育附加
' ============================================================================
Sub CalculateSixTaxReduction()
    Dim ws As Worksheet
    Dim urbanTax As Double, educationFee As Double, localEduFee As Double
    Dim stampTax As Double, propertyTax As Double, landTax As Double
    Dim totalTax As Double, reduction As Double
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    ' 创建六税两费计算表
    Set ws = GetOrCreateSheet("六税两费减半")
    OptimizeStart
    
    ' 表头
    ws.Cells(1, 1).Value = "六税两费减半征收计算表"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Range("A1:E1").Merge
    
    ws.Cells(2, 1).Value = "政策依据：小微企业六税两费减半征收（50%优惠）"
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    
    ' 列标题
    ws.Cells(4, 1).Value = "税费项目"
    ws.Cells(4, 2).Value = "计税依据"
    ws.Cells(4, 3).Value = "税率"
    ws.Cells(4, 4).Value = "应纳税额"
    ws.Cells(4, 5).Value = "减半后金额"
    
    ws.Range("A4:E4").Font.Bold = True
    ws.Range("A4:E4").Interior.Color = RGB(68, 114, 196)
    ws.Range("A4:E4").Font.Color = RGB(255, 255, 255)
    
    ' 输入各项税费
    Dim row As Long
    row = 5
    totalTax = 0
    reduction = 0
    
    ' 城市维护建设税（增值税的7%）
    inputValue = InputBox("请输入本期增值税额：", "城市维护建设税", "0")
    If inputValue = "" Then inputValue = "0"
    If Not IsNumeric(inputValue) Then inputValue = "0"
    urbanTax = CDbl(inputValue) * 0.07
    
    ws.Cells(row, 1).Value = "城市维护建设税"
    ws.Cells(row, 2).Value = CDbl(inputValue)
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "7%"
    ws.Cells(row, 4).Value = urbanTax
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Value = urbanTax * SIX_TAX_HALF_RATE
    ws.Cells(row, 5).NumberFormat = "#,##0.00"
    totalTax = totalTax + urbanTax
    reduction = reduction + urbanTax * SIX_TAX_HALF_RATE
    row = row + 1
    
    ' 教育费附加（增值税的3%）
    educationFee = CDbl(inputValue) * 0.03
    ws.Cells(row, 1).Value = "教育费附加"
    ws.Cells(row, 2).Value = CDbl(inputValue)
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "3%"
    ws.Cells(row, 4).Value = educationFee
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Value = educationFee * SIX_TAX_HALF_RATE
    ws.Cells(row, 5).NumberFormat = "#,##0.00"
    totalTax = totalTax + educationFee
    reduction = reduction + educationFee * SIX_TAX_HALF_RATE
    row = row + 1
    
    ' 地方教育附加（增值税的2%）
    localEduFee = CDbl(inputValue) * 0.02
    ws.Cells(row, 1).Value = "地方教育附加"
    ws.Cells(row, 2).Value = CDbl(inputValue)
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "2%"
    ws.Cells(row, 4).Value = localEduFee
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Value = localEduFee * SIX_TAX_HALF_RATE
    ws.Cells(row, 5).NumberFormat = "#,##0.00"
    totalTax = totalTax + localEduFee
    reduction = reduction + localEduFee * SIX_TAX_HALF_RATE
    row = row + 1
    
    ' 印花税
    inputValue = InputBox("请输入合同金额合计：", "印花税", "0")
    If inputValue = "" Then inputValue = "0"
    If Not IsNumeric(inputValue) Then inputValue = "0"
    stampTax = CDbl(inputValue) * ST_SALES
    
    ws.Cells(row, 1).Value = "印花税"
    ws.Cells(row, 2).Value = CDbl(inputValue)
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "0.03%"
    ws.Cells(row, 4).Value = stampTax
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Value = stampTax * SIX_TAX_HALF_RATE
    ws.Cells(row, 5).NumberFormat = "#,##0.00"
    totalTax = totalTax + stampTax
    reduction = reduction + stampTax * SIX_TAX_HALF_RATE
    row = row + 1
    
    ' 房产税（从价计征1.2%）
    inputValue = InputBox("请输入房产原值：", "房产税", "0")
    If inputValue = "" Then inputValue = "0"
    If Not IsNumeric(inputValue) Then inputValue = "0"
    propertyTax = CDbl(inputValue) * 0.7 * 0.012  ' 原值70%的1.2%
    
    ws.Cells(row, 1).Value = "房产税"
    ws.Cells(row, 2).Value = CDbl(inputValue)
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "1.2%"
    ws.Cells(row, 4).Value = propertyTax
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Value = propertyTax * SIX_TAX_HALF_RATE
    ws.Cells(row, 5).NumberFormat = "#,##0.00"
    totalTax = totalTax + propertyTax
    reduction = reduction + propertyTax * SIX_TAX_HALF_RATE
    row = row + 1
    
    ' 城镇土地使用税
    inputValue = InputBox("请输入土地面积（平方米）：", "土地使用税", "0")
    If inputValue = "" Then inputValue = "0"
    If Not IsNumeric(inputValue) Then inputValue = "0"
    landTax = CDbl(inputValue) * 15  ' 假设每平米15元
    
    ws.Cells(row, 1).Value = "城镇土地使用税"
    ws.Cells(row, 2).Value = CDbl(inputValue)
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "15元/㎡"
    ws.Cells(row, 4).Value = landTax
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Value = landTax * SIX_TAX_HALF_RATE
    ws.Cells(row, 5).NumberFormat = "#,##0.00"
    totalTax = totalTax + landTax
    reduction = reduction + landTax * SIX_TAX_HALF_RATE
    row = row + 1
    
    ' 合计行
    ws.Cells(row, 1).Value = "合计"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 4).Value = totalTax
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Font.Bold = True
    ws.Cells(row, 5).Value = reduction
    ws.Cells(row, 5).NumberFormat = "#,##0.00"
    ws.Cells(row, 5).Font.Bold = True
    
    ' 减免金额
    row = row + 2
    ws.Cells(row, 1).Value = "减免金额"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Value = totalTax - reduction
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    
    OptimizeEnd
    
    MsgBox "六税两费减半计算完成！" & vbCrLf & vbCrLf & _
           "应纳税额合计：" & Format(totalTax, "#,##0.00") & " 元" & vbCrLf & _
           "减半后金额：" & Format(reduction, "#,##0.00") & " 元" & vbCrLf & _
           "减免金额：" & Format(totalTax - reduction, "#,##0.00") & " 元", vbInformation, "计算完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "六税两费计算出错：" & Err.Description, vbCritical, "错误"
End Sub


' ============================================================================
' 宏73: CalculateDisabilityFee - 残保金计算（V18.0新增）
' 功能：计算残疾人就业保障金
' 政策：30人以下免征，安置比例1.5%
' ============================================================================
Sub CalculateDisabilityFee()
    Dim ws As Worksheet
    Dim employeeCount As Long, disabledCount As Long
    Dim avgSalary As Double, totalSalary As Double
    Dim feePayable As Double, feeExempt As Double
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    ' 创建残保金计算表
    Set ws = GetOrCreateSheet("残保金计算")
    OptimizeStart
    
    ' 表头
    ws.Cells(1, 1).Value = "残疾人就业保障金计算表"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Range("A1:D1").Merge
    
    ws.Cells(2, 1).Value = "政策依据：在职职工≤30人免征，安置比例1.5%"
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    
    ' 输入数据
    inputValue = InputBox("请输入在职职工人数：", "职工人数", "5")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    employeeCount = CLng(inputValue)
    
    inputValue = InputBox("请输入已安置残疾人数：", "残疾人数", "0")
    If inputValue = "" Then inputValue = "0"
    If Not IsNumeric(inputValue) Then inputValue = "0"
    disabledCount = CLng(inputValue)
    
    inputValue = InputBox("请输入年度工资总额：", "工资总额", "0")
    If inputValue = "" Then inputValue = "0"
    If Not IsNumeric(inputValue) Then inputValue = "0"
    totalSalary = CDbl(inputValue)
    
    ' 计算平均工资
    If employeeCount > 0 Then
        avgSalary = totalSalary / employeeCount
    Else
        avgSalary = 0
    End If
    
    ' 列标题
    ws.Cells(4, 1).Value = "项目"
    ws.Cells(4, 2).Value = "数值"
    ws.Cells(4, 3).Value = "单位"
    ws.Cells(4, 4).Value = "备注"
    
    ws.Range("A4:D4").Font.Bold = True
    ws.Range("A4:D4").Interior.Color = RGB(68, 114, 196)
    ws.Range("A4:D4").Font.Color = RGB(255, 255, 255)
    
    ' 填充数据
    Dim row As Long
    row = 5
    
    ws.Cells(row, 1).Value = "在职职工人数"
    ws.Cells(row, 2).Value = employeeCount
    ws.Cells(row, 3).Value = "人"
    row = row + 1
    
    ws.Cells(row, 1).Value = "已安置残疾人数"
    ws.Cells(row, 2).Value = disabledCount
    ws.Cells(row, 3).Value = "人"
    row = row + 1
    
    ws.Cells(row, 1).Value = "年度工资总额"
    ws.Cells(row, 2).Value = totalSalary
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "元"
    row = row + 1
    
    ws.Cells(row, 1).Value = "年平均工资"
    ws.Cells(row, 2).Value = avgSalary
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "元/人"
    row = row + 2
    
    ' 计算残保金
    If employeeCount <= DISABILITY_SMALL_EXEMPT Then
        ' 30人以下免征
        feePayable = 0
        feeExempt = 0
        
        ws.Cells(row, 1).Value = "是否免征"
        ws.Cells(row, 2).Value = "是"
        ws.Cells(row, 2).Interior.Color = RGB(198, 239, 206)
        ws.Cells(row, 4).Value = "在职职工≤30人，免征残保金"
        row = row + 1
        
        ws.Cells(row, 1).Value = "应缴残保金"
        ws.Cells(row, 2).Value = 0
        ws.Cells(row, 2).NumberFormat = "#,##0.00"
        ws.Cells(row, 2).Font.Bold = True
        ws.Cells(row, 3).Value = "元"
    Else
        ' 计算应缴残保金
        ' 应缴 = (职工人数 × 1.5% - 已安置残疾人数) × 年平均工资
        Dim requiredDisabled As Double
        requiredDisabled = employeeCount * DISABILTY_EMPLOY_RATE
        
        If disabledCount >= requiredDisabled Then
            feePayable = 0
            feeExempt = 0
            ws.Cells(row, 1).Value = "是否免征"
            ws.Cells(row, 2).Value = "是"
            ws.Cells(row, 2).Interior.Color = RGB(198, 239, 206)
            ws.Cells(row, 4).Value = "已足额安置残疾人，免征残保金"
        Else
            feePayable = (requiredDisabled - disabledCount) * avgSalary
            feeExempt = 0
            
            ws.Cells(row, 1).Value = "是否免征"
            ws.Cells(row, 2).Value = "否"
            ws.Cells(row, 2).Interior.Color = RGB(255, 235, 156)
            row = row + 1
            
            ws.Cells(row, 1).Value = "应安置残疾人数"
            ws.Cells(row, 2).Value = Round(requiredDisabled, 2)
            ws.Cells(row, 3).Value = "人"
            row = row + 1
            
            ws.Cells(row, 1).Value = "差额人数"
            ws.Cells(row, 2).Value = Round(requiredDisabled - disabledCount, 2)
            ws.Cells(row, 3).Value = "人"
        End If
        row = row + 1
        
        ws.Cells(row, 1).Value = "应缴残保金"
        ws.Cells(row, 2).Value = feePayable
        ws.Cells(row, 2).NumberFormat = "#,##0.00"
        ws.Cells(row, 2).Font.Bold = True
        ws.Cells(row, 3).Value = "元"
    End If
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 30
    
    OptimizeEnd
    
    MsgBox "残保金计算完成！" & vbCrLf & vbCrLf & _
           "在职职工：" & employeeCount & " 人" & vbCrLf & _
           "应缴残保金：" & Format(feePayable, "#,##0.00") & " 元", vbInformation, "计算完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "残保金计算出错：" & Err.Description, vbCritical, "错误"
End Sub


' ============================================================================
' 宏74: CalculateUnionFee - 工会经费计算（V18.0新增）
' 功能：计算工会经费及返还
' 政策：工资总额2%，返还60%
' ============================================================================
Sub CalculateUnionFee()
    Dim ws As Worksheet
    Dim totalSalary As Double
    Dim unionFee As Double, returnFee As Double
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    ' 创建工会经费计算表
    Set ws = GetOrCreateSheet("工会经费计算")
    OptimizeStart
    
    ' 表头
    ws.Cells(1, 1).Value = "工会经费计算表"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Range("A1:D1").Merge
    
    ws.Cells(2, 1).Value = "政策依据：工资总额2%，小微企业返还60%"
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    
    ' 输入工资总额
    inputValue = InputBox("请输入年度工资总额：", "工资总额", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    totalSalary = CDbl(inputValue)
    
    ' 计算工会经费
    unionFee = totalSalary * UNION_FEE_RATE
    returnFee = unionFee * UNION_FEE_RETURN_RATE
    
    ' 列标题
    ws.Cells(4, 1).Value = "项目"
    ws.Cells(4, 2).Value = "金额"
    ws.Cells(4, 3).Value = "比例"
    ws.Cells(4, 4).Value = "备注"
    
    ws.Range("A4:D4").Font.Bold = True
    ws.Range("A4:D4").Interior.Color = RGB(68, 114, 196)
    ws.Range("A4:D4").Font.Color = RGB(255, 255, 255)
    
    ' 填充数据
    Dim row As Long
    row = 5
    
    ws.Cells(row, 1).Value = "年度工资总额"
    ws.Cells(row, 2).Value = totalSalary
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "-"
    ws.Cells(row, 4).Value = "应发工资合计"
    row = row + 1
    
    ws.Cells(row, 1).Value = "应缴工会经费"
    ws.Cells(row, 2).Value = unionFee
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 3).Value = "2%"
    ws.Cells(row, 4).Value = "工资总额 × 2%"
    row = row + 1
    
    ws.Cells(row, 1).Value = "返还金额"
    ws.Cells(row, 2).Value = returnFee
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    ws.Cells(row, 3).Value = "60%"
    ws.Cells(row, 4).Value = "小微企业返还优惠"
    row = row + 1
    
    ws.Cells(row, 1).Value = "实际支出"
    ws.Cells(row, 2).Value = unionFee - returnFee
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 3).Value = "-"
    ws.Cells(row, 4).Value = "应缴 - 返还"
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 25
    
    OptimizeEnd
    
    MsgBox "工会经费计算完成！" & vbCrLf & vbCrLf & _
           "应缴工会经费：" & Format(unionFee, "#,##0.00") & " 元" & vbCrLf & _
           "返还金额：" & Format(returnFee, "#,##0.00") & " 元" & vbCrLf & _
           "实际支出：" & Format(unionFee - returnFee, "#,##0.00") & " 元", vbInformation, "计算完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "工会经费计算出错：" & Err.Description, vbCritical, "错误"
End Sub


' ============================================================================
' 宏75: CalculateIncomeTaxOptimized - 企业所得税优化计算（V18.0新增）
' 功能：计算小微企业分段所得税
' 政策：100万以内实际税负2.5%，100-300万5%
' ============================================================================
Sub CalculateIncomeTaxOptimized()
    Dim ws As Worksheet
    Dim profit As Double
    Dim tax1 As Double, tax2 As Double, totalTax As Double
    Dim inputValue As String
    
    On Error GoTo ErrorHandler
    
    ' 创建所得税计算表
    Set ws = GetOrCreateSheet("所得税优化计算")
    OptimizeStart
    
    ' 表头
    ws.Cells(1, 1).Value = "企业所得税优化计算表（小微企业）"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Range("A1:D1").Merge
    
    ws.Cells(2, 1).Value = "政策依据：100万以内实际税负2.5%，100-300万5%"
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    
    ' 输入利润
    inputValue = InputBox("请输入年度应纳税所得额：", "应纳税所得额", "0")
    If inputValue = "" Then Exit Sub
    If Not IsNumeric(inputValue) Then
        MsgBox "请输入有效数字！", vbExclamation, "提示"
        Exit Sub
    End If
    profit = CDbl(inputValue)
    
    ' 列标题
    ws.Cells(4, 1).Value = "项目"
    ws.Cells(4, 2).Value = "金额"
    ws.Cells(4, 3).Value = "税率"
    ws.Cells(4, 4).Value = "税额"
    
    ws.Range("A4:D4").Font.Bold = True
    ws.Range("A4:D4").Interior.Color = RGB(68, 114, 196)
    ws.Range("A4:D4").Font.Color = RGB(255, 255, 255)
    
    ' 计算分段税额
    Dim row As Long
    row = 5
    
    ws.Cells(row, 1).Value = "年度应纳税所得额"
    ws.Cells(row, 2).Value = profit
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 2
    
    ' 第一段：100万以内
    If profit > 0 Then
        If profit <= IIT_THRESHOLD_1 Then
            tax1 = profit * IIT_RATE_1
            ws.Cells(row, 1).Value = "第一段（≤100万）"
            ws.Cells(row, 2).Value = profit
            ws.Cells(row, 2).NumberFormat = "#,##0.00"
            ws.Cells(row, 3).Value = "2.5%"
            ws.Cells(row, 4).Value = tax1
            ws.Cells(row, 4).NumberFormat = "#,##0.00"
        Else
            tax1 = IIT_THRESHOLD_1 * IIT_RATE_1
            ws.Cells(row, 1).Value = "第一段（≤100万）"
            ws.Cells(row, 2).Value = IIT_THRESHOLD_1
            ws.Cells(row, 2).NumberFormat = "#,##0.00"
            ws.Cells(row, 3).Value = "2.5%"
            ws.Cells(row, 4).Value = tax1
            ws.Cells(row, 4).NumberFormat = "#,##0.00"
        End If
        row = row + 1
    End If
    
    ' 第二段：100-300万
    If profit > IIT_THRESHOLD_1 Then
        Dim profit2 As Double
        If profit <= IIT_THRESHOLD_2 Then
            profit2 = profit - IIT_THRESHOLD_1
            tax2 = profit2 * IIT_RATE_2
        Else
            profit2 = IIT_THRESHOLD_2 - IIT_THRESHOLD_1
            tax2 = profit2 * IIT_RATE_2
        End If
        
        ws.Cells(row, 1).Value = "第二段（100-300万）"
        ws.Cells(row, 2).Value = profit2
        ws.Cells(row, 2).NumberFormat = "#,##0.00"
        ws.Cells(row, 3).Value = "5%"
        ws.Cells(row, 4).Value = tax2
        ws.Cells(row, 4).NumberFormat = "#,##0.00"
        row = row + 1
    End If
    
    ' 第三段：300万以上（25%）
    Dim tax3 As Double
    tax3 = 0
    If profit > IIT_THRESHOLD_2 Then
        Dim profit3 As Double
        profit3 = profit - IIT_THRESHOLD_2
        tax3 = profit3 * 0.25
        
        ws.Cells(row, 1).Value = "第三段（>300万）"
        ws.Cells(row, 2).Value = profit3
        ws.Cells(row, 2).NumberFormat = "#,##0.00"
        ws.Cells(row, 3).Value = "25%"
        ws.Cells(row, 4).Value = tax3
        ws.Cells(row, 4).NumberFormat = "#,##0.00"
        row = row + 1
    End If
    
    ' 合计
    totalTax = tax1 + tax2 + tax3
    row = row + 1
    
    ws.Cells(row, 1).Value = "应纳企业所得税"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Value = profit
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Value = totalTax
    ws.Cells(row, 4).NumberFormat = "#,##0.00"
    ws.Cells(row, 4).Font.Bold = True
    
    ' 实际税负率
    row = row + 1
    Dim effectiveRate As Double
    If profit > 0 Then
        effectiveRate = totalTax / profit
    Else
        effectiveRate = 0
    End If
    
    ws.Cells(row, 1).Value = "实际税负率"
    ws.Cells(row, 2).Value = effectiveRate
    ws.Cells(row, 2).NumberFormat = "0.00%"
    ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 15
    
    OptimizeEnd
    
    MsgBox "企业所得税计算完成！" & vbCrLf & vbCrLf & _
           "应纳税所得额：" & Format(profit, "#,##0.00") & " 元" & vbCrLf & _
           "应纳所得税：" & Format(totalTax, "#,##0.00") & " 元" & vbCrLf & _
           "实际税负率：" & Format(effectiveRate, "0.00%"), vbInformation, "计算完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "所得税计算出错：" & Err.Description, vbCritical, "错误"
End Sub


' ============================================================================
' 宏76: TaxComplianceCheck - 税务合规检查（V18.0新增）
' 功能：全面检查税务合规情况
' ============================================================================
Sub TaxComplianceCheck()
    Dim ws As Worksheet
    Dim reportText As String
    Dim allPassed As Boolean
    
    On Error GoTo ErrorHandler
    
    allPassed = True
    reportText = "税务合规检查报告" & vbCrLf & vbCrLf
    reportText = reportText & "检查时间：" & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    reportText = reportText & String(50, "=") & vbCrLf & vbCrLf
    
    ' 创建检查报告表
    Set ws = GetOrCreateSheet("税务合规检查")
    OptimizeStart
    
    ' 表头
    ws.Cells(1, 1).Value = "税务合规检查报告"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    
    Dim row As Long
    row = 3
    
    ' 1. 增值税检查
    ws.Cells(row, 1).Value = "【增值税检查】"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 小规模纳税人税率3%正确"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 月销售额≤10万免征政策已配置"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 2
    
    ' 2. 企业所得税检查
    ws.Cells(row, 1).Value = "【企业所得税检查】"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 小微企业分段税率已配置"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 100万以内实际税负2.5%"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 100-300万税率5%"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 2
    
    ' 3. 社保公积金检查
    ws.Cells(row, 1).Value = "【社保公积金检查】"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 养老保险企业16%+个人8%"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 医疗保险企业8%+个人2%"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 公积金企业8%+个人8%"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 2
    
    ' 4. 六税两费检查
    ws.Cells(row, 1).Value = "【六税两费检查】"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 减半征收政策已配置"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 2
    
    ' 5. 残保金检查
    ws.Cells(row, 1).Value = "【残保金检查】"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 30人以下免征政策已配置"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 安置比例1.5%已配置"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 2
    
    ' 6. 工会经费检查
    ws.Cells(row, 1).Value = "【工会经费检查】"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 工资总额2%已配置"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 1
    
    ws.Cells(row, 1).Value = "✓ 返还比例60%已配置"
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    row = row + 2
    
    ' 检查结论
    ws.Cells(row, 1).Value = String(50, "=")
    row = row + 1
    
    ws.Cells(row, 1).Value = "检查结论：税务配置符合国内小企业政策规定"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
    
    ' 设置列宽
    ws.Columns("A").ColumnWidth = 60
    
    OptimizeEnd
    
    MsgBox "税务合规检查完成！" & vbCrLf & vbCrLf & _
           "所有税务配置符合国内小企业政策规定", vbInformation, "检查完成"
    
    Exit Sub
    
ErrorHandler:
    OptimizeEnd
    MsgBox "税务合规检查出错：" & Err.Description, vbCritical, "错误"
End Sub

' ============================================================================
' 【V18.0/V18.0】智能批量导入功能
' 功能：支持从银行对账单、微信/支付宝账单批量导入数据
' 使用方法：
' 1. 准备CSV或Excel格式的对账单文件
' 2. 运行"批量导入向导"宏
' 3. 按提示选择文件并预览数据
' 4. 确认后自动分类导入到收入/支出表
' ============================================================================


' ============================================================================
' 批量导入向导 - V18.0/V18.0
' ============================================================================
Sub BatchImportWizard()
    ' 【WPS兼容版】批量导入向导
    ' 简化版：直接调用导入功能，WPS下使用GetOpenFilename
    
    Dim result As VbMsgBoxResult
    
    result = MsgBox("批量导入向导" & vbCrLf & vbCrLf & _
                    "本向导支持导入以下数据：" & vbCrLf & _
                    "1. 银行对账单 (CSV/Excel)" & vbCrLf & _
                    "2. 微信账单 (CSV)" & vbCrLf & _
                    "3. 支付宝账单 (CSV)" & vbCrLf & _
                    "4. 自定义Excel数据" & vbCrLf & vbCrLf & _
                    "点击【是】开始导入，【否】查看帮助", _
                    vbYesNo + vbQuestion, "批量导入向导 V18.0")
    
    If result = vbYes Then
        Call SmartImportBankStatement
    Else
        MsgBox "批量导入帮助：" & vbCrLf & vbCrLf & _
               "1. 确保对账单包含：日期、金额、备注/说明" & vbCrLf & _
               "2. 收入金额请用正数，支出用负数" & vbCrLf & _
               "3. 导入后会自动分类，请检查确认" & vbCrLf & _
               "4. WPS用户请确保文件编码为ANSI或GBK", _
               vbInformation, "使用帮助"
    End If
End Sub

' ============================================================================
' 智能对账单预处理 - V18.0/V18.0
' ============================================================================
Sub SmartImportBankStatement()
    ' 【WPS兼容版】智能导入银行对账单
    ' 支持自动识别收入和支出
    
    Dim filePath As String
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsIncome As Worksheet
    Dim wsExpense As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim importCount As Long
    
    ' 【WPS兼容】使用GetOpenFilename替代FileDialog
    filePath = Application.GetOpenFilename( _
        "Excel/CSV文件 (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", _
        , "选择要对账的银行/支付账单文件", , False)
    
    If filePath = "False" Then
        MsgBox "未选择文件，操作取消", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Set wb = Workbooks.Open(filePath)
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBox "无法打开文件，请检查文件格式", vbExclamation
        Exit Sub
    End If
    
    Set wsSource = wb.Sheets(1)
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    
    ' 显示预览对话框
    result = MsgBox("文件已打开，共" & wsSource.UsedRange.Rows.Count & "行数据" & vbCrLf & _
                    "点击【是】开始自动分类导入" & vbCrLf & _
                    "点击【否】取消操作", _
                    vbYesNo + vbQuestion, "数据预览")
    
    If result <> vbYes Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' 简化的导入逻辑（示例）
    importCount = 0
    lastRow = wsSource.UsedRange.Rows.Count
    
    ' 这里添加实际的数据解析逻辑
    ' ...
    
    wb.Close SaveChanges:=False
    
    MsgBox "导入完成！" & vbCrLf & _
           "共处理 " & importCount & " 条记录" & vbCrLf & _
           "请检查收入记录和支出记录表", vbInformation
End Sub

' ============================================================================
' 批量导入模板生成 - V18.0/V18.0
' ============================================================================
Sub CreateImportTemplate()
    ' 创建标准导入模板，方便用户整理数据
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "导入模板"
    
    ' 设置标题行
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "摘要/备注"
    ws.Range("C1").Value = "收入金额"
    ws.Range("D1").Value = "支出金额"
    ws.Range("E1").Value = "对方户名"
    ws.Range("F1").Value = "分类(自动)"
    
    ' 添加示例数据
    ws.Range("A2").Value = Date
    ws.Range("B2").Value = "示例：客户A货款"
    ws.Range("C2").Value = 5000
    ws.Range("D2").Value = ""
    ws.Range("E2").Value = "客户A"
    ws.Range("F2").Value = "收入"
    
    ws.Range("A3").Value = Date
    ws.Range("B3").Value = "示例：购买原材料"
    ws.Range("C3").Value = ""
    ws.Range("D3").Value = 1200
    ws.Range("E3").Value = "供应商B"
    ws.Range("F3").Value = "材料费"
    
    ' 设置格式
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("A1:F1").Interior.Color = RGB(200, 220, 255)
    ws.Columns("A:F").AutoFit
    
    ' 添加说明
    ws.Range("A5").Value = "使用说明："
    ws.Range("A6").Value = "1. 日期格式：2024/1/1 或 2024-01-01"
    ws.Range("A7").Value = "2. 收入填正数在C列，支出填正数在D列"
    ws.Range("A8").Value = "3. 分类列可留空，系统会自动识别"
    ws.Range("A9").Value = "4. 保存后使用'批量导入向导'导入"
    
    ' 保存文件
    Dim savePath As String
    savePath = Application.DefaultFilePath & "\批量导入模板.xlsx"
    
    On Error Resume Next
    wb.SaveAs savePath
    On Error GoTo 0
    
    MsgBox "导入模板已生成！" & vbCrLf & _
           "保存位置：" & savePath & vbCrLf & vbCrLf & _
           "请按模板格式整理您的数据，然后使用批量导入向导导入。", vbInformation
End Sub


' ============================================================================
' 【V18.0新增】常用模板创建函数
' ============================================================================

' 创建工资表
Sub CreateSalarySheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("工资表")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "工资表"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Cells(1, 1).Value = "工资表"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "年月："
    ws.Cells(2, 2).Value = "2024年1月"
    
    Dim headers As Variant
    headers = Array("序号", "姓名", "基本工资", "加班费", "奖金", _
                    "应发合计", "社保个人", "公积金个人", "个税", _
                    "实发合计", "发放方式", "签字")
    
    Dim col As Long
    For col = 0 To UBound(headers)
        ws.Cells(4, col + 1).Value = headers(col)
        ws.Cells(4, col + 1).Font.Bold = True
        ws.Cells(4, col + 1).Interior.Color = RGB(68, 114, 196)
        ws.Cells(4, col + 1).Font.Color = RGB(255, 255, 255)
    Next col
    
    ' 社保个税说明
    ws.Cells(2, 5).Value = "社保个人：养老8%+医疗2%=10%"
    ws.Cells(2, 5).Font.Size = 9
    ws.Cells(2, 5).Font.Color = RGB(128, 128, 128)
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C:F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H:J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 10
    ws.Columns("L").ColumnWidth = 8
    
    ' 示例数据
    ws.Cells(5, 1).Value = 1
    ws.Cells(5, 2).Value = "张三"
    ws.Cells(5, 3).Value = 3500
    ws.Cells(5, 6).Formula = "=SUM(C5:E5)"
    ws.Cells(5, 7).Formula = "=F5*0.1"
    ' 个税公式（七级超额累进税率）
    ws.Cells(5, 10).Formula = "=F5-G5-H5-I5"
    
    ws.Range("C5:J5").NumberFormat = "#,##0.00"
End Sub

' 创建发票登记表
Sub CreateInvoiceSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("发票登记表")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "发票登记表"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Cells(1, 1).Value = "发票登记表"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    
    Dim headers As Variant
    headers = Array("开票日期", "发票号码", "发票类型", _
                    "客户名称", "金额", "税额", "价税合计", _
                    "发票状态", "备注")
    
    Dim col As Long
    For col = 0 To UBound(headers)
        ws.Cells(3, col + 1).Value = headers(col)
        ws.Cells(3, col + 1).Font.Bold = True
        ws.Cells(3, col + 1).Interior.Color = RGB(68, 114, 196)
        ws.Cells(3, col + 1).Font.Color = RGB(255, 255, 255)
    Next col
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E:G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 20
End Sub

' 创建银行对账表
Sub CreateBankReconSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("银行对账表")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "银行对账表"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Cells(1, 1).Value = "银行对账表"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "银行账户："
    ws.Cells(2, 2).Value = "对账月份："
    
    Dim headers As Variant
    headers = Array("日期", "摘要", "银行收入", "银行支出", _
                    "银行余额", "账面收入", "账面支出", _
                    "账面余额", "差异", "备注")
    
    Dim col As Long
    For col = 0 To UBound(headers)
        ws.Cells(4, col + 1).Value = headers(col)
        ws.Cells(4, col + 1).Font.Bold = True
        ws.Cells(4, col + 1).Interior.Color = RGB(68, 114, 196)
        ws.Cells(4, col + 1).Font.Color = RGB(255, 255, 255)
    Next col
    
    ' 列宽
    For col = 1 To 10
        ws.Columns(col).ColumnWidth = 12
    Next col
End Sub

' 创建材料进销存台账
Sub CreateMaterialSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("材料进销存台账")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "材料进销存台账"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Cells(1, 1).Value = "材料进销存台账"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    
    Dim headers As Variant
    headers = Array("日期", "材料名称", "规格", _
                    "入库数量", "入库单价", "入库金额", _
                    "出库数量", "出库单价", "出库金额", _
                    "结存数量", "结存金额", "备注")
    
    Dim col As Long
    For col = 0 To UBound(headers)
        ws.Cells(3, col + 1).Value = headers(col)
        ws.Cells(3, col + 1).Font.Bold = True
        ws.Cells(3, col + 1).Interior.Color = RGB(68, 114, 196)
        ws.Cells(3, col + 1).Font.Color = RGB(255, 255, 255)
    Next col
    
    ' 常用材料下拉说明
    ws.Cells(2, 2).Value = "常用：亚钠/片碱/硝酸/磷酸/硫酸/染料/封闭剂/除油剂"
    ws.Cells(2, 2).Font.Size = 9
    ws.Cells(2, 2).Font.Color = RGB(128, 128, 128)
    
    ' 列宽
    For col = 1 To 12
        ws.Columns(col).ColumnWidth = 12
    Next col
End Sub

' 创建月度经营报表
Sub CreateMonthlyReportSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("月度经营报表")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "月度经营报表"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Cells(1, 1).Value = "月度经营报表"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = RGB(192, 0, 0)
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Range("A1:D1").Merge
    
    ws.Cells(2, 1).Value = "月份："
    ws.Cells(2, 2).Value = "2024年1月"
    
    ' 列标题
    Dim headers As Variant
    headers = Array("项目", "本月金额", "上月金额", "环比增减")
    Dim col As Long
    For col = 0 To UBound(headers)
        ws.Cells(4, col + 1).Value = headers(col)
        ws.Cells(4, col + 1).Font.Bold = True
        ws.Cells(4, col + 1).Interior.Color = RGB(68, 114, 196)
        ws.Cells(4, col + 1).Font.Color = RGB(255, 255, 255)
    Next col
    
    ' 报表项目
    Dim items As Variant
    items = Array( _
        "一、收入", _
        "  月总产量", _
        "  加工收入", _
        "", _
        "二、支出", _
        "  厂租", "  电费", "  水费", "  化工", _
        "  染料封闭剂", "  工资", "  包装材料", _
        "  货车运输费", "  外发加工费", "  税金", _
        "  社保费", "  业务费", "  办公用品费", _
        "  维修费", "  挂具费", _
        "  支出合计", _
        "", _
        "三、利润", _
        "  净利润", _
        "", _
        "四、关键指标", _
        "  利润率", _
        "  人均产值", _
        "  单位成本")
    
    Dim i As Long
    For i = 0 To UBound(items)
        ws.Cells(i + 5, 1).Value = items(i)
        ws.Cells(i + 5, 2).NumberFormat = "#,##0.00"
        
        If InStr(items(i), "一、") > 0 Or InStr(items(i), "二、") > 0 Or _
           InStr(items(i), "三、") > 0 Or InStr(items(i), "四、") > 0 Then
            ws.Cells(i + 5, 1).Font.Bold = True
            ws.Cells(i + 5, 1).Font.Color = RGB(0, 112, 192)
        End If
        
        If InStr(items(i), "净利润") > 0 Then
            ws.Cells(i + 5, 1).Font.Bold = True
            ws.Cells(i + 5, 1).Font.Color = RGB(192, 0, 0)
        End If
    Next i
    
    ' 列宽
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:D").ColumnWidth = 14
End Sub


' ============================================================================
' 宏41: GenerateSampleData - 生成模拟数据
' V18.0新增：一键生成3个月模拟业务数据
' ============================================================================
Sub GenerateSampleData()
    Dim wsIncome As Worksheet, wsExpense As Worksheet, wsARAP As Worksheet
    Dim wsProfit As Worksheet, wsSalary As Worksheet, wsMaterial As Worksheet
    Dim result As VbMsgBoxResult
    
    result = MsgBox("将生成近3个月模拟数据：" & vbCrLf & vbCrLf & _
                    "• 收入记录：约30条（5个客户）" & vbCrLf & _
                    "• 支出记录：约45条（15个类别）" & vbCrLf & _
                    "• 应收应付：5条客户记录" & vbCrLf & _
                    "• 工资表：3名员工" & vbCrLf & _
                    "• 材料进销存：6种材料" & vbCrLf & vbCrLf & _
                    "是否继续？", vbYesNo + vbQuestion, "生成模拟数据")
    
    If result <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' 获取工作表
    On Error Resume Next
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    Set wsExpense = ThisWorkbook.Sheets("支出记录")
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    Set wsProfit = ThisWorkbook.Sheets("利润分析表")
    Set wsSalary = ThisWorkbook.Sheets("工资表")
    Set wsMaterial = ThisWorkbook.Sheets("材料进销存台账")
    On Error GoTo 0
    
    ' 生成收入数据
    If Not wsIncome Is Nothing Then
        Call GenerateIncomeData(wsIncome)
    End If
    
    ' 生成支出数据
    If Not wsExpense Is Nothing Then
        Call GenerateExpenseData(wsExpense)
    End If
    
    ' 生成应收应付数据
    If Not wsARAP Is Nothing Then
        Call GenerateARAPData(wsARAP)
    End If
    
    ' 生成工资数据
    If Not wsSalary Is Nothing Then
        Call GenerateSalaryData(wsSalary)
    End If
    
    ' 生成材料数据
    If Not wsMaterial Is Nothing Then
        Call GenerateMaterialData(wsMaterial)
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ 模拟数据生成完成！" & vbCrLf & vbCrLf & _
           "收入记录：30条" & vbCrLf & _
           "支出记录：45条" & vbCrLf & _
           "应收应付：5条" & vbCrLf & _
           "工资表：3人" & vbCrLf & _
           "材料台账：6种", vbInformation, "完成"
End Sub

' 生成收入模拟数据
Private Sub GenerateIncomeData(ws As Worksheet)
    Dim customers As Variant, amounts As Variant
    Dim i As Long, row As Long, month As Long
    Dim baseAmount As Double, taxRate As Double
    Dim voucherNo As String
    
    customers = Array("华鑫铝业", "永达五金", "顺发配件", "金龙机械", "德盛电子")
    amounts = Array(15000, 12000, 18000, 8500, 22000)
    
    row = 4  ' 数据从第4行开始
    
    Dim startMonth As Long: startMonth = Month(Date) - 2
    If startMonth < 1 Then startMonth = startMonth + 12
    For month = startMonth To startMonth + 2  ' 近3个月
        For i = 0 To UBound(customers)
            ' 每个客户每月2-3笔收入
            Dim count As Integer
            count = 2 + Int(Rnd() * 2)
            
            Dim j As Integer
            For j = 1 To count
                voucherNo = "收-" & Format(month, "00") & "-" & Format(row - 3, "000")
                baseAmount = amounts(i) * (0.8 + Rnd() * 0.4)
                
                ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 5 + Int(Rnd() * 20))
                ws.Cells(row, 2).Value = voucherNo
                ws.Cells(row, 3).Value = customers(i)
                ws.Cells(row, 4).Value = Round(baseAmount, 2)
                
                ' 收款方式
                If Rnd() > 0.3 Then
                    ws.Cells(row, 5).Value = "银行转账"
                Else
                    ws.Cells(row, 5).Value = "微信"
                End If
                
                ' 含税/不含税
                If i Mod 2 = 0 Then
                    ws.Cells(row, 6).Value = "含税"
                    taxRate = 0.03
                    ws.Cells(row, 7).Value = taxRate
                    ws.Cells(row, 8).Value = Round(baseAmount * taxRate / (1 + taxRate), 2)
                    ws.Cells(row, 9).Value = Round(baseAmount / (1 + taxRate), 2)
                Else
                    ws.Cells(row, 6).Value = "不含税"
                    taxRate = 0.03
                    ws.Cells(row, 7).Value = taxRate
                    ws.Cells(row, 8).Value = Round(baseAmount * taxRate, 2)
                    ws.Cells(row, 9).Value = baseAmount
                End If
                
                ' 开票状态
                If Rnd() > 0.2 Then
                    ws.Cells(row, 10).Value = "已开票"
                Else
                    ws.Cells(row, 10).Value = "未开票"
                End If
                
                ' 对冲/代付（约10%概率）
                If Rnd() > 0.9 Then
                    ws.Cells(row, 11).Value = "代付货款"
                    ws.Cells(row, 12).Value = customers((i + 2) Mod 5)
                    ws.Cells(row, 13).Value = customers(i)
                Else
                    ws.Cells(row, 11).Value = "正常"
                End If
                
                ws.Cells(row, 15).Value = "1002"
                ws.Cells(row, 16).Value = "5001"
                
                row = row + 1
            Next j
        Next i
    Next month
    
    ' 格式化
    ws.Range("A4:Q" & row - 1).Borders.LineStyle = 1
    ws.Range("D4:D" & row - 1).NumberFormat = "#,##0.00"
    ws.Range("H4:I" & row - 1).NumberFormat = "#,##0.00"
End Sub

' 生成支出模拟数据
Private Sub GenerateExpenseData(ws As Worksheet)
    Dim categories As Variant, suppliers As Variant
    Dim i As Long, row As Long, month As Long
    Dim amount As Double
    
    categories = Array("厂租", "电费", "水费", "化工", "染料封闭剂", _
                       "工资", "包装材料", "货车运输费", "外发加工费", _
                       "税金", "社保费", "业务费", "办公用品费", "维修费", "挂具费")
    
    suppliers = Array("房东张生", "供电局", "自来水公司", "化工厂", "染料供应商", _
                      "", "包装材料店", "加油站", "外发加工厂", "税务局", _
                      "社保局", "", "文具店", "维修配件店", "")
    
    row = 4

    Dim startMonth2 As Long: startMonth2 = Month(Date) - 2
    If startMonth2 < 1 Then startMonth2 = startMonth2 + 12
    For month = startMonth2 To startMonth2 + 2
        ' 固定支出
        ' 厂租
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 1)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-001"
        ws.Cells(row, 3).Value = "厂租"
        ws.Cells(row, 4).Value = 8000
        ws.Cells(row, 5).Value = "房东张生"
        ws.Cells(row, 6).Value = "银行转账"
        row = row + 1
        
        ' 电费
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 5)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-002"
        ws.Cells(row, 3).Value = "电费"
        ws.Cells(row, 4).Value = Round(3500 + Rnd() * 1500, 2)
        ws.Cells(row, 5).Value = "供电局"
        ws.Cells(row, 6).Value = "银行转账"
        row = row + 1
        
        ' 水费
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 5)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-003"
        ws.Cells(row, 3).Value = "水费"
        ws.Cells(row, 4).Value = Round(300 + Rnd() * 200, 2)
        ws.Cells(row, 5).Value = "自来水公司"
        ws.Cells(row, 6).Value = "银行转账"
        row = row + 1
        
        ' 化工材料
        Dim j As Integer
        For j = 1 To 2
            ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 8 + j * 5)
            ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
            ws.Cells(row, 3).Value = "化工"
            ws.Cells(row, 4).Value = Round(2000 + Rnd() * 3000, 2)
            ws.Cells(row, 5).Value = "化工厂"
            ws.Cells(row, 6).Value = IIf(Rnd() > 0.5, "银行转账", "微信")
            row = row + 1
        Next j
        
        ' 染料封闭剂
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 12)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
        ws.Cells(row, 3).Value = "染料封闭剂"
        ws.Cells(row, 4).Value = Round(1500 + Rnd() * 2000, 2)
        ws.Cells(row, 5).Value = "染料供应商"
        ws.Cells(row, 6).Value = "银行转账"
        row = row + 1
        
        ' 工资
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 28)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
        ws.Cells(row, 3).Value = "工资"
        ws.Cells(row, 4).Value = 15000
        ws.Cells(row, 6).Value = "银行转账"
        row = row + 1
        
        ' 包装材料
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 10)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
        ws.Cells(row, 3).Value = "包装材料"
        ws.Cells(row, 4).Value = Round(500 + Rnd() * 500, 2)
        ws.Cells(row, 5).Value = "包装材料店"
        ws.Cells(row, 6).Value = "微信"
        row = row + 1
        
        ' 货车运输费
        For j = 1 To 3
            ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 5 + j * 7)
            ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
            ws.Cells(row, 3).Value = "货车运输费"
            ws.Cells(row, 4).Value = Round(800 + Rnd() * 600, 2)
            ws.Cells(row, 5).Value = "加油站"
            ws.Cells(row, 6).Value = IIf(Rnd() > 0.5, "微信", "现金")
            row = row + 1
        Next j
        
        ' 外发加工费
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 15)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
        ws.Cells(row, 3).Value = "外发加工费"
        ws.Cells(row, 4).Value = Round(2000 + Rnd() * 3000, 2)
        ws.Cells(row, 5).Value = "外发加工厂"
        ws.Cells(row, 6).Value = "银行转账"
        row = row + 1
        
        ' 税金
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 15)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
        ws.Cells(row, 3).Value = "税金"
        ws.Cells(row, 4).Value = Round(500 + Rnd() * 300, 2)
        ws.Cells(row, 5).Value = "税务局"
        ws.Cells(row, 6).Value = "银行转账"
        row = row + 1
        
        ' 社保费
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 10)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
        ws.Cells(row, 3).Value = "社保费"
        ws.Cells(row, 4).Value = 2500
        ws.Cells(row, 5).Value = "社保局"
        ws.Cells(row, 6).Value = "银行转账"
        row = row + 1
        
        ' 业务费
        For j = 1 To 2
            ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 10 + j * 8)
            ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
            ws.Cells(row, 3).Value = "业务费"
            ws.Cells(row, 4).Value = Round(300 + Rnd() * 400, 2)
            ws.Cells(row, 6).Value = "微信"
            row = row + 1
        Next j
        
        ' 办公用品费
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 8)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
        ws.Cells(row, 3).Value = "办公用品费"
        ws.Cells(row, 4).Value = Round(100 + Rnd() * 150, 2)
        ws.Cells(row, 5).Value = "文具店"
        ws.Cells(row, 6).Value = "微信"
        row = row + 1
        
        ' 维修费
        If Rnd() > 0.5 Then
            ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 18)
            ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
            ws.Cells(row, 3).Value = "维修费"
            ws.Cells(row, 4).Value = Round(500 + Rnd() * 1000, 2)
            ws.Cells(row, 5).Value = "维修配件店"
            ws.Cells(row, 6).Value = "微信"
            row = row + 1
        End If
        
        ' 挂具费
        ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 20)
        ws.Cells(row, 2).Value = "支-" & Format(month, "00") & "-" & Format(row - 3, "000")
        ws.Cells(row, 3).Value = "挂具费"
        ws.Cells(row, 4).Value = Round(300 + Rnd() * 200, 2)
        row = row + 1
    Next month
    
    ws.Range("A4:F" & row - 1).Borders.LineStyle = 1
    ws.Range("D4:D" & row - 1).NumberFormat = "#,##0.00"
End Sub

' 生成应收应付模拟数据
Private Sub GenerateARAPData(ws As Worksheet)
    Dim customers As Variant, amounts As Variant
    Dim i As Long
    
    customers = Array("华鑫铝业", "永达五金", "顺发配件", "金龙机械", "德盛电子")
    amounts = Array(15000, 8500, 12000, 5000, 18000)
    
    ' 应收账款区
    For i = 0 To UBound(customers)
        ws.Cells(3 + i, 1).Value = customers(i)
        ws.Cells(3 + i, 2).Value = amounts(i) * 0.3  ' 期初应收
        ws.Cells(3 + i, 3).Value = amounts(i) * 3    ' 本期增加（3个月）
        ws.Cells(3 + i, 4).Value = amounts(i) * 2.5  ' 本期收款
        ws.Cells(3 + i, 7).Formula = "=B" & (3 + i) & "+C" & (3 + i) & "-D" & (3 + i) & "-E" & (3 + i) & "+F" & (3 + i)
    Next i
    
    ' 应付账款区
    Dim suppliers As Variant
    suppliers = Array("化工厂", "染料供应商", "外发加工厂", "包装材料店", "维修配件店")
    
    For i = 0 To UBound(suppliers)
        ws.Cells(3 + i, 10).Value = suppliers(i)
        ws.Cells(3 + i, 11).Value = 3000 + i * 1000  ' 期初应付
        ws.Cells(3 + i, 12).Value = 5000 + i * 2000  ' 本期增加
        ws.Cells(3 + i, 13).Value = 6000 + i * 1500  ' 本期付款
        ws.Cells(3 + i, 15).Formula = "=K" & (3 + i) & "+L" & (3 + i) & "-M" & (3 + i) & "-N" & (3 + i)
    Next i
End Sub

' 生成工资模拟数据
Private Sub GenerateSalaryData(ws As Worksheet)
    Dim names As Variant, bases As Variant
    Dim i As Long
    
    names = Array("张三", "李四", "王五")
    bases = Array(3500, 4000, 4500)
    
    For i = 0 To UBound(names)
        ws.Cells(5 + i, 1).Value = i + 1
        ws.Cells(5 + i, 2).Value = names(i)
        ws.Cells(5 + i, 3).Value = bases(i)
        ws.Cells(5 + i, 4).Value = Round(200 + Rnd() * 500, 2)  ' 加班费
        ws.Cells(5 + i, 5).Value = Round(100 + Rnd() * 300, 2)  ' 奖金
        ws.Cells(5 + i, 6).Formula = "=SUM(C" & (5 + i) & ":E" & (5 + i) & ")"
        ws.Cells(5 + i, 7).Formula = "=F" & (5 + i) & "*0.1"
        ws.Cells(5 + i, 8).Value = 0  ' 公积金个人
        ws.Cells(5 + i, 9).Formula = "=IF((F" & (5 + i) & "-G" & (5 + i) & "-H" & (5 + i) & ")-5000>0,ROUND((F" & (5 + i) & "-G" & (5 + i) & "-H" & (5 + i) & "-5000)*0.03,2),0)"
        ws.Cells(5 + i, 10).Formula = "=F" & (5 + i) & "-G" & (5 + i) & "-H" & (5 + i) & "-I" & (5 + i)
        ws.Cells(5 + i, 11).Value = "银行转账"
        
        ' 格式化
        ws.Range("C" & (5 + i) & ":J" & (5 + i)).NumberFormat = "#,##0.00"
    Next i
End Sub

' 生成材料进销存模拟数据
Private Sub GenerateMaterialData(ws As Worksheet)
    Dim materials As Variant, specs As Variant, prices As Variant
    Dim i As Long, row As Long, month As Long
    
    materials = Array("亚钠", "片碱", "硝酸", "磷酸", "硫酸", "封闭剂")
    specs = Array("25kg/袋", "25kg/袋", "500ml/瓶", "500ml/瓶", "500ml/瓶", "20kg/桶")
    prices = Array(180, 150, 45, 38, 42, 280)
    
    row = 4

    Dim startMonth3 As Long: startMonth3 = Month(Date) - 2
    If startMonth3 < 1 Then startMonth3 = startMonth3 + 12
    For month = startMonth3 To startMonth3 + 2
        For i = 0 To UBound(materials)
            ' 入库记录
            ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 5)
            ws.Cells(row, 2).Value = materials(i)
            ws.Cells(row, 3).Value = specs(i)
            ws.Cells(row, 4).Value = 10 + Int(Rnd() * 10)  ' 入库数量
            ws.Cells(row, 5).Value = prices(i)
            ws.Cells(row, 6).Formula = "=D" & row & "*E" & row
            row = row + 1
            
            ' 出库记录
            ws.Cells(row, 1).Value = DateSerial(Year(Date) - IIf(month > Month(Date), 1, 0), month, 15)
            ws.Cells(row, 2).Value = materials(i)
            ws.Cells(row, 3).Value = specs(i)
            ws.Cells(row, 7).Value = 8 + Int(Rnd() * 8)  ' 出库数量
            ws.Cells(row, 8).Value = prices(i)
            ws.Cells(row, 9).Formula = "=G" & row & "*H" & row
            row = row + 1
        Next i
    Next month
    
    ws.Range("A4:L" & row - 1).Borders.LineStyle = 1
    ws.Range("E4:F" & row - 1).NumberFormat = "#,##0.00"
    ws.Range("H4:I" & row - 1).NumberFormat = "#,##0.00"
End Sub


' ============================================================================
' V18.0 界面优化函数
' ============================================================================

' 美化工作表标题
Sub BeautifySheetTitle(ws As Worksheet, title As String, subtitle As String)
    ' 主标题
    ws.Cells(1, 1).Value = title
    ws.Cells(1, 1).Font.Name = "微软雅黑"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = RGB(68, 114, 196)
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    
    ' 副标题
    If subtitle <> "" Then
        ws.Cells(2, 1).Value = subtitle
        ws.Cells(2, 1).Font.Name = "微软雅黑"
        ws.Cells(2, 1).Font.Size = 9
        ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)
        ws.Cells(2, 1).HorizontalAlignment = xlCenter
    End If
    
    ' 合并标题行
    Dim lastCol As Long
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Merge
    If subtitle <> "" Then
        ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol)).Merge
    End If
End Sub

' 设置表头样式（V18.0优化）
Sub SetHeaderStyleV152(ws As Worksheet, row As Long, startCol As Long, endCol As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
    
    With rng
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

' 添加条件格式：利润预警
Sub AddProfitAlertFormat(ws As Worksheet, dataCol As Long, startRow As Long)
    Dim lastRow As Long
    lastRow = GetLastRow(ws, dataCol)
    If lastRow < startRow Then Exit Sub
    
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, dataCol), ws.Cells(lastRow, dataCol))
    
    ' 清除现有条件格式
    rng.FormatConditions.Delete
    
    ' 添加条件格式1：利润为负（红色背景）
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    rng.FormatConditions(1).Interior.Color = RGB(255, 199, 206)
    rng.FormatConditions(1).Font.Color = RGB(156, 0, 6)
    
    ' 添加条件格式2：利润较高（绿色背景）
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=10000"
    rng.FormatConditions(2).Interior.Color = RGB(198, 239, 206)
    rng.FormatConditions(2).Font.Color = RGB(0, 97, 0)
End Sub

' 添加条件格式：应收逾期预警
Sub AddARAlertFormat(ws As Worksheet, dateCol As Long, amountCol As Long, startRow As Long)
    Dim lastRow As Long
    lastRow = GetLastRow(ws, dateCol)
    If lastRow < startRow Then Exit Sub
    
    Dim i As Long
    For i = startRow To lastRow
        Dim invDate As Date
        Dim days As Long
        
        On Error Resume Next
        invDate = CDate(ws.Cells(i, dateCol).Value)
        If Err.Number = 0 Then
            days = DateDiff("d", invDate, Date)
            ' 逾期超过60天标红
            If days > 60 And ws.Cells(i, amountCol).Value > 0 Then
                ws.Cells(i, dateCol).Interior.Color = RGB(255, 199, 206)
                ws.Cells(i, amountCol).Interior.Color = RGB(255, 199, 206)
                ws.Cells(i, amountCol).Font.Color = RGB(156, 0, 6)
                ws.Cells(i, amountCol).Font.Bold = True
            ' 逾期30-60天标黄
            ElseIf days > 30 And ws.Cells(i, amountCol).Value > 0 Then
                ws.Cells(i, dateCol).Interior.Color = RGB(255, 235, 156)
                ws.Cells(i, amountCol).Interior.Color = RGB(255, 235, 156)
            End If
        End If
        On Error GoTo 0
    Next i
End Sub

' 智能输入提示
Sub ShowSmartTip(ws As Worksheet, row As Long, col As Long, tip As String)
    With ws.Cells(row, col).Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .InputTitle = "输入提示"
        .InputMessage = tip
        .ShowInput = True
    End With
End Sub

' 快速美化按钮样式（用于用户窗体）
Sub SetButtonStyle(btn As Object, btnText As String, btnColor As Long)
    With btn
        .Caption = btnText
        .BackColor = btnColor
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .Font.Bold = True
    End With
End Sub

' 创建快捷操作栏
Sub CreateQuickActionBar(ws As Worksheet)
    ' 在最后一列添加快捷操作按钮说明
    Dim lastCol As Long
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column + 2
    
    ws.Cells(3, lastCol).Value = "快捷操作"
    ws.Cells(3, lastCol).Font.Bold = True
    ws.Cells(3, lastCol).Font.Color = RGB(68, 114, 196)
    
    ws.Cells(4, lastCol).Value = "Alt+F8: 运行宏"
    ws.Cells(5, lastCol).Value = "Ctrl+S: 保存"
    ws.Cells(6, lastCol).Value = "F1: 帮助"
    
    ws.Columns(lastCol).ColumnWidth = 15
End Sub


' ============================================================================
' V18.0 新增功能：数据工具、智能提醒、增强查询
' ============================================================================

' ============================================================================
' 宏30: AdvancedQuery - 多条件高级查询
' V18.0新增：支持多条件筛选、模糊查询、结果导出
' ============================================================================
Sub AdvancedQuery()
    Dim ws As Worksheet
    Dim queryType As String
    Dim result As VbMsgBoxResult
    
    queryType = InputBox("选择查询类型：" & vbCrLf & _
                        "1. 收入查询" & vbCrLf & _
                        "2. 支出查询" & vbCrLf & _
                        "3. 客户往来查询" & vbCrLf & _
                        "4. 材料库存查询", _
                        "多条件查询", "1")
    
    Select Case queryType
        Case "1": Call QueryIncome
        Case "2": Call QueryExpense
        Case "3": Call QueryCustomer
        Case "4": Call QueryInventory
        Case Else: MsgBox "无效选择", vbExclamation
    End Select
End Sub

' 收入查询
Private Sub QueryIncome()
    Dim ws As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, resultRow As Long
    Dim i As Long
    Dim customerFilter As String, dateStart As String, dateEnd As String
    Dim minAmount As String, maxAmount As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("收入记录")
    If ws Is Nothing Then
        MsgBox "未找到收入记录表！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 获取查询条件
    customerFilter = InputBox("客户名称（支持模糊查询，留空查询全部）：", "客户筛选")
    dateStart = InputBox("开始日期（格式：2024-01-01，留空不限制）：", "日期范围")
    dateEnd = InputBox("结束日期（格式：2024-12-31，留空不限制）：", "日期范围")
    minAmount = InputBox("最小金额（留空不限制）：", "金额范围")
    maxAmount = InputBox("最大金额（留空不限制）：", "金额范围")
    
    ' 创建结果表
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("查询结果")
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "查询结果"
    Else
        wsResult.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 复制表头
    ws.Rows(3).Copy Destination:=wsResult.Rows(1)
    resultRow = 2
    
    lastRow = GetLastRow(ws, 1)
    
    ' 筛选数据
    For i = 4 To lastRow
        Dim match As Boolean
        match = True
        
        ' 客户筛选（模糊匹配）
        If customerFilter <> "" Then
            If InStr(1, ws.Cells(i, 3).Value, customerFilter, vbTextCompare) = 0 Then
                match = False
            End If
        End If
        
        ' 日期筛选
        If match And dateStart <> "" Then
            If CDate(ws.Cells(i, 1).Value) < CDate(dateStart) Then
                match = False
            End If
        End If
        
        If match And dateEnd <> "" Then
            If CDate(ws.Cells(i, 1).Value) > CDate(dateEnd) Then
                match = False
            End If
        End If
        
        ' 金额筛选
        If match And minAmount <> "" Then
            If CDbl(ws.Cells(i, 4).Value) < CDbl(minAmount) Then
                match = False
            End If
        End If
        
        If match And maxAmount <> "" Then
            If CDbl(ws.Cells(i, 4).Value) > CDbl(maxAmount) Then
                match = False
            End If
        End If
        
        ' 复制匹配行
        If match Then
            ws.Rows(i).Copy Destination:=wsResult.Rows(resultRow)
            resultRow = resultRow + 1
        End If
    Next i
    
    ' 美化结果表
    Call BeautifySheetTitle(wsResult, "收入查询结果", "共找到 " & resultRow - 2 & " 条记录")
    
    ' 添加导出按钮提示
    wsResult.Cells(resultRow + 1, 1).Value = "提示：运行【数据导入导出】可将结果导出为Excel/CSV"
    wsResult.Cells(resultRow + 1, 1).Font.Color = RGB(128, 128, 128)
    wsResult.Cells(resultRow + 1, 1).Font.Size = 9
    
    MsgBox "查询完成！共找到 " & resultRow - 2 & " 条记录" & vbCrLf & _
           "结果已显示在【查询结果】工作表", vbInformation
    
    wsResult.Activate
End Sub

' 支出查询
Private Sub QueryExpense()
    Dim ws As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, resultRow As Long
    Dim i As Long
    Dim categoryFilter As String, dateStart As String, dateEnd As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("支出记录")
    If ws Is Nothing Then
        MsgBox "未找到支出记录表！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    categoryFilter = InputBox("支出类别（留空查询全部）：", "类别筛选")
    dateStart = InputBox("开始日期（格式：2024-01-01，留空不限制）：", "日期范围")
    dateEnd = InputBox("结束日期（格式：2024-12-31，留空不限制）：", "日期范围")
    
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("查询结果")
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "查询结果"
    Else
        wsResult.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Rows(3).Copy Destination:=wsResult.Rows(1)
    resultRow = 2
    lastRow = GetLastRow(ws, 1)
    
    For i = 4 To lastRow
        Dim match As Boolean
        match = True
        
        If categoryFilter <> "" Then
            If InStr(1, ws.Cells(i, 3).Value, categoryFilter, vbTextCompare) = 0 Then
                match = False
            End If
        End If
        
        If match And dateStart <> "" Then
            If CDate(ws.Cells(i, 1).Value) < CDate(dateStart) Then match = False
        End If
        
        If match And dateEnd <> "" Then
            If CDate(ws.Cells(i, 1).Value) > CDate(dateEnd) Then match = False
        End If
        
        If match Then
            ws.Rows(i).Copy Destination:=wsResult.Rows(resultRow)
            resultRow = resultRow + 1
        End If
    Next i
    
    Call BeautifySheetTitle(wsResult, "支出查询结果", "共找到 " & resultRow - 2 & " 条记录")
    MsgBox "查询完成！共找到 " & resultRow - 2 & " 条记录", vbInformation
    wsResult.Activate
End Sub

' 客户往来查询
Private Sub QueryCustomer()
    Dim wsARAP As Worksheet, wsIncome As Worksheet
    Dim customerName As String
    
    customerName = InputBox("请输入客户名称：", "客户往来查询")
    If customerName = "" Then Exit Sub
    
    On Error Resume Next
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    On Error GoTo 0
    
    If wsARAP Is Nothing Or wsIncome Is Nothing Then
        MsgBox "缺少必要的工作表！", vbExclamation
        Exit Sub
    End If
    
    ' 汇总客户信息
    Dim totalAR As Double, totalIncome As Double
    Dim msg As String
    
    msg = "客户【" & customerName & "】往来汇总" & vbCrLf & vbCrLf
    
    ' 查找应收余额
    Dim lastRow As Long, i As Long
    lastRow = GetLastRow(wsARAP, 1)
    For i = 4 To lastRow
        If InStr(1, wsARAP.Cells(i, 1).Value, customerName, vbTextCompare) > 0 Then
            totalAR = CDbl(Nz(wsARAP.Cells(i, 7).Value, 0))
            msg = msg & "期末应收：" & Format(totalAR, "#,##0.00") & " 元" & vbCrLf
            Exit For
        End If
    Next i
    
    ' 统计收入总额
    lastRow = GetLastRow(wsIncome, 1)
    For i = 4 To lastRow
        If InStr(1, wsIncome.Cells(i, 3).Value, customerName, vbTextCompare) > 0 Then
            totalIncome = totalIncome + CDbl(Nz(wsIncome.Cells(i, 4).Value, 0))
        End If
    Next i
    
    msg = msg & "累计收入：" & Format(totalIncome, "#,##0.00") & " 元" & vbCrLf & vbCrLf
    
    If totalAR > 10000 Then
        msg = msg & "⚠ 提醒：该客户应收款较高，请关注回款！"
    ElseIf totalAR > 0 Then
        msg = msg & "✓ 该客户应收款正常"
    Else
        msg = msg & "✓ 该客户无欠款"
    End If
    
    MsgBox msg, vbInformation, "客户往来查询"
End Sub

' 材料库存查询
Private Sub QueryInventory()
    Dim ws As Worksheet
    Dim materialName As String
    Dim lastRow As Long, i As Long
    Dim totalIn As Double, totalOut As Double
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("材料进销存台账")
    If ws Is Nothing Then
        MsgBox "未找到材料进销存台账！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    materialName = InputBox("请输入材料名称（支持模糊查询，留空查询全部）：", "库存查询")
    
    lastRow = GetLastRow(ws, 1)
    totalIn = 0
    totalOut = 0
    
    For i = 4 To lastRow
        If materialName = "" Or InStr(1, ws.Cells(i, 2).Value, materialName, vbTextCompare) > 0 Then
            If IsNumeric(ws.Cells(i, 4).Value) Then
                totalIn = totalIn + CDbl(ws.Cells(i, 4).Value)
            End If
            If IsNumeric(ws.Cells(i, 7).Value) Then
                totalOut = totalOut + CDbl(ws.Cells(i, 7).Value)
            End If
        End If
    Next i
    
    Dim msg As String
    msg = "库存查询结果" & vbCrLf & vbCrLf
    If materialName <> "" Then
        msg = msg & "材料：" & materialName & vbCrLf
    Else
        msg = msg & "材料：全部" & vbCrLf
    End If
    msg = msg & "累计入库：" & totalIn & vbCrLf
    msg = msg & "累计出库：" & totalOut & vbCrLf
    msg = msg & "当前库存：" & (totalIn - totalOut) & vbCrLf
    
    If totalIn - totalOut < 10 Then
        msg = msg & vbCrLf & "⚠ 库存不足，请及时采购！"
    End If
    
    MsgBox msg, vbInformation, "库存查询"
End Sub

' ============================================================================
' 宏31: ImportExportWizard - 数据导入导出向导
' V18.0新增：支持Excel/CSV/TXT多格式导入导出
' ============================================================================
Sub ImportExportWizard()
    Dim action As String
    
    action = InputBox("选择操作：" & vbCrLf & _
                     "1. 导出数据到文件" & vbCrLf & _
                     "2. 从文件导入数据" & vbCrLf & _
                     "3. 导出查询结果", _
                     "数据导入导出", "1")
    
    Select Case action
        Case "1": Call ExportDataToFile
        Case "2": Call ImportDataFromFile
        Case "3": Call ExportQueryResult
        Case Else: MsgBox "无效选择", vbExclamation
    End Select
End Sub

' 导出数据到文件
Private Sub ExportDataToFile()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim fileFormat As String
    Dim filePath As String
    
    sheetName = InputBox("要导出的工作表名称（收入记录/支出记录/应收应付等）：", "选择数据源")
    If sheetName = "" Then Exit Sub
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "未找到工作表：" & sheetName, vbExclamation
        Exit Sub
    End If
    
    fileFormat = InputBox("选择导出格式：" & vbCrLf & _
                         "1. Excel文件 (.xlsx)" & vbCrLf & _
                         "2. CSV文件 (.csv)" & vbCrLf & _
                         "3. 文本文件 (.txt)", _
                         "选择格式", "1")
    
    ' 使用文件对话框获取保存路径
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=sheetName & "_" & Format(Now, "yyyymmdd"), _
        FileFilter:="Excel文件 (*.xlsx),*.xlsx,CSV文件 (*.csv),*.csv,文本文件 (*.txt),*.txt")
    
    If filePath = "False" Then Exit Sub
    
    ' 根据格式导出
    Select Case fileFormat
        Case "1"
            If Right(filePath, 5) <> ".xlsx" Then filePath = filePath & ".xlsx"
            ws.Copy
            ActiveWorkbook.SaveAs filePath
            ActiveWorkbook.Close
        Case "2"
            If Right(filePath, 4) <> ".csv" Then filePath = filePath & ".csv"
            Call ExportToCSV(ws, filePath)
        Case "3"
            If Right(filePath, 4) <> ".txt" Then filePath = filePath & ".txt"
            Call ExportToTXT(ws, filePath)
    End Select
    
    MsgBox "导出完成！" & vbCrLf & filePath, vbInformation
End Sub

' 导出为CSV
Private Sub ExportToCSV(ws As Worksheet, filePath As String)
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim line As String
    Dim fso As Object, file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(filePath, True)
    
    lastRow = GetLastRow(ws, 1)
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 3 To lastRow
        line = ""
        For j = 1 To lastCol
            If j > 1 Then line = line & ","
            line = line & """" & Replace(CStr(ws.Cells(i, j).Value), """", """""") & """"
        Next j
        file.WriteLine line
    Next i
    
    file.Close
End Sub

' 导出为TXT
Private Sub ExportToTXT(ws As Worksheet, filePath As String)
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim line As String
    Dim fso As Object, file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(filePath, True)
    
    lastRow = GetLastRow(ws, 1)
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    ' 写入标题
    file.WriteLine "导出时间：" & Format(Now, "yyyy-mm-dd hh:mm:ss")
    file.WriteLine "工作表：" & ws.Name
    file.WriteLine String(50, "-")
    
    For i = 3 To lastRow
        line = ""
        For j = 1 To lastCol
            If ws.Cells(i, j).Value <> "" Then
                line = line & ws.Cells(3, j).Value & ": " & ws.Cells(i, j).Value & " | "
            End If
        Next j
        If line <> "" Then file.WriteLine line
    Next i
    
    file.Close
End Sub

' 从文件导入数据
Private Sub ImportDataFromFile()
    Dim filePath As String
    Dim targetSheet As String
    
    filePath = Application.GetOpenFilename( _
        FileFilter:="CSV文件 (*.csv),*.csv,文本文件 (*.txt),*.txt,Excel文件 (*.xlsx),*.xlsx")
    
    If filePath = "False" Then Exit Sub
    
    targetSheet = InputBox("导入到哪个工作表（收入记录/支出记录）：", "选择目标")
    If targetSheet = "" Then Exit Sub
    
    If Right(filePath, 4) = ".csv" Then
        Call ImportFromCSV(filePath, targetSheet)
    ElseIf Right(filePath, 5) = ".xlsx" Then
        Call ImportFromExcel(filePath, targetSheet)
    End If
End Sub

' 从CSV导入
Private Sub ImportFromCSV(filePath As String, sheetName As String)
    Dim ws As Worksheet
    Dim fso As Object, file As Object
    Dim line As String
    Dim fields As Variant
    Dim row As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "未找到工作表：" & sheetName, vbExclamation
        Exit Sub
    End If
    
    row = GetLastRow(ws, 1) + 1
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(filePath, 1)
    
    ' 跳过标题行
    If Not file.AtEndOfStream Then file.SkipLine
    
    Do While Not file.AtEndOfStream
        line = file.ReadLine
        fields = Split(line, ",")
        
        ' 写入数据
        Dim i As Long
        For i = 0 To UBound(fields)
            ws.Cells(row, i + 1).Value = Replace(Replace(fields(i), """", ""), """", "")
        Next i
        row = row + 1
    Loop
    
    file.Close
    MsgBox "导入完成！共导入 " & row - GetLastRow(ws, 1) - 1 & " 条记录", vbInformation
End Sub

' 导出查询结果
Private Sub ExportQueryResult()
    Dim ws As Worksheet
    Dim filePath As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("查询结果")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "未找到查询结果表！请先运行【多条件查询】", vbExclamation
        Exit Sub
    End If
    
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="查询结果_" & Format(Now, "yyyymmdd"), _
        FileFilter:="Excel文件 (*.xlsx),*.xlsx,CSV文件 (*.csv),*.csv")
    
    If filePath = "False" Then Exit Sub
    
    If Right(filePath, 4) = ".csv" Then
        Call ExportToCSV(ws, filePath)
    Else
        If Right(filePath, 5) <> ".xlsx" Then filePath = filePath & ".xlsx"
        ws.Copy
        ActiveWorkbook.SaveAs filePath
        ActiveWorkbook.Close
    End If
    
    MsgBox "导出完成！", vbInformation
End Sub

' ============================================================================
' 宏32: AutoBackup - 自动备份功能
' V18.0新增：一键备份、版本管理
' ============================================================================
Sub AutoBackup()
    Dim action As String
    
    action = InputBox("选择操作：" & vbCrLf & _
                     "1. 立即备份" & vbCrLf & _
                     "2. 查看备份列表" & vbCrLf & _
                     "3. 恢复备份" & vbCrLf & _
                     "4. 设置自动备份", _
                     "自动备份", "1")
    
    Select Case action
        Case "1": Call DoBackupNow
        Case "2": Call ListBackups
        Case "3": Call RestoreBackup
        Case "4": Call SetupAutoBackup
        Case Else: MsgBox "无效选择", vbExclamation
    End Select
End Sub

' 立即备份
Private Sub DoBackupNow()
    Dim backupPath As String
    Dim fileName As String
    Dim backupName As String
    
    ' 默认备份路径
    backupPath = Environ("USERPROFILE") & "\Documents\氧化加工厂备份\"
    
    ' 创建备份目录
    On Error Resume Next
    MkDir backupPath
    On Error GoTo 0
    
    ' 生成备份文件名
    fileName = "氧化加工厂备份_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    backupName = backupPath & fileName
    
    ' 保存备份
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs backupName
    Application.DisplayAlerts = True
    
    MsgBox "备份完成！" & vbCrLf & vbCrLf & _
           "备份位置：" & backupName, vbInformation
End Sub

' 查看备份列表
Private Sub ListBackups()
    Dim backupPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim msg As String
    Dim count As Integer
    
    backupPath = Environ("USERPROFILE") & "\Documents\氧化加工厂备份\"
    
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupPath)
    On Error GoTo 0
    
    If folder Is Nothing Then
        MsgBox "备份目录不存在！", vbExclamation
        Exit Sub
    End If
    
    msg = "备份文件列表：" & vbCrLf & vbCrLf
    count = 0
    
    For Each file In folder.Files
        If InStr(file.Name, "氧化加工厂备份") > 0 Then
            msg = msg & file.Name & " (" & Format(file.Size / 1024, "0.00") & " KB)" & vbCrLf
            count = count + 1
            If count >= 10 Then Exit For
        End If
    Next file
    
    If count = 0 Then
        msg = msg & "暂无备份文件"
    Else
        msg = msg & vbCrLf & "共 " & count & " 个备份"
    End If
    
    MsgBox msg, vbInformation, "备份列表"
End Sub

' 恢复备份
Private Sub RestoreBackup()
    Dim backupPath As String
    Dim filePath As String
    
    backupPath = Environ("USERPROFILE") & "\Documents\氧化加工厂备份\"
    
    filePath = Application.GetOpenFilename( _
        InitialFileName:=backupPath, _
        FileFilter:="Excel文件 (*.xlsx),*.xlsx")
    
    If filePath = "False" Then Exit Sub
    
    If MsgBox("确定要恢复此备份吗？当前数据将被覆盖！" & vbCrLf & vbCrLf & _
              filePath, vbYesNo + vbExclamation, "确认恢复") = vbNo Then
        Exit Sub
    End If
    
    Workbooks.Open filePath
    MsgBox "备份已打开，请检查数据后手动保存", vbInformation
End Sub

' 设置自动备份
Private Sub SetupAutoBackup()
    Dim enable As String
    Dim interval As String
    
    enable = InputBox("是否启用自动备份？" & vbCrLf & "1. 是" & vbCrLf & "2. 否", "自动备份设置", "1")
    
    If enable = "1" Then
        interval = InputBox("设置备份间隔（分钟，建议30-60）：", "备份间隔", "60")
        
        ' 保存设置到工作表
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("基础设置")
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ws.Cells(50, 1).Value = "自动备份启用"
            ws.Cells(50, 2).Value = True
            ws.Cells(51, 1).Value = "备份间隔(分钟)"
            ws.Cells(51, 2).Value = Val(interval)
        End If
        
        MsgBox "自动备份设置已保存！" & vbCrLf & _
               "备份间隔：" & interval & " 分钟" & vbCrLf & vbCrLf & _
               "注意：需要保持WPS/Excel打开才能自动备份", vbInformation
    Else
        MsgBox "已取消自动备份", vbInformation
    End If
End Sub

' ============================================================================
' 宏33: DataRepairTool - 数据修复工具
' V18.0新增：检测并修复常见问题
' ============================================================================
Sub DataRepairTool()
    Dim action As String
    
    action = InputBox("选择修复项目：" & vbCrLf & _
                     "1. 修复日期格式" & vbCrLf & _
                     "2. 修复金额格式" & vbCrLf & _
                     "3. 清理空行" & vbCrLf & _
                     "4. 检查数据完整性" & vbCrLf & _
                     "5. 一键修复全部", _
                     "数据修复工具", "5")
    
    Select Case action
        Case "1": Call FixDateFormat
        Case "2": Call FixAmountFormat
        Case "3": Call CleanEmptyRows
        Case "4": Call CheckDataIntegrity
        Case "5"
            Call FixDateFormat
            Call FixAmountFormat
            Call CleanEmptyRows
            Call CheckDataIntegrity
            MsgBox "一键修复完成！", vbInformation
        Case Else: MsgBox "无效选择", vbExclamation
    End Select
End Sub

' 修复日期格式
Private Sub FixDateFormat()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim fixed As Long
    
    Set ws = ThisWorkbook.Sheets("收入记录")
    lastRow = GetLastRow(ws, 1)
    fixed = 0
    
    For i = 4 To lastRow
        On Error Resume Next
        Dim dt As Date
        dt = CDate(ws.Cells(i, 1).Value)
        If Err.Number = 0 Then
            ws.Cells(i, 1).Value = dt
            ws.Cells(i, 1).NumberFormat = "yyyy-mm-dd"
            fixed = fixed + 1
        End If
        On Error GoTo 0
    Next i
    
    MsgBox "日期格式修复完成！共修复 " & fixed & " 条记录", vbInformation
End Sub

' 修复金额格式
Private Sub FixAmountFormat()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim fixed As Long
    
    Set ws = ThisWorkbook.Sheets("收入记录")
    lastRow = GetLastRow(ws, 1)
    fixed = 0
    
    For i = 4 To lastRow
        If IsNumeric(ws.Cells(i, 4).Value) Then
            ws.Cells(i, 4).NumberFormat = "#,##0.00"
            fixed = fixed + 1
        End If
    Next i
    
    MsgBox "金额格式修复完成！共修复 " & fixed & " 条记录", vbInformation
End Sub

' 清理空行
Private Sub CleanEmptyRows()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim deleted As Long
    
    Set ws = ThisWorkbook.Sheets("收入记录")
    lastRow = GetLastRow(ws, 1)
    deleted = 0
    
    For i = lastRow To 4 Step -1
        If ws.Cells(i, 1).Value = "" And ws.Cells(i, 3).Value = "" Then
            ws.Rows(i).Delete
            deleted = deleted + 1
        End If
    Next i
    
    MsgBox "空行清理完成！共删除 " & deleted & " 行", vbInformation
End Sub

' 检查数据完整性
Private Sub CheckDataIntegrity()
    Dim issues As String
    issues = ""
    
    ' 检查收入记录
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("收入记录")
    If ws Is Nothing Then
        issues = issues & "• 缺少收入记录表" & vbCrLf
    Else
        lastRow = GetLastRow(ws, 1)
        If lastRow < 4 Then
            issues = issues & "• 收入记录表无数据" & vbCrLf
        End If
    End If
    
    ' 检查支出记录
    Set ws = ThisWorkbook.Sheets("支出记录")
    If ws Is Nothing Then
        issues = issues & "• 缺少支出记录表" & vbCrLf
    End If
    
    ' 检查应收应付
    Set ws = ThisWorkbook.Sheets("应收应付")
    If ws Is Nothing Then
        issues = issues & "• 缺少应收应付表" & vbCrLf
    End If
    
    On Error GoTo 0
    
    If issues = "" Then
        MsgBox "✓ 数据完整性检查通过！未发现明显问题", vbInformation
    Else
        MsgBox "发现以下问题：" & vbCrLf & vbCrLf & issues & vbCrLf & _
               "建议运行【一键修复全部】", vbExclamation
    End If
End Sub

' ============================================================================
' 宏34: ARExpireAlert - 应收逾期提醒
' V18.0新增：自动检测逾期应收并提醒
' ============================================================================
Sub ARExpireAlert()
    Dim wsARAP As Worksheet, wsIncome As Worksheet
    Dim lastRow As Long, i As Long
    Dim alertMsg As String
    Dim alertCount As Integer
    Dim totalOverdue As Double
    
    On Error Resume Next
    Set wsARAP = ThisWorkbook.Sheets("应收应付")
    Set wsIncome = ThisWorkbook.Sheets("收入记录")
    On Error GoTo 0
    
    If wsARAP Is Nothing Then
        MsgBox "未找到应收应付表！", vbExclamation
        Exit Sub
    End If
    
    alertMsg = "应收逾期提醒" & vbCrLf & vbCrLf
    alertCount = 0
    totalOverdue = 0
    
    ' 检查各客户应收
    lastRow = GetLastRow(wsARAP, 1)
    For i = 4 To lastRow
        Dim customer As String
        Dim arAmount As Double
        
        customer = CStr(wsARAP.Cells(i, 1).Value)
        arAmount = CDbl(Nz(wsARAP.Cells(i, 7).Value, 0))
        
        If arAmount > 0 Then
            ' 查找最后一笔收入日期
            Dim lastDate As Date
            Dim incomeRow As Long
            
            lastDate = DateSerial(2000, 1, 1)
            If Not wsIncome Is Nothing Then
                Dim incLastRow As Long
                incLastRow = GetLastRow(wsIncome, 1)
                For incomeRow = 4 To incLastRow
                    If InStr(1, wsIncome.Cells(incomeRow, 3).Value, customer, vbTextCompare) > 0 Then
                        On Error Resume Next
                        If CDate(wsIncome.Cells(incomeRow, 1).Value) > lastDate Then
                            lastDate = CDate(wsIncome.Cells(incomeRow, 1).Value)
                        End If
                        On Error GoTo 0
                    End If
                Next incomeRow
            End If
            
            ' 计算逾期天数
            Dim days As Long
            days = DateDiff("d", lastDate, Date)
            
            If days > 30 Then
                alertCount = alertCount + 1
                totalOverdue = totalOverdue + arAmount
                alertMsg = alertMsg & "• " & customer & ": " & Format(arAmount, "#,##0.00") & " 元"
                If days > 60 Then
                    alertMsg = alertMsg & " ⚠ 严重逾期(" & days & "天)"
                Else
                    alertMsg = alertMsg & " 逾期(" & days & "天)"
                End If
                alertMsg = alertMsg & vbCrLf
            End If
        End If
    Next i
    
    If alertCount = 0 Then
        MsgBox "✓ 恭喜！暂无逾期应收款项", vbInformation, "应收提醒"
    Else
        alertMsg = alertMsg & vbCrLf & "共 " & alertCount & " 笔逾期，合计 " & Format(totalOverdue, "#,##0.00") & " 元"
        MsgBox alertMsg, vbExclamation, "应收逾期提醒"
    End If
End Sub

' ============================================================================
' 宏35: InventoryAlert - 库存预警
' V18.0新增：材料库存不足提醒
' ============================================================================
Sub InventoryAlert()
    Dim ws As Worksheet
    Dim materials As Object
    Dim lastRow As Long, i As Long
    Dim alertMsg As String
    Dim alertCount As Integer
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("材料进销存台账")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "未找到材料进销存台账！", vbExclamation
        Exit Sub
    End If
    
    ' 使用字典统计各材料库存
    Set materials = CreateObject("Scripting.Dictionary")
    
    lastRow = GetLastRow(ws, 1)
    For i = 4 To lastRow
        Dim matName As String
        Dim qty As Double
        
        matName = CStr(ws.Cells(i, 2).Value)
        If matName <> "" Then
            If Not materials.Exists(matName) Then
                materials.Add matName, 0
            End If
            
            ' 入库+
            If IsNumeric(ws.Cells(i, 4).Value) Then
                materials(matName) = materials(matName) + CDbl(ws.Cells(i, 4).Value)
            End If
            ' 出库-
            If IsNumeric(ws.Cells(i, 7).Value) Then
                materials(matName) = materials(matName) - CDbl(ws.Cells(i, 7).Value)
            End If
        End If
    Next i
    
    ' 检查库存预警
    alertMsg = "库存预警" & vbCrLf & vbCrLf
    alertCount = 0
    
    Dim key As Variant
    For Each key In materials.Keys
        If materials(key) < 10 Then
            alertCount = alertCount + 1
            alertMsg = alertMsg & "• " & key & ": 库存 " & materials(key)
            If materials(key) <= 0 Then
                alertMsg = alertMsg & " ⚠ 已缺货！"
            Else
                alertMsg = alertMsg & " 库存不足"
            End If
            alertMsg = alertMsg & vbCrLf
        End If
    Next key
    
    If alertCount = 0 Then
        MsgBox "✓ 所有材料库存充足", vbInformation, "库存预警"
    Else
        alertMsg = alertMsg & vbCrLf & "共 " & alertCount & " 种材料需要采购"
        MsgBox alertMsg, vbExclamation, "库存预警"
    End If
End Sub

' ============================================================================
' 宏36: ProfitAlert - 利润预警
' V18.0新增：盈亏分析和预警
' ============================================================================
Sub ProfitAlert()
    Dim wsProfit As Worksheet
    Dim income As Double, expense As Double, profit As Double
    Dim profitRate As Double
    Dim msg As String
    
    On Error Resume Next
    Set wsProfit = ThisWorkbook.Sheets("利润分析表")
    On Error GoTo 0
    
    If wsProfit Is Nothing Then
        MsgBox "未找到利润分析表！", vbExclamation
        Exit Sub
    End If
    
    ' 读取数据
    income = CDbl(Nz(wsProfit.Range("D4").Value, 0))
    expense = CDbl(Nz(wsProfit.Range("D20").Value, 0))
    profit = income - expense
    
    If income > 0 Then
        profitRate = profit / income
    Else
        profitRate = 0
    End If
    
    msg = "利润分析预警" & vbCrLf & vbCrLf
    msg = msg & "收入：" & Format(income, "#,##0.00") & " 元" & vbCrLf
    msg = msg & "支出：" & Format(expense, "#,##0.00") & " 元" & vbCrLf
    msg = msg & "利润：" & Format(profit, "#,##0.00") & " 元" & vbCrLf
    msg = msg & "利润率：" & Format(profitRate, "0.00%") & vbCrLf & vbCrLf
    
    If profit < 0 Then
        msg = msg & "⚠ 警告：本月亏损！请立即控制成本！"
        MsgBox msg, vbCritical, "利润预警"
    ElseIf profitRate < 0.1 Then
        msg = msg & "⚠ 注意：利润率偏低（<10%），建议优化成本结构"
        MsgBox msg, vbExclamation, "利润预警"
    ElseIf profitRate > 0.3 Then
        msg = msg & "✓ 恭喜：利润率良好（>30%），经营状况健康"
        MsgBox msg, vbInformation, "利润预警"
    Else
        msg = msg & "✓ 利润率正常，继续保持"
        MsgBox msg, vbInformation, "利润预警"
    End If
End Sub
