Attribute VB_Name = "CalStoryStaffInput"
' 说明:
' 一、在执行主函数 CaluStoryStaffInput 前对表格需要进行排查，看是否符合以下要求
' 1、是否根据需求的名字进行排序,如果没有排序，相同的需求会有两行
' 2、第一行(CNST_FIRST_ROW)是否是有效行
' 3、端的类型是否与程序中定义的一致（详见Sub initJobTypes()中的初始化）
' 二、根据源表格的信息设置以下参数const参数
' 三、如果需要编辑端的分类，需要修改 CNST_JOB_TYPE_COUNT以及initJobTypes

Const CNST_SRC_SHEET = "IM项目20170303"     '源sheet名字
Const CNST_TRG_SHEET = "Sheet8"             '目标sheet名字
Const CNST_FIRST_ROW As Integer = 71        '开始行
Const CNST_LAST_ROW As Integer = 89         '结束行
Const CNST_STORY_NAME_COLUMN = "B"          '设置需求名字的位置
Const CNST_JOB_TYPE_COLUMN = "G"            '设置端的位置
Const CNST_SIGNED_COLUMN = "N"              '设置是否已签字的位置
Const CNST_TQDNONF_COLUMN = "D"             '设置是否是否是TQD或NONF位置
Const CNST_BEGING_DATE_COLUMN = "AT"        '设置开始的时间位置
Const CNST_WEEKS = 2                        '累计几周

Dim storyNameColumn As Integer
Dim jobTypeColumn As Integer
Dim beginDateColumn As Integer
Dim signedColumn As Integer
Dim tqdnonfColumn As Integer

Const CNST_JOB_TYPE_COUNT = 8               '端的数量 最后一个是“其他”
Dim CNST_JOB_TYPES(CNST_JOB_TYPE_COUNT - 1) As String

'定义需求类，存放与需求类相关的方法
Private Type CStory
    strTQDNonf As String
    strName As String
    strSigned As String
    dJobTypeDays(CNST_JOB_TYPE_COUNT - 1) As Double
    allJobTypeDays As Double
End Type

Dim storyList_() As CStory '''全局变量保存需求列表

'说明：
'1、定义端名称，表格中的端名称与这里要一直，如果名字不一致时间都会汇总到“其他”里，
'2、添加一种端类型CNST_JOB_TYPE_COUNT要加1，且要给CNST_JOB_TYPES赋初始值
Sub initJobTypes()
    CNST_JOB_TYPES(0) = "架构"
    CNST_JOB_TYPES(1) = "WEB后端"
    CNST_JOB_TYPES(2) = "PC端"
    CNST_JOB_TYPES(3) = "U3D"
    CNST_JOB_TYPES(4) = "安卓"
    CNST_JOB_TYPES(5) = "iOS"
    CNST_JOB_TYPES(6) = "web前端"
    
    
    CNST_JOB_TYPES(CNST_JOB_TYPE_COUNT - 1) = "其他"
End Sub

Function getJobTypeIndex(strJobName As String, CNST_JOB_TYPES() As String) As Integer
    Dim i As Integer
    Dim index As Integer
    index = CNST_JOB_TYPE_COUNT - 1
    For i = LBound(CNST_JOB_TYPES) To UBound(CNST_JOB_TYPES)
        If UCase(strJobName) = UCase(CNST_JOB_TYPES(i)) Then
            index = i
            Exit For
        End If
    Next i
    getJobTypeIndex = index
End Function


Private Sub addStory_(storyName As String) '''
    Dim maxIndex As Integer
    maxIndex = UBound(storyList_)
    maxIndex = maxIndex + 1
    ReDim Preserve storyList_(0 To maxIndex)
    storyList_(maxIndex).strName = storyName
End Sub

Sub addStory(storyName As String, storyList() As String, storyJobTypeDays() As Double)
    Dim maxIndex As Integer
    maxIndex = UBound(storyList)
    maxIndex = maxIndex + 1
    ReDim Preserve storyList(0 To maxIndex)
    ReDim Preserve storyJobTypeDays(0 To CNST_JOB_TYPE_COUNT - 1, 0 To maxIndex)
    storyList(maxIndex) = storyName
End Sub

Sub setJobTypeDays(storyNameIndex As Integer, jobTypeName As String, jobTypeDays As Double, storyJobTypeDays() As Double)
    Dim jobTypeIndex As Integer
    jobTypeIndex = getJobTypeIndex(jobTypeName, CNST_JOB_TYPES)
    storyJobTypeDays(jobTypeIndex, storyNameIndex) = jobTypeDays
End Sub

Private Sub setJobTypeDays_(jobTypeName As String, jobTypeDays As Double, aStory As CStory) '''obTypeDays As Double,
    Dim jobTypeIndex As Integer
    jobTypeIndex = getJobTypeIndex(jobTypeName, CNST_JOB_TYPES)
    aStory.dJobTypeDays(jobTypeIndex) = jobTypeDays
End Sub

Private Sub sumAllJobTypeDays(aStory As CStory)
    Dim sumDays As Double
    Dim i As Integer
    sumDays = 0
    
    For i = 0 To CNST_JOB_TYPE_COUNT - 1
        sumDays = sumDays + aStory.dJobTypeDays(i)
    Next i
    aStory.allJobTypeDays = sumDays
End Sub


Function sumJobTypeDays(nowRow As Integer, firstCoulumn As Integer) As Double
    Dim sumDays As Double
    sumDays = 0
    Dim i As Integer
    For i = 0 To CNST_WEEKS - 1
        sumDays = sumDays + Worksheets(CNST_SRC_SHEET).Cells(nowRow, firstCoulumn + i).Value
    Next i
    sumJobTypeDays = sumDays
End Function

Sub setStoryParamsByCell(iStoryIndex As Integer, iRow As Integer)
    Set srcSheet = Worksheets(CNST_SRC_SHEET)       '源表格

    With storyList_(iStoryIndex)
        .strName = srcSheet.Cells(iRow, storyNameColumn).Value
        .strSigned = srcSheet.Cells(iRow, signedColumn).Value
        .strTQDNonf = srcSheet.Cells(iRow, tqdnonfColumn).Value
    End With
End Sub

' 注意:
' 在执行函数前对表格的checklist
' 1、是否根据需求的名字进行排序
' 2、第一行(CNST_FIRST_ROW)是否是有效行
Sub CaluStoryStaffInput()
    Call initJobTypes
    
    Set srcSheet = Worksheets(CNST_SRC_SHEET)       '源表格
    Set tgtSheet = Worksheets(CNST_TRG_SHEET)       '目标表格
    
    Dim storyList() As String                       '保存需求列表
    ReDim storyList(0 To 0)
    Dim storyJobTypeDays() As Double                '保存某个需求各工作分类的工作时间
    ReDim storyJobTypeDays(0 To CNST_JOB_TYPE_COUNT - 1, 0 To 0)
    
    ReDim storyList_(0 To 0) '''
    
    storyNameColumn = srcSheet.Cells(1, CNST_STORY_NAME_COLUMN).Column      '设置需求名字的位置
    jobTypeColumn = srcSheet.Cells(1, CNST_JOB_TYPE_COLUMN).Column          '设置端的位置
    beginDateColumn = srcSheet.Cells(1, CNST_BEGING_DATE_COLUMN).Column     '设置开始的时间位置
    signedColumn = srcSheet.Cells(1, CNST_SIGNED_COLUMN).Column             '设置是否已签的位置
    tqdnonfColumn = srcSheet.Cells(1, CNST_TQDNONF_COLUMN).Column           '设置需求类型TQDNonf的位置
    
    
    Dim storyIndex As Integer
    storyIndex = 0
    Dim storyName As String
    storyName = srcSheet.Cells(CNST_FIRST_ROW, storyNameColumn).Value
    Dim jobTypeName As String
    jobTypeName = srcSheet.Cells(CNST_FIRST_ROW, jobTypeColumn).Value
    storyList(storyIndex) = storyName
    
    Call setStoryParamsByCell(storyIndex, CNST_FIRST_ROW)
    
    Dim jobTypeDays As Double
    jobTypeDays = sumJobTypeDays(CNST_FIRST_ROW, beginDateColumn)
    Call setJobTypeDays(storyIndex, jobTypeName, jobTypeDays, storyJobTypeDays)
    
    Call setJobTypeDays_(jobTypeName, jobTypeDays, storyList_(storyIndex)) '''
    
    Dim rowStoryName As String
    Dim iRow As Integer
    For iRow = CNST_FIRST_ROW + 1 To CNST_LAST_ROW
        rowStoryName = srcSheet.Cells(iRow, storyNameColumn).Value
        jobTypeName = srcSheet.Cells(iRow, jobTypeColumn).Value
        jobTypeDays = sumJobTypeDays(iRow, beginDateColumn)
        If (rowStoryName = storyName) Then
           
        Else
            storyIndex = storyIndex + 1
            storyName = rowStoryName
            Call addStory(storyName, storyList, storyJobTypeDays)
            
            Call addStory_(storyName) '''
            Call setStoryParamsByCell(storyIndex, iRow)

        End If
        Call setJobTypeDays(storyIndex, jobTypeName, jobTypeDays, storyJobTypeDays)
        
        Call setJobTypeDays_(jobTypeName, jobTypeDays, storyList_(storyIndex)) '''
        
        Call sumAllJobTypeDays(storyList_(storyIndex))
    Next iRow
    

    '输出结果
    ''输出标题
    Dim iWColumn, iWRow, i, j As Integer
    iWColumn = 1
    iWRow = 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "项目名称"
    iWColumn = iWColumn + 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "任务类型"
    iWColumn = iWColumn + 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "优先级"
    iWColumn = iWColumn + 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "任务内容（需求描述）"
    iWColumn = iWColumn + 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "版本计划是否已签"
    iWColumn = iWColumn + 1
    ''输出内容
    For i = 0 To CNST_JOB_TYPE_COUNT - 1
        tgtSheet.Cells(iWRow, iWColumn + i).Value = CNST_JOB_TYPES(i)
    Next i
    ''输出调试信息
    tgtSheet.Cells(iWRow, iWColumn + CNST_JOB_TYPE_COUNT).Value = srcSheet.Cells(1, beginDateColumn).Value + " (" + Str(CNST_WEEKS) + ")"
    For i = 0 To storyIndex
        If storyList_(i).allJobTypeDays <> 0 Then
            iWRow = iWRow + 1
            iWColumn = 1
            tgtSheet.Cells(iWRow, iWColumn) = "IM+音视频"
            iWColumn = iWColumn + 1
            tgtSheet.Cells(iWRow, iWColumn) = storyList_(i).strTQDNonf
            iWColumn = iWColumn + 1
            tgtSheet.Cells(iWRow, iWColumn) = "高"
            iWColumn = iWColumn + 1
            tgtSheet.Cells(iWRow, iWColumn) = storyList_(i).strName
            iWColumn = iWColumn + 1
            tgtSheet.Cells(iWRow, iWColumn) = storyList_(i).strSigned
            iWColumn = iWColumn + 1
            
            '''tgtSheet.Cells(iWRow, iWColumn).Value = storyList(i)
            For j = 0 To CNST_JOB_TYPE_COUNT - 1
               '''tgtSheet.Cells(iWRow, iWColumn + j) = storyJobTypeDays(j, i)
               tgtSheet.Cells(iWRow, iWColumn + j) = storyList_(i).dJobTypeDays(j) '''
            Next j
        End If '排除调时间为0的需求
    Next i
    
End Sub
