Attribute VB_Name = "CalStoryStaffInput"
' ˵��:
' һ����ִ�������� CaluStoryStaffInput ǰ�Ա����Ҫ�����Ų飬���Ƿ��������Ҫ��
' 1���Ƿ������������ֽ�������,���û��������ͬ�������������
' 2����һ��(CNST_FIRST_ROW)�Ƿ�����Ч��
' 3���˵������Ƿ�������ж����һ�£����Sub initJobTypes()�еĳ�ʼ����
' ��������Դ������Ϣ�������²���const����
' ���������Ҫ�༭�˵ķ��࣬��Ҫ�޸� CNST_JOB_TYPE_COUNT�Լ�initJobTypes

Const CNST_SRC_SHEET = "IM��Ŀ20170303"     'Դsheet����
Const CNST_TRG_SHEET = "Sheet8"             'Ŀ��sheet����
Const CNST_FIRST_ROW As Integer = 71        '��ʼ��
Const CNST_LAST_ROW As Integer = 89         '������
Const CNST_STORY_NAME_COLUMN = "B"          '�����������ֵ�λ��
Const CNST_JOB_TYPE_COLUMN = "G"            '���ö˵�λ��
Const CNST_SIGNED_COLUMN = "N"              '�����Ƿ���ǩ�ֵ�λ��
Const CNST_TQDNONF_COLUMN = "D"             '�����Ƿ��Ƿ���TQD��NONFλ��
Const CNST_BEGING_DATE_COLUMN = "AT"        '���ÿ�ʼ��ʱ��λ��
Const CNST_WEEKS = 2                        '�ۼƼ���

Dim storyNameColumn As Integer
Dim jobTypeColumn As Integer
Dim beginDateColumn As Integer
Dim signedColumn As Integer
Dim tqdnonfColumn As Integer

Const CNST_JOB_TYPE_COUNT = 8               '�˵����� ���һ���ǡ�������
Dim CNST_JOB_TYPES(CNST_JOB_TYPE_COUNT - 1) As String

'���������࣬�������������صķ���
Private Type CStory
    strTQDNonf As String
    strName As String
    strSigned As String
    dJobTypeDays(CNST_JOB_TYPE_COUNT - 1) As Double
    allJobTypeDays As Double
End Type

Dim storyList_() As CStory '''ȫ�ֱ������������б�

'˵����
'1����������ƣ�����еĶ�����������Ҫһֱ��������ֲ�һ��ʱ�䶼����ܵ����������
'2�����һ�ֶ�����CNST_JOB_TYPE_COUNTҪ��1����Ҫ��CNST_JOB_TYPES����ʼֵ
Sub initJobTypes()
    CNST_JOB_TYPES(0) = "�ܹ�"
    CNST_JOB_TYPES(1) = "WEB���"
    CNST_JOB_TYPES(2) = "PC��"
    CNST_JOB_TYPES(3) = "U3D"
    CNST_JOB_TYPES(4) = "��׿"
    CNST_JOB_TYPES(5) = "iOS"
    CNST_JOB_TYPES(6) = "webǰ��"
    
    
    CNST_JOB_TYPES(CNST_JOB_TYPE_COUNT - 1) = "����"
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
    Set srcSheet = Worksheets(CNST_SRC_SHEET)       'Դ���

    With storyList_(iStoryIndex)
        .strName = srcSheet.Cells(iRow, storyNameColumn).Value
        .strSigned = srcSheet.Cells(iRow, signedColumn).Value
        .strTQDNonf = srcSheet.Cells(iRow, tqdnonfColumn).Value
    End With
End Sub

' ע��:
' ��ִ�к���ǰ�Ա���checklist
' 1���Ƿ������������ֽ�������
' 2����һ��(CNST_FIRST_ROW)�Ƿ�����Ч��
Sub CaluStoryStaffInput()
    Call initJobTypes
    
    Set srcSheet = Worksheets(CNST_SRC_SHEET)       'Դ���
    Set tgtSheet = Worksheets(CNST_TRG_SHEET)       'Ŀ����
    
    Dim storyList() As String                       '���������б�
    ReDim storyList(0 To 0)
    Dim storyJobTypeDays() As Double                '����ĳ���������������Ĺ���ʱ��
    ReDim storyJobTypeDays(0 To CNST_JOB_TYPE_COUNT - 1, 0 To 0)
    
    ReDim storyList_(0 To 0) '''
    
    storyNameColumn = srcSheet.Cells(1, CNST_STORY_NAME_COLUMN).Column      '�����������ֵ�λ��
    jobTypeColumn = srcSheet.Cells(1, CNST_JOB_TYPE_COLUMN).Column          '���ö˵�λ��
    beginDateColumn = srcSheet.Cells(1, CNST_BEGING_DATE_COLUMN).Column     '���ÿ�ʼ��ʱ��λ��
    signedColumn = srcSheet.Cells(1, CNST_SIGNED_COLUMN).Column             '�����Ƿ���ǩ��λ��
    tqdnonfColumn = srcSheet.Cells(1, CNST_TQDNONF_COLUMN).Column           '������������TQDNonf��λ��
    
    
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
    

    '������
    ''�������
    Dim iWColumn, iWRow, i, j As Integer
    iWColumn = 1
    iWRow = 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "��Ŀ����"
    iWColumn = iWColumn + 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "��������"
    iWColumn = iWColumn + 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "���ȼ�"
    iWColumn = iWColumn + 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "�������ݣ�����������"
    iWColumn = iWColumn + 1
    tgtSheet.Cells(iWRow, iWColumn).Value = "�汾�ƻ��Ƿ���ǩ"
    iWColumn = iWColumn + 1
    ''�������
    For i = 0 To CNST_JOB_TYPE_COUNT - 1
        tgtSheet.Cells(iWRow, iWColumn + i).Value = CNST_JOB_TYPES(i)
    Next i
    ''���������Ϣ
    tgtSheet.Cells(iWRow, iWColumn + CNST_JOB_TYPE_COUNT).Value = srcSheet.Cells(1, beginDateColumn).Value + " (" + Str(CNST_WEEKS) + ")"
    For i = 0 To storyIndex
        If storyList_(i).allJobTypeDays <> 0 Then
            iWRow = iWRow + 1
            iWColumn = 1
            tgtSheet.Cells(iWRow, iWColumn) = "IM+����Ƶ"
            iWColumn = iWColumn + 1
            tgtSheet.Cells(iWRow, iWColumn) = storyList_(i).strTQDNonf
            iWColumn = iWColumn + 1
            tgtSheet.Cells(iWRow, iWColumn) = "��"
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
        End If '�ų���ʱ��Ϊ0������
    Next i
    
End Sub
