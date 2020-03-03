Attribute VB_Name = "导出为华为EMUI名片"
Option Explicit
Sub ExportVCFForEMUI()
'''导出为EMUI名片vcf格式
'''变量准备
Dim FSO
Dim ADSM, ADSM2
Dim FilePath, FileName
Dim TheTable, CtrlLogo As String
Dim LastContactCell As Range
Dim RelationType
Set FSO = CreateObject("scripting.filesystemobject")
Set ADSM = CreateObject("ADODB.Stream")
Set ADSM2 = CreateObject("ADODB.Stream")
FilePath = Range("功能设置!C2")
FileName = Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
TheTable = "基本联系信息"
ADSM.Type = 2
ADSM.Charset = "UTF-8"
ADSM.Open
Dim i, j
For i = 1 To Sheets.Item(TheTable).UsedRange.Rows.Count   '注意sheet对象，目前还不知道怎么以row对象为遍历单位
CtrlLogo = Sheets.Item(TheTable).Cells(i, 2)
Select Case CtrlLogo
Case Range("功能设置!C1")
ADSM.writetext "BEGIN:VCARD", 1
ADSM.writetext "VERSION:3.0", 1
If Range(TheTable & "!" & "C" & i) <> "" Or Range(TheTable & "!" & "D" & i) <> "" Then    '  确定姓名两栏不全为空
ADSM.writetext "FN:" & Range(TheTable & "!" & "C" & i) & Range(TheTable & "!" & "D" & i), 1    '写入FN字段用姓名作为值
ADSM.writetext "N:" & Range(TheTable & "!" & "C" & i) & ";" & Range(TheTable & "!" & "D" & i), 1  '写入姓名N字段
If Range(TheTable & "!" & "J" & i) <> "" Then ADSM.writetext "TEL;TYPE=CELL:" & Range(TheTable & "!" & "J" & i), 1 '写入移动电话
If Range(TheTable & "!" & "M" & i) <> "" Then ADSM.writetext "TEL;TYPE=WORK:" & Range(TheTable & "!" & "M" & i), 1 '写入工作电话，但此处仅是职务值
If Range(TheTable & "!" & "L" & i) <> "" Then ADSM.writetext "TEL;TYPE=HOME:" & Range(TheTable & "!" & "L" & i), 1 '写入住宅电话
If Range(TheTable & "!" & "O" & i) <> "" Then ADSM.writetext "EMAIL;TYPE=WORK:" & Range(TheTable & "!" & "O" & i), 1 '写入办公邮件
If Range(TheTable & "!" & "N" & i) <> "" Then ADSM.writetext "EMAIL;TYPE=HOME:" & Range(TheTable & "!" & "N" & i), 1 '写入个人邮件
If Range(TheTable & "!" & "G" & i) <> "" Then ADSM.writetext "ORG:" & Range(TheTable & "!" & "G" & i) & Range(TheTable & "!" & "H" & i), 1  '写入组织，包括单位和部门
If Range(TheTable & "!" & "I" & i) <> "" Then ADSM.writetext "TITLE:" & Range(TheTable & "!" & "I" & i), 1  '写入头衔，使用职务栏I
If Range(TheTable & "!" & "P" & i) <> "" Then ADSM.writetext "ADR;TYPE=HOME:;;" & Range(TheTable & "!" & "P" & i) & ";;;;", 1  '写入家庭住址栏，用P列
If Range(TheTable & "!" & "Q" & i) <> "" Then ADSM.writetext "ADR;TYPE=WORK:;;" & Range(TheTable & "!" & "Q" & i) & ";;;;", 1  '写入家庭住址栏，用Q列
If Range(TheTable & "!" & "R" & i) <> "" Then ADSM.writetext "BDAY:" & Format(Range(TheTable & "!" & "R" & i), "yyyy-m-d"), 1  '写入生日，用R列
If Range(TheTable & "!" & "S" & i) <> "" Then ADSM.writetext "URL:" & Range(TheTable & "!" & "S" & i), 1   '写入网站，使用S列，可以使用Mailto:等字头
If Range(TheTable & "!" & "K" & i) <> "" Then ADSM.writetext "X-QQ:" & Range(TheTable & "!" & "K" & i), 1   '写入QQ，小米能自动启动QQ程序
If Range(TheTable & "!" & "T" & i) <> "" Then ADSM.writetext "X-ANDROID-CUSTOM:vnd.android.cursor.item/contact_event;" & Range("详细信息!" & "T" & i) & ";1;（1则显示周年纪念日且可运行）;;;;;;;;;;;;", 1  '写入周年纪念日，使用详细信息T列
If Range(TheTable & "!" & "U" & i) <> "" Then ADSM.writetext "X-ANDROID-CUSTOM:vnd.android.cursor.item/contact_event;" & Range("详细信息!" & "U" & i) & ";3;（1则显示周年纪念日且可运行）;;;;;;;;;;;;", 1  '写入周年纪念日，使用详细信息U列
If Range(TheTable & "!" & "V" & i) <> "" Then ADSM.writetext "X-ANDROID-CUSTOM:vnd.android.cursor.item/contact_event;" & Range("详细信息!" & "V" & i) & ";0;（1则显示周年纪念日且可运行）;;;;;;;;;;;;", 1  '写入周年纪念日，使用详细信息V列
For j = 1 To Sheets.Item("关系链").UsedRange.Rows.Count
Select Case Range("关系链!C" & j).Value
Case "助理", "秘书"
RelationType = 1
Case "兄弟"
RelationType = 2
Case "子女", "儿子", "女儿"
RelationType = 3
Case "情人", "恋人", "伴侣"
RelationType = 4
Case "父亲", "爸爸"
RelationType = 5
Case "朋友", "密友"
RelationType = 6
Case "上司", "老板", "主管"
RelationType = 7
Case "母亲", "妈妈"
RelationType = 8
Case "父母"
RelationType = 9
Case "合作伙伴", "合伙人", "同事"
RelationType = 10
Case "介绍人"
RelationType = 11
Case "亲属"
RelationType = 12
Case "姐妹", "姐姐", "妹妹"
RelationType = 13
Case "配偶", "老公", "老婆", "夫人", "媳妇"
RelationType = 14
Case Else
RelationType = 0
End Select
If Range("关系链!A" & j) = Range(TheTable & "!" & "A" & i).Value Then ADSM.writetext "X-ANDROID-CUSTOM:vnd.android.cursor.item/relation;" & Range("关系链!D" & j) & ";" & RelationType & ";" & Range("关系链!C" & j) & ";;;;;;;;;;;;", 1 '遍历关系链表并写入此人的关系人
Next

'   Set LastContactCell = Range("关系链!A1")    '这段代码用于使用range.find函数来查找符合条件的单元格，发现还不如for循环来得简洁。想想算法根基上也是一样要遍历的，那就不如直接for了
'   Set FirstContactCell = Nothing
'   Do
'      Set LastContactCell = Range("关系链!A:A").Find(Range(TheTable & "!" & "A" & i).Value, LastContactCell, xlValues, xlWhole, xlByRows, xlNext)
'      If FirstContactCell Is Nothing Then Set FirstContactCell = LastContactCell
'      If Not LastContactCell Is Nothing Then Debug.Print LastContactCell.Address
'   Loop Until FirstContactCell Is Range("关系链!A:A").FindNext(LastContactCell)
End If
ADSM.writetext "END:VCARD", 1
Range(TheTable & "!" & "B" & i) = ""  '删除生成标记m
End Select
Next
ADSM.Position = 3
ADSM2.Type = 1
ADSM2.Open
ADSM.copyto ADSM2
ADSM2.savetofile FilePath & FileName & ".vcf", 2   '将adsm2写入到文件
Debug.Print ("ok")
End Sub
