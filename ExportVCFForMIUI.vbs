'the key is Vcard files in MIUI is built by UTF-8 wiht no BOM
'the other key is that the file export from you MIUI mobile is imcomplete,but is complete on  two-dimension code on screen.you can use another device to scan this  two-dimension code.
Sub ExportVCFForMIUI()
'''变量准备
Set fso = CreateObject("scripting.filesystemobject")
Set ADSM = CreateObject("ADODB.Stream")
Set ADSM2 = CreateObject("ADODB.Stream")
FilePath = Range("功能设置!C2")
FileName = Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
Dim TheTable, CtrlLogo As String
TheTable = "基本联系信息"
ADSM.Type = 2
ADSM.Charset = "UTF-8"
ADSM.Open
Dim i
For i = 1 To Sheets.Item(TheTable).UsedRange.Rows.Count   '注意sheet对象，目前还不知道怎么以row对象为遍历单位
  CtrlLogo = Sheets.Item(TheTable).Cells(i, 2)
  Select Case CtrlLogo
  Case Range("功能设置!C1")
    ADSM.writetext "BEGIN: VCARD", 1
    ADSM.writetext "VERSION:2.1", 1
    If Range(TheTable & "!" & "C" & i) <> "" Or Range(TheTable & "!" & "D" & i) <> "" Then    '  确定姓名两栏不全为空
    ADSM.writetext "FN:" & Range(TheTable & "!" & "C" & i) & Range(TheTable & "!" & "D" & i), 1    '写入FN字段用姓名作为值
    ADSM.writetext "N:" & Range(TheTable & "!" & "C" & i) & ";" & Range(TheTable & "!" & "D" & i), 1  '写入姓名N字段
    If Range(TheTable & "!" & "J" & i) <> "" Then ADSM.writetext "TEL;TYPE=CELL:" & Range(TheTable & "!" & "J" & i), 1 '写入移动电话  
    If Range(TheTable & "!" & "M" & i) <> "" Then ADSM.writetext "TEL;TYPE=WORK:" & Range(TheTable & "!" & "M" & i), 1 '写入工作电话，但此处仅是职务值
    If Range(TheTable & "!" & "L" & i) <> "" Then ADSM.writetext "TEL;TYPE=HOME:" & Range(TheTable & "!" & "L" & i), 1 '写入住宅电话
    If Range(TheTable & "!" & "O" & i) <> "" Then ADSM.writetext "EMAIL;TYPE=WORK:" & Range(TheTable & "!" & "O" & i), 1 '写入办公邮件
    If Range(ThexTable & "N" & i) <> "" Then ADSM.writetext "EMAIL;TYPE=HOME:" & Range(TheTable & "!" & "N" & i), 1 '写入个人邮件
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
  End If
    ADSM.writetext "End:VCARD", 1
  End Select
Next
ADSM.Position = 3   
ADSM2.Type = 1
ADSM2.Open
ADSM.copyto ADSM2
ADSM2.savetofile FilePath & FileName & ".vcf", 2   '将adsm2写入到文件
Debug.Print ("ok")
End Sub
