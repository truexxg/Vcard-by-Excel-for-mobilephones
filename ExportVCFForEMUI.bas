Attribute VB_Name = "����Ϊ��ΪEMUI��Ƭ"
Option Explicit
Sub ExportVCFForEMUI()
'''����ΪEMUI��Ƭvcf��ʽ
'''����׼��
Dim FSO
Dim ADSM, ADSM2
Dim FilePath, FileName
Dim TheTable, CtrlLogo As String
Dim LastContactCell As Range
Dim RelationType
Set FSO = CreateObject("scripting.filesystemobject")
Set ADSM = CreateObject("ADODB.Stream")
Set ADSM2 = CreateObject("ADODB.Stream")
FilePath = Range("��������!C2")
FileName = Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
TheTable = "������ϵ��Ϣ"
ADSM.Type = 2
ADSM.Charset = "UTF-8"
ADSM.Open
Dim i, j
For i = 1 To Sheets.Item(TheTable).UsedRange.Rows.Count   'ע��sheet����Ŀǰ����֪����ô��row����Ϊ������λ
CtrlLogo = Sheets.Item(TheTable).Cells(i, 2)
Select Case CtrlLogo
Case Range("��������!C1")
ADSM.writetext "BEGIN:VCARD", 1
ADSM.writetext "VERSION:3.0", 1
If Range(TheTable & "!" & "C" & i) <> "" Or Range(TheTable & "!" & "D" & i) <> "" Then    '  ȷ������������ȫΪ��
ADSM.writetext "FN:" & Range(TheTable & "!" & "C" & i) & Range(TheTable & "!" & "D" & i), 1    'д��FN�ֶ���������Ϊֵ
ADSM.writetext "N:" & Range(TheTable & "!" & "C" & i) & ";" & Range(TheTable & "!" & "D" & i), 1  'д������N�ֶ�
If Range(TheTable & "!" & "J" & i) <> "" Then ADSM.writetext "TEL;TYPE=CELL:" & Range(TheTable & "!" & "J" & i), 1 'д���ƶ��绰
If Range(TheTable & "!" & "M" & i) <> "" Then ADSM.writetext "TEL;TYPE=WORK:" & Range(TheTable & "!" & "M" & i), 1 'д�빤���绰�����˴�����ְ��ֵ
If Range(TheTable & "!" & "L" & i) <> "" Then ADSM.writetext "TEL;TYPE=HOME:" & Range(TheTable & "!" & "L" & i), 1 'д��סլ�绰
If Range(TheTable & "!" & "O" & i) <> "" Then ADSM.writetext "EMAIL;TYPE=WORK:" & Range(TheTable & "!" & "O" & i), 1 'д��칫�ʼ�
If Range(TheTable & "!" & "N" & i) <> "" Then ADSM.writetext "EMAIL;TYPE=HOME:" & Range(TheTable & "!" & "N" & i), 1 'д������ʼ�
If Range(TheTable & "!" & "G" & i) <> "" Then ADSM.writetext "ORG:" & Range(TheTable & "!" & "G" & i) & Range(TheTable & "!" & "H" & i), 1  'д����֯��������λ�Ͳ���
If Range(TheTable & "!" & "I" & i) <> "" Then ADSM.writetext "TITLE:" & Range(TheTable & "!" & "I" & i), 1  'д��ͷ�Σ�ʹ��ְ����I
If Range(TheTable & "!" & "P" & i) <> "" Then ADSM.writetext "ADR;TYPE=HOME:;;" & Range(TheTable & "!" & "P" & i) & ";;;;", 1  'д���ͥסַ������P��
If Range(TheTable & "!" & "Q" & i) <> "" Then ADSM.writetext "ADR;TYPE=WORK:;;" & Range(TheTable & "!" & "Q" & i) & ";;;;", 1  'д���ͥסַ������Q��
If Range(TheTable & "!" & "R" & i) <> "" Then ADSM.writetext "BDAY:" & Format(Range(TheTable & "!" & "R" & i), "yyyy-m-d"), 1  'д�����գ���R��
If Range(TheTable & "!" & "S" & i) <> "" Then ADSM.writetext "URL:" & Range(TheTable & "!" & "S" & i), 1   'д����վ��ʹ��S�У�����ʹ��Mailto:����ͷ
If Range(TheTable & "!" & "K" & i) <> "" Then ADSM.writetext "X-QQ:" & Range(TheTable & "!" & "K" & i), 1   'д��QQ��С�����Զ�����QQ����
If Range(TheTable & "!" & "T" & i) <> "" Then ADSM.writetext "X-ANDROID-CUSTOM:vnd.android.cursor.item/contact_event;" & Range("��ϸ��Ϣ!" & "T" & i) & ";1;��1����ʾ����������ҿ����У�;;;;;;;;;;;;", 1  'д����������գ�ʹ����ϸ��ϢT��
If Range(TheTable & "!" & "U" & i) <> "" Then ADSM.writetext "X-ANDROID-CUSTOM:vnd.android.cursor.item/contact_event;" & Range("��ϸ��Ϣ!" & "U" & i) & ";3;��1����ʾ����������ҿ����У�;;;;;;;;;;;;", 1  'д����������գ�ʹ����ϸ��ϢU��
If Range(TheTable & "!" & "V" & i) <> "" Then ADSM.writetext "X-ANDROID-CUSTOM:vnd.android.cursor.item/contact_event;" & Range("��ϸ��Ϣ!" & "V" & i) & ";0;��1����ʾ����������ҿ����У�;;;;;;;;;;;;", 1  'д����������գ�ʹ����ϸ��ϢV��
For j = 1 To Sheets.Item("��ϵ��").UsedRange.Rows.Count
Select Case Range("��ϵ��!C" & j).Value
Case "����", "����"
RelationType = 1
Case "�ֵ�"
RelationType = 2
Case "��Ů", "����", "Ů��"
RelationType = 3
Case "����", "����", "����"
RelationType = 4
Case "����", "�ְ�"
RelationType = 5
Case "����", "����"
RelationType = 6
Case "��˾", "�ϰ�", "����"
RelationType = 7
Case "ĸ��", "����"
RelationType = 8
Case "��ĸ"
RelationType = 9
Case "�������", "�ϻ���", "ͬ��"
RelationType = 10
Case "������"
RelationType = 11
Case "����"
RelationType = 12
Case "����", "���", "����"
RelationType = 13
Case "��ż", "�Ϲ�", "����", "����", "ϱ��"
RelationType = 14
Case Else
RelationType = 0
End Select
If Range("��ϵ��!A" & j) = Range(TheTable & "!" & "A" & i).Value Then ADSM.writetext "X-ANDROID-CUSTOM:vnd.android.cursor.item/relation;" & Range("��ϵ��!D" & j) & ";" & RelationType & ";" & Range("��ϵ��!C" & j) & ";;;;;;;;;;;;", 1 '������ϵ����д����˵Ĺ�ϵ��
Next

'   Set LastContactCell = Range("��ϵ��!A1")    '��δ�������ʹ��range.find���������ҷ��������ĵ�Ԫ�񣬷��ֻ�����forѭ�����ü�ࡣ�����㷨������Ҳ��һ��Ҫ�����ģ��ǾͲ���ֱ��for��
'   Set FirstContactCell = Nothing
'   Do
'      Set LastContactCell = Range("��ϵ��!A:A").Find(Range(TheTable & "!" & "A" & i).Value, LastContactCell, xlValues, xlWhole, xlByRows, xlNext)
'      If FirstContactCell Is Nothing Then Set FirstContactCell = LastContactCell
'      If Not LastContactCell Is Nothing Then Debug.Print LastContactCell.Address
'   Loop Until FirstContactCell Is Range("��ϵ��!A:A").FindNext(LastContactCell)
End If
ADSM.writetext "END:VCARD", 1
Range(TheTable & "!" & "B" & i) = ""  'ɾ�����ɱ��m
End Select
Next
ADSM.Position = 3
ADSM2.Type = 1
ADSM2.Open
ADSM.copyto ADSM2
ADSM2.savetofile FilePath & FileName & ".vcf", 2   '��adsm2д�뵽�ļ�
Debug.Print ("ok")
End Sub
