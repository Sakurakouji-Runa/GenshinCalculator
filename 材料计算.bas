Attribute VB_Name = "���ϼ���"

Sub ���ϼ���()
    
    Application.ScreenUpdating = False
    Application.Interactive = False
    
    'Part0 ���״̬Ԥ����
    Sheets("��ҳ").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("����-���ɵ���").Select
    ActiveSheet.AutoFilterMode = False
    '���Ⱦɫ״̬
    i = 2
    Do While Cells(i, 1).Value <> ""
        Cells(i, 6).Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        i = i + 1
    Loop
    Sheets("����-����").Select
    ActiveSheet.AutoFilterMode = False
    '���Ⱦɫ״̬
    i = 2
    Do While Cells(i, 1).Value <> ""
        Cells(i, 4).Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        i = i + 1
    Loop
    Sheets("��ɫ").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("��ɫ�ȼ�").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("��ɫ����").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("�츳����").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("����").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("�����ȼ�").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("��������").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("ʥ����").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    '�����һ�εļ�����
    Sheets("����-���ɵ���").Range("E:E").Clear
    Sheets("����-���ɵ���").Range("E1") = "��������"
    Sheets("����-���ɵ���").Range("F:F").Clear
    Sheets("����-���ɵ���").Range("F1") = "ȱ������"
    Sheets("����-���ɵ���").Range("H:H").Clear
    Sheets("����-���ɵ���").Range("H1") = "������Դ"
    Sheets("����-���ɵ���").Range("I:I").Clear
    Sheets("����-���ɵ���").Range("I1") = "��Դ����"
    Sheets("����-����").Range("C:C").Clear
    Sheets("����-����").Range("C1") = "��������"
    Sheets("����-����").Range("D:D").Clear
    Sheets("����-����").Range("D1") = "ȱ������"
    Sheets("����-����").Range("E:E").Clear
    Sheets("����-����").Range("E1") = "������Դ"
    Sheets("����-����").Range("F:F").Clear
    Sheets("����-����").Range("F1") = "��Դ����"
    
    'Part1 ������������
    Sheets("��ҳ").Select
    'ֻҪ��һ�в��վͻ�һֱ��
    i = 2
    Do While Cells(i, 1) <> ""
        expSingle = 0
        molaSingle = 0
        roleName = Cells(i, 1).Value
        '�Ƿ���ͻ���ж�
        lvStart = Cells(i, 2).Value
        If Right(lvStart, 1) = "+" Then
            lvStart = CInt(Left(lvStart, Len(lvStart) - 1))
            lvStartAscensions = 1
        Else
            lvStartAscensions = 0
        End If
        lvEnd = Cells(i, 3).Value
        If Right(lvEnd, 1) = "+" Then
            lvEnd = CInt(Left(lvEnd, Len(lvEnd) - 1))
            lvEndAscensions = 1
        Else
            lvEndAscensions = 0
        End If
        'ֻ�е�Ŀ��ȼ�>��ǰ�ȼ�ʱ�ż���
        If (lvEnd > lvStart Or (lvEnd = lvStart And lvEndAscensions = 1)) Then
            matchStart = Application.WorksheetFunction.Match(lvStart, Sheets("��ɫ�ȼ�").Range("A:A"), 0)
            matchEnd = Application.WorksheetFunction.Match(lvEnd, Sheets("��ɫ�ȼ�").Range("A:A"), 0)
            'Part1-1 ������������Ħ��
            expSingle = Sheets("��ɫ�ȼ�").Cells(matchEnd, 2).Value - Sheets("��ɫ�ȼ�").Cells(matchStart, 2).Value
            molaSingle = Sheets("��ɫ�ȼ�").Cells(matchEnd, 3).Value - Sheets("��ɫ�ȼ�").Cells(matchStart, 3).Value
            If lvStartAscensions = 0 Then
                molaSingle = molaSingle + Sheets("��ɫ�ȼ�").Cells(matchStart, 6).Value
            End If
            If lvEndAscensions = 0 Then
                molaSingle = molaSingle - Sheets("��ɫ�ȼ�").Cells(matchEnd, 6).Value
            End If
            If expSingle <> 0 Then
                '��������������д
                If Sheets("����-���ɵ���").Cells(3, 5).Value = "" Then
                    Sheets("����-���ɵ���").Cells(3, 5).Value = expSingle
                Else
                    Sheets("����-���ɵ���").Cells(3, 5).Value = Sheets("����-���ɵ���").Cells(3, 5).Value + expSingle
                End If
                '����������Դ��д
                If Sheets("����-���ɵ���").Cells(3, 8).Value = "" Then
                    Sheets("����-���ɵ���").Cells(3, 8).Value = roleName & "����"
                Else
                    Sheets("����-���ɵ���").Cells(3, 8).Value = Sheets("����-���ɵ���").Cells(3, 8).Value & Chr(10) & roleName & "����"
                End If
                '����������Դ������д
                If Sheets("����-���ɵ���").Cells(3, 9).Value = "" Then
                    Sheets("����-���ɵ���").Cells(3, 9).Value = expSingle
                Else
                    Sheets("����-���ɵ���").Cells(3, 9).Value = Sheets("����-���ɵ���").Cells(3, 9).Value & Chr(10) & expSingle
                End If
            End If
            If molaSingle <> 0 Then
            'Ħ������������д
                If Sheets("����-����").Cells(2, 3).Value = "" Then
                    Sheets("����-����").Cells(2, 3).Value = molaSingle
                Else
                    Sheets("����-����").Cells(2, 3).Value = Sheets("����-����").Cells(2, 3).Value + molaSingle
                End If
                'Ħ��������Դ��д
                If Sheets("����-����").Cells(2, 5).Value = "" Then
                    Sheets("����-����").Cells(2, 5).Value = roleName & "����"
                Else
                    Sheets("����-����").Cells(2, 5).Value = Sheets("����-����").Cells(2, 5).Value & Chr(10) & roleName & "����"
                End If
                'Ħ��������Դ������д
                If Sheets("����-����").Cells(2, 6).Value = "" Then
                    Sheets("����-����").Cells(2, 6).Value = molaSingle
                Else
                    Sheets("����-����").Cells(2, 6).Value = Sheets("����-����").Cells(2, 6).Value & Chr(10) & molaSingle
                End If
            End If
            'Part1-2 ������������
            Dim eliteProb(1 To 4)
            Dim eliteMust(1 To 1)
            Dim gathering(1 To 1)
            Dim normalDrop(1 To 3)
            eliteProb(1) = 0
            eliteProb(2) = 0
            eliteProb(3) = 0
            eliteProb(4) = 0
            eliteMust(1) = 0
            gathering(1) = 0
            normalDrop(1) = 0
            normalDrop(2) = 0
            normalDrop(3) = 0
            '�ȼ�ȡ������
            j = 2
            Do While lvStart >= Sheets("��ɫ����").Cells(j, 1).Value And Sheets("��ɫ����").Cells(j, 1).Value <> ""
                j = j + 1
            Loop
            If lvStart > Sheets("��ɫ����").Cells(j - 1, 1).Value Then
                lvStartAscensions = 1
            End If
            lvStart = Sheets("��ɫ����").Cells(j - 1, 1).Value
            j = 2
            Do While lvEnd >= Sheets("��ɫ����").Cells(j, 1).Value And Sheets("��ɫ����").Cells(j, 1).Value <> ""
                j = j + 1
            Loop
            If lvEnd > Sheets("��ɫ����").Cells(j - 1, 1).Value Then
                lvEndAscensions = 1
            End If
            lvEnd = Sheets("��ɫ����").Cells(j - 1, 1).Value
            matchStart = Application.WorksheetFunction.Match(lvStart, Sheets("��ɫ����").Range("A:A"), 0)
            matchEnd = Application.WorksheetFunction.Match(lvEnd, Sheets("��ɫ����").Range("A:A"), 0)
            '�Ƿ���ͻ�ƴ���
            matchStart = matchStart + lvStartAscensions
            matchEnd = matchEnd + lvEndAscensions
            'ͳ��������Ҫ����
            For j = matchStart To (matchEnd - 1)
                If Sheets("��ɫ����").Cells(j, 2).Value <> "" Then
                    eliteProb(Sheets("��ɫ����").Cells(j, 2).Value) = eliteProb(Sheets("��ɫ����").Cells(j, 2).Value) + Sheets("��ɫ����").Cells(j, 3).Value
                End If
                If Sheets("��ɫ����").Cells(j, 4).Value <> "" Then
                    eliteMust(Sheets("��ɫ����").Cells(j, 4).Value) = eliteMust(Sheets("��ɫ����").Cells(j, 4).Value) + Sheets("��ɫ����").Cells(j, 5).Value
                End If
                If Sheets("��ɫ����").Cells(j, 6).Value <> "" Then
                    gathering(Sheets("��ɫ����").Cells(j, 6).Value) = gathering(Sheets("��ɫ����").Cells(j, 6).Value) + Sheets("��ɫ����").Cells(j, 7).Value
                End If
                If Sheets("��ɫ����").Cells(j, 8).Value <> "" Then
                    normalDrop(Sheets("��ɫ����").Cells(j, 8).Value) = normalDrop(Sheets("��ɫ����").Cells(j, 8).Value) + Sheets("��ɫ����").Cells(j, 9).Value
                End If
            Next j
            '��ȡ��������
            roleMatch = Application.WorksheetFunction.Match(roleName, Sheets("��ɫ").Range("A:A"), 0)
            eliteProbName = Sheets("��ɫ").Cells(roleMatch, 3).Value
            eliteMustName = Sheets("��ɫ").Cells(roleMatch, 4).Value
            gatheringName = Sheets("��ɫ").Cells(roleMatch, 5).Value
            normalDropName = Sheets("��ɫ").Cells(roleMatch, 6).Value
            '��������λ�ò�ѯ
            eliteProbPos = Application.WorksheetFunction.Match(eliteProbName, Sheets("����-���ɵ���").Range("A:A"), 0) + 3
            eliteMustPos = Application.WorksheetFunction.Match(eliteMustName, Sheets("����-���ɵ���").Range("A:A"), 0)
            gatheringPos = Application.WorksheetFunction.Match(gatheringName, Sheets("����-����").Range("A:A"), 0)
            normalDropPos = Application.WorksheetFunction.Match(normalDropName, Sheets("����-���ɵ���").Range("A:A"), 0) + 2
            '����������������Դ��д
            For j = 1 To 4
                If eliteProb(j) <> 0 Then
                    If Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 5).Value = "" Then
                        Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 5).Value = eliteProb(j)
                    Else
                        Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 5).Value = Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 5).Value + eliteProb(j)
                    End If
                    If Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 8).Value = "" Then
                        Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 8).Value = roleName & "ͻ��"
                    Else
                        Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 8).Value = Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 8).Value & Chr(10) & roleName & "ͻ��"
                    End If
                    If Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 9).Value = "" Then
                        Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 9).Value = eliteProb(j)
                    Else
                        Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 9).Value = Sheets("����-���ɵ���").Cells(eliteProbPos - j + 1, 9).Value & Chr(10) & eliteProb(j)
                    End If
                End If
            Next j
            If eliteMust(1) <> 0 Then
                If Sheets("����-���ɵ���").Cells(eliteMustPos, 5).Value = "" Then
                    Sheets("����-���ɵ���").Cells(eliteMustPos, 5).Value = eliteMust(1)
                Else
                    Sheets("����-���ɵ���").Cells(eliteMustPos, 5).Value = Sheets("����-���ɵ���").Cells(eliteMustPos, 5).Value + eliteMust(1)
                End If
                If Sheets("����-���ɵ���").Cells(eliteMustPos, 8).Value = "" Then
                    Sheets("����-���ɵ���").Cells(eliteMustPos, 8).Value = roleName & "ͻ��"
                Else
                    Sheets("����-���ɵ���").Cells(eliteMustPos, 8).Value = Sheets("����-���ɵ���").Cells(eliteMustPos, 8).Value & Chr(10) & roleName & "ͻ��"
                End If
                If Sheets("����-���ɵ���").Cells(eliteMustPos, 9).Value = "" Then
                    Sheets("����-���ɵ���").Cells(eliteMustPos, 9).Value = eliteMust(1)
                Else
                    Sheets("����-���ɵ���").Cells(eliteMustPos, 9).Value = Sheets("����-���ɵ���").Cells(eliteMustPos, 9).Value & Chr(10) & eliteMust(1)
                End If
            End If
            If gathering(1) <> 0 Then
                If Sheets("����-����").Cells(gatheringPos, 3).Value = "" Then
                    Sheets("����-����").Cells(gatheringPos, 3).Value = gathering(1)
                Else
                    Sheets("����-����").Cells(gatheringPos, 3).Value = Sheets("����-����").Cells(gatheringPos, 3).Value + gathering(1)
                End If
                If Sheets("����-����").Cells(gatheringPos, 5).Value = "" Then
                    Sheets("����-����").Cells(gatheringPos, 5).Value = roleName & "ͻ��"
                Else
                    Sheets("����-����").Cells(gatheringPos, 5).Value = Sheets("����-����").Cells(gatheringPos, 5).Value & Chr(10) & roleName & "ͻ��"
                End If
                If Sheets("����-����").Cells(gatheringPos, 6).Value = "" Then
                    Sheets("����-����").Cells(gatheringPos, 6).Value = gathering(1)
                Else
                    Sheets("����-����").Cells(gatheringPos, 6).Value = Sheets("����-����").Cells(gatheringPos, 6).Value & Chr(10) & gathering(1)
                End If
            End If
            For j = 1 To 3
                If normalDrop(j) <> 0 Then
                    If Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 5).Value = "" Then
                        Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 5).Value = normalDrop(j)
                    Else
                        Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 5).Value = Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 5).Value + normalDrop(j)
                    End If
                    If Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 8).Value = "" Then
                        Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 8).Value = roleName & "ͻ��"
                    Else
                        Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 8).Value = Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 8).Value & Chr(10) & roleName & "ͻ��"
                    End If
                    If Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 9).Value = "" Then
                        Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 9).Value = normalDrop(j)
                    Else
                        Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 9).Value = Sheets("����-���ɵ���").Cells(normalDropPos - j + 1, 9).Value & Chr(10) & normalDrop(j)
                    End If
                End If
            Next j
        End If
        i = i + 1
    Loop
    '������Ϊ��ʱ����д����ȱ������
    If Sheets("����-���ɵ���").Cells(3, 5).Value <> "" Then
        '��Ҫ��������
        expRole = Sheets("����-���ɵ���").Cells(3, 5).Value
        '���о�������
        exp3 = Sheets("����-���ɵ���").Cells(2, 4).Value
        exp2 = Sheets("����-���ɵ���").Cells(3, 4).Value
        exp1 = Sheets("����-���ɵ���").Cells(4, 4).Value
        expHave = exp3 * 20000 + exp2 * 5000 + exp1 * 1000
        'ȱ�ھ�������������Ӣ�۾������
        expNeed = expRole - expHave
        If expNeed > 0 Then
            Sheets("����-���ɵ���").Cells(3, 6).Value = expNeed
            Sheets("����-���ɵ���").Cells(2, 6).Value = "�ۺ�" & Application.WorksheetFunction.RoundUp((expRole - expHave) / 20000, 0)
        Else
            Sheets("����-���ɵ���").Cells(3, 6).Value = "����"
        End If
    End If
    'Ħ��ȱ������������ȱ�����������һ����
    
    'Part2 �����츳����
    i = 2
    Do While Cells(i, 1) <> ""
        roleName = Cells(i, 1).Value
        '��ʼ������������
        Dim talentBook(1 To 3)
        Dim talentDrop(1 To 3)
        Dim weekDrop(1 To 1)
        Dim activityGet(1 To 1)
        talentBook(1) = 0
        talentBook(2) = 0
        talentBook(3) = 0
        talentDrop(1) = 0
        talentDrop(2) = 0
        talentDrop(3) = 0
        weekDrop(1) = 0
        activityGet(1) = 0
        molaTalent = 0
        '��ȡ��������
        roleMatch = Application.WorksheetFunction.Match(roleName, Sheets("��ɫ").Range("A:A"), 0)
        talentBookName = Sheets("��ɫ").Cells(roleMatch, 7).Value
        talentDropName = Sheets("��ɫ").Cells(roleMatch, 8).Value
        weekDropName = Sheets("��ɫ").Cells(roleMatch, 9).Value
        activityGetName = Sheets("��ɫ").Cells(roleMatch, 10).Value
        '��������λ�ò�ѯ
        talentBookPos = Application.WorksheetFunction.Match(talentBookName, Sheets("����-���ɵ���").Range("A:A"), 0) + 2
        talentDropPos = Application.WorksheetFunction.Match(talentDropName, Sheets("����-���ɵ���").Range("A:A"), 0) + 2
        If weekDropName = "" Then
            weekDropPos = -1
        Else
            weekDropPos = Application.WorksheetFunction.Match(weekDropName, Sheets("����-���ɵ���").Range("A:A"), 0)
        End If
        activityGetPos = Application.WorksheetFunction.Match(activityGetName, Sheets("����-���ɵ���").Range("A:A"), 0)
        For j = 4 To 8 Step 2
            talentStart = Cells(i, j).Value
            talentEnd = Cells(i, j + 1).Value
            'ֻ�е�Ŀ��ȼ�>��ǰ�ȼ�ʱ�ż���
            If talentEnd > talentStart Then
                matchStart = Application.WorksheetFunction.Match(talentStart, Sheets("�츳����").Range("A:A"), 0)
                matchEnd = Application.WorksheetFunction.Match(talentEnd, Sheets("�츳����").Range("A:A"), 0)
                'ͳ���츳����Ħ��
                If matchStart = 2 Then
                    molaSingle = Sheets("�츳����").Cells(matchEnd, 10).Value
                Else
                    molaSingle = Sheets("�츳����").Cells(matchEnd, 10).Value - Sheets("�츳����").Cells(matchStart, 10).Value
                End If
                molaTalent = molaTalent + molaSingle
                'ͳ���츳��������
                For k = (matchStart + 1) To matchEnd
                    If Sheets("�츳����").Cells(k, 2).Value <> "" Then
                        talentBook(Sheets("�츳����").Cells(k, 2).Value) = talentBook(Sheets("�츳����").Cells(k, 2).Value) + Sheets("�츳����").Cells(k, 3).Value
                    End If
                    If Sheets("�츳����").Cells(k, 4).Value <> "" Then
                        talentDrop(Sheets("�츳����").Cells(k, 4).Value) = talentDrop(Sheets("�츳����").Cells(k, 4).Value) + Sheets("�츳����").Cells(k, 5).Value
                    End If
                    If Sheets("�츳����").Cells(k, 6).Value <> "" Then
                        weekDrop(Sheets("�츳����").Cells(k, 6).Value) = weekDrop(Sheets("�츳����").Cells(k, 6).Value) + Sheets("�츳����").Cells(k, 7).Value
                    End If
                    If Sheets("�츳����").Cells(k, 8).Value <> "" Then
                        activityGet(Sheets("�츳����").Cells(k, 8).Value) = activityGet(Sheets("�츳����").Cells(k, 8).Value) + Sheets("�츳����").Cells(k, 9).Value
                    End If
                Next k
            End If
        Next j
        '����������������Դ��д
        For j = 1 To 3
            If talentBook(j) <> 0 Then
                If Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 5).Value = "" Then
                    Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 5).Value = talentBook(j)
                Else
                    Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 5).Value = Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 5).Value + talentBook(j)
                End If
                If Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 8).Value = "" Then
                    Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 8).Value = roleName & "�츳"
                Else
                    Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 8).Value = Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 8).Value & Chr(10) & roleName & "�츳"
                End If
                If Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 9).Value = "" Then
                    Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 9).Value = talentBook(j)
                Else
                    Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 9).Value = Sheets("����-���ɵ���").Cells(talentBookPos - j + 1, 9).Value & Chr(10) & talentBook(j)
                End If
            End If
        Next j
        For j = 1 To 3
            If talentDrop(j) <> 0 Then
                If Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 5).Value = "" Then
                    Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 5).Value = talentDrop(j)
                Else
                    Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 5).Value = Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 5).Value + talentDrop(j)
                End If
                If Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 8).Value = "" Then
                    Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 8).Value = roleName & "�츳"
                Else
                    Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 8).Value = Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 8).Value & Chr(10) & roleName & "�츳"
                End If
                If Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 9).Value = "" Then
                    Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 9).Value = talentDrop(j)
                Else
                    Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 9).Value = Sheets("����-���ɵ���").Cells(talentDropPos - j + 1, 9).Value & Chr(10) & talentDrop(j)
                End If
            End If
        Next j
        If (weekDrop(1) <> 0 And weekDropPos > 0) Then
            If Sheets("����-���ɵ���").Cells(weekDropPos, 5).Value = "" Then
                Sheets("����-���ɵ���").Cells(weekDropPos, 5).Value = weekDrop(1)
            Else
                Sheets("����-���ɵ���").Cells(weekDropPos, 5).Value = Sheets("����-���ɵ���").Cells(weekDropPos, 5).Value + weekDrop(1)
            End If
            If Sheets("����-���ɵ���").Cells(weekDropPos, 8).Value = "" Then
                Sheets("����-���ɵ���").Cells(weekDropPos, 8).Value = roleName & "�츳"
            Else
                Sheets("����-���ɵ���").Cells(weekDropPos, 8).Value = Sheets("����-���ɵ���").Cells(weekDropPos, 8).Value & Chr(10) & roleName & "�츳"
            End If
            If Sheets("����-���ɵ���").Cells(weekDropPos, 9).Value = "" Then
                Sheets("����-���ɵ���").Cells(weekDropPos, 9).Value = weekDrop(1)
            Else
                Sheets("����-���ɵ���").Cells(weekDropPos, 9).Value = Sheets("����-���ɵ���").Cells(weekDropPos, 9).Value & Chr(10) & weekDrop(1)
            End If
        End If
        If activityGet(1) <> 0 Then
            If Sheets("����-���ɵ���").Cells(activityGetPos, 5).Value = "" Then
                Sheets("����-���ɵ���").Cells(activityGetPos, 5).Value = activityGet(1)
            Else
                Sheets("����-���ɵ���").Cells(activityGetPos, 5).Value = Sheets("����-���ɵ���").Cells(activityGetPos, 5).Value + activityGet(1)
            End If
            If Sheets("����-���ɵ���").Cells(activityGetPos, 8).Value = "" Then
                Sheets("����-���ɵ���").Cells(activityGetPos, 8).Value = roleName & "�츳"
            Else
                Sheets("����-���ɵ���").Cells(activityGetPos, 8).Value = Sheets("����-���ɵ���").Cells(activityGetPos, 8).Value & Chr(10) & roleName & "�츳"
            End If
            If Sheets("����-���ɵ���").Cells(activityGetPos, 9).Value = "" Then
                Sheets("����-���ɵ���").Cells(activityGetPos, 9).Value = activityGet(1)
            Else
                Sheets("����-���ɵ���").Cells(activityGetPos, 9).Value = Sheets("����-���ɵ���").Cells(activityGetPos, 9).Value & Chr(10) & activityGet(1)
            End If
        End If
        If molaTalent <> 0 Then
            'Ħ������������д
            If Sheets("����-����").Cells(2, 3).Value = "" Then
                Sheets("����-����").Cells(2, 3).Value = molaTalent
            Else
                Sheets("����-����").Cells(2, 3).Value = Sheets("����-����").Cells(2, 3).Value + molaTalent
            End If
            'Ħ��������Դ��д
            If Sheets("����-����").Cells(2, 5).Value = "" Then
                Sheets("����-����").Cells(2, 5).Value = roleName & "�츳"
            Else
                Sheets("����-����").Cells(2, 5).Value = Sheets("����-����").Cells(2, 5).Value & Chr(10) & roleName & "�츳"
            End If
            'Ħ����Դ������д
            If Sheets("����-����").Cells(2, 6).Value = "" Then
                Sheets("����-����").Cells(2, 6).Value = molaTalent
            Else
                Sheets("����-����").Cells(2, 6).Value = Sheets("����-����").Cells(2, 6).Value & Chr(10) & molaTalent
            End If
        End If
        i = i + 1
    Loop
    
    'Part3 ������������
    i = 2
    Do While Cells(i, 11) <> ""
        weaponName = Cells(i, 11).Value
        '��ʼ������������
        Dim dungenDrop(1 To 4)
        Dim eliteThing(1 To 3)
        Dim normalThing(1 To 3)
        dungenDrop(1) = 0
        dungenDrop(2) = 0
        dungenDrop(3) = 0
        dungenDrop(4) = 0
        eliteThing(1) = 0
        eliteThing(2) = 0
        eliteThing(3) = 0
        normalThing(1) = 0
        normalThing(2) = 0
        normalThing(3) = 0
        expSingle = 0
        molaSingle = 0
        '��ȡ��������
        weaponMatch = Application.WorksheetFunction.Match(weaponName, Sheets("����").Range("B:B"), 0)
        dungenDropName = Sheets("����").Cells(weaponMatch, 4).Value
        eliteThingName = Sheets("����").Cells(weaponMatch, 5).Value
        normalThingName = Sheets("����").Cells(weaponMatch, 6).Value
        '��������λ�ò�ѯ
        dungenDropPos = Application.WorksheetFunction.Match(dungenDropName, Sheets("����-���ɵ���").Range("A:A"), 0) + 3
        eliteThingPos = Application.WorksheetFunction.Match(eliteThingName, Sheets("����-���ɵ���").Range("A:A"), 0) + 2
        normalThingPos = Application.WorksheetFunction.Match(normalThingName, Sheets("����-���ɵ���").Range("A:A"), 0) + 2
        '�Ƿ���ͻ���ж�
        lvStart = Cells(i, 12).Value
        If Right(lvStart, 1) = "+" Then
            lvStart = CInt(Left(lvStart, Len(lvStart) - 1))
            lvStartAscensions = 1
        Else
            lvStartAscensions = 0
        End If
        lvEnd = Cells(i, 13).Value
        If Right(lvEnd, 1) = "+" Then
            lvEnd = CInt(Left(lvEnd, Len(lvEnd) - 1))
            lvEndAscensions = 1
        Else
            lvEndAscensions = 0
        End If
        '��Ŀ��ȼ�>��ǰ�ȼ�ʱ�ż���
        If (lvEnd > lvStart Or (lvEnd = lvStart And lvEndAscensions = 1)) Then
            '�����Ǽ���ѯ
            weaponStar = Sheets("����").Cells(Application.WorksheetFunction.Match(weaponName, Sheets("����").Range("B:B"), 0), 3).Value
            starMatch = Application.WorksheetFunction.Match(weaponStar, Sheets("�����ȼ�").Range("A:A"), 0)
            weaponUpmax = Application.WorksheetFunction.CountIf(Sheets("�����ȼ�").Range("A:A"), weaponStar)
            matchStart = Application.WorksheetFunction.Match(lvStart, Sheets("�����ȼ�").Cells(starMatch, 2).Resize(weaponUpmax, 1), 0) + starMatch - 1
            matchEnd = Application.WorksheetFunction.Match(lvEnd, Sheets("�����ȼ�").Cells(starMatch, 2).Resize(weaponUpmax, 1), 0) + starMatch - 1
            'Part3-1 ������������Ħ��
            expSingle = Sheets("�����ȼ�").Cells(matchEnd, 3).Value - Sheets("�����ȼ�").Cells(matchStart, 3).Value
            molaSingle = Sheets("�����ȼ�").Cells(matchEnd, 4).Value - Sheets("�����ȼ�").Cells(matchStart, 4).Value
            If expSingle <> 0 Then
                '��������������д
                If Sheets("����-����").Cells(4, 3).Value = "" Then
                    Sheets("����-����").Cells(4, 3).Value = expSingle
                Else
                    Sheets("����-����").Cells(4, 3).Value = Sheets("����-����").Cells(4, 3).Value + expSingle
                End If
                '����������Դ��д
                If Sheets("����-����").Cells(4, 5).Value = "" Then
                    Sheets("����-����").Cells(4, 5).Value = weaponName & "����"
                Else
                    Sheets("����-����").Cells(4, 5).Value = Sheets("����-����").Cells(4, 5).Value & Chr(10) & weaponName & "����"
                End If
                '������Դ������д
                If Sheets("����-����").Cells(4, 6).Value = "" Then
                    Sheets("����-����").Cells(4, 6).Value = expSingle
                Else
                    Sheets("����-����").Cells(4, 6).Value = Sheets("����-����").Cells(4, 6).Value & Chr(10) & expSingle
                End If
            End If
            If molaSingle <> 0 Then
            'Ħ������������д
                If Sheets("����-����").Cells(2, 3).Value = "" Then
                    Sheets("����-����").Cells(2, 3).Value = molaSingle
                Else
                    Sheets("����-����").Cells(2, 3).Value = Sheets("����-����").Cells(2, 3).Value + molaSingle
                End If
                'Ħ��������Դ��д
                If Sheets("����-����").Cells(2, 5).Value = "" Then
                    Sheets("����-����").Cells(2, 5).Value = weaponName & "����"
                Else
                    Sheets("����-����").Cells(2, 5).Value = Sheets("����-����").Cells(2, 5).Value & Chr(10) & weaponName & "����"
                End If
                'Ħ����Դ������д
                If Sheets("����-����").Cells(2, 6).Value = "" Then
                    Sheets("����-����").Cells(2, 6).Value = molaSingle
                Else
                    Sheets("����-����").Cells(2, 6).Value = Sheets("����-����").Cells(2, 6).Value & Chr(10) & molaSingle
                End If
            End If
            'Part3-2 ������������
            '�Ƿ���ͻ�ƴ���
            matchStart = matchStart + lvStartAscensions
            matchEnd = matchEnd + lvEndAscensions
            For j = matchStart To (matchEnd - 1)
                If Sheets("��������").Cells(j, 3).Value <> "" Then
                    dungenDrop(Sheets("��������").Cells(j, 3).Value) = dungenDrop(Sheets("��������").Cells(j, 3).Value) + Sheets("��������").Cells(j, 4).Value
                End If
                If Sheets("��������").Cells(j, 5).Value <> "" Then
                    eliteThing(Sheets("��������").Cells(j, 5).Value) = eliteThing(Sheets("��������").Cells(j, 5).Value) + Sheets("��������").Cells(j, 6).Value
                End If
                If Sheets("��������").Cells(j, 7).Value <> "" Then
                    normalThing(Sheets("��������").Cells(j, 7).Value) = normalThing(Sheets("��������").Cells(j, 7).Value) + Sheets("��������").Cells(j, 8).Value
                End If
            Next j
            '����������������Դ��д
            For j = 1 To 4
                If dungenDrop(j) <> 0 Then
                    If Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 5).Value = "" Then
                        Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 5).Value = dungenDrop(j)
                    Else
                        Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 5).Value = Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 5).Value + dungenDrop(j)
                    End If
                    If Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 8).Value = "" Then
                        Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 8).Value = weaponName & "ͻ��"
                    Else
                        Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 8).Value = Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 8).Value & Chr(10) & weaponName & "ͻ��"
                    End If
                    If Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 9).Value = "" Then
                        Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 9).Value = dungenDrop(j)
                    Else
                        Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 9).Value = Sheets("����-���ɵ���").Cells(dungenDropPos - j + 1, 9).Value & Chr(10) & dungenDrop(j)
                    End If
                End If
            Next j
            For j = 1 To 3
                If eliteThing(j) <> 0 Then
                    If Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 5).Value = "" Then
                        Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 5).Value = eliteThing(j)
                    Else
                        Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 5).Value = Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 5).Value + eliteThing(j)
                    End If
                    If Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 8).Value = "" Then
                        Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 8).Value = weaponName & "ͻ��"
                    Else
                        Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 8).Value = Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 8).Value & Chr(10) & weaponName & "ͻ��"
                    End If
                    If Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 9).Value = "" Then
                        Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 9).Value = eliteThing(j)
                    Else
                        Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 9).Value = Sheets("����-���ɵ���").Cells(eliteThingPos - j + 1, 9).Value & Chr(10) & eliteThing(j)
                    End If
                End If
            Next j
            For j = 1 To 3
                If normalThing(j) <> 0 Then
                    If Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 5).Value = "" Then
                        Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 5).Value = normalThing(j)
                    Else
                        Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 5).Value = Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 5).Value + normalThing(j)
                    End If
                    If Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 8).Value = "" Then
                        Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 8).Value = weaponName & "ͻ��"
                    Else
                        Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 8).Value = Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 8).Value & Chr(10) & weaponName & "ͻ��"
                    End If
                    If Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 9).Value = "" Then
                        Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 9).Value = normalThing(j)
                    Else
                        Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 9).Value = Sheets("����-���ɵ���").Cells(normalThingPos - j + 1, 9).Value & Chr(10) & normalThing(j)
                    End If
                End If
            Next j
        End If
        i = i + 1
    Loop
    '������Ϊ��ʱ����д��������ȱ������
    If Sheets("����-����").Cells(4, 3).Value <> "" Then
        '��Ҫ��������
        expWeapon = Sheets("����-����").Cells(4, 3).Value
        '���о�������
        exp3 = Sheets("����-����").Cells(5, 2).Value
        exp2 = Sheets("����-����").Cells(4, 2).Value
        exp1 = Sheets("����-����").Cells(3, 2).Value
        expHave = exp3 * 10000 + exp2 * 2000 + exp1 * 400
        'ȱ�ھ�������������Ӣ�۾������
        expNeed = expWeapon - expHave
        If expNeed > 0 Then
            Sheets("����-����").Cells(4, 4).Value = expNeed
            Sheets("����-����").Cells(5, 4).Value = "�ۺ�" & Application.WorksheetFunction.RoundUp((expWeapon - expHave) / 10000, 0)
        Else
            Sheets("����-����").Cells(4, 4).Value = "����"
        End If
    End If
    
    'Part4 ʥ������������
    i = 2
    Do While Cells(i, 15) <> ""
        relicName = Cells(i, 15).Value
        relicStar = Cells(i, 16).Value
        relicCount = Cells(i, 17).Value
        lvStart = Cells(i, 18).Value
        lvEnd = Cells(i, 19).Value
        expSingle = 0
        molaSingle = 0
         'ֻ�е�Ŀ��ȼ�>��ǰ�ȼ�ʱ�ż���
        If lvEnd > lvStart Then
            '��ȡ��������λ��
            starMatch = Application.WorksheetFunction.Match(relicStar, Sheets("ʥ����").Range("A:A"), 0)
            relicUpmax = Application.WorksheetFunction.CountIf(Sheets("ʥ����").Range("A:A"), relicStar)
            matchStart = Application.WorksheetFunction.Match(lvStart, Sheets("ʥ����").Cells(starMatch, 2).Resize(relicUpmax, 1), 0) + starMatch - 1
            matchEnd = Application.WorksheetFunction.Match(lvEnd, Sheets("ʥ����").Cells(starMatch, 2).Resize(relicUpmax, 1), 0) + starMatch - 1
            '��������ľ���Ħ������
            expSingle = Sheets("ʥ����").Cells(matchEnd, 3).Value - Sheets("ʥ����").Cells(matchStart, 3).Value
            molaSingle = Sheets("ʥ����").Cells(matchEnd, 4).Value - Sheets("ʥ����").Cells(matchStart, 4).Value
            expSingle = expSingle * relicCount
            molaSingle = molaSingle * relicCount
            '��������������д
            If Sheets("����-����").Cells(8, 3).Value = "" Then
                Sheets("����-����").Cells(8, 3).Value = expSingle
            Else
                Sheets("����-����").Cells(8, 3).Value = Sheets("����-����").Cells(8, 3).Value + expSingle
            End If
            '����������Դ��д
            If Sheets("����-����").Cells(8, 5).Value = "" Then
                Sheets("����-����").Cells(8, 5).Value = relicName & "����"
            Else
                Sheets("����-����").Cells(8, 5).Value = Sheets("����-����").Cells(8, 5).Value & Chr(10) & relicName & "����"
            End If
            '����������Դ������д
            If Sheets("����-����").Cells(8, 6).Value = "" Then
                Sheets("����-����").Cells(8, 6).Value = expSingle
            Else
                Sheets("����-����").Cells(8, 6).Value = Sheets("����-����").Cells(8, 6).Value & Chr(10) & expSingle
            End If
            'Ħ������������д
            If Sheets("����-����").Cells(2, 3).Value = "" Then
                Sheets("����-����").Cells(2, 3).Value = molaSingle
            Else
                Sheets("����-����").Cells(2, 3).Value = Sheets("����-����").Cells(2, 3).Value + molaSingle
            End If
            'Ħ��������Դ��д
            If Sheets("����-����").Cells(2, 5).Value = "" Then
                Sheets("����-����").Cells(2, 5).Value = relicName & "����"
            Else
                Sheets("����-����").Cells(2, 5).Value = Sheets("����-����").Cells(2, 5).Value & Chr(10) & relicName & "����"
            End If
            'Ħ��������Դ������д
            If Sheets("����-����").Cells(2, 6).Value = "" Then
                Sheets("����-����").Cells(2, 6).Value = molaSingle
            Else
                Sheets("����-����").Cells(2, 6).Value = Sheets("����-����").Cells(2, 6).Value & Chr(10) & molaSingle
            End If
        End If
        i = i + 1
    Loop
    '������Ϊ��ʱ����дʥ���ﾭ��ȱ������
    If Sheets("����-����").Cells(8, 3).Value <> "" Then
        '��Ҫ��������
        expRelic = Sheets("����-����").Cells(8, 3).Value
        '���о�������
        exp5 = Sheets("����-����").Cells(10, 2).Value
        exp4 = Sheets("����-����").Cells(9, 2).Value
        exp3 = Sheets("����-����").Cells(8, 2).Value
        exp2 = Sheets("����-����").Cells(7, 2).Value
        exp1 = Sheets("����-����").Cells(6, 2).Value
        expHave = exp5 * 3780 + exp4 * 2520 + exp3 * 1260 + exp2 * 840 + exp1 * 420
        'ȱ�ھ�������������Ӣ�۾������
        expNeed = expRelic - expHave
        If expNeed > 0 Then
            Sheets("����-����").Cells(8, 4).Value = expNeed
            Sheets("����-����").Cells(6, 4).Value = "�ۺ�" & Application.WorksheetFunction.RoundUp((expRelic - expHave) / 420, 0)
        Else
            Sheets("����-����").Cells(8, 4).Value = "����"
        End If
    End If
    
    'Part5 Ħ���Ͳ���ȱ����������
    Sheets("����-���ɵ���").Select
    'Part5-1 ����-���ɵ���ҳ����
    i = 5
    Do While Cells(i, 1) <> ""
        '����������Ϊ�ղŴ���
        If Cells(i, 5) <> "" Then
            '�ԿɺϳɺͲ��ɺϳ���ֱ���
            If Cells(i, 2) = "" Then
                needNumber = Cells(i, 5).Value - Cells(i, 4).Value
                If needNumber > 0 Then
                    Cells(i, 6).Value = needNumber
                Else
                    Cells(i, 6).Value = "����"
                End If
            Else
                '��ȡ�ɺϳɼ���
                itemName = Cells(i, 1).Value
                itemMax = Application.WorksheetFunction.CountIf(Sheets("����-���ɵ���").Range("A:A"), itemName)
                itemStart = Application.WorksheetFunction.Match(itemName, Sheets("����-���ɵ���").Range("A:A"), 0)
                itemEnd = itemStart + itemMax - 1
                Dim itemNeed()
                ReDim itemNeed(1 To itemMax)
                For j = 1 To itemMax
                    itemNeed(j) = 0
                Next j
                k = itemMax
                '����ϳ�ǰȱ�����
                For j = itemStart To itemEnd
                    '����ϳ�ǰȱ������
                    itemNeed(k) = Cells(j, 5).Value - Cells(j, 4).Value
                    '��д�ϳ�ǰȱ������
                    If Cells(j, 5) <> "" Then
                        If itemNeed(k) > 0 Then
                            Cells(j, 6).Value = itemNeed(k)
                        Else
                            Cells(j, 6).Value = "����"
                        End If
                    End If
                    k = k - 1
                Next j
                '����ϳɹ���
                Dim upRule()
                ReDim upRule(1 To itemMax)
                upRule(1) = 1
                For j = 2 To itemMax
                    upRule(j) = upRule(j - 1) * Cells(itemEnd - j + 1, 3).Value
                Next j
                '�ӵͼ����߼������ϳɲ���
                For j = 2 To itemMax
                    If itemNeed(j) > 0 Then
                        '��Ҫ�ϳɵ�����
                        compoundNeed = itemNeed(j) * upRule(j)
                        '֧�ֺϳɵ�����
                        compoundHave = 0
                        For k = 1 To j - 1
                            If itemNeed(k) < 0 Then
                                compoundHave = compoundHave - itemNeed(k) * upRule(k)
                            End If
                        Next k
                        '����ʵ�ʺϳɸ���������
                        compoundNumber = 0
                        If compoundHave >= upRule(j) Then
                            compoundNumber = Application.WorksheetFunction.Min(itemNeed(j), Int(compoundHave / upRule(j)))
                            compoundCostAll = compoundNumber * upRule(j)
                        End If
                        If compoundNumber <> 0 Then
                            '�Ӹߵ��Ϳ۳��ϳ�����
                            itemNeed(j) = itemNeed(j) - compoundNumber
                            For k = (j - 1) To 1 Step -1
                                If itemNeed(k) < 0 Then
                                    compoundCostSingle = Application.WorksheetFunction.Min(compoundCostAll, -(itemNeed(k) * upRule(k)))
                                    itemNeed(k) = itemNeed(k) + compoundCostSingle / upRule(k)
                                    compoundCostAll = compoundCostAll - compoundCostSingle
                                End If
                            Next k
                            '��д�ϳɸ�����ʾ
                            Cells(itemEnd - j + 1, 6).Value = Cells(itemEnd - j + 1, 6).Value & "���ɺϳ�" & compoundNumber & "��"
                        End If
                    End If
                Next j
                '�ɺϳ���i����
                i = itemEnd
            End If
        End If
        i = i + 1
    Loop
    'Part5-1 ����-����ҳ����
    Sheets("����-����").Select
    'Ħ��ȱ����������
    If Cells(2, 3).Value <> "" Then
        needNumber = Cells(2, 3).Value - Cells(2, 2).Value
        If needNumber > 0 Then
            Cells(2, 4).Value = needNumber
        Else
            Cells(2, 4).Value = "����"
        End If
        Else
    End If
    '�ز�ȱ�ڼ���
    i = 11
    Do While Cells(i, 1) <> ""
        If Cells(i, 3) <> "" Then
            needNumber = Cells(i, 3).Value - Cells(i, 2).Value
            If needNumber > 0 Then
                Cells(i, 4).Value = needNumber
            Else
                Cells(i, 4).Value = "����"
            End If
        End If
        i = i + 1
    Loop
    
    'Part6 ��ʽ����
    Sheets("����-���ɵ���").Select
    '��ͷ��ʽ�ָ�
    Range("D1").Select
    Selection.Copy
    Range("E1:I1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    '���뷽ʽ�ָ�
    Columns("F:F").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With
    Range("F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Columns("H:H").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With
    Range("H1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Columns("I:I").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With
    Range("I1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    '�п��Զ�����
    Columns("F:F").EntireColumn.AutoFit
    Cells(1, 1).Select
    '��������ȱ��ʱ����Ⱦɫ
    If (Cells(3, 6).Value <> "" And Cells(3, 6).Value <> "����") Then
        Cells(3, 6).Select
        With Selection.Interior
            .Color = 13496575
        End With
    End If
    '��������ȱ��ʱ����Ⱦɫ
    i = 5
    Do While Cells(i, 1).Value <> ""
        If Cells(i, 6).Value <> "" Then
            needNumber = Cells(i, 6).Value
            If Right(needNumber, 1) = "��" Then
                'case1 ���˵�ʱ��������
            ElseIf Right(needNumber, 1) = "��" Then
                'case2 �кϳ��Ƽ�ʱ�ж��Ƿ�Ҫ��
                originNumberPos = InStr(needNumber, "��")
                compoundNumberPos = InStr(needNumber, "��")
                originNumber = Left(needNumber, originNumberPos - 1)
                compoundNumber = Right(needNumber, Len(needNumber) - compoundNumberPos)
                compoundNumber = Left(compoundNumber, Len(compoundNumber) - 1)
                If CInt(originNumber) > CInt(compoundNumber) Then
                    Cells(i, 6).Select
                    With Selection.Interior
                        .Color = 13496575
                    End With
                End If
            Else
                'case3 Ϊ����ʱֱ��Ⱦɫ
                Cells(i, 6).Select
                With Selection.Interior
                    .Color = 13496575
                End With
            End If
        End If
        i = i + 1
    Loop
    Application.ScreenUpdating = True
    Range("A2").Select
    Cells(1, 1).Select
    '����ҳ����ͬ����
    Sheets("����-����").Select
    Range("B1").Select
    Selection.Copy
    Range("C1:F1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With
    Range("D1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Columns("E:E").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With
    Range("E1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Columns("F:F").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With
    Range("F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    '�п��Զ�����
    Columns("D:D").EntireColumn.AutoFit
    '��Ħ����ȱ��ʱ����Ⱦɫ
    If (Cells(2, 4).Value <> "" And Cells(2, 4).Value <> "����") Then
        Cells(2, 4).Select
        With Selection.Interior
            .Color = 13496575
        End With
    End If
    '������������ȱ��ʱ����Ⱦɫ
    If (Cells(4, 4).Value <> "" And Cells(4, 4).Value <> "����") Then
        Cells(4, 4).Select
        With Selection.Interior
            .Color = 13496575
        End With
    End If
    '��ʥ���ﾭ����ȱ��ʱ����Ⱦɫ
    If (Cells(8, 4).Value <> "" And Cells(8, 4).Value <> "����") Then
        Cells(8, 4).Select
        With Selection.Interior
            .Color = 13496575
        End With
    End If
    i = 11
    Do While Cells(i, 1).Value <> ""
        If (Cells(i, 4).Value <> "" And Cells(i, 4).Value <> "����") Then
            Cells(i, 4).Select
            With Selection.Interior
                .Color = 13496575
            End With
        End If
        i = i + 1
    Loop
    Range("A2").Select
    Cells(1, 1).Select
    
    '���������ʾ
    Application.Interactive = True
    Sheets("����-���ɵ���").Select
    MsgBox ("����")
    
End Sub



