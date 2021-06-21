Attribute VB_Name = "材料计算"

Sub 材料计算()
    
    Application.ScreenUpdating = False
    Application.Interactive = False
    
    'Part0 表格状态预处理
    Sheets("首页").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("背包-养成道具").Select
    ActiveSheet.AutoFilterMode = False
    '清除染色状态
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
    Sheets("背包-材料").Select
    ActiveSheet.AutoFilterMode = False
    '清除染色状态
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
    Sheets("角色").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("角色等级").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("角色材料").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("天赋材料").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("武器").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("武器等级").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("武器材料").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    Sheets("圣遗物").Select
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).Select
    '清除上一次的计算结果
    Sheets("背包-养成道具").Range("E:E").Clear
    Sheets("背包-养成道具").Range("E1") = "需求数量"
    Sheets("背包-养成道具").Range("F:F").Clear
    Sheets("背包-养成道具").Range("F1") = "缺口数量"
    Sheets("背包-养成道具").Range("H:H").Clear
    Sheets("背包-养成道具").Range("H1") = "需求来源"
    Sheets("背包-养成道具").Range("I:I").Clear
    Sheets("背包-养成道具").Range("I1") = "来源数量"
    Sheets("背包-材料").Range("C:C").Clear
    Sheets("背包-材料").Range("C1") = "需求数量"
    Sheets("背包-材料").Range("D:D").Clear
    Sheets("背包-材料").Range("D1") = "缺口数量"
    Sheets("背包-材料").Range("E:E").Clear
    Sheets("背包-材料").Range("E1") = "需求来源"
    Sheets("背包-材料").Range("F:F").Clear
    Sheets("背包-材料").Range("F1") = "来源数量"
    
    'Part1 人物升级计算
    Sheets("首页").Select
    '只要第一列不空就会一直算
    i = 2
    Do While Cells(i, 1) <> ""
        expSingle = 0
        molaSingle = 0
        roleName = Cells(i, 1).Value
        '是否已突破判断
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
        '只有当目标等级>当前等级时才计算
        If (lvEnd > lvStart Or (lvEnd = lvStart And lvEndAscensions = 1)) Then
            matchStart = Application.WorksheetFunction.Match(lvStart, Sheets("角色等级").Range("A:A"), 0)
            matchEnd = Application.WorksheetFunction.Match(lvEnd, Sheets("角色等级").Range("A:A"), 0)
            'Part1-1 人物升级经验摩拉
            expSingle = Sheets("角色等级").Cells(matchEnd, 2).Value - Sheets("角色等级").Cells(matchStart, 2).Value
            molaSingle = Sheets("角色等级").Cells(matchEnd, 3).Value - Sheets("角色等级").Cells(matchStart, 3).Value
            If lvStartAscensions = 0 Then
                molaSingle = molaSingle + Sheets("角色等级").Cells(matchStart, 6).Value
            End If
            If lvEndAscensions = 0 Then
                molaSingle = molaSingle - Sheets("角色等级").Cells(matchEnd, 6).Value
            End If
            If expSingle <> 0 Then
                '经验需求数量填写
                If Sheets("背包-养成道具").Cells(3, 5).Value = "" Then
                    Sheets("背包-养成道具").Cells(3, 5).Value = expSingle
                Else
                    Sheets("背包-养成道具").Cells(3, 5).Value = Sheets("背包-养成道具").Cells(3, 5).Value + expSingle
                End If
                '经验需求来源填写
                If Sheets("背包-养成道具").Cells(3, 8).Value = "" Then
                    Sheets("背包-养成道具").Cells(3, 8).Value = roleName & "升级"
                Else
                    Sheets("背包-养成道具").Cells(3, 8).Value = Sheets("背包-养成道具").Cells(3, 8).Value & Chr(10) & roleName & "升级"
                End If
                '经验需求来源数量填写
                If Sheets("背包-养成道具").Cells(3, 9).Value = "" Then
                    Sheets("背包-养成道具").Cells(3, 9).Value = expSingle
                Else
                    Sheets("背包-养成道具").Cells(3, 9).Value = Sheets("背包-养成道具").Cells(3, 9).Value & Chr(10) & expSingle
                End If
            End If
            If molaSingle <> 0 Then
            '摩拉需求数量填写
                If Sheets("背包-材料").Cells(2, 3).Value = "" Then
                    Sheets("背包-材料").Cells(2, 3).Value = molaSingle
                Else
                    Sheets("背包-材料").Cells(2, 3).Value = Sheets("背包-材料").Cells(2, 3).Value + molaSingle
                End If
                '摩拉需求来源填写
                If Sheets("背包-材料").Cells(2, 5).Value = "" Then
                    Sheets("背包-材料").Cells(2, 5).Value = roleName & "升级"
                Else
                    Sheets("背包-材料").Cells(2, 5).Value = Sheets("背包-材料").Cells(2, 5).Value & Chr(10) & roleName & "升级"
                End If
                '摩拉需求来源数量填写
                If Sheets("背包-材料").Cells(2, 6).Value = "" Then
                    Sheets("背包-材料").Cells(2, 6).Value = molaSingle
                Else
                    Sheets("背包-材料").Cells(2, 6).Value = Sheets("背包-材料").Cells(2, 6).Value & Chr(10) & molaSingle
                End If
            End If
            'Part1-2 人物升级材料
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
            '等级取整处理
            j = 2
            Do While lvStart >= Sheets("角色材料").Cells(j, 1).Value And Sheets("角色材料").Cells(j, 1).Value <> ""
                j = j + 1
            Loop
            If lvStart > Sheets("角色材料").Cells(j - 1, 1).Value Then
                lvStartAscensions = 1
            End If
            lvStart = Sheets("角色材料").Cells(j - 1, 1).Value
            j = 2
            Do While lvEnd >= Sheets("角色材料").Cells(j, 1).Value And Sheets("角色材料").Cells(j, 1).Value <> ""
                j = j + 1
            Loop
            If lvEnd > Sheets("角色材料").Cells(j - 1, 1).Value Then
                lvEndAscensions = 1
            End If
            lvEnd = Sheets("角色材料").Cells(j - 1, 1).Value
            matchStart = Application.WorksheetFunction.Match(lvStart, Sheets("角色材料").Range("A:A"), 0)
            matchEnd = Application.WorksheetFunction.Match(lvEnd, Sheets("角色材料").Range("A:A"), 0)
            '是否已突破处理
            matchStart = matchStart + lvStartAscensions
            matchEnd = matchEnd + lvEndAscensions
            '统计升级需要数量
            For j = matchStart To (matchEnd - 1)
                If Sheets("角色材料").Cells(j, 2).Value <> "" Then
                    eliteProb(Sheets("角色材料").Cells(j, 2).Value) = eliteProb(Sheets("角色材料").Cells(j, 2).Value) + Sheets("角色材料").Cells(j, 3).Value
                End If
                If Sheets("角色材料").Cells(j, 4).Value <> "" Then
                    eliteMust(Sheets("角色材料").Cells(j, 4).Value) = eliteMust(Sheets("角色材料").Cells(j, 4).Value) + Sheets("角色材料").Cells(j, 5).Value
                End If
                If Sheets("角色材料").Cells(j, 6).Value <> "" Then
                    gathering(Sheets("角色材料").Cells(j, 6).Value) = gathering(Sheets("角色材料").Cells(j, 6).Value) + Sheets("角色材料").Cells(j, 7).Value
                End If
                If Sheets("角色材料").Cells(j, 8).Value <> "" Then
                    normalDrop(Sheets("角色材料").Cells(j, 8).Value) = normalDrop(Sheets("角色材料").Cells(j, 8).Value) + Sheets("角色材料").Cells(j, 9).Value
                End If
            Next j
            '获取材料名称
            roleMatch = Application.WorksheetFunction.Match(roleName, Sheets("角色").Range("A:A"), 0)
            eliteProbName = Sheets("角色").Cells(roleMatch, 3).Value
            eliteMustName = Sheets("角色").Cells(roleMatch, 4).Value
            gatheringName = Sheets("角色").Cells(roleMatch, 5).Value
            normalDropName = Sheets("角色").Cells(roleMatch, 6).Value
            '升级材料位置查询
            eliteProbPos = Application.WorksheetFunction.Match(eliteProbName, Sheets("背包-养成道具").Range("A:A"), 0) + 3
            eliteMustPos = Application.WorksheetFunction.Match(eliteMustName, Sheets("背包-养成道具").Range("A:A"), 0)
            gatheringPos = Application.WorksheetFunction.Match(gatheringName, Sheets("背包-材料").Range("A:A"), 0)
            normalDropPos = Application.WorksheetFunction.Match(normalDropName, Sheets("背包-养成道具").Range("A:A"), 0) + 2
            '升级材料数量和来源填写
            For j = 1 To 4
                If eliteProb(j) <> 0 Then
                    If Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 5).Value = "" Then
                        Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 5).Value = eliteProb(j)
                    Else
                        Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 5).Value = Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 5).Value + eliteProb(j)
                    End If
                    If Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 8).Value = "" Then
                        Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 8).Value = roleName & "突破"
                    Else
                        Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 8).Value = Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 8).Value & Chr(10) & roleName & "突破"
                    End If
                    If Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 9).Value = "" Then
                        Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 9).Value = eliteProb(j)
                    Else
                        Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 9).Value = Sheets("背包-养成道具").Cells(eliteProbPos - j + 1, 9).Value & Chr(10) & eliteProb(j)
                    End If
                End If
            Next j
            If eliteMust(1) <> 0 Then
                If Sheets("背包-养成道具").Cells(eliteMustPos, 5).Value = "" Then
                    Sheets("背包-养成道具").Cells(eliteMustPos, 5).Value = eliteMust(1)
                Else
                    Sheets("背包-养成道具").Cells(eliteMustPos, 5).Value = Sheets("背包-养成道具").Cells(eliteMustPos, 5).Value + eliteMust(1)
                End If
                If Sheets("背包-养成道具").Cells(eliteMustPos, 8).Value = "" Then
                    Sheets("背包-养成道具").Cells(eliteMustPos, 8).Value = roleName & "突破"
                Else
                    Sheets("背包-养成道具").Cells(eliteMustPos, 8).Value = Sheets("背包-养成道具").Cells(eliteMustPos, 8).Value & Chr(10) & roleName & "突破"
                End If
                If Sheets("背包-养成道具").Cells(eliteMustPos, 9).Value = "" Then
                    Sheets("背包-养成道具").Cells(eliteMustPos, 9).Value = eliteMust(1)
                Else
                    Sheets("背包-养成道具").Cells(eliteMustPos, 9).Value = Sheets("背包-养成道具").Cells(eliteMustPos, 9).Value & Chr(10) & eliteMust(1)
                End If
            End If
            If gathering(1) <> 0 Then
                If Sheets("背包-材料").Cells(gatheringPos, 3).Value = "" Then
                    Sheets("背包-材料").Cells(gatheringPos, 3).Value = gathering(1)
                Else
                    Sheets("背包-材料").Cells(gatheringPos, 3).Value = Sheets("背包-材料").Cells(gatheringPos, 3).Value + gathering(1)
                End If
                If Sheets("背包-材料").Cells(gatheringPos, 5).Value = "" Then
                    Sheets("背包-材料").Cells(gatheringPos, 5).Value = roleName & "突破"
                Else
                    Sheets("背包-材料").Cells(gatheringPos, 5).Value = Sheets("背包-材料").Cells(gatheringPos, 5).Value & Chr(10) & roleName & "突破"
                End If
                If Sheets("背包-材料").Cells(gatheringPos, 6).Value = "" Then
                    Sheets("背包-材料").Cells(gatheringPos, 6).Value = gathering(1)
                Else
                    Sheets("背包-材料").Cells(gatheringPos, 6).Value = Sheets("背包-材料").Cells(gatheringPos, 6).Value & Chr(10) & gathering(1)
                End If
            End If
            For j = 1 To 3
                If normalDrop(j) <> 0 Then
                    If Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 5).Value = "" Then
                        Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 5).Value = normalDrop(j)
                    Else
                        Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 5).Value = Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 5).Value + normalDrop(j)
                    End If
                    If Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 8).Value = "" Then
                        Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 8).Value = roleName & "突破"
                    Else
                        Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 8).Value = Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 8).Value & Chr(10) & roleName & "突破"
                    End If
                    If Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 9).Value = "" Then
                        Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 9).Value = normalDrop(j)
                    Else
                        Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 9).Value = Sheets("背包-养成道具").Cells(normalDropPos - j + 1, 9).Value & Chr(10) & normalDrop(j)
                    End If
                End If
            Next j
        End If
        i = i + 1
    Loop
    '当需求不为空时，填写经验缺口数量
    If Sheets("背包-养成道具").Cells(3, 5).Value <> "" Then
        '需要经验数量
        expRole = Sheets("背包-养成道具").Cells(3, 5).Value
        '已有经验数量
        exp3 = Sheets("背包-养成道具").Cells(2, 4).Value
        exp2 = Sheets("背包-养成道具").Cells(3, 4).Value
        exp1 = Sheets("背包-养成道具").Cells(4, 4).Value
        expHave = exp3 * 20000 + exp2 * 5000 + exp1 * 1000
        '缺口经验数量，按大英雄经验计算
        expNeed = expRole - expHave
        If expNeed > 0 Then
            Sheets("背包-养成道具").Cells(3, 6).Value = expNeed
            Sheets("背包-养成道具").Cells(2, 6).Value = "折合" & Application.WorksheetFunction.RoundUp((expRole - expHave) / 20000, 0)
        Else
            Sheets("背包-养成道具").Cells(3, 6).Value = "好了"
        End If
    End If
    '摩拉缺口数量，材料缺口数量得最后一起算
    
    'Part2 人物天赋计算
    i = 2
    Do While Cells(i, 1) <> ""
        roleName = Cells(i, 1).Value
        '初始化各材料数量
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
        '获取材料名称
        roleMatch = Application.WorksheetFunction.Match(roleName, Sheets("角色").Range("A:A"), 0)
        talentBookName = Sheets("角色").Cells(roleMatch, 7).Value
        talentDropName = Sheets("角色").Cells(roleMatch, 8).Value
        weekDropName = Sheets("角色").Cells(roleMatch, 9).Value
        activityGetName = Sheets("角色").Cells(roleMatch, 10).Value
        '升级材料位置查询
        talentBookPos = Application.WorksheetFunction.Match(talentBookName, Sheets("背包-养成道具").Range("A:A"), 0) + 2
        talentDropPos = Application.WorksheetFunction.Match(talentDropName, Sheets("背包-养成道具").Range("A:A"), 0) + 2
        If weekDropName = "" Then
            weekDropPos = -1
        Else
            weekDropPos = Application.WorksheetFunction.Match(weekDropName, Sheets("背包-养成道具").Range("A:A"), 0)
        End If
        activityGetPos = Application.WorksheetFunction.Match(activityGetName, Sheets("背包-养成道具").Range("A:A"), 0)
        For j = 4 To 8 Step 2
            talentStart = Cells(i, j).Value
            talentEnd = Cells(i, j + 1).Value
            '只有当目标等级>当前等级时才计算
            If talentEnd > talentStart Then
                matchStart = Application.WorksheetFunction.Match(talentStart, Sheets("天赋材料").Range("A:A"), 0)
                matchEnd = Application.WorksheetFunction.Match(talentEnd, Sheets("天赋材料").Range("A:A"), 0)
                '统计天赋升级摩拉
                If matchStart = 2 Then
                    molaSingle = Sheets("天赋材料").Cells(matchEnd, 10).Value
                Else
                    molaSingle = Sheets("天赋材料").Cells(matchEnd, 10).Value - Sheets("天赋材料").Cells(matchStart, 10).Value
                End If
                molaTalent = molaTalent + molaSingle
                '统计天赋升级材料
                For k = (matchStart + 1) To matchEnd
                    If Sheets("天赋材料").Cells(k, 2).Value <> "" Then
                        talentBook(Sheets("天赋材料").Cells(k, 2).Value) = talentBook(Sheets("天赋材料").Cells(k, 2).Value) + Sheets("天赋材料").Cells(k, 3).Value
                    End If
                    If Sheets("天赋材料").Cells(k, 4).Value <> "" Then
                        talentDrop(Sheets("天赋材料").Cells(k, 4).Value) = talentDrop(Sheets("天赋材料").Cells(k, 4).Value) + Sheets("天赋材料").Cells(k, 5).Value
                    End If
                    If Sheets("天赋材料").Cells(k, 6).Value <> "" Then
                        weekDrop(Sheets("天赋材料").Cells(k, 6).Value) = weekDrop(Sheets("天赋材料").Cells(k, 6).Value) + Sheets("天赋材料").Cells(k, 7).Value
                    End If
                    If Sheets("天赋材料").Cells(k, 8).Value <> "" Then
                        activityGet(Sheets("天赋材料").Cells(k, 8).Value) = activityGet(Sheets("天赋材料").Cells(k, 8).Value) + Sheets("天赋材料").Cells(k, 9).Value
                    End If
                Next k
            End If
        Next j
        '升级材料数量和来源填写
        For j = 1 To 3
            If talentBook(j) <> 0 Then
                If Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 5).Value = "" Then
                    Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 5).Value = talentBook(j)
                Else
                    Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 5).Value = Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 5).Value + talentBook(j)
                End If
                If Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 8).Value = "" Then
                    Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 8).Value = roleName & "天赋"
                Else
                    Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 8).Value = Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 8).Value & Chr(10) & roleName & "天赋"
                End If
                If Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 9).Value = "" Then
                    Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 9).Value = talentBook(j)
                Else
                    Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 9).Value = Sheets("背包-养成道具").Cells(talentBookPos - j + 1, 9).Value & Chr(10) & talentBook(j)
                End If
            End If
        Next j
        For j = 1 To 3
            If talentDrop(j) <> 0 Then
                If Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 5).Value = "" Then
                    Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 5).Value = talentDrop(j)
                Else
                    Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 5).Value = Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 5).Value + talentDrop(j)
                End If
                If Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 8).Value = "" Then
                    Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 8).Value = roleName & "天赋"
                Else
                    Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 8).Value = Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 8).Value & Chr(10) & roleName & "天赋"
                End If
                If Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 9).Value = "" Then
                    Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 9).Value = talentDrop(j)
                Else
                    Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 9).Value = Sheets("背包-养成道具").Cells(talentDropPos - j + 1, 9).Value & Chr(10) & talentDrop(j)
                End If
            End If
        Next j
        If (weekDrop(1) <> 0 And weekDropPos > 0) Then
            If Sheets("背包-养成道具").Cells(weekDropPos, 5).Value = "" Then
                Sheets("背包-养成道具").Cells(weekDropPos, 5).Value = weekDrop(1)
            Else
                Sheets("背包-养成道具").Cells(weekDropPos, 5).Value = Sheets("背包-养成道具").Cells(weekDropPos, 5).Value + weekDrop(1)
            End If
            If Sheets("背包-养成道具").Cells(weekDropPos, 8).Value = "" Then
                Sheets("背包-养成道具").Cells(weekDropPos, 8).Value = roleName & "天赋"
            Else
                Sheets("背包-养成道具").Cells(weekDropPos, 8).Value = Sheets("背包-养成道具").Cells(weekDropPos, 8).Value & Chr(10) & roleName & "天赋"
            End If
            If Sheets("背包-养成道具").Cells(weekDropPos, 9).Value = "" Then
                Sheets("背包-养成道具").Cells(weekDropPos, 9).Value = weekDrop(1)
            Else
                Sheets("背包-养成道具").Cells(weekDropPos, 9).Value = Sheets("背包-养成道具").Cells(weekDropPos, 9).Value & Chr(10) & weekDrop(1)
            End If
        End If
        If activityGet(1) <> 0 Then
            If Sheets("背包-养成道具").Cells(activityGetPos, 5).Value = "" Then
                Sheets("背包-养成道具").Cells(activityGetPos, 5).Value = activityGet(1)
            Else
                Sheets("背包-养成道具").Cells(activityGetPos, 5).Value = Sheets("背包-养成道具").Cells(activityGetPos, 5).Value + activityGet(1)
            End If
            If Sheets("背包-养成道具").Cells(activityGetPos, 8).Value = "" Then
                Sheets("背包-养成道具").Cells(activityGetPos, 8).Value = roleName & "天赋"
            Else
                Sheets("背包-养成道具").Cells(activityGetPos, 8).Value = Sheets("背包-养成道具").Cells(activityGetPos, 8).Value & Chr(10) & roleName & "天赋"
            End If
            If Sheets("背包-养成道具").Cells(activityGetPos, 9).Value = "" Then
                Sheets("背包-养成道具").Cells(activityGetPos, 9).Value = activityGet(1)
            Else
                Sheets("背包-养成道具").Cells(activityGetPos, 9).Value = Sheets("背包-养成道具").Cells(activityGetPos, 9).Value & Chr(10) & activityGet(1)
            End If
        End If
        If molaTalent <> 0 Then
            '摩拉需求数量填写
            If Sheets("背包-材料").Cells(2, 3).Value = "" Then
                Sheets("背包-材料").Cells(2, 3).Value = molaTalent
            Else
                Sheets("背包-材料").Cells(2, 3).Value = Sheets("背包-材料").Cells(2, 3).Value + molaTalent
            End If
            '摩拉需求来源填写
            If Sheets("背包-材料").Cells(2, 5).Value = "" Then
                Sheets("背包-材料").Cells(2, 5).Value = roleName & "天赋"
            Else
                Sheets("背包-材料").Cells(2, 5).Value = Sheets("背包-材料").Cells(2, 5).Value & Chr(10) & roleName & "天赋"
            End If
            '摩拉来源数量填写
            If Sheets("背包-材料").Cells(2, 6).Value = "" Then
                Sheets("背包-材料").Cells(2, 6).Value = molaTalent
            Else
                Sheets("背包-材料").Cells(2, 6).Value = Sheets("背包-材料").Cells(2, 6).Value & Chr(10) & molaTalent
            End If
        End If
        i = i + 1
    Loop
    
    'Part3 武器升级计算
    i = 2
    Do While Cells(i, 11) <> ""
        weaponName = Cells(i, 11).Value
        '初始化各材料数量
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
        '获取材料名称
        weaponMatch = Application.WorksheetFunction.Match(weaponName, Sheets("武器").Range("B:B"), 0)
        dungenDropName = Sheets("武器").Cells(weaponMatch, 4).Value
        eliteThingName = Sheets("武器").Cells(weaponMatch, 5).Value
        normalThingName = Sheets("武器").Cells(weaponMatch, 6).Value
        '升级材料位置查询
        dungenDropPos = Application.WorksheetFunction.Match(dungenDropName, Sheets("背包-养成道具").Range("A:A"), 0) + 3
        eliteThingPos = Application.WorksheetFunction.Match(eliteThingName, Sheets("背包-养成道具").Range("A:A"), 0) + 2
        normalThingPos = Application.WorksheetFunction.Match(normalThingName, Sheets("背包-养成道具").Range("A:A"), 0) + 2
        '是否已突破判断
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
        '当目标等级>当前等级时才计算
        If (lvEnd > lvStart Or (lvEnd = lvStart And lvEndAscensions = 1)) Then
            '武器星级查询
            weaponStar = Sheets("武器").Cells(Application.WorksheetFunction.Match(weaponName, Sheets("武器").Range("B:B"), 0), 3).Value
            starMatch = Application.WorksheetFunction.Match(weaponStar, Sheets("武器等级").Range("A:A"), 0)
            weaponUpmax = Application.WorksheetFunction.CountIf(Sheets("武器等级").Range("A:A"), weaponStar)
            matchStart = Application.WorksheetFunction.Match(lvStart, Sheets("武器等级").Cells(starMatch, 2).Resize(weaponUpmax, 1), 0) + starMatch - 1
            matchEnd = Application.WorksheetFunction.Match(lvEnd, Sheets("武器等级").Cells(starMatch, 2).Resize(weaponUpmax, 1), 0) + starMatch - 1
            'Part3-1 武器升级经验摩拉
            expSingle = Sheets("武器等级").Cells(matchEnd, 3).Value - Sheets("武器等级").Cells(matchStart, 3).Value
            molaSingle = Sheets("武器等级").Cells(matchEnd, 4).Value - Sheets("武器等级").Cells(matchStart, 4).Value
            If expSingle <> 0 Then
                '经验需求数量填写
                If Sheets("背包-材料").Cells(4, 3).Value = "" Then
                    Sheets("背包-材料").Cells(4, 3).Value = expSingle
                Else
                    Sheets("背包-材料").Cells(4, 3).Value = Sheets("背包-材料").Cells(4, 3).Value + expSingle
                End If
                '经验需求来源填写
                If Sheets("背包-材料").Cells(4, 5).Value = "" Then
                    Sheets("背包-材料").Cells(4, 5).Value = weaponName & "升级"
                Else
                    Sheets("背包-材料").Cells(4, 5).Value = Sheets("背包-材料").Cells(4, 5).Value & Chr(10) & weaponName & "升级"
                End If
                '经验来源数量填写
                If Sheets("背包-材料").Cells(4, 6).Value = "" Then
                    Sheets("背包-材料").Cells(4, 6).Value = expSingle
                Else
                    Sheets("背包-材料").Cells(4, 6).Value = Sheets("背包-材料").Cells(4, 6).Value & Chr(10) & expSingle
                End If
            End If
            If molaSingle <> 0 Then
            '摩拉需求数量填写
                If Sheets("背包-材料").Cells(2, 3).Value = "" Then
                    Sheets("背包-材料").Cells(2, 3).Value = molaSingle
                Else
                    Sheets("背包-材料").Cells(2, 3).Value = Sheets("背包-材料").Cells(2, 3).Value + molaSingle
                End If
                '摩拉需求来源填写
                If Sheets("背包-材料").Cells(2, 5).Value = "" Then
                    Sheets("背包-材料").Cells(2, 5).Value = weaponName & "升级"
                Else
                    Sheets("背包-材料").Cells(2, 5).Value = Sheets("背包-材料").Cells(2, 5).Value & Chr(10) & weaponName & "升级"
                End If
                '摩拉来源数量填写
                If Sheets("背包-材料").Cells(2, 6).Value = "" Then
                    Sheets("背包-材料").Cells(2, 6).Value = molaSingle
                Else
                    Sheets("背包-材料").Cells(2, 6).Value = Sheets("背包-材料").Cells(2, 6).Value & Chr(10) & molaSingle
                End If
            End If
            'Part3-2 武器升级材料
            '是否已突破处理
            matchStart = matchStart + lvStartAscensions
            matchEnd = matchEnd + lvEndAscensions
            For j = matchStart To (matchEnd - 1)
                If Sheets("武器材料").Cells(j, 3).Value <> "" Then
                    dungenDrop(Sheets("武器材料").Cells(j, 3).Value) = dungenDrop(Sheets("武器材料").Cells(j, 3).Value) + Sheets("武器材料").Cells(j, 4).Value
                End If
                If Sheets("武器材料").Cells(j, 5).Value <> "" Then
                    eliteThing(Sheets("武器材料").Cells(j, 5).Value) = eliteThing(Sheets("武器材料").Cells(j, 5).Value) + Sheets("武器材料").Cells(j, 6).Value
                End If
                If Sheets("武器材料").Cells(j, 7).Value <> "" Then
                    normalThing(Sheets("武器材料").Cells(j, 7).Value) = normalThing(Sheets("武器材料").Cells(j, 7).Value) + Sheets("武器材料").Cells(j, 8).Value
                End If
            Next j
            '升级材料数量和来源填写
            For j = 1 To 4
                If dungenDrop(j) <> 0 Then
                    If Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 5).Value = "" Then
                        Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 5).Value = dungenDrop(j)
                    Else
                        Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 5).Value = Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 5).Value + dungenDrop(j)
                    End If
                    If Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 8).Value = "" Then
                        Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 8).Value = weaponName & "突破"
                    Else
                        Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 8).Value = Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 8).Value & Chr(10) & weaponName & "突破"
                    End If
                    If Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 9).Value = "" Then
                        Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 9).Value = dungenDrop(j)
                    Else
                        Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 9).Value = Sheets("背包-养成道具").Cells(dungenDropPos - j + 1, 9).Value & Chr(10) & dungenDrop(j)
                    End If
                End If
            Next j
            For j = 1 To 3
                If eliteThing(j) <> 0 Then
                    If Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 5).Value = "" Then
                        Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 5).Value = eliteThing(j)
                    Else
                        Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 5).Value = Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 5).Value + eliteThing(j)
                    End If
                    If Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 8).Value = "" Then
                        Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 8).Value = weaponName & "突破"
                    Else
                        Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 8).Value = Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 8).Value & Chr(10) & weaponName & "突破"
                    End If
                    If Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 9).Value = "" Then
                        Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 9).Value = eliteThing(j)
                    Else
                        Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 9).Value = Sheets("背包-养成道具").Cells(eliteThingPos - j + 1, 9).Value & Chr(10) & eliteThing(j)
                    End If
                End If
            Next j
            For j = 1 To 3
                If normalThing(j) <> 0 Then
                    If Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 5).Value = "" Then
                        Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 5).Value = normalThing(j)
                    Else
                        Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 5).Value = Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 5).Value + normalThing(j)
                    End If
                    If Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 8).Value = "" Then
                        Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 8).Value = weaponName & "突破"
                    Else
                        Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 8).Value = Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 8).Value & Chr(10) & weaponName & "突破"
                    End If
                    If Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 9).Value = "" Then
                        Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 9).Value = normalThing(j)
                    Else
                        Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 9).Value = Sheets("背包-养成道具").Cells(normalThingPos - j + 1, 9).Value & Chr(10) & normalThing(j)
                    End If
                End If
            Next j
        End If
        i = i + 1
    Loop
    '当需求不为空时，填写武器经验缺口数量
    If Sheets("背包-材料").Cells(4, 3).Value <> "" Then
        '需要经验数量
        expWeapon = Sheets("背包-材料").Cells(4, 3).Value
        '已有经验数量
        exp3 = Sheets("背包-材料").Cells(5, 2).Value
        exp2 = Sheets("背包-材料").Cells(4, 2).Value
        exp1 = Sheets("背包-材料").Cells(3, 2).Value
        expHave = exp3 * 10000 + exp2 * 2000 + exp1 * 400
        '缺口经验数量，按大英雄经验计算
        expNeed = expWeapon - expHave
        If expNeed > 0 Then
            Sheets("背包-材料").Cells(4, 4).Value = expNeed
            Sheets("背包-材料").Cells(5, 4).Value = "折合" & Application.WorksheetFunction.RoundUp((expWeapon - expHave) / 10000, 0)
        Else
            Sheets("背包-材料").Cells(4, 4).Value = "好了"
        End If
    End If
    
    'Part4 圣遗物升级计算
    i = 2
    Do While Cells(i, 15) <> ""
        relicName = Cells(i, 15).Value
        relicStar = Cells(i, 16).Value
        relicCount = Cells(i, 17).Value
        lvStart = Cells(i, 18).Value
        lvEnd = Cells(i, 19).Value
        expSingle = 0
        molaSingle = 0
         '只有当目标等级>当前等级时才计算
        If lvEnd > lvStart Then
            '获取升级数据位置
            starMatch = Application.WorksheetFunction.Match(relicStar, Sheets("圣遗物").Range("A:A"), 0)
            relicUpmax = Application.WorksheetFunction.CountIf(Sheets("圣遗物").Range("A:A"), relicStar)
            matchStart = Application.WorksheetFunction.Match(lvStart, Sheets("圣遗物").Cells(starMatch, 2).Resize(relicUpmax, 1), 0) + starMatch - 1
            matchEnd = Application.WorksheetFunction.Match(lvEnd, Sheets("圣遗物").Cells(starMatch, 2).Resize(relicUpmax, 1), 0) + starMatch - 1
            '升级需求的经验摩拉计算
            expSingle = Sheets("圣遗物").Cells(matchEnd, 3).Value - Sheets("圣遗物").Cells(matchStart, 3).Value
            molaSingle = Sheets("圣遗物").Cells(matchEnd, 4).Value - Sheets("圣遗物").Cells(matchStart, 4).Value
            expSingle = expSingle * relicCount
            molaSingle = molaSingle * relicCount
            '经验需求数量填写
            If Sheets("背包-材料").Cells(8, 3).Value = "" Then
                Sheets("背包-材料").Cells(8, 3).Value = expSingle
            Else
                Sheets("背包-材料").Cells(8, 3).Value = Sheets("背包-材料").Cells(8, 3).Value + expSingle
            End If
            '经验需求来源填写
            If Sheets("背包-材料").Cells(8, 5).Value = "" Then
                Sheets("背包-材料").Cells(8, 5).Value = relicName & "升级"
            Else
                Sheets("背包-材料").Cells(8, 5).Value = Sheets("背包-材料").Cells(8, 5).Value & Chr(10) & relicName & "升级"
            End If
            '经验需求来源数量填写
            If Sheets("背包-材料").Cells(8, 6).Value = "" Then
                Sheets("背包-材料").Cells(8, 6).Value = expSingle
            Else
                Sheets("背包-材料").Cells(8, 6).Value = Sheets("背包-材料").Cells(8, 6).Value & Chr(10) & expSingle
            End If
            '摩拉需求数量填写
            If Sheets("背包-材料").Cells(2, 3).Value = "" Then
                Sheets("背包-材料").Cells(2, 3).Value = molaSingle
            Else
                Sheets("背包-材料").Cells(2, 3).Value = Sheets("背包-材料").Cells(2, 3).Value + molaSingle
            End If
            '摩拉需求来源填写
            If Sheets("背包-材料").Cells(2, 5).Value = "" Then
                Sheets("背包-材料").Cells(2, 5).Value = relicName & "升级"
            Else
                Sheets("背包-材料").Cells(2, 5).Value = Sheets("背包-材料").Cells(2, 5).Value & Chr(10) & relicName & "升级"
            End If
            '摩拉需求来源数量填写
            If Sheets("背包-材料").Cells(2, 6).Value = "" Then
                Sheets("背包-材料").Cells(2, 6).Value = molaSingle
            Else
                Sheets("背包-材料").Cells(2, 6).Value = Sheets("背包-材料").Cells(2, 6).Value & Chr(10) & molaSingle
            End If
        End If
        i = i + 1
    Loop
    '当需求不为空时，填写圣遗物经验缺口数量
    If Sheets("背包-材料").Cells(8, 3).Value <> "" Then
        '需要经验数量
        expRelic = Sheets("背包-材料").Cells(8, 3).Value
        '已有经验数量
        exp5 = Sheets("背包-材料").Cells(10, 2).Value
        exp4 = Sheets("背包-材料").Cells(9, 2).Value
        exp3 = Sheets("背包-材料").Cells(8, 2).Value
        exp2 = Sheets("背包-材料").Cells(7, 2).Value
        exp1 = Sheets("背包-材料").Cells(6, 2).Value
        expHave = exp5 * 3780 + exp4 * 2520 + exp3 * 1260 + exp2 * 840 + exp1 * 420
        '缺口经验数量，按大英雄经验计算
        expNeed = expRelic - expHave
        If expNeed > 0 Then
            Sheets("背包-材料").Cells(8, 4).Value = expNeed
            Sheets("背包-材料").Cells(6, 4).Value = "折合" & Application.WorksheetFunction.RoundUp((expRelic - expHave) / 420, 0)
        Else
            Sheets("背包-材料").Cells(8, 4).Value = "好了"
        End If
    End If
    
    'Part5 摩拉和材料缺口数量计算
    Sheets("背包-养成道具").Select
    'Part5-1 背包-养成道具页处理
    i = 5
    Do While Cells(i, 1) <> ""
        '需求数量不为空才处理
        If Cells(i, 5) <> "" Then
            '对可合成和不可合成类分别处理
            If Cells(i, 2) = "" Then
                needNumber = Cells(i, 5).Value - Cells(i, 4).Value
                If needNumber > 0 Then
                    Cells(i, 6).Value = needNumber
                Else
                    Cells(i, 6).Value = "好了"
                End If
            Else
                '获取可合成级数
                itemName = Cells(i, 1).Value
                itemMax = Application.WorksheetFunction.CountIf(Sheets("背包-养成道具").Range("A:A"), itemName)
                itemStart = Application.WorksheetFunction.Match(itemName, Sheets("背包-养成道具").Range("A:A"), 0)
                itemEnd = itemStart + itemMax - 1
                Dim itemNeed()
                ReDim itemNeed(1 To itemMax)
                For j = 1 To itemMax
                    itemNeed(j) = 0
                Next j
                k = itemMax
                '计算合成前缺口情况
                For j = itemStart To itemEnd
                    '保存合成前缺口数量
                    itemNeed(k) = Cells(j, 5).Value - Cells(j, 4).Value
                    '填写合成前缺口数量
                    If Cells(j, 5) <> "" Then
                        If itemNeed(k) > 0 Then
                            Cells(j, 6).Value = itemNeed(k)
                        Else
                            Cells(j, 6).Value = "好了"
                        End If
                    End If
                    k = k - 1
                Next j
                '计算合成规则
                Dim upRule()
                ReDim upRule(1 To itemMax)
                upRule(1) = 1
                For j = 2 To itemMax
                    upRule(j) = upRule(j - 1) * Cells(itemEnd - j + 1, 3).Value
                Next j
                '从低级到高级搜索合成策略
                For j = 2 To itemMax
                    If itemNeed(j) > 0 Then
                        '需要合成的消耗
                        compoundNeed = itemNeed(j) * upRule(j)
                        '支持合成的消耗
                        compoundHave = 0
                        For k = 1 To j - 1
                            If itemNeed(k) < 0 Then
                                compoundHave = compoundHave - itemNeed(k) * upRule(k)
                            End If
                        Next k
                        '计算实际合成个数和消耗
                        compoundNumber = 0
                        If compoundHave >= upRule(j) Then
                            compoundNumber = Application.WorksheetFunction.Min(itemNeed(j), Int(compoundHave / upRule(j)))
                            compoundCostAll = compoundNumber * upRule(j)
                        End If
                        If compoundNumber <> 0 Then
                            '从高到低扣除合成消耗
                            itemNeed(j) = itemNeed(j) - compoundNumber
                            For k = (j - 1) To 1 Step -1
                                If itemNeed(k) < 0 Then
                                    compoundCostSingle = Application.WorksheetFunction.Min(compoundCostAll, -(itemNeed(k) * upRule(k)))
                                    itemNeed(k) = itemNeed(k) + compoundCostSingle / upRule(k)
                                    compoundCostAll = compoundCostAll - compoundCostSingle
                                End If
                            Next k
                            '填写合成个数提示
                            Cells(itemEnd - j + 1, 6).Value = Cells(itemEnd - j + 1, 6).Value & "（可合成" & compoundNumber & "）"
                        End If
                    End If
                Next j
                '可合成类i管理
                i = itemEnd
            End If
        End If
        i = i + 1
    Loop
    'Part5-1 背包-材料页处理
    Sheets("背包-材料").Select
    '摩拉缺口数量计算
    If Cells(2, 3).Value <> "" Then
        needNumber = Cells(2, 3).Value - Cells(2, 2).Value
        If needNumber > 0 Then
            Cells(2, 4).Value = needNumber
        Else
            Cells(2, 4).Value = "好了"
        End If
        Else
    End If
    '特产缺口计算
    i = 11
    Do While Cells(i, 1) <> ""
        If Cells(i, 3) <> "" Then
            needNumber = Cells(i, 3).Value - Cells(i, 2).Value
            If needNumber > 0 Then
                Cells(i, 4).Value = needNumber
            Else
                Cells(i, 4).Value = "好了"
            End If
        End If
        i = i + 1
    Loop
    
    'Part6 格式调整
    Sheets("背包-养成道具").Select
    '表头格式恢复
    Range("D1").Select
    Selection.Copy
    Range("E1:I1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    '对齐方式恢复
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
    '列宽自动适配
    Columns("F:F").EntireColumn.AutoFit
    Cells(1, 1).Select
    '当经验有缺口时进行染色
    If (Cells(3, 6).Value <> "" And Cells(3, 6).Value <> "好了") Then
        Cells(3, 6).Select
        With Selection.Interior
            .Color = 13496575
        End With
    End If
    '当材料有缺口时进行染色
    i = 5
    Do While Cells(i, 1).Value <> ""
        If Cells(i, 6).Value <> "" Then
            needNumber = Cells(i, 6).Value
            If Right(needNumber, 1) = "了" Then
                'case1 好了的时候不做处理
            ElseIf Right(needNumber, 1) = "）" Then
                'case2 有合成推荐时判断是否还要锄
                originNumberPos = InStr(needNumber, "（")
                compoundNumberPos = InStr(needNumber, "成")
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
                'case3 为数字时直接染色
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
    '材料页做相同处理
    Sheets("背包-材料").Select
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
    '列宽自动适配
    Columns("D:D").EntireColumn.AutoFit
    '当摩拉有缺口时进行染色
    If (Cells(2, 4).Value <> "" And Cells(2, 4).Value <> "好了") Then
        Cells(2, 4).Select
        With Selection.Interior
            .Color = 13496575
        End With
    End If
    '当武器经验有缺口时进行染色
    If (Cells(4, 4).Value <> "" And Cells(4, 4).Value <> "好了") Then
        Cells(4, 4).Select
        With Selection.Interior
            .Color = 13496575
        End With
    End If
    '当圣遗物经验有缺口时进行染色
    If (Cells(8, 4).Value <> "" And Cells(8, 4).Value <> "好了") Then
        Cells(8, 4).Select
        With Selection.Interior
            .Color = 13496575
        End With
    End If
    i = 11
    Do While Cells(i, 1).Value <> ""
        If (Cells(i, 4).Value <> "" And Cells(i, 4).Value <> "好了") Then
            Cells(i, 4).Select
            With Selection.Interior
                .Color = 13496575
            End With
        End If
        i = i + 1
    Loop
    Range("A2").Select
    Cells(1, 1).Select
    
    '计算完成提示
    Application.Interactive = True
    Sheets("背包-养成道具").Select
    MsgBox ("好了")
    
End Sub



