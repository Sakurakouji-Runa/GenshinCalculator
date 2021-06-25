Attribute VB_Name = "副词条动态规划"

Sub 副词条动态规划()
Attribute 副词条动态规划.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 动态规划 宏
'
    '选择 伤害计算 工作表
    Worksheets("伤害计算").Activate
    '重置solver规划求解模块
    SolverReset
    '设置规划求解计算精度
    SolverOptions precision:=0.000001
    
    '规划求解 中的 设置目标 和 可变单元格
    SolverOk SetCell:=Range("O25"), _
        MaxMinVal:=1, _
        ByChange:=Range("P2:P11"), _
        Engine:=1, _
        EngineDesc:="GRG Nonlinear"
    '规划求解 中的 遵守约束
    SolverAdd CellRef:=Range("P12"), _
        Relation:=2, _
        FormulaText:=Range("Q12")
    SolverAdd CellRef:=Range("P2:P11"), _
        Relation:=3, _
        FormulaText:=0

    '不显示 规划求解 对话框
    SolverSolve UserFinish:=True
    
    SolverReset
    
    MsgBox ("好了")

End Sub
