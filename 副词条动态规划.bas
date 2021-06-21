Attribute VB_Name = "副词条动态规划"

Sub 副词条动态规划()
Attribute 副词条动态规划.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 动态规划 宏
'
    Worksheets("伤害计算").Activate

    SolverReset
    SolverOptions precision:=0.000001
    
    SolverOk SetCell:=Range("O25"), _
        MaxMinVal:=1, _
        ByChange:=Range("P2:P11"), _
        Engine:=1, _
        EngineDesc:="GRG Nonlinear"

    SolverAdd CellRef:=Range("P12"), _
        Relation:=2, _
        FormulaText:=Range("Q12")
    SolverAdd CellRef:=Range("P2:P11"), _
        Relation:=3, _
        FormulaText:=0

    SolverSolve UserFinish:=True
    
    SolverReset
    
    MsgBox ("好了")

End Sub
