Attribute VB_Name = "��������̬�滮"

Sub ��������̬�滮()
Attribute ��������̬�滮.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��̬�滮 ��
'
    Worksheets("�˺�����").Activate

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
    
    MsgBox ("����")

End Sub
