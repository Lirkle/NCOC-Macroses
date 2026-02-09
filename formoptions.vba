Option Explicit

Public ColorLines As Boolean
Public OnlyThisSlide As Boolean
Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = True
    chkColorLines.Value = False
    chkOnlyThisSlide.Value = False
End Sub

Private Sub btnOK_Click()
    ColorLines = (chkColorLines.Value = True)
    OnlyThisSlide = (chkOnlyThisSlide.Value = True)
    Cancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Закрытие крестиком = Cancel
    Cancelled = True
End Sub
