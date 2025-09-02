VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_choices 
   Caption         =   "Start Menu"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   -75
   ClientWidth     =   21075
   OleObjectBlob   =   "form_choices.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_choices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btn_cancel_Click()

    Unload Me

End Sub

Private Sub btn_run_Click()

    Dim picked_designations() As String
    picked_designations = collect_input_designations()
    
    If option_caGrid = True Then
        Call BillingGridToCAFileConverterSynkronizerPrepared(picked_designations, checkbox_remove_footnotes.Value, checkbox_trim_procedure_names.Value, checkbox_sort.Value)
    ElseIf option_internalBudgetGridFrequencies = True Then
        Call BillingGridToInternalBudgetGridFrequencies(picked_designations, checkbox_remove_footnotes.Value, checkbox_trim_procedure_names.Value, checkbox_sort.Value)
    ElseIf option_internalBudgetGridTotals = True Then
        Call BillingGridToInternalBudgetGridTotals(picked_designations, checkbox_remove_footnotes.Value, checkbox_trim_procedure_names.Value, checkbox_sort.Value)
    End If
    
    Unload Me
    MsgBox ("The file has been prepared and is open as a separate Excel workbook!"), vbInformation         'added sound to alert the user that import has finished

End Sub

Private Sub UserForm_Initialize()

    Height = 350
    Width = 855
    option_caGrid = True
    designation_NNR.Value = True
    designation_S0.Value = True
    designation_S1.Value = True
    designation_floating.Value = True
    checkbox_trim_procedure_names = True
    checkbox_remove_footnotes = True
    checkbox_sort = False

End Sub

Private Function collect_input_designations()
    
    Dim picked_designations() As String
    
    picked_designations = Split("-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1", ",")

    If designation_S0.Value = True Then
        picked_designations(0) = "S0"
    End If

    If designation_S1.Value = True Then
        picked_designations(1) = "S1"
    End If

    If designation_R.Value = True Then
        picked_designations(2) = "R"
    End If

    If designation_RF.Value = True Then
        picked_designations(3) = "R(F)"
    End If

    If designation_RCL.Value = True Then
        picked_designations(4) = "R(CL)"
    End If

    If designation_NNA.Value = True Then
        picked_designations(5) = "N(NA)"
    End If

    If designation_NNB.Value = True Then
        picked_designations(6) = "N(NB)"
    End If

    If designation_NNR.Value = True Then
        picked_designations(7) = "N(NR)"
    End If

    If designation_NNU.Value = True Then
        picked_designations(8) = "N(NU)"
    End If

    If designation_NWO.Value = True Then
        picked_designations(9) = "N(WO)"
    End If

    If designation_floating.Value = True Then
        picked_designations(10) = "floating"
    End If
    
    collect_input_designations = picked_designations

End Function
