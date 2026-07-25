Attribute VB_Name = "modDesignsInit"
Option Explicit

Private Const DESIGNS_DOMAIN_CONTRACT_VERSION As String = "R1-DESIGNS-1"

Public Sub Auto_Open()
    ' Designs Domain is a background engine. Loading the XLAM must not inspect,
    ' activate, classify, or mutate any open role/operator workbook.
End Sub

Public Function GetDesignsDomainContractVersion() As String
    GetDesignsDomainContractVersion = DESIGNS_DOMAIN_CONTRACT_VERSION
End Function

Public Function DiagnoseDesignsDomain() As String
    DiagnoseDesignsDomain = _
        "ContractVersion=" & DESIGNS_DOMAIN_CONTRACT_VERSION & _
        "|Workbook=" & ThisWorkbook.Name & _
        "|IsAddin=" & CStr(ThisWorkbook.IsAddin) & _
        "|StartupMutation=False" & _
        "|Authority=WHx.invSys.Data.Designs.xlsb"
End Function
