            ' temporary patch for general fiction short stories collections
            CS.GetField "948", 1, sValue
            If InStr(sValue, "808.831") <> 0 Then
               Msgbox "Effective July 1 2018, NYPL has ceased to use '808.831' for general collections of short stories. Use the 'FIC' call number instead. Your record has not been exported."
               GoTo Done
            End If