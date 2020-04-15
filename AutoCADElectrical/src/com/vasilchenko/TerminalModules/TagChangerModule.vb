
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.TerminalModules
    Module TagChangerModule

        Friend Sub ChangeTag()
            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

            Dim acFilterList(0) As TypedValue
            acFilterList.SetValue(New TypedValue(CInt(DxfCode.Start), "INSERT"), 0)
            Dim acSelectionFilter As New SelectionFilter(acFilterList)

            Dim acPromptPntOpt As New PromptSelectionOptions With {
                .MessageForAdding = vbLf & "Выбери элементы"
            }

            Dim acPromptSelRes As PromptSelectionResult = acEditor.GetSelection(acPromptPntOpt, acSelectionFilter)

            If acPromptSelRes.Status <> PromptStatus.OK Then
                acEditor.WriteMessage("Не выбран элемент" & vbCrLf)
                Exit Sub
            Else
                Const i As Short = 1
                While i <> 0
                    Dim ufTagChanger As New ufTagChanger
                    ufTagChanger.ShowDialog()
                    If ufTagChanger.tbReplace.TextLength <> 0 AndAlso ufTagChanger.tbSearch.TextLength <> 0 Then
                        Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction

                            For Each acSelSetValue In acPromptSelRes.Value
                                Dim acBlockRef As BlockReference = acTransaction.GetObject(acSelSetValue.ObjectId, OpenMode.ForWrite)
                                For Each id As ObjectId In acBlockRef.AttributeCollection
                                    Dim acAttrReference As AttributeReference = acTransaction.GetObject(id, OpenMode.ForWrite)
                                    Select Case acAttrReference.Tag
                                        Case "TAG1", "TAG1F", "TAG_1", "TAG2", "TAG_2"
                                            acAttrReference.TextString = acAttrReference.TextString.Replace(ufTagChanger.tbSearch.Text, ufTagChanger.tbReplace.Text)
                                    End Select
                                Next
                            Next
                            acEditor.Regen()
                            acTransaction.Commit()
                            ufTagChanger.Close()
                        End Using
                    Else
                        Exit Sub
                    End If
                End While
            End If
        End Sub

    End Module
End Namespace
