Imports AutoCADTerminalBuilder.com.vasilchenko.Classes
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.ModifyTerminal
    Module RedrawTerminalModule

        Friend Sub StartRedraw()
            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

            Dim i As Short = 1
            While i <> 0
                Dim acPromptSelOpt As New PromptSelectionOptions()
                acPromptSelOpt.MessageForAdding = vbLf & "Select Terminal for Redraw:"
                acPromptSelOpt.SingleOnly = True
                Dim acResult As PromptSelectionResult = acEditor.GetSelection(acPromptSelOpt)

                If acResult.Status <> PromptStatus.OK Then
                    If i = 1 Then Application.ShowAlertDialog("Try to select a block next time.")
                    Exit Sub
                End If

                Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
                    Dim acBlckRef As BlockReference = acTransaction.GetObject(acResult.Value(0).ObjectId, OpenMode.ForRead)

                    Dim strTagstrp = "", strNum = "", strCat = "", strLocation = ""
                    Dim dblWidth As Double = 0, dblHeight As Double = 0

                    For Each id As ObjectId In acBlckRef.AttributeCollection
                        Dim acAttrbReference As AttributeReference = acTransaction.GetObject(id, OpenMode.ForRead)
                        Select Case acAttrbReference.Tag
                            Case "P_TAGSTRIP"
                                strTagstrp = acAttrbReference.TextString
                            Case "TERM"
                                strNum = acAttrbReference.TextString
                            Case "LOC"
                                strLocation = acAttrbReference.TextString
                            Case "CAT"
                                strCat = acAttrbReference.TextString
                            Case "WIDTH"
                                dblWidth = CDbl(acAttrbReference.TextString)
                            Case "HEIGHT"
                                dblHeight = CDbl(acAttrbReference.TextString)
                        End Select
                    Next


                    Dim newTerminal As MultilevelTerminalClass = New AcadConnector().FillTerminalData(strTagstrp, strNum, strLocation)

                    If newTerminal.Catalog <> strCat Then
                        Dim response = MsgBox("Каталожный номер изменен! Перестроить клеммник полностью?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Information)
                        If response = MsgBoxResult.Yes Then
                            MsgBox("Пока не реализовано. Удалите собственноручно =)")
                        ElseIf response = MsgBoxResult.No Then
                            RedrawCurrentBlock(acDatabase, acEditor, acTransaction, newTerminal, acBlckRef, dblWidth, dblHeight)
                        ElseIf response = MsgBoxResult.Cancel Then
                            Exit Sub
                        End If
                    Else
                        RedrawCurrentBlock(acDatabase, acEditor, acTransaction, newTerminal, acBlckRef, dblWidth, dblHeight)
                    End If
                    acTransaction.Commit()
                End Using
                i = 2
            End While
        End Sub

        Private Sub RedrawCurrentBlock(acDatabase As Database, acEditor As Editor, acTransaction As Transaction, newTerminal As MultilevelTerminalClass, acBlckRef As BlockReference, dblWidth As Double, dblHeight As Double)
            Dim acPointPositionRight = New Point3d(acBlckRef.Position.X + (dblWidth / 2), acBlckRef.Position.Y - (dblHeight / 2), 0)
            Dim acPointPositionLeft = New Point3d(acBlckRef.Position.X - (dblWidth / 2), acBlckRef.Position.Y - (dblHeight / 2), 0)
            RemoveLine(acDatabase, acEditor, acPointPositionLeft, acPointPositionRight)
            Dim objAcadConnetctor As New AcadConnector()
            objAcadConnetctor.DrawTerminalBlock(acDatabase, acTransaction, newTerminal, acBlckRef.Position, acBlckRef)
            acBlckRef.UpgradeOpen()
            acBlckRef.Erase()
        End Sub

        Private Sub RemoveLine(acDatabase As Database, acEditor As Editor, acStartPointLeft As Point3d, acStartPointRight As Point3d)
            Dim acBlockTable As BlockTable = acDatabase.BlockTableId.GetObject(OpenMode.ForRead)
            Dim acBlockTableRecord As BlockTableRecord = acBlockTable(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)

            For Each objId As ObjectId In acBlockTableRecord
                Dim acEntity As Entity = objId.GetObject(OpenMode.ForRead)
                If TypeOf acEntity Is Line Then
                    Dim acLine As Line = TryCast(acEntity, Line)
                    If (Math.Round(acLine.StartPoint.X, 4) = Math.Round(acStartPointRight.X, 4) And
                        Math.Round(acLine.StartPoint.Y, 4) = Math.Round(acStartPointRight.Y, 4)) OrElse
                        (Math.Round(acLine.EndPoint.X, 4) = Math.Round(acStartPointRight.X, 4) And
                        Math.Round(acLine.EndPoint.Y, 4) = Math.Round(acStartPointRight.Y, 4)) OrElse
                        (Math.Round(acLine.StartPoint.X, 4) = Math.Round(acStartPointLeft.X, 4) And
                        Math.Round(acLine.StartPoint.Y, 4) = Math.Round(acStartPointLeft.Y, 4)) OrElse
                        (Math.Round(acLine.EndPoint.X, 4) = Math.Round(acStartPointLeft.X, 4) And
                        Math.Round(acLine.EndPoint.Y, 4) = Math.Round(acStartPointLeft.Y, 4)) Then

                        acLine.UpgradeOpen()
                        acLine.Erase()
                    End If
                End If
            Next
        End Sub
    End Module
End Namespace
