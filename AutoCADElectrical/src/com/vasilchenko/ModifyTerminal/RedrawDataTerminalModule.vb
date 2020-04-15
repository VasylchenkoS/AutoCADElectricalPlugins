Imports System.Collections.Generic
Imports AutoCADTerminalBuilder.com.vasilchenko.Classes
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.ModifyTerminal
    Module RedrawDataTerminalModule

        Friend Sub StartRedrawData()
            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

            Dim i As Short = 1
            While i <> 0
                Dim acPromptSelOpt As New PromptSelectionOptions With {
                    .MessageForAdding = vbLf & "Select Terminal for Redraw:",
                    .SingleOnly = True
                }
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
                    Dim objSingleTerminal = newTerminal.Terminal.Values(0)

                    If newTerminal.Catalog <> strCat Then
                        Dim response = MsgBox("Каталожный номер изменен! Перестройте клеммник полностью", MsgBoxStyle.Critical + MsgBoxStyle.Information)
                    Else
                        For Each id As ObjectId In acBlckRef.AttributeCollection
                            Dim acAttrReference As AttributeReference = acTransaction.GetObject(id, OpenMode.ForRead)
                            Select Case acAttrReference.Tag
                                Case "WIRENOL"
                                    If objSingleTerminal.WiresLeftList.Count <> 0 AndAlso
                                        objSingleTerminal.WiresLeftList.Item(0).WireNumber.ToLower <> "pe" Then
                                        acAttrReference.TextString = objSingleTerminal.WiresLeftList.Item(0).WireNumber
                                    End If
                                Case "WIRENOR"
                                    If objSingleTerminal.WiresRigthList.Count <> 0 AndAlso
                                        objSingleTerminal.WiresRigthList.Item(0).WireNumber.ToLower <> "pe" Then
                                        acAttrReference.TextString = objSingleTerminal.WiresRigthList.Item(0).WireNumber
                                    End If
                                Case "TERMDESCL"
                                    Dim isCable = False
                                    If objSingleTerminal.WiresLeftList.Count <> 0 Then
                                        acAttrReference.TextString = SetTermdesc(objSingleTerminal.WiresLeftList, isCable)
                                    End If
                                Case "TERMDESCR"
                                    Dim isCable = False
                                    If objSingleTerminal.WiresRigthList.Count <> 0 Then
                                        acAttrReference.TextString = SetTermdesc(objSingleTerminal.WiresRigthList, isCable)
                                    End If
                            End Select

                            acBlckRef.AttributeCollection.AppendAttribute(acAttrReference)
                            acTransaction.AddNewlyCreatedDBObject(acAttrReference, True)
                        Next
                    End If
                    acTransaction.Commit()
                End Using
                i = 2
            End While
        End Sub
        Private Function SetTermdesc(wireList As List(Of WireClass), ByRef isCable As Boolean) As String
            Dim strTermdesc = ""
            If wireList.Any(Function(x As WireClass) x.HasCable) Then
                strTermdesc = "в " & wireList.Find(Function(x As WireClass) x.HasCable).Cable.Mark
                isCable = True
            Else
                For lngA = 0 To wireList.Count - 1
                    If wireList.Item(lngA).Termdesc <> "" Then
                        If strTermdesc = "" Then
                            strTermdesc = "к " & wireList.Item(lngA).Termdesc & ", "
                        Else
                            strTermdesc = strTermdesc & wireList.Item(lngA).Termdesc & ", "
                        End If
                    End If
                Next
                If strTermdesc <> "" Then strTermdesc = strTermdesc.Remove(strTermdesc.Length - 2)
            End If
            Return strTermdesc
        End Function
    End Module
End Namespace
