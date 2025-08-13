Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Imports VB = Microsoft.VisualBasic
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.Colors
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions
Imports IOEx

Public Class Form1
    Public db As Database = HostApplicationServices.WorkingDatabase
    Public doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    Public ed As Editor = doc.Editor  'Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

    Private m_list As DriveListEx
    Private m_lis As DriveInfoEx
    Public IDFLAG As Boolean
    Public zbid As String

    Public Sub New()

        ' 此调用是 Windows 窗体设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        With Grid1
            .Cell(0, 1).Text = "桩号"
            .Cell(0, 2).Text = "面积"
        End With

    End Sub

    Private Sub Grid1_CellChange(ByVal Sender As Object, ByVal e As FlexCell.Grid.CellChangeEventArgs) Handles Grid1.CellChange
        With Grid1
            If .Cell(.Rows - 1, 1).Text <> "" Then .AddItem("")
        End With
    End Sub


    Private Sub Grid1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles Grid1.Click
        With Grid1
            Select Case .ActiveCell.Col
                Case 1
                    .ContextMenuStrip = Me.ContextMenuStrip2
                Case 2
                    .ContextMenuStrip = Me.ContextMenuStrip1
            End Select
        End With
    End Sub

    ''' <summary>
    ''' 框选
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请选择多段线")


        Dim acDocEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor


        Dim acTypValAr(0) As TypedValue
        acTypValAr.SetValue(New TypedValue(DxfCode.Start, "*POLYLINE"), 0)

        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)


        Dim acSSPrompt As PromptSelectionResult
        acSSPrompt = acDocEd.GetSelection(acSelFtr)


        If acSSPrompt.Status = PromptStatus.OK Then
            Dim acSSet As SelectionSet = acSSPrompt.Value


            ed.WriteMessage(vbCrLf & "共选择" & acSSet.Count.ToString & "个图形")

            Using trans As Transaction = db.TransactionManager.StartTransaction

                Dim i As Short
                For i = 0 To acSSet.Count - 1

                    Dim ent As Entity = trans.GetObject(acSSet.Item(i).ObjectId, OpenMode.ForRead)

                    Dim pl As Polyline = CType(ent, Polyline)
                    If Me.CheckBox1.Checked = True Then
                        If pl.Closed = True Then
                            With Grid1
                                If .Cell(.Rows - 1, 2).Text <> "" Then
                                    .AddItem("")
                                End If
                                .Cell(.Rows - 1, 2).Text = Math.Round(pl.Area / Me.TextBox1.Text, 4)
                            End With
                        End If
                    Else
                        With Grid1
                            If .Cell(.Rows - 1, 2).Text <> "" Then
                                .AddItem("")
                            End If
                            .Cell(.Rows - 1, 2).Text = Math.Round(pl.Area / Me.TextBox1.Text, 4)
                        End With
                    End If

                Next
            End Using
        Else
            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("没有合适的图形")
        End If
        Me.Focus()
    End Sub

   

    ''' <summary>
    ''' 获取桩号
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Try
            Dim acDocEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim acTypValAr(0) As TypedValue

            acTypValAr.SetValue(New TypedValue(DxfCode.Start, "TEXT"), 0)

            Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)


            Dim acSSPrompt As PromptSelectionResult
            acSSPrompt = acDocEd.GetSelection(acSelFtr)
            Using trans As Transaction = db.TransactionManager.StartTransaction
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    Dim i As Short
                    For i = 0 To acSSet.Count - 1
                        Dim ent As Entity = trans.GetObject(acSSet.Item(i).ObjectId, OpenMode.ForRead)
                        Dim pl As DBText = CType(ent, DBText)
                        With Grid1

                            .Cell(.ActiveCell.Row, 1).Text = pl.TextString
                        End With
                    Next
                Else
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Numberof objects selected: 0")
                End If

            End Using
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' '选择类型
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim db As Database = doc.Database

        Dim ed As Editor = doc.Editor

        Dim peo As New PromptEntityOptions("请选择一个实体")

        Dim per As PromptEntityResult = Nothing
        Try
            per = ed.GetEntity(peo)
            If per.Status = PromptStatus.OK Then

                Dim id As ObjectId = per.ObjectId

                Dim trans As Transaction = db.TransactionManager.StartTransaction()

                Dim ent As Entity = DirectCast(trans.GetObject(id, OpenMode.ForRead, True), Entity)
                ed.WriteMessage((vbLf & "实体ObjectId为：" & ent.ObjectId.ToString & vbLf & "实体类型为：") & ent.[GetType]().FullName)
                trans.Commit()
                trans.Dispose()
            End If
        Catch exc As Autodesk.AutoCAD.Runtime.Exception
            ed.WriteMessage("发生异常，原因为：" + exc.Message)
        End Try

    End Sub






#Region "查询方式"
    ''' <summary>
    ''' 查询方式=桩号+面积
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        Dim i As Double
        Dim yesno As Short

        For i = 1 To 100000
            yesno = MsgBox("进行面积查询", 4, "系统提示")
            If yesno = 6 Then
                ksmjcx()
            Else
                Exit Sub
            End If
        Next
        Me.Focus()
    End Sub

    ''' <summary>
    ''' 循环开始=桩号+面积
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ksmjcx()
        Call zh()
        Call ddx()
        Me.Focus()
    End Sub

    ''' <summary>
    ''' 获取桩号
    ''' </summary>
    ''' <remarks></remarks>
    Sub zh()
        Dim zhwz() As String
        Dim zhwz1 As String
        Try
            Dim ED As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("点击桩号文字")

            Dim acDocEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim acTypValAr(0) As TypedValue


            acTypValAr.SetValue(New TypedValue(DxfCode.Start, "*TEXT"), 0)

            Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)

            Dim acSSPrompt As PromptSelectionResult
            acSSPrompt = acDocEd.GetSelection(acSelFtr)

            Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction


                If acSSPrompt.Status = PromptStatus.OK Then

                    Dim acSSet As SelectionSet = acSSPrompt.Value

                    For Each selObj In acSSet
                        Dim ent As Entity = DirectCast(trans.GetObject(selObj.ObjectId, OpenMode.ForRead, True), Entity)


                        Select Case ent.[GetType]().Name

                            Case "DBText"
                                Dim textObj As DBText = trans.GetObject(selObj.ObjectId, OpenMode.ForRead, False, True)
                                'MsgBox("Text = " & textObj.TextString)
                                If textObj.TextString.Contains("{") Then
                                    zhwz = textObj.TextString.Split(";")
                                    '^[0-9]*[\+\-\*][0-9]*[\+\-\*\.][0-9]*
                                    zhwz1 = WZFZ(zhwz(1).Replace("}", ""))
                                Else
                                    zhwz1 = textObj.TextString
                                End If

                                With Grid1
                                    For j As Short = 1 To .Rows - 1
                                        If .Cell(j, 1).Text = "" Then
                                            .Cell(j, 1).Text = zhwz1.Replace("断面", "").Replace("横", "").Replace("图", "").Replace("桩号", "")
                                        End If
                                    Next


                                End With
                                textObj.Dispose()
                            Case "MText"
                                Dim mtextObj As MText = trans.GetObject(selObj.ObjectId, OpenMode.ForRead, False, True)

                                If mtextObj.Contents.Contains("{") Then
                                    zhwz = mtextObj.Contents.Split(";")
                                    '^[0-9]*[\+\-\*][0-9]*[\+\-\*\.][0-9]*
                                    zhwz1 = WZFZ(zhwz(1).Replace("}", "")).Replace("断面", "").Replace("横", "").Replace("图", "").Replace("桩号", "")
                                Else
                                    If mtextObj.Contents.Contains("\P") Then
                                        zhwz = mtextObj.Contents.Split("\P")
                                        zhwz1 = WZFZ(zhwz(0).Replace("}", "")).Replace("断面", "").Replace("横", "").Replace("图", "").Replace("桩号", "")
                                    Else
                                        zhwz1 = WZFZ(mtextObj.Contents).Replace("断面", "").Replace("横", "").Replace("图", "").Replace("桩号", "")

                                    End If
                                End If


                                With Grid1
                                    For j As Short = 1 To .Rows - 1
                                        If .Cell(j, 1).Text = "" Then
                                            .Cell(j, 1).Text = zhwz1 ' mtextObj.Contents.Replace("\P", vbCrLf)
                                        End If
                                    Next

                                End With
                                mtextObj.Dispose()
                        End Select
                    Next selObj
                Else
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("你没有选择合适的文字")
                    With Grid1
                        For j As Short = 1 To .Rows - 1
                            If .Cell(j, 1).Text = "" Then
                                .Cell(j, 1).Text = "0+000"
                            End If
                        Next


                    End With
                End If

            End Using

            '
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    ''' <summary>
    ''' 正则表达式0+000提取
    ''' </summary>
    ''' <param name="Str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WZFZ(ByVal Str As String) As String

        Dim strPattern As String = "^[0-9]*[\+\-\*][0-9]*[\+\-\*\.][0-9]*"

        '声明一个不可变的正则表达式 

        Dim oRegex As New Regex(strPattern, RegexOptions.Multiline)

        '声明一个表示返回的单个匹配值 

        Dim oMatch As Match

        '声明一个表示所有匹配值得集合 

        Dim oMatches As MatchCollection

        If oRegex.IsMatch(Str) = True Then

            oMatches = oRegex.Matches(Str)

            For Each oMatch In oMatches
                Dim strTemp As String = oMatch.Value
                Str = Strings.Replace(Str, strTemp, "", , 1)
            Next

        End If

        Return Str

    End Function

    ''' <summary>
    ''' 获取面积
    ''' </summary>
    ''' <remarks></remarks>
    Sub ddx()
        Try
            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请选择多段线")
            Dim ED As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            '线集合
            Dim summj As Double = 0
            Dim sl As Double = 0
            Dim polyList As New List(Of ObjectId)



            Dim ts As String = vbCrLf & "请选择多段线："

            Dim opt As New PromptEntityOptions(ts)

            opt.SetRejectMessage(vbCrLf & "你选择的不是多段线！")

            opt.AddAllowedClass(GetType(Polyline), True)



            Dim Res As PromptEntityResult = ED.GetEntity(opt)

            While Res.Status = PromptStatus.OK
                sl += 1
                ED.WriteMessage("共选择" & sl & "个实体")
                'Me.Label1.Text = "共选择" & sl & "个实体"
                Me.Refresh()
                polyList.Add(Res.ObjectId)
                Res = ED.GetEntity(opt)

            End While

            If polyList.Count > 0 Then
                summj = 0
                For Each pLID As ObjectId In polyList

                    '锁定文档

                    'Using docLoc As DocumentLock = doc.LockDocument

                    Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction

                        Try

                            Dim ent As Entity = trans.GetObject(pLID, OpenMode.ForRead)

                            Dim pl As Polyline = CType(ent, Polyline)

                            '如果是闭合的曲线
                            If Me.CheckBox1.Checked = True Then
                                If pl.Closed = True Then summj += pl.Area
                            Else
                                summj += pl.Area
                            End If


                        Catch ex As Exception
                            MsgBox(ex.ToString)
                            'ED.WriteMessage(vbCrLf & ex.ToString)

                        End Try

                        trans.Commit()

                    End Using

                    'End Using

                Next
                With Grid1
                    For j As Short = 1 To .Rows - 1
                        If .Cell(j, 1).Text = "" Then
                            .Cell(j - 1, 2).Text = Math.Round(summj / Val(Me.TextBox1.Text) / Val(Me.TextBox1.Text), 4)
                        End If
                    Next

                End With
            Else

                ED.WriteMessage(vbCrLf & "<<<你没有选择线>>>")

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    ''' <summary>
    ''' 单个查询
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Try

            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请选择多段线")
            Dim ED As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            '线集合
            Dim summj As Double = 0
            Dim sl As Double = 0
            Dim polyList As New List(Of ObjectId)

            Dim ts As String = vbCrLf & "请选择多段线："

            Dim opt As New PromptEntityOptions(ts)

            opt.SetRejectMessage(vbCrLf & "你选择的不是多段线！")

            opt.AddAllowedClass(GetType(Polyline), True)


            Dim Res As PromptEntityResult = ED.GetEntity(opt)

            While Res.Status = PromptStatus.OK
                sl += 1
                ed.WriteMessage("共选择" & sl & "个实体")
                Me.Refresh()
                polyList.Add(Res.ObjectId)
                Res = ed.GetEntity(opt)

            End While

            If polyList.Count > 0 Then
                summj = 0
                For Each pLID As ObjectId In polyList

                    '锁定文档



                    Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction

                        Try

                            Dim ent As Entity = trans.GetObject(pLID, OpenMode.ForRead)

                            Dim pl As Polyline = CType(ent, Polyline)

                            '如果是闭合的曲线
                            If Me.CheckBox1.Checked = True Then
                                If pl.Closed = True Then summj += pl.Area
                            Else
                                summj += pl.Area
                            End If

                            'If pl.Closed = True Then
                            ' summj += pl.Area

                            'Else

                            'ed.WriteMessage(vbCrLf & "<<<不是闭合的曲线，请先闭合>>>")

                            'End If



                        Catch ex As Exception

                            ed.WriteMessage(vbCrLf & ex.ToString)

                        End Try

                        trans.Commit()

                    End Using



                Next
                With Grid1

                    .Cell(.ActiveCell.Row, 2).Text = Math.Round(summj / Val(Me.TextBox1.Text) / Val(Me.TextBox1.Text), 4)
                End With
            Else

                ed.WriteMessage(vbCrLf & "<<<你没有选择线>>>")

            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' 查询方式=面积
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        kxcxmj()
    End Sub


    ''' <summary>
    ''' 框选查询面积
    ''' </summary>
    ''' <remarks></remarks>
    Sub kxcxmj()
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请选择多段线")
        Dim ED As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        '' Get the current document editor
        Dim acDocEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        '' Create a TypedValue array to define thefilter criteria
        Dim acTypValAr(0) As TypedValue
        acTypValAr.SetValue(New TypedValue(DxfCode.Start, "*POLYLINE"), 0)
        'acTypValAr.SetValue(New TypedValue(DxfCode.Start, "CIRCLE"), 1)
        'Circle
        '' Assign the filter criteria to aSelectionFilter object
        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)

        '' Request for objects to be selected in thedrawing area
        Dim acSSPrompt As PromptSelectionResult
        acSSPrompt = acDocEd.GetSelection(acSelFtr)

        '' If the prompt status is OK, objects wereselected
        If acSSPrompt.Status = PromptStatus.OK Then
            Dim acSSet As SelectionSet = acSSPrompt.Value

            'Application.ShowAlertDialog("Number ofobjects selected: " & acSSet.Count.ToString())
            acDocEd.WriteMessage(vbCrLf & "共选择" & acSSet.Count.ToString & "个图形")

            Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction

                Dim i As Short
                For i = 0 To acSSet.Count - 1
                    ' MsgBox(acSSet.Item(i).ObjectId.ToString)
                    Dim ent As Entity = trans.GetObject(acSSet.Item(i).ObjectId, OpenMode.ForRead)

                    Dim pl As Polyline = CType(ent, Polyline)
                    If Me.CheckBox1.Checked = True Then
                        If pl.Closed = True Then

                            'MsgBox(pl.Area)

                            With Grid1
                                If .Cell(.Rows - 1, 2).Text <> "" Then
                                    .AddItem("")
                                End If
                                .Cell(.Rows - 1, 2).Text = Math.Round(pl.Area / Val(Me.TextBox1.Text) / Val(Me.TextBox1.Text), 4)
                            End With
                        End If
                    Else
                        With Grid1
                            If .Cell(.Rows - 1, 2).Text <> "" Then
                                .AddItem("")
                            End If
                            .Cell(.Rows - 1, 2).Text = Math.Round(pl.Area / Val(Me.TextBox1.Text) / Val(Me.TextBox1.Text), 4)
                        End With
                    End If

                Next
            End Using
        Else
            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("没有合适的图形")
        End If
        Me.Focus()
    End Sub


    ''' <summary>
    ''' 面积标注-多个，未用到
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请选择多段线")
            Dim ED As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor


            Dim acDocEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            Dim acTypValAr(0) As TypedValue
            acTypValAr.SetValue(New TypedValue(DxfCode.Start, "*POLYLINE"), 0)

            Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)


            Dim acSSPrompt As PromptSelectionResult
            acSSPrompt = acDocEd.GetSelection(acSelFtr)


            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value


                acDocEd.WriteMessage(vbCrLf & "共选择" & acSSet.Count.ToString & "个图形")

                Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction

                    Dim i As Short
                    For i = 0 To acSSet.Count - 1
                        ' MsgBox(acSSet.Item(i).ObjectId.ToString)
                        Dim id As ObjectId = acSSet.Item(i).ObjectId
                        Dim ent As Entity = DirectCast(trans.GetObject(id, OpenMode.ForRead, True), Entity)

                        ' Dim ent As Entity = trans.GetObject(acSSet.Item(i).ObjectId, OpenMode.ForRead)

                        Dim pl As Polyline = CType(ent, Polyline)
                        If Me.CheckBox1.Checked = True Then
                            If pl.Closed = True Then
                                ' getAllVertices(ent)
                            End If

                        Else
                            'getAllVertices(ent)
                        End If

                    Next
                End Using
            Else
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("没有合适的图形")
            End If
            Me.Focus()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region

    Private Sub BG_HNT_KeyDown(ByVal Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Grid1.KeyDown
        If e.Control And e.KeyCode = Keys.C Then
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' 快速勾边
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click

        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim pd As Point3d

        Try

        
            Call layerxj()

            Dim prPointOptions As PromptPointOptions = New PromptPointOptions("\n请选择范围内的点：")
            Dim prPointRes As PromptPointResult = ed.GetPoint(prPointOptions)
            pd = prPointRes.Value
            ' GetColsedBoundary(pd.X, pd.Y)
            '"_-Boundary" & vbCr & Pt(0) & "," & Pt(1) & vbCr & vbCr

            doc.SendStringToExecute("_-Boundary" & vbCr & pd.X & "," & pd.Y & vbCr & vbCr, True, False, False)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    <Runtime.InteropServices.DllImport("acad.exe", SetLastError:=True)> _
  Private Shared Function acedCmd(ByVal vlist As System.IntPtr) As Integer
    End Function

    Public Shared Function GetColsedBoundary(ByVal InnerX As Double, ByVal InnerY As Double) As ObjectId
        Dim IslandDetection As Short = 0
        Dim rb As ResultBuffer = New ResultBuffer
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database


        Try
            doc.SendStringToExecute("_-Boundary" & vbCr & InnerX & "," & InnerX & vbCr & vbCr, True, False, False)

            acedCmd(rb.UnmanagedObject)

            Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim psr As PromptSelectionResult = ed.SelectLast
            Dim ObjIds As ObjectId() = psr.Value.GetObjectIds
            If ObjIds IsNot Nothing And ObjIds.Length > 0 Then
                Dim ObjId As ObjectId = ObjIds(0)
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Dim obj As DBObject = trans.GetObject(ObjId, OpenMode.ForRead)
                    If TypeOf (obj) Is Polyline Then
                        Dim ObjectPolyPolyline As Polyline = obj
                        Dim msg As String = String.Format("Area = {0} ", ObjectPolyPolyline.Area)
                        ed.WriteMessage(msg)
                        Return ObjId
                    End If
                End Using

            End If

        Catch ex As System.Exception
            MsgBox("Hatch Boundary Error " & ex.Message)
            Return ObjectId.Null
        Finally
            rb.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 图层新建
    ''' </summary>
    ''' <remarks></remarks>
    Sub layerxj()
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim LayerName As String = "华宇_面积"
        Try
            Using doc.LockDocument()

                Using trans As Transaction = db.TransactionManager.StartTransaction()
                    Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForWrite), LayerTable)
                    Dim layerId As ObjectId
                    If lt.Has(LayerName.Trim()) = False Then
                        Dim ltr As New LayerTableRecord
                        ltr.Name = LayerName.Trim
                        ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 90)
                        layerId = lt.Add(ltr)
                        '将图层记录添加到层表中         '将图层表记录添加到事务处理中                    
                        trans.AddNewlyCreatedDBObject(ltr, True)

                        db.Clayer = ltr.ObjectId
                    Else

                    End If
                    trans.Commit()
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' 面积标注-单个
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 图层ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 图层ToolStripMenuItem.Click
        Try
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim db As Database = doc.Database

            Dim ed As Editor = doc.Editor

            Dim peo As New PromptEntityOptions("请选择一个实体")

            Dim per As PromptEntityResult = Nothing

            per = ed.GetEntity(peo)
            If per.Status = PromptStatus.OK Then

                Dim id As ObjectId = per.ObjectId

                Dim trans As Transaction = db.TransactionManager.StartTransaction()

                Dim ent As Entity = DirectCast(trans.GetObject(id, OpenMode.ForRead, True), Entity)


                Dim opt As PromptPointOptions = New PromptPointOptions("选择插入点")
                Dim res As PromptPointResult = ed.GetPoint(opt)


                getAllVertices(ent, res.Value.X, res.Value.Y)
                'ed.WriteMessage((vbLf & "实体ObjectId为：" & ent.ObjectId.ToString & vbLf & "实体类型为：") & ent.[GetType]().FullName)
                trans.Commit()
                trans.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' 获取图层列表
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetLayers() As String()

        Using Trans As Transaction = DB.TransactionManager.StartTransaction
            Dim LayT As LayerTable = Trans.GetObject(DB.LayerTableId, OpenMode.ForRead)
            Dim ID As ObjectId, LayTR As LayerTableRecord, Lays As New List(Of String)
            For Each ID In LayT
                LayTR = Trans.GetObject(ID, OpenMode.ForRead)
                Lays.Add(LayTR.Name)
            Next
            Trans.Commit()
            GetLayers = Lays.ToArray
        End Using
    End Function

    ''' <summary>
    ''' 写入面积数据
    ''' </summary>
    ''' <param name="ent"></param>
    ''' <param name="XX"></param>
    ''' <param name="YY"></param>
    ''' <remarks></remarks>
    Public Sub getAllVertices(ByVal ent As Polyline, ByVal XX As String, ByVal YY As String) 'As Point2dCollection
        Try

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Using doc.LockDocument()
                Dim mytext As New DBText

                mytext.Position = New Point3d(XX, YY, 0)
                Dim Db As Database = HostApplicationServices.WorkingDatabase


                Using Trans As Transaction = Db.TransactionManager.StartTransaction
                    'Dim TSB As TextStyleTable = Trans.GetObject(Db.TextStyleTableId, OpenMode.ForRead)
                    mytext.TextString = "面积：" & Math.Round(ent.Area / TextBox1.Text, 4)
                    'MsgBox("面积：" & Math.Round(ent.Area / TextBox1.Text, 4))
                    'mytext.TextStyle = TSB("工程图")
                    'mytext.Height = 10
                    Dim spc As BlockTableRecord = Trans.GetObject(Db.CurrentSpaceId, OpenMode.ForWrite, False)
                    spc.AppendEntity(mytext)
                    Trans.AddNewlyCreatedDBObject(mytext, True)
                    'Trans.AddNewlyCreatedDBObject(mytext, True)
                    Trans.Commit()

                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        'Return verCollection
    End Sub

    Sub Read3DPolyline(ByVal poly3d As Polyline3d)
        '' Start a transaction
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()

            Dim vId As ObjectId
            For Each vId In poly3d
                Dim v3d As PolylineVertex3d = acTrans.GetObject(vId, OpenMode.ForRead)
                doc.Editor.WriteMessage(vbCrLf + v3d.Position.ToString())
            Next
            acTrans.Commit()
        End Using
    End Sub

    Public Sub addtext(ByVal x As Double, ByRef y As Double)
        Try

            doc.LockDocument()


            Dim ed As Editor = doc.Editor

            dim opt as  PromptPointOptions = new PromptPointOptions("选择插入点")
            Dim res As PromptPointResult = ed.GetPoint(opt)


            Dim mytext As New DBText
            mytext.TextString = "测试"
            mytext.Position = New Point3d(res.Value.X, res.Value.Y, 0)
            Dim Db As Database = HostApplicationServices.WorkingDatabase
            Using Trans As Transaction = Db.TransactionManager.StartTransaction
                'Dim TSB As TextStyleTable = Trans.GetObject(Db.TextStyleTableId, OpenMode.ForRead)
                mytext.TextString = "工程图"
                'mytext.TextStyle = TSB.Item("工程图")
                Dim spc As BlockTableRecord = Trans.GetObject(Db.CurrentSpaceId, OpenMode.ForWrite)
                spc.AppendEntity(mytext)
                Trans.AddNewlyCreatedDBObject(mytext, True)
                Trans.Commit()
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

   

    ''' <summary>
    ''' 注册
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem14.Click
        Dim a As New ZCCT
        a.ShowDialog()
    End Sub
    ''' <summary>
    ''' 导出到EXCEL
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click
        If Regsfzc() = True Then
            'MsgBox("用户已经注册，欢迎使用华宇系列软件", 1 + 64, "系统提示")
            Grid1.ExportToExcel("")
        Else
            'MsgBox("请注册，才能够获得更多功能", 1 + 64, "系统提示")

        End If
    End Sub
    ''' <summary>
    ''' 打印
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ToolStripMenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem11.Click
        If Regsfzc() = True Then
            'MsgBox("用户已经注册，欢迎用华宇系列软件", 1 + 64, "系统提示")
            Grid1.Print()
        Else
            'MsgBox("请注册，才能够获得更多功能", 1 + 64, "系统提示")

        End If
    End Sub

    Private Sub ToolStripMenuItem7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Using doc.LockDocument()
                Dim optPt1 As New PromptPointOptions(vbLf & "请选取第一个点:")
                Dim resPt1 As PromptPointResult = ed.GetPoint(optPt1)
                If resPt1.Status <> PromptStatus.OK Then
                    Return
                End If
                Dim pt1 As Point3d = resPt1.Value

                Dim optPt2 As New PromptPointOptions(vbLf & "请选取第二个点:")
                optPt2.BasePoint = pt1
                optPt2.UseBasePoint = True
                Dim resPt2 As PromptPointResult = ed.GetPoint(optPt2)
                If resPt2.Status <> PromptStatus.OK Then
                    Return
                End If
                Dim pt2 As Point3d = resPt2.Value

                MsgBox(Math.Sqrt((pt1.X - pt2.X) ^ 2 + (pt1.Y - pt2.Y) ^ 2 + (pt1.Z - pt2.Z) ^ 2))

                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead, False), BlockTable)
                    Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite, False), BlockTableRecord)
                    'pt1->pt2方向的单位矢量
                    Dim vec1 As Vector3d = (pt2 - pt1).GetNormal()

                    '垂直方向的向量,长度为5
                    Dim vec2 As Vector3d = 5 * vec1.RotateBy(Math.PI / 2, Vector3d.ZAxis)
                    Dim lines As New List(Of Line)()
                    lines.Add(New Line(pt1 - 2 * vec1, pt2 + 2 * vec1))
                    lines.Add(New Line(pt1 + vec2, pt2 + vec2))
                    lines.Add(New Line(pt1 - vec2, pt2 - vec2))
                    For Each line As Line In lines
                        btr.AppendEntity(line)
                        tr.AddNewlyCreatedDBObject(line, True)
                    Next
                    tr.Commit()
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripMenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem15.Click
        Dim KSCK As New AboutBox1  ''打开窗口
        KSCK.ShowDialog()

    End Sub

    Private Sub ToolStripMenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem13.Click

    End Sub

    Private Sub longToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles longToolStripMenuItem.Click
        Try
            Dim distanceoptions As New PromptDistanceOptions("请选择第一个点：")
            Dim PriDisRes As PromptDoubleResult
            PriDisRes = ed.GetDistance(distanceoptions)
            If PriDisRes.Status <> PromptStatus.OK Then
                ed.WriteMessage("对不起，有错误，请重新输入")
            Else
                ed.WriteMessage("这条直线的距离是：" & PriDisRes.Value.ToString())
            End If
        Catch ex As System.Exception
            ed.WriteMessage("存在错误" & ex.Message)
        End Try
    End Sub
End Class


