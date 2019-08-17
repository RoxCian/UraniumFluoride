Imports ExcelDna.Integration

Namespace Helper
#Disable Warning IDE0051
    Public Class ValueCircularListCollection
        Inherits CircularListCollection(Of Integer?)

        Public Sub New(listCount As Integer, listCapacity As Integer, value As Integer?)
            MyBase.New(listCount, listCapacity, value)
        End Sub
    End Class

    Public Class CircularListCollection(Of T)
        Inherits List(Of CircularList(Of T))
        Public ReadOnly Property ListCapacity As Integer
            Get
                If Me.Count = 0 Then Return 0 Else Return Me.Item(0).Capacity
            End Get
        End Property

        Private Sub New()
            MyBase.New
        End Sub
        Public Sub New(listCount As Integer, listCapacity As Integer)
            MyBase.New(listCount)
            For i = 0 To listCount - 1
                Me.Add(New CircularList(Of T)(listCapacity))
            Next
        End Sub
        Public Sub New(listCount As Integer, listCapacity As Integer, value As T)
            MyBase.New(listCount)
            For i = 0 To listCount - 1
                Me.Add(New CircularList(Of T)(listCapacity, value))
            Next
        End Sub
        Private Shadows Sub Add(item As CircularList(Of T))
            MyBase.Add(item)
        End Sub
        Private Shadows Sub RemoveAt(index As Integer)
            MyBase.RemoveAt(index)
        End Sub
    End Class
    Public Class CircularList(Of T)
        Inherits List(Of T)
        Default Public Shadows Property Item(index As Integer) As T
            Get
                Return MyBase.Item(Me.GetAbsoluteIndex(index))
            End Get
            Set(value As T)
                MyBase.Item(Me.GetAbsoluteIndex(index)) = value
            End Set
        End Property
        Public Property AbsolutePointerIndex As Integer
            Get
                Return _AbsolutePointerIndex
            End Get
            Private Set(value As Integer)
                Dim valueInRange As Integer = value - Math.Floor(value / Capacity) * Capacity
                If Me.Count < Me.Capacity AndAlso valueInRange > Me.Count Then
                    For i = Me.Count - 1 To valueInRange
                        If Not Me.ValuedIndex.Contains(i) Then Me.ValuedIndex.Add(i)
                    Next
                End If
                Me._AbsolutePointerIndex = valueInRange
            End Set
        End Property
        Public Property Current As T
            Get
                Return MyBase.Item(AbsolutePointerIndex)
            End Get
            Set(value As T)
                MyBase.Item(AbsolutePointerIndex) = value
                If Not Me.ValuedIndex.Contains(AbsolutePointerIndex) Then Me.ValuedIndex.Add(AbsolutePointerIndex)
            End Set
        End Property
        Public Overloads ReadOnly Property Count As Integer
            Get
                Return Me.ValuedIndex.Count
            End Get
        End Property
        Public Overloads ReadOnly Property Capacity As Integer
        Public Overloads ReadOnly Property CountUnderZero As Integer
            Get
                If Me.Count < Me.Capacity Then Return Me.AbsolutePointerIndex Else Return 0
            End Get
        End Property
        Public Overloads ReadOnly Property CountAboveZero As Integer
            Get
                If Me.Count < Me.Capacity Then Return Me.Count - Me.AbsolutePointerIndex - 1 Else Return Me.Count - 1
            End Get
        End Property
        Private Property ValuedIndex As New List(Of Integer)(Me.Capacity)

        Dim _AbsolutePointerIndex As Integer = 0

        Private Sub New()
            MyBase.New()
        End Sub
        Public Sub New(capacity As Integer)
            MyBase.New(capacity)
            For i = 0 To capacity - 1
                MyBase.Add(Nothing)
            Next
            Me.Capacity = capacity
        End Sub
        Public Sub New(capacity As Integer, value As T)
            MyBase.New(capacity)
            For i = 0 To capacity - 1
                MyBase.Add(value)
            Next
            Me.Capacity = capacity
        End Sub
        Private Sub New(collection As IEnumerable(Of T))
            MyBase.New(collection)
        End Sub

        Public Function GetAbsoluteIndex(relativeIndex As Integer) As Integer
            Dim absoluteIndex As Integer = Me.AbsolutePointerIndex + relativeIndex
            If Me.Count = 0 Then Return Me.AbsolutePointerIndex
            If absoluteIndex > Me.Count - 1 Then
                Do Until absoluteIndex <= Me.Count - 1
                    absoluteIndex -= Me.Count
                Loop
            ElseIf absoluteIndex < 0 Then
                Do Until absoluteIndex > 0
                    absoluteIndex += Me.Count
                Loop
            End If
            Return absoluteIndex
        End Function

        Public Sub MoveNext()
            If Me.ValuedIndex.Contains(Me.AbsolutePointerIndex) Then Me.AbsolutePointerIndex += 1
        End Sub
        Public Sub MovePrevious()
            Me.AbsolutePointerIndex -= 1
        End Sub
        Public Sub MoveNext(current As T)
            Me.Current = current
            Me.MoveNext()
        End Sub
        Public Sub MovePrevious(current As T)
            Me.Current = current
            Me.MovePrevious()
        End Sub
        Public Sub Skip([step] As Integer)
            Me.AbsolutePointerIndex += [step]
        End Sub

        Public Shadows Sub Add(item As T)
            Me.MoveNext()
            Me.Current = item
        End Sub

        Public Overloads Iterator Function GetEnumerator() As IEnumerator(Of T)
            For i = 0 To Me.Count - 1
                Yield MyBase.Item(i)
            Next
        End Function
    End Class

    Public Class SortingExcelCollection
        'Incompleted
    End Class

    Public Structure OpenInterval(Of T As IComparable(Of T))
        Public ReadOnly Property Left As T
        Public ReadOnly Property Right As T
        Public Sub New(left As T, right As T)

            Me.Left = left
            Me.Right = right
        End Sub
        Public Function Contains(value As T) As Boolean
            Return value.CompareTo(Me.Left) > 0 And value.CompareTo(Me.Right) < 0
        End Function
    End Structure
    Public Structure CloseInterval(Of T As IComparable(Of T))
        Public ReadOnly Property Left As T
        Public ReadOnly Property Right As T
        Public Sub New(left As T, right As T)

            Me.Left = left
            Me.Right = right
        End Sub
        Public Function Contains(value As T) As Boolean
            Return value.CompareTo(Me.Left) >= 0 And value.CompareTo(Me.Right) <= 0
        End Function
    End Structure
    Public Structure OpenInterval2(Of T As IComparable(Of T))
        Public ReadOnly Property Left As T
        Public ReadOnly Property Right As T
        Public ReadOnly Property Top As T
        Public ReadOnly Property Bottom As T
        Public Sub New(left As T, top As T, right As T, bottom As T)
            Me.Left = left
            Me.Right = right
            Me.Top = top
            Me.Bottom = bottom
        End Sub
        Public Function Contains(x As T, y As T) As Boolean
            Return x.CompareTo(Me.Left) > 0 And x.CompareTo(Me.Right) < 0 And y.CompareTo(Me.Top) > 0 And y.CompareTo(Me.Bottom) < 0
        End Function
    End Structure
    Public Structure CloseInterval2(Of T As IComparable(Of T))
        Public ReadOnly Property Left As T
        Public ReadOnly Property Right As T
        Public ReadOnly Property Top As T
        Public ReadOnly Property Bottom As T
        Public Sub New(left As T, top As T, right As T, bottom As T)
            Me.Left = left
            Me.Right = right
            Me.Top = top
            Me.Bottom = bottom
        End Sub
        Public Function Contains(x As T, y As T) As Boolean
            Return x.CompareTo(Me.Left) >= 0 And x.CompareTo(Me.Right) <= 0 And y.CompareTo(Me.Top) >= 0 And y.CompareTo(Me.Bottom) <= 0
        End Function
    End Structure

    Public Class ClosedXMLWorkbookLibrary
        Private Shared ReadOnly WorkbookCollection As New Dictionary(Of String, ClosedXML.Excel.XLWorkbook)
        Public Shared Function Create(path As String, Optional [alias] As String = "", Optional isReadOnly As Boolean = True) As ClosedXML.Excel.XLWorkbook
            If [alias] = "" Then [alias] = path
            If Not FileIO.FileSystem.FileExists(path) Then Return Nothing
            If Not WorkbookCollection.ContainsKey([alias]) Then WorkbookCollection.Add([alias], New ClosedXML.Excel.XLWorkbook(path, isReadOnly))
            Return WorkbookCollection([alias])
        End Function
        Private Sub New() : End Sub
    End Class

    Public Class CatmullRomSpline
        Public Const CentripetalAlpha As Single = 0.5
        Public Const UniformAlpha As Single = 0
        Public Const ChordalAlpha As Single = 1
        Public ReadOnly Property P As Numerics.Vector2()
        Public Property Alpha As Single
        Public ReadOnly Property T(index As Integer) As Single
            Get
                Static ResultDictionary As New Dictionary(Of Single, Single())
                If ResultDictionary.ContainsKey(Alpha) Then
                    Return ResultDictionary(Alpha)(index)
                Else
                    Dim tvalue(P.Count - 1) As Single
                    tvalue(0) = 0
                    For i = 1 To P.Count - 1
                        tvalue(i) = GetTParameter(tvalue(i - 1), P(i - 1), P(i), Alpha)
                    Next
                    ResultDictionary.Add(Alpha, tvalue)
                    Return tvalue(index)
                End If
            End Get
        End Property
        Public ReadOnly Property A(index As Integer, tValue As Single) As Numerics.Vector2
            Get
                Dim t0 = T(index - 1)
                Dim t1 = T(index)
                Dim p0 = P(index - 1)
                Dim p1 = P(index)
                Return (t1 - tValue) / (t1 - t0) * p0 + (tValue - t0) / (t1 - t0) * p1
            End Get
        End Property
        Private Function GetAFunc(index As Integer) As Func(Of Single, Numerics.Vector2)
            Return Function(tValue As Single)
                       Dim t0 = T(index - 1)
                       Dim t1 = T(index)
                       Dim p0 = P(index - 1)
                       Dim p1 = P(index)
                       Return (t1 - tValue) / (t1 - t0) * p0 + (tValue - t0) / (t1 - t0) * p1
                   End Function
        End Function
        Public ReadOnly Property B(index As Integer, tValue As Single) As Numerics.Vector2
            Get
                Dim t0 = T(index - 1)
                Dim t2 = T(index + 1)
                Dim a1 = A(index, tValue)
                Dim a2 = A(index + 1, tValue)
                Return (t2 - tValue) / (t2 - t0) * a1 + (tValue - t0) / (t2 - t0) * a2
            End Get
        End Property
        Private Function GetBFunc(index As Integer) As Func(Of Single, Numerics.Vector2)
            Return Function(tValue As Single)
                       Dim t0 = T(index - 1)
                       Dim t2 = T(index + 1)
                       Dim a1 = A(index, tValue)
                       Dim a2 = A(index + 1, tValue)
                       Return (t2 - tValue) / (t2 - t0) * a1 + (tValue - t0) / (t2 - t0) * a2
                   End Function
        End Function
        Public ReadOnly Property C(index As Integer, tValue As Single) As Numerics.Vector2
            Get
                Dim t1 = T(index)
                Dim t2 = T(index + 1)
                Dim b1 = B(index, tValue)
                Dim b2 = B(index + 1, tValue)
                Return (t2 - tValue) / (t2 - t1) * b1 + (tValue - t1) / (t2 - t1) * b2
            End Get
        End Property
        Private Function GetCFunc(index As Integer) As Func(Of Single, Numerics.Vector2)
            Return Function(tValue As Single)
                       Dim t1 = T(index)
                       Dim t2 = T(index + 1)
                       Dim b1 = B(index, tValue)
                       Dim b2 = B(index + 1, tValue)
                       Return (t2 - tValue) / (t2 - t1) * b1 + (tValue - t1) / (t2 - t1) * b2
                   End Function
        End Function

        Public Function GetPlot(tValue As Single) As Numerics.Vector2
            If tValue >= T(0) And tValue < T(1) Then Return B(1, tValue)
            For i = 1 To P.Count - 3
                If T(i) <= tValue And T(i + 1) > tValue Then Return C(i, tValue)
            Next
            If tValue >= T(P.Count - 2) And tValue <= T(P.Count - 1) Then Return B(P.Count - 2, tValue)
            Return New Numerics.Vector2(Single.MinValue)
        End Function
        Public Function GetAllPlots([step] As Single) As Numerics.Vector2()
            Dim result As New List(Of Numerics.Vector2)
            For tvalue = T(0) To T(P.Count - 1) Step [step]
                result.Add(GetPlot(tvalue))
            Next
            Return result.ToArray
        End Function
        Public Function GetYMaxPlot([step] As Single) As Numerics.Vector2
            Static allPlotsDictionary As New Dictionary(Of (Alpha As Single, [Step] As Single), Numerics.Vector2())
            Dim plots As Numerics.Vector2()
            If allPlotsDictionary.ContainsKey((Me.Alpha, [step])) Then plots = allPlotsDictionary((Me.Alpha, [step])) Else plots = GetAllPlots([step]).ToArray
            Dim pYMax As New Numerics.Vector2(Single.MinValue)
            For Each i In plots
                If i.Y > pYMax.Y Then pYMax = i
            Next
            Return pYMax
        End Function
        Public Function GetXMaxPlot([step] As Single) As Numerics.Vector2
            Static allPlotsDictionary As New Dictionary(Of (Alpha As Single, [Step] As Single), Numerics.Vector2())
            Dim plots As Numerics.Vector2()
            If allPlotsDictionary.ContainsKey((Me.Alpha, [step])) Then plots = allPlotsDictionary((Me.Alpha, [step])) Else plots = GetAllPlots([step]).ToArray
            Dim pXMax As New Numerics.Vector2(Single.MinValue)
            For Each i In plots
                If i.X > pXMax.X Then pXMax = i
            Next
            Return pXMax
        End Function
        Public Function GetXValue(y As Single, [step] As Single) As Double
            Static allPlotsDictionary As New Dictionary(Of (Alpha As Single, [Step] As Single), Numerics.Vector2())
            Dim plots As Numerics.Vector2()
            If allPlotsDictionary.ContainsKey((Me.Alpha, [step])) Then plots = allPlotsDictionary((Me.Alpha, [step])) Else plots = GetAllPlots([step]).ToArray
            For i = 0 To plots.Count - 2
                Dim y0 = plots(i).Y
                Dim y1 = plots(i + 1).Y
                Dim x0 = plots(i).X
                Dim x1 = plots(i + 1).X
                If y >= y0 And y <= y1 Then Return (x0 - x1) / (y0 - y1) * y + (y0 * x1 - y1 * x0) / (y0 - y1)
            Next
            Return Single.MinValue
        End Function
        Public Function GetYValue(x As Single, [step] As Single) As Double
            Static allPlotsDictionary As New Dictionary(Of (Alpha As Single, [Step] As Single), Numerics.Vector2())
            Dim plots As Numerics.Vector2()
            If allPlotsDictionary.ContainsKey((Me.Alpha, [step])) Then plots = allPlotsDictionary((Me.Alpha, [step])) Else plots = GetAllPlots([step]).ToArray
            For i = 0 To plots.Count - 2
                Dim y0 = plots(i).Y
                Dim y1 = plots(i + 1).Y
                Dim x0 = plots(i).X
                Dim x1 = plots(i + 1).X
                If x >= x0 And x <= x1 Then Return (y0 - y1) / (x0 - x1) * x + (y1 * x0 - y0 * x1) / (x0 - x1)
            Next
            Return Single.MinValue
        End Function
        Public Sub New(plots As (X As Single, Y As Single)(), Optional alpha As Single = CentripetalAlpha)
            Me.P = (From _p In plots Select New Numerics.Vector2(_p.X, _p.Y)).ToArray
            Me.Alpha = alpha
        End Sub
        Public Sub New(plots As Numerics.Vector2(), Optional alpha As Single = CentripetalAlpha)
            Me.P = plots
            Me.Alpha = alpha
        End Sub

        Public Shared Function GetTParameter(tl As Single, pl As Numerics.Vector2, pr As Numerics.Vector2, alpha As Single) As Single
            Return ((pr.X - pl.X) ^ 2 + (pr.Y + pl.Y) ^ 2) ^ (alpha / 2) + tl
        End Function
    End Class

    Public Class Interval
        Private Shared ReadOnly OperatorCodeCollection As New Dictionary(Of String, (InitializingParameterCount As Integer, IntervalToString As Func(Of String(), String), IsInInterval As Func(Of String(), Double, Boolean))) From {
            {"gt", (1, Function(p As String()) "＞" & p(0), Function(p As String(), d As Double) d > p(0))},
            {"lt", (1, Function(p As String()) "＜" & p(0), Function(p As String(), d As Double) d < p(0))},
            {"ge", (1, Function(p As String()) If(Environment.OSVersion.Version.Major = 10, "⩾", "≥") & p(0), Function(p As String(), d As Double) d >= p(0))}, '"⩾"
            {"le", (1, Function(p As String()) If(Environment.OSVersion.Version.Major = 10, "⩽", "≤") & p(0), Function(p As String(), d As Double) d <= p(0))}, '"⩽"
            {"in", (2, Function(p As String()) p(0) & "～" & p(1), Function(p As String(), d As Double) d >= p(0) And d <= p(1))},
            {"eq", (2, Function(p As String()) p(0), Function(p As String(), d As Double) d = p(0))},
            {"lc", (2, Function(p As String()) "[" & p(0) & ", +∞)", Function(p As String(), d As Double) d >= p(0))},
            {"lo", (2, Function(p As String()) "(" & p(0) & ", +∞)", Function(p As String(), d As Double) d > p(0))},
            {"rc", (2, Function(p As String()) "(-∞, " & p(0) & "]", Function(p As String(), d As Double) d <= p(0))},
            {"ro", (2, Function(p As String()) "(-∞, " & p(0) & ")", Function(p As String(), d As Double) d < p(0))},
            {"lcro", (2, Function(p As String()) "[" & p(0) & ", " & p(1) & ")", Function(p As String(), d As Double) d >= p(0) And d < p(1))},
            {"lorc", (2, Function(p As String()) "(" & p(0) & ", " & p(1) & "]", Function(p As String(), d As Double) d > p(0) And d <= p(1))},
            {"lcrc", (2, Function(p As String()) "[" & p(0) & ", " & p(1) & "]", Function(p As String(), d As Double) d >= p(0) And d <= p(1))},
            {"loro", (2, Function(p As String()) "(" & p(0) & ", " & p(1) & ")", Function(p As String(), d As Double) d > p(0) And d < p(1))},
            {"/", (0, Function(p As String()) "/", Function(p As String(), d As Double) True)}
        }
        Public ReadOnly Property OperatorCode As String
        Public ReadOnly Property Parameters As String()

        Public Sub New(intervalCode As String)
            Dim parts = intervalCode.Split(" ")
            If Not OperatorCodeCollection.ContainsKey(parts(0)) Then Throw New Exception
            Me.OperatorCode = parts(0)
            Me.Parameters = (From i In parts Where IsNumeric(i) Select i).ToArray
        End Sub

        Public Overrides Function ToString() As String
            Return OperatorCodeCollection(Me.OperatorCode).IntervalToString(Me.Parameters)
        End Function

        Public Function IsInInterval(number As Double) As Boolean
            Return OperatorCodeCollection(Me.OperatorCode).IsInInterval(Me.Parameters, number)
        End Function
    End Class

    Public Class ExcelReferenceInterop
        Private Sub New()
        End Sub
        Public Shared Function Rows(r As ExcelReference) As ExcelReference

        End Function

    End Class

    Public Class ConcatModelElement
        Public ReadOnly Property Value As String
        Public ReadOnly Property IsColumnSplitter As Boolean
        Public ReadOnly Property IsRowSplitter As Boolean
        Public Sub New(value As String, isRowSplitter As Boolean, isColumnSplitter As Boolean)
            Me.Value = value
            Me.IsRowSplitter = isRowSplitter
            If isRowSplitter Then Me.IsColumnSplitter = False Else Me.IsColumnSplitter = isColumnSplitter
        End Sub
        Public Overrides Function ToString() As String
            Return Me.Value
        End Function
        Public Shared Narrowing Operator CType(e As ConcatModelElement) As String
            Return e.ToString
        End Operator
        Public Shared Widening Operator CType(e As String) As ConcatModelElement
            Return New ConcatModelElement(e, False, False)
        End Operator
    End Class
End Namespace