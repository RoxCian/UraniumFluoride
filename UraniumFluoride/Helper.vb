#Region "RandNoRepeat helper class"

Namespace Helper

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
                If Me.Count = 0 Then Return 0 Else Return Me.Item(0).Count
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
                        Me._IsValued(i) = True
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
                Me._IsValued(AbsolutePointerIndex) = True
            End Set
        End Property
        Public Overloads ReadOnly Property Count As Integer
            Get
                Dim r As Integer
                For i = 0 To Me.Capacity - 1
                    If Me._IsValued(i) Then r += 1
                Next
                Return r
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
        Dim _AbsolutePointerIndex As Integer = 0
        Dim _IsValued As New List(Of Boolean)(Me.Capacity)

        Private Sub New()
            MyBase.New()
        End Sub
        Public Sub New(capacity As Integer)
            MyBase.New(capacity)
            For i = 0 To capacity - 1
                MyBase.Add(Nothing)
                Me._IsValued.Add(False)
            Next
            Me.Capacity = capacity
        End Sub
        Public Sub New(capacity As Integer, value As T)
            MyBase.New(capacity)
            For i = 0 To capacity - 1
                MyBase.Add(value)
                Me._IsValued.Add(False)
            Next
        End Sub
        Private Sub New(collection As IEnumerable(Of T))
            MyBase.New(collection)
        End Sub

        Public Function GetAbsoluteIndex(relativeIndex As Integer) As Integer
            Dim absoluteIndex As Integer = Me.AbsolutePointerIndex + relativeIndex
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
            If Me._IsValued(Me.AbsolutePointerIndex) Then Me.AbsolutePointerIndex += 1
        End Sub
        Public Sub MovePrevious()
            Me._AbsolutePointerIndex -= 1
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
#End Region
End Namespace