Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.classes
    Public Class PrintedPoligonClass
        Private _objID As ObjectId
        Private _dblBottonLeft As Point2d
        Private _dblTopRight As Point2d
        Private _dblArea As Double
        Private _dblXSize As Double
        Private _dblYSize As Double
        Private _acNormal As Vector3d


        Public Property ID As ObjectId
            Get
                Return _objID
            End Get
            Set(value As ObjectId)
                Me._objID = value
            End Set
        End Property

        Public Property BottonLeft As Point2d
            Get
                Return _dblBottonLeft
            End Get
            Set(value As Point2d)
                Me._dblBottonLeft = value
            End Set
        End Property

        Public Property TopRight As Point2d
            Get
                Return _dblTopRight
            End Get
            Set(value As Point2d)
                Me._dblTopRight = value
            End Set
        End Property

        Public Property Area As Double
            Get
                Return _dblArea
            End Get
            Set(value As Double)
                Me._dblArea = value
            End Set
        End Property

        Public Property XSize As Double
            Get
                Return _dblXSize
            End Get
            Set(value As Double)
                Me._dblXSize = value
            End Set
        End Property

        Public Property YSize As Double
            Get
                Return _dblYSize
            End Get
            Set(value As Double)
                Me._dblYSize = value
            End Set
        End Property

        Public Property Normal As Vector3d
            Get
                Return _acNormal
            End Get
            Set(value As Vector3d)
                Me._acNormal = value
            End Set
        End Property

    End Class

End Namespace