Public Class HuffmanTreeNode
    Public ByteCode As Byte
    Public Counter As ULong
    Public Parent As HuffmanTreeNode
    Public LeftChild As HuffmanTreeNode
    Public RightChild As HuffmanTreeNode

    Public Sub New(NodeCounter As Integer, NodeByteCode As Integer)
        Counter = NodeCounter
        ByteCode = NodeByteCode
    End Sub

End Class