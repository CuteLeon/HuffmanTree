Public Class MainForm
    Dim RootNode As HuffmanTreeNode
    Dim ByteCounter(255) As ULong
    Dim HuffmanCode(255) As String
    Dim HuffmanNodeList As List(Of HuffmanTreeNode) = New List(Of HuffmanTreeNode)

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Index As Integer, ListSubcript As Byte
        Dim FileBytes() As Byte

        FileBytes = IO.File.ReadAllBytes(Application.StartupPath & "\HuffmanTreeTestFile.png")
        'FileBytes = System.Text.Encoding.UTF8.GetBytes("测试数据！！！ABC")

        Debug.Print("读取文件结束！文件大小： " & FileBytes.Count & " 字节")

        '使用并行计算统计 0000 0000~1111 1111 共256种字节出现的次数，但是并行统计错误较大，暂不使用
        'Parallel.[ForEach](Of Byte)(FileBytes, New ParallelOptions With {.MaxDegreeOfParallelism = 8},
        '    Sub(X As Byte)
        '        ByteCounter(X) += 1
        '    End Sub)

        For Each FileByte In FileBytes
            ByteCounter(FileByte) += 1
        Next
        Debug.Print("构建哈希表完毕！输出哈希表：")

        For Index = 0 To 255
            Debug.Print(Index.ToString("000") & " / " & GetBitArray(Index) & " ： " & ByteCounter(Index))
        Next
        Debug.Print("哈希表输出完毕！")

        HuffmanNodeList = New List(Of HuffmanTreeNode) With {.Capacity = 256}
        For Index = 0 To 255
            If ByteCounter(Index) > 0 Then
                Dim NewHuffmanNode As HuffmanTreeNode = New HuffmanTreeNode(ByteCounter(Index), Index)
                ListSubcript = GetInsertIndex(HuffmanNodeList, ByteCounter(Index))
                HuffmanNodeList.Insert(ListSubcript, NewHuffmanNode)
            End If
        Next
        Debug.Print("构建字节优先列表完毕！输出优先队列并构建二叉树：")

        PrintNodeList()

        RootNode = CreateHuffmanTree(HuffmanNodeList)
        Debug.Print("哈弗曼树生成完毕！")

        PrintHuffmanCode()

        '计算哈夫曼编码压缩率
        Dim HuffmanLength As Integer
        For Each FileByte In FileBytes
            HuffmanLength += HuffmanCode(FileByte).Length
        Next
        Debug.Print("压缩为原数据的：" & (HuffmanLength / (FileBytes.Length * 8)).ToString("###.##%"))

        Erase ByteCounter
        HuffmanNodeList.Clear()
        GC.Collect()
    End Sub

    ''' <summary>
    ''' 输出字节的二进制形式
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    Private Function GetBitArray(Value As Byte) As String
        Dim BitArray As BitArray = New BitArray(New Byte() {Value})
        Dim BitString As String = vbNullString
        For Index As Byte = 0 To BitArray.Length - 1
            BitString = IIf(BitArray.Get(Index), "1", "0") & BitString
        Next
        Return BitString
    End Function

    ''' <summary>
    ''' 定位字节排序插入的位置（从小到大的排列顺序）
    ''' </summary>
    ''' <returns></returns>
    Private Function GetInsertIndex(ByteList As List(Of HuffmanTreeNode), CounterValue As ULong) As Byte
        Dim Left As Short = 0, Mid As Short,
               Right As Short = ByteList.Count
        Do While Left < Right
            [Mid] = ((Left + Right) \ 2)
            If CounterValue > ByteList(Mid).Counter Then
                Left = Mid + 1
            ElseIf CounterValue < ByteList(Mid).Counter Then
                Right = Mid
            Else
                Return Mid
            End If
        Loop
        Return Right
    End Function

    ''' <summary>
    ''' 从哈弗曼树节点列表生成哈夫曼树
    ''' </summary>
    ''' <param name="NodeList">哈弗曼树节点列表（必须）</param>
    Private Function CreateHuffmanTree(NodeList As List(Of HuffmanTreeNode)) As HuffmanTreeNode
        Do Until NodeList.Count = 1
            Dim ParentNode As HuffmanTreeNode = New HuffmanTreeNode(NodeList(0).Counter + NodeList(1).Counter, 0)
            CreateHuffmanCode(NodeList(0), "0")
            ParentNode.LeftChild = NodeList(0)
            CreateHuffmanCode(NodeList(1), "1")
            ParentNode.RightChild = NodeList(1)
            NodeList.RemoveRange(0, 2)
            NodeList.Insert(GetInsertIndex(NodeList, ParentNode.Counter), ParentNode)

            'PrintNodeList()
        Loop
        Return NodeList.First
    End Function

    ''' <summary>
    ''' 输出节点队列
    ''' </summary>
    Private Sub PrintNodeList()
        For Index = 0 To HuffmanNodeList.Count - 1
            Debug.Print(HuffmanNodeList(Index).ByteCode & " : " & HuffmanNodeList(Index).Counter)
        Next
        Debug.Print("优先队列输出完毕！共 {0} 个节点！", HuffmanNodeList.Count)
    End Sub

    ''' <summary>
    ''' 输出哈夫曼编码
    ''' </summary>
    Private Sub PrintHuffmanCode()
        For Index = 0 To 255
            Debug.Print(Index & " : " & HuffmanCode(Index))
        Next
        Debug.Print("输出哈夫曼编码完毕！")
    End Sub

    ''' <summary>
    ''' 建树时顺便计算哈弗曼编码
    ''' </summary>
    ''' <param name="ChildNode"></param>
    ''' <param name="CodeChar"></param>
    Private Sub CreateHuffmanCode(ChildNode As HuffmanTreeNode, CodeChar As String)
        If IsNothing(ChildNode.LeftChild) And IsNothing(ChildNode.RightChild) Then
            HuffmanCode(ChildNode.ByteCode) = CodeChar
        Else
            CreateHuffmanCode(ChildNode.LeftChild, CodeChar & "0")
            CreateHuffmanCode(ChildNode.RightChild, CodeChar & "1")
        End If
    End Sub

End Class
