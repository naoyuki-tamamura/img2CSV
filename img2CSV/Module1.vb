Imports System
Imports System.IO

Module Module1

    Sub Main(args As String())
        Dim SourceFilePath As String = Nothing
        If args.Length = 0 Then
            Console.WriteLine("CSVに変換するファイルを指定してください。")
            Environment.Exit(1)
        Else
            SourceFilePath = args(0)
        End If

        Dim BaseFileName = System.IO.Path.GetFileNameWithoutExtension(SourceFilePath)
        Dim DestFilePath As String
        If Path.IsPathRooted(SourceFilePath) = True Then
            DestFilePath = System.IO.Path.GetDirectoryName(SourceFilePath) & "\" & BaseFileName & ".csv"
        Else
            DestFilePath = System.IO.Directory.GetCurrentDirectory() & "\" & BaseFileName & ".csv"
        End If

        If System.IO.File.Exists(SourceFilePath) = False Then
            Console.WriteLine(SourceFilePath & "が見つかりません。")
            Environment.Exit(2)
        End If

        If System.IO.File.Exists(DestFilePath) Then
            Console.WriteLine("同じフォルダに '" + BaseFileName & ".csv" + "' が存在するため、処理を中止します。")
            Environment.Exit(2)
        End If

        'ファイルのサイズを取得
        Dim fi As New System.IO.FileInfo(SourceFilePath)
        Dim FileSize As Long = fi.Length

        Dim Buff As String
        ' キーボードから文字列の入力

        Dim HeaderSize As Long = 0

        Console.WriteLine("ヘッダサイズを半角整数で入力してください：")
        Buff = Console.ReadLine()
        If IsNumeric(Buff) Then
            HeaderSize = Integer.Parse(Buff)
        Else
            Console.WriteLine("整数で入力してください。")
            Environment.Exit(3)
        End If

        Console.WriteLine("横方向の画素数を半角整数で入力してください：")
        Buff = Console.ReadLine()

        Dim MatrixX As Long = 0
        If IsNumeric(Buff) Then
            MatrixX = Integer.Parse(Buff)
        Else
            Console.WriteLine("整数で入力してください。")
            Environment.Exit(3)
        End If

        Dim DataFormat As Long = 0
        Console.WriteLine("データフォーマットを選択してください 1:Byte 2:Integer 3：Float 4：UnsignedInteger ")
        Buff = Console.ReadLine()
        If IsNumeric(Buff) Then
            DataFormat = Integer.Parse(Buff)
            If DataFormat <= 0 And 5 <= DataFormat Then
                Console.WriteLine("1～4を入力してください")
                Environment.Exit(3)
            End If
        Else
            Console.WriteLine("整数で入力してください。")
            Environment.Exit(2)
        End If

        Dim BytesPerPixel As Long = 0
        Select Case DataFormat
            Case 1
                BytesPerPixel = 1
            Case 2
                BytesPerPixel = 2
            Case 3
                BytesPerPixel = 4
            Case 4
                BytesPerPixel = 2
        End Select

        If (FileSize - HeaderSize) Mod (MatrixX * BytesPerPixel) <> 0 Then
            Console.WriteLine("ヘッダサイズの指定、横方向の画素数、データフォーマットの指定のいずれかに誤りがあります。")
            Console.WriteLine("Enterキーを押すと終了します。")
            Console.ReadLine()
            Environment.Exit(3)
        End If

        Dim DataEndian As Long = 0
        If DataFormat > 1 Then
            Console.WriteLine("バイトオーダーを指定してください 1:Little Endian 2:Big Endian ")
            Buff = Console.ReadLine()
            If IsNumeric(Buff) Then
                DataEndian = Integer.Parse(Buff)
                If DataEndian <= 0 And 3 <= DataEndian Then
                    Console.WriteLine("1もしくは2を入力してください")
                    Environment.Exit(3)
                End If
            Else
                Console.WriteLine("整数で入力してください。")
                Environment.Exit(2)
            End If
        End If

        Dim MatrixY As Long = (FileSize - HeaderSize) / (MatrixX * BytesPerPixel)

        Dim arrPixelValueText As String()() = Nothing
        ReDim arrPixelValueText(MatrixY - 1)
        For CurrentY = 0 To MatrixY - 1
            ReDim arrPixelValueText(CurrentY)(MatrixX - 1)
        Next

        Using stream As Stream = File.OpenRead(SourceFilePath)
            ' streamから読み込むためのBinaryReaderを作成
            Using reader As New BinaryReader(stream)
                If HeaderSize > 0 Then
                    reader.ReadBytes(HeaderSize)
                End If
                For CurrentY = 0 To MatrixY - 1
                    For CurrentX = 0 To MatrixX - 1
                        Select Case DataFormat
                            Case 1
                                arrPixelValueText(CurrentY)(CurrentX) = CStr(reader.ReadByte)
                            Case 2
                                Dim TempArray(1) As Byte
                                TempArray(0) = reader.ReadByte
                                TempArray(1) = reader.ReadByte
                                If DataEndian = 2 Then
                                    Array.Reverse(TempArray)
                                End If
                                arrPixelValueText(CurrentY)(CurrentX) = CStr(BitConverter.ToInt16(TempArray, 0))
                            Case 3
                                Dim TempArray(3) As Byte
                                TempArray(0) = reader.ReadByte
                                TempArray(1) = reader.ReadByte
                                TempArray(2) = reader.ReadByte
                                TempArray(3) = reader.ReadByte
                                If DataEndian = 2 Then
                                    Array.Reverse(TempArray)
                                End If
                                arrPixelValueText(CurrentY)(CurrentX) = CStr(BitConverter.ToSingle(TempArray, 0))
                            Case 4
                                Dim TempArray(1) As Byte
                                TempArray(0) = reader.ReadByte
                                TempArray(1) = reader.ReadByte
                                If DataEndian = 2 Then
                                    Array.Reverse(TempArray)
                                End If
                                arrPixelValueText(CurrentY)(CurrentX) = CStr(BitConverter.ToUInt16(TempArray, 0))
                        End Select
                    Next
                Next
            End Using
        End Using

        WriteCsv(DestFilePath, arrPixelValueText)
        Console.WriteLine("完了しました。")
        Console.WriteLine("Enterキーを押すと終了します。")
        Console.ReadLine()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' CSVファイルの書込処理
    ''' </summary>
    ''' <param name="astrFileName">ファイル名</param>
    ''' <param name="aarrData">書込データ文字列の2次元配列</param>
    ''' <returns>True:結果OK, False:NG</returns>
    ''' <remarks>カラム名をファイルに出力したい場合は、書込データの先頭に設定すること</remarks>
    ''' -----------------------------------------------------------------------------
    Private Function WriteCsv(ByVal astrFileName As String, ByVal aarrData As String()()) As Boolean
        WriteCsv = False
        'ファイルStreamWriter
        Dim sw As System.IO.StreamWriter = Nothing

        Try
            'CSVファイル書込に使うEncoding
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("UTF-8")
            '書き込むファイルを開く
            sw = New System.IO.StreamWriter(astrFileName, False, enc)

            For Each arrLine() As String In aarrData
                Dim blnFirst As Boolean = True
                Dim strLIne As String = ""
                For Each str As String In arrLine
                    If blnFirst = False Then
                        '「,」(カンマ)の書込
                        sw.Write(",")
                    End If
                    blnFirst = False
                    '1カラムデータの書込
                    'str = """" & str & """"
                    sw.Write(str)
                Next
                '改行の書込
                sw.Write(vbCrLf)
            Next

            '正常終了
            Return True

        Catch ex As Exception
            'エラー
            MsgBox(ex.Message)
        Finally
            '閉じる
            If sw IsNot Nothing Then
                sw.Close()
            End If
        End Try
    End Function

End Module
