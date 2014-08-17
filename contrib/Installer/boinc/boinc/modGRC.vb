﻿Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.IO
Imports System.Reflection
Imports System.Net
Imports System.Text
Imports System.Security.Cryptography

Module modGRC

    Public mclsUtilization As Utilization

    Public mfrmMining As frmMining

    Public mfrmProjects As frmProjects
    Public mfrmSql As frmSQL
    Public mfrmGridcoinMiner As frmGridcoinMiner
    Public mfrmLeaderboard As frmLeaderboard
    Public MerkleRoot As String = "0xda43abf15a2fcd57ceae9ea0b4e0d872981e2c0b72244466650ce6010a14efb8"
    Public merkleroot2 As String = "0xda43abf15abcdefghjihjklmnopq872981e2c0b72244466650ce6010a14efb8"

    Structure xGridcoinMiningStructure
        Public Shared device As String = "0"
        Public Shared gpu_thread_concurrency As String = "8192"
        Public Shared worksize As String = "256"
        Public Shared intensity As String = "13"
        Public Shared lookup_gap As String = "2"
    End Structure

    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()
        Dim sha1 As New SHA1CryptoServiceProvider
        ' Hash the key. 
        Dim keyBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)
        ' Truncate or pad the hash. 
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

    Private Function CreateAESEncryptor(ByVal sSalt As String) As ICryptoTransform

        Try
            Dim encryptor As ICryptoTransform = Aes.Create.CreateEncryptor(TruncateHash(MerkleRoot + Right(sSalt, 4), Aes.Create.KeySize \ 8), TruncateHash("", Aes.Create.BlockSize \ 8))

            'Aes.Create.Key = TruncateHash(MerkleRoot + Right(sSalt, 4), Aes.Create.KeySize \ 8)
            'Aes.Create.IV = TruncateHash("", Aes.Create.BlockSize \ 8)
            Return encryptor

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Private Function CreateAESDecryptor(ByVal sSalt As String) As ICryptoTransform

        Try
            Dim decryptor As ICryptoTransform = Aes.Create.CreateDecryptor(TruncateHash(MerkleRoot + Right(sSalt, 4), Aes.Create.KeySize \ 8), TruncateHash("", Aes.Create.BlockSize \ 8))

            'Aes.Create.Key = TruncateHash(MerkleRoot + Right(sSalt, 4), Aes.Create.KeySize \ 8)
            'Aes.Create.IV = TruncateHash("", Aes.Create.BlockSize \ 8)
            Return decryptor

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function StringToByte(sData As String)
        Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(sData)
        Return keyBytes
    End Function
    Public Function xReturnMiningValue(iDeviceId As Integer, sKey As String, bUsePrefix As Boolean) As String
        Dim g As New xGridcoinMiningStructure
        Dim sOut As String = ""
        Dim sLookupKey As String = LCase("dev" + Trim(iDeviceId) + "_" + sKey)
        If Not bUsePrefix Then sLookupKey = LCase(sKey)
        sOut = KeyValue(sLookupKey)
        Return sOut
    End Function
    Public Function GetGRCAppDir() As String
        Try
            Dim fi As New System.IO.FileInfo(Assembly.GetExecutingAssembly().Location)
            Return fi.DirectoryName
        Catch ex As Exception
        End Try
    End Function

    Public Function Base64File(sFileName As String)
        Dim sFilePath As String = GetGRCAppDir() + "\" + sFileName
        Dim b() As Byte
        b = System.IO.File.ReadAllBytes(sFilePath)
        Dim sBase64 As String = System.Convert.ToBase64String(b, 0, b.Length)
        b = System.Text.Encoding.ASCII.GetBytes(sBase64)
        System.IO.File.WriteAllBytes(sFilePath, b)
    End Function
    Public Function UnBase64File(sSourceFileName As String, sTargetFileName As String)
        Dim b() As Byte
        b = System.IO.File.ReadAllBytes(sSourceFileName)
        Dim value As String = System.Text.ASCIIEncoding.ASCII.GetString(b)
        b = System.Convert.FromBase64String(value)
        System.IO.File.WriteAllBytes(sTargetFileName, b)

    End Function

    Public Function GetMd5String(ByVal sData As String) As String
        Try
            Dim objMD5 As New System.Security.Cryptography.MD5CryptoServiceProvider
            Dim arrData() As Byte
            Dim arrHash() As Byte

            ' first convert the string to bytes (using UTF8 encoding for unicode characters)
            arrData = System.Text.Encoding.UTF8.GetBytes(sData)

            ' hash contents of this byte array
            arrHash = objMD5.ComputeHash(arrData)

            ' thanks objects
            objMD5 = Nothing
            Dim sOut As String = ByteArrayToWierdHexString(arrHash)

            ' return formatted hash
            Return sOut

        Catch ex As Exception
            Return "MD5Error"
        End Try
    End Function

    ' utility function to convert a byte array into a hex string
    Private Function ByteArrayToWierdHexString(ByVal arrInput() As Byte) As String
        Dim strOutput As New System.Text.StringBuilder(arrInput.Length)
        For i As Integer = 0 To arrInput.Length - 1
            strOutput.Append(arrInput(i).ToString("X2"))
        Next
        Return strOutput.ToString().ToLower
    End Function



    Public Function ByteToString(b() As Byte)
        Dim sReq As String
        sReq = System.Text.Encoding.UTF8.GetString(b)
        Return sReq
    End Function

    Public Function RestartWallet1(sParams As String)
        Dim p As Process = New Process()
        Dim pi As ProcessStartInfo = New ProcessStartInfo()
        pi.WorkingDirectory = GetGRCAppDir()
        pi.UseShellExecute = True
        Log("Restarting wallet with params " + sParams)

        pi.Arguments = sParams
        pi.FileName = Trim("GRCRestarter.exe")
        p.StartInfo = pi
        p.Start()
    End Function
    Public Function ConfigPath() As String
        Dim sFolder As String
        sFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\Gridcoinstake"
        Dim sPath As String
        sPath = sFolder + "\gridcoinstake.conf"
        Return sPath
    End Function
    Public Function GetGridPath(ByVal sType As String) As String
        Dim sTemp As String
        sTemp = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\Gridcoinstake\" + sType
        If System.IO.Directory.Exists(sTemp) = False Then
            Try
                System.IO.Directory.CreateDirectory(sTemp)
            Catch ex As Exception

            End Try
        End If
        Return sTemp
    End Function
    Public Function GetGridFolder() As String
        Dim sTemp As String
        sTemp = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\Gridcoinstake\"
        Return sTemp
    End Function
    Public Function KeyValue(ByVal sKey As String) As String
        Try
            Dim sPath As String = ConfigPath()
            Dim sr As New StreamReader(sPath)
            Dim sRow As String
            Dim vRow() As String
            Do While sr.EndOfStream = False
                sRow = sr.ReadLine
                vRow = Split(sRow, "=")
                If LCase(vRow(0)) = LCase(sKey) Then
                    sr.Close()
                    Return Trim(vRow(1) & "")
                End If
            Loop
            sr.Close()
            Return ""

        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Function UpdateKey(ByVal sKey As String, ByVal sValue As String)
        Try
            Dim sInPath As String = ConfigPath()
            Dim sOutPath As String = ConfigPath() + ".bak"

            Dim bFound As Boolean

            Dim sr As New StreamReader(sInPath)
            Dim sw As New StreamWriter(sOutPath, False)

            Dim sRow As String
            Dim vRow() As String
            Do While sr.EndOfStream = False
                sRow = sr.ReadLine
                vRow = Split(sRow, "=")
                If UBound(vRow) > 0 Then
                    If LCase(vRow(0)) = LCase(sKey) Then
                        sw.WriteLine(sKey + "=" + sValue)
                        bFound = True
                    Else
                        sw.WriteLine(sRow)
                    End If
                End If

            Loop
            If bFound = False Then
                sw.WriteLine(sKey + "=" + sValue)
            End If
            sr.Close()
            sw.Close()
            Kill(sInPath)
            FileCopy(sOutPath, sInPath)
            Kill(sOutPath)
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Function cBOO(data As Object) As Boolean
        Dim bOut As Boolean
        Try
            Dim sBoo As String = data.ToString
            bOut = CBool(sBoo)
        Catch ex As Exception
            Return False
        End Try
        Return bOut
    End Function
    Public Sub KillGuiMiner()
        Dim tKill As New System.Threading.Thread(AddressOf ThreadKillGuiMiner)
        tKill.Start()
    End Sub
    Public Sub ThreadKillGuiMiner()
        Log("Closing all miners (guiminer,cgminer,conhost,reaper).")

        For x = 1 To 6
            Try
                KillProcess("guiminer*")
                KillProcess("cgminer*")
                KillProcess("conhost*")
                KillProcess("reaper*")
                mCGMinerHwnd(x - 1) = 0
            Catch ex As Exception
            End Try
        Next x

    End Sub

    Public Sub Snapshot()
        Dim sDataDir As String = GRCDataDir()
        Dim sBlocks = sDataDir + "blocks"
        Dim sChain = sDataDir + "chainstate"
        '   Dim sDatabase = sDataDir + "database"
        Dim sSnapshotBlocks = sDataDir + "snapshot\blocks"
        Dim sSnapshotChain = sDataDir + "snapshot\chainstate"
        Dim sSnapshotDatabase = sDataDir + "snapshot\database"
        Dim sSnapshotDir = sDataDir + "snapshot"

        Try
            Dim dDestination2 As New System.IO.DirectoryInfo(sSnapshotDir)
            dDestination2.Delete()

        Catch ex As Exception

        End Try

        DirectorySnapshot(sBlocks, sSnapshotBlocks, True)
        DirectorySnapshot(sChain, sSnapshotChain, True)

        'DirectorySnapshot(sDatabase, sSnapshotDatabase, True)
    End Sub
    Public Sub RestoreFromSnapshot()
        Dim sDataDir As String = GRCDataDir()
        Dim sBlocks = sDataDir + "blocks"
        Dim sChain = sDataDir + "chainstate"
        Dim sSnapshotBlocks = sDataDir + "snapshot\blocks"
        Dim sSnapshotChain = sDataDir + "snapshot\chainstate"
        DirectorySnapshot(sSnapshotBlocks, sBlocks, True)
        DirectorySnapshot(sSnapshotChain, sChain, True)
    End Sub


    Private Sub DirectorySnapshot( _
        ByVal sourceDirName As String, _
        ByVal destDirName As String, _
        ByVal copySubDirs As Boolean)

        ' Get the subdirectories for the specified directory. 
        Dim dir As DirectoryInfo = New DirectoryInfo(sourceDirName)
        Dim dirs As DirectoryInfo() = dir.GetDirectories()

        If Not dir.Exists Then
            'Throw New DirectoryNotFoundException( _               "Source directory does not exist or could not be found: " _        + sourceDirName)

        End If
        'Remove the destination directory
        Dim dDestination As New System.IO.DirectoryInfo(destDirName)

        Try
            RemoveBlocksDir(dDestination)

        Catch ex As Exception

        End Try


        ' If the destination directory doesn't exist, create it. 
        If Not Directory.Exists(destDirName) Then
            Directory.CreateDirectory(destDirName)
        End If
        ' Get the files in the directory and copy them to the new location. 
        Dim files As FileInfo() = dir.GetFiles()
        For Each file In files
            Dim temppath As String = Path.Combine(destDirName, file.Name)
            Try
                file.CopyTo(temppath, False)

            Catch ex As Exception
                Log("Error while copying " + file.Name + " to " + temppath)

            End Try
        Next file

        ' If copying subdirectories, copy them and their contents to new location. 
        If copySubDirs Then
            For Each subdir In dirs
                Dim temppath As String = Path.Combine(destDirName, subdir.Name)
                DirectorySnapshot(subdir.FullName, temppath, copySubDirs)
            Next subdir
        End If
    End Sub


    Private Sub RemoveBlocksDir(d As System.IO.DirectoryInfo)
        Try
            d.Delete(True)
        Catch ex As Exception
        End Try
    End Sub


    Public Function GRCDataDir() As String
        Dim sFolder As String
        sFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\Gridcoinstake\"
        Return sFolder
    End Function
    Public Function GetBlockExplorerBlockNumber() As Long
        Try
            Dim w As New MyWebClient
            Dim sWap As String
            Try
                sWap = w.DownloadString("http://explorer.gridcoin.us")
                'Blocks:</TD><TD> 35671</TD></TR>
            Catch ex As Exception
                Return -3
            End Try
            Dim iStart As Integer
            iStart = InStr(1, sWap, "Blocks:</TD>") + 12
            Dim iEnd As Integer
            iEnd = InStr(iStart, sWap, "</TD>")
            Dim sBlocks As String
            sBlocks = Mid(sWap, iStart, iEnd - iStart)
            sBlocks = Replace(sBlocks, "</TD>", "")
            sBlocks = Replace(sBlocks, "</TR>", "")
            sBlocks = Replace(sBlocks, "<TD>", "")
            sBlocks = Trim(sBlocks)
            Dim lBlock As Long
            lBlock = Val(sBlocks)
            Return lBlock
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Public Sub Log(sData As String)
        Try
            Dim sPath As String
            sPath = GetGridFolder() + "debug2.log"
            Dim sw As New System.IO.StreamWriter(sPath, True)
            sw.WriteLine(Trim(Now) + ", " + sData)
            sw.Close()
        Catch ex As Exception
        End Try

    End Sub



    Public Function LoadNodes(data As String)
        Exit Function '
        'Fix:1/5/2014 8:01:18 AM, getnewid:Argument 'Expression' cannot be converted to type 'DBNull'.

        Dim mData As New Sql

        Dim vData() As String
        vData = Split(data, vbCrLf)
        Dim sRow As String
        Dim vRow() As String
        Dim sNode As String
        Dim sVersion As String
        Dim sql As String
        Dim lNewId As Long

        Try

            For x = 0 To UBound(vData)
                sRow = vData(x)
                vRow = Split(sRow, "|")
                If UBound(vRow) >= 1 Then
                    sNode = vRow(0)
                    sVersion = vRow(1)
                    lNewId = GetNewId("peers", mData)

                    sql = "Insert into peers values ('" + Trim(lNewId) + "','" + Trim(sNode) + "','" + Trim(sVersion) + "',date('now'))"
                    Log(sql)

                    mData.Exec(sql)
                End If

            Next x

            Exit Function

        Catch ex As Exception
            Log("LoadNodes:" + ex.Message)

        End Try

    End Function
    Public Function GetNewId(sTable As String, mData As Sql) As Long
        Try
            Dim sql As String
            sql = "Select max(id) as maxid from " + sTable
            Dim vID As Long
            vID = Val(mData.QueryFirstRow(sql, "maxid")) + 1
            Return vID
        Catch ex As Exception
            Log("getnewid:" + ex.Message)
        End Try
    End Function


    Public Function RetrieveSiteSecurityInformation(sURL As String) As String
        Dim u As New Uri(sURL)
        Dim sp As ServicePoint = ServicePointManager.FindServicePoint(u)
        Dim groupName As String = Guid.NewGuid().ToString()
        Dim req As HttpWebRequest = TryCast(HttpWebRequest.Create(u), HttpWebRequest)
        req.ConnectionGroupName = groupName
        Try

            Using resp As WebResponse = req.GetResponse()
            End Using
            sp.CloseConnectionGroup(groupName)
            Dim key As Byte() = sp.Certificate.GetPublicKey()
            Dim sOut As String
            sOut = ByteArrayToHexString(key)
            Return sOut
        Catch ex As Exception
            'Usually due to either HTTP, 501, Not Implemented...etc.
            Return ""
        End Try

    End Function

    Public Function ByteArrayToHexString(ByVal ba As Byte()) As String
        Dim hex As StringBuilder
        hex = New StringBuilder(ba.Length * 2)
        For Each b As Byte In ba
            hex.AppendFormat("{0:x2}", b)
        Next
        Return hex.ToString()
    End Function

    Public Function GetBoincDataFolder() As String
        Dim sAppDir As String
        sAppDir = KeyValue("boincdatafolder")
        If Len(sAppDir) > 0 Then Return sAppDir
        Dim bigtime3f7o6l0daedrf4597acff2affbb5ed209f439aFroBearden0edd44ae1167a1e9be6eeb5cc2acd9c9 As String
        bigtime3f7o6l0daedrf4597acff2affbb5ed209f439aFroBearden0edd44ae1167a1e9be6eeb5cc2acd9c9 = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
        bigtime3f7o6l0daedrf4597acff2affbb5ed209f439aFroBearden0edd44ae1167a1e9be6eeb5cc2acd9c9 = bigtime3f7o6l0daedrf4597acff2affbb5ed209f439aFroBearden0edd44ae1167a1e9be6eeb5cc2acd9c9 + "\Boinc\"
        If Not System.IO.Directory.Exists(bigtime3f7o6l0daedrf4597acff2affbb5ed209f439aFroBearden0edd44ae1167a1e9be6eeb5cc2acd9c9) Then
            bigtime3f7o6l0daedrf4597acff2affbb5ed209f439aFroBearden0edd44ae1167a1e9be6eeb5cc2acd9c9 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) + mclsUtilization.Des3Decrypt("sEl7B/roaQaNGPo+ckyQBA==")

        End If
        Return bigtime3f7o6l0daedrf4597acff2affbb5ed209f439aFroBearden0edd44ae1167a1e9be6eeb5cc2acd9c9
    End Function

    Public Function ValidatePoolURL(sURL As String) As Boolean
        Dim Pools(100) As String
        Pools(1) = "gridcoin.us,30 82 01 0a 02 82 01 01 00 e1 91 3f 65 da 2b cc de 81 10 be 21 bd 8a 22 00 c5 8d 5f d6 72 5d 1c 3c e4 0b 3a 03 c8 07 c1 e1 69 54 22 d3 ff 9e d7 55 55 c2 2e 62 bd 5c bc f5 3f 93 3d f1 2c 39 0b 66 04 a8 50 7e f5 19 ca 97 a5 99 02 0b 11 39 37 5e df a2 74 14 f1 ed be eb af 4b 53 c2 cc a9 ea 5f c0 0a cb 92 cf 7f 21 fc 96 4f 79 47 e9 15 97 58 65 ef 10 a3 3e 46 6a 1d 5b 34 ea ff 6d c6 10 08 b8 60 dd 40 d5 b3 43 73 96 70 9f ce f1 2c 3b 8e 09 e0 14 97 9e b3 c6 6c a2 d9 81 4d d4 71 f1 46 ae ec b9 cf 0b 59 bd 7a 85 88 48 0f aa fa 6e f5 1a 75 18 f0 c9 94 79 6c 8b 11 86 de 3f ab 76 62 77 99 5a c4 fb 10 79 35 3d 61 33 15 ed a8 0c ce 45 cd 3e fc 64 62 72 07 a2 05 b4 df 3c f8 97 7c f9 20 43 b6 93 c2 2a 67 b7 9c 64 36 2f 9f 2d c3 d1 82 1a 9c 85 bb 3f d6 b7 07 aa 23 a3 a9 6a 49 18 f1 46 b5 b3 11 b6 61 02 03 01 00 01"

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Verify SSL Certificate'''''''''''''''''''''''''''
        Dim sHexSSLPublicKey As String = RetrieveSiteSecurityInformation(sURL)
        If Len(sHexSSLPublicKey) = 0 Then Return False

        For x As Integer = 0 To 100
            If Len(Pools(x)) > 0 Then
                Dim vPools() As String = Split(Pools(x), ",")
                If UBound(vPools) = 1 Then
                    Dim sPoolPublicKey As String = vPools(1)
                    If Len(sPoolPublicKey) > 0 Then
                        sPoolPublicKey = Replace(sPoolPublicKey, " ", "")
                        If LCase(sPoolPublicKey) = LCase(sHexSSLPublicKey) Then
                            Return True
                        End If
                    End If
                End If
            End If
        Next
        Return False

    End Function
    Public Function CalcMd5(sMd5 As String) As String

        Try
            Dim md5 As Object
            md5 = System.Security.Cryptography.MD5.Create()
            Dim b() As Byte
            b = StringToByte(sMd5)
            md5 = md5.ComputeHash(b)
            Dim sOut As String
            sOut = ByteArrayToHexString(md5)
            Return sOut
        Catch ex As Exception
            Return "MD5Error"
        End Try
    End Function


    Public Function ExtractFilename(ByVal sStartElement As String, ByVal sEndElement As String, ByVal sData As String, ByVal minOutLength As Integer) As String
        Try
            Dim sDataBackup As String
            sDataBackup = LCase(sData)
            Dim iStart As Integer
            Dim iEnd As Long
            Dim sOut As String
            iStart = InStr(1, sDataBackup, sStartElement) + Len(sStartElement) + 1
            iEnd = InStr(iStart + minOutLength, sDataBackup, sEndElement)
            sOut = Mid(sData, iStart, iEnd - iStart)
            sOut = Replace(sOut, ",", "")
            sOut = Replace(sOut, "br/>", "")
            sOut = Replace(sOut, "</a>", "")
            Dim iPrefix As Long
            iPrefix = InStr(1, sOut, ">")
            Dim sPrefix As String
            sPrefix = Mid(sOut, 1, iPrefix)
            sOut = Replace(sOut, sPrefix, "")
            Dim sExt As String
            sExt = LCase(Mid(sOut, Len(sOut) - 2, 3))
            sOut = LCase(sOut)
            If sExt = "pdf" Or LCase(sOut).Contains("to parent directory") Or sExt = "msi" Or sExt = "pdb" Or sExt = "xml" Or LCase(sOut).Contains("vshost") Or sExt = "txt" Or sOut = "gridcoin" Or sOut = "gridcoin_ro" Or sOut = "older" Or sExt = "cpp" Or sOut = "web.config" Then sOut = ""
            If sOut = "gridcoin.zip" Then sOut = ""
            If sOut = "gridcoinrdtestharness.exe.exe" Or sOut = "gridcoinrdtestharness.exe" Then sOut = ""
            If sOut = "cgminer_base64.zip" Then sOut = ""
            If sOut = "signed" Then sOut = ""

            Return Trim(sOut)
        Catch ex As Exception
            Dim message As String = ex.Message


        End Try
    End Function

    Public Function ParseDate(sDate As String)
        'parses microsofts IIS date to a date, globally
        Dim vDate() As String
        vDate = Split(sDate, " ")
        If UBound(vDate) > 0 Then
            Dim sEle1 As String = vDate(0)
            Dim vEle() As String
            vEle = Split(sEle1, "/")
            If UBound(vEle) > 1 Then
                Dim dt1 As Date
                dt1 = DateSerial(vEle(2), vEle(0), vEle(1))
                Return dt1

            End If
        End If
        Return CDate("1-1-2031")

    End Function

    Public Function NeedsUpgrade() As Boolean
        Try

            Dim sMsg As String
            Dim sURL As String = "http://download.gridcoin.us/download/downloadstake/"
            Dim w As New MyWebClient
            Dim sFiles As String
            sFiles = w.DownloadString(sURL)
            Dim vFiles() As String = Split(sFiles, "<br>")
            If UBound(vFiles) < 10 Then
                Return False
            End If

            sMsg = ""
            For iRow As Integer = 0 To UBound(vFiles)
                Dim sRow As String = vFiles(iRow)
                Dim sFile As String = ExtractFilename("<a", "</a>", sRow, 5)
                If Len(sFile) > 1 Then
                    If sFile = "boinc.dll" Then
                        Dim sDT As String
                        sDT = Mid(sRow, 1, 20)
                        sDT = Trim(sDT)

                        Dim dDt As DateTime
                        dDt = ParseDate(Trim(sDT))
                        dDt = TimeZoneInfo.ConvertTime(dDt, System.TimeZoneInfo.Utc)
                        'Hosting server is PST, so subtract Utc - 7 to achieve PST:
                        dDt = DateAdd(DateInterval.Hour, -12, dDt)
                        'local file time
                        Dim sLocalPath As String = GetGRCAppDir()
                        Dim sLocalFile As String = sFile
                        If LCase(sLocalFile) = "grcrestarter.exe" Then sLocalFile = "grcrestarter_copy.exe"
                        Dim sLocalPathFile As String = sLocalPath + "\" + sLocalFile
                        Dim dtLocal As DateTime
                        Try
                            dtLocal = System.IO.File.GetLastWriteTime(sLocalPathFile)
                            dtLocal = TimeZoneInfo.ConvertTime(dtLocal, System.TimeZoneInfo.Utc)
                            Log("Gridcoin.us boinc.dll timestamp (UTC) : " + Trim(dDt) + ", VS : Local boinc.dll timestamp (UTC) : " + Trim(dtLocal))
                            If dDt < dtLocal Then
                                Log("Not upgrading.")
                            End If

                        Catch ex As Exception
                            Return False
                        End Try
                        If dDt > dtLocal Then
                            Log("Client needs upgrade.")

                            Return True
                        End If

                    End If
                End If
            Next iRow
        Catch ex As Exception
            Return False

        End Try
    End Function


    Public Function AES512EncryptData(ByVal plaintext As String) As String

        Try

            ' Convert the plaintext string to a byte array. 
            Dim encryptor As ICryptoTransform = CreateAESEncryptor("salt")
            
            'Call MerkleAES(MerkleRoot)
            Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)
            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream. 
            Dim encStream As New CryptoStream(ms, encryptor, System.Security.Cryptography.CryptoStreamMode.Write)
            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()
            Try
                Return Convert.ToBase64String(ms.ToArray)

            Catch ex As Exception

            End Try


        Catch ex As Exception

        End Try
    End Function

    Public Function AES512DecryptData(ByVal encryptedtext As String) As String
        Try

            ' Convert the encrypted text string to a byte array. 
            '            MerkleAES(MerkleRoot)
            Dim decryptor As ICryptoTransform = CreateAESDecryptor("salt")
            Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)
            Dim ms As New System.IO.MemoryStream
            Dim decStream As New CryptoStream(ms, decryptor, System.Security.Cryptography.CryptoStreamMode.Write)
            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()
            ' Convert the plaintext stream to a string. 
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function



End Module

Public Class MyWebClient
    Inherits System.Net.WebClient
    Protected Overrides Function GetWebRequest(ByVal uri As Uri) As System.Net.WebRequest
        Dim w As System.Net.WebRequest = MyBase.GetWebRequest(uri)
        w.Timeout = 7000
        Return w
    End Function
End Class


