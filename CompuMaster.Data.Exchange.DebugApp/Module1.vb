Option Explicit On
Option Strict On

Imports CompuMaster.Data.MsExchange
Imports CompuMaster.Data.MsExchange.Exchange2007SP1OrHigher
'Imports Microsoft.Exchange.WebServices.Data

Module Module1

    Sub Main()
        Try
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AppStart")
            Dim he As New HlsMsExchangeDataAccess("srv20")
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " BeforeQuery")
            Dim t As DataTable = he.MsExchangeActivities(New Date(2016, 03, 03), New Date(2016, 03, 30, 23, 59, 59))
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AfterQuery")
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " TableOutput (Y/N)?")
            If Console.ReadKey().KeyChar.ToString.ToLowerInvariant = "y" Then
                Console.WriteLine()
                Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " BeforeOutput")
                Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(t))
            End If
            Console.WriteLine()
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " RowCount=" & t.Rows.Count)
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AppEnd")
        Catch ex As Exception
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AppError")
            Console.WriteLine(ex.ToString)
        End Try

        'TestExchange2007()
    End Sub



    Sub TestExchange2007()
        Try
            'Dim e2007 As New CompuMaster.Data.Exchange2007SP1OrHigher(CompuMaster.Data.Exchange2007SP1OrHigher.ExchangeVersion.Exchange2007_SP1,"", "jwezel@compumaster.de")
            Dim e2007 As New CompuMaster.Data.MsExchange.Exchange2007SP1OrHigher(CompuMaster.Data.MsExchange.Exchange2007SP1OrHigher.ExchangeVersion.Exchange2010_SP1, "srv20")

            Dim folderRoot As CompuMaster.Data.MsExchange.FolderPathRepresentation = e2007.LookupFolder(WellKnownFolderName.Root)
            Dim dirRoot As Directory = folderRoot.Directory.SelectSubFolder("AllItems", False, e2007.DirectorySeparatorChar)
            Dim dirInbox As Directory = dirRoot.InitialRootDirectory.SelectSubFolder("Oberste Ebene des Informationsspeichers\Inbox", False, e2007.DirectorySeparatorChar)

            'ShowItems(dirRoot, e2007)

            Console.WriteLine(dirRoot.DisplayPath)
            ForEachSubDirectory(dirRoot.InitialRootDirectory, e2007)

            'Dim folderRoot As CompuMaster.Data.MsExchange.FolderPathRepresentation = e2007.LookupFolder(WellKnownFolderName.MsgFolderRoot)
            'Dim dirRoot As Directory = folderRoot.Directory.SelectSubFolder("Inbox", False, e2007.DirectorySeparatorChar)

            ShowItems(dirRoot, e2007)
            ShowItems(Convert2Items(dirRoot, e2007, New Microsoft.Exchange.WebServices.Data.Item() {dirInbox.ItemsAsExchangeItem()(0)}))
            ShowItems(New Item() {dirInbox.Items()(0)})

            ShowItems(New Item() {dirInbox.MailboxItems(SearchDefault, ItemViewDefault)(0)})

            'Dim foldersBelowRoot As Directory() = e2007.ListFolderItems(folderRoot)
            'Dim foldersBelowRoot As Directory() = e2007.ListSubFoldersRecursively(folderRoot)
            'Dim foldersBelowRoot As Directory() = dirRoot.SubFolders
            Dim testSubFolder As Directory = dirRoot
            Console.WriteLine("TEST SUBS FOR: " & testSubFolder.DisplayName)
            Console.WriteLine("TEST SUBS FOR: " & testSubFolder.ID)
            Console.WriteLine("TEST SUBS FOR: " & testSubFolder.SubFolderCount)
            'Console.WriteLine("TEST SUBS FOR: " & testSubFolder.SubFolderCoun)
            'foldersBelowRoot = e2007.ListSubFolders(New FolderPathRepresentation(testSubFolder.ExchangeFolder.))

            'Dim itemView As New Microsoft.Exchange.WebServices.Data.ItemView(Integer.MaxValue, 0, Microsoft.Exchange.WebServices.Data.OffsetBasePoint.Beginning)
            'Dim searchFilter As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016 - 03 - 18))
            'Items = folderRoot.ExchangeFolder.FindItems(searchFilter, itemView)



            End

            Dim u As Uri = e2007.SaveMailAsDraft("test", "test <b>plain</b>", "", Nothing, Nothing, Nothing)
            e2007.SaveMailAsDraft("test", "", "text <b>html</b>", Nothing, Nothing, Nothing)
            Console.WriteLine(u.ToString)
            End
            'e2007.ResolveMailboxOrContactNames("jochen")
            'e2007.CreateFolder("Test", e2007.LookupFolder(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Inbox, "__Christ-Sein\CVJM\!Archiv", False))
            'e2007.CreateFolder("__Christ-Sein\CVJM\!Archiv\Test\Sub-Test", e2007.LookupFolder(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Inbox, "", False))
            'e2007.EmptyFolder(e2007.LookupFolder(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Inbox, "__Christ-Sein\CVJM\!Archiv\Test", False), DeleteMode.MoveToDeletedItems, False)
            'e2007.DeleteFolder(e2007.LookupFolder(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Inbox, "__Christ-Sein\CVJM\!Archiv\Test", False), DeleteMode.MoveToDeletedItems)
            'Dim MyFolder As FolderPathRepresentation = e2007.LookupFolder(WellKnownFolderName.PublicFoldersRoot, "Company Contacts", False)
            Dim MyFolder As Directory = dirRoot.SelectSubFolder("Inbox", False, e2007.DirectorySeparatorChar)
            'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(e2007.ListFolderItems(MyFolder)))
            Dim dt As DataTable
            dt = Directory.ItemsAsDataTable(MyFolder.Items)
            'dt = CompuMaster.Data.DataTables.CreateDataTableClone(e2007.ListFolderItems(MyFolder), "subject like '*schüppenhauer*' or subject like '*rund um boppard*'", "", 3)
            'dt = CompuMaster.Data.DataTables.CreateDataTableClone(e2007.ListFolderItems(MyFolder), "subject='Michael Höfler' or subject = 'Claudia Lamberti'", "", 3)
            'CompuMaster.Data.Csv.WriteDataTableToCsvFile("g:\cc.csv", dt)
            Dim ht As Hashtable = CompuMaster.Data.DataTables.FindDuplicates(dt.Columns("ID"))

            'dt.Rows.Add(dt.NewRow)
            'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(dt))
            'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(dt.Clone))
            Console.WriteLine(vbNewLine & "Data Rows: 2 first exemplary IDs:")
            Console.WriteLine(dt.Rows(0)("ID"))
            Console.WriteLine(dt.Rows(1)("ID"))
            Dim IDsAreEqual As Boolean = (dt.Rows(0)("ID").ToString = dt.Rows(1)("ID").ToString)
            If IDsAreEqual = False Then Console.WriteLine(Space(FirstDifferentChar(dt.Rows(0)("ID").ToString, dt.Rows(1)("ID").ToString)) & "^")
            Console.WriteLine("IDs are equal=" & IDsAreEqual.ToString.ToUpper)

            Console.WriteLine(vbNewLine & "DUPS:")
            For Each key As Object In ht.Keys
                Console.WriteLine(key.ToString & "=" & ht(key).ToString)
            Next
            'e2007.VerifyUniqueItemIDs(dt)

            'Console.WriteLine(vbnewline & "Re-Loading ID")
            'Dim c As Microsoft.Exchange.WebServices.Data.Contact = e2007.LoadContactData(Utils.NoDBNull(dt.Rows(0)("ID").ToString, ""))
            'Console.WriteLine(c.Subject)
            'c.Update(Microsoft.Exchange.WebServices.Data.ConflictResolutionMode.AutoResolve)

            End
            'e2007.SendMail("Test", "from CompuMaster.Data.Exchange2007SP1OrHigher" & vbNewLine & "on " & Now.ToString, New Recipient() {New Recipient("jwezel@compumaster.de")}, Nothing, Nothing)
            'e2007.CreateAppointment("Test-Appointment", "nowhere", "from CompuMaster.Data.Exchange2007SP1OrHigher" & vbNewLine & "on " & Now.ToString, Now.AddMinutes(5), New TimeSpan(0, 30, 0))
            'e2007.CreateMeetingAppointment("Test-Meeting", "nowhere", "from CompuMaster.Data.Exchange2007SP1OrHigher" & vbNewLine & "on " & Now.ToString, Now.AddMinutes(5), New TimeSpan(0, 30, 0), New Recipient() {New Recipient("jwezel@compumaster.de")}, Nothing, Nothing)
        Catch ex As Exception
            Console.WriteLine("Error: " + ex.ToString)
        End Try
    End Sub

    Private Function SearchDefault() As Microsoft.Exchange.WebServices.Data.SearchFilter
        Dim searchFilterCollection As New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.And)
        Dim searchFilterEarlierDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsGreaterThanOrEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016, 03, 10, 14, 00, 0))
        Dim searchFilterLaterDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsLessThanOrEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016, 03, 24, 14, 59, 59))
        searchFilterCollection.Add(searchFilterEarlierDate)
        searchFilterCollection.Add(searchFilterLaterDate)
        Return searchFilterCollection
    End Function
    Private Function ItemViewDefault() As Microsoft.Exchange.WebServices.Data.ItemView
        Dim itemView As New Microsoft.Exchange.WebServices.Data.ItemView(Integer.MaxValue, 0, Microsoft.Exchange.WebServices.Data.OffsetBasePoint.Beginning)
        itemView.OrderBy.Add(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, Microsoft.Exchange.WebServices.Data.SortDirection.Descending)
        'itemView.Traversal = Microsoft.Exchange.WebServices.Data.ItemTraversal.Associated
        Return itemView
    End Function

    Private Sub ShowItems(dir As Directory, e2007 As Exchange2007SP1OrHigher)

        Dim items As ObjectModel.Collection(Of Microsoft.Exchange.WebServices.Data.Item) = dir.ItemsAsExchangeItem(SearchDefault, ItemViewDefault)
        ShowItems(Convert2Items(dir, e2007, items))
    End Sub

    Private Function Convert2Items(dir As Directory, e2007 As Exchange2007SP1OrHigher, items As Microsoft.Exchange.WebServices.Data.Item()) As Item()
        Dim Result As New List(Of Item)
        For MyItemCounter As Integer = 0 To System.Math.Min(1, items.Length) - 1
            Result.Add(New Item(e2007, items(MyItemCounter), dir))
        Next
        Return Result.ToArray
    End Function
    Private Function Convert2Items(dir As Directory, e2007 As Exchange2007SP1OrHigher, items As ObjectModel.Collection(Of Microsoft.Exchange.WebServices.Data.Item)) As Item()
        Dim Result As New List(Of Item)
        For MyItemCounter As Integer = 0 To System.Math.Min(1, items.Count) - 1
            Result.Add(New Item(e2007, items(MyItemCounter), dir))
        Next
        Return Result.ToArray
    End Function

    Private Sub ShowItems(items As Item())

        Console.WriteLine("    ---")
        For MyItemCounter As Integer = 0 To System.Math.Min(1, items.Length) - 1
            Dim entryItem As Item = items(MyItemCounter)
            Console.WriteLine("    " & entryItem.Subject & " / DC:" & entryItem.DateTimeCreated & " / DR:" & entryItem.DateTimeReceived & " / DS:" & entryItem.DateTimeSent)
            'Console.WriteLine("    Co:" & entryItem.MimeContent)
            'Console.WriteLine("    BT: " & entryItem.BodyType)
            'Console.WriteLine("    BC: " & entryItem.Body)
            'Console.WriteLine("    Fr: " & Utils.ObjectNotNothingOrEmptyString(entryItem.FromSender).ToString)

            Console.WriteLine("    Fr: " & entryItem.FromExchangeSender)
            Console.WriteLine("    To: " & entryItem.DisplayTo)
            Console.WriteLine("    Cc: " & entryItem.DisplayCc)
            Console.WriteLine("    Pa: " & entryItem.ParentDirectory.DisplayPath)
            Console.WriteLine("    ---")
        Next
    End Sub

    Private Sub ForEachSubDirectory(dir As Directory, e2007 As Exchange2007SP1OrHigher)

        For Each dirItem As Directory In dir.SubFolders
            Console.Write(dirItem.ToString)
            Console.WriteLine(" (F:" & dirItem.SubFolderCount & " / U:" & dirItem.ItemUnreadCount & " / T:" & dirItem.ItemCount & ")")

            'Dim itemView As New Microsoft.Exchange.WebServices.Data.ItemView(Integer.MaxValue, 0, Microsoft.Exchange.WebServices.Data.OffsetBasePoint.Beginning)
            'Dim searchFilterCollection As New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.And)
            'Dim searchFilterEarlierDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsGreaterThanOrEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016, 03, 18, 14, 00, 0))
            'Dim searchFilterLaterDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsLessThanOrEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016, 03, 18, 14, 59, 59))
            'searchFilterCollection.Add(searchFilterEarlierDate)
            'searchFilterCollection.Add(searchFilterLaterDate)

            ''Dim itemsEApi As Microsoft.Exchange.WebServices.Data.FindItemsResults(Of Microsoft.Exchange.WebServices.Data.Item) = dirItem.ExchangeFolder.FindItems(searchFilterCollection, itemView)
            'Dim items As ObjectModel.Collection(Of Microsoft.Exchange.WebServices.Data.Item) = dirItem.Items(searchFilterCollection, itemView)
            ''Dim items As ObjectModel.Collection(Of Microsoft.Exchange.WebServices.Data.Item) = dirItem.Items()
            ''If itemsEApi.Items.Count <> items.Count Or items.Count <> e2007.ListFolderItemsAsExchangeItems(dirItem).Length Then
            ''    Console.WriteLine("!!" & dirItem.ToString & " (" & e2007.ListFolderItemsAsExchangeItems(dirItem).Length & " of " & dirItem.ItemCount & ")")
            ''End If
            ''Console.WriteLine("    FType: " & dirItem.FolderClass)

            'Dim EndCounter As Integer
            'EndCounter += 1

            ''For Each editem As Generic.KeyValuePair(Of String, Object) In dirItem.ExtendedData
            ''    If editem.Value Is Nothing Then
            ''        Console.WriteLine("         " & editem.Key & "={NULL}")
            ''    Else
            ''        Console.WriteLine("         " & editem.Key & "=" & editem.Value.ToString)
            ''    End If
            ''Next
            ''If EndCounter >= 10 Then End

            'If False AndAlso True OrElse dirItem.DisplayPath.Contains("Inbox") Then

            '    For MyItemCounter As Integer = 0 To System.Math.Min(1, items.Count) - 1
            '        Dim entryItem As Item
            '        entryItem = New Item(e2007, items.Item(MyItemCounter), dirItem)
            '        Console.WriteLine("    " & entryItem.Subject & " / DC:" & entryItem.DateTimeCreated & " / DR:" & entryItem.DateTimeReceived & " / DS:" & entryItem.DateTimeSent)
            '        'Console.WriteLine("    Co:" & entryItem.MimeContent)
            '        'Console.WriteLine("    BT: " & entryItem.BodyType)
            '        'Console.WriteLine("    BC: " & entryItem.Body)
            '        'Console.WriteLine("    Fr: " & Utils.ObjectNotNothingOrEmptyString(entryItem.FromSender).ToString)

            '        Console.WriteLine("    Fr: " & entryItem.FromExchangeSender)
            '        Console.WriteLine("    To: " & entryItem.DisplayTo)
            '        Console.WriteLine("    Cc: " & entryItem.DisplayCc)
            '        Console.WriteLine("    Fr: " & entryItem.ParentDirectory.DisplayPath)
            '        'For Each addr As System.Net.Mail.MailAddress In entryItem.RecipientTo
            '        '    Console.WriteLine("    TO: " & addr.ToString)
            '        'Next
            '        'For Each addr As System.Net.Mail.MailAddress In entryItem.RecipientCc
            '        '    Console.WriteLine("    CC: " & addr.ToString)
            '        'Next
            '        'For Each addr As System.Net.Mail.MailAddress In entryItem.RecipientBcc
            '        '    Console.WriteLine("    BCC: " & addr.ToString)
            '        'Next
            '        'For Each addr As System.Net.Mail.MailAddress In entryItem.ReplyTo
            '        '    Console.WriteLine("    Repl: " & addr.ToString)
            '        'Next

            '        ''Console.WriteLine("T: " & entryItem.BodyText)
            '        ''Console.WriteLine("H: " & entryItem.BodyHtml)
            '        'For Each editem As Generic.KeyValuePair(Of String, Object) In entryItem.ExtendedData
            '        '    If editem.Value Is Nothing Then
            '        '        Console.WriteLine("         " & editem.Key & "={NULL}")
            '        '    Else
            '        '        Console.WriteLine("         " & editem.Key & "=" & editem.Value.ToString)
            '        '    End If
            '        'Next

            '    Next

            'End If

            If dirItem.DisplayPath.Contains("Technik") Then
                ForEachSubDirectory(dirItem, e2007)
            Else
                ForEachSubDirectory(dirItem, e2007)
            End If
        Next

    End Sub

    Public Function FirstDifferentChar(ByVal value1 As String, ByVal value2 As String) As Integer
        Dim charCounter As Integer
        For charCounter = 0 To value1.Length
            If value1(charCounter) <> value2(charCounter) Then Return charCounter
        Next
        Return charCounter
    End Function
End Module
