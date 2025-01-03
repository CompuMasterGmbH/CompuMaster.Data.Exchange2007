﻿Option Strict On
Option Explicit On

Imports Microsoft.Exchange.WebServices.Data
Imports System.Net
Imports System.Data

Namespace CompuMaster.Data.MsExchange

    ''' <summary>
    ''' Represents a folder in Exchange
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Directory

        Private ReadOnly _exchangeWrapper As Exchange2007SP1OrHigher
        Private ReadOnly _IsRootElementForSubFolderQuery As Boolean = False
        Private _SubFoldersAlreadyPutIntoHierarchy As Boolean = False
        Private ReadOnly _folder As Folder
        Private _parentDirectory As Directory
        Private _parentFolder As Folder

        Public Sub New(exchangeWrapper As Exchange2007SP1OrHigher, ByVal folder As Folder)
            _folder = folder
            _IsRootElementForSubFolderQuery = True
            _exchangeWrapper = exchangeWrapper
        End Sub

        Friend Sub New(exchangeWrapper As Exchange2007SP1OrHigher, ByVal folder As Folder, ByVal parentFolder As Folder)
            _folder = folder
            _parentFolder = parentFolder
            If parentFolder Is Nothing Then
                _IsRootElementForSubFolderQuery = True
            End If
            _exchangeWrapper = exchangeWrapper
        End Sub

        Friend Sub New(exchangeWrapper As Exchange2007SP1OrHigher, ByVal folder As Folder, ByVal parentDirectory As Directory)
            _folder = folder
            _parentDirectory = parentDirectory
            If parentDirectory Is Nothing Then
                _IsRootElementForSubFolderQuery = True
            End If
            _exchangeWrapper = exchangeWrapper
        End Sub

        ''' <summary>
        ''' Gives full access to the Exchange Managed API for this folder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ExchangeFolder() As Folder
            Get
                Return _folder
            End Get
        End Property

        Private Shared Sub ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(ByVal value As String, ByVal row As DataRow, ByVal columnName As String)
            If row.Table.Columns.Contains(columnName) = False Then
                row.Table.Columns.Add(columnName, GetType(String))
            End If
            row(columnName) = value
        End Sub

        ''' <summary>
        ''' List available items of a folder
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ItemsAsDataTable(items As Item()) As System.Data.DataTable
            Dim Result As New DataTable("items")
            Dim ProcessedSchemas As New ArrayList
            Dim Columns As New Hashtable
            'Add all items into the result table with all of their properties as complete as possible
            For Each MyItem As Item In items
                'Add required additional columns if not yet done
                If ProcessedSchemas.Contains(MyItem.ExchangeItem.Schema) = False Then
                    For Each prop As PropertyDefinition In MyItem.ExchangeItem.Schema
                        Dim ColName As String = prop.Name
                        If prop.Version <> 0 Then ColName &= "_V" & prop.Version
                        If Not Result.Columns.Contains(ColName) Then
                            If prop.Type.ToString.StartsWith("System.Nullable") Then
                                'Dataset doesn't support System.Nullable --> use System.Object
                                Columns.Add(ColName, New FolderItemPropertyToColumn(prop, Result.Columns.Add(ColName, GetType(Object))))
                            Else
                                'Use the property type as regular
                                Columns.Add(ColName, New FolderItemPropertyToColumn(prop, Result.Columns.Add(ColName, prop.Type)))
                            End If
                        End If
                    Next
                End If
                'Add item as new data row
                Dim row As System.Data.DataRow = Result.NewRow
                For Each key As Object In Columns.Keys
                    Dim MyColumn As FolderItemPropertyToColumn = CType(Columns(key), FolderItemPropertyToColumn)
                    If MyColumn.SchemaProperty IsNot Nothing Then
                        Try
                            If MyItem.ExchangeItem.Item(MyColumn.SchemaProperty) Is Nothing Then
                                row(MyColumn.Column) = DBNull.Value
                            Else
                                Select Case MyItem.ExchangeItem.Item(MyColumn.SchemaProperty).GetType.ToString
                                    Case GetType(Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection
                                        value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection)
                                        'Dim comp As String = MyItem.Item(CType(Columns("Subject"), FolderItemPropertyToColumn).SchemaProperty).ToString
                                        'If comp.IndexOf("Wezel") > -1 Then
                                        '    Debug.Print(value.ToString)
                                        'End If
                                        'For Each valueKey As ExtendedProperty In value
                                        '    Debug.Print(valueKey.PropertyDefinition.Name & "=" & valueKey.Value.ToString)
                                        'Next
                                    Case GetType(Microsoft.Exchange.WebServices.Data.CompleteName).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.CompleteName
                                        value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.CompleteName)
                                        ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value.Title, row, MyColumn.Column.ColumnName & "_Title")
                                    Case GetType(Microsoft.Exchange.WebServices.Data.EmailAddressDictionary).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.EmailAddressDictionary
                                        value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.EmailAddressDictionary)
                                        If value.Contains(EmailAddressKey.EmailAddress1) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(EmailAddressKey.EmailAddress1).Address, row, MyColumn.Column.ColumnName & "_Email1")
                                        If value.Contains(EmailAddressKey.EmailAddress2) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(EmailAddressKey.EmailAddress2).Address, row, MyColumn.Column.ColumnName & "_Email2")
                                        If value.Contains(EmailAddressKey.EmailAddress3) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(EmailAddressKey.EmailAddress3).Address, row, MyColumn.Column.ColumnName & "_Email3")
                                    Case GetType(Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary
                                        value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary)
                                        If value.Contains(PhysicalAddressKey.Business) Then
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).Street, row, MyColumn.Column.ColumnName & "_Business_Street")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).PostalCode, row, MyColumn.Column.ColumnName & "_Business_PostalCode")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).City, row, MyColumn.Column.ColumnName & "_Business_City")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).State, row, MyColumn.Column.ColumnName & "_Business_State")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Business).CountryOrRegion, row, MyColumn.Column.ColumnName & "_Business_CountryOrRegion")
                                        End If
                                        If value.Contains(PhysicalAddressKey.Home) Then
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).Street, row, MyColumn.Column.ColumnName & "_Home_Street")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).PostalCode, row, MyColumn.Column.ColumnName & "_Home_PostalCode")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).City, row, MyColumn.Column.ColumnName & "_Home_City")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).State, row, MyColumn.Column.ColumnName & "_Home_State")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Home).CountryOrRegion, row, MyColumn.Column.ColumnName & "_Home_CountryOrRegion")
                                        End If
                                        If value.Contains(PhysicalAddressKey.Other) Then
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).Street, row, MyColumn.Column.ColumnName & "_Other_Street")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).PostalCode, row, MyColumn.Column.ColumnName & "_Other_PostalCode")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).City, row, MyColumn.Column.ColumnName & "_Other_City")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).State, row, MyColumn.Column.ColumnName & "_Other_State")
                                            ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhysicalAddressKey.Other).CountryOrRegion, row, MyColumn.Column.ColumnName & "_Other_CountryOrRegion")
                                        End If
                                    Case GetType(Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary
                                        value = CType(MyItem.ExchangeItem.Item(MyColumn.SchemaProperty), Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary)
                                        If value.Contains(PhoneNumberKey.BusinessPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.BusinessPhone), row, MyColumn.Column.ColumnName & "_BusinessPhone")
                                        If value.Contains(PhoneNumberKey.BusinessPhone2) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.BusinessPhone2), row, MyColumn.Column.ColumnName & "_BusinessPhone2")
                                        If value.Contains(PhoneNumberKey.BusinessFax) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.BusinessFax), row, MyColumn.Column.ColumnName & "_BusinessFax")
                                        If value.Contains(PhoneNumberKey.CompanyMainPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.CompanyMainPhone), row, MyColumn.Column.ColumnName & "_CompanyMainPhone")
                                        If value.Contains(PhoneNumberKey.CarPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.CarPhone), row, MyColumn.Column.ColumnName & "_CarPhone")
                                        If value.Contains(PhoneNumberKey.Callback) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.Callback), row, MyColumn.Column.ColumnName & "_Callback")
                                        If value.Contains(PhoneNumberKey.AssistantPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.AssistantPhone), row, MyColumn.Column.ColumnName & "_AssistantPhone")
                                        If value.Contains(PhoneNumberKey.HomeFax) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.HomeFax), row, MyColumn.Column.ColumnName & "_HomeFax")
                                        If value.Contains(PhoneNumberKey.HomePhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.HomePhone), row, MyColumn.Column.ColumnName & "_HomePhone")
                                        If value.Contains(PhoneNumberKey.HomePhone2) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.HomePhone2), row, MyColumn.Column.ColumnName & "_HomePhone2")
                                        If value.Contains(PhoneNumberKey.MobilePhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.MobilePhone), row, MyColumn.Column.ColumnName & "_MobilePhone")
                                        If value.Contains(PhoneNumberKey.OtherFax) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.OtherFax), row, MyColumn.Column.ColumnName & "_OtherFax")
                                        If value.Contains(PhoneNumberKey.OtherTelephone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.OtherTelephone), row, MyColumn.Column.ColumnName & "_OtherTelephone")
                                        If value.Contains(PhoneNumberKey.PrimaryPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.PrimaryPhone), row, MyColumn.Column.ColumnName & "_PrimaryPhone")
                                        If value.Contains(PhoneNumberKey.RadioPhone) Then ItemsAsDataTable_AssignValueToColumnOrJitCreateColumn(value(PhoneNumberKey.RadioPhone), row, MyColumn.Column.ColumnName & "_RadioPhone")
                                    Case Else
                                        row(MyColumn.Column) = MyItem.ExchangeItem.Item(MyColumn.SchemaProperty)
                                End Select
                            End If
                        Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                            'Mark this column to be killed at the end because it only contains non-sense
                            MyColumn.SchemaProperty = Nothing
                        Catch ex As Microsoft.Exchange.WebServices.Data.ServiceVersionException
                            'Mark this column to be killed at the end because it only contains non-sense
                            MyColumn.SchemaProperty = Nothing
                        End Try
                    End If
                Next
                Result.Rows.Add(row)
            Next
            'Remove all columns which are marked to be deleted
            For Each key As Object In Columns.Keys
                If CType(Columns(key), FolderItemPropertyToColumn).SchemaProperty Is Nothing Then
                    'Missing data indicates a column to be deleted
                    Result.Columns.Remove(CType(Columns(key), FolderItemPropertyToColumn).Column)
                End If
            Next
            Result.Columns("ID").Unique = True
            Return Result
        End Function

        ''' <summary>
        ''' Schema information to a column
        ''' </summary>
        ''' <remarks></remarks>
        Private Class FolderItemPropertyToColumn
            Public Sub New(ByVal schemaProperty As PropertyDefinition, ByVal column As System.Data.DataColumn)
                Me.Column = column
                Me.SchemaProperty = schemaProperty
            End Sub
            Public Column As System.Data.DataColumn
            Public SchemaProperty As PropertyDefinition
        End Class

        ''' <summary>
        ''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        ''' </summary>
        ''' <returns></returns>
        Public Function ItemsAsExchangeItem() As ObjectModel.Collection(Of Microsoft.Exchange.WebServices.Data.Item)
            Return Me.ExchangeFolder.FindItems(New ItemView(Integer.MaxValue)).Result.Items
        End Function

        ''' <summary>
        ''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        ''' </summary>
        ''' <returns></returns>
        Public Function Items() As Item()
            Dim Result As New List(Of Item)
            For Each ExchangeItem As Microsoft.Exchange.WebServices.Data.Item In ItemsAsExchangeItem()
                Result.Add(New Item(Me._exchangeWrapper, ExchangeItem, Me))
            Next
            Return Result.ToArray
        End Function

        ''' <summary>
        ''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        ''' </summary>
        ''' <returns></returns>
        Public Function ItemsAsExchangeItem(searchFilter As Microsoft.Exchange.WebServices.Data.SearchFilter, itemView As Microsoft.Exchange.WebServices.Data.ItemView) As List(Of Microsoft.Exchange.WebServices.Data.Item)
            Dim FoundItems As New List(Of Microsoft.Exchange.WebServices.Data.Item)
            Dim MaxQueryItems As Integer = itemView.PageSize
            Dim MoreResultsAvailable As Boolean = True

            'Repeatedly query all partly results and combine them
            Do While MoreResultsAvailable
                Dim ItemViewWithOffset As Microsoft.Exchange.WebServices.Data.ItemView = itemView
                If MaxQueryItems = Integer.MaxValue Then
                    ItemViewWithOffset.Offset = FoundItems.Count
                End If
                Dim QueryResult As FindItemsResults(Of Microsoft.Exchange.WebServices.Data.Item) = Me.ExchangeFolder.FindItems(searchFilter, ItemViewWithOffset).Result
                For Each item As Microsoft.Exchange.WebServices.Data.Item In QueryResult.Items
                    FoundItems.Add(item)
                Next
                MoreResultsAvailable = QueryResult.MoreAvailable AndAlso FoundItems.Count < MaxQueryItems
            Loop

            Return FoundItems
        End Function

        ''' <summary>
        ''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        ''' </summary>
        ''' <returns></returns>
        Public Function Items(searchFilter As Microsoft.Exchange.WebServices.Data.SearchFilter, itemView As Microsoft.Exchange.WebServices.Data.ItemView) As Item()
            Dim Result As New List(Of Item)
            For Each ExchangeItem As Microsoft.Exchange.WebServices.Data.Item In ItemsAsExchangeItem(searchFilter, itemView)
                Result.Add(New Item(Me._exchangeWrapper, ExchangeItem, Me))
            Next
            Return Result.ToArray
        End Function

        ''' <summary>
        ''' All items of a folder (might be limited due to exchange default to e.g. 1,000 items)
        ''' </summary>
        ''' <returns></returns>
        Public Function MailboxItems(searchFilter As Microsoft.Exchange.WebServices.Data.SearchFilter, itemView As Microsoft.Exchange.WebServices.Data.ItemView) As Item()
            Dim searchFolder As Directory = Me.InitialRootDirectory.SelectSubFolder("AllItems", False, Me._exchangeWrapper.DirectorySeparatorChar)
            Dim Result As New List(Of Item)
            For Each ExchangeItem As Microsoft.Exchange.WebServices.Data.Item In searchFolder.ItemsAsExchangeItem(searchFilter, itemView)
                Result.Add(New Item(Me._exchangeWrapper, ExchangeItem, Me))
            Next
            Return Result.ToArray
        End Function

        ''' Total amount of items of a folder
        Public Function ItemCount() As Integer
            Return Me.ExchangeFolder.TotalCount
        End Function

        ''' <summary>
        ''' Number of unread items of a folder
        ''' </summary>
        ''' <returns></returns>
        Public Function ItemUnreadCount() As Integer
            Try
                Return Me.ExchangeFolder.UnreadCount
            Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                Return 0
            End Try
        End Function

        Private _ParentFolderID As String
        ''' <summary>
        ''' The unique ID of the parent folder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ParentFolderID() As String
            Get
                If _ParentFolderID Is Nothing Then
                    _ParentFolderID = _folder.ParentFolderId.UniqueId
                End If
                Return _ParentFolderID
            End Get
        End Property

        Private _ID As String
        ''' <summary>
        ''' The unique ID of the folder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ID() As String
            Get
                If _ID Is Nothing Then
                    _ID = _folder.Id.UniqueId
                End If
                Return _ID
            End Get
        End Property

        Private _FolderClass As String
        ''' <summary>
        ''' The folder class name
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FolderClass() As String
            Get
                If _FolderClass Is Nothing Then
                    _FolderClass = _folder.FolderClass
                End If
                Return _FolderClass
            End Get
            Set(ByVal value As String)
                _FolderClass = value
                _folder.FolderClass = value
                Me.CachedFolderDisplayPath = Nothing
            End Set
        End Property

        Private _DisplayName As String
        ''' <summary>
        ''' The display name of the folder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DisplayName() As String
            Get
                If _DisplayName Is Nothing Then
                    _DisplayName = _folder.DisplayName
                End If
                Return _DisplayName
            End Get
            Set(ByVal value As String)
                _DisplayName = value
                _folder.DisplayName = value
                Me.CachedFolderDisplayPath = Nothing
            End Set
        End Property

        Private CachedFolderDisplayPath As String = Nothing
        ''' <summary>
        ''' The path of the user separated by back-slashes (\)
        ''' </summary>
        ''' <returns>Existing back-slashes in folder's display names might confuse here - in case that back-slahes are possible in display names</returns>
        Public ReadOnly Property DisplayPath As String
            Get
                If Me.ParentDirectory IsNot Nothing Then
                    Return Me.ParentDirectory.DisplayPath & "\" & Me.DisplayName
                Else
                    Return Me.DisplayName
                End If
            End Get
        End Property

        ''' <summary>
        ''' The relative folder path starting from your initial top directory
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>Parent folder structure can't be looked up till root folder, that's why it's only up to your initial top directory</remarks>
        Public ReadOnly Property ParentDirectory As Directory
            Get
                If InitialRootDirectory._SubFoldersAlreadyPutIntoHierarchy = False Then
                    InitialRootDirectory.PutSubFoldersIntoHierarchy()
                End If
                If _parentDirectory Is Nothing AndAlso _parentFolder IsNot Nothing Then
                    _parentDirectory = New Directory(_exchangeWrapper, _parentFolder)
                ElseIf _parentDirectory Is Nothing AndAlso _parentFolder Is Nothing Then
                    '_parentFolder = New Directory(Me.ExchangeFolder.ParentFolderId.UniqueId)
                End If
                Return _parentDirectory
            End Get
        End Property

        ''' <summary>
        ''' Sorting required in advance before an item of a directory (e.g. in \AllItems but pointing to {MsgFolderRoot}\Inbox) looks up its parent directory
        ''' </summary>
        Private Sub PutSubFoldersIntoHierarchy()
            For Each subDir As Directory In Me.SubFolders
                subDir.PutSubFoldersIntoHierarchy()
            Next
            _SubFoldersAlreadyPutIntoHierarchy = True
        End Sub

        ''' <summary>
        ''' The default view for folders
        ''' </summary>
        ''' <returns></returns>
        Friend Shared Function DefaultFolderView(folderTraversal As FolderTraversal, offSet As Integer) As FolderView
            Dim Result As New FolderView(Integer.MaxValue, offSet) With {
                .PropertySet = DefaultPropertySet(),
                .Traversal = folderTraversal
            }
            Return Result
        End Function

        Friend Shared Function DefaultPropertySet() As PropertySet
            Dim AdditionalProperties As New List(Of Microsoft.Exchange.WebServices.Data.PropertyDefinition) From {
                Microsoft.Exchange.WebServices.Data.FolderSchema.ChildFolderCount,
                Microsoft.Exchange.WebServices.Data.FolderSchema.TotalCount,
                Microsoft.Exchange.WebServices.Data.FolderSchema.UnreadCount,
                Microsoft.Exchange.WebServices.Data.FolderSchema.FolderClass,
                Microsoft.Exchange.WebServices.Data.FolderSchema.Id,
                Microsoft.Exchange.WebServices.Data.FolderSchema.ParentFolderId,
                Microsoft.Exchange.WebServices.Data.FolderSchema.DisplayName
            }
            Return New PropertySet(BasePropertySet.FirstClassProperties, AdditionalProperties.ToArray)
        End Function

        ''' <summary>
        ''' Query the sub directories of this directory - deep traversal
        ''' </summary>
        Private Function QuerySubFoldersOfSeveralHierachyLevels() As List(Of Directory)
            Dim folderTraversal As FolderTraversal = FolderTraversal.Deep
            Dim FoundFolders As New List(Of Microsoft.Exchange.WebServices.Data.Folder)
            Dim MoreResultsAvailable As Boolean = True

            'Repeatedly query all partly results and combine them
            Do While MoreResultsAvailable
                Dim folders As FindFoldersResults = Me.ExchangeFolder.FindFolders(DefaultFolderView(folderTraversal, FoundFolders.Count)).Result
                For Each folder As Folder In folders
                    FoundFolders.Add(folder)
                Next
                MoreResultsAvailable = folders.MoreAvailable
            Loop

            Return SubFolders2DirectoryHierarchy(FoundFolders, Me)
        End Function

        Private Function SubFolders2DirectoryHierarchy(folders As List(Of Microsoft.Exchange.WebServices.Data.Folder), parentDirectory As Directory) As List(Of Directory)
            If parentDirectory Is Nothing Then Throw New ArgumentNullException(NameOf(parentDirectory))
            'hierarchy tree -> folder results might be (sub-)grand-children
            Dim FoundDirectories As New List(Of Directory)
            For Each folder As Folder In folders
                FoundDirectories.Add(New Directory(_exchangeWrapper, folder, parentDirectory))
            Next

            Return FoundDirectories
        End Function

        'Private Function FindSubFoldersInDataOfSeveralHierachyLevels(uniqueFolderID As String) As List(Of Directory)
        '    If _SubFoldersOfSeveralHierachyLevels Is Nothing Then

        '    End If
        'End Function

        Private _SubFoldersOfSeveralHierachyLevels As List(Of Directory)
        Friend ReadOnly Property SubFoldersOfSeveralHierachyLevels As List(Of Directory)
            Get
                If _SubFoldersOfSeveralHierachyLevels IsNot Nothing Then
                    'return cached list
                    Return _SubFoldersOfSeveralHierachyLevels
                ElseIf Me._IsRootElementForSubFolderQuery = True Then
                    'query list
                    _SubFoldersOfSeveralHierachyLevels = QuerySubFoldersOfSeveralHierachyLevels()
                    Return _SubFoldersOfSeveralHierachyLevels
                ElseIf Me._parentDirectory IsNot Nothing Then
                    'use parent's list
                    Return Me._parentDirectory.SubFoldersOfSeveralHierachyLevels
                Else
                    'no list avaible and this is no root dir and there is no other parent - should not happen !?!
                    Throw New InvalidOperationException("no list avaible and this is no root dir and there is no other parent")
                End If
            End Get
        End Property

        Private _SubFolders As List(Of Directory)
        Public ReadOnly Property SubFolders As Directory()
            Get
                If _SubFolders Is Nothing Then
                    If Me.SubFoldersOfSeveralHierachyLevels IsNot Nothing Then
                        'fill from hierarchy list
                        _SubFolders = New List(Of Directory)
                        For MyCounter As Integer = 0 To Me.SubFoldersOfSeveralHierachyLevels.Count - 1
                            Dim childDir As Directory = Me.SubFoldersOfSeveralHierachyLevels(MyCounter)
                            If childDir.ParentFolderID = Me.ID Then
                                'found a child folder
                                childDir.Internal_SetParentDirectory(Me)
                                _SubFolders.Add(childDir)
                            End If
                        Next
                    Else
                        Throw New InvalidOperationException("No folder hierarchy structure available")
                    End If
                End If
                Return _SubFolders.ToArray
            End Get
        End Property

        Private Sub Internal_SetParentDirectory(parentDirectory As Directory)
            Me._parentDirectory = parentDirectory
            Me._parentFolder = parentDirectory.ExchangeFolder
            Me._ParentFolderID = parentDirectory.ExchangeFolder.Id.UniqueId
        End Sub

        Public ReadOnly Property InitialRootDirectory As Directory
            Get
                If _parentDirectory IsNot Nothing Then
                    Return _parentDirectory.InitialRootDirectory
                Else
                    Return Me
                End If
            End Get
        End Property

        ''' <summary>
        ''' Lookup a directory based on its directory structure
        ''' </summary>
        ''' <param name="subfolder">A string containing the relative folder path, e.g. &quot;Inbox\Done&quot;</param>
        ''' <param name="searchCaseInsensitive">Ignore upper/lower case differences</param>
        ''' <param name="directorySeparatorChar"></param>
        ''' <returns></returns>
        Public Function SelectSubFolder(subFolder As String, ByVal searchCaseInsensitive As Boolean, directorySeparatorChar As Char) As Directory
            If subFolder = Nothing Then
                Return Me
            ElseIf subFolder.StartsWith(directorySeparatorChar) Then
                Throw New ArgumentException("subFolder can't start with a directorySeparatorChar", NameOf(subFolder))
            Else
                Dim subfoldersSplitted As String() = subFolder.Split(directorySeparatorChar)
                Dim nextSubFolder As String = subfoldersSplitted(0)
                For Each mySubfolder As Directory In Me.SubFolders
                    If mySubfolder.DisplayName = nextSubFolder OrElse (searchCaseInsensitive AndAlso mySubfolder.DisplayName.ToLowerInvariant = nextSubFolder.ToLowerInvariant) Then
                        If subfoldersSplitted.Length > 1 Then
                            'recursive call required

                            Return mySubfolder.SelectSubFolder(String.Join(directorySeparatorChar, subfoldersSplitted, 1, subfoldersSplitted.Length - 1), searchCaseInsensitive, directorySeparatorChar)
                        Else
                            'this is the last recursion - just return our current path item
                            Return mySubfolder
                        End If
                    End If
                Next
                Throw New Exception("Folder """ & subFolder & """ hasn't been found in " & Me.DisplayPath)
            End If
        End Function

        ''' <summary>
        ''' The number of child folders
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SubFolderCount() As Integer
            Get
                Return _folder.ChildFolderCount
            End Get
        End Property

        ''' <summary>
        ''' Save changes to this folder
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Save()
            _folder.Update()
        End Sub

        ''' <summary>
        ''' Save this folder as sub folder of the specified one
        ''' </summary>
        ''' <param name="parentFolder"></param>
        ''' <remarks></remarks>
        Public Sub Save(ByVal parentFolder As Directory)
            _folder.Save(New FolderId(parentFolder.ID))
        End Sub

        Public Overrides Function ToString() As String
            Return Me.DisplayPath
        End Function

        Private _ExtendedData As Generic.Dictionary(Of String, Object)
        Public Function ExtendedData() As Generic.Dictionary(Of String, Object)
            If _ExtendedData Is Nothing Then
                'Load first class props
                'Dim propSet As New Microsoft.Exchange.WebServices.Data.PropertySet(Microsoft.Exchange.WebServices.Data.BasePropertySet.FirstClassProperties)
                'Microsoft.Exchange.WebServices.Data.EmailMessage.Bind(_service.CreateConfiguredExchangeService, _exchangeItem.Id, propSet)
                _ExtendedData = New Generic.Dictionary(Of String, Object)
                'Add all items into the result table with all of their properties as complete as possible
                'Add required additional columns if not yet done
                For Each prop As PropertyDefinition In Me.ExchangeFolder.Schema
                    Dim ColName As String = prop.Name
                    If prop.Version <> 0 Then ColName &= "_V" & prop.Version
                    Try
                        If Me.ExchangeFolder.Item(prop) Is Nothing Then
                            _ExtendedData.Add(ColName, Nothing)
                        Else
                            Select Case Me.ExchangeFolder.Item(prop).GetType.ToString
                                Case GetType(Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection).ToString
                                    Dim value As Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection
                                    value = CType(Me.ExchangeFolder.Item(prop), Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection)
                                Case GetType(Microsoft.Exchange.WebServices.Data.CompleteName).ToString
                                    Dim value As Microsoft.Exchange.WebServices.Data.CompleteName
                                    value = CType(Me.ExchangeFolder.Item(prop), Microsoft.Exchange.WebServices.Data.CompleteName)
                                    _ExtendedData.Add(ColName & "_Title", value.Title)
                                Case GetType(Microsoft.Exchange.WebServices.Data.EmailAddressDictionary).ToString
                                    Dim value As Microsoft.Exchange.WebServices.Data.EmailAddressDictionary
                                    value = CType(Me.ExchangeFolder.Item(prop), Microsoft.Exchange.WebServices.Data.EmailAddressDictionary)
                                    If value.Contains(EmailAddressKey.EmailAddress1) Then _ExtendedData.Add(ColName & "_Email1", value(EmailAddressKey.EmailAddress1).Address)
                                    If value.Contains(EmailAddressKey.EmailAddress2) Then _ExtendedData.Add(ColName & "_Email2", value(EmailAddressKey.EmailAddress2).Address)
                                    If value.Contains(EmailAddressKey.EmailAddress3) Then _ExtendedData.Add(ColName & "_Email3", value(EmailAddressKey.EmailAddress3).Address)
                                Case GetType(Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary).ToString
                                    Dim value As Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary
                                    value = CType(Me.ExchangeFolder.Item(prop), Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary)
                                    If value.Contains(PhysicalAddressKey.Business) Then
                                        _ExtendedData.Add(ColName & "_Business_Street", value(PhysicalAddressKey.Business).Street)
                                        _ExtendedData.Add(ColName & "_Business_PostalCode", value(PhysicalAddressKey.Business).PostalCode)
                                        _ExtendedData.Add(ColName & "_Business_City", value(PhysicalAddressKey.Business).City)
                                        _ExtendedData.Add(ColName & "_Business_State", value(PhysicalAddressKey.Business).State)
                                        _ExtendedData.Add(ColName & "_Business_CountryOrRegion", value(PhysicalAddressKey.Business).CountryOrRegion)
                                    End If
                                    If value.Contains(PhysicalAddressKey.Home) Then
                                        _ExtendedData.Add(ColName & "_Home_Street", value(PhysicalAddressKey.Home).Street)
                                        _ExtendedData.Add(ColName & "_Home_PostalCode", value(PhysicalAddressKey.Home).PostalCode)
                                        _ExtendedData.Add(ColName & "_Home_City", value(PhysicalAddressKey.Home).City)
                                        _ExtendedData.Add(ColName & "_Home_State", value(PhysicalAddressKey.Home).State)
                                        _ExtendedData.Add(ColName & "_Home_CountryOrRegion", value(PhysicalAddressKey.Home).CountryOrRegion)
                                    End If
                                    If value.Contains(PhysicalAddressKey.Other) Then
                                        _ExtendedData.Add(ColName & "_Other_Street", value(PhysicalAddressKey.Other).Street)
                                        _ExtendedData.Add(ColName & "_Other_PostalCode", value(PhysicalAddressKey.Other).PostalCode)
                                        _ExtendedData.Add(ColName & "_Other_City", value(PhysicalAddressKey.Other).City)
                                        _ExtendedData.Add(ColName & "_Other_State", value(PhysicalAddressKey.Other).State)
                                        _ExtendedData.Add(ColName & "_Other_CountryOrRegion", value(PhysicalAddressKey.Other).CountryOrRegion)
                                    End If
                                Case GetType(Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary).ToString
                                    Dim value As Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary
                                    value = CType(Me.ExchangeFolder.Item(prop), Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary)
                                    If value.Contains(PhoneNumberKey.BusinessPhone) Then _ExtendedData.Add(ColName & "_BusinessPhone", value(PhoneNumberKey.BusinessPhone))
                                    If value.Contains(PhoneNumberKey.BusinessPhone2) Then _ExtendedData.Add(ColName & "_BusinessPhone2", value(PhoneNumberKey.BusinessPhone2))
                                    If value.Contains(PhoneNumberKey.BusinessFax) Then _ExtendedData.Add(ColName & "_BusinessFax", value(PhoneNumberKey.BusinessFax))
                                    If value.Contains(PhoneNumberKey.CompanyMainPhone) Then _ExtendedData.Add(ColName & "_CompanyMainPhone", value(PhoneNumberKey.CompanyMainPhone))
                                    If value.Contains(PhoneNumberKey.CarPhone) Then _ExtendedData.Add(ColName & "_CarPhone", value(PhoneNumberKey.CarPhone))
                                    If value.Contains(PhoneNumberKey.Callback) Then _ExtendedData.Add(ColName & "_Callback", value(PhoneNumberKey.Callback))
                                    If value.Contains(PhoneNumberKey.AssistantPhone) Then _ExtendedData.Add(ColName & "_AssistantPhone", value(PhoneNumberKey.AssistantPhone))
                                    If value.Contains(PhoneNumberKey.HomeFax) Then _ExtendedData.Add(ColName & "_HomeFax", value(PhoneNumberKey.HomeFax))
                                    If value.Contains(PhoneNumberKey.HomePhone) Then _ExtendedData.Add(ColName & "_HomePhone", value(PhoneNumberKey.HomePhone))
                                    If value.Contains(PhoneNumberKey.HomePhone2) Then _ExtendedData.Add(ColName & "_HomePhone2", value(PhoneNumberKey.HomePhone2))
                                    If value.Contains(PhoneNumberKey.MobilePhone) Then _ExtendedData.Add(ColName & "_MobilePhone", value(PhoneNumberKey.MobilePhone))
                                    If value.Contains(PhoneNumberKey.OtherFax) Then _ExtendedData.Add(ColName & "_OtherFax", value(PhoneNumberKey.OtherFax))
                                    If value.Contains(PhoneNumberKey.OtherTelephone) Then _ExtendedData.Add(ColName & "_OtherTelephone", value(PhoneNumberKey.OtherTelephone))
                                    If value.Contains(PhoneNumberKey.PrimaryPhone) Then _ExtendedData.Add(ColName & "_PrimaryPhone", value(PhoneNumberKey.PrimaryPhone))
                                    If value.Contains(PhoneNumberKey.RadioPhone) Then _ExtendedData.Add(ColName & "_RadioPhone", value(PhoneNumberKey.RadioPhone))
                                Case Else
                                    _ExtendedData.Add(ColName, Me.ExchangeFolder.Item(prop))
                            End Select
                        End If
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        'Mark this column to be killed at the end because it only contains non-sense
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceVersionException
                        'Mark this column to be killed at the end because it only contains non-sense
                    Catch ex As NullReferenceException
                        _ExtendedData.Add(ColName, Nothing)
                    End Try
                Next
            End If
            Return _ExtendedData
        End Function

    End Class

End Namespace