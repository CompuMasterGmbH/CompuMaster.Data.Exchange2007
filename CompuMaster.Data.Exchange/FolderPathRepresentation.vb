Option Strict On
Option Explicit On

Imports Microsoft.Exchange.WebServices.Data
Imports System.Net

Namespace CompuMaster.Data.MsExchange

    ''' <summary>
    ''' A representation of a folder path/ID
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FolderPathRepresentation

        Private _root As CompuMaster.Data.MsExchange.Exchange2007SP1OrHigher.WellKnownFolderName = Nothing
        'Private _subfolder As String = Nothing
        Private _folderID As String = Nothing
        Private _exchangeWrapper As CompuMaster.Data.MsExchange.Exchange2007SP1OrHigher = Nothing
        Private _exchangeFolder As Microsoft.Exchange.WebServices.Data.Folder

        Public Sub New(ByVal exchange As Exchange2007SP1OrHigher, ByVal folderID As String)
            _exchangeWrapper = exchange
            _folderID = folderID
        End Sub

        Friend Sub New(ByVal exchange As Exchange2007SP1OrHigher, ByVal folder As Folder)
            Me.New(exchange, folder.Id)
            _exchangeFolder = folder
        End Sub

        Friend Sub New(ByVal exchange As Exchange2007SP1OrHigher, ByVal folderID As FolderId)
            Me.New(exchange, folderID.UniqueId)
        End Sub

        Public Sub New(ByVal exchange As Exchange2007SP1OrHigher, ByVal root As CompuMaster.Data.MsExchange.Exchange2007SP1OrHigher.WellKnownFolderName)
            _root = root
            _exchangeWrapper = exchange
        End Sub

        'Public Sub New(ByVal exchange As Exchange2007SP1OrHigher, ByVal root As CompuMaster.Data.MsExchange.Exchange2007SP1OrHigher.WellKnownFolderName, ByVal subfolderName As String)
        '    _root = root
        '    _subfolder = subfolderName
        '    _exchangeWrapper = exchange
        'End Sub

        Public ReadOnly Property Folder As Microsoft.Exchange.WebServices.Data.Folder
            Get
                If _exchangeFolder Is Nothing Then
                    _exchangeFolder = _exchangeWrapper.LookupFolder(Me._root).Folder
                End If
                Return _exchangeFolder
            End Get
        End Property


        ''' <summary>
        ''' The folder ID as used in Exchange
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FolderID() As String
            If Not _folderID Is Nothing Then
                Return _folderID
            Else
                Return _exchangeWrapper.LookupFolder(Me._root).FolderID
            End If
        End Function

        Private _Directory As Directory
        Public ReadOnly Property Directory As Directory
            Get
                If _Directory Is Nothing Then

                    _Directory = New Directory(_exchangeWrapper, Me.Folder, CType(Nothing, Folder))
                End If
                Return _Directory
            End Get
        End Property
    End Class

End Namespace