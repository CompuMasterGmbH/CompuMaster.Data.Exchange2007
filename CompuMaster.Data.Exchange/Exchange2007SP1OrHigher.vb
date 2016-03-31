Option Strict On
Option Explicit On

Imports Microsoft.Exchange.WebServices.Data
Imports System.Net

Namespace CompuMaster.Data.MsExchange

    ''' <summary>
    ''' This solution works only for Exchange 2007, 2010, 2010 SP1, 2013
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Exchange2007SP1OrHigher
        'Documenation available at: http://msdn.microsoft.com/en-us/library/dd633696(v=EXCHG.80).aspx

#Region "Base exchange connection"
        Protected _serverName As String
        Protected _autoDiscoverServerByEMailAddress As String
        Protected _exchangeVersion As ExchangeVersion

        'Private OutlookPropertySetID As New Guid("{00062004-0000-0000-C000-000000000046}")

        Public Enum ExchangeVersion As Integer
            Exchange2007_SP1 = Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2007_SP1
            Exchange2010 = Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2010
            Exchange2010_SP1 = Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2010_SP1
            Exchange2010_SP2 = Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2010_SP2
            Exchange2013 = Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2013
            Exchange2013_SP1 = Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2013_SP1
        End Enum

        Public Sub New(ByVal exchangeVersion As ExchangeVersion, ByVal serverName As String)
            Me.New(exchangeVersion, serverName, Nothing)
        End Sub

        Public Sub New(ByVal exchangeVersion As ExchangeVersion, ByVal serverName As String, ByVal autoDiscoverServerByEMailAddress As String)
            If Not (serverName <> Nothing Xor autoDiscoverServerByEMailAddress <> Nothing) Then Throw New ArgumentException("Either serverName or autoDiscoverServerByEMailAddress must be defined")
            _serverName = serverName
            _autoDiscoverServerByEMailAddress = autoDiscoverServerByEMailAddress
            _exchangeVersion = exchangeVersion
        End Sub

        ''' <summary>
        ''' Create an instance targetting the correct Exchange version
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Function CreateExchangeService() As ExchangeService
            Return New ExchangeService(CType(_exchangeVersion, Microsoft.Exchange.WebServices.Data.ExchangeVersion))
        End Function

        ''' <summary>
        ''' Configure a valid connection to the related Exchange server
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Overridable Function CreateConfiguredExchangeService() As ExchangeService
            Static service As ExchangeService
            If service Is Nothing Then
                service = CreateExchangeService()
                ServicePointManager.ServerCertificateValidationCallback = AddressOf CertificateValidationCallBack
                service.UseDefaultCredentials = True
                If _autoDiscoverServerByEMailAddress <> Nothing Then
                    service.AutodiscoverUrl(_autoDiscoverServerByEMailAddress)
                Else
                    service.Url = New Uri("https://" & _serverName & "/ews/exchange.asmx")
                End If
            End If
            Return service
        End Function

        ''' <summary>
        ''' Ignore any certificate errors
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="certificate"></param>
        ''' <param name="chain"></param>
        ''' <param name="sslPolicyErrors"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Function CertificateValidationCallBack(ByVal sender As Object, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
            Return True
            ''SECURITY NOTE BY MICROSOFT: The certificate validation callback method in this example provides sufficient security for development and testing of EWS Managed API applications. However, it may not provide sufficient security for your deployed application. You should always make sure that the certificate validation callback method that you use meets the security requirements of your organization.  
            'Requires Imports System.Net.Security
            'Requires Imports System.Security.Cryptography.X509Certificates
            ''If the certificate is a valid, signed certificate, return true.
            'If (sslPolicyErrors = System.Net.Security.SslPolicyErrors.None) Then
            '    Return True
            'End If

            ''If there are errors in the certificate chain, look at each error to determine the cause.
            'If ((sslPolicyErrors And System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) <> Net.Security.SslPolicyErrors.None) Then
            '    If (Not chain Is Nothing AndAlso Not chain.ChainStatus Is Nothing) Then
            '        For Each status As System.Security.Cryptography.X509Certificates.X509ChainStatus In chain.ChainStatus
            '            If ((certificate.Subject = certificate.Issuer) AndAlso (status.Status = System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot)) Then
            '                'Self-signed certificates with an untrusted root are valid. 
            '                Continue For
            '            Else
            '                If (status.Status <> System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError) Then
            '                    'If there are any other errors in the certificate chain, the certificate is invalid,
            '                    'so the method returns false.
            '                    Return False
            '                End If
            '            End If
            '        Next
            '    End If

            '    'When processing reaches this line, the only errors in the certificate chain are 
            '    'untrusted root errors for self-signed certificates. These certificates are valid
            '    'for default Exchange server installations, so return true.
            '    Return True
            'Else

            '    'In all other cases, return false.
            '    Return False
            'End If
        End Function
#End Region

        ''' <summary>
        ''' The list of well known folder names in Exchange
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum WellKnownFolderName As Integer
            ArchiveDeletedItems = 22
            ArchiveMsgFolderRoot = 21
            ArchiveRecoverableItemsDeletions = 24
            ArchiveRecoverableItemsPurges = 26
            ArchiveRecoverableItemsRoot = 23
            ArchiveRecoverableItemsVersions = 25
            ArchiveRoot = 20
            Calendar = 0
            Contacts = 1
            DeletedItems = 2
            Drafts = 3
            Inbox = 4
            Journal = 5
            JunkEmail = 13
            MsgFolderRoot = 10
            Notes = 6
            Outbox = 7
            PublicFoldersRoot = 11
            RecoverableItemsDeletions = 17
            RecoverableItemsPurges = 19
            RecoverableItemsRoot = 16
            RecoverableItemsVersions = 18
            Root = 12
            SearchFolders = 14
            SentItems = 8
            Tasks = 9
            VoiceMail = 15
        End Enum

        ''' <summary>
        ''' Send an e-mail message
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="bodyPlainText"></param>
        ''' <param name="recipientsTo"></param>
        ''' <param name="recipientsCc"></param>
        ''' <param name="recipientsBcc"></param>
        ''' <remarks></remarks>
        Public Sub SendMail(ByVal subject As String, ByVal bodyPlainText As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient())
            CreateMessage(subject, bodyPlainText, String.Empty, recipientsTo, recipientsCc, recipientsBcc, Nothing).SendAndSaveCopy()
        End Sub

        ''' <summary>
        ''' Send an e-mail message
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="bodyPlainText"></param>
        ''' <param name="bodyHtml"></param>
        ''' <param name="recipientsTo"></param>
        ''' <param name="recipientsCc"></param>
        ''' <param name="recipientsBcc"></param>
        ''' <remarks></remarks>
        Public Sub SendMail(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient())
            CreateMessage(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, Nothing).SendAndSaveCopy()
        End Sub

        ''' <summary>
        ''' Send an e-mail message with attachment
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="bodyPlainText"></param>
        ''' <param name="bodyHtml"></param>
        ''' <param name="recipientsTo"></param>
        ''' <param name="recipientsCc"></param>
        ''' <param name="recipientsBcc"></param>
        ''' <param name="attachment"></param>
        ''' <remarks></remarks>
        Public Sub SendMail(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient(), ByVal attachment() As EMailAttachment)
            CreateMessage(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, attachment).SendAndSaveCopy()
        End Sub

        ''' <summary>
        ''' Create a new e-mail message
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="bodyPlainText"></param>
        ''' <param name="bodyHtml"></param>
        ''' <param name="recipientsTo"></param>
        ''' <param name="recipientsCc"></param>
        ''' <param name="recipientsBcc"></param>
        ''' <param name="attachment"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CreateMessage(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient(), ByVal attachment() As EMailAttachment) As EmailMessage
            If bodyPlainText = Nothing And bodyHtml = Nothing Then
                Throw New ArgumentNullException("plain text or html body required")
            ElseIf Not (bodyPlainText = Nothing Xor bodyHtml = Nothing) Then
                Throw New ArgumentException("either plain text or html body required, but not both")
            End If
            If recipientsTo Is Nothing Then recipientsTo = New Recipient() {}
            If recipientsCc Is Nothing Then recipientsCc = New Recipient() {}
            If recipientsBcc Is Nothing Then recipientsBcc = New Recipient() {}
            'Create the e-mail message, set its properties, and send it to user2@contoso.com, saving a copy to the Sent Items folder. 
            Dim message As New EmailMessage(Me.CreateConfiguredExchangeService())
            message.Subject = subject
            If bodyHtml <> Nothing Then
                message.Body = New MessageBody(BodyType.HTML, bodyHtml)
            Else
                message.Body = New MessageBody(BodyType.Text, bodyPlainText)
            End If
            For Each recipient As Recipient In recipientsTo
                If recipient.Name = Nothing Then
                    message.ToRecipients.Add(recipient.EMailAddress)
                Else
                    message.ToRecipients.Add(recipient.Name, recipient.EMailAddress)
                End If
            Next
            For Each recipient As Recipient In recipientsCc
                If recipient.Name = Nothing Then
                    message.CcRecipients.Add(recipient.EMailAddress)
                Else
                    message.CcRecipients.Add(recipient.Name, recipient.EMailAddress)
                End If
            Next
            For Each recipient As Recipient In recipientsBcc
                If recipient.Name = Nothing Then
                    message.BccRecipients.Add(recipient.EMailAddress)
                Else
                    message.BccRecipients.Add(recipient.Name, recipient.EMailAddress)
                End If
            Next
            If Not attachment Is Nothing Then
                For Each Item As EMailAttachment In attachment
                    If Not Item Is Nothing Then
                        If Item.FilePath <> "" AndAlso Item.FileName = Nothing Then
                            message.Attachments.AddFileAttachment(Item.FilePath)
                        ElseIf Item.FileName <> "" AndAlso Not Item.FileData Is Nothing Then
                            message.Attachments.AddFileAttachment(Item.FileName, Item.FileData)
                        ElseIf Item.FilePath <> "" AndAlso Item.FileName <> "" Then
                            message.Attachments.AddFileAttachment(Item.FileName, Item.FilePath)
                        ElseIf Item.FileName <> "" AndAlso Not Item.FileStream Is Nothing Then
                            message.Attachments.AddFileAttachment(Item.FileName, Item.FileStream)
                        End If
                    End If
                Next
            End If
            Return message
        End Function

        ''' <summary>
        ''' Save an e-mail message as draft in the drafts folder
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="bodyPlainText"></param>
        ''' <param name="bodyHtml"></param>
        ''' <param name="recipientsTo"></param>
        ''' <param name="recipientsCc"></param>
        ''' <param name="recipientsBcc"></param>
        ''' <remarks></remarks>
        Public Function SaveMailAsDraft(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient()) As Uri
            Return SaveMailAsDraft(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, Nothing, Nothing)
        End Function

        ''' <summary>
        ''' Save an e-mail message with attachments as draft in the drafts folder
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="bodyPlainText"></param>
        ''' <param name="bodyHtml"></param>
        ''' <param name="recipientsTo"></param>
        ''' <param name="recipientsCc"></param>
        ''' <param name="recipientsBcc"></param>
        ''' <param name="attachment"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SaveMailAsDraft(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient(), ByVal attachment() As EMailAttachment) As Uri
            Return SaveMailAsDraft(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, Nothing, attachment)
        End Function

        ''' <summary>
        ''' Save an e-mail message as a draft
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="bodyPlainText"></param>
        ''' <param name="bodyHtml"></param>
        ''' <param name="recipientsTo"></param>
        ''' <param name="recipientsCc"></param>
        ''' <param name="recipientsBcc"></param>
        ''' <param name="folder"></param>
        ''' <remarks></remarks>
        Public Function SaveMailAsDraft(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient(), ByVal folder As CompuMaster.Data.MsExchange.FolderPathRepresentation, ByVal attatchment() As EMailAttachment) As Uri
            Dim message As EmailMessage = CreateMessage(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, attatchment)
            If folder Is Nothing Then
                message.Save()
            Else
                message.Save(folder.FolderID)
            End If
            If Me._exchangeVersion = ExchangeVersion.Exchange2007_SP1 Then
                Return Nothing 'Not supported for exchange 2007
            Else
                'exchange 2010 supports the lookup of a web client url
                Dim url As String = message.WebClientEditFormQueryString
                Dim uri As New Uri(url)
                Throw New NotImplementedException("URL lookup of mail still to be implemented")
                Return uri
            End If
        End Function

        ''' <summary>
        ''' Create an appointment
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="location"></param>
        ''' <param name="body"></param>
        ''' <param name="start"></param>
        ''' <param name="duration"></param>
        ''' <returns>The unique ID of the appointment for later reference</returns>
        ''' <remarks></remarks>
        Public Function CreateAppointment(ByVal subject As String, ByVal location As String, ByVal body As String, ByVal start As DateTime, ByVal duration As TimeSpan) As String
            Return CreateMeetingAppointment(subject, location, body, start, duration, New Recipient() {}, New Recipient() {}, New Recipient() {})
        End Function

        ''' <summary>
        ''' Create a meeting appointment
        ''' </summary>
        ''' <param name="subject"></param>
        ''' <param name="location"></param>
        ''' <param name="body"></param>
        ''' <param name="start"></param>
        ''' <param name="duration"></param>
        ''' <param name="requiredAttendees"></param>
        ''' <param name="optionalAttendees"></param>
        ''' <param name="resources"></param>
        ''' <returns>The unique ID of the appointment for later reference</returns>
        ''' <remarks></remarks>
        Public Function CreateMeetingAppointment(ByVal subject As String, ByVal location As String, ByVal body As String, ByVal start As DateTime, ByVal duration As TimeSpan, ByVal requiredAttendees As Recipient(), ByVal optionalAttendees As Recipient(), ByVal resources As Recipient()) As String
            If start = Nothing Then Throw New ArgumentNullException("start")
            If requiredAttendees Is Nothing Then requiredAttendees = New Recipient() {}
            If optionalAttendees Is Nothing Then optionalAttendees = New Recipient() {}
            If resources Is Nothing Then resources = New Recipient() {}
            Dim appointment As New Appointment(Me.CreateConfiguredExchangeService())
            appointment.Subject = subject
            appointment.Body = body
            appointment.Location = location
            appointment.Start = start
            appointment.End = appointment.Start.Add(duration)
            For Each Attendee As Recipient In requiredAttendees
                If Attendee.Name = Nothing Then
                    appointment.RequiredAttendees.Add(Attendee.EMailAddress)
                Else
                    appointment.RequiredAttendees.Add(Attendee.Name, Attendee.EMailAddress)
                End If
            Next
            For Each Attendee As Recipient In optionalAttendees
                If Attendee.Name = Nothing Then
                    appointment.OptionalAttendees.Add(Attendee.EMailAddress)
                Else
                    appointment.OptionalAttendees.Add(Attendee.Name, Attendee.EMailAddress)
                End If
            Next
            For Each Attendee As Recipient In resources
                If Attendee.Name = Nothing Then
                    appointment.Resources.Add(Attendee.EMailAddress)
                Else
                    appointment.Resources.Add(Attendee.Name, Attendee.EMailAddress)
                End If
            Next
            If requiredAttendees.Length = 0 AndAlso optionalAttendees.Length = 0 AndAlso resources.Length = 0 Then
                appointment.Save(SendInvitationsMode.SendToNone)
            Else
                appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy)
            End If
            Return appointment.Id.UniqueId
        End Function

        ''' <summary>
        ''' Save a contact
        ''' </summary>
        ''' <param name="contact"></param>
        ''' <param name="folder"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SaveContact(ByVal contact As Contact, ByVal folder As FolderPathRepresentation) As String
            If contact.IsNew AndAlso Not folder Is Nothing Then
                'Create new entry in specified folder
                contact.Save(folder.FolderID)
            ElseIf contact.IsNew AndAlso folder Is Nothing Then
                'Create new entry in default folder
                contact.Save()
            ElseIf folder Is Nothing Then
                'Overwrite existing item
                contact.Update(ConflictResolutionMode.AutoResolve)
            ElseIf Not folder Is Nothing AndAlso contact.ParentFolderId.UniqueId <> folder.FolderID Then
                'Save additional item in different folder instead of overwriting existing item
                contact.Save(folder.FolderID)
            Else
                'Overwrite existing item
                contact.Update(ConflictResolutionMode.AutoResolve)
            End If
            Return contact.Id.UniqueId
        End Function

        ''' <summary>
        ''' Create a new contact
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CreateNewContact() As Contact
            Return New Contact(Me.CreateConfiguredExchangeService())
        End Function

        ''' <summary>
        ''' Load a contact
        ''' </summary>
        ''' <param name="itemID"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LoadContactData(ByVal itemID As ItemId) As Contact
            Return Contact.Bind(Me.CreateConfiguredExchangeService, itemID)
        End Function

        ''' <summary>
        ''' Load a contact
        ''' </summary>
        ''' <param name="itemID"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LoadContactData(ByVal itemID As String) As Contact
            Return Me.LoadContactData(New Microsoft.Exchange.WebServices.Data.ItemId(itemID))
        End Function

        ''' <summary>
        ''' Well known folder classes
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum FolderClass As Integer
            Undefined = 0
            Generic = 1
            Contacts = 5
            Tasks = 2
            Search = 3
            Calendar = 4
            Notices = 6
            Journal = 7
            Configuration = 8
            Custom = 9
        End Enum

        ''' <summary>
        ''' Lookup the folder class
        ''' </summary>
        ''' <param name="folder"></param>
        ''' <returns>The well known folder class or otherwise Custom for a custom FolderClass name at Exchange</returns>
        ''' <remarks></remarks>
        <Obsolete("Better use Directory class", False)> Public Function LookupFolderClass(ByVal folder As FolderPathRepresentation) As Exchange2007SP1OrHigher.FolderClass
            Select Case LookupFolderClassName(folder)
                Case "IPF.Appointment"
                    Return Exchange2007SP1OrHigher.FolderClass.Calendar
                Case "IPF.Contact"
                    Return Exchange2007SP1OrHigher.FolderClass.Contacts
                Case "IPF.Note"
                    Return Exchange2007SP1OrHigher.FolderClass.Generic
                Case "IPF.Journal"
                    Return Exchange2007SP1OrHigher.FolderClass.Journal
                    'Case ""
                    '    Return Exchange2007SP1OrHigher.FolderClass.Search
                Case "IPF.Task"
                    Return Exchange2007SP1OrHigher.FolderClass.Tasks
                Case "IPF.StickyNote"
                    Return Exchange2007SP1OrHigher.FolderClass.Notices
                Case "IPF.Configuration"
                    Return Exchange2007SP1OrHigher.FolderClass.Configuration
                Case Else
                    Return FolderClass.Custom
            End Select
        End Function

        ''' <summary>
        ''' Lookup the folder class
        ''' </summary>
        ''' <param name="folder"></param>
        ''' <returns>The real folder class name as defined at Exchange</returns>
        ''' <remarks></remarks>
        <Obsolete("Better use Directory class", False)> Public Function LookupFolderClassName(ByVal folder As FolderPathRepresentation) As String
            Dim lookupfolder As Microsoft.Exchange.WebServices.Data.Folder
            lookupfolder = Microsoft.Exchange.WebServices.Data.Folder.Bind(Me.CreateConfiguredExchangeService, folder.FolderID)
            Return lookupfolder.FolderClass
        End Function

        ''' <summary>
        ''' Convert the well known folder class into the official folder class name as Exchange
        ''' </summary>
        ''' <param name="folderClass"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FolderClassName(ByVal folderClass As FolderClass) As String
            Select Case folderClass
                Case Exchange2007SP1OrHigher.FolderClass.Calendar
                    Return "IPF.Appointment"
                Case Exchange2007SP1OrHigher.FolderClass.Contacts
                    Return "IPF.Contact"
                Case Exchange2007SP1OrHigher.FolderClass.Generic
                    Return "IPF.Note"
                Case Exchange2007SP1OrHigher.FolderClass.Journal
                    Return "IPF.Journal"
                Case Exchange2007SP1OrHigher.FolderClass.Search
                    Throw New NotSupportedException("Search classes use custom names, name resolution not supported")
                Case Exchange2007SP1OrHigher.FolderClass.Tasks
                    Return "IPF.Task"
                Case Exchange2007SP1OrHigher.FolderClass.Notices
                    Return "IPF.StickyNote"
                Case Exchange2007SP1OrHigher.FolderClass.Configuration
                    Return "IPF.Configuration"
                Case Exchange2007SP1OrHigher.FolderClass.Custom
                    Throw New NotSupportedException("A custom folder class requires you to use custom folder class name. The purpose of this method is not intended for this folder class type.")
                Case Else
                    Throw New ArgumentOutOfRangeException("folderClass")
            End Select
        End Function

        ''' <summary>
        ''' Create a new subfolder
        ''' </summary>
        ''' <param name="folderName"></param>
        ''' <param name="baseFolder"></param>
        ''' <returns>The unique ID of the folder</returns>
        ''' <remarks></remarks>
        <Obsolete("Better use Directory class", False)> Public Function CreateFolder(ByVal folderName As String, ByVal baseFolder As FolderPathRepresentation) As String
            Return CreateFolder(folderName, baseFolder, FolderClass.Generic)
        End Function

        ''' <summary>
        ''' Create a new subfolder
        ''' </summary>
        ''' <param name="folderName"></param>
        ''' <param name="baseFolder"></param>
        ''' <returns>The unique ID of the folder</returns>
        ''' <remarks></remarks>
        <Obsolete("Better use Directory class", False)> Public Function CreateFolder(ByVal folderName As String, ByVal baseFolder As FolderPathRepresentation, ByVal folderClass As FolderClass) As String
            Dim customFolderClass As String = Nothing
            If folderClass <> Exchange2007SP1OrHigher.FolderClass.Undefined Then
                customFolderClass = FolderClassName(folderClass)
            End If
            Return CreateFolder(folderName, baseFolder, customFolderClass)
        End Function

        ''' <summary>
        ''' Create a new subfolder
        ''' </summary>
        ''' <param name="folderName"></param>
        ''' <param name="baseFolder"></param>
        ''' <returns>The unique ID of the folder</returns>
        ''' <remarks></remarks>
        <Obsolete("Better use Directory class", False)> Public Function CreateFolder(ByVal folderName As String, ByVal baseFolder As FolderPathRepresentation, ByVal customFolderClass As String) As String
            Dim folder As New Microsoft.Exchange.WebServices.Data.Folder(Me.CreateConfiguredExchangeService)
            folder.DisplayName = folderName
            If customFolderClass <> Nothing Then
                folder.FolderClass = customFolderClass
            End If
            folder.Save(baseFolder.FolderID)
            Return folder.Id.UniqueId
        End Function

        ''' <summary>
        ''' Delete a folder
        ''' </summary>
        ''' <param name="folder"></param>
        ''' <param name="deletionMode"></param>
        ''' <remarks></remarks>
        <Obsolete("Better use Directory class", False)> Public Sub DeleteFolder(ByVal folder As FolderPathRepresentation, ByVal deletionMode As DeleteMode)
            Dim dropFolder As Folder = Microsoft.Exchange.WebServices.Data.Folder.Bind(Me.CreateConfiguredExchangeService, folder.FolderID)
            dropFolder.Delete(deletionMode)
        End Sub

        ''' <summary>
        ''' Empty a folder
        ''' </summary>
        ''' <param name="folder"></param>
        ''' <param name="deletionMode"></param>
        ''' <remarks></remarks>
        <Obsolete("Better use Directory class", False)> Public Sub EmptyFolder(ByVal folder As FolderPathRepresentation, ByVal deletionMode As DeleteMode, ByVal deleteSubFolders As Boolean)
            Dim emptyFolder As Folder = Microsoft.Exchange.WebServices.Data.Folder.Bind(Me.CreateConfiguredExchangeService, folder.FolderID)
            emptyFolder.Empty(deletionMode, deleteSubFolders)
        End Sub

        ''' <summary>
        ''' Lookup a folder path representation based on its directory structure
        ''' </summary>
        ''' <param name="baseFolder"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LookupFolder(ByVal baseFolder As WellKnownFolderName) As FolderPathRepresentation
            Dim folder As Folder = Microsoft.Exchange.WebServices.Data.Folder.Bind(Me.CreateConfiguredExchangeService, CType(baseFolder, Microsoft.Exchange.WebServices.Data.WellKnownFolderName))
            Return New FolderPathRepresentation(Me, folder)
        End Function

        Private _DirectorySeparatorChar As Char = "\"c
        ''' <summary>
        ''' The directory separator char which shall be used in all method calls to this class with parameters specifying a directory structure
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>The exchange server supports typical directory separator chars like \, / as a normal character within a directory name. That's why the programmer may want to define his own separator char to support required directories correctly.</remarks>
        Public Property DirectorySeparatorChar() As Char
            Get
                Return _DirectorySeparatorChar
            End Get
            Set(ByVal value As Char)
                _DirectorySeparatorChar = value
            End Set
        End Property

        ''' <summary>
        ''' Enumerates possible matches of mailbox accounts/contacts for the searched name
        ''' </summary>
        ''' <param name="searchedName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ResolveMailboxOrContactNames(ByVal searchedName As String) As Mailbox()
            'Identify the mailbox folders to search for potential name resolution matches.
            Dim folders As List(Of FolderId) = New List(Of FolderId)
            folders.Add(New FolderId(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Contacts))

            'Search for all contact entries in the default mailbox contacts folder and in Active Directory. This results in a call to EWS.
            Dim coll As NameResolutionCollection = Me.CreateConfiguredExchangeService.ResolveName(searchedName, folders, ResolveNameSearchLocation.ContactsThenDirectory, False)

            Dim Results As New ArrayList
            For Each nameRes As NameResolution In coll
                Results.Add(nameRes.Mailbox)
                Console.WriteLine("Contact name: " + nameRes.Mailbox.Name)
                Console.WriteLine("Contact e-mail address: " + nameRes.Mailbox.Address)
                Console.WriteLine("Mailbox type: " + nameRes.Mailbox.MailboxType.ToString)
            Next
            Return CType(Results.ToArray(GetType(Mailbox)), Mailbox())
        End Function

    End Class

End Namespace