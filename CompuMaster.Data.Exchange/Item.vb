﻿Option Strict On
Option Explicit On

Imports Microsoft.Exchange.WebServices.Data
Imports System.Net

Namespace CompuMaster.Data.MsExchange

    Public Class Item
        Private ReadOnly _parentDirectory As Directory
        Private ReadOnly _exchangeItem As Microsoft.Exchange.WebServices.Data.Item
        Private ReadOnly _service As Exchange2007SP1OrHigher
        Public Sub New(service As Exchange2007SP1OrHigher, item As Microsoft.Exchange.WebServices.Data.Item)
            _exchangeItem = item
            _service = service
        End Sub
        Public Sub New(service As Exchange2007SP1OrHigher, item As Microsoft.Exchange.WebServices.Data.Item, parentDirectory As Directory)
            _parentDirectory = parentDirectory
            _exchangeItem = item
            _service = service
        End Sub

        Public ReadOnly Property ParentDirectory As Directory
            Get
                If _exchangeItem.ParentFolderId.UniqueId = _parentDirectory.ID Then
                    'pointer to parent directory is really the parent directory
                    Return _parentDirectory
                Else
                    'pointer to parent directory is not the real directory, but the requested directory might be a child directory of it
                    For MyCounter As Integer = 0 To _parentDirectory.SubFoldersOfSeveralHierachyLevels.Count - 1
                        If _exchangeItem.ParentFolderId.UniqueId = _parentDirectory.SubFoldersOfSeveralHierachyLevels(MyCounter).ID Then
                            Return _parentDirectory.SubFoldersOfSeveralHierachyLevels(MyCounter)
                        End If
                    Next
                    'pointer to parent directory is not the real directory, and no matching child directory found
                    Throw New InvalidOperationException("item's parent directory information doesn't match to the referenced parent directory")
                End If
            End Get
        End Property

        Public ReadOnly Property ExchangeItem As Microsoft.Exchange.WebServices.Data.Item
            Get
                Return _exchangeItem
            End Get
        End Property

        Public ReadOnly Property Subject As String
            Get
                Return _exchangeItem.Subject
            End Get
        End Property

        Public ReadOnly Property DateTimeSent As DateTime
            Get
                If Me.IsDraft Then
                    Return Nothing
                Else
                    Try
                        Return _exchangeItem.DateTimeSent
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        Return Nothing
                    End Try
                End If
            End Get
        End Property
        Public ReadOnly Property DateTimeReceived As DateTime
            Get
                Try
                    Return _exchangeItem.DateTimeReceived
                Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                    Return Nothing
                End Try
            End Get
        End Property
        Public ReadOnly Property DateTimeCreated As DateTime
            Get
                Try
                    Return _exchangeItem.DateTimeCreated
                Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                    Return Nothing
                End Try
            End Get
        End Property

        Public ReadOnly Property CalendarEntryBegin As DateTime
            Get
                If Me.IsAppointment = False OrElse ExtendedData.ContainsKey("Start") = False OrElse ExtendedData.Item("Start") Is Nothing Then
                    Return Nothing
                Else
                    Return CType(ExtendedData.Item("Start"), DateTime)
                End If
            End Get
        End Property
        Public ReadOnly Property CalendarEntryEnd As DateTime
            Get
                If Me.IsAppointment = False OrElse ExtendedData.ContainsKey("End") = False OrElse ExtendedData.Item("End") Is Nothing Then
                    Return Nothing
                Else
                    Return CType(ExtendedData.Item("End"), DateTime)
                End If
            End Get
        End Property
        Public ReadOnly Property IsAppointment As Boolean
            Get
                Try
                    If _exchangeItem.ItemClass = "IPM.Appointment" Then
                        Return True
                    Else
                        Return False
                    End If
                Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                    Return False
                End Try
            End Get
        End Property

        Public ReadOnly Property IsDraft As Boolean
            Get
                Return _exchangeItem.IsDraft
            End Get
        End Property

        'Public ReadOnly Property MimeContent As String
        '    Get
        '        Try
        '            Return System.Text.Encoding.GetEncoding(_exchangeItem.MimeContent.CharacterSet).GetString(_exchangeItem.MimeContent.Content)
        '        Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
        '            Return Nothing
        '        End Try
        '    End Get
        'End Property
        'Public ReadOnly Property BodyType As String
        '    Get
        '        Try
        '            Return _exchangeItem.Body.BodyType.ToString
        '        Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
        '            Return Nothing
        '        End Try
        '    End Get
        'End Property
        'Public ReadOnly Property Body As String
        '    Get
        '        Try
        '            Return _exchangeItem.Body
        '        Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
        '            Return Nothing
        '        End Try
        '    End Get
        'End Property
        Public ReadOnly Property BodyText As String
            Get
                Static _Result As String = Nothing
                If _Result Is Nothing Then
                    Try
                        Dim message As Microsoft.Exchange.WebServices.Data.EmailMessage = SenderRecipientsDataAndPlainTextBody()
                        _Result = message.Body.Text
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        _Result = ""
                    End Try
                End If
                Return _Result
            End Get
        End Property

        Public ReadOnly Property BodyHtml As String
            Get
                Static _Result As String = Nothing
                If _Result Is Nothing Then
                    Try
                        Dim propSet As New Microsoft.Exchange.WebServices.Data.PropertySet(Microsoft.Exchange.WebServices.Data.BasePropertySet.IdOnly, Microsoft.Exchange.WebServices.Data.EmailMessageSchema.Body) With {
                            .RequestedBodyType = Microsoft.Exchange.WebServices.Data.BodyType.HTML
                        }
                        Dim message As Microsoft.Exchange.WebServices.Data.EmailMessage = Microsoft.Exchange.WebServices.Data.EmailMessage.Bind(_service.CreateConfiguredExchangeService, _exchangeItem.Id, propSet).Result
                        _Result = message.Body.Text
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        _Result = ""
                    End Try
                End If
                Return _Result
            End Get
        End Property

        Private Function SenderRecipientsDataAndPlainTextBody() As Microsoft.Exchange.WebServices.Data.EmailMessage
            Static _Result As Microsoft.Exchange.WebServices.Data.EmailMessage = Nothing
            If _Result Is Nothing Then
                Dim AdditionalProperties As New List(Of Microsoft.Exchange.WebServices.Data.PropertyDefinition) From {
                    Microsoft.Exchange.WebServices.Data.EmailMessageSchema.From,
                    Microsoft.Exchange.WebServices.Data.EmailMessageSchema.DisplayTo,
                    Microsoft.Exchange.WebServices.Data.EmailMessageSchema.DisplayCc,
                    Microsoft.Exchange.WebServices.Data.EmailMessageSchema.ToRecipients,
                    Microsoft.Exchange.WebServices.Data.EmailMessageSchema.CcRecipients,
                    Microsoft.Exchange.WebServices.Data.EmailMessageSchema.BccRecipients,
                    Microsoft.Exchange.WebServices.Data.EmailMessageSchema.ReplyTo,
                    Microsoft.Exchange.WebServices.Data.EmailMessageSchema.Body,
                    Microsoft.Exchange.WebServices.Data.AppointmentSchema.Start,
                    Microsoft.Exchange.WebServices.Data.AppointmentSchema.StartTimeZone,
                    Microsoft.Exchange.WebServices.Data.AppointmentSchema.End,
                    Microsoft.Exchange.WebServices.Data.AppointmentSchema.EndTimeZone
                }
                Dim propSet As New Microsoft.Exchange.WebServices.Data.PropertySet(Microsoft.Exchange.WebServices.Data.BasePropertySet.FirstClassProperties, AdditionalProperties.ToArray) With {
                    .RequestedBodyType = Microsoft.Exchange.WebServices.Data.BodyType.Text
                }
                _Result = Microsoft.Exchange.WebServices.Data.EmailMessage.Bind(_service.CreateConfiguredExchangeService, _exchangeItem.Id, propSet).Result
            End If
            Return _Result
        End Function

        Public ReadOnly Property FromSender As System.Net.Mail.MailAddress
            Get
                Static _Result As System.Net.Mail.MailAddress = Nothing
                If False And _Result Is Nothing Then
                    Try
                        Dim message As Microsoft.Exchange.WebServices.Data.EmailMessage = SenderRecipientsDataAndPlainTextBody()
                        If message.From Is Nothing Then
                            _Result = Nothing
                        Else
                            _Result = New System.Net.Mail.MailAddress(message.From.Address, message.From.Name)
                        End If
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        _Result = Nothing
                    End Try
                End If
                Return _Result
            End Get
        End Property

        Public ReadOnly Property ReplyTo As List(Of System.Net.Mail.MailAddress)
            Get
                Static _Result As List(Of System.Net.Mail.MailAddress) = Nothing
                If _Result Is Nothing Then
                    Try
                        Dim message As Microsoft.Exchange.WebServices.Data.EmailMessage = SenderRecipientsDataAndPlainTextBody()
                        _Result = New List(Of System.Net.Mail.MailAddress)
                        For Each addr As EmailAddress In message.ReplyTo
                            _Result.Add(New System.Net.Mail.MailAddress(addr.Address, addr.Name))
                        Next
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        _Result = Nothing
                    End Try
                End If
                Return _Result
            End Get
        End Property

        Public ReadOnly Property FromExchangeSender As String
            Get
                If ExtendedData.ContainsKey("From") = False OrElse ExtendedData.Item("From") Is Nothing Then
                    Return String.Empty
                Else
                    Return CType(ExtendedData.Item("From"), Microsoft.Exchange.WebServices.Data.EmailAddress).ToString
                End If
            End Get
        End Property

        Public ReadOnly Property DisplayTo As String
            Get
                Return CType(ExtendedData.Item("DisplayTo"), String)
            End Get
        End Property

        Public ReadOnly Property DisplayCc As String
            Get
                Return CType(ExtendedData.Item("DisplayCc"), String)
            End Get
        End Property

        Public ReadOnly Property RecipientTo As List(Of System.Net.Mail.MailAddress)
            Get
                Static _Result As List(Of System.Net.Mail.MailAddress) = Nothing
                If _Result Is Nothing Then
                    Try
                        Dim message As Microsoft.Exchange.WebServices.Data.EmailMessage = SenderRecipientsDataAndPlainTextBody()
                        _Result = New List(Of System.Net.Mail.MailAddress)
                        For Each addr As EmailAddress In message.ToRecipients
                            _Result.Add(New System.Net.Mail.MailAddress(addr.Address, addr.Name))
                        Next
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        _Result = Nothing
                    End Try
                End If
                Return _Result
            End Get
        End Property

        Public ReadOnly Property RecipientCc As List(Of System.Net.Mail.MailAddress)
            Get
                Static _Result As List(Of System.Net.Mail.MailAddress) = Nothing
                If _Result Is Nothing Then
                    Try
                        Dim message As Microsoft.Exchange.WebServices.Data.EmailMessage = SenderRecipientsDataAndPlainTextBody()
                        _Result = New List(Of System.Net.Mail.MailAddress)
                        For Each addr As EmailAddress In message.CcRecipients
                            _Result.Add(New System.Net.Mail.MailAddress(addr.Address, addr.Name))
                        Next
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        _Result = Nothing
                    End Try
                End If
                Return _Result
            End Get
        End Property

        Public ReadOnly Property RecipientBcc As List(Of System.Net.Mail.MailAddress)
            Get
                Static _Result As List(Of System.Net.Mail.MailAddress) = Nothing
                If _Result Is Nothing Then
                    Try
                        Dim message As Microsoft.Exchange.WebServices.Data.EmailMessage = SenderRecipientsDataAndPlainTextBody()
                        _Result = New List(Of System.Net.Mail.MailAddress)
                        For Each addr As EmailAddress In message.BccRecipients
                            _Result.Add(New System.Net.Mail.MailAddress(addr.Address, addr.Name))
                        Next
                    Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                        _Result = Nothing
                    End Try
                End If
                Return _Result
            End Get
        End Property

        Private _ExtendedData As Generic.Dictionary(Of String, Object)
        Public Function ExtendedData() As Generic.Dictionary(Of String, Object)
            If _ExtendedData Is Nothing Then
                'Load first class props
                'Dim propSet As New Microsoft.Exchange.WebServices.Data.PropertySet(Microsoft.Exchange.WebServices.Data.BasePropertySet.FirstClassProperties)
                'Microsoft.Exchange.WebServices.Data.EmailMessage.Bind(_service.CreateConfiguredExchangeService, _exchangeItem.Id, propSet)
                _ExtendedData = New Generic.Dictionary(Of String, Object)

                'Add all items into the result table with all of their properties as complete as possible
                'Add required additional columns if not yet done
                For Each prop As PropertyDefinition In Me.ExchangeItem.Schema
                    Dim ColName As String = prop.Name
                    If prop.Version <> 0 Then
                        ColName &= "_V" & prop.Version
                    End If
                    If prop.Version > Me._service.ExchangeServiceVersion Then
                        'service version cannot read from fields with version number of a higher/newer exchange server version (e.g. using service version for Exchange 2010, but reading an Exchange 2013 field)
                        'causing Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyExceptions
                    Else
                        'read as usual
                        Try
                            If Me.ExchangeItem.Item(prop) Is Nothing Then
                                _ExtendedData.Add(ColName, Nothing)
                            Else
                                Select Case Me.ExchangeItem.Item(prop).GetType.ToString
                                    Case GetType(Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection
                                        value = CType(Me.ExchangeItem.Item(prop), Microsoft.Exchange.WebServices.Data.ExtendedPropertyCollection)
                                    Case GetType(Microsoft.Exchange.WebServices.Data.CompleteName).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.CompleteName
                                        value = CType(Me.ExchangeItem.Item(prop), Microsoft.Exchange.WebServices.Data.CompleteName)
                                        _ExtendedData.Add(ColName & "_Title", value.Title)
                                    Case GetType(Microsoft.Exchange.WebServices.Data.EmailAddressDictionary).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.EmailAddressDictionary
                                        value = CType(Me.ExchangeItem.Item(prop), Microsoft.Exchange.WebServices.Data.EmailAddressDictionary)
                                        If value.Contains(EmailAddressKey.EmailAddress1) Then _ExtendedData.Add(ColName & "_Email1", value(EmailAddressKey.EmailAddress1).Address)
                                        If value.Contains(EmailAddressKey.EmailAddress2) Then _ExtendedData.Add(ColName & "_Email2", value(EmailAddressKey.EmailAddress2).Address)
                                        If value.Contains(EmailAddressKey.EmailAddress3) Then _ExtendedData.Add(ColName & "_Email3", value(EmailAddressKey.EmailAddress3).Address)
                                    Case GetType(Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary).ToString
                                        Dim value As Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary
                                        value = CType(Me.ExchangeItem.Item(prop), Microsoft.Exchange.WebServices.Data.PhysicalAddressDictionary)
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
                                        value = CType(Me.ExchangeItem.Item(prop), Microsoft.Exchange.WebServices.Data.PhoneNumberDictionary)
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
                                        _ExtendedData.Add(ColName, Me.ExchangeItem.Item(prop))
                                End Select
                            End If
                        Catch ex As Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException
                            'Mark this column to be killed at the end because it only contains non-sense
                            'DEBUG NOTE: This exception might appear several times in debug sessions but can't be stopped from throwing -> JUST IGNORE THEM!
                        Catch ex As Microsoft.Exchange.WebServices.Data.ServiceVersionException
                            'Mark this column to be killed at the end because it only contains non-sense
                        Catch ex As NullReferenceException
                            _ExtendedData.Add(ColName, Nothing)
                        End Try
                    End If
                Next
            End If
            Return _ExtendedData
        End Function

    End Class

End Namespace