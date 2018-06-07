'=====================================================================

'  This file is part of the Autodesk Vault API Code Samples.

'  Copyright (C) Autodesk Inc.  All rights reserved.

'THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
'KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'PARTICULAR PURPOSE.
'=====================================================================

Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports System.Linq

'wB added
Imports System.IO


Imports Autodesk.Connectivity.WebServices
Imports Autodesk.Connectivity.WebServicesTools
Imports VDF = Autodesk.DataManagement.Client.Framework

'wB added
Imports ACW = Autodesk.Connectivity.WebServices

''' <summary>
''' This is our main form.
''' </summary>
Public Class Form1
    Inherits System.Windows.Forms.Form

    'wB moved here
    'Dim connection As VDF.Vault.Currency.Connections.Connection

    Public Sub New()
        '
        ' Required for Windows Form Designer support
        '
        InitializeComponent()

    End Sub

    ''' <summary>
    ''' The main entry point for the application.
    ''' </summary>
    <STAThread()> _
    Private Shared Sub Main()
        Dim appobject As New Form1()

        Application.Run(appobject)
    End Sub

    Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
        ListAllFiles()
    End Sub

    ''' <summary>
    ''' This function lists all the files in the Vault and displays them in the form's ListBox.
    ''' </summary>
    Public Sub ListAllFiles()
        ' For demonstration purposes, the information is hard-coded.
        Dim results As VDF.Vault.Results.LogInResult = VDF.Vault.Library.ConnectionManager.LogIn("localhost", "Vault", "Administrator", "", VDF.Vault.Currency.Connections.AuthenticationFlags.Standard, Nothing)

        If Not results.Success Then
            Return
        End If

        'wB commented
        Dim connection As VDF.Vault.Currency.Connections.Connection = results.Connection
        'wB added
        ' connection = results.Connection

        Try
            ' Start at the root Folder.
            Dim root As VDF.Vault.Currency.Entities.Folder = connection.FolderManager.RootFolder

            Me.m_listBox.Items.Clear()

            ' Call a function which prints all files in a Folder and sub-Folders.
            PrintFilesInFolder(root, connection)
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error")
            Return
        End Try

        VDF.Vault.Library.ConnectionManager.LogOut(connection)
    End Sub


    ''' <summary>
    ''' Prints all the files in the current Folder and any sub Folders.
    ''' </summary>
    ''' <param name="parentFolder">The folder we want to print.</param>
    ''' <param name="connection">The manager object for making Vault Server calls.</param>
    Private Sub PrintFilesInFolder(parentFolder As VDF.Vault.Currency.Entities.Folder, connection As VDF.Vault.Currency.Connections.Connection)
        ' get all the Files in the current Folder.
        Dim childFiles As ACW.File() = connection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(parentFolder.Id, False)

        'wB added to test
        Dim Defs As ACW.PropDef() = connection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")


        ' print out any Files we find.
        If childFiles IsNot Nothing AndAlso childFiles.Any() Then
            For Each file As ACW.File In childFiles
                ' print the full path, which is Folder name + File name.
                Me.m_listBox.Items.Add((Convert.ToString(parentFolder.FullName) & "/") + file.Name)
            Next
        End If

        ' check for any sub Folders.
        Dim folders As IEnumerable(Of VDF.Vault.Currency.Entities.Folder) = connection.FolderManager.GetChildFolders(parentFolder, False, False)
        If folders IsNot Nothing AndAlso folders.Any() Then
            For Each folder As VDF.Vault.Currency.Entities.Folder In folders
                ' recursively print the files in each sub-Folder
                PrintFilesInFolder(folder, connection)
            Next
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        ' For demonstration purposes, the information is hard-coded.
        Dim results As VDF.Vault.Results.LogInResult = VDF.Vault.Library.ConnectionManager.LogIn("localhost", "Vault", "Administrator", "", VDF.Vault.Currency.Connections.AuthenticationFlags.Standard, Nothing)

        If Not results.Success Then
            Return
        End If

        Dim connection As VDF.Vault.Currency.Connections.Connection = results.Connection

        ' Get the FileIteration 
        Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration = Nothing
        oFileIteration = getFileIteration("12345 Testing XXXX.csv", connection)

        Dim fldrPath As VDF.Currency.FolderPathAbsolute = New VDF.Currency.FolderPathAbsolute("c:\Temp")
        Dim filePath As String = Path.Combine(fldrPath.ToString(), oFileIteration.EntityName)


        If System.IO.File.Exists(filePath) Then
            'we'll try to delete the file so we can get the latest copy
            Try
                System.IO.File.Delete(filePath)
            Catch generatedExceptionName As System.IO.IOException
                VDF.Vault.Library.ConnectionManager.LogOut(connection)
                Throw New Exception("The file you are attempting to open already exists and can not be overwritten. This file may currently be open, try closing any application you are using to view this file and try opening the file again.")
            End Try
        End If

        ' Get settings 
        Dim oSettings As VDF.Vault.Settings.AcquireFilesSettings =
                                          New VDF.Vault.Settings.AcquireFilesSettings(connection)
        'Going to Checkout and download
        oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout Or _
                                    VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download

        oSettings.AddEntityToAcquire(oFileIteration)
        'set the path to the working folder
        oSettings.LocalPath = fldrPath 'connection.WorkingFoldersManager.GetWorkingFolder(oFileIteration)
        'Do the download
        connection.FileManager.AcquireFiles(oSettings)


        VDF.Vault.Library.ConnectionManager.LogOut(connection)
    End Sub

    'wB added this procedure, initial code from blog post
    ' Public Sub getFileToDownload(connection As VDF.Vault.Currency.Connections.Connection)
    Public Function getFileIteration(nameOfFile As String, connection As VDF.Vault.Currency.Connections.Connection) _
                                                                        As VDF.Vault.Currency.Entities.FileIteration
        'wB commented
        'Dim conditions As SrchCond()
        'wB added
        Dim conditions As ACW.SrchCond()

        ReDim conditions(0)

        Dim lCode As Long = 1
        'wB commented
        'Dim Defs As PropDef() = serviceManager.PropertyService.GetPropertyDefiniti?onsByEntityClassId("FILE")
        'wB added
        Dim Defs As ACW.PropDef() = connection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")

        'wB commented
        'Dim Prop As PropDef = Nothing
        'wB added
        Dim Prop As ACW.PropDef = Nothing
        'wB commented
        'For Each def As PropDef In Defs
        'wB added
        For Each def As ACW.PropDef In Defs
            If def.DispName = "File Name" Then
                Prop = def
            End If
        Next def
        'wB commented
        'Dim searchCondition As SrchCond = New SrchCond()
        'wB added
        Dim searchCondition As ACW.SrchCond = New ACW.SrchCond()
        searchCondition.PropDefId = Prop.Id
        'wB commented
        'searchCondition.PropTyp = PropertySearchType.SingleProperty
        'wB added
        searchCondition.PropTyp = ACW.PropertySearchType.SingleProperty
        searchCondition.SrchOper = lCode
        'wB commented
        ' searchCondition.SrchTxt = "12345 Testing XXXX.csv"
        'wB added
        searchCondition.SrchTxt = nameOfFile

        conditions(0) = searchCondition

        ' search for files
        Dim FileList As List(Of Autodesk.Connectivity.WebServices.File) = New List(Of Autodesk.Connectivity.WebServices.File)()
        Dim sBookmark As String = String.Empty
        'wB commented
        'Dim Status As SrchStatus = Nothing
        'wB added
        Dim Status As ACW.SrchStatus = Nothing

        While (Status Is Nothing OrElse FileList.Count < Status.TotalHits)
            'wB commented
            ' Dim files As Autodesk.Connectivity.WebServices.File() = serviceManager.DocumentService.FindFilesBySearchCo?nditions( _
            '    conditions, Nothing, Nothing, True, True, sBookmark, Status)
            Dim files As Autodesk.Connectivity.WebServices.File() = connection.WebServiceManager.
                                                      DocumentService.FindFilesBySearchConditions(conditions,
                                                             Nothing, Nothing, True, True, sBookmark, Status)

            'wB end added
            If (Not files Is Nothing) Then
                FileList.AddRange(files)
            End If
        End While

        'wB added
        Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration = New VDF.Vault.Currency.Entities.FileIteration(connection, FileList(0))
        Return oFileIteration

    End Function

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ' For demonstration purposes, the information is hard-coded.
        Dim results As VDF.Vault.Results.LogInResult = VDF.Vault.Library.ConnectionManager.LogIn("localhost", "Vault", "Administrator", "", VDF.Vault.Currency.Connections.AuthenticationFlags.Standard, Nothing)

        If Not results.Success Then
            Return
        End If

        Dim connection As VDF.Vault.Currency.Connections.Connection = results.Connection

        ' Get the FileIteration 
        Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration =
                                        getFileIteration("12345 Testing XXXX.csv", connection)

        'Dim fldrPath As Autodesk.DataManagement.Client.Framework.Currency.FolderPathAbsolute = connection.WorkingFoldersManager.GetWorkingFolder(oFileIteration)
        Dim fldrPath As VDF.Currency.FolderPathAbsolute = New VDF.Currency.FolderPathAbsolute("c:\Temp")
        Dim filePath As String = Path.Combine(fldrPath.ToString(), oFileIteration.EntityName)
        Dim filePathAbs As VDF.Currency.FilePathAbsolute = New VDF.Currency.FilePathAbsolute(filePath)

        'determine if the file exists
        If Not System.IO.File.Exists(filePath) Then
            MessageBox.Show("The file you are attempting to check in does not exist on disk.")
            'Logout
            VDF.Vault.Library.ConnectionManager.LogOut(connection)
            Return
        End If

        If oFileIteration.IsCheckedOut = True Then

            connection.FileManager.CheckinFile(oFileIteration, "wbTesting", False, New ACW.FileAssocParam() {},
                                                                             Nothing, False, Nothing,
                                                  ACW.FileClassification.None, False, filePathAbs)

        End If

        'Logout
        VDF.Vault.Library.ConnectionManager.LogOut(connection)
    End Sub


End Class
