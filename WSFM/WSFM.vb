Imports System.Diagnostics

Module WSFM

    Private Const _SPOTLIGHT_ROOT_FOLDER As String = "\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"
    Private Const _LANDSCAPE_ROOT_FOLDER As String = "\Wallpaper\Landscape"
    Private Const _PORTRAIT_ROOT_FOLDER As String = "\Wallpaper\Portrait"

    Private _portraitFolder, _landscapeFolder, _sourceFolder As String
    Private _numInvalidFiles, _numDuplicatedFiles, _numLandscapePictures, _numCopiedLandscapePictures, _numPortraitPictures, _numCopiedPortraitPictures, _numValidPictures As Integer
    Private eLog As New EventLog("Application")

#Region "  Properties  "

    Public Property SourceFolder As String
        Get
            Return _sourceFolder
        End Get
        Set(value As String)
            Try
                If System.IO.Directory.Exists(value) Then
                    _sourceFolder = value
                End If
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    Public Property PortraitFolder As String
        Get
            Return _portraitFolder
        End Get
        Set(value As String)
            Try
                If System.IO.Directory.Exists(value) Then
                    _portraitFolder = value
                End If
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    Public Property LandscapeFolder As String
        Get
            Return _landscapeFolder
        End Get
        Set(value As String)
            Try
                If System.IO.Directory.Exists(value) Then
                    _landscapeFolder = value

                End If
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    Private ReadOnly Property LandscapeRootFolder As String
        Get
            Return _LANDSCAPE_ROOT_FOLDER
        End Get
    End Property

    Private ReadOnly Property PortraitRootFolder As String
        Get
            Return _PORTRAIT_ROOT_FOLDER
        End Get
    End Property

    Public ReadOnly Property SpotlightRootFolder As String
        Get
            Return _SPOTLIGHT_ROOT_FOLDER
        End Get
    End Property

#End Region

    Public Sub Main()
        SourceFolder = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), SpotlightRootFolder)
        LandscapeFolder = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), LandscapeRootFolder)
        PortraitFolder = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), PortraitRootFolder)
        eLog.Source = "WSFM"

        StartCopyProcess()

    End Sub

    Private Sub StartCopyProcess()
        Dim sourceFiles As String() = System.IO.Directory.GetFiles(SourceFolder)

        _numInvalidFiles = 0
        _numDuplicatedFiles = 0
        _numLandscapePictures = 0
        _numCopiedLandscapePictures = 0
        _numPortraitPictures = 0
        _numCopiedPortraitPictures = 0
        _numValidPictures = 0

        For Each file As String In sourceFiles
            Try
                Dim fileInfo As New System.IO.FileInfo(file)
                Dim picture As New System.Drawing.Bitmap(file)

                If (picture.Width = 1920 Or picture.Width = 1080) And (picture.Height = 1080 Or picture.Height = 1920) Then
                    _numValidPictures += 1

                    If picture.Width = 1920 And picture.Height = 1080 Then
                        _numLandscapePictures += 1
                        CopyPictureToTargetDirectory(fileInfo.Name, PictureType.Landscape)
                    Else
                        _numPortraitPictures += 1
                        CopyPictureToTargetDirectory(fileInfo.Name, PictureType.Portrait)
                    End If

                Else
                    _numInvalidFiles += 1
                End If

            Catch ex As Exception
                _numInvalidFiles += 1
            End Try

        Next

        Dim sEvent As String = String.Concat("Total Files read: ", sourceFiles.Count, vbCrLf, "Valid Pictures: ", _numValidPictures, vbCrLf, "Duplicated Files: ",
                                             _numDuplicatedFiles, vbCrLf, "Copied Lanscape Pictures: ", _numCopiedLandscapePictures, vbCrLf, "Copied Portrait Pictures: ",
                                             _numCopiedPortraitPictures, vbCrLf, "Number of Invalid Files: ", _numInvalidFiles)

        eLog.WriteEntry(sEvent, EventLogEntryType.Information, 1)

    End Sub

    Private Sub CopyPictureToTargetDirectory(ByVal filename As String, type As PictureType)

        Dim destinationFile As String
        Dim sourceFile As String = String.Concat(SourceFolder, "\", filename)

        If type = PictureType.Landscape Then
            destinationFile = String.Concat(LandscapeFolder, "\", filename, ".jpg")
        Else
            destinationFile = String.Concat(PortraitFolder, "\", filename, ".jpg")
        End If

        Try
            If Not System.IO.File.Exists(destinationFile) Then
                System.IO.File.Copy(sourceFile, destinationFile)
                If type = PictureType.Landscape Then
                    _numCopiedLandscapePictures += 1
                Else
                    _numCopiedPortraitPictures += 1
                End If
            Else
                _numDuplicatedFiles += 1
            End If

        Catch ex As Exception
            _numInvalidFiles += 1
            Dim sEvent As String = String.Concat("Error in copying file ", filename)
            eLog.WriteEntry(sEvent, EventLogEntryType.Error, 100)

        End Try

    End Sub

End Module

Public Enum PictureType

    Landscape = 0
    Portrait = 1

End Enum