Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports System.Text
Public Class Dateisortierer
    Public errorDateiName As String
    Public lngCounter As Long
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Anwendung schließen
        Application.Exit()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim fileCreatedDate As DateTime 'Erstellungsdatum auf dem Medium
        Dim fileChangedDate As DateTime 'Letztes Änderungsdatum

        Dim firstPathString As String   'Pfad, in dem sich die Anwendung und damit die zu sortierenden Dateien befinden.
        Dim secPathString As String     'Ordnername mit Erstellungsdatum des Bildes
        Dim errPathString As String

        Dim zielPfadString As String
        Dim zielDateiString As String

        Dim extDatei As String
        Dim idExists As Boolean = False

        Dim files() As String = IO.Directory.GetFiles(Application.StartupPath) 'Files-Collection erstellen
        Dim anzahlFiles As Long
        Dim nrFile As Long = 0

        Dim imageData As Image
        Dim propItem As PropertyItem
        Dim Encoder As Encoding = Encoding.Default
        Dim imgBinString As String

        Dim ayear As String
        Dim aMonth As String
        Dim aDay As String
        Dim aDateString As String
        Dim fileExistsSince As DateTime = Nothing

        Try
            anzahlFiles = files.Count
            Label2.Text = nrFile & " / " & anzahlFiles

            'Den Ordnerpfad bilden: 1. Teil

            '1. Teil des Zielpfades: Startpfad der Anwendung zuweisen.
            firstPathString = Application.StartupPath

            'Erstellungsdatum aller Dateien (keine Ordner!) im Ordner auslesen.
            For Each f As String In files

                nrFile = nrFile + 1
                Label2.Text = nrFile & " / " & anzahlFiles
                Application.DoEvents()
                extDatei = IO.Path.GetExtension(f)


                If extDatei.ToLower = ".bmp" Or extDatei.ToLower = ".jpg" Or extDatei.ToLower = ".jpeg" _
                    Or extDatei.ToLower = ".tiff" Or extDatei.ToLower = ".png" Or extDatei.ToLower = ".gif" Then
                    errorDateiName = IO.Path.GetFileName(f)

                    'Erstellungsdatum der Datei f auslesen.
                    fileCreatedDate = IO.File.GetCreationTime(f)
                    'Änderungsdatum der Datei f auslesen.
                    fileChangedDate = IO.File.GetLastWriteTime(f)

                    'Aufnahmedatum aus EXIF-Metadaten auslesen: ID36867 (binär)
                    Try 'Um korrumpierte Bilder aufzufangen
                        imageData = Image.FromFile(f)
                        Application.DoEvents()
                        Using imageData
                            For Each info In imageData.PropertyItems
                                If info.Id = 36867 Then
                                    idExists = True
                                    imageData.GetPropertyItem(36867)
                                    propItem = imageData.GetPropertyItem(36867)
                                    imgBinString = Encoder.GetString(propItem.Value, 0, propItem.Len - 1)
                                    ayear = imgBinString.Substring(0, 4)
                                    aMonth = imgBinString.Substring(5, 2)
                                    aDay = imgBinString.Substring(8, 2)

                                    Try
                                        aDateString = ayear & "-" & aMonth & "-" & aDay
                                        fileExistsSince = DateTime.Parse(aDateString, System.Globalization.CultureInfo.InvariantCulture)
                                    Catch ex As Exception
                                        idExists = False
                                    End Try
                                End If
                            Next
                            imageData.Dispose()
                        End Using
                        'Den Ordnerpfad bilden: 2. Teil
                        '2. Teil des Zielpfades: Entscheidung zwischen 3 Daten.
                        secPathString = fileCreatedDate.ToString("yyyy-MM-dd") 'mindestens ein Datum gesichert


                        If fileChangedDate < fileCreatedDate Then
                            secPathString = fileChangedDate.ToString("yyyy-MM-dd")
                        End If

                        If idExists Then
                            'If String.IsNullOrEmpty(fileExistsSince) = False And fileExistsSince.ToString("yyyy-MM-dd") <> "0001-01-01" Then
                            If DateTime.Parse(secPathString, System.Globalization.CultureInfo.InvariantCulture) > fileExistsSince Then
                                secPathString = fileExistsSince.ToString("yyyy-MM-dd")
                            End If
                            'End If
                        End If
                        '3. Zielpfad nun zusammensetzen
                        zielPfadString = firstPathString & "\" & secPathString

                        'Erstelle diesen Ordner, sofern er noch nicht existiert.
                        If System.IO.Directory.Exists(zielPfadString) = False Then
                            System.IO.Directory.CreateDirectory(zielPfadString)
                        End If

                        'Existiert der Ordner (kann ja sein, dass etwas schief gegangen ist)?
                        If IO.Directory.Exists(zielPfadString) = True Then
                            'Verschiebe die Datei, sofern sie noch nicht im Zielordner existiert.
                            zielDateiString = zielPfadString & "\" & IO.Path.GetFileName(f)
                            If IO.File.Exists(zielDateiString) = False Then
                                IO.File.Move(f, zielDateiString)
                            End If
                        End If
                        imageData.Dispose()
                        'imageData = Nothing
                        'propItem = Nothing
                        'Encoder = Nothing
                    Catch ex As Exception
                        firstPathString = Application.StartupPath
                        errPathString = "unbekannt"

                        zielPfadString = firstPathString & "\" & errPathString

                        'Erstelle diesen Ordner, sofern er noch nicht existiert.
                        If System.IO.Directory.Exists(zielPfadString) = False Then
                            System.IO.Directory.CreateDirectory(zielPfadString)
                        End If

                        'Existiert der Ordner (kann ja sein, dass etwas schief gegangen ist)?
                        If IO.Directory.Exists(zielPfadString) = True Then
                            'Verschiebe die Datei, sofern sie noch nicht im Zielordner existiert.
                            zielDateiString = zielPfadString & "\" & IO.Path.GetFileName(f)
                            If IO.File.Exists(zielDateiString) = False Then
                                IO.File.Move(f, zielDateiString)
                            End If
                        End If
                    End Try
                End If
                idExists = False
                fileExistsSince = Nothing
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message & " Datei: " & errorDateiName)
        End Try
    End Sub

    Private Sub Dateisortierer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lngCounter = 0
    End Sub

End Class
