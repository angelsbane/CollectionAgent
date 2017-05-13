Imports Atlas2
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections





Public Class Form1
    'GLOBAL DECLARATIONS
    Dim WithEvents MyPage As Page
    Dim InputArguments As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    Dim WorkingFolder As String 'location of Current Active Folder
    Dim lToday As Boolean = False
    Dim ArchiveFolder As String
    Dim ComparisonFolder As String
    Dim matches As MatchCollection 'Used as holder for scrapped websites
    Dim lists As New ArrayList
    Dim siteNum As Integer 'Used to keep track of site numbers for ad naming
    'Test Application for Master Log
    Dim archiveLog As String
    'END TESTING
    Declare Function AttachConsole Lib "kernel32.dll" (ByVal dwProcessId As Int32) As Boolean
    Declare Function FreeConsole Lib "kernel32.dll" () As Boolean


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not Debugger.IsAttached Then
            Me.WindowState = FormWindowState.Minimized
            Me.Visible = False
        End If

        InputArguments = My.Application.CommandLineArgs

    End Sub

    ''' <summary>
    ''' returns the Input argument which starts with the supplied code.  Argumens should follow a format of /u:example.intuli.com where u is the argumentcode and example.intuli.com would be the value
    ''' </summary>
    ''' <param name="ArgumentCode"></param>
    ''' <remarks></remarks>
    Private Function GetArgument(ByVal ArgumentCode As String) As String
        ArgumentCode = String.Concat("/", ArgumentCode, ":")

        'if we have less than 1 input argument return empty string.
        If InputArguments.Count < 1 Then Return ""

        Dim ReturnValue As String

        For n = 0 To InputArguments.Count - 1
            If InputArguments(n).StartsWith(ArgumentCode) Then
                ReturnValue = InputArguments(n)
                Return ReturnValue.Substring(ReturnValue.IndexOf(":") + 1)
            End If
        Next

        Return ""
    End Function


    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim ValidCollection As Boolean = False
        Dim Spider As Boolean = False
        'x is used as iterator for Lists for the spider
        Dim x As Integer = 0
        'initialize page object 
        'I.E. Mode is defaulted as IE11; possibly add manually selected versions
        'Just a test
        SetIEMode(IEVersion.IE11)
        'end test
        MyPage = New Page(Me.SplitContainer1.Panel1, Me.lst_StatusLog)
        'Changed by Bryan - testing 11-15-2016
        ' MyPage.Silent = True
        'End Change
        'Add handler for Status Updates
        AddHandler ApplicationLog.NewEntry, AddressOf StatusUpdate
        'Attach output console for returning data to the manager program
        If Not Debugger.IsAttached Then AttachConsole(-1)

        'get the folders and id from the input arguments
        WorkingFolder = String.Format("{0}\{1}", GetArgument("wf"), GetArgument("id"))
        Dim ArchiveParentFolder As String = String.Format("{0}\Working\", GetArgument("af"))
        ArchiveFolder = String.Format("{0}\Working\{1}", GetArgument("af"), GetArgument("id"))
        'Added ArchiveFolderlog
        Dim ArchiveFolderLog = String.Format("{0}\Archive\logs\{1}", GetArgument("af"), GetArgument("id"))
        ComparisonFolder = String.Format("{0}\{1}", GetArgument("cf"), GetArgument("id"))


        'if archive parent folder doesn't exist, create it.
        If Not Directory.Exists(ArchiveParentFolder) Then
            Directory.CreateDirectory(ArchiveParentFolder)
        End If

        'if archive folder exists, delete it so it can be replaced by the existing working folder
        If Directory.Exists(ArchiveFolder) Then
            Directory.Delete(ArchiveFolder, True)
        End If

        If Not Directory.Exists(ArchiveFolderLog) Then
            Directory.CreateDirectory(ArchiveFolderLog)
        End If

        'if working folder exists, back it up to the Archive
        If Directory.Exists(WorkingFolder) Then
            '      Directory.Move(WorkingFolder, ArchiveFolder)
        End If

        'Create a new Working Folder
        'ADDED By Bryan 05-09-16
        If File.Exists(WorkingFolder + "\log.txt") Then
            Dim logFile As String
            Dim Test As String
            logFile = ReadTextFile(WorkingFolder + "\log.txt")
            'TEMP VAR: Temporary Variable for Usage for date storage
            Test = String.Concat(Date.Today.Month.ToString + "/" + Date.Today.Day.ToString + "/" + Date.Today.Year.ToString)
            'END TEMP
            If logFile.Contains(Test) Then
                lToday = True
            Else
                '1st Step:  Move Old Directory
                '2nd Step:  Create New Directory
                Directory.Move(WorkingFolder, ArchiveFolder)
                Directory.CreateDirectory(WorkingFolder)
            End If
        Else
            Directory.CreateDirectory(WorkingFolder)
        End If
        'FINISH ADDED BY BRYAN 05/09/2016
        'Directory.CreateDirectory(WorkingFolder)

        'Stream the log to a file

        ApplicationLog.StreamtoFile(String.Concat(WorkingFolder, "\log.txt"))


        Atlas2.SuppressErrors = False



        Do While lists.Count > x - 1
            Try
                '0 is used as the defining point to spider, afterwards it is skipped
                'TODO:  Case Works: Rewrite to Regular Expression Identification
                If GetArgument("cv").ToString = "ads" Or GetArgument("cv").ToString = "ADS" Then
                    If x = 0 Then
                        MyPage.Load(GetArgument("u"))
                        MyPage.ForceWait(2000)
                        siteNum = 0
                        Try
                            Spider = webScrape()
                        Catch ex As Exception
                            AddError("Spider Failed")
                            '     lists.Add(GetArgument("u"))
                        Finally
                            AddEntry("Sites Found: " + CStr(lists.Count))
                        End Try

                        Try
                            File.Create(String.Concat(WorkingFolder + "\" + "sitefile.txt")).Dispose()
                            Dim sw As StreamWriter = File.AppendText(WorkingFolder + "\" + "sitefile.txt")
                            AddEntry("Creating Site File: ")
                            '  File.Create(String.Concat(WorkingFolder + "\" + "sitefile.txt")).Dispose()
                            For num As Integer = 0 To lists.Count - 1 Step 1

                                sw.WriteLine("(" + CStr(num) + ")-" + lists.Item(num))
                            Next

                            sw.Dispose()
                        Catch ex As Exception
                            AddError("Cannot Create SiteFile")
                        End Try
                    Else
                        siteNum = x
                        MyPage.Load(lists(x - 1).ToString)
                        MyPage.ForceWait(3000)
                        AddEntry("Loading Page: " + lists(x - 1).ToString)
                    End If
                Else

                End If

            Catch ex As Exception
                AddError("Web Spider Failed: ", ex)
            End Try
            'Should skip Spidering
            Try
                If GetArgument("cv").ToString = "ads" Or GetArgument("cv").ToString = "ADS" Then
                    ValidCollection = CollectAds()
                    siteNum = x - 1
                Else
                    ValidCollection = CollectSite()
                End If
            Catch ex As Exception
                AddError("Collection Error: ", ex)
                x += 1
            Finally
                If ValidCollection Then
                    'important to end with this so the manager knows the collection was successful.
                    AddEntry("Completed " + CStr(x))
                    x += 1
                Else
                    AddError("Unable to collect site: " + CStr(x))
                    x += 1
                End If



                'close application 
                '  Me.Close()

            End Try
        Loop
        AddEntry("Completed Page Scrape")
        If Not Debugger.IsAttached Then
            'release the console
            FreeConsole()
        End If
        'close application 
        Me.Close()
    End Sub

    'handle status updates by adding the dattime and passing them to the console
    Private Sub StatusUpdate(ByVal Sender As Object, ByVal e As LogEventArgs)
        System.Console.WriteLine(e.Message)
    End Sub

    Private Sub ArchiveUpdate(ByVal Sender As Object, ByVal e As LogEventArgs)
        System.Console.WriteLine(e.Message)
    End Sub

    ''' <summary>
    ''' This function is where the site collection code should be.  
    ''' It has access to the Input Arguments collection and the MyPage object.
    ''' </summary>
    ''' <remarks></remarks>
    Private Function webScrape() As Boolean
        ' Dim url As String = MyPage.FetchHTML(GetArgument("u"))
        ' Dim url As String = MyPage.HTML.ToString
        Dim crawl As Array = MyPage.GetCrawlLinks
        'End Spidering delcarations
        'THIS WILL BE PUT INTO A SEPERATE FUNCTION
        '  lists.Add(url)
        ' matches = Regex.Matches(url, "<a.*?href=""(.*?)"".*?>(.*?)</a>")
        Try
            For Each match As String In crawl
                ''For Each match As match In crawl
                ''  Dim matchUrl As String = match.Groups(1).Value
                Dim matchUrl As String = match.ToString
                ''Ignore all anchor links
                'If matchUrl.StartsWith("#") Then
                '    Continue For
                'End If
                'TEST: get rid of .css links
                If matchUrl.Contains(".css?") Then
                    Continue For
                End If
                ''TEst for JAMA - Jama has large amount aspx sites and they seem to be problematic
                ' If matchUrl.Contains("aspx") Or matchUrl.Contains("click?") Then
                '  Continue For
                ' End If
                ''Ignore all javascript calls
                'If matchUrl.ToLower.StartsWith("javascript:") Then
                '    Continue For
                'End If
                ''Ignore all email links
                'If matchUrl.ToLower.StartsWith("mailto:") Then
                '    Continue For
                'End If
                lists.Add(matchUrl.ToString)

                If lists.Count = 10 Then
                    Exit For
                End If
            Next

        Catch ex As Exception

            Return False
        End Try




        Return True
        'END TRIAL TEST - if Successful, put into a seperate function
    End Function
    Private Function CollectAds() As Boolean
        Dim ImageCounter As Integer = 1
        Dim BannerCounter As Integer = 1
        Dim SideCounter As Integer = 1
        Dim GroupCounter As Integer = 1
        Dim ImageHashCollection As New List(Of String)
        'TESTING MULTIPLE HASHTABLES
        Dim BannerHashCollection As New List(Of String)
        Dim GroupHashCollection As New List(Of String)
        Dim SideHashCollection As New List(Of String)
        Dim adList As New List(Of String) 'Used to hold Ads that are going to be put into MainArticle.Txt
        'END TABLES
        Dim filelist As New ArrayList()


        'load the page supplied with the /u: input argument

        ' MyPage.Load(GetArgument("u"))
        AddEntry("Waiting 5 seconds for page scripts to run.")
        'wait for 5 seconds for the page to fully load.
        MyPage.ForceWait(2000)



        'Test to see if Failed Script errors are disabled
        'CHANGE BACK TO DIM
       

        'ADDED BY BRYAN - Test for stopping navigation
        'if the Page HTML is empty then, then the page did not load propery, make another attempt



        If String.IsNullOrEmpty(MyPage.HTML) Then
            AddWarning("Page failed to load property, retrying page load.")
            MyPage.Load(GetArgument("u"))
            AddEntry("Waiting 5 seconds for page scripts to run.")
            MyPage.ForceWait(5000)
        End If

        'if Page is still empty then exit with an error
        If String.IsNullOrWhiteSpace(MyPage.HTML) Then
            AddError("Unable to Load page.")
            Return False
        End If
        'ADDITION:  Adding a page check to see if it has "Cannot load", "No Internet Connection" and other parameters

        'click on healthcare professional button, if found
        'will probably need to expand the RegEx pattern and/or add additional checks
        'CHANGE BACK TO DIM LATER
        Dim ProfButton As Element = MyPage.GetElementbyRegEx("am.+?professional", "a")
        If ProfButton.IsValid Then
            ProfButton.Click()
            'give the page time to load after the click, seems some sites have scripts and/or animations that run.
            MyPage.ForceWait(5000)
            AddEntry("Bypassed Professional Prompt")
        End If
        'TEST TO DISAbLE SCRIPTING ON PAGE
        ProfButton = MyPage.GetElementbyRegEx("Yes", "a")
        If ProfButton.IsValid Then
            ProfButton.Click()
            'give the page time to load after the click, seems some sites have scripts and/or animations that run.
            MyPage.ForceWait(5000)
            AddEntry("Bypassed Script Prompt")
        End If
        'TEST TO DISABLE PROMPT SCRIPT
        ProfButton = MyPage.GetElementbyRegEx("continue running scripts", "a")
        If ProfButton.IsValid Then
            ProfButton.Click()
            'give the page time to load after the click, seems some sites have scripts and/or animations that run.
            MyPage.ForceWait(5000)
            AddEntry("Bypassed Script Prompt")
        End If
        'click on continue to page link, if found
        ProfButton = MyPage.GetElementbyRegEx("continue\sto.+?page", "a")
        If ProfButton.IsValid Then
            ProfButton.Click()
            'give the page time to load after the click, seems some sites have scripts and/or animations that run.
            MyPage.ForceWait(5000)
            AddEntry("Bypassed Professional Prompt")
        End If

        'Click for Patient Site Indication : If condition to ensure it doesn't click Patient on HCP site
        If Not MyPage.URL.Contains("hcp") Then
            ProfButton = MyPage.GetElementbyRegEx("patient\s.+continue", "a")
            If ProfButton.IsValid Then
                ProfButton.Click()
                MyPage.ForceWait(5000)
                AddEntry("Bypassed Patient Prompt")
            End If
        End If

        'click on radio button Example:  http://www.adempas-us.com/hcp/
        ProfButton = MyPage.GetElementbyRegEx("hcpProvider")
        If ProfButton.IsValid Then
            ProfButton.Click()
            If MyPage.GetElementbyID("hcpConfirmation").IsValid Then
                MyPage.GetElementbyID("hcpConfirmation").Click()
            End If
        End If

        If File.Exists(WorkingFolder + "\Banners.md5") Then
            BannerHashCollection = LoadBinaryObject(String.Concat(WorkingFolder, "\Banners.md5"))
            BannerCounter = BannerHashCollection.Count + 1
        End If

        If File.Exists(WorkingFolder + "\Groups.md5") Then
            GroupHashCollection = LoadBinaryObject(String.Concat(WorkingFolder, "\Groups.md5"))
            GroupCounter = GroupHashCollection.Count + 1
        End If

        If File.Exists(WorkingFolder + "\Sides.md5") Then
            SideHashCollection = LoadBinaryObject(String.Concat(WorkingFolder, "\Sides.md5"))
            SideCounter = SideHashCollection.Count + 1
        End If


        'TESTING Browser Resizes
        '  For Each Site As String In lists

        ' MyPage.Load(GetArgument("u"))

        If MyPage.Height > 500 Then

            MyPage.ResizeBrowsertoDocument()
        Else
            MyPage.ResizeBrowser(1280)
        End If

        'MyPage.LockBrowserSize = True

        Dim CollectTime As DateTime = Now

        '   Dim compval As String = GetArgument("cv")
        '********************************
        'IDEA Build as a series of Cases 
        ' For Each Site In matches

        ' Next
        If GetArgument("cv") = "ads" Or GetArgument("cv") = "ADS" Then
            'Variable Declarations for Ads
            Dim BanId As String = "nothing" 'Used as Case for Banners
            Dim BanTest As Element ' New Test Case for Banners
            Dim groupTest As Element
            Dim sideTest As Element 'Used for SideBanner(s)
            Dim GroId As String = "nothing" 'used as Case for Group Id's
            Dim pageInfo As String = MyPage.HTML.ToString
            Dim noGroup As Boolean = False
            Dim noSide As Boolean = False
            Dim noBanner As Boolean = False

            '  Dim noSide As Boolean = False

            'End Variable Declaration
            'NOTE:  At this point, this is where the sites should start to loop?
            'tempElement used to gather the amount of elements on the page to sort
            Dim tempElement As ElementCollection
            '= MyPage.GetAllElementsbyAttributeRegEx("img", "src", "ad\w")
            'USED for google ads primarily and websites that use Google Ad's containers
            ''Iframes are easy to work with-
            '' bloodjournal.org,webmd.com,jama.network.com0
            If pageInfo.Contains("iframe") Then

                If MyPage.GetAllElementsbyAttributeRegEx("iframe", "id", "ad\w").IsValid Then
                    tempElement = MyPage.GetAllElementsbyAttributeRegEx("iframe", "id", "ad\w")
                ElseIf MyPage.GetAllElementsbyAttributeRegEx("iframe", "src", "ad\w").IsValid Then
                    tempElement = MyPage.GetAllElementsbyAttributeRegEx("iframe", "src", "ad\w")
                ElseIf MyPage.GetAllElementsbyAttributeRegEx("iframe", "src", "html\w").IsValid Then
                    tempElement = MyPage.GetAllElementsbyAttributeRegEx("iframe", "src", "html\w")
                Else
                    'This is out of desperation so it doesn't error out
                    'tempElement = MyPage.GetAllElementsbyAttributeRegEx("iframe", "scrolling", "no")
                    tempElement = MyPage.GetAllElementsbyTagRegEx("iframe")

                End If

                For Each item As Element In tempElement.Items
                    If BanTest Is Nothing Then
                        If item.Height >= 90 And item.Width >= 700 Then
                            'testcase for Identifying Banners
                            BanTest = item
                        End If
                    End If

                    If groupTest Is Nothing Then
                        If item.Height >= 250 And item.Height <= 300 Then
                            If item.Width >= 300 Then
                                groupTest = item
                            End If
                        End If
                    End If

                    '  If sideTest Is Nothing Then
                    If item.Height >= 600 And item.Width = 160 Then
                        sideTest = item
                    End If

                Next
                tempElement = Nothing
            End If

            'NOTES: This is really a brute force method but should still work in the short term until refined
            If BanTest IsNot Nothing And groupTest IsNot Nothing Then
                'do nothing
            Else
                If MyPage.GetAllElementsbyAttributeRegEx("img", "src", "ad\w").IsValid Then
                    tempElement = MyPage.GetAllElementsbyAttributeRegEx("img", "src", "ad\w")
                    For Each item As Element In tempElement.Items
                        If BanTest Is Nothing Then
                            If item.Height >= 90 And item.Width >= 700 Then
                                'testcase for Identifying Banners
                                BanTest = item
                            End If
                        End If
                        If groupTest Is Nothing Then
                            If item.Height >= 250 And item.Height <= 300 Then
                                If item.Width >= 300 Then
                                    groupTest = item
                                End If
                            End If
                        End If
                        'If sideTest Is Nothing Then
                        '    If item.Height >= 600 And item.Width = 160 Then
                        '        sideTest = item
                        '    End If
                        'End If
                    Next

                End If
            End If
            'Define MainArticle File for use of files:

            Do Until Now.Subtract(CollectTime).TotalSeconds >= 120
                Dim myElem As Atlas2.Element
                Dim myGroup As Atlas2.Element
                Dim found As Boolean = True

                'CASE Banner Tree

                '  If MyPage.IsMatch("iframe") Then
                'Test Procedure to ensure that sites on .aspx are valid


                'This is targetting google banner ads - this is the first test
                'Successful with IFRAME X
                'Successful with AdClick X
                'Successful with Double-Click X

                Dim bannerImage, GroupImage, sideImage As System.Drawing.Bitmap
                'IMAGE GRABBING BLOCK:  This starts the process of grabbing the image
                'ORIGINAL Dim currentIMage as System.Drawing.Bitmap = Bantest.Snapshot
                ' Dim bannerImage As System.Drawing.Bitmap = BanTest.Snapshot
                'Dim groupImage As System.Drawing.Bitmap = groupTest.Snapshot
                ' Dim sideImage As System.Drawing.Bitmap = sideTest.Snapshot


                Dim quickCounter As Integer = 0
                'reject the image if it is one solid color
                Do While quickCounter <= 3

                    If BanTest Is Nothing And groupTest Is Nothing And sideTest Is Nothing Then
                        AddEntry("No Elements Found - skipping page")
                        quickCounter = 3
                        Exit Do
                    End If

                    Try
                        bannerImage = BanTest.Snapshot
                        If Not IsAllOneColor(bannerImage) Then
                            'Do Nothing
                        Else
                            BanTest = Nothing
                        End If
                    Catch ex As Exception
                        bannerImage = Nothing
                        BanTest = Nothing
                    End Try

                    If BanTest IsNot Nothing Then
                        'BanTest.ToPNGSeries(String.Concat(WorkingFolder, "\", GetArgument("id"), "-BA-", siteNum.ToString, ".png"), 60, 30)
                        'if the current image is different than the last image, save it
                        Dim CurrentMD5 As String = MD5Image(bannerImage)
                        Try
                            If Not BannerHashCollection.Contains(CurrentMD5) Then
                                'FOR SAVETOPNGSERIES
                                ' BanTest.ToPNGSeries(String.Concat(WorkingFolder, "\", GetArgument("id"), "-BA-", siteNum.ToString, ".png"), 60, 30)
                                'AddEntry("Banner Captured")
                                'BannerHashCollection.Add(CurrentMD5)
                                'END SAVE TO PNGSERIES
                                BannerHashCollection.Add(CurrentMD5)
                                AddEntry("MD5(Banner): " + CurrentMD5.ToString)
                                bannerImage.Save(String.Concat(WorkingFolder, "\", GetArgument("id"), "-BA-", BannerCounter, ".png"), System.Drawing.Imaging.ImageFormat.Png)
                                adList.Add(String.Concat(GetArgument("id"), "-BA-", BannerCounter))
                                BannerCounter += 1
                            End If

                        Catch ex As Exception
                            BanTest = Nothing
                        End Try

                    End If
                    Try
                        GroupImage = groupTest.Snapshot
                        If Not IsAllOneColor(GroupImage) Then
                            'Do Nothing
                        Else
                            groupTest = Nothing
                        End If
                    Catch ex As Exception
                        GroupImage = Nothing
                        groupTest = Nothing
                    End Try
                    If groupTest IsNot Nothing Then
                        GroupImage = groupTest.Snapshot
                        Dim CurrentMD5 = MD5Image(GroupImage)
                        Try
                            If Not GroupHashCollection.Contains(CurrentMD5) Then
                                'FOR SAVETOPNG SERIES
                                '  groupTest.ToPNGSeries(String.Concat(WorkingFolder, "\", GetArgument("id"), "-GP-", siteNum.ToString, ".png"), 60, 30)
                                ' AddEntry("Group Captured...")
                                'END SAVETOPNG SERIES
                                GroupHashCollection.Add(CurrentMD5)
                                AddEntry("MD5(Group): " + CurrentMD5.ToString)
                                GroupImage.Save(String.Concat(WorkingFolder, "\", GetArgument("id"), "-GP-", GroupCounter, ".png"), System.Drawing.Imaging.ImageFormat.Png)
                                adList.Add(String.Concat(GetArgument("id"), "-GP-", GroupCounter))
                                GroupCounter += 1
                            End If
                        Catch ex As Exception
                            groupTest = Nothing
                        End Try

                    End If
                    Try
                        sideImage = sideTest.Snapshot
                        If Not IsAllOneColor(sideImage) Then
                            'Do Nothing
                        Else
                            sideTest = Nothing
                        End If
                    Catch ex As Exception
                        sideImage = Nothing
                        sideTest = Nothing
                    End Try
                    If sideTest IsNot Nothing Then
                        Dim CurrentMD5 = MD5Image(sideImage)

                        Try
                            If Not SideHashCollection.Contains(CurrentMD5) Then
                                'START SAVE TO PNGSERIES
                                'sideTest.ToPNGSeries(String.Concat(WorkingFolder, "\", GetArgument("id"), "-SI-", siteNum.ToString, ".png"), 60, 30)
                                ' AddEntry("Side Image")
                                ' SideHashCollection.Add(CurrentMD5)
                                'END PNG SERIES
                                SideHashCollection.Add(CurrentMD5)
                                AddEntry("MD5(Side): " + CurrentMD5.ToString)
                                sideImage.Save(String.Concat(WorkingFolder, "\", GetArgument("id"), "-SI-", SideCounter, ".png"), System.Drawing.Imaging.ImageFormat.Png)
                                adList.Add(String.Concat(GetArgument("id"), "-SI-", SideCounter))
                                SideCounter += 1
                            End If
                        Catch ex As Exception
                            sideTest = Nothing
                        End Try

                    End If

                    quickCounter += 1
                Loop


            Loop

        End If


        'sleep for 2 sec
        ' ThreadSleep(2000)

        ' If noSide And noGroup And noBanner = True Then
        'AddEntry("Skipping Page")
        ' Exit Do
        'End If



        ' End If

        ''save the body Inner Text as the page text
        WriteTextFile(MyPage.Body.InnerText, String.Concat(WorkingFolder, "\body.txt"))

        ''Save the main article text
        ' WriteTextFile(CleanHTML(MyPage.GetMainArticleCollection.HTML), String.Concat(WorkingFolder, "\MainArticle.txt"))

        If Not My.Computer.FileSystem.FileExists(String.Concat(WorkingFolder, "\MainArticle.txt")) Then
            File.Create(String.Concat(WorkingFolder + "\" + "MainArticle.txt")).Dispose()
        End If

        Dim aw As StreamWriter = File.AppendText(WorkingFolder + "\" + "MainArticle.txt")
        If adList.Count > 0 Then
            Try
                '  File.Create(String.Concat(WorkingFolder + "\" + "sitefile.txt")).Dispose()
                aw.WriteLine("(" + CStr(siteNum) + ")-" + MyPage.URL.ToString)
                For num As Integer = 0 To adList.Count - 1 Step 1

                    aw.WriteLine("(" + CStr(siteNum) + ")-" + adList.Item(num))
                Next
            Catch ex As Exception
                AddError("Can't Modify adList")
            End Try
        End If
        aw.Dispose()

        ''Save the MD5 Collections for advertisements
        SaveBinaryObject(BannerHashCollection, String.Concat(WorkingFolder, "\Banners.md5"))
        SaveBinaryObject(GroupHashCollection, String.Concat(WorkingFolder, "\Groups.md5"))
        SaveBinaryObject(SideHashCollection, String.Concat(WorkingFolder, "\Sides.md5"))

        Dim ArchiveFolderLog = String.Format("{0}\Archive\logs\{1}", GetArgument("af"), GetArgument("id"))

        WriteTextFile(ApplicationLog.toString, String.Concat(ArchiveFolderLog, "\archivelog.txt"), True)
        Return True
    End Function
    'This is the function for the advertisement collection series, will probably move this on to a whole sepereate program but for now will utilize
    'the collection series as a testing grounds

    Private Function CollectSite() As Boolean
        Dim ImageCounter As Integer = 1
        Dim ImageHashCollection As New List(Of String)

        'load the page supplied with the /u: input argument
        MyPage.Load(GetArgument("u"))

        AddEntry("Waiting 5 seconds for page scripts to run.")
        'wait for 5 seconds for the page to fully load.
        MyPage.ForceWait(5000)
        MyPage.LoadPopups = False
        'ADDED BY BRYAN - Test for stopping navigation
        'if the Page HTML is empty then, then the page did not load propery, make another attempt
        If String.IsNullOrEmpty(MyPage.HTML) Then
            AddWarning("Page failed to load property, retrying page load.")
            MyPage.Load(GetArgument("u"))
            AddEntry("Waiting 5 seconds for page scripts to run.")
            MyPage.ForceWait(5000)
        End If


        'if Page is still empty then exit with an error
        If String.IsNullOrWhiteSpace(MyPage.HTML) Then
            AddError("Unable to Load page.")
            Return False
        End If
        'ADDITION:  Adding a page check to see if it has "Cannot load", "No Internet Connection" and other parameters

        'click on healthcare professional button, if found
        'will probably need to expand the RegEx pattern and/or add additional checks
        Dim ProfButton As Element = MyPage.GetElementbyRegEx("am.+?professional", "a")
        If ProfButton.IsValid Then
            ProfButton.Click()
            'give the page time to load after the click, seems some sites have scripts and/or animations that run.
            MyPage.ForceWait(5000)
            AddEntry("Bypassed Professional Prompt")
        End If

        'click on continue to page link, if found
        ProfButton = MyPage.GetElementbyRegEx("continue\sto.+?page", "a")
        If ProfButton.IsValid Then
            ProfButton.Click()
            'give the page time to load after the click, seems some sites have scripts and/or animations that run.
            MyPage.ForceWait(5000)
            AddEntry("Bypassed Professional Prompt")
        End If
        'Click for Patient Site Indication : If condition to ensure it doesn't click Patient on HCP site
        If Not MyPage.URL.Contains("hcp") Then
            ProfButton = MyPage.GetElementbyRegEx("patient\s.+continue", "a")
            If ProfButton.IsValid Then
                ProfButton.Click()
                MyPage.ForceWait(5000)
                AddEntry("Bypassed Patient Prompt")
            End If
        End If

        'click on radio button Example:  http://www.adempas-us.com/hcp/
        ProfButton = MyPage.GetElementbyRegEx("hcpProvider")
        If ProfButton.IsValid Then
            ProfButton.Click()
            If MyPage.GetElementbyID("hcpConfirmation").IsValid Then
                MyPage.GetElementbyID("hcpConfirmation").Click()
            End If
        End If



        'resize the page to a standard width with an autodetected height
        '   MyPage.ResizeBrowser(1280)

        'TESTING Broweser Resizes
        If MyPage.Height > 500 Then
            'MyPage.ResizeBrowser(1280)
            MyPage.ResizeBrowsertoDocument(True)
        Else
            MyPage.ResizeBrowser(1280)
        End If

        MyPage.LockBrowserSize = True

        Dim CollectTime As DateTime = Now

        'loop image collection for 60 seconds
        AddEntry("Beginning image collection: ")
        Do Until Now.Subtract(CollectTime).TotalSeconds >= 60
            'Revision Notes for V.1.03:  Switching to MyElementtoPNGSERIES test
            Try
                'get the current image
                Dim CurrentImage As System.Drawing.Bitmap = MyPage.ToImage

                'reject the image if it is one solid color
                If Not IsAllOneColor(CurrentImage) Then

                    Dim CurrentMD5 As String = MD5Image(CurrentImage)

                    'if the current image is different than the last image, save it
                    If Not ImageHashCollection.Contains(CurrentMD5) Then
                        ImageHashCollection.Add(CurrentMD5)
                        CurrentImage.Save(String.Concat(WorkingFolder, "\FullImage", ImageCounter, ".png"), System.Drawing.Imaging.ImageFormat.Png)
                        ImageCounter += 1
                    End If
                End If
            Catch
                AddError("Unable to capture image.")
            End Try

            'sleep for 2 sec
            ThreadSleep(2000)

        Loop

        'save the page
        Try
            MyPage.SaveAsHTML(String.Concat(WorkingFolder, "\FullSite.html"))
        Catch
            AddError("Unable to save HTML Fullesite page.")
        End Try
        '####Added by Bryan 11-17-2015
        'Test to see if Mfg Number matches exclusively
        Try
            If GetArgument("mfg") = "na" Then


            ElseIf MyPage.Body.InnerText.Contains(String.Format("{0}", GetArgument("mfg"))) Then
                Dim Manufound As String = GetArgument("mfg")
                AddEntry("Manufacturer Number Found: " + Manufound.ToString)
                'Added a read text file for testing purposes ' 3-15-18 *DELETE LATER
                WriteTextFile(Manufound, WorkingFolder + "\manu.txt")
            Else
                AddEntry("Manufacturer Number Not Found")
            End If

        Catch ex As Exception
            AddError("Unable to look for Manufacturer Number")
        End Try
        '####
        'save the body Inner Text as the page text
        WriteTextFile(MyPage.Body.InnerText, String.Concat(WorkingFolder, "\body.txt"))

        'Save the main article text // FOR ADS:  loop through for loop and add to MainArticle.txt
        WriteTextFile(CleanHTML(MyPage.GetMainArticleCollection.HTML), String.Concat(WorkingFolder, "\MainArticle.txt"))
        
        'Save the MD5 Collection
        SaveBinaryObject(ImageHashCollection, String.Concat(WorkingFolder, "\Images.md5"))


        Dim ArchiveFolderLog = String.Format("{0}\Archive\logs\{1}", GetArgument("af"), GetArgument("id"))

        WriteTextFile(ApplicationLog.toString, String.Concat(ArchiveFolderLog, "\archivelog.txt"), True)
        Return True
    End Function

    Private Shared Function RetrieveNumber(ByVal value As String) As Integer
        Dim returnVal As String = String.Empty
        Dim collection As MatchCollection = Regex.Matches(value, "\d+")
        For Each m As Match In collection
            returnVal += m.ToString()
        Next
        Return Convert.ToInt32(returnVal)
    End Function


End Class
