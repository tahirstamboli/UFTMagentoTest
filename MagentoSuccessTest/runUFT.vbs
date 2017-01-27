testResourcePath = "C:\"



'Getting the test path

Dim objArgs

Set objArgs = wscript.Arguments

testPath = objArgs(0)



'Determining that the test does exist

Dim objFSO

Set objFSO = CreateObject("Scripting.FileSystemObject")

DoesFolderExist = objFSO.FolderExists(testPath)

Set objFSO = Nothing



If DoesFolderExist Then

    Dim uftApp 'Declare the Application object variable

    Dim uftTest 'Declare a Test object variable

    Set uftApp = CreateObject("QuickTest.Application") 'Create the Application object

    uftApp.Launch 'Start QuickTest

    uftApp.Visible = false 'Make the QuickTest application visible

    uftApp.Open testPath, False 'Open the test in read-only mode

    Set uftTest = uftApp.Test



    Set uftResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object

    uftResultsOpt.ResultsLocation = testResourcePath ' Specify the location to save the test results.

    uftTest.Run uftResultsOpt,True 'Run the test and wait until end of the test run

    If uftTest.LastRunResults.Status = "Failed" Then
        'qtTest.Run 'Run the test
        uftTest.Close 'Close the test
        uftApp.Quit
        WScript.Quit 1
    Else
        uftTest.Close 'Close the test
        uftApp.Quit
    End If

Else

    'Couldn't find the test folder. That's bad. Guess we'll have to report on how we couldn't find the test.

    'Insert reporting mechanism here.

End If



