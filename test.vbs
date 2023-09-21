option explicit
dim objShell, objFile, objJson
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
dim directory
directory = objFile.GetParentFolderName(wscript.ScriptFullName)
Include(directory & "\aspJSON.vbs")
set objJson = new aspJSON

sub Main()
    dim oJSON
    Set oJSON = New aspJSON

    With oJSON.data

        .Add "familyName", "Smith"                      'Create value
        .Add "familyMembers", oJSON.Collection()

        With oJSON.data("familyMembers")

            .Add 0, oJSON.Collection()                  'Create unnamed object
            With .item(0)
                .Add "firstName", "John"
                .Add "age", 41

                .Add "job", oJSON.Collection()          'Create named object
                With .item("job")
                    .Add "function", "Webdeveloper"
                    .Add "salary", 70000
                End With
            End With


            .Add 1, oJSON.Collection()
            With .item(1)
                .Add "firstName", "Suzan"
                .Add "age", 38
                .Add "interests", oJSON.Collection()    'Create array
                With .item("interests")
                    .Add 0, "Reading"
                    .Add 1, "Tennis"
                    .Add 2, "Painting"
                End With
            End With

            .Add 2, oJSON.Collection()
            With .item(2)
                .Add "firstName", "John Jr."
                .Add "age", 2.5
            End With

        End With

    End With

    wscript.Echo oJSON.JSONoutput() 
end sub

sub Include(scriptName)
    ExecuteGlobal objFile.OpenTextFile(scriptName).ReadAll()
End Sub

sub Breakpoint(message)
    wscript.Echo(message)
    wscript.Quit
end sub

Main()