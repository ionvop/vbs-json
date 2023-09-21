option explicit
dim objShell, objFile, objHttp, objJson
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
set objHttp = CreateObject("MSXML2.XMLHTTP.6.0")
dim directory
directory = objFile.GetParentFolderName(wscript.ScriptFullName)
Include(directory & "\aspJSON.vbs")
set objJson = new aspJSON

sub Main()
    dim data, messages, message, res

    'Example: Generating a JSON string from a dictionary object
    set data = CreateObject("Scripting.Dictionary")
    call data.Add("model", "gpt-3.5-turbo")
    messages = array()
    set message = CreateObject("Scripting.Dictionary")

    call message.Add("role", "user")
    call message.Add("content", "Hello")
    call Push(messages, message)
    message.RemoveAll()

    call message.Add("role", "assistant")
    call message.Add("content", "Hi! How may I assist you?")
    call Push(messages, message)
    message.RemoveAll()

    call message.Add("role", "user")
    call message.Add("content", "Say this is a test!")
    call Push(messages, message)
    message.RemoveAll()

    call data.Add("messages", messages)
    call data.Add("temperature", 0.7)
    data = JsonSerialize(data)
    wscript.Echo(data)

    'Example: Accessing values from a JSON string using the aspJSON class
    res = objFile.OpenTextFile(directory & "\res.json").ReadAll()
    objJson.loadJSON(res)
    wscript.Echo(objJson.data("choices")(0).item("message").item("content"))
end sub

function Push(inputArray, pushData)
    dim newData
    redim preserve inputArray(ubound(inputArray) + 1)

    if typename(pushData) = "Dictionary" then
        set newData = CloneDict(pushData)
        set inputArray(ubound(inputArray)) = newData
        Push = inputArray
        exit function
    end if

    inputArray(ubound(inputArray)) = pushData
    Push = inputArray
end function

function CloneDict(data)
    dim element, result
    set result = CreateObject("Scripting.Dictionary")

    for each element in data.Keys()
        call result.Add(element, data.Item(element))
    next

    set CloneDict = result
end function

function JsonSerialize(data)
    dim i, element, result
    result = ""

    select case typename(data)
        case "Dictionary"
            result = result & "{"

            for each element in data.Keys()
                result = result & """" & JsonEscape(element) & """: "

                select case typename(data.Item(element))
                    case "Boolean"
                        result = result & lcase(data.Item(element))
                    case "Integer", "Double"
                        result = result & data.Item(element)
                    case "String"
                        result = result & """" & JsonEscape(data.Item(element)) & """"
                    case "Variant()", "Dictionary"
                        result = result & JsonSerialize(data.Item(element))
                end select

                result = result & ", "
            next

            result = left(result, len(result) - 2)
            result = result & "}"
        case "Variant()"
            result = result & "["

            for each element in data
                select case typename(element)
                    case "Boolean"
                        result = result & lcase(element)
                    case "Integer", "Double"
                        result = result & element
                    case "String"
                        result = result & """" & JsonEscape(element) & """"
                    case "Variant()", "Dictionary"
                        result = result & JsonSerialize(element)
                end select

                result = result & ", "
            next

            result = left(result, len(result) - 2)
            result = result & "]"
    end select

    JsonSerialize = result
end function

function JsonEscape(data)
    dim result
    result = data
    result = replace(result, "\", "\\")
    result = replace(result, vbcrlf, "\n")
    result = replace(result, vbtab, "\t")
    result = replace(result, """", "\""")
    JsonEscape = result
end function

sub Include(scriptName)
    ExecuteGlobal objFile.OpenTextFile(scriptName).ReadAll()
End Sub

sub Breakpoint(message)
    if typename(message) = "Variant()" then
        message = join(message, vbcrlf)
    end if

    wscript.Echo(message)
    wscript.Quit()
end sub

Main()