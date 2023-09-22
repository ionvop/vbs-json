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

    Debug()

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

function JsonParse(data)
    dim i, element, element2, content, isString, level, result, key, value, list
    i = 1
    content = replace(data, vbcrlf, "")
    content = replace(content, vbtab, "")
    content = trim(content)
    isString = false
    level = 0

    select case left(content, 1)
        case "{"
            set result = CreateObject("Scripting.Dictionary")
            list = array()
            value = ""

            for i = 2 to len(content)
                element = mid(content, i, 1)

                select case element
                    case """"
                        if isString then
                            isString = false
                        else
                            isString = true
                        end if

                        value = value & element
                    case "\"
                        key = key & mid(content, i, 2)
                        i = i + 1
                    case "{", "["
                        if isString = false then
                            level = level + 1
                        end if

                        value = value & element
                    case "}", "]"
                        if isString = false then
                            if level <= 0 then
                                level = 0
                                value = trim(value)
                                Push list, value
                                value = ""
                            else
                                value = value & element
                            end if

                            level = level - 1
                        else
                            value = value & element
                        end if
                    case ","
                        if isString = false then
                            if level <= 0 then
                                value = trim(value)
                                Push list, value
                                value = ""
                            else
                                value = value & element
                            end if
                        else
                            value = value & element
                        end if
                    case else
                        value = value & element
                end select
            next

            for each element in list
                key = ""

                for i = 1 to len(element)
                    element2 = mid(element, i, 1)

                    select case element2
                        case """"
                            if isString then
                                isString = false
                            else
                                isString = true
                            end if

                            key = key & element2
                        case "\"
                            key = key & mid(element, i, 2)
                            i = i + 1
                        case ":"
                            if isString = false then
                                exit for
                            end if

                            key = key & element2
                        case else
                            key = key & element2
                    end select
                next

                key = trim(key)

                if left(key, 1) = """" then
                    key = mid(key, 2, len(key) - 2)
                end if

                value = ""

                for i = i + 1 to len(element)
                    element2 = mid(element, i, 1)

                    select case element2
                        case """"
                            if isString then
                                isString = false
                            else
                                isString = true
                            end if

                            value = value & element2
                        case "\"
                            value = value & mid(element, i, 2)
                            i = i + 1
                        case "{", "["
                            if isString = false then
                                level = level + 1
                            end if

                            value = value & element2
                        case "}", "]"
                            if isString = false then
                                if level <= 0 then
                                    level = 0
                                    exit for
                                end if

                                level = level - 1
                            end if

                            value = value & element2
                        case ","
                            if isString = false then
                                if level <= 0 then
                                    exit for
                                end if
                            end if

                            value = value & element2
                        case else
                            value = value & element2
                    end select
                next

                value = trim(value)

                select case left(value, 1)            
                    case """"
                        value = mid(value, 2, len(value) - 2)
                    case "{"
                        set value = CloneDict(JsonParse(value))
                    case "["
                        value = JsonParse(value)
                    case else
                        if isnumeric(value) then
                            if instr(value, ".") then
                                value = cdbl(value)
                            else
                                value = clng(value)
                            end if
                        end if
                end select

                if typename(value) = "String" then
                    value = JsonUnescape(value)
                end if
                
                result.Add key, value
            next
        case "["
            result = array()
            value = ""

            for i = 2 to len(content)
                element = mid(content, i, 1)

                select case element
                    case """"
                        if isString then
                            isString = false
                        else
                            isString = true
                        end if

                        value = value & element
                    case "\"
                        value = value & mid(content, i, 2)
                        i = i + 1
                    case "{", "["
                        if isString = false then
                            level = level + 1
                        end if

                        value = value & element
                    case "}", "]"
                        if isString = false then
                            if level <= 0 then
                                level = 0
                                value = trim(value)

                                select case left(value, 1)            
                                    case """"
                                        value = mid(value, 2, len(value) - 2)
                                    case "{"
                                        set value = CloneDict(JsonParse(value))
                                    case "["
                                        value = JsonParse(value)
                                    case else
                                        if isnumeric(value) then
                                            if instr(value, ".") then
                                                value = cdbl(value)
                                            else
                                                value = clng(value)
                                            end if
                                        end if
                                end select

                                Push result, value
                                value = ""
                            else
                                value = value & element
                            end if

                            level = level - 1
                        else
                            value = value & element
                        end if
                    case ","
                        if isString = false then
                            if level <= 0 then
                                value = trim(value)

                                select case left(value, 1)            
                                    case """"
                                        value = mid(value, 2, len(value) - 2)
                                    case "{"
                                        set value = CloneDict(JsonParse(value))
                                    case "["
                                        value = JsonParse(value)
                                    case else
                                        if isnumeric(value) then
                                            if instr(value, ".") then
                                                value = cdbl(value)
                                            else
                                                value = clng(value)
                                            end if
                                        end if
                                end select

                                Push result, value
                                value = ""
                            else
                                value = value & element
                            end if
                        else
                            value = value & element
                        end if
                    case else
                        value = value & element
                end select
            next
    end select

    if typename(result) = "Dictionary" then
        set JsonParse = result
        exit function
    end if
    
    JsonParse = result
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
                    case "Integer", "Double", "Long"
                        result = result & data.Item(element)
                    case "String"
                        result = result & """" & JsonEscape(data.Item(element)) & """"
                    case "Variant()", "Dictionary"
                        result = result & JsonSerialize(data.Item(element))
                end select

                result = result & ", "
            next

            if len(result) > 2 then
                result = left(result, len(result) - 2)
            end if

            result = result & "}"
        case "Variant()"
            result = result & "["

            for each element in data
                select case typename(element)
                    case "Boolean"
                        result = result & lcase(element)
                    case "Integer", "Double", "Long"
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

function JsonUnescape(data)
    dim i, element, result
    result = ""

    for i = 1 to len(data)
        element = mid(data, i, 1)

        if element = "\" then
            i = i + 1
            element = mid(data, i, 1)

            select case element
                case "\"
                    result = result & "\"
                case "n"
                    result = result & vbcrlf
                case "t"
                    result = result & vbcrlf
                case """"
                    result = result & """"
            end select
        else
            result = result & element
        end if
    next

    JsonUnescape = result
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

sub Debug()
    dim test, test2
    test = objFile.OpenTextFile(directory & "\res.json").ReadAll()
    set test2 = JsonParse(test)
    Breakpoint(JsonSerialize(test2))
end sub

Main()