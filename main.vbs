option explicit
dim objShell, objFile, objHttp
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
set objHttp = CreateObject("Msxml2.XMLHTTP.6.0")
dim directory
directory = objFile.GetParentFolderName(wscript.ScriptFullName)

sub Main()
    dim headers, test, test2
    set headers = CreateObject("Scripting.Dictionary")
    call headers.Add("Content-Type", "application/json")
    set test = CreateObject("Scripting.Dictionary")
    call test.Add("role", "user")
    call test.Add("content", "Hello, world!")
    test2 = array("foo", "bar")
    call test.Add("list", 0.3)
    Breakpoint(Json(0, test))
end sub

function Json(method, data)
dim element, result, dataKeys
result = ""

    select case method
        case 0 'Serialize
            select case typename(data)
                case "Dictionary"
                    result = result & "{"
                    dataKeys = data.Keys()

                    for each element in dataKeys
                        result = result & """" & element & """: "

                        select case typename(data.Item(element))
                            case "Boolean"
                                result = result & lcase(data.Item(element))
                            case "Integer"
                                result = result & data.Item(element)
                            case "Double"
                                result = result & data.Item(element)
                            case "String"
                                result = result & """" & data.Item(element) & """"
                            case "Variant()"
                                result = result & Json(0, data.Item(element))
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
                            case "Integer"
                                result = result & element
                            case "Double"
                                result = result & element
                            case "String"
                                result = result & """" & element & """"
                            case "Variant()"
                                result = result & Json(0, element)
                        end select

                        result = result & ", "
                    next

                    result = left(result, len(result) - 2)
                    result = result & "]"
            end select
    end select

Json = result
end function

function Curl(url, method, headers, data)
    dim element, headerKeys
    call objHtml.Open(method, url, false)
    headerKeys = headers.Keys()
    
    for each element in headerKeys
        call objHtml.SetRequestHeader(element, headers.Item(element))
    next


end function

sub Breakpoint(message)
    if typename(message) = "Variant()" then
        message = join(message, vbcrlf)
    end if

    wscript.Echo(message)
    wscript.Quit()
end sub

Main()