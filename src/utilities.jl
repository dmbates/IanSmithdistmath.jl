function remap(ln::AbstractString)
    return replace(
        ln,
        "Const " => "const ",
        "Private Function" => "function ",
        "Public Function " => "function ",
        "End Function" => "end",
        "If " => "if ",
        "End If" => "end",
        " Then" => "",
        "Else" => "else",
        "ElseIf " => "elseif ",
        "Do While" => "while",
        "Loop" => "end",
        "While " => "while ",
        "Wend" => "end",
        " Not " => " !",
        " And " => " && ",
        " Or " => " || ",
        "False" => "false",
        "True" => "true",
        "ByVal " => "",
        " As Boolean" => "::Bool",
        " As Double" => "::Float64",
        "Abs(" => "abs(",
        "Exp(" => "exp(",
        "Log(" => "log(",
        "Max(" => "max(",
        "Min(" => "min(",
        "Sqr(" => "abs2(",
        r"([0-9])#" => s"\1.0",
        "'" => "#",
    )
end