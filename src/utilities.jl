remap(ln::AbstractString) = foldl(replace, REPLACEMENTS; init=ln)

# `is_dim` specifies whether we're in a `Dim` statement. If that's the case, provide
# a default value, otherwise including only the type.
# NOTE: Does not currently handle multiple declarations on the same line, which would
# be separated by commas.
function remap_declaration(spec::AbstractString, is_dim::Bool=startswith(spec, "Dim"))
    spec = spec[(is_dim ? 4 : 1):something(findfirst(==('#'), spec), end)]
    m = match(VAR_DEF_RGX, lstrip(spec))
    name, indices, as_type = m.captures
    if as_type === nothing
        is_dim && indices === nothing && return "local " * name
        type = "Any"
    else
        type = remap_type(strip(as_type[5:end]))
    end
    if indices !== nothing
        inds = split(strip(ispunct, indices), ',')
        n = length(inds)
        if n == 0
            type = "Vector{" * type * "}"
            value = type * "()"
        else
            # Array indices can be set as `lower To upper`, which overrides `Option Base`.
            # The default array base is 0, and this code ignores the possibility of the
            # VB code having set it to 1. To construct a Julia array, we need the size
            # in each dimension rather than the last index, hence the +1s below.
            sizes = map(inds) do ind
                if any(isspace, ind)
                    lower, _, upper = split(ind)
                    return parse(Int, upper) - parse(Int, lower) + 1
                else
                    return parse(Int, ind) + 1
                end
            end
            type = string("Array{", type, ',', n, "}")
            value = string(type, "(undef, ", join(sizes, ", "), ')')
        end
    else
        value = "convert(" * type * ", 0)"
    end
    out = name * "::" * type
    if is_dim
        out *= " = " * value
    end
    return out
end

function remap_type(vbt::AbstractString)
    return replace(
        vbt,
        "Boolean" => "Bool",
        "Byte" => "UInt8",
        "Double" => "Float64",
        "Integer" => "Int16",
        "Long" => "Int32",
        "LongLong" => "Int64",
        "LongPtr" => "Int",
        "Object" => "Any",  # close enough
    )
end

const VAR_DEF_RGX = r"(\w+)(\(.*\))?(\s+As\s+\w+)?"

const REPLACEMENTS = [
    "'" => "#",
    "Const " => "const ",
    r"(Private|Public) (Function|Sub)" => "function ",
    r"End (Function|Sub|If|Type)" => "end",
    "If " => "if ",
    " Then" => "",
    "Else" => "else",
    "ElseIf " => "elseif ",
    "Do While" => "while",
    "Loop" => "end",
    "Exit Do" => "break",
    # TODO: `Exit Function` is `return` but the thing to return is a variable with
    # the same name as the function which we don't have access to in this context,
    # so we can't map the statements 1-1
    "While " => "while ",
    "Wend" => "end",
    "Type" => "mutable struct",
    " Not " => " !",
    " And " => " && ",
    " Or " => " || ",
    "False" => "false",
    "True" => "true",
    " Null" => "nothing",
    "IsNull" => "isnothing",
    "ByVal " => "",
    "Abs(" => "abs(",
    "Exp(" => "exp(",
    "Log(" => "log(",
    "Max(" => "max(",
    "Min(" => "min(",
    "Sqr(" => "abs2(",
    r"([0-9])#" => s"\1.0",
    Regex("\\bDim\\s+" * VAR_DEF_RGX.pattern) => remap_declaration,
    r" As \w+" => (x -> "::" * remap_type(last(split(x)))),
]
