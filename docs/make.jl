using IanSmithdistmath
using Documenter

DocMeta.setdocmeta!(IanSmithdistmath, :DocTestSetup, :(using IanSmithdistmath); recursive=true)

makedocs(;
    modules=[IanSmithdistmath],
    authors="Douglas Bates <dmbates@gmail.com> and contributors",
    repo="https://github.com/dmbates/IanSmithdistmath.jl/blob/{commit}{path}#{line}",
    sitename="IanSmithdistmath.jl",
    format=Documenter.HTML(;
        prettyurls=get(ENV, "CI", "false") == "true",
        canonical="https://dmbates.github.io/IanSmithdistmath.jl",
        assets=String[],
    ),
    pages=[
        "Home" => "index.md",
    ],
)

deploydocs(;
    repo="github.com/dmbates/IanSmithdistmath.jl",
    devbranch="main",
)
