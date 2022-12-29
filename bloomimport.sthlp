{smcl}
{* *! version 1.2.6  17mar2021}{...}
{p2colset 1 21 18 2}{...}
{p2col:{bf: bloomimport} {hline 2}}Import Bloomberg data{p_end}
{p2colreset}{...}


{marker syntax}{...}
{title:Syntax}


{p 8 32 2}
{cmd:bloomimport} {cmd:using} {it:{help filename}}
    [{cmd:,}
      {it:{help bloomimport##bloomimport_options:bloomimport options}}]


{synoptset 37}{...}
{marker bloomimport_options}{...}
{synopthdr :bloomimport options}
{synoptline}
{synopt :{opt sh:eet("sheetname")}} Excel worksheet to load{p_end}
{synopt :{opt cellra:nge([start][:end])}} Excel cell range to load{p_end}
{synopt :{opt datastart(column)}} Excel cell where data start{p_end}
{synopt :{opt nvar(#)}} Number of variables for each ticker{p_end}
{synopt :{opt lasttick(string)}} Excel column of last ticker{p_end}
{synopt :{opt from(varlist)}} data ticker to rename{p_end}
{synopt :{opt to(varlist)}} new names for data ticker specified in {opt from(varlist)}{p_end}
{synopt :{opt clear}}replace data in memory{p_end}
{synoptline}
{p2colreset}{...}


{marker description}{...}
{title:Description}

{pstd}
{cmd:bloomimport} loads an Excel file containing data exported by bloomberg into Stata.

{marker importoptions}{...}
{title:Options for bloomimport}

{phang}
{opt "cellrange([start][:end])"} specifies a range of cells within
the worksheet to load.  {it:start} and {it:end} are specified using
standard Excel cell notation, for example, {cmd:A1}, {cmd:BC2000}, and
{cmd:C23}.

{phang}
{cmd:sheet("}{it:sheetname}{cmd:")} imports the worksheet named
{it:sheetname} in the workbook.  The default is to import the first
worksheet.

{phang}
{cmd:datastart(column)} specifies the first column of data in the Excel worksheet.
Quindi è la colonna dopo la data, dove iniziano i dati. Forse non serve

{phang}
{cmd:nvar(integer)} specifies the number of data item foreach ticker.

{phang}
{cmd:lastick(column)} specifies the column dove inizia l'ultimo ticker.

{phang}
{cmd:from(varlist)} lista dei campi che devono essere rinominati, tipicamente perchè
il nome è troppo lungo o perchè il nome sarebbe incompatibile con le regole di Stata sui
nomi delle variabili (è ancora veder bene come farlo e dove piazzarlo)

{phang}
{cmd:to(varlist)} lista dei nuovi nomi da assegnare ai campi specificati in {cmd:from(varlist)}

{phang}
{cmd:clear} clears data in memory before loading data from the Excel workbook.

{marker remarks}{...}
{title:Remarks/Examples}

{pstd}

        {cmd:. bloomimport using "data/Vantaggio competitivo e WACC.xlsx", cellrange(A4) sheet("Foglio1") datastart(B) nvar(32) lasttick(DEP)}

{pstd}

        {cmd:. bloomimport using "data/us_banks.xlsx", cellrange(A4) sheet("US1") datastart(B) nvar(13) lasttick(GQO) }
