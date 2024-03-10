{smcl}
{* *! version 1.2  09mar2024}{...}
{p2colset 1 21 18 2}{...}
{p2col:{bf: bloomimport} {hline 2}}Import Bloomberg data{p_end}
{p2colreset}{...}


{marker syntax}{...}
{title:Syntax}


{p 8 32 2}
{cmd:bloomimport} {cmd:using} {it:excel_filename},
      {it:{help bloomimport##bloomimport_options:bloomimport_options}}


{synoptset 37}{...}
{marker bloomimport_options}{...}
{synopthdr :bloomimport_options}
{synoptline}
{synopt :{opt sh:eet("sheetname")}} Excel worksheet to load (mandatory){p_end}
{synopt :{opt cellra:nge([start][:end])}} Excel cell range to load (mandatory){p_end}
{synopt :{opt nvar(#)}} Number of variables for each ticker (mandatory){p_end}
{synopt :{opt lasttick(string)}} Excel column of last ticker (mandatory){p_end}
{synopt :{opt datastart(column)}} Excel cell where data start{p_end}
{synopt :{opt from(varlist)}} data ticker to rename{p_end}
{synopt :{opt to(varlist)}} new names for data ticker specified in {opt from(varlist)}{p_end}
{synopt :{opt clear}} replace data in memory{p_end}
{synoptline}
{p2colreset}{...}


{marker description}{...}
{title:Description}

{pstd}
{cmd:bloomimport} loads an Excel file containing data exported by Bloomberg into Stata.

{marker importoptions}{...}
{title:Options for bloomimport}

{phang}
{opt "cellrange([start][:end])"} specifies a range of cells within
the worksheet to load. {it:start} and {it:end} are specified using
standard Excel cell notation, for example, {cmd:A1}, {cmd:BC2000}, and
{cmd:C23}.

{phang}
{cmd:sheet("}{it:sheetname}{cmd:")} imports the worksheet named
{it:sheetname} in the workbook.  The default is to import the first
worksheet.

{phang}
{cmd:nvar(integer)} specifies the number of data item foreach ticker.

{phang}
{cmd:lastick(column)} specifies the column of the last ticker.

{phang}
{cmd:datastart(column)} specifies the first column of data in the Excel worksheet.
If not specified, next column with respect to cellrange()

{phang}
{cmd:from(varlist)} list of fields that need to be renamed, typically because the name is too long or
because the name would be incompatible with Stata's variable naming rules.

{phang}
{cmd:to(varlist)} list of new names to assign to the fields specified in from(varlist). (Under development, coming soon)

{phang}
{cmd:clear} clear data in memory before loading data from the Excel workbook. (Under development, coming soon)

{pstd}

{marker examples}{...}
{title:Examples}

{cmd:bloomimport using "data/Vantaggio competitivo e WACC.xlsx", cellrange(A4) sheet("Foglio1") nvar(32) lasttick(DEP)}
{pstd}

{cmd:bloomimport using "data/us_banks.xlsx", cellrange(A4) sheet("US1") datastart(B) nvar(13) lasttick(GQO) }
{pstd}

{cmd:bloomimport using "data/ISIN.xlsx", cellrange(A4) sheet("estrazione1") datastart(B) nvar(6) lasttick(HOF) clear}
{cmd:compress}
{cmd:order ticker date}
{cmd:gen tmp=date(date,"MDY")}
{cmd:order tmp, after(date)}
{cmd:format tmp %td}
{cmd:drop date}
{cmd:rename tmp date}
{cmd:export excel using ISIS_longform.xlsx, firstrow(variables) replace}


{marker remarks}{...}
{title:Remarks}
The command will automatically download and install the xframeappend package {helpb xframeappend}, if it is not already installed.
See {net "describe xframeappend, from(http://fmwww.bc.edu/repec/bocode/x/)":ssc describe xframeappend}.

{cmd:bloomimport} includes a piece of code written by William Matsuoka to convert Excel column letter to
their corresponding column numbers. Find the article and code {browse "http://www.wmatsuoka.com/stata/putexcel-part-ii-numofbase26":here}.







