*! version 1.5  Nicola Tommasi  14sep2024
*   -dates(single|multi)
*   -minor changes
*! version 1.4  Nicola Tommasi  09mar2024
*   -export(wide|long)
*! version 1.2  Nicola Tommasi  09mar2024
*   -datastart() optional. If not specified, datastart=cellrange+1
*! version 1.1  Nicola Tommasi  07mar2023
*   -prevent xframeappend error "shared variables in frames being combined must be both numeric or both string"
*! version 1.0b  Nicola Tommasi  29nov2022
*   -version 17 use frames, 15 & 16 tempfile
*! version 0.5   Nicola Tommasi  17nov2022

program bloomimport
**version 15
set tracedepth 1
**only if version >= 17
**which name_ado_file

syntax  using/, sheet(string) cellrange(string) [export(string) nvar(integer 0) lasttick(string) /*datastart(string)*/ from(string) to(string) dates(string) debug ]

tempname temp fr_fusion ABS ABE
tempfile buildingDB
capture frames reset `temp'
capture frames reset `fr_fusion'

local version `c(stata_version)'
if `version'>=17 {
  capture which xframeappend
  if _rc==111 {
    di in yellow "xframeappend not installed.... installing..."
    ssc inst xframeappend
    di in yellow "xframeappend has been correctly installed!"
  }
}

if "`export'"=="" local export wide
if "`dates'"=="" local dates single /*multi*/
if "`dates'" != "single" & "`dates'" != "multi" {
  di "dates() must be single or mutli"
  exit
}


if "`export'"=="wide" {
  qui import excel using "`using'", sheet("`sheet'") cellrange(`cellrange') clear allstring

  if regexm("`cellrange'","(^[A-Z]*)") local firstrow = regexs(1)
  mata: `ABS' = "`firstrow'"
  mata: st_numscalar("datastartN", numofbase26(`ABS'))
  if "`dates'" == "single" local  datastartN = datastartN+1
  else  local  datastartN = datastartN
  if regexm("`lasttick'","(^[A-Z]*)") local lasttickS=  regexs(1)
  mata: `ABE' = "`lasttickS'"
  mata: st_numscalar("lasttickmata", numofbase26(`ABE'))
  local lasttickmata = lasttickmata
  local sta = `datastartN' /*colonna da cui partono i dati per dates(single), servono per mata */
  local STA=`sta'
  local LAST = `lasttickmata' /* è la colonna dove iniziano i dati dell'ultimo ticker --> numofbase26() */

  if "`dates'" == "single" {
    while `STA'<= `LAST'  {
      local END = `STA' + `nvar' - 1
      mata: st_local("cellname", numtobase26(`STA'))
      mata: st_local("cellfine", numtobase26(`STA'+`nvar'-1))
      local Vname = `cellname' in 1

      preserve
        keep `firstrow' `cellname'-`cellfine'
        local VtoDESTR = ""
        forvalues c=`STA'/`END' {
  	      mata: st_local("clnm", numtobase26(`c'))
  	      local token =`clnm' in 2
          capture confirm variable `token'
          if _rc rename `clnm' `token'
          else {
            local token `token'2
            rename `clnm' `token'
          }
          qui replace `token'="" if strmatch(`token',"*N/A*")
          local VtoDESTR  "`VtoDESTR' `token'"
        }
        qui gen ticker="`Vname'"
        qui drop in 1/2
        rename `firstrow' date

        if `version'>=17 {
          if `STA'==`sta' qui frame copy default `fr_fusion', replace
          else {
            qui frame copy default `temp'
            if "`debug'"!=""{
              di "temp"
              summ
              desc
              fre ticker
            }
            frame change `fr_fusion'
            if "`debug'"!=""{
              di "fr_fusion"
              desc
              summ
            }
            xframeappend `temp', drop fast
            frame change default
          }
        }
        else {
          if `STA'==`sta' qui save `buildingDB', replace
          else {
            qui append using `buildingDB', force
            qui save `buildingDB', replace
          }
        }
      restore
      if `STA'<=`LAST' local STA = `STA' + `nvar'
      }
  }

  else { /*dates=multi*/
    **set trace on
    local STAi = `STA'
    while `STAi'<= `LAST'  {
      local ENDi = `STAi' + `nvar'
      mata: st_local("cellname", numtobase26(`STAi'))
      mata: st_local("cellfine", numtobase26(`ENDi'))
      local Vname = `cellname' in 1
      preserve
        keep `cellname'-`cellfine'
        local VtoDESTR = ""
        forvalues c=`STAi'/`ENDi' {
  	      mata: st_local("clnm", numtobase26(`c'))
  	      if `c'==`STAi' local token date
          else local token = `clnm' in 2
          capture confirm variable `token'
          if _rc rename `clnm' `token'
          else {
            local token `token'2
            rename `clnm' `token'
          }
          qui replace `token'="" if strmatch(`token',"*N/A*")
          local VtoDESTR  "`VtoDESTR' `token'"
        }
        qui gen ticker="`Vname'"
        qui drop in 1/2
        **qui destring `VtoDESTR', replace
        qui drop if date==""
        if `version'>=17 {
          if `STAi'==`sta' qui frame copy default `fr_fusion', replace
          else {
            qui frame copy default `temp'
            if "`debug'"!=""{
              di "temp"
              summ
              desc
              fre ticker
            }
            frame change `fr_fusion'
            if "`debug'"!=""{
              di "fr_fusion"
              desc
              summ
            }
            xframeappend `temp', drop fast
            frame change default
          }
        }
        else {
          if `STAi'==`sta' qui save `buildingDB', replace
          else {
            qui append using `buildingDB', force
            qui save `buildingDB', replace
          }
        }
      restore
      if `STAi'<=`LAST' local STAi = `ENDi' + 2  /*empy col */
    }
  }

  if `version'>=17 {
    frame change `fr_fusion'
    frame copy `fr_fusion' default, replace
    frame change default
    qui destring, replace
  }
  else {
    use  `buildingDB', clear
    qui destring, replace
  }
  order ticker date
}



else { /*"`export'"=="long"*/
  if "`dates'" == "single" {
    qui import excel using "`using'", sheet("`sheet'") cellrange(`cellrange') clear allstring firstrow case(upper)
    qui missings dropvars, force

    if regexm("`cellrange'","(^[A-Z]*)") local firstrow = regexs(1)

    rename `firstrow' ticker
    label var ticker "Ticker"
    qui carryforward ticker, replace
    rename DATES field
    label var field "Field"

    qui ds, not(varl Ticker Field)
    foreach V of varlist `r(varlist)' {
      local vdesc : variable label `V'
      local date = subinstr("`vdesc'","/","_",.)
      local date = subinstr("`date'"," ","",.)
      rename `V' _`date'
      qui replace _`date'="" if strmatch(_`date',"#*")
    }

    if `version'>=18 {
      qui reshape long _@, i(ticker field) j(date) favor(speed) string
      qui reshape wide _, i(ticker date) j(field) string favor(speed)
    }
    else {
      qui reshape long _@, i(ticker field) j(date) string
      qui reshape wide _, i(ticker date) j(field) string
    }

    foreach V of varlist _* {
      qui destring `V', replace
    }
    rename _* *
  }

  else { /*"`dates'" == "multi"*/
    qui import excel using "`using'", sheet("`sheet'") cellrange(`cellrange') clear allstring case(upper)
    if regexm("`cellrange'","(^[A-Z]*)") local firstrow = regexs(1)
    rename `firstrow' ticker
    label var ticker "Ticker"
    qui carryforward ticker, replace
    qui ds
    local field = word("`r(varlist)'",2)
    rename `field' field
    label var field "Field"
    qui drop if field==""
    local tag = field in 1
    tempname ticker_id field_id id fieldlen
    qui gen `id'=_n
    qui egen `ticker_id' = seq() if field=="`tag'"
    qui carryforward `ticker_id', replace
    qui bysort `ticker_id' (`id'): gen `field_id'=_n
    label var `id' "tempvar"
    label var `ticker_id' "tempvar"
    label var `field_id' "tempvar"

    qui levelsof `ticker_id', local(ticker_list)
    foreach t of local ticker_list {
      preserve
        qui keep if `ticker_id'==`t'
        qui missings dropvars, force
        qui ds, not(varl Ticker Field tempvar)
        foreach V of varlist `r(varlist)' {
          local vname = `V' in 1
          local vname = strtrim("`vname'")
          local vname = subinstr("`vname'","/","",.)
          rename `V' D`vname'
        }
        qui drop in 1
        qui reshape long D@, i(ticker field) j(date) string
        qui replace D="" if strmatch(D,"*N/A*")
        qui gen `fieldlen'=strlen(field)
        qui summ `fieldlen'
        if `r(max)'>=32 local flag=1
        else local flag=0
        if `flag'==1 {
          qui levelsof field if `fieldlen'>=32, local(lista)
          local cnt=1
          foreach F of local lista {
            qui replace field = "tmp`cnt'" if field=="`F'"
            local name`cnt' `F'
            local cnt `++cnt'
          }
        }
        capture drop `id' `field_id' `fieldlen'
        qui reshape wide D, i(ticker date) j(field) string
        if `flag'==1 {
          local cnt `--cnt'
          forvalues i=1/`cnt' {
            rename Dtmp`i' `name`i''
          }
        }

        if `version'>=17 {
          if `ticker_id'==1 qui frame copy default `fr_fusion', replace
          else {
            qui frame copy default `temp'
            if "`debug'"!=""{
              di "temp"
              summ
              desc
              fre ticker
            }
            frame change `fr_fusion'
            if "`debug'"!=""{
              di "fr_fusion"
              desc
              summ
            }
            xframeappend `temp', drop fast
            frame change default
          }
        }
        else {
          if `ticker_id'==1 qui save `buildingDB', replace
          else {
            qui append using `buildingDB', force
            qui save `buildingDB', replace
          }
        }
      restore
    }
    **to avoid dates destring
    if `version'>=17 {
      frame change `fr_fusion'
      frame copy `fr_fusion' default, replace
      frame change default
      qui ds
      local VARLIST = "`r(varlist)'"
      local VARLIST = subinstr("`VARLIST'"," date ", " ",1)
      qui destring `VARLIST', replace
    }
    else {
      use  `buildingDB', clear
      qui ds
      local VARLIST = "`r(varlist)'"
      local VARLIST = subinstr("`VARLIST'"," date ", " ",1)
      qui destring `VARLIST', replace
    }
  }
}

end


/****
William Matsuoka
Putexcel Part II: numofbase26()
http://www.wmatsuoka.com/stata/putexcel-part-ii-numofbase26
****/
mata
real matrix numofbase26(string matrix base)
{
  real matrix output, pwr, b
  real scalar i, j, k, l

  base = strupper(base)
  output = J(rows(base), cols(base), .)

  for (i=1; i<=rows(base); i++) {
    for (j=1; j<=cols(base); j++) {
      if (strlen(base[i,j]) == 1) output[i,j] = ascii(base[i,j]) - 64
      else {
        l = strlen(base[i,j])
        b = pwr = J(1, l, .)
        for (k=1; k<=l; k++) {
          b[1,k] = ascii(substr(base[i,j], k, 1)) - 64
          pwr[1, k] = l - k
        }
        output[i,j] = rowsum(b :* (26:^pwr))
      }
    }
  }
  return(output)
}
end



/***********
  numtobase26() is an undocumented Mata function
   that converts column numbers to Excel's column letters. For example, if
   you want to know what the 27th column will be called in Excel you can
  type in Stata:
  . mata : numtobase26(27)
  AA
*********/


exit


