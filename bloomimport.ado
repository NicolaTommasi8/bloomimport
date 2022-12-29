*! version 1.0b  Nicola Tommasi  29nov2022
*   -version 17 use frames, 15 & 16 tempfile
*! version 0.5   Nicola Tommasi  17nov2022

program bloomimport
version 17
**set tracedepth 1
**only if version >= 17
**which name_ado_file

syntax using/  , cellrange(string) datastart(string) nvar(integer) lasttick(string)   ///
       [sheet(string) clear  from(string) to(string) ///
        debug /*undocumented*/ ]

tempname temp fr_fusion ABS
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

qui import excel using "`using'", sheet("`sheet'") cellrange(`cellrange') `clear'


/***********
numtobase26() is an undocumented Mata function
 that converts column numbers to Excel's column letters. For example, if
 you want to know what the 27th column will be called in Excel you can
type in Stata:
. mata : numtobase26(27)
AA
*********/

if regexm("`cellrange'","(^[A-Z]*)") local firstrow=  regexs(1)

if regexm("`datastart'","(^[A-Z]*)") local datastartS=  regexs(1)
mata: `ABS' = "`datastartS'"
mata: st_numscalar("datastartN", numofbase26(`ABS'))
local  datastartN = datastartN

if regexm("`lasttick'","(^[A-Z]*)") local lasttickS=  regexs(1)
mata: `ABS' = "`lasttickS'"
mata: st_numscalar("lasttickmata", numofbase26(`ABS'))
local lasttickmata = lasttickmata

local sta = `datastartN' /*colonna da cui partono i dati, servono per mata */
local STA=`sta'
local LAST = `lasttickmata' /* è la colonna dove iniziano i dati dell'ultimo ticker --> numofbase26() */


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
	    rename `clnm' `token'
      qui replace `token'="" if strmatch(`token',"*N/A*")
      local VtoDESTR  "`VtoDESTR' `token'"
    }

    qui gen ticker="`Vname'"
    qui drop in 1/2
    qui destring `VtoDESTR', replace
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
        append using `buildingDB', force
        qui save `buildingDB', replace
      }
    }
  restore

  if `STA'<=`LAST' local STA = `STA' + `nvar'
}

if `version'>=17 {
  frame change `fr_fusion'
  frame copy `fr_fusion' default, replace
}
else use  `buildingDB', clear

end


/***
Putexcel Part II: numofbase26()
http://www.wmatsuoka.com/stata/putexcel-part-ii-numofbase26
***/
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




exit


