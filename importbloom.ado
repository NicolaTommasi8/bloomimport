*! version 0.5  Nicola Tommasi  17nov2022

program importbloom
version 17

set tracedepth 1

syntax using/  , cellrange(string) datastart(string) nvar(integer) lasttick(string)   ///
       [sheet(string) clear  from(string) to(string) ///
        debug /*undocumented*/ ]


tempname temp fr_fusion ABS
capture frames reset `temp'
capture frames reset `fr_fusion'

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


**if regexm("`datastart'","([0-9]*$)") local datastartN = regexs(1)
if regexm("`datastart'","(^[A-Z]*)") local datastartS=  regexs(1)
mata: `ABS' = "`datastartS'"
mata: st_numscalar("datastartN", numofbase26(`ABS'))
local  datastartN = datastartN

**if regexm("`lasttick'","([0-9]*$)") local lasttickN = regexs(1)
if regexm("`lasttick'","(^[A-Z]*)") local lasttickS=  regexs(1)
mata: `ABS' = "`lasttickS'"
mata: st_numscalar("lasttickmata", numofbase26(`ABS'))
local lasttickmata = lasttickmata

local sta = `datastartN' /*colonna da cui partono i dati, servono per mata */
**local nvar = 32 /* numero di vars che si ripetono, serve per mata */
local STA=`sta'
local LAST = `lasttickmata' /* è la colonna dove iniziano i dati dell'ultimo ticker --> numofbase26() */


while `STA'<= `LAST'  {
	**di `STA'
  local END = `STA' + `nvar' - 1

  mata: st_local("cellname", numtobase26(`STA'))
  mata: st_local("cellfine", numtobase26(`STA'+`nvar'-1))
  **di "inizio: `cellname'"
  **di "fine: `cellfine'"

  local Vname = `cellname' in 1
  **local Vname = subinstr("`Vname'"," Equity","",1)
  **di "Vname: `Vname'"


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
  restore

  if `STA'<=`LAST' local STA = `STA' + `nvar'
}

frame change `fr_fusion'
frame copy `fr_fusion' default, replace


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


