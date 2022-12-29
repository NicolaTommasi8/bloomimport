clear all

discard

capture log close
qui log using guida/logs/example.txt, text replace
bloomimport using "data/Vantaggio competitivo e WACC.xlsx", cellrange(A4) ///
  sheet("Foglio1") datastart(B) nvar(32) lasttick(DEP)
summ
qui log close
filefilter guida/logs/example.txt guida/logs/example1.txt, replace from("\W. qui log close\W") to("")
erase guida/logs/example.txt


clear all
capture log close
qui log using guida/logs/example.txt, text replace
set maxvar 8000
bloomimport using "data/us_banks.xlsx", cellrange(A4) sheet("US1") ///
  datastart(B) nvar(13) lasttick(GQO)
summ
qui log close
filefilter guida/logs/example.txt guida/logs/example2.txt, replace from("\W. qui log close\W") to("")
erase guida/logs/example.txt





bloomimport using "data/testXbloomimport.xlsx", cellrange(A4) sheet("quotazioni") datastart(B) nvar(4) lasttick(ST) clear
summ

**con range di dati e un solo campo
bloomimport using "data/testXbloomimport.xlsx", cellrange(A4:Q1503) sheet("gua") datastart(B) nvar(1) lasttick(Q) clear
summ

**con range di dati e un solo campo
bloomimport using "data/testXbloomimport.xlsx", cellrange(T4:AA1503) sheet("gua") datastart(U) nvar(1) lasttick(AA) clear
summ

bloomimport using "data/testXbloomimport.xlsx", cellrange(A4) sheet("daily") datastart(B) nvar(6) lasttick(RB) clear
summ

bloomimport using "data/testXbloomimport.xlsx", cellrange(A4) sheet("daily2") datastart(B) nvar(4) lasttick(BJ) clear
summ
exit




gen data = date(date,"DMY")
format data %td
order data
drop date



gen year=real(substr(A,-4,.))
fre year
duplicates report ticker year
duplicates drop ticker year, force
fre year
drop A
order ticker year



