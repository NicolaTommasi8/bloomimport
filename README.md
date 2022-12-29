# bloomimport
Package to import some Bloomberg data exported via Excel

Example:

    bloomimport using "data/testXbloomimport.xlsx", cellrange(A4) sheet("daily") datastart(B) nvar(6) lasttick(RB) clear
    
    summ
    
        Variable |        Obs        Mean    Std. dev.       Min        Max
    -------------+---------------------------------------------------------
            date |          0
         PX_OPEN |    194,636    21.49725    131.4422      .1375    3384.22
         PX_LAST |    194,640    21.46353    131.1024      .1375   3296.123
     CUR_MKT_CAP |    198,071    7627.696    10532.53    97.7533   87087.78
         EQY_DPS |     29,889    .0820321    .2797269          0        2.1
    -------------+---------------------------------------------------------
    IS_DIV_PER~R |     29,889    .0820321    .2797269          0        2.1
    PX_TO_BOOK~O |    186,372    2.319075    4.859172       .009   110.5107
          ticker |          0

