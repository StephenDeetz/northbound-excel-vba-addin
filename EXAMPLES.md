# Pretty Print Examples

Every case below is drawn from the add-in's own test suite
(`PrettyPrintAnswers.txt`), so what you see here is exactly what Pretty
Print produces. Some formulas already read fine on one line -- Pretty
Print leaves those alone.

## Example 1

Before:
```
=IF(A1>0, SUM((A1:B1) + (C1:D1)), "Test,,")
```

After:
```
=IF(
    A1 > 0,
    SUM(
        (A1:B1) + (C1:D1)
    ),
    "Test,,"
)
```

## Example 2

Before:
```
=CONCAT("Hello ""World,""","End")
```

After:
```
=CONCAT("Hello ""World,""", "End")
```

## Example 3

Before:
```
=IF(TRIM([@[UTC Offset]])<>"", $B$2 + ([@[UTC Offset]]/24), "")
```

After:
```
=IF(
    TRIM([@[UTC Offset]]) <> "",
    $B$2 + ([@[UTC Offset]] / 24),
    ""
)
```

## Example 4

Before:
```
=IF(TRUE,IF(A1=1,B1+C1,D1),FALSE)
```

After:
```
=IF(TRUE, IF(A1 = 1, B1 + C1, D1), FALSE)
```

## Example 5

Before:
```
=LAMBDA(Char,CharStatus,Guess,GameArr,CharStatusArr,SUM(MAKEARRAY(1,5,LAMBDA(r,C,IF(AND(INDEX(GameArr,r,C)=Char,INDEX(CharStatusArr,r,C)<>"Not in Word"),1,0)))))(Game!D5,D5,$AP5,Game!$B5:$F5,Interpret!$B5:$F5)
```

After:
```
=LAMBDA(
    Char,
    CharStatus,
    Guess,
    GameArr,
    CharStatusArr,
    SUM(
        MAKEARRAY(
            1,
            5,
            LAMBDA(
                r,
                C,
                IF(
                    AND(
                        INDEX(GameArr, r, C) = Char,
                        INDEX(CharStatusArr, r, C) <> "Not in Word"
                    ),
                    1,
                    0
                )
            )
        )
    )
)(
    Game!D5,
    D5,
    $AP5,
    Game!$B5:$F5,
    Interpret!$B5:$F5
)
```

## Example 6

Before:
```
=IFERROR(IF($H18<>"",XLOOKUP($H18,'Client Forecast'!$K:$K,'Client Forecast'!C:C),""),"")
```

After:
```
=IFERROR(
    IF(
        $H18 <> "",
        XLOOKUP(
            $H18,
            'Client Forecast'!$K:$K,
            'Client Forecast'!C:C
        ),
        ""
    ),
    ""
)
```

## Example 7

Before:
```
=IF(TRIM($A2)<>"",INDEX(TEXTSPLIT($A2,"|"),1,3),"")
```

After:
```
=IF(
    TRIM($A2) <> "",
    INDEX(TEXTSPLIT($A2, "|"), 1, 3),
    ""
)
```

## Example 8

Before:
```
=@IF($AF2,IF($AJ2,IF($AP2>0,INDEX(CSV_Std_Translations,1,$AP2),"Blank"),IF(TRIM($C2)<>"",$C2,"Blank")),"")
```

After:
```
=@IF(
    $AF2,
    IF(
        $AJ2,
        IF(
            $AP2 > 0,
            INDEX(CSV_Std_Translations, 1, $AP2),
            "Blank"
        ),
        IF(TRIM($C2) <> "", $C2, "Blank")
    ),
    ""
)
```

## Example 9

Before:
```
=IF(IFERROR(FIND(" - ",INDEX(TEXTSPLIT($A2,"|"),1,2)),0)>=0,INDEX(TEXTSPLIT($A2,"|"),1,2),"")
```

After:
```
=IF(
    IFERROR(
        FIND(
            " - ",
            INDEX(TEXTSPLIT($A2, "|"), 1, 2)
        ),
        0
    ) >= 0,
    INDEX(TEXTSPLIT($A2, "|"), 1, 2),
    ""
)
```

## Example 10

Before:
```
=IF(A9,IF(K9<>0,TRIM(MID(B9,J9,FIND(K9,B9,J9)-J9)),MID(B9,J9,2)),"")
```

After:
```
=IF(
    A9,
    IF(
        K9 <> 0,
        TRIM(
            MID(B9, J9, FIND(K9, B9, J9) - J9)
        ),
        MID(B9, J9, 2)
    ),
    ""
)
```

## Example 11

Before:
```
=IF(A1<>"", IF(B1>0, IF(C1=1, "Valid", "Invalid"), "Error"), "Blank")
```

After:
```
=IF(
    A1 <> "",
    IF(B1 > 0, IF(C1 = 1, "Valid", "Invalid"), "Error"),
    "Blank"
)
```

## Example 12

Before:
```
=IF(A1 <> "", CONCAT("Prefix-", TEXTJOIN(", ", TRUE, B1, C1, D1)), "No Data")
```

After:
```
=IF(
    A1 <> "",
    CONCAT("Prefix-", TEXTJOIN(", ", TRUE, B1, C1, D1)),
    "No Data"
)
```

## Example 13

Before:
```
=SUM(XLOOKUP(A1, Table1[Lookup], Table1[Value], 0) * B1, C1)
```

After:
```
=SUM(XLOOKUP(A1, Table1[Lookup], Table1[Value], 0) * B1, C1)
```

## Example 14

Before:
```
=IF([@[Net Profit]] > 0, "Profitable", IF([@[Net Profit]] = 0, "Break Even", "Loss"))
```

After:
```
=IF(
    [@[Net Profit]] > 0,
    "Profitable",
    IF([@[Net Profit]] = 0, "Break Even", "Loss")
)
```

## Example 15

Before:
```
=INDIRECT(ADDRESS(A1,B1,4))
```

After:
```
=INDIRECT(ADDRESS(A1, B1, 4))
```

## Example 16

Before:
```
=SUM(SEQUENCE(5, 1, A1, 1) * B1)
```

After:
```
=SUM(SEQUENCE(5, 1, A1, 1) * B1)
```

## Example 17

Before:
```
=LET(x, A1 + B1, y, x * 2, y - C1)
```

After:
```
=LET(
    x, A1 + B1,
    y, x * 2,
    y - C1
)
```

## Example 18

Before:
```
=CHOOSE(A1, 10, 20, "Text Value", "Another Text")
```

After:
```
=CHOOSE(
    A1,
    10,
    20,
    "Text Value",
    "Another Text"
)
```

## Example 19

Before:
```
=IF(AND([@[% In]],[@[Votes In]]),[@[Rep Need]]/[@[VL-3VL]],0)
```

After:
```
=IF(
    AND([@[% In]], [@[Votes In]]),
    [@[Rep Need]] / [@[VL-3VL]],
    0
)
```

## Example 20

Before:
```
=LAMBDA(Input,Cnt,Pos,Incl,IFS((Cnt+Pos)<0,"[]",(Cnt+Pos)>=8,"[]",NOT(Incl),"["&MID(Input,(Pos+Cnt),(ABS(Cnt)))&"]",TRUE,"["&MID(Input,(Pos+Cnt+1),(ABS(Cnt)))&"]"))($N70,$P70,$Q70,$O70)
```

After:
```
=LAMBDA(
    Input,
    Cnt,
    Pos,
    Incl,
    IFS(
        (Cnt + Pos) < 0, "[]",
        (Cnt + Pos) >= 8, "[]",
        NOT(Incl), "[" & MID(
            Input,
            (Pos + Cnt),
            (ABS(Cnt))
        ) & "]",
        TRUE, "[" & MID(
            Input,
            (Pos + Cnt + 1),
            (ABS(Cnt))
        ) & "]"
    )
)($N70, $P70, $Q70, $O70)
```

## Example 21

Before:
```
=LET(x, ((A1+B1)), y, IF(x>0, LET(u, x*2, v, u+3, v), ((x))), y)
```

After:
```
=LET(
    x, ((A1 + B1)),
    y, IF(
        x > 0,
        LET(
            u, x * 2,
            v, u + 3,
            v
        ),
        ((x))
    ),
    y
)
```

## Example 22

Before:
```
=LET(x, (A1 + B1) * ((C1 + D1)), y, LET(u, x + 1, v, u * 3, v), y + ((x)))
```

After:
```
=LET(
    x, (A1 + B1) * ((C1 + D1)),
    y, LET(
        u, x + 1,
        v, u * 3,
        v
    ),
    y + ((x))
)
```

## Example 23

Before:
```
=LET(x, A1 + (B1 + (C1)), y, x * 2, y - ((C1)))
```

After:
```
=LET(
    x, A1 + (B1 + (C1)),
    y, x * 2,
    y - ((C1))
)
```

## Example 24

Before:
```
=IF([@Active]=FALSE,0,IFERROR(LET(Rid,[@Id],Keep,(tblRecipeBOM[Recipe Id]=Rid)*(1-((tblRecipeBOM[Item Type]="Recipe")*(tblRecipeBOM[Item Id]=Rid))),Qty,FILTER(tblRecipeBOM[Qty],Keep),Type,FILTER(tblRecipeBOM[Item Type],Keep),ID,FILTER(tblRecipeBOM[Item Id],Keep),val,SWITCH(Type,"Ingredient",XLOOKUP(ID,tblIngredients[Id],tblIngredients[Calories],0),"Recipe",XLOOKUP(ID,[Id],[Calories],0),0),ROUND(SUM(Qty*val),1)),""))
```

After:
```
=IF(
    [@Active] = FALSE,
    0,
    IFERROR(
        LET(
            Rid, [@Id],
            Keep, (tblRecipeBOM[Recipe Id] = Rid) * (
                1 - (
                    (tblRecipeBOM[Item Type] = "Recipe") * (tblRecipeBOM[Item Id] = Rid)
                )
            ),
            Qty, FILTER(tblRecipeBOM[Qty], Keep),
            Type, FILTER(tblRecipeBOM[Item Type], Keep),
            ID, FILTER(tblRecipeBOM[Item Id], Keep),
            val, SWITCH(
                Type,
                "Ingredient", XLOOKUP(ID, tblIngredients[Id], tblIngredients[Calories], 0),
                "Recipe", XLOOKUP(ID, [Id], [Calories], 0),
                0
            ),
            ROUND(SUM(Qty * val), 1)
        ),
        ""
    )
)
```

## Example 25

Before:
```
=IF(EOMONTH(K$2,0)<=EOMONTH('FTM P&L'!$O$2,0),XLOOKUP(EOMONTH(K$2,0),'Import Values from WB4'!$D$3300:$AM$3300,'Import Values from WB4'!$D3309:$AM3309,0,-1),0)
```

After:
```
=IF(
    EOMONTH(K$2, 0) <= EOMONTH('FTM P&L'!$O$2, 0),
    XLOOKUP(
        EOMONTH(K$2, 0),
        'Import Values from WB4'!$D$3300:$AM$3300,
        'Import Values from WB4'!$D3309:$AM3309,
        0,
        -1
    ),
    0
)
```

## Example 26

Before:
```
=SUMPRODUCT(--(tblData[Status]="Active"),--(tblData[Amount]>100),tblData[Value])
```

After:
```
=SUMPRODUCT(
    --(tblData[Status] = "Active"),
    --(tblData[Amount] > 100),
    tblData[Value]
)
```

## Example 27

Before:
```
=SUM([@[Revenue, Total]])
```

After:
```
=SUM([@[Revenue, Total]])
```

## Example 28

Before:
```
=SUM([@[Total (Net)]])
```

After:
```
=SUM([@[Total (Net)]])
```

## Example 29

Before:
```
=CHOOSE(MATCH(Data!E3,{"A","B","C"},0),"Alpha","Beta","Gamma")
```

After:
```
=CHOOSE(
    MATCH(Data!E3, {"A","B","C"}, 0),
    "Alpha",
    "Beta",
    "Gamma"
)
```

## Example 30

Before:
```
=LET(a, 1, b, 2, a + b)
```

After:
```
=LET(
    a, 1,
    b, 2,
    a + b
)
```

## Example 31

Before:
```
=IFS(A1 > 10, "High", A1 > 5, "Mid", TRUE, "Low")
```

After:
```
=IFS(
    A1 > 10, "High",
    A1 > 5, "Mid",
    TRUE, "Low"
)
```

## Example 32

Before:
```
=SWITCH(A1, 1, "One", 2, "Two", "Other")
```

After:
```
=SWITCH(
    A1,
    1, "One",
    2, "Two",
    "Other"
)
```
