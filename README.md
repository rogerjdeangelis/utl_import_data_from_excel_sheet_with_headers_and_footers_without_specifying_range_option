# utl_import_data_from_excel_sheet_with_headers_and_footers_without_specifying_range_option
Import data from excel sheet with headers and footers without specifying RANGE option.  Keywords: sas sql join merge big data analytics macros oracle teradata mysql sas communities stackoverflow statistics artificial inteligence AI Python R Java Javascript WPS Matlab SPSS Scala Perl C C# Excel MS Access JSON graphics maps NLP natural language processing machine learning igraph DOSUBL DOW loop stackoverflow SAS community.


    Import data from excel sheet with headers and footers without specifying RANGE option

    NOTE: You can progamatically create, read and write to ranges in R package XLConnect.
          This package has very strong support.

    https://tinyurl.com/ya3ev7ms
    https://github.com/rogerjdeangelis/utl_import_data_from_excel_sheet_with_headers_and_footers_without_specifying_range_option

    To answer the rest of the question see
    https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel&type=&language=

    see SAS Forum
    https://tinyurl.com/y7cqkho9
    https://communities.sas.com/t5/Base-SAS-Programming/How-to-import-data-from-excel-xlsx-sheet-with-headers-and/td-p/246864


    INPUT
    =====

     WORKBOOK d:/xls/class.xlsx with sheet class (you can use [sheet1])

       d:/xls/class.xlsx
          +-------------------------+----------------------------------------------------------------+
          |    A       |    B       |     C      |    D       |     E      |    F       |    g       |
          +-------------------------+----------------------------------------------------------------+
       1  |            |     X      |            |            |            |            |            |
          +------------+------------+------------+------------+------------+------------+------------+
       1  |            |            |            |            |            |  HEIGHT    |  WEIGHT    |
          +------------+------------+------------+------------+------------+------------+------------+
       1  |            |            | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
          +------------+------------+------------+------------+------------+------------+------------+
       2  |            |            | ALFRED     |    M       |    14      |    69      |  112.5     |
          +------------+------------+------------+------------+------------+------------+------------+
       3  |            |            | BARBARA    |    F       |    13      |    58      |  101.5     |
          +------------+------------+------------+------------+------------+------------+------------+
       20 |            |            | WILLIAM    |    M       |    15      |   66.5     |  112       |
          +------------+------------+------------+------------+------------+------------+------------+
       [CLASS]



    EXAMPLE OUTPUT

      SASHELP.CLASS total obs=19

       NAME       SEX    AGE    HEIGHT    WEIGHT

       Alfred      M      14     69.0      112.5
       Barbara     F      13     65.3       98.0
       William     M      15     66.5      112.0



    PROCESS (WORKING CODE)
    ======================
      WPS/PROC R - WPS Express does not limit the size of SAS dataset from R?)


        wb <- loadWorkbook("d:/xls/class.xlsx", create = FALSE);
        box<-getBoundingBox(wb, sheet="class");
        want <- readWorksheet(wb, sheet = "class", startRow=box[1,1]);


    OUTPUT
    ======

      SASHELP.CLASS total obs=19

       NAME       SEX    AGE    HEIGHT    WEIGHT

       Alfred      M      14     69.0      112.5
       Alice       F      13     56.5       84.0
       Barbara     F      13     65.3       98.0

    *                _              _       _
     _ __ ___   __ _| | _____    __| | __ _| |_ __ _
    | '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
    | | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
    |_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

    ;

    * output starts at C3;
    ods excel file="d:/xls/class.xlsx";
    ods excel options(sheet_name="class" sheet_interval="none" start_at="C3"); * a1 not needed;
    PROC REPORT DATA=sashelp.class(obs=3) NOCENTER missing;
    RUN;quit;
    ods excel close;

    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    ;

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.have;

    %utl_submit_wps64('
    libname sd1 "d:/sd1";
    options set=R_HOME "C:/Program Files/R/R-3.3.2";
    libname wrk  sas7bdat "%sysfunc(pathname(work))";
    proc r;
    submit;
    source("C:/Program Files/R/R-3.3.2/etc/Rprofile.site", echo=T);
    library(haven);
    library(XLConnect);
    wb <- loadWorkbook("d:/xls/class.xlsx", create = FALSE);
    box<-getBoundingBox(wb, sheet="class");
    want <- readWorksheet(wb, sheet = "class", startRow=box[1,1]);
    endsubmit;
    import r=want data=wrk.want;
    run;quit;
    ');

    proc print data=want;
    run;quit;

