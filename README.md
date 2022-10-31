# utl-powershell-unzip-one-meber-of-a-winzip-archive
Extract just one file from a winzip archive using drop down to powershell
    %let pgm=utl-powershell-unzip-one-meber-of-a-winzip-archive;

    Extract just one file from a winzip archive win 10 64bit ;

    github
    https://tinyurl.com/3nec9mf4
    https://github.com/rogerjdeangelis/utl-powershell-unzip-one-meber-of-a-winzip-archive

    Drop down to powershell macro at end of message;

    Archive has to have more than one member;

    * create a c:/temp folder;
    data _null_;
      rc=dcreate("temp","c:\");
    run;quit;

    * copy sashelp shoes and class to folder;
    libname tmp "c:\temp";
    proc copy in=sashelp out=tmp;
      select shoes class;
    run;quit;

    /*
    c:\temp
      class.sas7bdat
      shoes.sas7bdat
    */

    * create winzip archive;;
    %utl_submit_ps64("
    Add-Type -Assembly System.IO.Compression.FileSystem;
    Compress-Archive -Path c:\temp\*.* -DestinationPath c:\zip\tbl.zip;
    ");

    * get list of files in archive members;
    %utl_submit_ps64("
    Add-Type -Assembly System.IO.Compression.FileSystem;
    $archive = [System.IO.Compression.ZipFile]::OpenRead('c:\zip\tbl.zip');
    $archive.Entries |
    Select-Object -Property Name;
    ");

    /* sas log
      Name
      ----
      class.sas7bdat
      shoes.sas7bdat
    */

    * extract just shoes.sas7bdat out of c:\zip\tbl.zip and into c:\zop\shoes.sas7bdat;

    %utl_submit_ps64("
    Add-Type -Assembly System.IO.Compression.FileSystem;
    $archivePath = 'c:\zip\tbl.zip';
    $destinationDir = 'c:\zip';
    $fileToExtract = 'shoes.sas7bdat';
    $null = New-Item $destinationDir -ItemType Directory -Force;
    $resolvedArchivePath    = Convert-Path -LiteralPath $archivePath;
    $resolvedDestinationDir = Convert-Path -LiteralPath $destinationDir;
    $archive = [IO.Compression.ZipFile]::OpenRead( $resolvedArchivePath );
    try { ;
     if( $foundFile = $archive.Entries.Where({ $_.FullName -eq $fileToExtract }, 'First') ) {;
     $destinationFile = Join-Path $resolvedDestinationDir $foundFile.Name;
     [IO.Compression.ZipFileExtensions]::ExtractToFile( $foundFile[ 0 ], $destinationFile );
     };
     else {;
     Write-Error 'File not found in ZIP: $fileToExtract';
     };
    };
    finally {;
     if( $archive ) {;
     $archive.Close();
     $archive.Dispose();
     };
    };
    ");

    /*
     _ __ ___   __ _  ___ _ __ ___
    | `_ ` _ \ / _` |/ __| `__/ _ \
    | | | | | | (_| | (__| | | (_) |
    |_| |_| |_|\__,_|\___|_|  \___/

    */

    %macro utl_submit_ps64(
          pgm
         ,return=  /* name for the macro variable from Powershell */
         )/des="Semi colon separated set of python commands - drop down to python";

      * write the program to a temporary file;
      filename py_pgm "%sysfunc(pathname(work))/py_pgm.ps1" lrecl=32766 recfm=v;
      data _null_;
        length pgm  $32755 cmd $1024;
        file py_pgm ;
        pgm=&pgm;
        semi=countc(pgm,';');
          do idx=1 to semi;
            cmd=cats(scan(pgm,idx,';'));
            if cmd=:'. ' then
               cmd=trim(substr(cmd,2));
             put cmd $char384.;
             putlog cmd $char384.;
          end;
      run;quit;
      %let _loc=%sysfunc(pathname(py_pgm));
      %put &_loc;
      filename rut pipe  "powershell.exe -executionpolicy bypass -file &_loc ";
      data _null_;
        file print;
        infile rut;
        input;
        put _infile_;
        putlog _infile_;
      run;
      filename rut clear;
      filename py_pgm clear;


      * use the clipboard to create macro variable;
      %if "&return" ^= "" %then %do;
        filename clp clipbrd ;
        data _null_;
         length txt $200;
         infile clp;
         input;
         putlog "*******  " _infile_;
         call symputx("&return",_infile_,"G");
        run;quit;
      %end;

    %mend utl_submit_ps64;

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */



