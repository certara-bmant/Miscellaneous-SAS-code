%macro merck_kgaa_P21(dir=,mdfile=,outfile=,type=);
%put &dir.\&mdfile;
/*datsets imports*/

options validvarname=any;

libname xl xlsx "&dir./&mdfile." ;
data cover;
set xl.cover;
run;


data spec;
set xl.'Data Structure'n;
run;

data _vlm;
set xl.VLM;
if domain ne '';
run;


data cl;
set xl.codelist;

run;

data vis;
set xl.visits;
run;

data variables_tmp;
attrib 
'Variable Order'n label='Order'
'Variable Name'n label='Variable'
'Variable Label'n label='Label'
Codelist format=$20.
Format format=$50.
'IO_Variable Implementation Instr'n label='Developer Notes'
Type label='Data Type' format=$20.;
;
set spec;
if substr('Codelist or Format'n,1,1)='(' then codelist=compress(codelist,'()');
else format='Codelist or Format'n;
if upcase(type)='CHAR' then type='text';
if upcase(type)='NUM' then type='integer';

run;
                        /* sheet name containing the VLM data */



/* ===== GET VARIABLE LIST, EXCLUDING NAMES THAT START WITH IO (case-insensitive) ===== */
proc contents data=_vlm out=_vars(keep=name type) noprint; run;

proc sql noprint;
  /* TYPE: 1 = numeric, 2 = character */
  create table _keep as
  select name, type
  from _vars
  where upcase(substr(strip(name),1,2)) ne 'IO';

  /* Load names and types into macro variables */
  select name, type into :var1-:var999, :type1-:type999 from _keep;
  %let nvars = &sqlobs;
quit;

/* ===== HELPER: CHECK IF A VARIABLE EXISTS IN A DATASET ===== */
%macro var_exists(dsn, var);
  %local dsid rc varnum exists;
  %let dsid = %sysfunc(open(&dsn,i));
  %if &dsid %then %do;
    %let varnum = %sysfunc(varnum(&dsid,&var));
    %let rc = %sysfunc(close(&dsid));
    %if &varnum > 0 %then 1;
    %else 0;
  %end;
  %else 0;
%mend;

/* ===== BUILD CODE LIST =====
   Rules:
   - ID = column name
   - Term = distinct, non-missing values from that column, after converting numerics to text
   - Ignore any Term containing '(' or ')'
   - For variables ending with TESTCD, if matching TEST exists, also output Decode = corresponding TEST
*/
proc datasets lib=work nolist; delete CodeList _one; quit;

%macro build_codelist;
  %do i = 1 %to &nvars;
    %let v = &&var&i;
    %let t = &&type&i;

    /* Determine if this is a __TESTCD-style variable and compute matching __TEST name */
    %let UPPER_V = %upcase(&v);
    %let is_testcd = 0;
    %let decodeVar = ;
    %if %sysfunc(prxmatch(/TESTCD$/, &UPPER_V)) %then %do;
      %let is_testcd = 1;
      /* Replace trailing TESTCD with TEST to form the decode variable name */
      %let prefix_len = %eval(%length(&v) - 6); /* "TESTCD" has 6 chars */
      %let decodeVar = %substr(&v, 1, &prefix_len)TEST;
    %end;

    /* Build expressions for Term & (optionally) Decode */
    %if &t = 2 %then %do;  /* character */
      %let term_expr = strip(&v);
    %end;
    %else %do;             /* numeric -> character */
      %let term_expr = strip(put(&v, best32.));
    %end;

    /* If TESTCD and decodeVar exists, add Decode; else blank Decode */
    %let have_decode = 0;
    %if &is_testcd %then %do;
      %if %var_exists(work._vlm, &decodeVar) %then %let have_decode = 1;
    %end;

    proc sql;
      create table _one as
      select distinct
             "&v" as ID length=128,
             &term_expr as Term length=400
             %if &have_decode %then %do;
               , strip(&decodeVar) as Decode length=400
             %end;
             %else %do;
               , "" as Decode length=400
             %end;
      from _vlm
      /* Exclude missing, and exclude values that contain brackets-fixed */
      where not missing(&v)
        and substr(&term_expr, 1,1) ne '('
      ;
    quit;

    /* Append into a single CodeList table */
    proc append base=CodeList data=_one force; run;
    proc datasets lib=work nolist; delete _one; quit;
  %end;

  /* Clean, collapse whitespace, and de-duplicate on (ID, Term, Decode) */
  data CodeList;
    set CodeList;
    ID     = strip(ID);
    Term   = compbl(strip(Term));
    Decode = compbl(strip(Decode));
	'Data Type'n='text';
  run;

  proc sort data=CodeList nodupkey;
    by ID Term Decode;
  run;
%mend;

%build_codelist;

/* ===== WRITE BACK TO THE EXCEL FILE AS A NEW SHEET ===== */
/*data xl.CodeListX;
  set CodeList;
run;

libname xl clear;*/

data cl_(rename=('Codelist Name'n=ID 'Codelist Description'n=Name));
attrib type label='Data type'
'Codelist Description'n label='Name'
;
set cl;
*if 'Codelist Type'n ne 'CDISC';
'Data Type'n='text';
if 'Codelist Name'n='' then delete;
run;


proc sql;
create table codelist_name as select distinct a.*,b.'Variable Label'n as Name from codelist a left join spec b on a.id=b.'Variable Name'n;
quit;
proc sort data=cl_;
by id;
run;
data codelist_all;
format term id $200.;
set cl_ codelist_name ;
count +1;

if first.id then count=1;
by id;
run;

data codelist_all;
format blank $1.;
set codelist_all;

id=compress(id,'()');
run;

data codelist_all_;
retain count id term blank 'Decoded Value'n; 
set codelist_all;


run;

proc sort data=codelist_all_ out=cl_list(keep= ID Name) nodupkey;
by id;
run;

proc export data=codelist_all_ outfile="&dir/codelists_&type..xlsx" dbms=excel replace;
run;

proc export data=cl_list outfile="&dir/codelists_&type..xlsx" dbms=excel replace;
run;

proc export data=Variables_tmp outfile="&dir/spec_&type..xlsx" 
dbms=excel replace;
run;
%mend;

%merck_kgaa_p21(dir=C:\users\bmant\merck_kgaa,mdfile=CPGX101_Cell Phenotyping_DTS_Template.xlsx,outfile=,type=CP);
