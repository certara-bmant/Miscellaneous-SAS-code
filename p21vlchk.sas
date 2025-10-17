%macro p21vldchk(dir=,specfile=,specsh=,vendfile=,key=);

/*Import vendor data*/
proc import datafile="&dir./&vendfile..csv" out=vendor dbms=csv replace;
guessingrows=max;
run;

/*Import specifications data*/
proc import datafile="&dir./&specfile" out=specdata dbms=xlsx replace;sheet="&specsh";
run;


proc sort data=specdata(keep=&key);
by &key;
run;

/*Subset vendor data by keys*/
proc sort data=vendor(keep=&key) nodupkey;
by &key;
run;

/*Merge and compare value from transposed dataset valuelevel by predefined key variables*/
data vlnomatch vlmatch;
format comment &key $200.;
merge vendor (in=a) specdata (in=b);
by &key;
if a and not b then comment='Value level combination present in vendor data but NOT specifications';
if b and not a then comment='Value level combination present in specifications data but NOT vendor data';
if a and b then comment='Value level combination matches vendor data';


keep &key comment ;

if (a and not b) or (b and not a) then output vlnomatch;
if a and b then output vlmatch;

run;

proc export data=vlmatch outfile="&dir./vlchk" dbms=xlsx replace;
sheet=vlmatch;
run;
proc export data=vlnomatch outfile="&dir./vlchk" dbms=xlsx replace;
sheet=vlnomatch;
run;


/*SQL version to find in vendor data but not in spec*/
/*Convert delimiter to comma for sql statement*/
%let keyc=%sysfunc(translate(&key,',',' '));

/*Generate where clause for each key variable*/
%do i = 1 %to %sysfunc(countw(&key));
%if &i eq 1 %then %let keywc=a.%sysfunc(scan(&key,&i))=b.%sysfunc(scan(&key,&i));  
%if &i gt 1 %then %let keywc=&keywc. and a.%sysfunc(scan(&key,&i))=b.%sysfunc(scan(&key,&i));  

%end;

%put &keywc;

/*Create table from vendor data where no match ffound*/
proc sql;
create table query as
select &keyc from

vendor a
where not exists (select &keyc from specdata b

where &keywc);
quit;


/*Create table from spec data where no match found in vendor*/
proc sql;
create table spquery as
select &keyc from

specdata a
where not exists (select &keyc from vendor b

where &keywc);
quit;


/*Find wc in spec but not in data*/
proc import datafile="&dir./&specfile" out=vl dbms=xlsx replace;sheet="ValueLevel";
run;


data wcs_;
format where_clause_ $500.;
set vl;
count+1;
/*Add quotes for values to run in where clause*/
where_clause_=tranwrd(where_clause,'EQ ','EQ "');
where_clause_=tranwrd(where_clause_,' and','" and');
where_clause_=tranwrd(where_clause_,'""','"');
if substr(where_clause_,length(where_clause_),1) ne '"' then where_clause_=trim(left(where_clause_))||'"';

run;

/*Convert to sequenced macro variables*/
data _null_;
set wcs_;
count+1;
call symput('wc'||compress(_n_),where_clause_);

run;


proc sql;
create table wcs as select distinct where_clause from vl;
quit;

data squery_all;
if 0 then output;
run;


/*If no records found then create issue table with relavant where clause*/
%do t=1 %to &sqlobs;
proc sql;

create table squery&t as
select where_clause_ from

wcs_ a
where not exists (select &keyc from vendor b

where &&wc&t ) and a.count=&t;


quit;

/*Accumulate issues from each loop*/
data squery_all;
set squery_all squery&t ;
run;

%end;





data wcs__;
set wcs_;
format where_clause_;
where_clause_='('||trim(left(where_clause_))||')';
run;

proc transpose data=wcs__ out=wcall;
var where_clause_;
run;


data wcall_;
format wcmax $5000.;
set wcall;

wcmax=catx(' OR ',of col:);

 call symput('wcall',wcmax);

run;


proc sql;

create table vquery as
select  &keyc from vendor a

where not (&wcall ) ;


quit;




/*Generate vendor WC based on keys*/
data vendorwc;
format wc $500.;
set vendor;
count+1;
do z = 1 to countw("&key");
if z eq 1 then wc=scan("&key", z)||'="'||trim(left(vvaluex(scan("&key", z))))||'"';

if z gt 1 and vvaluex(scan("&key", z)) ne '' then wc=trim(left(wc))||' and '||scan("&key", z)||'="'||trim(left(vvaluex(scan("&key", z))))||'"';
end;
run;


/*Convert to sequenced macro variables*/
data _null_;
set vendorwc end=eof;
call symput('vwc'||compress(_n_),wc);
if eof then call symput('vwcr',count);
run;


data vquery_all;
if 0 then output;
run;


/*If no records found then create issue table with relavant where clause*/
%do x=1 %to &vwcr;
proc sql;

create table vquery&x as
select wc from

vendorwc a
where not exists (select &keyc from specdata b

where &&vwc&x ) and a.count=&x;


quit;

/*Accumulate issues from each loop*/
data vquery_all;
set vquery_all vquery&x ;
run;

%end;






/*Find keys for VLM-in practice would need to loop per dataset*/
data wcvars;
   set wcs;
 n=count(where_clause,'EQ');/*Limited in scope by only defining variables as any word directly before a 'EQ' comparator*/

  p = 0;
   do i=1 to n until(p=0); 
      p = find(where_clause,"EQ", p+1);
	output;
   end;
   run;

  data wcvars_;
   set wcvars;
   count+1;
var=scan(substr(where_clause,1,p-1),-1,' ');

   run;
proc sort data=wcvars_ out=wcvars__(keep=var count) nodupkey;
by var;
run;

proc sort data=wcvars__ ;
by count;
run;

proc transpose data=wcvars__ out=wcvar(drop=_name_);
id count;
var var;
run;

data spkey;
set wcvar;
  result=catx(' ',of _:);
  call symput('spkey',result);/*create macro variable*/
  run;

  %put &spkey;




%mend;

%p21vldchk(dir=C:\Users\bmant\OneDrive - Certara\P21TESTS, /*working directory*/
specfile=Ben Testing (6), /*Name of specifications file-do NOT include file extension*/
specsh=LB_CENTRAL, /*Name of worksheet containing dataset level Valuelevel data*/
vendfile=LB_CENTRAL, /*Name of vendor file. do NOT include file extension. Must be in csv format.*/
key=LBTESTCD LBCAT LBSPEC LBMETHOD); /*Key combination of variables providing uniqueness in VL data*/



