proc import datafile="C:\Users\bmant\Downloads\Reference Metadata.xlsx" out=nov_fvld dbms=xlsx replace;sheet='TEST_CATEGORY';
run;
proc import datafile="C:\Users\bmant\Downloads\Reference Metadata.xlsx" out=nov_qvld dbms=xlsx replace;sheet='QUESTIONS';
run;

proc import datafile="C:\Users\bmant\Downloads\Reference Metadata.xlsx" out=nov_prec dbms=xlsx replace;sheet='PRECISION';
run;


proc import datafile="C:\Users\bmant\Downloads\Reference Metadata.xlsx" out=nov_units dbms=xlsx replace;sheet='TEST_UNITS';
run;


/*Basic VLM for findings for ORRES-note precision/significant digits dependant on source unit. This could be technically merged and unit 
incorporated to where clause but means multiple interpretations of same BC so gets messy*/
data f_wc;
set nov_fvld(rename=(DATA_DOMAIN=DATASET testdef=Label controlled_terminology=Codelist t_datatype=DataType));
format where_clause $500.;
if status='A';
where_clause=trim(left(dataset))||'TESTCD EQ ' ||trim(left(testcd));
if cat ne '' then where_clause=trim(left(where_clause))||' AND '||trim(left(dataset))||'CAT EQ '||trim(left(cat));
if scat ne '' then where_clause=trim(left(where_clause))||' AND '||trim(left(dataset))||'SCAT EQ '||trim(left(scat));
if spec ne '' then where_clause=trim(left(where_clause))||' AND '||trim(left(dataset))||'SPEC EQ '||trim(left(spec));
if method ne '' then where_clause=trim(left(where_clause))||' AND '||trim(left(dataset))||'METHOD EQ '||trim(left(method));
variable=trim(left(dataset))||'ORRES';
Order +1;
run;

/*Basic VLM for questions for ORRES*/
data q_wc;
set nov_qvld(rename=(DATA_DOMAIN=DATASET tstlg=Label controlled_terminology=Codelist t_length=Length signdgt=SignificantDigits));
format where_clause $500.;
if status='A';
where_clause=trim(left(dataset))||'TESTCD EQ ' ||trim(left(testcd));
if scat ne '' then where_clause=trim(left(where_clause))||' AND '||trim(left(dataset))||'SCAT EQ '||trim(left(scat));
if evlint ne '' then where_clause=trim(left(where_clause))||' AND '||trim(left(dataset))||'EVLINT EQ '||trim(left(evlint));
variable=trim(left(dataset))||'ORRES';
Order +1;
run;



proc sql;
create table p21_unit_vld as select distinct a.*,b.preunit,b.unittp from f_wc a left join nov_units b on a.parm=b.parm and b.preunit not in ('NO UNIT' 'CODED');
quit;

proc sql;
create table p21_unit_sf as select distinct a.*,b.pcsnum as signficantdigits from p21_unit_vld a left join nov_prec b on a.parm = b.parm and b.pcsnum gt 0 and a.preunit=b.preunit ;
quit;

data p21_unit_cl_or;
set p21_unit_sf(drop=id); 
ID='UNIT_'||trim(left(parm))||'_OR';
name='Unit subset for '||trim(left(test))||' (Original)';
term=Preunit;
run;

data p21_unitwc;
set p21_unit_sf;
ucount +1;
by order;
if first.order then ucount=1;
run;




proc sql;
create table oru_vld as select distinct a.*,b.id as unitcl from f_wc a left join p21_unit_cl_or b on a.parm=b.parm;
run;

data oru_vld;
set oru_vld;
datatype='text';
drop variable;
variable=trim(left(dataset))||'ORRESU';
run;

data p21_unit_cl_st;
set p21_unit_vld(drop=id); 
if unittp eq 'SI';
ID='UNIT_'||trim(left(parm))||'_ST';
name='Unit subset for '||trim(left(test))||' (Standardized)';
term=Preunit;
run;





