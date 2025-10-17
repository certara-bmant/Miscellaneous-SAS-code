

%let fpath=C:\Users\bmant\Downloads;
%let dssfile=cdisc_sdtm_dataset_specializations_latest (4);

/*Import dss*/
proc import datafile="&fpath.\&dssfile..xlsx" out=dss dbms=xlsx replace;sheet='SDTM Dataset Specializations';
run;

/*Create subcodelists for vlm in P21 */
data unit_cl;
format name $200. id $200.;
set dss;
if vlm_target='Y' and codelist_submission_value='UNIT' ;
if substr(sdtm_variable,3,6)='ORRESU' then name='Unit, subset for '||trim(left(short_name))||' - Original';
if substr(sdtm_variable,3,6)='STRESU' then name='Unit, subset for '||trim(left(short_name))||' - Standardized';
if substr(sdtm_variable,3,6)='ORRESU' then id=trim(left(codelist_submission_value))||'_'||trim(left(vlm_group_id))||'_OR';
if substr(sdtm_variable,3,6)='STRESU' then id=trim(left(codelist_submission_value))||'_'||trim(left(vlm_group_id))||'_ST';
if value_list='' and assigned_value ne '' then value_list=assigned_value;
listcount=countc(value_list,';')+1;
do order=1 to listcount;
term=scan(value_list,order,';');
output;

end;

keep id name  order term vlm_group_id sdtm_variable;

run;

data dss_ chks;
format typchk $200.;
set dss;
if significant_digits ne '' and length ne '200' and data_type ne 'float' then typchk="Data type should be float";
if length='200' and significant_digits ne '' then typchk="Data type likely text and digits not required"; 
if typchk ne '' then output chks;

if significant_digits ne '' and length ne '200' then data_type='float';/*Fixes for discrepant details i.e. non float with digits*/
if length='200' then significant_digits =''; 
output dss_;

run;

proc sort data=dss_ out=dss__;
by vlm_group_id sdtm_variable ;
run;

data wc;

set dss__;

*if comparator ='' and sdtm_variable eq 'LBLOINC' and assigned_value ne '' then comparator='EQ';/*fix for duplicates-better qualifiers should be used i.e. LBRESTYP*/


if comparator ne '';
if index(trim(left(assigned_value)),' ') gt 0 then assigned_value='"'||trim(left(assigned_value))||'"';
if value_list ne '' then vlx='("'||trim(left(tranwrd(value_list,';','","')))||'")';
if comparator='EQ' then wc=trim(left(sdtm_variable))||' '||trim(left(comparator))||' '||trim(left(assigned_value));
if comparator='IN' then wc=trim(left(sdtm_variable))||' '||trim(left(comparator))||' '||trim(left(vlx));

run;





/*Generate where clause*/

data wc_;
format combinedwc $500.;
set wc;
retain combinedwc;
by   vlm_group_id sdtm_variable  ;
if first.vlm_group_id then combinedwc=wc;
  else  combinedwc=catx(' and ', combinedwc, wc);

if last.vlm_group_id then output;
run;

/*Merges*/
proc sql;
create table vlmerge(keep=domain short_name vlm_group_id combinedwc sdtm_variable data_type length format codelist_submission_value significant_digits mandatory_value assigned_value) 
as select distinct a.*,b.combinedwc from dss_ a left join wc_ b on a.vlm_group_id=b.vlm_group_id where a.vlm_target='Y';/*Merge full where caluse to targets*/
create table vlmerge2 as select  distinct a.*,b.id from vlmerge a left join  unit_cl b on a.vlm_group_id=b.vlm_group_id and a.sdtm_variable=b.sdtm_variable;/*merge unit codelist*/
quit;

/*Format to P21 VL template*/
data vlm(rename=(domain=Dataset sdtm_variable=variable combinedwc=where_clause codelist_submission_value=codelist short_name=label mandatory_value=mandatory));
format codelist_submission_value Decoded_Variable Codelist_Expected	Expected_Codelist_ID Expected_Codelist_Name	Origin	Source	Method	Predecessor	Comment	Developer_Notes variant $100.;
set vlmerge2;
if id ne '' then codelist_submission_value=id;
origin='Collected';
order+1;
drop id vlm_group_id;
run;

/*Reorder*/
data vlmfinal;
retain Order	Dataset	Variable	Variant	Where_Clause	Label	Data_Type	Length	Significant_Digits	Format	Mandatory	Assigned_Value	Codelist	Decoded_Variable
Codelist_Expected	Expected_Codelist_ID	Expected_Codelist_Name	Origin	Source	Method	Predecessor	Comment	Developer_Notes;
set vlm;
run;

/*Export of vlm in P21 format*/
proc export data=vlmfinal outfile="&fpath.\bc_vlm.xlsx" dbms=xlsx replace;
run;


proc sort data=vlmerge2 dupout=vlmdups nodupkey;
by sdtm_variable combinedwc;
run;
proc export data=vlmdups outfile="&fpath.\dss_checks.xlsx" dbms=xlsx replace;
sheet='Duplicates';
run;

proc export data=chks outfile="&fpath.\dss_checks.xlsx" dbms=xlsx replace;
sheet='Data attribute Check';
run;
