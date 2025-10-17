%macro nov_P21(dir=,mdfile=,outfile=,mdtype=,stver=,endver=);
%put &dir.\&mdfile;
/*datsets imports*/

options validvarname=any;

libname xl xlsx "&dir./&mdfile." ;

data dsets;
format endver stver $3.;
set xl.DOMAIN;
/*Convert to text for filtering*/
endver=end_version;
stver=start_version;
%if stver  ne %then %do;
if findw("&stver",stver) gt 0;
%end;
%if endver  ne %then %do;
if findw("&endver",endver) gt 0;
%end;

%if mdtype  ne %then %do;
if find(upcase(standard),"&mdtype") gt 0;
%end;
run;


data vars;
format endver stver $3. data_type $10.;
set xl.ELEMENT;
endver=end_version;
stver=start_version;
%if stver  ne %then %do;
if findw("&stver",stver) gt 0;
%end;
%if endver  ne %then %do;
if findw("&endver",endver) gt 0;
%end;


%if mdtype  ne %then %do;
if find(upcase(standard),"&mdtype") gt 0;
%end;

/*Determine core SDTM/ADAM*/
if element_core='PERM' then core='Permitted';
if element_core='REQ' then core='Required';
if element_core='EXP' then core='Expected';

/*CDASH*/
if element_core='O' then core='Optional';
if element_core='R/C' then core='Recommended/Conditional';
if element_core='HR' then core='Highly Recommended';


/*Determine Mandatory*/
if core='Required' then Mandatory='Yes';else Mandatory='No';

/*Determine Data Type*/
if index(element_format,'CHAR') then data_type='text';
if index(element_format,'DATE') then data_type='date';
if index(element_format,'DATETIME') then data_type='datetime';
if index(element_format,'TIME') then data_type='time';
if index(element_format,'NUM') then data_type='integer';

/*Determine Float Data Type attributes*/
if index(element_format,',') then do;
data_type='float';
length=substr(element_format,index(element_format,'(')+1,length(element_format)-index(element_format,','));
significant_digits=substr(element_format,index(element_format,',')+1,length(element_format)-index(element_format,',')-1);
end;

/*Determine Lengthe where not float attributes*/
if index(element_format,',') eq 0 then do;
length=substr(element_format,index(element_format,'(')+1,length(element_format)-index(element_format,'(')-1);
end;


if controlled_terminology ne 'N/A' then Codelist=controlled_terminology;
if element_role ne 'N/A' then role=element_role;
run;




data p21_dsets;
set dsets(rename=(data_domain=dataset keys=key_variables));
attrib 
dataset label='Dataset'
key_variables label='Key Variables'
structure label='Structure'
repeating label='Repeating'
Variant format=$1. label='Variant'
Class format=$1. label='Class'
Subclass format=$1. label='Subclass'
Prototype format=$1. label='Prototype'
Reference_data format=$1. label='Reference Data'
Comment format=$1. label='Comment'
Developer_Notes format=$1. label='Developer Notes'
Description format=$1. label='Description'

;


run;

data p21_dsets_out;
retain Dataset	Variant	Label	Description	Class	Subclass	Structure	Key_Variables	Prototype	Standard	Repeating	Reference_Data	Comment	Developer_Notes;
set p21_dsets;
keep Dataset	Variant	Label	Description	Class	Subclass	Structure	Key_Variables	Prototype	Standard	Repeating	Reference_Data	Comment	Developer_Notes;


run;


proc sort data=vars;
by data_domain id;
run;

data p21_vars;
set vars(rename=(data_domain=dataset data_element=Variable element_label=Label element_definition=description));
attrib 
dataset label='Dataset'
variable label='Variable'
Variant format=$1. label='Variant'
Label label='Label'
Comment format=$1. label='Comment'
Developer_Notes format=$1. label='Developer Notes'
Description  label='Description'
data_type  label='Data Type'
significant_digits  label='Significant Digits'
Assigned_value format=$1. label='Assigned Value'
Codelist_expected format=$1. label='Codelist Expected'
Expected_codelist_id format=$1. label='Expected Codelist ID'
Expected_codelist_name format=$1. label='Expected Codelist Name'
Decoded_variable format=$1. label='Decoded Variable'
Origin format=$1. label='Origin'
Format format=$1. label='Format'
Method format=$1. label='Method'
Source format=$1. label='Source'
role  label='Role'
Predecessor format=$1. label='Predecessor'
raw_variables format=$1. label='Raw Variables'
;
order+1;
if first.dataset then order=1;
by dataset;

run;


data p21_vars_out;
retain Order	Dataset	Variable	Variant	Label	Description	Data_Type	Length	Significant_Digits	Format	Core	Mandatory	Assigned_Value	Codelist	Decoded_Variable	Codelist_Expected	Expected_Codelist_ID	Expected_Codelist_Name	Origin	Source	Method	Predecessor	Role	Comment	Raw_Variables	Developer_Notes;

set p21_vars;
keep Order	Dataset	Variable	Variant	Label	Description	Data_Type	Length	Significant_Digits	Format	Core	Mandatory	Assigned_Value	Codelist	Decoded_Variable	Codelist_Expected	Expected_Codelist_ID	Expected_Codelist_Name	Origin	Source	Method	Predecessor	Role	Comment	Raw_Variables	Developer_Notes;

run;

proc export data=p21_dsets_out outfile="&dir./&outfile..xlsx" dbms=xlsx label replace ;sheet='Datasets'; 
run;

proc export data=p21_vars_out outfile="&dir./&outfile..xlsx" dbms=xlsx label replace; sheet='Variables';
run;




%mend;

options mprint symbolgen;
%nov_P21(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis SDTM 3.2,mdtype=SDTM,stver=3.2,endver=3.2 999);
%nov_P21(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis SDTM 3.3,mdtype=SDTM,stver=3.2 3.3,endver=999);
%nov_P21(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis ADAM 1,mdtype=ADAM,stver=1,endver=1 999);
%nov_P21(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis ADAM 1.1,mdtype=ADAM,stver=1 1.1,endver=1.1 999);
%nov_P21(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis ADAM 1.2,mdtype=ADAM,stver=1 1.1 1.2,endver=1.2 999);
%nov_P21(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis CDASH 2.1,mdtype=CDASH,stver=2.1,endver=2.1 999);
