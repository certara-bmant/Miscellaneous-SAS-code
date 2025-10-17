%macro nov_P21_cl(dir=,mdfile=,outfile=,cltype=,stver=,endver=,nm=,type=,ver=);

/*codelist imports*/

options validvarname=any;

libname xl xlsx "&dir./&mdfile." ;

data codelist cldescs;
format terminology $50.;
set xl.codelist(rename=(id=seq code=nci_term_code controlled_terminology=id cdisc_definition=term_description sdtm=term codelist_code=nci_codelist_codeid decode_submission_value=decoded_value
numeric_codelist=order CDISC_SYNONYM=Synonyms CONTROLLED_TERM_DESCRIPTION=name codelist_extensible=extensible));

if upcase(ct_type)=upcase("&cltype");
Terminology="&outfile";
stver=substr(nci_start_version,1,10);
%if &cltype = GLOSSARY %then %do;
if nci_codelist_codeid ne '' then id='CDISC Glossary';
%end;
if id='' then output cldescs;
if id='' then delete; /*Codelist details*/
output codelist;
run;



proc sql;
create table cl_merge as select distinct a.*, b.term_description as codelist_description
from codelist a left join cldescs b on a.id =b.term ;
quit;

/*fix for repeated terms*/
proc sort nodupkey data=cl_merge;
by name term;
run;

/*Add label and extra variables for export*/
data codelist_ ;
set cl_merge;
attrib 
NCI_Codelist_CodeID label='NCI Codelist CodeID'
name label='Name'
term label='Term'
synonyms label='Synonyms'
decoded_value label='Decoded Value'
nci_term_code label='NCI Term Code'
term_description label='Term Description'
extensible label='Extensible'
data_type format=$1. label='Data Type'
id label='ID'
codelist_description label='Codelist Description'
order label='Order'
;


run;




/*Order variables*/
data codelist_p21_out;
retain ID	Name	Codelist_Description	Extensible	NCI_Codelist_CodeID	Terminology	Data_Type	Order	Term	Term_Description	NCI_Term_Code	Decoded_Value	Synonyms;
set codelist_;
keep ID	Name	Codelist_Description	Extensible	NCI_Codelist_CodeID	Terminology	Data_Type	Order	Term	Term_Description	NCI_Term_Code	Decoded_Value	Synonyms;

run;

proc sort data=codelist_p21_out;
by id order;
run;

/*Export*/
proc export data=codelist_p21_out outfile="&dir./&outfile..xlsx" dbms=xlsx label replace;sheet='Codelists';
run;

/*Standard Properties*/

%include 'C:\Users\bmant\Downloads\termcode.sas';

data term_out(rename=(test=value));
set term;
test=dequote(resolve(value));
drop value;
run;

proc export data=term_out outfile="&dir./&outfile..xlsx" dbms=xlsx label replace;sheet='Terminology';
run;


data _null_;
fname = 'todelete';
rc = filename(fname, "&dir./&outfile..xlsx.bak");
rc = fdelete(fname);
rc = filename(fname);
run;



%mend;

options mprint symbolgen;
%nov_P21_cl(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis SDTM Terminology,cltype=SDTM,nm=Novartis SDTM Terminology,type=SDTM,ver=999999);
%nov_P21_cl(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis ADAM Terminology,cltype=ADAM,nm=Novartis ADAM Terminology,type=ADaM,ver=999999);
%nov_P21_cl(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis CDASH Terminology,cltype=CDASH,nm=Novartis CDASH Terminology,type=non-CDISC,ver=999999);
%nov_P21_cl(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis Glossary Terminology,cltype=GLOSSARY,nm=Novartis Glossary Terminology,type=non-CDISC,ver=999999);
%nov_P21_cl(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis Protocol Terminology,cltype=PROTOCOL,nm=Novartis Protocol Terminology,type=non-CDISC,ver=999999);
%nov_P21_cl(dir=C:\Users\bmant\Downloads,mdfile=Stage Metadata.xlsx,outfile=Novartis Define Terminology,cltype=DEFINE,nm=Novartis Define Terminology,type=non-CDISC,ver=999999);
