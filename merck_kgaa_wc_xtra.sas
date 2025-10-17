
%let xlsx_path = c:/users/bmant/merck_kgaa/LBGX101_IQVIALab_SafetyData_DTS_Template_highlighted_dups.xlsx;   /* <-- set your file path */
%let vlm_sheet = VLM;
options validvarname=any;

/* ========= OPEN EXCEL & READ VLM ========= */
libname xl xlsx "&xlsx_path";
data _vlm; set xl.&vlm_sheet; run;

/* ======================= FROM _VLM ONLY (consolidates identical-term TESTCD groups) ======================= */
%macro build_from_vlm_wide(
  vlm            = _vlm,                 /* imported VLM sheet (wide) */
  outds          = Where_Clauses_ByVar,  /* final per-clause table */
  outcodes       = CodeLists_ByVar,      /* raw CodeList_ID -> Term rows */
  codelistds     = CodeLists_Terms,      /* two-col code list: ID, Term */
  domaincol      = DOMAIN,               /* domain column name in _vlm */
  exclude_suffix = RES RESU ORRES ORRESU TESTCD, /* variables to ignore for vlm target generation*/
  outpath =, 
  type=,
  prefix=
);

  /* ---- discover <prefix>TESTCD and candidate variable columns ---- */
  proc contents data=&vlm out=_vls(keep=name varnum type length) noprint; run;

%if %upcase(&prefix)=Y %then %do;
data _prefixes;
    set _vls;
    length  prefix $8;
   
    if substr(name,3,6) eq 'TESTCD' then do;
      prefix = substr(name, 1, 2);
      output;
    end;
    keep prefix;
  run;
  run;
  proc sort data=_prefixes nodupkey; by prefix; run;

  proc sql;
    create table _varmap as
    select a.name as varname length=128,
           upcase(strip(a.name)) as name_up length=128,
           upcase(substr(a.name,1,2)) as prefix length=8,
           substr(upcase(strip(a.name)),3) as suffix length=64,
           a.varnum
    from _vls a
    inner join _prefixes b
      on upcase(substr(a.name,1,2)) = b.prefix
    where upcase(strip(a.name)) ne 'DOMAIN'
  ;
  quit;

  data _varmap;
    set _varmap;
    if findw(upcase("&exclude_suffix"), strip(suffix), ' ') then delete;
  run;

  %local __n varcnt;
  proc sql noprint; select count(*) into :__n from _varmap; quit;
  %if &__n=0 %then %do;
    %put NOTE: No candidate variables after exclusions (&exclude_suffix).;
    data &outds; length Dataset Variable Assigned_Value CodeList_ID Where_Clause $1; stop; run;
    data &outcodes; length CodeList_ID $1 Value $1; stop; run;
    data &codelistds; length ID $1 Term $1; stop; run;
    %return;
  %end;

  proc sort data=_varmap; by varnum; run;
  proc sql noprint;
    select varname, prefix, suffix
      into :vm_var1-:vm_var999, :vm_pre1-:vm_pre999, :vm_suf1-:vm_suf999
    from _varmap
    order by varnum;
    %let varcnt=&sqlobs;
  quit;

  /* ---- LONG: (Dataset, Variable, Assigned_Value, TestCD) ---- */
  data _long;
    length Dataset $32 Variable $64 Assigned_Value $4000 TestCD $200;
    set &vlm;
    Dataset = upcase(vvaluex("&domaincol"));
    %do i=1 %to &varcnt;
      TestCD         = strip(vvaluex(cats("&&vm_pre&i",'TESTCD')));
      Variable       = "&&vm_var&i";
      Assigned_Value = strip(vvaluex("&&vm_var&i"));
      if not missing(Dataset) and not missing(TestCD) and not missing(Assigned_Value) then output;
    %end;
    keep Dataset Variable Assigned_Value TestCD;
  run;


%end;

%if %upcase(&prefix) ne Y %then %do;

  data _varmap;
    set _vls;
    if findw(upcase("&exclude_suffix"), strip(name), ' ') then delete;
	if substr(name,1,2) eq 'IO' then delete;
  run;



  proc sort data=_varmap; by varnum; run;
  proc sql noprint;
    select name
      into :vm_var1-:vm_var999
    from _varmap
    order by varnum;
    %let varcnt=&sqlobs;
  quit;

  data _long;
    length Dataset $32 Variable $64 Assigned_Value $4000 TestCD $200;
    set &vlm;
    Dataset = 'SAFETY';
    %do i=1 %to &varcnt;
      TestCD         = strip(vvaluex('TESTCD'));
      Variable       = "&&vm_var&i";
      Assigned_Value = strip(vvaluex("&&vm_var&i"));
      if not missing(Dataset) and not missing(TestCD) and not missing(Assigned_Value) then output;
    %end;
    keep Dataset Variable Assigned_Value TestCD;
  run;

%end;

  proc sort data=_long nodupkey; by Dataset Variable Assigned_Value TestCD; run;



  /* =================== A) AMBIGUOUS TESTCDS (same TESTCD -> >1 terms) =================== */
  /* find TESTCDs that map to multiple terms */
  proc sql;
    create table _ambig as
    select Dataset, Variable, TestCD,
           count(distinct Assigned_Value) as nvals
    from _long
    group by 1,2,3
    having calculated nvals > 1;
  quit;

  /* terms for ambiguous TESTCDs, with normalized token for set-comparison */
  proc sql;
    create table _ambig_terms as
    select distinct
      l.Dataset, l.Variable, l.TestCD,
      l.Assigned_Value as Term length=4000
    from _long l inner join _ambig a
      on l.Dataset=a.Dataset and l.Variable=a.Variable and l.TestCD=a.TestCD
    ;
  quit;
  proc sort data=_ambig_terms nodupkey; by Dataset Variable TestCD Term; run;

  /* build a canonical signature of the term set per (Dataset, Variable, TestCD) */
  data _ambig_sig;
    length Dataset $32 Variable $64 TestCD $200 Term $4000 Term_Norm $4000 Sig $200;
    set _ambig_terms;
    by Dataset Variable TestCD Term;
    Term_Norm = upcase(compbl(strip(Term)));
    retain Sig;
    if first.TestCD then Sig='';
    Sig = catx('|', Sig, Term_Norm);    /* sorted by Term due to BY order */
    if last.TestCD then output;
    keep Dataset Variable TestCD Sig;
  run;

  /* group TESTCDs that have the exact same term set signature */
  proc sort data=_ambig_sig; by Dataset Variable Sig TestCD; run;

  data _ambig_groups;
    length Dataset $32 Variable $64 Sig $32000 n 8 tcd_list $2000;
    retain n tcd_list;
    set _ambig_sig;
    by Dataset Variable Sig;
    if first.Sig then do; n=0; tcd_list=''; end;
    n+1;
    tcd_list = catx('_', tcd_list, upcase(compress(TestCD, , 'kad'))); /* deterministic order */
    if last.Sig then output;
    keep Dataset Variable Sig n tcd_list;
  run;

  /* map each ambiguous TESTCD to its grouped ID */
  proc sql;
    create table _ambig_map as
    select s.Dataset, s.Variable, s.TestCD, g.n,
           cats(s.Variable,'_',g.tcd_list) as Group_ID length=256,
           g.Sig
    from _ambig_sig s
    inner join _ambig_groups g
      on s.Dataset=g.Dataset and s.Variable=g.Variable and s.Sig=g.Sig
    ;
  quit;

  /* build IN(...) list per group for the main table */
  proc sort data=_ambig_sig; by Dataset Variable Sig TestCD; run;
  data _ambig_list;
    length Dataset $32 Variable $64 Sig $200 list $500;
    retain list;
    set _ambig_sig;
    by Dataset Variable Sig;
    if first.Sig then list='';
    list = catx(', ', list, cats("'", tranwrd(TestCD,"'","''"), "'"));
    if last.Sig then output;
    keep Dataset Variable Sig list;
  run;

  /* assemble rows for AMBIG part (consolidated when n>1) */
  proc sql;
    create table _rows_ambig as
    select m.Dataset, m.Variable,
           ' ' as Assigned_Value length=200,
           m.Group_ID as CodeList_ID length=256,
      case
  when countc(l.list, ',') = 0
  %if %upcase(&prefix) eq Y %then %do;
    then catx(' ', upcase(substr(m.Variable,1,2))||'TESTCD', 'EQ', strip(l.list))
  else
    trim(upcase(substr(m.Variable,1,2))) || 'TESTCD IN (' || strip(l.list) || ')'
end as Where_Clause length=2000
%end;

  %if %upcase(&prefix) ne Y %then %do;
    then catx(' ', 'TESTCD', 'EQ', strip(l.list))
  else
     'TESTCD IN (' || strip(l.list) || ')'
end as Where_Clause length=2000
%end;

    from _ambig_map m
    inner join _ambig_list l
      on m.Dataset=l.Dataset and m.Variable=l.Variable and m.Sig=l.Sig
    group by m.Dataset, m.Variable, m.Group_ID, l.list
    ;
  quit;

  /* terms for each consolidated ambiguous group -> CodeLists tables */
  proc sql;
    create table _ambig_terms_gid as
    select distinct
      m.Group_ID as CodeList_ID length=256,
      t.Term      as Value     length=4000
    from _ambig_terms t
    inner join _ambig_map m
      on t.Dataset=m.Dataset and t.Variable=m.Variable and t.TestCD=m.TestCD
    ;
  quit;

  /* =================== B) NON-AMBIG TESTCDS (usual consolidation) =================== */
  proc sql;
    create table _long_nonamb as
    select *
    from _long l
    where not exists (
      select 1 from _ambig a
      where a.Dataset=l.Dataset and a.Variable=l.Variable and a.TestCD=l.TestCD
    );
  quit;

  data _byval_clause;
    length Dataset $32 Variable $64 Assigned_Value $4000
           Where_Clause $20000 base $16 _list $20000
           tcd_suffix $2000 tok $200;
    do until (last.Assigned_Value);
      set _long_nonamb;
      by Dataset Variable Assigned_Value;
      if first.Assigned_Value then call missing(_list, tcd_suffix);
      _list = catx(', ', _list, cats("'", tranwrd(TestCD,"'","''"), "'"));
      tok   = upcase(compress(TestCD, , 'kad'));
      tcd_suffix = catx('_', tcd_suffix, tok);
    end;
	  %if %upcase(&prefix) eq Y %then %do;
    base = cats(upcase(substr(Variable,1,2)),'TESTCD');
	%end;

	  %if %upcase(&prefix) ne Y %then %do;
    base = 'TESTCD';
	%end;

    	if countc(_list,',') ge 1 then  Where_Clause = trim(base) || ' IN (' || strip(_list) || ')';
	else Where_Clause = trim(base) || ' EQ ' ||strip(_list);


    keep Dataset Variable Assigned_Value Where_Clause tcd_suffix;
  run;

  proc sort data=_byval_clause; by Dataset Variable Where_Clause Assigned_Value; run;
  data _counts_wc;
    length Dataset $32 Variable $64 Where_Clause $2000;
    do until (last.Where_Clause);
      set _byval_clause;
      by Dataset Variable Where_Clause;
      if first.Where_Clause then nvals=0;
      nvals+1;
    end;
    keep Dataset Variable Where_Clause nvals;
  run;

  proc sort data=_byval_clause; by Dataset Variable Where_Clause; run;
  data _merged_wc;
    merge _byval_clause(in=a) _counts_wc(in=b);
    by Dataset Variable Where_Clause;
    if a;
    length CodeList_ID $256;
    if nvals>1 then CodeList_ID = cats(Variable, '_', tcd_suffix);
    else CodeList_ID = ' ';
  run;

  proc sort data=_merged_wc; by Dataset Variable Where_Clause; run;
  data _rows_nonamb;
    length Dataset $32 Variable $64 Assigned_Value $200 _Assigned $200 CodeList_ID $256 Where_Clause $2000;
    do until (last.Where_Clause);
      set _merged_wc;
      by Dataset Variable Where_Clause;
      if first.Where_Clause then do; _CodeList_ID = CodeList_ID; _Assigned=' '; _nvals=nvals; end;
      if _nvals=1 then _Assigned = Assigned_Value;
    end;
    Dataset      = Dataset;
    Variable     = Variable;
    Where_Clause = Where_Clause;
    if _nvals=1 then do; Assigned_Value=_Assigned; CodeList_ID=' '; end;
    else               do; Assigned_Value=' ';      CodeList_ID=_CodeList_ID; end;
    output;
    drop _CodeList_ID _Assigned _nvals nvals;
  run;

  /* =================== C) Final outputs =================== */
  data &outds; set _rows_ambig _rows_nonamb; drop tcd_suffix;run;

  data &outcodes;
    length CodeList_ID $256 Value $4000;
    set
      _ambig_terms_gid
      _merged_wc (where=(CodeList_ID ne ' ') keep=CodeList_ID Assigned_Value rename=(Assigned_Value=Value))
    ;
  run;
  proc sort data=&outcodes nodupkey; by CodeList_ID Value; run;

  data &codelistds;
    length ID $256 Term $4000;
    set &outcodes(rename=(CodeList_ID=ID Value=Term));
	
  run;

proc export data=&codelistds outfile="&outpath./vlm_extra_&type..xlsx" dbms=xlsx replace;sheet='Codelists';
run;
proc export data=&outds outfile="&outpath./vlm_extra_&type..xlsx" dbms=xlsx replace;sheet='Value Level';
run;



%mend build_from_vlm_wide;

/* Run it */
%build_from_vlm_wide(
  vlm=_vlm,
  outds=Where_Clauses_ByVar,
  outcodes=CodeLists_ByVar,
  codelistds=CodeLists_Terms,
  domaincol=DOMAIN,
  exclude_suffix=RES RESU ORRES ORRESU TESTCD TEST STRESC STRESU,
  outpath=c:\users\bmant\merck_kgaa,
  type=lb,
  prefix=
);
