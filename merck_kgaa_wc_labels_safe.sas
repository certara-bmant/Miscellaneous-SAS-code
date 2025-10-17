/* ========= CONFIG ========= */
%let xlsx_path = c:/users/bmant/merck_kgaa/LBGX101_IQVIALab_SafetyData_DTS_Template.xlsx;   /* <-- set your file path */
%let vlm_sheet = VLM;
options validvarname=any;

/* ========= OPEN EXCEL & READ VLM ========= */
libname xl xlsx "&xlsx_path";
data _vlm; set xl.&vlm_sheet; run;

/* ========= DISCOVER VARIABLES (exclude IO*) ========= */
proc contents data=_vlm out=_vars(keep=name type varnum) noprint; run;

proc sql noprint;
  create table _keep as
  select name, type, varnum
  from _vars
  where upcase(substr(strip(name),1,2)) ne 'IO'
  order by varnum;

  /* target variables that become rows */
  create table _resvars as
  select name as resvar length=128
  from _keep
  where prxmatch('/(ORRESU|ORRES|STRESU|STRESC|STRESN|STRES)$/i', strip(name)) > 0;

  /* macro lists */
  select name, type into :v1-:v999, :t1-:t999 from _keep; %let nvars=&sqlobs;
  select resvar into :r1-:r999 from _resvars;             %let nres=&sqlobs;
quit;

/* ========= MAIN BUILDER ========= */
%macro make_valuelevel(typ=);

  %if &nres = 0 %then %do;
    %put NOTE: No _RES/_STRES variables found in VLM. Nothing to build.;
    %return;
  %end;

  data _ValueLevel_raw;
    length Dataset $32 Variable $64 Where_Clause $4000
           Label $800 Base_Label $400
           Data_Type $20 Format $40 Mandatory $3
           Assigned_Value $400 Codelist $256 Comment $1000
           _prefix $64
           _testvar $64 _specvar $64 _methodvar $64 _anmethvar $64
           _evalvar $64 _evalvar_alt $64
           _namvar $64 _testcdvar $64 _catvar $64 _scatvar $64
           _tstdtlvar $64 _tstdtlvar_alt $64
           test_val $400 spec_val $400 method_val $400 anmeth_val $400
           eval_val $400 nam_val $400 testcd_val $400
           cat_val $400 scat_val $400 tstdtl_val $400
           _have_test 8 _have_spec 8 _have_method 8 _have_anmeth 8
           _have_eval 8 _have_nam 8 _have_testcd 8
           _have_cat 8 _have_scat 8 _have_tstdtl 8
           ;
    set _vlm;

    /* ---- Initialize optional columns ---- */
    Base_Label = ''; Data_Type = ''; Format = ''; Mandatory = ''; Comment = '';

    /* ---- Hash of existing column names ---- */
    if _n_=1 then do;
      declare hash hv(dataset:"_vars(keep=name)");
      hv.defineKey('name');
      hv.defineData('name');
      hv.defineDone();
    end;
    length name $64;

    /* Dataset from DOMAIN if present */
    Dataset = strip(vvaluex('DOMAIN'));

    /* ---------- Where Clause (kept as Where_Clause; rename at the end) ---------- */
    length Where_Clause $4000 piece $600 valc $600 _vnc $128 _x $128;
    Where_Clause = "";
%macro build_where;
  %do i=1 %to &nvars;
    %let _vn=&&v&i; %let _tp=&&t&i;
    _vnc = "&&v&i";
    _x   = strip(upcase(_vnc));

    /* exclude any *RESU variables */
    if lengthn(_x) < 4 or substrn(_x, lengthn(_x)-3) ne 'RESU' then do;
      valc = strip(vvaluex(_vnc));
      if valc='.' then valc='';

      /* retain original rule: skip parentheses etc. */
      if not missing(valc)
         and substr(valc,1,1) ne '('
         and indexc(valc,'<>')=0
         and substr(valc,1,4) ne 'e.g.' then do;

        /* add quotes only if there are 2+ words */
        if countw(valc) > 1 then do;
          /* optional: escape internal quotes */
          /* valc = tranwrd(valc,'"','""'); */
          valc = quote(trim(valc));
        end;

        %if &_tp=2 %then %do;
          piece = catx(' ', _vnc, 'EQ', trim(valc));
        %end;
        %else %do;
          piece = catx(' ', _vnc, 'EQ', valc);
        %end;

        if Where_Clause="" then Where_Clause=piece;
        else Where_Clause=catx(' and ', Where_Clause, piece);
      end;
    end;
  %end;
%mend;
%build_where;

    /* ---------- emit one row for each target variable ---------- */
    %macro emit_rows;
      %do j=1 %to &nres;
        Variable = "&&r&j";

        /* suffix/prefix detection */
        length _xname $64 _sfx $7 data_type $8;
        _xname = strip(upcase(Variable));
        select;
          when (lengthn(_xname)>=6 and substrn(_xname,lengthn(_xname)-5)='ORRESU') _sfx='ORRESU';
          when (lengthn(_xname)>=6 and substrn(_xname,lengthn(_xname)-5)='STRESU') _sfx='STRESU';
          when (lengthn(_xname)>=6 and substrn(_xname,lengthn(_xname)-5)='STRESN') _sfx='STRESN';
          when (lengthn(_xname)>=6 and substrn(_xname,lengthn(_xname)-5)='STRESC') _sfx='STRESC';
          when (lengthn(_xname)>=5 and substrn(_xname,lengthn(_xname)-4)='ORRES')  _sfx='ORRES';
          when (lengthn(_xname)>=5 and substrn(_xname,lengthn(_xname)-4)='STRES')  _sfx='STRES';
          otherwise _sfx='';
        end;

        if _sfx ne '' then _prefix = substrn(Variable,1,lengthn(Variable)-lengthn(_sfx)); else _prefix='';

        /* companion var names (domain-specific first) */
        _testvar    = cats(_prefix,'TEST');
		 _symvar    = cats(_prefix,'SYM');
		  _gensrvar    = cats(_prefix,'GENSR');
		_tstopovar		= cats(_prefix,'TSTOPO');
		_TSTCNDvar		= cats(_prefix,'TSTCND');
		_BDAGNTvar		= cats(_prefix,'BDAGNT');
		_CNDAGTvar		= cats(_prefix,'CNDAGT');
        _specvar    = cats(_prefix,'SPEC');
        _methodvar  = cats(_prefix,'METHOD');
        _anmethvar  = cats(_prefix,'ANMETH');
        _evalvar    = cats(_prefix,'EVAL');     /* prefer EGEVAL */
        _evalvar_alt= 'EVAL';                   /* fallback if only generic exists */
        _namvar     = cats(_prefix,'NAM');
        _testcdvar  = cats(_prefix,'TESTCD');
        _catvar     = cats(_prefix,'CAT');
        _scatvar    = cats(_prefix,'SCAT');
        _tstdtlvar  = cats(_prefix,'TSTDTL');   /* prefer EGTSTDTL */
        _tstdtlvar_alt = 'TSTDTL';              /* fallback */
		_MRKSTRvar = cats(_prefix,'MRKSTR');
		_GATEvar = cats(_prefix,'GATE');
		_TSTPNLvar = cats(_prefix,'TSTPNL');
		_ANTREGvar = cats(_prefix,'ANTREG');
		_SPCCNDvar = cats(_prefix,'SPCCND');

        /* existence flags via hash (domain-specific) */
        name = _testvar;    _have_test   = (hv.find()=0);
        name = _specvar;    _have_spec   = (hv.find()=0);
        name = _methodvar;  _have_method = (hv.find()=0);
        name = _anmethvar;  _have_anmeth = (hv.find()=0);
        name = _evalvar;    _have_eval   = (hv.find()=0);
        name = _namvar;     _have_nam    = (hv.find()=0);
        name = _testcdvar;  _have_testcd = (hv.find()=0);
        name = _catvar;     _have_cat    = (hv.find()=0);
        name = _scatvar;    _have_scat   = (hv.find()=0);
        name = _tstdtlvar;  _have_tstdtl = (hv.find()=0);
		name = _tstopovar;  _have_tstopo = (hv.find()=0);
		name = _BDAGNTvar;  _have_BDAGNT = (hv.find()=0);
		name = _CNDAGTvar;  _have_CNDAGT = (hv.find()=0);
		name = _TSTCNDvar;  _have_TSTCND = (hv.find()=0);
		name = _SYMvar;  _have_SYM = (hv.find()=0);
		name = _GENSRvar;  _have_GENSR = (hv.find()=0);
	    name = _MRKSTRvar;  _have_MRKSTR = (hv.find()=0);
	    name = _GATEvar;  _have_GATE = (hv.find()=0);
	    name = _TSTPNLvar;  _have_TSTPNL = (hv.find()=0);
	    name = _ANTREGvar;  _have_ANTREG = (hv.find()=0);
	    name = _SPCCNDvar;  _have_SPCCND = (hv.find()=0);


		/* fallbacks to generic names for whitelistable facets */

        if not _have_eval   then do; name=_evalvar_alt;   if hv.find()=0 then do; _have_eval=1;   _evalvar=_evalvar_alt;   end; end;
        if not _have_tstdtl then do; name=_tstdtlvar_alt; if hv.find()=0 then do; _have_tstdtl=1; _tstdtlvar=_tstdtlvar_alt; end; end;

        /* safe value pulls */
        if _have_test   then test_val    = strip(vvaluex(_testvar));    else call missing(test_val);
        if _have_spec   then spec_val    = strip(vvaluex(_specvar));    else call missing(spec_val);
        if _have_method then method_val  = strip(vvaluex(_methodvar));  else call missing(method_val);
        if _have_anmeth then anmeth_val  = strip(vvaluex(_anmethvar));  else call missing(anmeth_val);
        if _have_eval   then eval_val    = strip(vvaluex(_evalvar));    else call missing(eval_val);
        if _have_nam    then nam_val     = strip(vvaluex(_namvar));     else call missing(nam_val);
        if _have_testcd then testcd_val  = strip(vvaluex(_testcdvar));  else call missing(testcd_val);
        if _have_cat    then cat_val     = strip(vvaluex(_catvar));     else call missing(cat_val);
        if _have_scat   then scat_val    = strip(vvaluex(_scatvar));    else call missing(scat_val);
        if _have_tstdtl then tstdtl_val  = strip(vvaluex(_tstdtlvar));  else call missing(tstdtl_val);
		if _have_tstopo then tstopo_val  = strip(vvaluex(_tstopovar));  else call missing(tstopo_val);
 		if _have_BDAGNT then BDAGNT_val  = strip(vvaluex(_BDAGNTvar));  else call missing(BDAGNT_val);
		if _have_CNDAGT then CNDAGT_val  = strip(vvaluex(_CNDAGTvar));  else call missing(CNDAGT_val);
		if _have_TSTCND then TSTCND_val  = strip(vvaluex(_TSTCNDvar));  else call missing(TSTCND_val);
		if _have_SYM then sym_val  = strip(vvaluex(_symvar));  else call missing(SYM_val);
		if _have_GENSR then GENSR_val  = strip(vvaluex(_GENSRvar));  else call missing(GENSR_val);
		if _have_MRKSTR then MRKSTR_val  = strip(vvaluex(_MRKSTRvar));  else call missing(MRKSTR_val);
		if _have_GATE then GATE_val  = strip(vvaluex(_GATEvar));  else call missing(GATE_val);
		if _have_TSTPNL then TSTPNL_val  = strip(vvaluex(_TSTPNLvar));  else call missing(TSTPNL_val);
		if _have_ANTREG then ANTREG_val  = strip(vvaluex(_ANTREGvar));  else call missing(ANTREG_val);
		if _have_SPCCND then SPCCND_val  = strip(vvaluex(_SPCCNDvar));  else call missing(SPCCND_val);

        /* variable's own value / codelist assign */
        length val_res $600 base_name $400;
        val_res = strip(vvaluex(Variable));
        call missing(Assigned_Value, Codelist);

        /* base label: prefer TEST, then TESTCD, else the variable */
        if _have_test   and not missing(test_val)   then base_name = test_val;
        else if _have_testcd and not missing(testcd_val) then base_name = testcd_val;
        else base_name = Variable;

        /* Assigned Value / Codelist */
        retain rx_sq rx_par;
        if _n_=1 then do;
          rx_sq  = prxparse('/\[(.*?)\]/');
          rx_par = prxparse('/\((.*?)\)/');
        end;
        if not missing(val_res) and indexc(val_res,'<>')=0  then do;
          if prxmatch(rx_sq, val_res) then Codelist = strip(prxposn(rx_sq,1,val_res));
          else if prxmatch(rx_par, val_res) then Codelist = strip(prxposn(rx_par,1,val_res));
          else Assigned_Value = val_res;
        end;

        /* emit rule: non-RESU always; RESU only when Assigned_Value present */
        _is_resu = (lengthn(_xname)>=4 and substrn(_xname,lengthn(_xname)-3)='RESU');
        if _is_resu=0 then output;
        else if not missing(Assigned_Value) then output;

		*if not missing(codelist) or not missing(assigned_value) then data_type='text';



      %end;  /* j loop */
    %mend;
    %emit_rows;
  run;

  /* stable row id (no longer used for label suffix) */
  data _ValueLevel_raw; set _ValueLevel_raw; retain RID 0; RID+1; run;

  /* ---------------- Minimal-append labeling within (Dataset, Variable) ---------------- */

  /* Whitelist facets to use for disambiguation (in order). Example: TSTDTL EVAL */
  %let facet_order = TSTDTL SCAT ANMETH ANTREG SPCCND ;

  /* Build progressive candidate labels (CATX -> proper spacing) */
  data _ValueLevel_raw;
    set _ValueLevel_raw;

    length label0-label20 $800;

    /* Base phrase */
    label0 = catx(' ', 'Measurement of', strip(base_name));

    %macro _build_labels;
      %let m=%sysfunc(countw(&facet_order,%str( )));
      %do k=1 %to &m;
        %let facet=%upcase(%scan(&facet_order,&k));
        label&k = label%eval(&k-1);

        %if &facet = TSTDTL %then %do;
          if not missing(tstdtl_val) then label&k = catx(' ', label%eval(&k-1), 'by', strip(tstdtl_val));
        %end;
        %else %if &facet = EVAL %then %do;
          if not missing(eval_val) then label&k = catx(' ', label%eval(&k-1), 'by', strip(eval_val));
        %end;
        %else %if &facet = METHOD %then %do;
          if not missing(method_val) then label&k = catx(' ', label%eval(&k-1), 'by', strip(method_val));
        %end;
        %else %if &facet = ANMETH %then %do;
          if not missing(anmeth_val) then label&k = catx(' ', label%eval(&k-1), 'by', strip(anmeth_val));
        %end;
        %else %if &facet = NAM %then %do;
          if not missing(nam_val) then label&k = catx(' ', label%eval(&k-1), 'by', strip(nam_val));
        %end;
        %else %if &facet = SPEC %then %do;
          if not missing(spec_val) then label&k = catx(' ', label%eval(&k-1), 'in', strip(spec_val));
        %end;
        %else %if &facet = SCAT %then %do;
          if not missing(scat_val) then label&k = catx(' ', label%eval(&k-1), 'in', strip(scat_val));
        %end;
        %else %if &facet = CAT %then %do;
          if not missing(cat_val) then label&k = catx(' ', label%eval(&k-1), 'in', strip(cat_val));
        %end;
		%else %if &facet = TSTOPO %then %do;
          if not missing(tstopo_val) then label&k = catx(' ', label%eval(&k-1), 'by', strip(tstopo_val));
        %end;
        %else %if &facet = BDAGNT %then %do;
          if not missing(BDAGNT_val) then label&k = catx(' ', label%eval(&k-1), 'for', strip(BDAGNT_val));
        %end;
        %else %if &facet = CNDAGT %then %do;
          if not missing(CNDAGT_val) then label&k = catx(' ', label%eval(&k-1), 'for', strip(CNDAGT_val));
        %end;
		%else %if &facet = TSTCND %then %do;
          if not missing(TSTCND_val) then label&k = catx(' ', label%eval(&k-1), ' ', strip(TSTCND_val));
        %end;
		 %else %if &facet = SYM %then %do;
          if not missing(SYM_val) then label&k = catx(' ', label%eval(&k-1), 'for', strip(SYM_val));
        %end;
		 %else %if &facet = GENSR %then %do;
          if not missing(GENSR_val) then label&k = catx(' ', label%eval(&k-1), 'in', strip(GENSR_val));
        %end;
		%else %if &facet = MRKSTR %then %do;
          if not missing(MRKSTR_val) then label&k = catx(' ', label%eval(&k-1), 'for', strip(MRKSTR_val));
        %end;
		%else %if &facet = GATE %then %do;
          if not missing(GATE_val) then label&k = catx(' ', label%eval(&k-1), 'by', strip(GATE_val));
        %end;
		%else %if &facet = TSTPNL %then %do;
          if not missing(TSTPNL_val) then label&k = catx(' ', label%eval(&k-1), 'by', strip(TSTPNL_val));
        %end;

		%else %if &facet = ANTREG %then %do;
         if not missing(ANTREG_val) then label&k = catx(' ', label%eval(&k-1), 'in', strip(ANTREG_val));
        %end;
				%else %if &facet = SPCCND %then %do;
         if not missing(SPCCND_val) then label&k = catx(' ', label%eval(&k-1), 'in', strip(SPCCND_val),' ',strip(SPEC_val));
        %end;


      %end;
    %mend;
    %_build_labels
  run;

  /* Count duplicates per stage within (Dataset, Variable, labelK) ONLY */
  %macro _mk_counts(k);
  proc sql noprint;
    create table _cnt&k as
    select Dataset, Variable, label&k as labelK, count(*) as cnt&k
    from _ValueLevel_raw
    group by 1,2,3;
  quit;

  proc sort data=_cnt&k;          by Dataset Variable labelK; run;
  proc sort data=_ValueLevel_raw; by Dataset Variable label&k; run;

  data _ValueLevel_raw;
    merge _ValueLevel_raw(in=a)
          _cnt&k(in=b rename=(labelK=label&k));
    by Dataset Variable label&k;
    if a;
  run;
  %mend;

  /* Run counts for label0..labelM */
  %let __m=%sysfunc(countw(&facet_order,%str( )));
  %macro _run_counts;
    %_mk_counts(0)
    %do __i=1 %to &__m;
      %_mk_counts(&__i)
    %end;
  %mend;
  %_run_counts

  /* Pick the shortest unique within (Dataset, Variable); else keep longest allowed (no record_ suffix) */
  data _ValueLevel_raw;
    set _ValueLevel_raw;
    length Label $800;

    if cnt0=1 then Label = label0;
    %macro _pick;
      %do __i=1 %to &__m;
        else if cnt&__i=1 then Label = label&__i;
      %end;
    %mend; %_pick
    else Label = coalescec(label&__m, label0);

    drop label0-label20 cnt0-cnt20;
  run;

  /* ========= Finalize & Write ========= */
  data ValueLevel;
    set _ValueLevel_raw;
	format data_type $10.;
  
    /* NOTE: Final uniqueness key is Dataset, Variable, Label */
	 if &typ.orresu ne '' then data_type='float';

		   if Assigned_Value ne '' or codelist ne '' then data_type='text';
		if data_type='' then data_type='text';

		     rename Where_Clause   = 'Where Clause'n
           Data_Type      = 'Data Type'n
           Assigned_Value = 'Assigned Value'n;
    keep Dataset Variable Where_Clause Label Assigned_Value Codelist Comment data_type;
  run;

  proc sort data=ValueLevel /*nodupkey*/;
    by Dataset Variable Label;
  run;

  /* replace the ValueLevel sheet */
  proc datasets lib=xl nolist; delete ValueLevel; quit;

proc export data= valuelevel outfile="c:/users/bmant/merck_kgaa/valuelevel_&typ..xlsx" dbms=xlsx replace;
run;
%mend make_valuelevel;

/* Run it */



*%starting;
%make_valuelevel(typ=safe);
*%ending;

libname xl clear;
