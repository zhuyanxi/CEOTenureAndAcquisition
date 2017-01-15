

 cd "C:\Users\Administrator\Desktop\����\����"
 set more off
 foreach file in FI_T5 FI_T51 FI_T52 FI_T53 {
   import excel using `file'.xls, firstrow clear
   foreach x of varlist _all{
    local lab=`x'[1]
	label var `x' "`lab'"
	}
	drop in 1/2
	rename (Stkcd Accper Indcd F050201B)(stkcd acc indu roa)
	destring stkcd roa,force replace
	split acc,parse("-") 
	rename acc1 year
	keep if acc2=="12"
	drop acc*
	duplicates drop  stkcd year,force
	order stkcd year
	sort stkcd year
	save `file'.dta,replace
	}
	
	use FI_T5.dta,clear
	foreach file in FI_T51 FI_T52 FI_T53 {
	 append using `file'.dta
	 }
	 destring year,force replace
	 
	 sort stkcd year
	 save fin.dta,replace

	 
	 
	 cd "C:\Users\Administrator\Desktop\����\�߹�"
	 
	  import excel using CG_Ceo.xls,firstrow clear
	   foreach x of varlist _all{
    local lab=`x'[1]
	label var `x' "`lab'"
	}
	drop in 1/2

	  
	  *-�����ܾ�������
	  keep if Position=="2"
	  drop Position
	  
	  *-������������
	  keep if Changtyp=="2"
	  drop Changtyp
	  
	  keep Stkcd Annodt Chgdt Name Edca Entele Ifagent
	  
	  *-��ְ��ʼ���
	  gen year=real(substr(Chgdt,1,4))
	  gen mon=real(substr(Chgdt,6,2))
	  replace year=year+1 if mon>6 //7-12�¸�����Ϊ����һ�����
	
	 destring Stkcd Edca -Ifagent,force replace
	  drop Annodt Chgdt
	  rename Stkcd stkcd
	  
	 *-һ������������
	 bys stkcd year : egen count=count(year)
	 
	 *-(1)��������һ�µ�����
	 duplicates report  stkcd year Name
	 bys stkcd year Name: egen min_mon=min(mon)
	 keep if mon==min_mon
	 drop min_mon
	 duplicates drop stkcd year Name,force
	 
	 *-(2)һ���θ��水�ո��������·�
	 bys stkcd year: egen max_mon=max(mon)
	 keep if mon==max_mon
	 drop max_mon
	 duplicates drop stkcd year,force
	 drop  count
	 duplicates report stkcd year
	 
	  sort stkcd year
	  order stkcd year
	
	drop mon
	
	save ceo.dta,replace
	
	*=======================
	cd "C:\Users\����\Desktop\����-20170112\����"
	 use "C:\Users\����\Desktop\����-20170112\����\fin.dta",replace
	  merge 1:1 stkcd year using "C:\Users\����\Desktop\����-20170112\�߹�\ceo.dta"
	  drop if _merge==2
	  drop _merge
	  sort stkcd year
	  
	  merge m:1 stkcd using "C:\Users\����\Desktop\����-20170112\�߹�\ipo.dta"
	  keep if _merge==3
	  drop _merge
	  
	  gen ipo_year=year(ipo)
	  *-ɾ��IPo�����2014���Ժ������
	  drop if ipo_year>=2014
	 xtset stkc year 
		
	bys stkcd :gen start_year = year if Name!=""
	forvalues i = 1/16{
     bys stkcd:replace start_year = start_year[_n - `i'] if start_year ==. 
}	
     replace start_year=ipo_year if start_year==.
     keep if year>=2000&year<=2015
	 
	 gen tenure=year-start_year+1
	 drop if year<ipo_year
	 
	 save �߹���������.dta,replace
	 

