@echo off
REM **** Automated Report parameters
set SnapshotReport=%1\%4_SnapshotReport_%5.csv
set SummaryReport=%1\%4_SummaryReport_%5.csv
set NewViolationsReport=%1\%4_NewViolations_%5.csv
set FixedViolationsReport=%1\%4_FixedViolations_%5.csv
set ChangedViolationsReport=%1\%4_ChangedViolations_%5.csv
set separator=;

set ApplicationName=%4
set DashboardService=%3
set CSSServer=%6
set userLogin=%7
set database=%8
set portNumber=%9
@for /L %%i in (0,1,8) do shift
set EDURL=%1
set CSSPSQL=%2

REM if newSnapshot == 0 then added violations will be calculated from the current snapshot and the previous snapshots,
REM if you have more than 2 snapshots and you want to calculate added violations between 2 snapshots, set snapshot_id for new snapshot and snapshot_id for old snapshot
REM To have the list of snapshot_id available for your application, connect to the central base and execute the following query: select * from dss_snapshots
set newSnapshot=0
set oldSnapshot=0
set PGPASSWORD=CastAIP


REM creation of structure to store added violations---------------------------------------------------------------
set requete="DO LANGUAGE plpgsql $$ BEGIN IF EXISTS (SELECT * FROM information_schema.tables WHERE  table_schema = '%DashboardService%'  AND table_name = 'aaa_recent_variations') THEN TRUNCATE TABLE %DashboardService%.AAA_RECENT_VARIATIONS; DROP TABLE %DashboardService%.AAA_RECENT_VARIATIONS; END IF; END; $$;"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber% -U %userLogin% -t %database%

set requete="CREATE TABLE %DashboardService%.AAA_RECENT_VARIATIONS (
set requete=%requete%		  SNAPSHOT_ID                                   integer              NOT NULL,
set requete=%requete%		  DIAG_ID                                       integer              NOT NULL,
set requete=%requete%		  OBJECT_ID                                     integer              NOT NULL,
set requete=%requete%		  MODULE_ID                                     integer              NOT NULL,
set requete=%requete%		  APP_ID                                        integer              NOT NULL,
set requete=%requete%		  PREVIOUS_SCORE                                integer              NOT NULL,
set requete=%requete%		  CURRENT_SCORE                                 integer              NOT NULL,
set requete=%requete%		  VIOLATION_STATUS                              integer              NOT NULL,
set requete=%requete%		  VIOLATION_STATUS_DESC                         VARCHAR(30)	    	 NOT NULL,
set requete=%requete%		  OBJECT_STATUS                                 integer              NOT NULL);"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

set requete="DO LANGUAGE plpgsql $$ BEGIN IF EXISTS (SELECT * FROM information_schema.tables WHERE  table_schema = '%DashboardService%'  AND table_name = 'aaa_object_statuses') THEN TRUNCATE TABLE %DashboardService%.AAA_OBJECT_STATUSES; DROP TABLE %DashboardService%.AAA_OBJECT_STATUSES; END IF; END; $$;"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 

set requete="CREATE TABLE %DashboardService%.AAA_OBJECT_STATUSES (	  
set requete=%requete%		  OBJECT_ID                                     integer           NOT NULL,
set requete=%requete%		  OBJECT_STATUS                                 integer           NOT NULL,
set requete=%requete%		  OBJECT_STATUS_DESC                            VARCHAR(10)       NOT NULL);"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

echo TABLES CREATED

rem filling tables with data ---------------------------------------------------------------
if  %newSnapshot% NEQ 0 goto :ReportPreparation

echo Report on last 2 snapshots

set newSnapshot="select snapshot_id from %DashboardService%.dss_snapshots s, %DashboardService%.dss_objects o where o.object_type_id = -102	and upper(o.object_name) = upper('%ApplicationName%') and s.application_id = o.object_id order by FUNCTIONAL_DATE desc limit 1"
set command=%CCSPSQL%  -c %newSnapshot% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
for /f "delims=" %%a in ('%command%') do set newSnapshot=%%a
for /f "tokens=* delims= " %%a in ("%newSnapshot%") do set newSnapshot=%%a
echo New Snapshot = %newSnapshot%

set oldSnapshot=select snapshot_id from %DashboardService%.dss_snapshots s, %DashboardService%.dss_objects o where o.object_type_id = -102	and upper(o.object_name) = upper('%ApplicationName%') and s.application_id = o.object_id and snapshot_id ^<^> %newSnapshot% order by FUNCTIONAL_DATE desc limit 1
set command=%CCSPSQL%  -c "%oldSnapshot%" -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database%
for /f "delims=" %%a in ('%command%') do set oldSnapshot=%%a
for /f "tokens=* delims= " %%a in ("%oldSnapshot%") do set oldSnapshot=%%a
echo Old Snapshot = %oldSnapshot%

:ReportPreparation

echo Report on Snapshot %oldSnapshot% and %newSnapshot% 

REM Added Objects
set requete="insert 	into %DashboardService%.AAA_OBJECT_STATUSES 
set requete=%requete%	select  OBJECT_ID, 1, 'Added' 
set requete=%requete%	from 	%DashboardService%.DSS_OBJECT_INFO 
set requete=%requete%	where 	SNAPSHOT_ID = %newSnapshot% 
set requete=%requete%	and 	OBJECT_ID not in (select OBJECT_ID from %DashboardService%.DSS_OBJECT_INFO where SNAPSHOT_ID = %oldSnapshot%);"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

REM Removed Objects
set requete="insert 	into %DashboardService%.AAA_OBJECT_STATUSES
set requete=%requete%	select 	OBJECT_ID, 2, 'Removed' 
set requete=%requete%	from 	%DashboardService%.DSS_OBJECT_INFO
set requete=%requete%	where 	SNAPSHOT_ID = %oldSnapshot%
set requete=%requete%	and 	OBJECT_ID not in (select OBJECT_ID from %DashboardService%.DSS_OBJECT_INFO where SNAPSHOT_ID = %newSnapshot%);"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

REM Updated Objects
set requete="insert 	into %DashboardService%.AAA_OBJECT_STATUSES
set requete=%requete%	select 	s1.OBJECT_ID, 3, 'Updated'
set requete=%requete%	from 	%DashboardService%.DSS_OBJECT_INFO s1, %DashboardService%.DSS_OBJECT_INFO s2 
set requete=%requete%	where 	s1.SNAPSHOT_ID = %newSnapshot%
set requete=%requete%	and		s2.SNAPSHOT_ID = %oldSnapshot%
set requete=%requete%	and 	s1.OBJECT_ID = s2.OBJECT_ID 
set requete=%requete%	and		s1.OBJECT_CHECKSUM <> s2.OBJECT_CHECKSUM;"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

REM Unchanged Objects
set requete="insert 	into %DashboardService%.AAA_OBJECT_STATUSES
set requete=%requete%	select 	s1.OBJECT_ID, 3, 'Unchanged'
set requete=%requete%	from 	%DashboardService%.DSS_OBJECT_INFO s1, %DashboardService%.DSS_OBJECT_INFO s2 
set requete=%requete%	where 	s1.SNAPSHOT_ID = %newSnapshot%
set requete=%requete%	and		s2.SNAPSHOT_ID = %oldSnapshot%
set requete=%requete%	and 	s1.OBJECT_ID = s2.OBJECT_ID 
set requete=%requete%	and		s1.OBJECT_CHECKSUM <> 0
set requete=%requete%	and		s1.OBJECT_CHECKSUM = s2.OBJECT_CHECKSUM;"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

REM N\A Status Objects
set requete="insert 	into %DashboardService%.AAA_OBJECT_STATUSES
set requete=%requete%	select 	s1.OBJECT_ID, 3, 'N/A'
set requete=%requete%	from 	%DashboardService%.DSS_OBJECT_INFO s1 
set requete=%requete%	where 	s1.SNAPSHOT_ID = %newSnapshot%
set requete=%requete%	and 	OBJECT_ID not in (select OBJECT_ID from %DashboardService%.AAA_OBJECT_STATUSES)
set requete=%requete%	and		s1.OBJECT_CHECKSUM = 0;"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

set requete="TRUNCATE TABLE %DashboardService%.AAA_RECENT_VARIATIONS"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 

REM New Violation
set requete="insert into %DashboardService%.AAA_RECENT_VARIATIONS
set requete=%requete%	SELECT DISTINCT	pt.SNAPSHOT_ID, res.DIAG_ID, res.OBJECT_ID, pt.MODULE_ID, pt.APP_ID, -1, Round(CAST(res.DIAG_VALUE as numeric),2), 1, 'New Violation' ,st.OBJECT_STATUS
set requete=%requete%	FROM  	%DashboardService%.DSS_PORTF_TREE pt, %DashboardService%.CSV_DIAGDETAILS res, %DashboardService%.AAA_OBJECT_STATUSES st
set requete=%requete%	where	pt.APP_ID 		= (select distinct OBJECT_ID from %DashboardService%.DSS_OBJECTS o where o.OBJECT_TYPE_ID = -102 and upper(o.object_name) = upper('%ApplicationName%'))
set requete=%requete%	AND		pt.SNAPSHOT_ID 	= %newSnapshot%
set requete=%requete%	AND		pt.MODULE_ID 	= res.CONTEXT_ID
set requete=%requete%	AND		res.SNAPSHOT_ID = %newSnapshot%
set requete=%requete%	AND		pt.SNAPSHOT_ID	= res.SNAPSHOT_ID
set requete=%requete%	AND		res.OBJECT_ID	= st.OBJECT_ID
set requete=%requete%	And 	not exists (Select 1 From %DashboardService%.CSV_DIAGDETAILS res2
set requete=%requete%			Where res2.SNAPSHOT_ID = %oldSnapshot% 
set requete=%requete%	And res2.DIAG_ID = res.DIAG_ID And res2.OBJECT_ID = res.OBJECT_ID  );"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

REM Fixed Violation
set requete="insert into %DashboardService%.AAA_RECENT_VARIATIONS
set requete=%requete%	SELECT DISTINCT	pt.SNAPSHOT_ID, res.DIAG_ID, res.OBJECT_ID, pt.MODULE_ID, pt.APP_ID, Round(CAST(res.DIAG_VALUE as numeric),2), -1, 2, 'Fixed Violation' ,st.OBJECT_STATUS
set requete=%requete%	FROM  	%DashboardService%.DSS_PORTF_TREE pt, %DashboardService%.CSV_DIAGDETAILS res, %DashboardService%.AAA_OBJECT_STATUSES st
set requete=%requete%	where	pt.APP_ID 		= (select distinct OBJECT_ID from %DashboardService%.DSS_OBJECTS o where o.OBJECT_TYPE_ID = -102 and upper(o.object_name) = upper('%ApplicationName%'))
set requete=%requete%	AND		pt.SNAPSHOT_ID 	= %newSnapshot%
set requete=%requete%	AND		pt.MODULE_ID 	= res.CONTEXT_ID
set requete=%requete%	AND		res.SNAPSHOT_ID = %oldSnapshot%
set requete=%requete%	AND		res.OBJECT_ID	= st.OBJECT_ID
set requete=%requete%	And 	not exists (Select 1 From %DashboardService%.CSV_DIAGDETAILS res2
set requete=%requete%			Where res2.SNAPSHOT_ID = %newSnapshot% And res2.DIAG_ID = res.DIAG_ID And res2.OBJECT_ID = res.OBJECT_ID );"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

REM Changed Violation
set requete="insert into %DashboardService%.AAA_RECENT_VARIATIONS
set requete=%requete%	SELECT DISTINCT	pt.SNAPSHOT_ID, res1.DIAG_ID, res1.OBJECT_ID, pt.MODULE_ID, pt.APP_ID, Round(cast(res1.DIAG_VALUE as numeric),2) , Round(cast(res2.DIAG_VALUE as numeric),2), 3, 'Changed Violation' ,st.OBJECT_STATUS
set requete=%requete%	FROM  	%DashboardService%.DSS_PORTF_TREE pt, %DashboardService%.CSV_DIAGDETAILS res1
set requete=%requete%	INNER JOIN 	%DashboardService%.CSV_DIAGDETAILS res2 	ON res1.OBJECT_ID 	= res2.OBJECT_ID AND res1.DIAG_ID = res2.DIAG_ID, 
set requete=%requete%	%DashboardService%.AAA_OBJECT_STATUSES st
set requete=%requete%	WHERE	pt.APP_ID 			= (select distinct OBJECT_ID from %DashboardService%.DSS_OBJECTS o where o.OBJECT_TYPE_ID = -102 and upper(o.object_name) = upper('%ApplicationName%'))
set requete=%requete%	AND		pt.SNAPSHOT_ID 		= %newSnapshot%
set requete=%requete%	AND		res1.SNAPSHOT_ID 	= %oldSnapshot%
set requete=%requete%	AND		pt.MODULE_ID 		= res1.CONTEXT_ID
set requete=%requete%	AND		pt.MODULE_ID 		= res2.CONTEXT_ID
set requete=%requete%	AND		res2.SNAPSHOT_ID 	= %newSnapshot%
set requete=%requete%	AND		pt.SNAPSHOT_ID		= res2.SNAPSHOT_ID
set requete=%requete%	AND		res2.OBJECT_ID		= st.OBJECT_ID
set requete=%requete%	AND  	Round(cast(res2.DIAG_VALUE as numeric),2)  -   Round(cast(res1.DIAG_VALUE as numeric),2) != 0;"
%CCSPSQL%  -c %requete% -h %CSSServer% -p %portNumber%  -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

rem Extract data into csv files ---------------------------------------------------------------

REM Snapshot Report
set requete="select 'Snapshot Id', 'Snapshot Order', 'Application Name', 'Version Name', 'Capture Date','Snapshot Date','CAST Release', 'Snapshot Name', 'Snapshpot Description' 
set requete=%requete% UNION
set requete=%requete% select CAST(s.snapshot_id as varchar), 'Last Snapshot', o.object_name, i.object_version, CAST(s.functional_date as varchar), CAST(s.snapshot_date as varchar), CAST(s.version as varchar), s.snapshot_name, s.snapshot_description
set requete=%requete% from %DashboardService%.dss_snapshots s, %DashboardService%.dss_objects o, %DashboardService%.dss_snapshot_info i
set requete=%requete% where o.object_type_id = -102
set requete=%requete% and upper(o.object_name) = upper('%ApplicationName%')
set requete=%requete% and o.object_id = s.application_id
set requete=%requete% and s.snapshot_id = (select snapshot_id from %DashboardService%.dss_snapshots where application_id = s.application_id order by FUNCTIONAL_DATE desc limit 1)
set requete=%requete% and s.application_id = i.object_id
set requete=%requete% and s.snapshot_id = i.snapshot_id
set requete=%requete% UNION
set requete=%requete% select CAST(s.snapshot_id as varchar), 'Previous Snapshot', o.object_name, i.object_version, CAST(s.functional_date as varchar), CAST(s.snapshot_date as varchar), CAST(s.version as varchar), s.snapshot_name, s.snapshot_description
set requete=%requete% from %DashboardService%.dss_snapshots s, %DashboardService%.dss_objects o, %DashboardService%.dss_snapshot_info i
set requete=%requete% where o.object_type_id = -102
set requete=%requete% and upper(o.object_name) = upper('%ApplicationName%')
set requete=%requete% and o.object_id = s.application_id
set requete=%requete% and s.snapshot_id = (select snapshot_id from %DashboardService%.dss_snapshots where application_id = s.application_id and snapshot_id <> (select snapshot_id from %DashboardService%.dss_snapshots order by FUNCTIONAL_DATE desc limit 1) order by FUNCTIONAL_DATE desc limit 1)
set requete=%requete% and s.application_id = i.object_id
set requete=%requete% and s.snapshot_id = i.snapshot_id
set requete=%requete% order by 5 desc;" 
%CCSPSQL% -AF %separator% -c %requete% -h %CSSServer% -p %portNumber% -o %SnapshotReport% -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

REM Summary Report
set requete="select distinct 'Rule Id','Critical', 'Weight', 'Rule Name','Number of New Violations','Number of Fixed Violations','Rule Priority','Education Comment','Rule Selection Date' 
set requete=%requete% UNION
set requete=%requete% select CAST(COALESCE(nv.diag_id, fv.diag_id) as varchar),
set requete=%requete% CAST(dmtt.metric_critical as varchar), CAST(dmtt.aggregate_weight as varchar), dmd.metric_description, 
set requete=%requete% CAST(COALESCE(nv.NumNew, 0) as varchar), CAST(COALESCE(fv.NumFixed, 0) as varchar), 'N/A', 'N/A', 'N/A'
set requete=%requete% from 	(select diag_id, snapshot_id, count(object_id) as NumNew from %DashboardService%.AAA_RECENT_VARIATIONS where violation_status =1 group by diag_id, snapshot_id ) nv
set requete=%requete% FULL OUTER JOIN 
set requete=%requete% (select diag_id, snapshot_id, count(object_id) as NumFixed from %DashboardService%.AAA_RECENT_VARIATIONS where violation_status =2 group by diag_id, snapshot_id) fv on nv.diag_id = fv.diag_id,
set requete=%requete% %DashboardService%.dss_metric_descriptions dmd,
set requete=%requete% (select metric_id, metric_critical, aggregate_weight from %DashboardService%.dss_metric_type_trees a where aggregate_weight = (select max(aggregate_weight) from %DashboardService%.dss_metric_type_trees b where a.metric_id = b.metric_id)) dmtt
set requete=%requete% where	dmd.metric_id = COALESCE(nv.diag_id, fv.diag_id)
set requete=%requete% and dmd.language='ENGLISH'
set requete=%requete% and dmd.description_type_id = 0
set requete=%requete% and dmtt.metric_id = dmd.metric_id
set requete=%requete% order by 2 desc , 3 desc, 5 desc, 6 desc;"

%CCSPSQL% -AF %separator% -c %requete% -h %CSSServer% -p %portNumber% -o %SummaryReport% -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

REM List of New Violations
set requete="select distinct 'Violation Status', 'Rule ID','Critical','Weight','Rule Name', 'Object Full Name', 'Object Type','Object Status','PRI','Value','Module Name','URL' 
set requete=%requete% UNION
set requete=%requete% select distinct v.VIOLATION_STATUS_DESC, CAST(dmd.metric_id as varchar),CAST(DMTT.METRIC_CRITICAL as varchar), CAST(DMTT.AGGREGATE_WEIGHT as varchar),
set requete=%requete% replace(metric_description,',', ''), d.object_full_name, dot.object_type_name, s.object_status_desc, CAST(greatest(pri_60011, pri_60012, pri_60013, pri_60014, pri_60016) as varchar), CAST(v.CURRENT_SCORE as varchar), modu.object_name,
set requete=%requete% '%EDURL%applications/' || CAST(v.APP_ID as varchar) || '/snapshots/' || CAST(v.SNAPSHOT_ID as varchar) || '/business/60017/qualityInvestigation/0/60017/all/' || CAST(v.DIAG_ID as varchar) || '/' || CAST(v.OBJECT_ID as varchar)
set requete=%requete% from 	%DashboardService%.AAA_OBJECT_STATUSES s, %DashboardService%.dss_objects d, %DashboardService%.AAA_RECENT_VARIATIONS v LEFT OUTER JOIN  %DashboardService%.dss_pri pri on v.snapshot_id = pri.snapshot_id and v.object_id = pri.object_id, 
set requete=%requete% %DashboardService%.dss_metric_descriptions dmd, %DashboardService%.dss_object_types dot, %DashboardService%.dss_objects modu, %DashboardService%.dss_metric_type_trees DMTT
set requete=%requete% where d.object_id = s.object_id
set requete=%requete% and v.object_id = s.object_id
set requete=%requete% and v.violation_status = 1 
set requete=%requete% and dmd.metric_id = v.diag_id
set requete=%requete% and DMTT.metric_id = v.diag_id
set requete=%requete% and DMTT.aggregate_weight = (select max(aggregate_weight) from %DashboardService%.dss_metric_type_trees tt where tt.metric_id = DMTT.metric_id)
set requete=%requete% and dmd.language='ENGLISH'
set requete=%requete% and dmd.description_type_id = 0
set requete=%requete% and modu.object_id = v.module_id
set requete=%requete% and dot.object_type_id = d.object_type_id order by 3 desc, 4 desc, 5 asc, 9 desc, 6 desc;"
%CCSPSQL% -AF %separator% -c %requete% -h %CSSServer% -p %portNumber% -o %NewViolationsReport% -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

if  "%NewViolationsReport%"=="" goto :ChangedViolationsReport

REM List of Fixed Violations

set requete="select distinct 'Violation Status','Rule ID','Critical','Weight','Rule Name','Object Full Name','Object Type','Object Status','PRI','Previous Value','Module Name' 
set requete=%requete% UNION
set requete=%requete% select distinct v.VIOLATION_STATUS_DESC,CAST(dmd.metric_id as varchar), CAST(DMTT.METRIC_CRITICAL as varchar), CAST(DMTT.AGGREGATE_WEIGHT as varchar), 
set requete=%requete% replace(metric_description,',', ''), d.object_full_name, dot.object_type_name, s.object_status_desc, CAST(greatest(pri_60011, pri_60012, pri_60013, pri_60014, pri_60016) as varchar), CAST(v.PREVIOUS_SCORE as varchar), modu.object_name 
set requete=%requete% from %DashboardService%.AAA_OBJECT_STATUSES s, %DashboardService%.dss_objects d, %DashboardService%.AAA_RECENT_VARIATIONS v LEFT OUTER JOIN  %DashboardService%.dss_pri pri on v.snapshot_id = pri.snapshot_id and v.object_id = pri.object_id, 
set requete=%requete% %DashboardService%.dss_metric_descriptions dmd, %DashboardService%.dss_object_types dot, %DashboardService%.dss_objects modu, %DashboardService%.dss_metric_type_trees DMTT
set requete=%requete% where d.object_id = s.object_id
set requete=%requete% and v.object_id = s.object_id
set requete=%requete% and v.violation_status = 2 	
set requete=%requete% and dmd.metric_id = v.diag_id
set requete=%requete% and DMTT.metric_id = v.diag_id
set requete=%requete% and DMTT.aggregate_weight = (select max(aggregate_weight) from %DashboardService%.dss_metric_type_trees tt where tt.metric_id = DMTT.metric_id)
set requete=%requete% and dmd.language='ENGLISH'
set requete=%requete% and dmd.description_type_id = 0
set requete=%requete% and modu.object_id = v.module_id
set requete=%requete% and dot.object_type_id = d.object_type_id order by 3 desc, 4 desc, 5 asc, 9 desc, 6 desc;"
%CCSPSQL% -AF %separator% -c %requete% -h %CSSServer% -p %portNumber% -o "%FixedViolationsReport%" -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV

:ChangedViolationsReport

if  "%FixedViolationsReport%"=="" goto :EndReport

REM List of Changed Violations

set requete="select distinct 'Violation Status','Rule Name','Object Full Name','Object Type','Object Status','Previous Value','Current Value','Critical','Weight','Module Name' 
set requete=%requete% UNION
set requete=%requete% select distinct v.VIOLATION_STATUS_DESC, replace(metric_description,',', ''), d.object_full_name, dot.object_type_name, s.object_status_desc, CAST(v.PREVIOUS_SCORE as varchar),CAST(v.CURRENT_SCORE as varchar), CAST(DMTT.METRIC_CRITICAL as varchar), CAST(DMTT.AGGREGATE_WEIGHT as varchar), modu.object_name 
set requete=%requete% from %DashboardService%.AAA_OBJECT_STATUSES s, %DashboardService%.dss_objects d, %DashboardService%.AAA_RECENT_VARIATIONS v, %DashboardService%.dss_metric_descriptions dmd,  %DashboardService%.dss_object_types dot, %DashboardService%.dss_objects modu, %DashboardService%.dss_metric_type_trees DMTT
set requete=%requete% where d.object_id = s.object_id
set requete=%requete% and v.object_id = s.object_id
set requete=%requete% and v.violation_status = 3 	
set requete=%requete% and dmd.metric_id = v.diag_id
set requete=%requete% and DMTT.metric_id = v.diag_id
set requete=%requete% and DMTT.aggregate_weight = (select max(aggregate_weight) from %DashboardService%.dss_metric_type_trees tt where tt.metric_id = DMTT.metric_id)
set requete=%requete% and dmd.language='ENGLISH'
set requete=%requete% and dmd.description_type_id = 0
set requete=%requete% and modu.object_id = v.module_id
set requete=%requete% and dot.object_type_id = d.object_type_id order by 8 desc, 9 desc, 2 asc, 3 asc;"
%CCSPSQL% -AF %separator% -c %requete% -h %CSSServer% -p %portNumber% -o %ChangedViolationsReport% -U %userLogin% -t %database% 
if %errorlevel% NEQ 0 echo FAILING Query: %requete% & goto failedCSV
goto EndReport

:failedCSV
echo FAILED CSV Generation Batch
exit /B 1

:EndReport
echo Done
