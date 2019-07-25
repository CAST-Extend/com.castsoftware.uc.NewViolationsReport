'''
Created on July 9th 2019

@author: NBI
'''
import os
import time
import subprocess
from cast.application import publish_report
from cast.application import ApplicationLevelExtension
import logging

class Report(ApplicationLevelExtension):
    
    def __init__(self):
        pass
    
    def start_application(self, application):
        pass

    def write_to_file(self, file_path, content, erase_content=False):
        """Write end results to a provided file."""
        if erase_content:
            open(file_path, 'w').close()

        fp = open(file_path, 'a')
        fp.write(content)
        fp.close()

    def ConnectionInfos(self, mngt=None):

        conn_rslt = mngt.execute_query("""select host,port from cms_inf_store_css;""")
        if bool(conn_rslt.rowcount):
            for conn_infos in conn_rslt:
                self.host = conn_infos[0]
                self.port = conn_infos[1]

            logging.info('* CSS Host : ' + str(self.host))
            logging.info('* CSS Port : ' + str(self.port))
        pass

    def GetDashboardURL(self, mngt=None, centraldb=None):

        url_rslt = mngt.execute_query("""select portal_url from cms_inf_css_centraldb where object_name='%s';""" % centraldb)
        if bool(url_rslt.rowcount):
            for url_info in url_rslt:
                if url_info[0] is not None:
                    logging.info('* Dashboard URL : ' + str(url_info[0]))
                    self.EDURL = url_info[0]
                else:
                    logging.warning('* Dashboard URL : No Engineering Dashboard URL defined in CAST Management Studio')
                    self.EDURL = "https://[ENGINEERING-DASHBOARD-URL-TO-CHANGE]/Engineering/engineering/index.html#AED[PUT NUMBER]/"
        pass

    def HasTwoSnapshots(self, central=None, appname=None):

        snapshot_rslt = central.execute_query("""select snapshot_id from dss_snapshots s, dss_objects o 
                                                 where s.application_id = o.object_id 
                                                 and o.object_name = '%s';""" % appname)
        if snapshot_rslt.rowcount > 1:
            return True
        else:
            return False

    def ViolationsCount(self, central=None):

        count_rslt = central.execute_query("""select count(object_id) from AAA_RECENT_VARIATIONS where violation_status =1""")
        if count_rslt.rowcount > 0:
            for count_info in count_rslt:
                self.NewViolationsCount = str(count_info[0])
        else:
            self.NewViolationsCount = str(-1)

        count_rslt = central.execute_query("""select count(object_id) from AAA_RECENT_VARIATIONS where violation_status =2""")
        if count_rslt.rowcount > 0:
            for count_info in count_rslt:
                self.FixedViolationsCount = str(count_info[0])
        else:
            self.FixedViolationsCount = str(-1)
        pass

    def ConfigFilePath(self, centraldb, appname):
        #BatchPath = E:\WorkingFolder\Automated_Reports\NewViolations
        #ReportPath = E:\WorkingFolder\Automated_Reports\NewViolations\Reports
        #ReportLogPath = E:\WorkingFolder\Automated_Reports\NewViolations\Logs
        #ApplicationName = Webgoat
        #DashboardService = webgoat_830_1362_central
        #CSSServer = localhost
        #userLogin = operator
        #database = postgres
        #portNumber = 2280
        #EDURL = https: // demo - eu.castsoftware.com / Engineering / engineering / index.html  # AED5/
        #CCSPSQL = E:\WorkingFolder\Automated_Reports\NewViolations\PSQL\psql.exe
        #SevenZip = "C:\Program Files\7-Zip\7z.exe"
        #StatusMail = n.bidaux @ castsoftware.com
        #ReportMail = n.bidaux @ castsoftware.com

        self.config_file_path = os.path.join(self.rptdir, "NewViolations_%s.config" % (appname))
        self.batch_path = os.path.join(self.get_plugin().get_plugin_directory(), 'Batch')
        zip_path=os.path.join(self.batch_path, '7-Zip\\7z.exe')
        psql_path=os.path.join(self.batch_path, 'PSQL\\psql.exe')
        self.write_to_file(self.config_file_path, 'BatchPath=%s' % self.batch_path, True)
        self.write_to_file(self.config_file_path, '\nReportPath=%s' % self.rptdir)
        self.write_to_file(self.config_file_path, '\nReportLogPath=%s' % self.rptdir)
        self.write_to_file(self.config_file_path, '\nApplicationName=%s' % appname)
        self.write_to_file(self.config_file_path, '\nDashboardService=%s' % centraldb)
        self.write_to_file(self.config_file_path, '\nCSSServer=%s' % self.host)
        self.write_to_file(self.config_file_path, '\nuserLogin=operator')
        self.write_to_file(self.config_file_path, '\ndatabase=postgres')
        self.write_to_file(self.config_file_path, '\nportNumber=%s' % self.port)
        #self.write_to_file(self.config_file_path, '\nEDURL=https://[ENGINEERING-DASHBOARD-URL-TO-CHANGE]/Engineering/engineering/index.html#AED[PUT THE NUMBER HERE]/')
        self.write_to_file(self.config_file_path, '\nEDURL=%s' % self.EDURL)
        self.write_to_file(self.config_file_path, '\nCCSPSQL=%s' % psql_path)
        self.write_to_file(self.config_file_path, '\nSevenZip=%s' % zip_path)
        self.write_to_file(self.config_file_path, '\nStatusMail=n.bidaux@castsoftware.com')
        self.write_to_file(self.config_file_path, '\nReportMail=n.bidaux@castsoftware.com')
    pass

    def GenerateReport(self):
        logging.info('* Command line : ' + self.batch_path + '\LaunchReportGeneration.bat ' + self.config_file_path)
        return_code = subprocess.call([self.batch_path + '\LaunchReportGeneration.bat', self.config_file_path])
        return return_code
    pass
    
    def end_application(self, application):   
        pass

    def after_snapshot(self, application):    
        kb = application.get_application_configuration().get_analysis_service()
        mngt = application.get_managment_base()
        central = application.get_central()

        appnames = kb.get_applications()
        appname = appnames[0].name
        centraldb = kb.name.replace("local", "central")
        logging.info('***** New Violations Report Generation *****')
        logging.info('* Application name : ' + str(appname))
        logging.info('* Management schema name : ' + mngt.name)
        logging.info('* Local schema name : ' + kb.name)
        logging.info('* Central schema name : ' + centraldb)

        self.rptdir = os.path.join(self.get_plugin().intermediate, "NewViolations")
        if not os.path.exists(self.rptdir):
            os.mkdir(self.rptdir)

        logging.info('* Report path : ' + self.rptdir)

        self.ConnectionInfos(mngt)
        self.GetDashboardURL(mngt, centraldb)
        self.ConfigFilePath(centraldb, appname)
        logging.info('* Configuration File path : ' + self.config_file_path)

        # status = "Warning"
        # status = "KO"
        status = "OK"

        if self.HasTwoSnapshots(central, appname):
            return_code = self.GenerateReport()
            if return_code == 0:
                latest_report_cdate = 0.0
                for file in os.listdir(self.rptdir):
                    if (file.endswith('xlsx')):
                        current_date = os.path.getctime(os.path.join(self.rptdir, file))
                        if latest_report_cdate < current_date:
                            latest_report = file
                            latest_report_cdate = current_date
                self.ViolationsCount(central)
                logging.info('* Excel report generated : ' + os.path.join(self.rptdir, latest_report))
                msg = "New Violations : " + self.NewViolationsCount + ' - Fixed violations : ' + self.FixedViolationsCount
                logging.info('* ' + msg)
                latest_report = os.path.join(self.rptdir, latest_report)
                os.rename(latest_report, latest_report.replace(" CI Review ", " NV Review "))
                latest_report = latest_report.replace(" CI Review ", " NV Review ")
                publish_report('New Violations Report', status, "List of the New Violations between the last 2 snapshots", msg, detail_report_path=latest_report)
            else:
                status = "KO"
                logging.warning('* Excel report generated : ERROR')
                publish_report('New Violations Report', status, "List of the New Violations between the last 2 snapshots", 'Error in the report generation', detail_report_path='')
        else:
            status = "Warning"
            logging.warning('* Excel report generated : report not generated because there is only one snapshot')
            publish_report('New Violations Report', status, "List of the New Violations between the last 2 snapshots", 'No report : only one snasphot', detail_report_path='')
        logging.info('***** End New Violations Report  *****')
        pass
