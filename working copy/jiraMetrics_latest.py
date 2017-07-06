import os
import re
import sys
import time
import datetime
from Constants import Constants
from ConfigParser import SafeConfigParser
from lib.JiraAPIHandler import JiraAPIHandler
from lib.ExcelWorkBookManager import ExcelWorkBookManager
from lib.ProjectProperties import ProjectProperties

if __name__ == "__main__":
    CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__))
    config_file = os.path.join(CURRENT_DIRECTORY, 'config', 'jiraMetrics.ini')
    if os.path.exists(config_file):
        config = SafeConfigParser()
        config.optionxform = str
        config.read(config_file)

    else:
        print config_file + " not found"
        time.sleep(5)
        sys.exit(0)

    output_dir = os.path.join(CURRENT_DIRECTORY, 'output')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    today = datetime.date.today()
    days_to_subtract = config.get('BUG_TRACKER', 'day_difference')
    days_to_subtract = int(days_to_subtract) if days_to_subtract else 0
    program_run_date = today - datetime.timedelta(days=days_to_subtract)
    run_date_yyyy_mm_dd = program_run_date.strftime("%Y-%m-%d")
    out_put_file_name = os.path.join(output_dir, config.get('OUTPUT', 'output_file_name'))
    out_put_file_name = out_put_file_name.replace('yyyy-mm-dd', run_date_yyyy_mm_dd)
    if os.path.exists(out_put_file_name):
        try:
            os.rename(out_put_file_name, out_put_file_name)
        except OSError as e:
            print out_put_file_name + ' already in use. Please close it'
            time.sleep(5)
            sys.exit(0)
    project_properties = ProjectProperties(config)
    project_properties.initialize_project_properties()
    with open(config_file, 'wb') as configfile:
        config.write(configfile)
    workbook_manager = ExcelWorkBookManager(config)
    if not os.path.exists(out_put_file_name):
        workbook_manager.create_empty_workbook(out_put_file_name)
        # Extract data from manually created workbook, if file exists
        old_workbook_file_name = None
        for filename in os.listdir(CURRENT_DIRECTORY):
            match = re.match('(From.*?.xlsm)', filename, re.I)
            if match:
                old_workbook_file_name = match.group(0)
        if old_workbook_file_name:
            old_workbook_file_name = os.path.join(CURRENT_DIRECTORY, old_workbook_file_name)
            if os.path.exists(old_workbook_file_name):
                workbook_manager.extract_data_from_old_file_and_insert_into_new_file(
                    old_workbook_file_name, out_put_file_name
                )
            else:
                print "{} File Not exists".format(old_workbook_file_name)
        else:
            print "Manually created workbook not exists in current directory"
    jira_api = JiraAPIHandler(config)
    workbook_manager.populate_latest_metrics_from_jira_for_date(program_run_date, jira_api, out_put_file_name)
    metrics = Constants.METRICS
    for metric_name, metric_properties in metrics.iteritems():
        workbook_manager.create_or_update_pivot_table_for(metric_name, out_put_file_name, program_run_date)

    workbook_manager.update_charts_for(metrics, out_put_file_name)

    print "Task Completed"
