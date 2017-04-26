import collections
class Constants(object):
    ROLLUP_SHEET_TITLE = "Rollup"
    ROLLUP_HEADER_ROWS = "2"
    ROLLUP_HEADER_WKLY_DIFF_STATUSES = ["Current Week", "Last Week", "Difference"]
    ROLLUP_HEADER_WKLY_GROWTH_STATUSES = ["Current Week", "Last Week", "Growth"]
    ROLLUP_SHEET_HEADERS = ["Project", "Run Date", "New", "In Progress", "Closed", "New & In Progress", "Total"]
    ROLLUP_SHEET_HEADER_PROPERTIES = {
        "Project": {
            "CELL_RANGE": "A1:A2"
        },
        "Run Date": {
            "CELL_RANGE": "B1:B2"
        },
        "New": {
            "CELL_RANGE": "C1:E1", "SUB_HEADER": ROLLUP_HEADER_WKLY_DIFF_STATUSES
        },
        "In Progress": {
            "CELL_RANGE": "F1:H1", "SUB_HEADER": ROLLUP_HEADER_WKLY_DIFF_STATUSES
        },
        "Closed": {
            "CELL_RANGE": "I1:K1", "SUB_HEADER": ROLLUP_HEADER_WKLY_DIFF_STATUSES
        },
        "New & In Progress": {
            "CELL_RANGE": "L1:N1", "SUB_HEADER": ROLLUP_HEADER_WKLY_DIFF_STATUSES
        },
        "Total": {
            "CELL_RANGE": "O1:Q1", "SUB_HEADER": ROLLUP_HEADER_WKLY_GROWTH_STATUSES
        }
    }
    ROLLUP_FREEZE_PANE_CELL = "A3"
    CELL_RANGE = "CELL_RANGE"
    SUB_HEADER = "SUB_HEADER"
    PROJECT_SHEET_HEADER = ["Week#", "Run Date", "New", "diff", "In Progress", "diff", "Closed", "diff"]
    PROJECT_SHEET_FREEZE_PANE_CELL = 'A2'
    STATUS_NEW = "New"
    STATUS_INPROGRESS = "InProgress"
    STATUS_CLOSED = "closed"
    STATUS_NEW_COLUMN = 2
    STATUS_INPROGRESS_COLUMN = 4
    STATUS_CLOSED_COLUMN = 6
    # WEEKLY_TOTAL_ALL_TICKETS_CHART = "Charts-WeeklyTotals"
    ALL_TICKETS_WEEKLY_TOTALS = "Weekly-Totals-AllTickets"
    CLOSED_WEEKLY_TOTALS = "Closed-WeeklyTotals"
    CLOSED_WEEKLY_CHANGE = "Closed-WeeklyChange"
    IN_PROGRESS_WEEKLY_TOTALS = "InProgress-WeeklyTotals"
    IN_PROGRESS_WEEKLY_CHANGE = "InProg-WeeklyChange"
    NEW_WEEKLY_TOTALS = "New-WeeklyTotals"
    NEW_WEEKLY_CHANGE = "New-WeeklyChange"
    CLOSED_ELAPSED = "ClosedElapsed"
    METRICS = collections.OrderedDict()
    METRICS[ALL_TICKETS_WEEKLY_TOTALS] = {
            "pivot_sheet_name": "Pivot-Weekly-Totals-AllTickets",
            "pivot_sheet_position": 0,
            "charts_sheet_name": "Charts-WeeklyTotals",
            "charts_sheet_position": 0
        }
    METRICS[CLOSED_WEEKLY_TOTALS] = {
            "pivot_sheet_name": "Pivot-Closed-WeeklyTotals",
            "pivot_sheet_position": 1,
            "charts_sheet_name": "Chart-Closed-WeeklyTotals",
            "charts_sheet_position": 1
        }
    METRICS[CLOSED_WEEKLY_CHANGE] = {
            "pivot_sheet_name": "Pivot-Closed-WeeklyChange",
            "pivot_sheet_position": 2,
            "charts_sheet_name": "Chart-Closed-WeeklyChange",
            "charts_sheet_position": 2
        }
    METRICS[IN_PROGRESS_WEEKLY_TOTALS] = {
            "pivot_sheet_name": "Pivot-InProgress-WeeklyTotals",
            "pivot_sheet_position": 3,
            "charts_sheet_name": "Chart-InProgress-WeeklyTotals",
            "charts_sheet_position": 3
        }
    METRICS[IN_PROGRESS_WEEKLY_CHANGE] = {
            "pivot_sheet_name": "Pivot-InProg-WeeklyChange",
            "pivot_sheet_position": 4,
            "charts_sheet_name": "Chart-InProg-WeeklyChange",
            "charts_sheet_position": 4
        }
    METRICS[NEW_WEEKLY_TOTALS] = {
            "pivot_sheet_name": "Pivot-New-WeeklyTotals",
            "pivot_sheet_position": 5,
            "charts_sheet_name": "Chart-New-WeeklyTotals",
            "charts_sheet_position": 5
        }
    METRICS[NEW_WEEKLY_CHANGE] = {
            "pivot_sheet_name": "Pivot-New-WeeklyChange",
            "pivot_sheet_position": 6,
            "charts_sheet_name": "Chart-New-WeeklyChange",
            "charts_sheet_position": 6
        }
    METRICS[CLOSED_ELAPSED] = {
            "pivot_sheet_name": "Pivot-ClosedElapsed",
            "pivot_sheet_position": 7,
            "charts_sheet_name": "Chart-ClosedElapsed",
            "charts_sheet_position": 7
        }

if __name__ == '__main__':
    print Constants.CLOSED_ELAPSED
