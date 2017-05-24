import collections


class Constants(object):
    ROLLUP_SHEET_TITLE = "Rollup"
    ROLLUP_SHEET_COLOR = "BFBFBF"
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
    CLOSED_ELAPSED_ROLLUP_SHEET_TITLE = "ClosedElapsed_Rollup"
    CLOSED_ELAPSED_ROLLUP_SHEET_HEADERS = ["Project", "Run Date", "Average of Days Elapsed",
                                           "Max of Days Elapsed", "Min of Days Elapsed", "Median"]
    CLOSED_ELAPSED_ROLLUP_FREEZE_PANE_CELL = "A2"
    CLOSED_ELAPSED_STATISTICS = ["Average", "Max", "Median"]

    CELL_RANGE = "CELL_RANGE"
    SUB_HEADER = "SUB_HEADER"
    PROJECT_SHEET_PROPERTIES = {
        "EXPRT" : {
            "SHEET_HEADER" : ["Week#", "Run Date", "New", "diff", "In Progress", "diff", "Closed", "diff"],
            "SHEET_FREEZE_PANE_CELL" : "A2",
            "SHEET_COLOR" : "9BBB59"
        },
        "EPR": {
            "SHEET_HEADER": ["Week#", "Run Date", "New", "diff", "In Progress", "diff", "Closed", "diff"],
            "SHEET_FREEZE_PANE_CELL": "A2",
            "SHEET_COLOR": "C0504D"
        },
        "MPORT": {
            "SHEET_HEADER": ["Week#", "Run Date", "New", "diff", "In Progress", "diff", "Closed", "diff"],
            "SHEET_FREEZE_PANE_CELL": "A2",
            "SHEET_COLOR": "8064A2"
        },
        "RCVS": {
            "SHEET_HEADER": ["Week#", "Run Date", "New", "diff", "In Progress", "diff", "Closed", "diff"],
            "SHEET_FREEZE_PANE_CELL": "A2",
            "SHEET_COLOR": "4BACC6"
        },
        "SPOR": {
            "SHEET_HEADER": ["Week#", "Run Date", "New", "diff", "In Progress", "diff", "Closed", "diff"],
            "SHEET_FREEZE_PANE_CELL": "A2",
            "SHEET_COLOR": "F79646"
        },
        "CRQST": {
            "SHEET_HEADER": ["Week#", "Run Date", "New", "diff", "In Progress", "diff", "Closed", "diff"],
            "SHEET_FREEZE_PANE_CELL": "A2",
            "SHEET_COLOR": "4F81BD"
        }

    }

    CHART_PROPERTIES = {
        "LINE_CHART" : {
            "EXPRT": {
                "MARKER_SYMBOL": "triangle",
                "COLOR": "9BBB59"
            },
            "EPR": {
                "MARKER_SYMBOL": "square",
                "COLOR": "C0504D"
            },
            "MPORT": {
                "MARKER_SYMBOL": "x",
                "COLOR": "8064A2"
            },
            "RCVS": {
                "MARKER_SYMBOL": "star",
                "COLOR": "4BACC6"
            },
            "SPOR": {
                "MARKER_SYMBOL": "circle",
                "COLOR": "F79646"
            },
            "CRQST": {
                "MARKER_SYMBOL": "diamond",
                "COLOR": "4F81BD"
            },
            "Max": {
                "MARKER_SYMBOL": "triangle",
                "COLOR": "9BBB59"
            },
            "Average": {
                "MARKER_SYMBOL": "diamond",
                "COLOR": "4F81BD"
            },
            "Median": {
                "MARKER_SYMBOL": "square",
                "COLOR": "C0504D"
            }

        },
        "BAR_CHART": {
            "EXPRT": {
                "COLOR": "9BBB59"
            },
            "EPR": {
                "COLOR": "C0504D"
            },
            "MPORT": {
                "COLOR": "8064A2"
            },
            "RCVS": {
                "COLOR": "4BACC6"
            },
            "SPOR": {
                "COLOR": "F79646"
            },
            "CRQST": {
                "COLOR": "4F81BD"
            }
        }
    }


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
            "pivot_sheet_color": "948A54",
            "charts_sheet_name": "Charts-WeeklyTotals",
            "charts_sheet_position": 0,
            "charts_sheet_color": "DDD9C4",
            "chart_weekly_total_title": "Weekly Total - All Tickets",
            "chart_weekly_growth_title": "Weekly Growth of Tickets"
        }
    METRICS[CLOSED_WEEKLY_TOTALS] = {
            "pivot_sheet_name": "Pivot-Closed-WeeklyTotals",
            "pivot_sheet_position": 1,
            "pivot_sheet_color": "CC0099",
            "charts_sheet_name": "Chart-Closed-WeeklyTotals",
            "charts_sheet_position": 2,
            "charts_sheet_color": "FF99CC",
            "chart_weekly_total_title": "Weekly Totals - Closed Tickets"
        }
    METRICS[CLOSED_WEEKLY_CHANGE] = {
            "pivot_sheet_name": "Pivot-Closed-WeeklyChange",
            "pivot_sheet_position": 2,
            "pivot_sheet_color": "CC0099",
            "charts_sheet_name": "Chart-Closed-WeeklyChange",
            "charts_sheet_position": 4,
            "charts_sheet_color": "FF99CC",
            "chart_weekly_change_title": "Weekly Change of Totals for Closed Tickets"
        }
    METRICS[IN_PROGRESS_WEEKLY_TOTALS] = {
            "pivot_sheet_name": "Pivot-InProgress-WeeklyTotals",
            "pivot_sheet_position": 3,
            "pivot_sheet_color": "339933",
            "charts_sheet_name": "Chart-InProgress-WeeklyTotals",
            "charts_sheet_position": 6,
            "charts_sheet_color": "99CC00",
            "chart_weekly_total_title": "Weekly Totals - In Progress Tickets",
        }
    METRICS[IN_PROGRESS_WEEKLY_CHANGE] = {
            "pivot_sheet_name": "Pivot-InProg-WeeklyChange",
            "pivot_sheet_position": 4,
            "pivot_sheet_color": "339933",
            "charts_sheet_name": "Chart-InProg-WeeklyChange",
            "charts_sheet_position": 8,
            "charts_sheet_color": "99CC00",
            "chart_weekly_change_title": "Weekly Change of Totals for \"In Progress\" Tickets"
        }
    METRICS[NEW_WEEKLY_TOTALS] = {
            "pivot_sheet_name": "Pivot-New-WeeklyTotals",
            "pivot_sheet_position": 5,
            "pivot_sheet_color": "E26B0A",
            "charts_sheet_name": "Chart-New-WeeklyTotals",
            "charts_sheet_position": 10,
            "charts_sheet_color": "FCD5B4",
            "chart_weekly_total_title": "Weekly Total - New Tickets",
        }
    METRICS[NEW_WEEKLY_CHANGE] = {
            "pivot_sheet_name": "Pivot-New-WeeklyChange",
            "pivot_sheet_position": 6,
            "pivot_sheet_color": "E26B0A",
            "charts_sheet_name": "Chart-New-WeeklyChange",
            "charts_sheet_position": 12,
            "charts_sheet_color": "FCD5B4",
            "chart_weekly_change_title": "Weekly Change of Totals for \"New\" Tickets"
        }
    METRICS[CLOSED_ELAPSED] = {
            "pivot_sheet_name": "Pivot-ClosedElapsed",
            "pivot_sheet_position": 7,
            "pivot_sheet_color": "FFFFFF",
            "charts_sheet_name": "Chart-ClosedElapsed",
            "charts_sheet_position": 14,
            "charts_sheet_color": "F2F2F2",
            "chart_current_week_title": "Analysis of Days Elapsed, Per Project, for Current Week"
        }

if __name__ == '__main__':
    print Constants.CLOSED_ELAPSED
    print Constants.CHART_PROPERTIES['LINE_CHART']
