from Constants import Constants
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import BarChart, LineChart, Series, Reference


class ChartManager(object):

    def __init__(self, data_sheet, charts_sheet):
        self.data_sheet = data_sheet
        self.charts_sheet = charts_sheet

    def draw_barchart(self, chart_properties):
        print " creating Bar Chart"
        chart = BarChart()
        chart.height = 12
        chart.width = 30
        chart.style = 10
        chart.title = chart_properties['title']
        #chart.y_axis.title = 'Total'
        #chart.x_axis.title = 'Run Date'

        data = Reference(
            self.data_sheet, min_col=chart_properties['data_min_column'],
            min_row=chart_properties['data_min_row'],
            max_row=chart_properties['data_max_row'],
            max_col=chart_properties['data_max_column'])
        cats = Reference(
            self.data_sheet,
            min_col=chart_properties['cats_min_column'],
            min_row=chart_properties['cats_min_row'],
            max_row=chart_properties['cats_max_row'])
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.shape = 4
        projects = chart_properties['projects']
        series = chart.series
        for index, current_series in enumerate(series):
            if projects:
                project = projects[index]
                project_properties = self.get_barchart_properties_for(project)
                project_color = project_properties["COLOR"]
                current_series.graphicalProperties.solidFill = project_color
            else:
                current_series.graphicalProperties.solidFill = "1381BD"

            if chart_properties['trendline']:
                current_series.trendline = Trendline()
                current_series.trendline.trendlineType = 'linear'
        if chart_properties['data_labels']:
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showVal = True
        self.charts_sheet.add_chart(chart, chart_properties['cell'])

    def draw_linechart(self, chart_properties):
        print "Creating line chart"
        chart = LineChart()
        chart.height = 12
        chart.width = 30
        chart.title = chart_properties['title']
        chart.style = 12
        chart.x_axis.tickLblPos = "low"
        #chart.y_axis.title = 'Growth'
        #chart.x_axis.title = 'Run Date'
        data = Reference(
            self.data_sheet, min_col=chart_properties['data_min_column'],
            min_row=chart_properties['data_min_row'],
            max_row=chart_properties['data_max_row'],
            max_col=chart_properties['data_max_column'])
        cats = Reference(
            self.data_sheet,
            min_col=chart_properties['cats_min_column'],
            min_row=chart_properties['cats_min_row'],
            max_row=chart_properties['cats_max_row'])

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        projects = chart_properties['projects']
        series = chart.series
        for index, current_series in enumerate(series):
            if projects:
                project = projects[index]
                project_properties = self.get_linechart_properties(project)
                current_series.marker.symbol = project_properties["MARKER_SYMBOL"]
                current_series.marker.size = 7
                if project_properties["MARKER_SYMBOL"] in ("triangle", "diamond", "circle"):
                    current_series.marker.graphicalProperties.solidFill = project_properties["COLOR"] # Marker filling
                current_series.marker.graphicalProperties.line.solidFill = project_properties["COLOR"]
                current_series.graphicalProperties.line.solidFill = project_properties["COLOR"]
            else:
                current_series.marker.symbol = "diamond"
                current_series.marker.graphicalProperties.solidFill = "1381BD"  # Marker filling
                current_series.marker.graphicalProperties.line.solidFill = "1381BD"
                current_series.graphicalProperties.line.solidFill = "1381BD"
            current_series.graphicalProperties.line.width = 28568
            if chart_properties['trendline']:
                current_series.trendline = Trendline()
                current_series.trendline.trendlineType = 'linear'

        self.charts_sheet.add_chart(chart, chart_properties['cell'])

    @staticmethod
    def get_linechart_properties(project):
        linechart_properties = Constants.CHART_PROPERTIES['LINE_CHART']
        return linechart_properties[project]

    @staticmethod
    def get_barchart_properties_for(project):
        barchart_properties = Constants.CHART_PROPERTIES['BAR_CHART']
        return barchart_properties[project]
