from openpyxl.chart.trendline import Trendline
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import BarChart, LineChart, Series, Reference


class ChartManager(object):

    def __init__(self, data_sheet, charts_sheet):
        self.data_sheet = data_sheet
        self.charts_sheet = charts_sheet

    def draw_barchart(self, chart_properties):
        print " creating Bar Chart"
        chart1 = BarChart()
        chart1.height = 12
        chart1.width = 30
        chart1.style = 10
        chart1.title = "Weekly Total - All Tickets"
        chart1.y_axis.title = 'Total'
        chart1.x_axis.title = 'Run Date'

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
        chart1.add_data(data, titles_from_data=True)
        chart1.set_categories(cats)
        chart1.shape = 4
        if chart_properties['trendline']:
            chart1.series[0].trendline = Trendline()
            chart1.series[0].trendline.trendlineType = 'linear'
        if chart_properties['data_labels']:
            chart1.dataLabels = DataLabelList()
            chart1.dataLabels.showVal = True
        self.charts_sheet.add_chart(chart1, chart_properties['cell'])

    def draw_linechart(self, chart_properties):
        print "Creating line chart"
        c1 = LineChart()
        c1.height = 12
        c1.width = 30
        c1.title = "Weekly Growth"
        c1.style = 12
        c1.y_axis.title = 'Growth'
        c1.x_axis.title = 'Run Date'
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

        c1.add_data(data, titles_from_data=True)
        c1.set_categories(cats)
        # Style the lines
        s1 = c1.series[0]
        s1.marker.symbol = "circle"
        s1.marker.graphicalProperties.solidFill = "360AD2"  # Marker filling
        s1.marker.graphicalProperties.line.solidFill = "360AD2"
        s1.graphicalProperties.line.solidFill = "360AD2"
        s1.graphicalProperties.line.width = 28568  # width in EMUs
        if chart_properties['trendline']:
            s1.trendline = Trendline()
            s1.trendline.trendlineType = 'linear'
        # s1.smooth = True # Make the line smooth
        self.charts_sheet.add_chart(c1, chart_properties['cell'])
