import traceback
import warnings
import os
import sys
import collections
import SX_environ

with warnings.catch_warnings():
    #jpype module uses deprecated sets module
    warnings.simplefilter("ignore")
    import jpype

gIsJavaInitialized = False

def initJava(JAVA_HOME=None, SX_JAR=None):
    global gIsJavaInitialized
    if gIsJavaInitialized:
        return
    if not JAVA_HOME:
        JAVA_HOME = SX_environ.get_JAVA_HOME()
    if not SX_JAR:
        SX_JAR = SX_environ.get_SX_JAR() 
    os.environ['JAVA_HOME'] = JAVA_HOME
    
    options = [
        '-Djava.class.path=%s' % SX_JAR
    ]
    jvmPath =  jpype.getDefaultJVMPath()
    if not jvmPath:
        jvmPath = os.path.join(JAVA_HOME, "bin", "client", "jvm.dll")
        if not os.path.isfile(jvmPath):
            print >>sys.stderr, "JVM dll doesn't exist at %s" % jvmPath
            return False
    try:
        print "Before Start"
        jpype.startJVM(jvmPath, *options)
        print "After Start"
    except:
        traceback.print_exc()
        return False
    gIsJavaInitialized = True
    return True

def time2excel(timestamp):
    import datetime
    dt = datetime.datetime.utcfromtimestamp(timestamp)
    dt1900 = datetime.datetime(1900,1,1)
    delta = dt - dt1900

    # Excel handles time in number of days since 1, January 1900
    # Where 1, Jan 1900 itself is 1
    # Additional +1 is fix to bug in Excel that treats 1900 as a leap year
    value = delta.days + 2 + delta.seconds/24./3600.
    return value

def identity(val):
    return val

kKB = 1024
kMB = 1024*kKB
kGB = 1024*kMB
kDay = 24*3600

class oExcel:
    
    class ChartInfo(object):
        def __init__(self):
            self.topRow = -1 
            self.chartType = None
            self.chartTitle = None
            self.chartWidth = 12
            self.chartHeight = 24
            self.chartAlign = "top"
            self.firstColumn = 1
            self.numDataColumns = 999
        def __str__(self):
            return "width=%d height=%d" % (self.chartWidth, self.chartHeight)

    formatMapping = { ",": ("#,##0", identity),
                      "%": ("0.0%", identity),
                      "K": ('#,##0.0,"K"', identity),
                      "M": ('#,##0.00,,"M"', identity),
                      "G": ('#,##0.00,,,"G"', identity),
                      "KB": ('#,##0"KB"', lambda x: float(x)/kKB),
                      "MB": ('#,##0.0"MB"', lambda x: float(x)/kMB),
                      "GB": ('#,##0.00"GB"', lambda x: float(x)/kGB),
                      "n": ('[<200000]#,##0;[<500000000]0.00,,"M";0.00,,,"G"', identity),
                      "N": ('[<500000000]0.00,,"M";[<500000000000]0.00,,,"G";0.00,,,,"T"', identity),
                      "T": ('yyyy-mm-dd hh:mm:ss', time2excel),
                      "Tm": ('yyyy-mm-dd hh:mm:ss.000', time2excel),
                      "t": ('[m]:ss', lambda x: float(x)/kDay),
                      "mt": ('[m]:ss.000', lambda x: float(x)/1000/kDay),
                    }
    def __init__(self, fileName, variableNames, **moreParams):
        initJava()
        self.myFileName = fileName
        self.mySheetName = moreParams.get("sheetName", None)
        self.initWorkbook()
        self.myVars = variableNames
        self.lastDataColumn = len(self.myVars) - 1
        self.myNumColumns = len(variableNames)
        self.chartInfoDict = collections.defaultdict(oExcel.ChartInfo)
        self.formattedColumns = {}
        self.columnTitles = {}
        self.conversionFunctions = [identity] * len(self.myVars)
        for (param, value) in moreParams.iteritems():
            if param.endswith('_format'):
                value = oExcel.formatMapping.get(value, value)
                self.formattedColumns[param[:-7]] = value[0]
            elif param.endswith('_title'):
                self.columnTitles[param[:-6]] = value
            elif param.startswith("chartType"):
                self.chartInfoDict[param[9:]].chartType = value
            elif param.startswith("chartTitle"):
                self.chartInfoDict[param[10:]].chartTitle = value
            elif param.startswith("chartWidth"):
                self.chartInfoDict[param[10:]].chartWidth = value
            elif param.startswith("chartHeight"):
                self.chartInfoDict[param[11:]].chartHeight = value
            elif param.startswith("chartAlign"):
                self.chartInfoDict[param[10:]].chartAlign = value
            elif param.startswith("firstColumn"):
                if isinstance(value, str) or isinstance(value, unicode):
                    try:
                        firstColumn = self.myVars.index(value)
                    except ValueError:
                        print "Variable %s is unknown" % value
                        continue
                else:
                    firstColumn = value
                #print "firstColumn", firstColumn
                self.chartInfoDict[param[11:]].firstColumn = firstColumn
            elif param.startswith("numColumns"):
                self.chartInfoDict[param[10:]].numDataColumns = value
            elif param in self.myVars:
                # consider this as format
                value = oExcel.formatMapping.get(value, value)
                self.formattedColumns[param] = value
        for param, value in self.formattedColumns.iteritems():
            self.setConversion(param, value)
        # calculate top tows for all charts that appear at the top
        topRow = 0
        for chartId, chartInfo in sorted(self.chartInfoDict.iteritems()):
            if not chartInfo.chartType:
                print "No chart type appears for %s chart" % (chartId if chartId else "default")
                continue
            #print "Chart Info", chartInfo
            if chartInfo.chartAlign == "top":
                chartInfo.topRow = topRow
                topRow += chartInfo.chartHeight + 1
            chartInfo.numDataColumns = min(self.myNumColumns-chartInfo.firstColumn, chartInfo.numDataColumns)
        
        if topRow:
            self.myTitleRow = topRow + 1
        else:
            self.myTitleRow = 0
        for i in range(len(variableNames)):
            title = self.columnTitles.get(variableNames[i], variableNames[i])
            self.myWorkbook.setText(self.myTitleRow, i, title)
        self.myNextRow = self.myTitleRow + 1
        
            
    def save(self, record):
        for i in range(len(record)):
            val = self.conversionFunctions[i](record[i])
            if isinstance(val, float) or isinstance(val, int):
                self.myWorkbook.setNumber(self.myNextRow, i, float(val))
            else:
                self.myWorkbook.setText(self.myNextRow, i, str(val))
            
        self.myNextRow += 1
    def close(self):
        for i in range(len(self.myVars)):
            self.myWorkbook.setColWidthAutoSize(i, True)
        topRowForBottomCharts = self.myNextRow+1
        for chartId,chartInfo in sorted(self.chartInfoDict.iteritems()):
            if not chartInfo.chartType:
                continue
            if chartInfo.chartAlign == "bottom":
                chartInfo.topRow = topRowForBottomCharts
                topRowForBottomCharts += chartInfo.chartHeight+1
            self.createChart(chartInfo)

        for (columnName, format) in self.formattedColumns.iteritems():
            self.applyFormat(columnName, format)
        self.applyTitleFormat()
        try:
            self.myWorkbook.writeXLSX(self.myFileName)
        except:
            s = "Failed to write to %s, may be file is opened?" % self.myFileName
            ioerr = IOError( s )
            raise ioerr
    def createChart(self, chartInfo):
        left = 0.5
        right = left + chartInfo.chartWidth
        top = chartInfo.topRow
        bottom = top + chartInfo.chartHeight
        chart = self.myWorkbook.addChart(float(left),float(top),float(right),float(bottom))
        if chartInfo.chartTitle:
            chart.setTitle(chartInfo.chartTitle)
        if chartInfo.chartType in ["column", "bar"]:
            self.createColumnChart(chart, chartInfo)
        elif chartInfo.chartType in ["pie", "doughnut"]:
            self.createPieChart(chart, chartInfo)
        elif chartInfo.chartType in ["line", "stackedLine", "area", "stackedArea"]:
            self.createLineChart(chart, chartInfo)
        
    def setLinkRange(self, chart, chartInfo, limitColumns=1000):
        #chart.setLinkRange("%s!$%s$%d:$%s$%d" %
        #    (self.mySheetName, chr(ord('a')+self.firstColumn), self.myTitleRow+1, chr(ord('a')+self.firstColumn + self.numDataColumns-1),self.myNextRow),False)
        numColumns = min(chartInfo.numDataColumns, limitColumns)
        for series in range(numColumns):
            chart.addSeries()
            seriesColName = chr(ord('a')+chartInfo.firstColumn+series)
            chart.setSeriesName(series,"%s!$%s$%d" % (self.mySheetName, seriesColName, self.myTitleRow+1))
            chart.setSeriesYValueFormula(series, "%s!$%s$%d:$%s$%d" %
                (self.mySheetName, seriesColName, self.myTitleRow+2, seriesColName, self.myNextRow) )
    def getNumDataRows(self):
        return self.myNextRow-self.myTitleRow-1
    def setCategoryFormula(self, chart, chartInfo):
        formula = "%s!$a$%d:$a$%d" % (self.mySheetName,self.myTitleRow+2,self.myNextRow)
        #print "categoryFormula = " + formula
        chart.setCategoryFormula(formula)
        if self.getNumDataRows() > 20:
            interval = self.getNumDataRows()/20
            ChartShape = jpype.JClass('com.smartxls.ChartShape')
            ChartFormat = jpype.JClass('com.smartxls.ChartFormat')
            for series in range(chartInfo.numDataColumns):
                format = chart.getSeriesFormat(series)
                format.setMarkerStyle(ChartFormat.MarkerNone)
                chart.setSeriesFormat(series, format)
            chart.setAxisScaleType(ChartShape.XAxis, 0, ChartShape.TimeScale)
            chart.setTimeScaleMajorUnit(ChartShape.XAxis,0,5,5)
    def createColumnChart(self, chart, chartInfo):
        ChartShape = jpype.JClass('com.smartxls.ChartShape')
        if chartInfo.chartType == "column":
            chart.setChartType(ChartShape.Column)
        elif chartInfo.chartType == "bar":
            chart.setChartType(ChartShape.Bar)
        self.setLinkRange(chart, chartInfo)
        #set axis title
        chart.setAxisTitle(ChartShape.XAxis, 0, self.myVars[0]);
        self.setCategoryFormula(chart,chartInfo)
    def createPieChart(self, chart, chartInfo):
        ChartShape = jpype.JClass('com.smartxls.ChartShape')
        ChartFormat = jpype.JClass('com.smartxls.ChartFormat')
        format = chart.getChartFormat()
        limitColumns = 1000
        if chartInfo.chartType == "pie":
            limitColumns = 1
            chart.setChartType(ChartShape.Pie)
            self.setLinkRange(chart, chartInfo, limitColumns)
            plotFormat = chart.getPlotFormat()
            plotFormat.setDataLabelPosition(ChartFormat.DataLabelPositionAuto)
            plotFormat.setDataLabelType(ChartFormat.DataLabelPercent)
            chart.setPlotFormat(plotFormat)
            
        elif chartInfo.chartType == "doughnut":
            chart.setChartType(ChartShape.Doughnut)
            format.setDataLabelType(ChartFormat.DataLabelPercent)
            self.setLinkRange(chart, chartInfo, limitColumns)
            plotFormat = chart.getPlotFormat()
            plotFormat.setDataLabelPosition(ChartFormat.DataLabelPositionAuto)
            for i in range(min(chartInfo.numDataColumns, limitColumns)):
                dataLabelFormat = chart.getDataLabelFormat(1)
                dataLabelFormat.setDataLabelType(ChartFormat.DataLabelPercent)
                dataLabelFormat.setDataLabelPosition(ChartFormat.DataLabelPositionDefault)
                chart.setDataLabelFormat(i, dataLabelFormat)
        chart.setLegendPosition(ChartShape.LegendPlacementLeft)
        chart.setChartFormat(format)

        
        self.setCategoryFormula(chart, chartInfo)
        chart.setVaryColors(True)

    def createLineChart(self, chart, chartInfo):
        ChartShape = jpype.JClass('com.smartxls.ChartShape')
        ChartFormat = jpype.JClass('com.smartxls.ChartFormat')
        format = chart.getChartFormat()
        isSingleLine = False
        if chartInfo.chartType == "line":
            if chartInfo.numDataColumns == 1:
                isSingleLine = True
                chart.setChartType(ChartShape.Area)
            else:
                chart.setChartType(ChartShape.Line)
        elif chartInfo.chartType == "stackedLine":
            chart.setChartType(ChartShape.Line)
            chart.setPlotStacked(True)
        elif chartInfo.chartType == "area":
            chart.setChartType(ChartShape.Area)
        elif chartInfo.chartType == "stackedArea":
            chart.setChartType(ChartShape.Area)
            chart.setPlotStacked(True)
        chart.setChartFormat(format)
        self.setLinkRange(chart, chartInfo)
        if isSingleLine:
            seriesFormat = chart.getSeriesFormat(0)
            seriesFormat.setLineAuto()
            seriesFormat.setFillAuto(False)
            
            seriesFormat.setFillVisible(False)
            seriesFormat.setLineStyle(ChartFormat.LineSolid)
            seriesFormat.setLineWeight(ChartFormat.Medium)
            seriesFormat.setLineColor(0x004E79BE)
            chart.setSeriesFormat(0, seriesFormat)
        elif chartInfo.chartType in ["stackedLine", "stackedArea"]:
            for i in range(chartInfo.numDataColumns):
                chart.setPlotGroupStack(i, True)
        self.setCategoryFormula(chart, chartInfo)
        chart.setVaryColors(True)
        chart.setAxisTitle(ChartShape.XAxis, 0, self.myVars[0])

    def createScatterChart(self, chart, chartInfo):
        ChartShape = jpype.JClass('com.smartxls.ChartShape')
        ChartFormat = jpype.JClass('com.smartxls.ChartFormat')
        chart.setChartType(ChartShape.Scatter)
        format = chart.getChartFormat()
        format.setLineAuto()
        formula = "%s!$a$%d:$a$%d" % (self.mySheetName,self.myTitleRow+2,self.myNextRow)
        #print "categoryFormula = " + formula
        self.setLinkRange(chart, chartInfo, 1)
        chart.setSeriesXValueFormula(0, formula)
        chart.setAxisTitle(ChartShape.XAxis, 0, self.myVars[0])
    
        
    def setConversion(self, columnName, format):
        try:
            col = self.myVars.index(columnName)
        except:
            return
        if isinstance(format, tuple):
            self.conversionFunctions[col] = format[1]
    def applyFormat(self, columnName, format):
        try:
            col = self.myVars.index(columnName)
        except:
            return
        if isinstance(format, tuple):
            format = format[0]
        style = self.myWorkbook.getRangeStyle(self.myTitleRow+1, col, self.myNextRow-1, col)
        style.setCustomFormat(format)
        self.myWorkbook.setRangeStyle(style, self.myTitleRow+1, col, self.myNextRow-1, col)
    def applyTitleFormat(self):
        style = self.myWorkbook.getRangeStyle(self.myTitleRow, 0, self.myTitleRow, self.lastDataColumn)
        style.setFontBold(True)
        RangeStyle = jpype.JClass('com.smartxls.RangeStyle')
        #style.setHorizontalAlignment(RangeStyle.HorizontalAlignmentCenter)
        style.setBottomBorder(RangeStyle.BorderMedium)
        self.myWorkbook.setRangeStyle(style, self.myTitleRow, 0, self.myTitleRow, self.lastDataColumn)
    def setNewSheet(self):
        self.mySheetId = -1
        for i in range(self.myWorkbook.getNumSheets()):
            if self.mySheetName == self.myWorkbook.getSheetName(i):
                if self.myWorkbook.getNumSheets() == 1:
                    self.myWorkbook.insertSheets(0,1)
                    self.myWorkbook.setNumSheets(1)
                else:
                    self.myWorkbook.deleteSheets(i, 1)
                    self.myWorkbook.insertSheets(i, 1)
                self.myWorkbook.setSheetName(i, self.mySheetName)
                self.mySheetId = i
                break
        if self.mySheetId == -1:
            i = self.myWorkbook.getNumSheets()
            self.myWorkbook.setNumSheets(i+1)
            self.myWorkbook.setSheetName(i, self.mySheetName)
            self.mySheetId = i
        self.myWorkbook.setSheet(self.mySheetId)
    def initWorkbook(self):
        WorkBook = jpype.JClass('com.smartxls.WorkBook')
        self.myWorkbook = WorkBook()
        if self.mySheetName:
            # try to load workbook from file
            
            try:
                self.myWorkbook.readXLSX(self.myFileName)
                self.setNewSheet()
            except:
                #print "Failed to load new file %s " % self.myFileName
                #print "Unexpected error:", sys.exc_info()[0]
                #raise
                self.myWorkbook.setSheetName(0, self.mySheetName)
        else:
            self.mySheetName = "Sheet1"

def outputExcelFromJson(jsonFileName):
    import json
    f = open(jsonFileName, "rb")
    inputData = json.load(f)
    f.close()
    if not initJava(JAVA_HOME=inputData["JAVA_HOME"], SX_JAR=inputData["SX_JAR"]):
        return False
    excel = oExcel(inputData['fileName'], inputData['variableNames'], **inputData['moreParams'])
    for record in inputData['data']:
        excel.save(tuple(record))
    excel.close()
    return True

if __name__=="__main__":
    outputExcelFromJson(sys.argv[1])
    