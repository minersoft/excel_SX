import json
import tempfile
import os.path
import subprocess
import SX_environ
 
class oExcelProxy:
    def __init__(self, fileName, variableNames, **moreParams):
        self.f = tempfile.NamedTemporaryFile(prefix="miner_excel_proxy_data", suffix=".json", delete=False)
        self.jsonFileName = self.f.name
        #print "Excel data going to", self.jsonFileName 
        self.data = []
        self.inputData = {
            'fileName': fileName,
            'variableNames': list(variableNames),
            'moreParams': moreParams,
            'data': self.data,
            'SX_JAR': SX_environ.get_SX_JAR(),
            'JAVA_HOME': SX_environ.get_JAVA_HOME(),
        }
        
    def save(self, record):
        dataRecord = []
        for v in record:
            if isinstance(v, int) or isinstance(v, long) or isinstance(v, float) or isinstance(v, str):
                dataRecord.append(v)
            else:
                dataRecord.append(str(v))
        self.data.append(dataRecord)
    def close(self):
        json.dump(self.inputData, self.f)
        self.f.close()
        subprocess.call(["python", os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel_target.py"),
                         self.jsonFileName])
        os.unlink(self.jsonFileName)