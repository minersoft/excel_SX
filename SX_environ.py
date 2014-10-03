def get_JAVA_HOME():
    import miner_globals
    from m.common import MiningError
    #os.environ['JAVA_HOME'] = "/usr/lib/jvm/jre"
    #os.environ['JAVA_HOME'] = r"C:\Program Files (x86)\Java\jre7"
    JAVA_HOME = miner_globals.getScriptParameter("EXTERNAL_JAVA_HOME_PATH", None)
    if not JAVA_HOME:
        raise MiningError("JAVA_HOME path is not defined (parameter EXTERNAL_JAVA_HOME_PATH)")
    return JAVA_HOME

def get_SX_JAR():
    import miner_globals
    from m.common import MiningError
    #SX_JAR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SX.jar')
    SX_JAR = miner_globals.getScriptParameter("EXTERNAL_SX_JAR_PATH", None)
    if not SX_JAR:
        raise MiningError("SX.jar path is not defined (parameter EXTERNAL_SX_JAR_PATH)")
    return SX_JAR
