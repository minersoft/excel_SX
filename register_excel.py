import miner_globals
# define targets
miner_globals.addExtensionToTargetMapping(".xlsx", "excel")

# always generate excel via proxy
if True or miner_globals.runsUnderPypy:
    miner_globals.addTargetToClassMapping("excel", None, "excel_target_proxy.oExcelProxy", "Creates Excel spreadsheet")
else:
    miner_globals.addTargetToClassMapping("excel", None, "excel_target.oExcel", "Creates Excel spreadsheet")
