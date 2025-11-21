import sys
sys.path.insert(0, '.')
try:
    import src.data_cleaner as dc
    import src.stat_export as se
    print('IMPORT_OK')
except Exception as e:
    print('IMPORT_FAIL', e)
