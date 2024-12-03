# PureLiving Fapiao Updater

A tool for weekly fapiao update, synchronize contents and highlights

## Notes

Format specific, suitable for fapiao excels in Fall 2024, cannot ensure future workflows

The procedure can be concluded as copy notes from original fapiao excel to new fapiao excel provided by finance department

Make sure the table always starts at A1, with title row included, elements that are not belonged to the dataset should be removed.

No duplicated title names!!!

## Workflow

The desired columns to be copied and pivot during copy process can be configured in  "FapiaoUpdaterConfig.cfg"

The 2 programs will follow the workflow of copying columns & highlight from src to dest

ContentSync.py aims for this procedure, which will generate a file named "Content Updated.xlsx" file with desired columns updated

If the columns specified in FapiaoUpdaterConfig.cfg doesn't exist in src, the program will raise error and quit

If the columns specified in FapiaoUpdaterConfig.cfg doesn't exist in dest, the program will create corresponding columns and continue copying
```
python ContentSync.py
```

If also want to sync highlights, can use HighlightSync.py
```
python HighlightSync.py
```

