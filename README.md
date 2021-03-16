# Buto-Plugin-ExcelPhp_excel_v18
Using PHPExcel-1.8.

## Widget demo
Widget to show how to open file, select a sheet, and get a range of data.
```
type: widget
data:
  plugin: excel/php_excel_v18
  method: demo
  data:
    file: /plugin/excel/php_excel_v18/demo.ods
    sheet: demo
    range: 'A1:C5'
```
Optional first row can be used as keys.
```
    first_row_as_key: true
```
Optional key_id if first_row_as_key is set.
```
    key_id: (An existing key)
```

## Method columnToFile
Method to create file from column data.

## Method download_excel
```
$data = array(array('id' => '1234', 'name' => 'John'));
wfPlugin::includeonce('excel/php_excel_v18');
$excel = new PluginExcelPhp_excel_v18();
$excel->download_excel($data, 'people.xlsx');
```
## Method save_excel
```
$data = array(array('id' => '1234', 'name' => 'John'));
wfPlugin::includeonce('excel/php_excel_v18');
$excel = new PluginExcelPhp_excel_v18();
$excel->save_excel($data, '/theme/my/theme/people.xlsx');
```
