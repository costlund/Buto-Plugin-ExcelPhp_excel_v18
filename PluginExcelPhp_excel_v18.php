<?php
/**
 * PHPExcel 1.8
 */
class PluginExcelPhp_excel_v18{
  /**
   * Including PHPExcel 1.8 library.
   */
  function __construct($buto = false) {
    if($buto){
      include_once __DIR__.'/PHPExcel-1.8/Classes/PHPExcel.php';
    }
  }
  /**
   * Widget to show how to open file, select a sheet, and get a range of data.
   */
  public function widget_demo($data){
    wfPlugin::includeonce('wf/array');
    $data = new PluginWfArray($data);
    wfHelp::yml_dump($data);
    $xls_data = $this->getRange($data->get('data'));
    wfHelp::yml_dump($xls_data);
  }
  /**
   <p>Array with data.</p>
   <p>Example:</p>
  #code-yml#
   file: /plugin/excel/php_excel_v18/demo.ods
   sheet: demo
   range: 'A1:C5'
  #code#
   <p>Optional first row can be used as keys.</p>
  #code-yml#
   first_row_as_key: true
  #code#
   <p>Returning an array.</p>
   */
  public function getRange($data){
    wfPlugin::includeonce('wf/array');
    $data = new PluginWfArray($data);
    $objPHPExcel = new PHPExcel();
    $dataFfile = wfArray::get($GLOBALS, 'sys/app_dir').$data->get('file');
    $objPHPExcel = PHPExcel_IOFactory::load($dataFfile);
    $objPHPExcel->setActiveSheetIndexByName($data->get('sheet'));
    $active_sheet = $objPHPExcel->getActiveSheet();
    $range = $active_sheet->rangeToArray($data->get('range'));
    if(!$data->get('first_row_as_key')){
      return $range;
    }else{
      $as_key = array();
      $array = array();
      foreach ($range as $key => $value) {
        if($key==0){
          foreach ($value as $key2 => $value2) {
            $as_key[] = $value2;
          }
        }else{
          $temp = array();
          foreach ($value as $key2 => $value2) {
            $temp[$as_key[$key2]] = $value2;
          }
          $array[] = $temp;
        }
      }
      return $array;
    }
  }
  /**
   <p>Method to create file from column data.</p>
   */
  public function columnToFile($file = '/plugin/excel/php_excel_v18/demo.ods', $sheet = 'yml', $column = 'A', $range_from = 1, $range_to = 5, $to_file = "/theme/[theme]/column_to_file.yml", $yml_validator = true ){
    $range = $this->getRange(array('file' => $file, 'sheet' => $sheet, 'range' => "$column$range_from:$column$range_to" ));
    $str = null;
    foreach ($range as $key => $value) {
      $str .= $value[0]."\n";
    }
    if($to_file){
      $data_save = wfArray::get($GLOBALS, 'sys/app_dir').wfSettings::replaceTheme($to_file);
      $save = true;
      if($yml_validator){
        try {
          sfYaml::load($str);
          $save = true;
        } catch (Exception $exc) {
          $save = false;
        }
      }
      if($save){
        wfFilesystem::saveFile($data_save, $str);
        return array('success' => true, 'file' => wfSettings::replaceTheme($to_file), 'str' => $str);
      }else{
        return array('success' => false, 'description' => 'Could not validate as YML!', 'str' => $str);
      }
    }else{
      return array('success' => false, 'description' => 'to_file is not set!', 'str' => $str);
    }
    
  }
  
  
  
}















