<?php
class PluginExcelPhp_excel_v18{
  function __construct($buto = false) {
    include_once __DIR__.'/PHPExcel-1.8/Classes/PHPExcel.php';
  }
  public function widget_demo($data){
    wfPlugin::includeonce('wf/array');
    $data = new PluginWfArray($data);
    wfHelp::yml_dump($data);
    $xls_data = $this->getRange($data->get('data'));
    wfHelp::yml_dump($xls_data);
  }
  public function getRange($data){
    wfPlugin::includeonce('wf/array');
    $data = new PluginWfArray($data);
    $objPHPExcel = new PHPExcel();
    $dataFfile = wfArray::get($GLOBALS, 'sys/app_dir'). wfSettings::replaceDir($data->get('file'));
    $objPHPExcel = PHPExcel_IOFactory::load($dataFfile);
    $objPHPExcel->setActiveSheetIndexByName($data->get('sheet'));
    $active_sheet = $objPHPExcel->getActiveSheet();
    $range = $active_sheet->rangeToArray($data->get('range'));
    $key_id = null;
    if($data->get('key_id')){
      $key_id = $data->get('key_id');
    }
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
          if(is_null($key_id)){
            $array[] = $temp;
          }else{
            $array[$temp[$key_id]] = $temp;
          }
        }
      }
      return $array;
    }
  }
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
  private function get_obj_from_data($rs){
    $objPHPExcel = new PHPExcel();
    $objPHPExcel->setActiveSheetIndex(0);
    $has_data = false;
    if(sizeof($rs)){
      $has_data = true;
    }
    /**
     * 
     */
    if($has_data){
      /**
       * 
       */
      $rowCount = 1;
      /**
       * First row.
       */
      $column = 0;
      foreach($rs[0] as $k => $v){
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($column, $rowCount, $k);
        $column++;
      }
      $rowCount++;
      /**
       * Data
       */
      foreach($rs as $v){
        $column = 0;
        foreach($v as $v2){
          $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($column, $rowCount, $v2);
          $column++;
        }
        $rowCount++;
      }
    }
    return $objPHPExcel;
  }
  private function get_filename($filename){
    if(!$filename){
      $filename = 'export_'.date('ymdHis').'.xlsx';
    }
    return $filename;
  }
  public function download_excel($rs, $filename = null){
    $objPHPExcel = $this->get_obj_from_data($rs);
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    /**
     * 
     */
    $filename = $this->get_filename($filename);
    /**
     * 
     */
   header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
    header('Content-Disposition: attachment;filename="'.$filename.'"');
    header('Cache-Control: max-age=0');
    /**
     * IE 9 issue.
     */
    header('Cache-Control: max-age=1');
    /**
     * IE over SSL issue.
     */
    header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
    header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT');
    header ('Cache-Control: cache, must-revalidate');
    header ('Pragma: public');
    /**
     * 
     */
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('php://output');
  }
  public function save_excel($rs, $filename){
    $objPHPExcel = $this->get_obj_from_data($rs);
    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
    $objWriter->save(wfGlobals::getAppDir().$filename);
  }
}
