<?php
   require_once 'lib/PHPExcel.php';
   require_once 'lib/PHPExcel/IOFactory.php';
   
 
  

 /**
  * 获取文件
  */ 
  function getfile($filename){
   if (file_exists("files/" . $_FILES[$filename]["name"]))
      {
      unlink ( "files/" . $_FILES[$filename]["name"]);
      }
      move_uploaded_file($_FILES[$filename]["tmp_name"],
      "files/" . $_FILES[$filename]["name"]);
      echo "Stored in: " . "files/" . $_FILES[$filename]["name"];
      
    $file_path= "files/" . $_FILES[$filename]["name"];
    return $file_path;
  }   
 
   
   
/**
 * 文件读成二维数组
 */
    function TxtToArray($file_path){
    $strs[]="";   
    if(file_exists($file_path)){
       $file_arr= file($file_path);   
       for($i=0;$i<count($file_arr);$i++){//按行读取
           $str =array();    
           $file_arr[$i]=preg_replace ( "/\s(?=\s)/","\\1", $file_arr[$i] );//删除多余的空格只留下1个
           $str = explode(" ", $file_arr[$i]);        
            array_push($strs,$str);
         }            
       }
       $strs=array_splice ($strs,1);     
       return $strs;
    }

  
/**
 * 查找excel文件内容
 */
 //匹配Excel数据
function FindIdInExcel($filePath,$findname,$col1,$col2){
  $PHPReader=new \PHPExcel_Reader_Excel2007(); 
  
    //判断文件类型
   if (!$PHPReader->canRead($filePath)) {
   $PHPReader = new \PHPExcel_Reader_Excel5(); 
   if (!$PHPReader->canRead($filePath)) {
    echo 'no Excel';
    return false;
   }
}
  
  $PHPExcel = $PHPReader->load($filePath);  
  /**读取excel文件中的第一个工作表*/ 
  $currentSheet = $PHPExcel->getSheet(0);
  
  /**取得一共有多少行*/  
  $allRow = $currentSheet->getHighestRow();
  //查找状态
  $findstatus=0;
//echo $allColumn."+".$allRow;
  for($rowIndex = 1; $rowIndex <= $allRow; $rowIndex ++){
  	$location = $col1.$rowIndex;
  	$cell = $currentSheet->getCell($location)->getValue();
  	if($cell==$findname){
  		$findlocaltion=$col2.$rowIndex;
  		$findvalue = $currentSheet->getCell($findlocaltion)->getValue();
  		$findstatus=1;
		return $findvalue;		
  	}
  }
    if($findstatus==0){
  		$errormsg="nofind";
        return $errormsg;
  	}
	unset($PHPReader);
 }
 
 /**
  * 检查表的类型
  */
 function isExcel($PHPReader,$filePath){
   //判断文件类型
   if (!$PHPReader->canRead($filePath)) {
   $PHPReader = new \PHPExcel_Reader_Excel5(); 
   if (!$PHPReader->canRead($filePath)) {
    echo 'no Excel';
    return false;
    exit;
   }
  }
 }
 
 
/**
 * 读取excel文件
 */
   function readExcelToArray($file_path){    

    try {
      $inputFileType = PHPExcel_IOFactory::identify($file_path);
      $objReader = PHPExcel_IOFactory::createReader($inputFileType);
      $objPHPExcel = $objReader->load($file_path);
     }
    catch(Exception $e)
    {
      die("加载文件发生错误：".pathinfo($file_path,PATHINFO_BASENAME).": ".$e->getMessage());
    }
    $currentSheet = $objPHPExcel->getSheet(0);
    $allRow = $currentSheet->getHighestRow();
    $allColumn = $currentSheet->getHighestColumn();     
//  echo $highestRow."+".$highestColumn;

// $rowData = $sheet->rangeToArray('A' . $row . ":" . $highestColumn . $rowIndex, NULL, TRUE, FALSE);
// $cell = $currentSheet->getCell($location)->getValue();

   $dataExcel = $currentSheet->toArray();
   
   $dataExcel=array_splice ($dataExcel,1);   
// print_r($dataExcel);
   
   return $dataExcel;
   }
?>