<?php
   require_once 'ExcelFun.php';


$file_path="";
   if(!empty($_FILES['file3']['tmp_name'])){    
    $file3=getfile('file3');
    $file_path=$file3;
    }
  
   
  if(file_exists($file_path)){  	
    if(!empty($_FILES['file4']['tmp_name'])){
     $file4=getfile('file4');     
     $result=readExcelToArray($file4);
     createExcel2($result,"file4",$file_path);
    } 
  }
/**
 * 单位提回文件处理 回应结果
 * 转成
 * 代收 *
 *  createExcel($data,$filename,$referfile);
 *  要生成的数组，判断上传的文件，是否有需要匹配的文件
 *  */

function createExcel2($data,$filename,$referfile){
        ob_end_clean();
    	error_reporting(E_ALL);
    	ini_set('display_errors', TRUE);
    	ini_set('display_startup_errors', TRUE);
    	date_default_timezone_set('PRC');
//      print_r($data);  
//      exit;
     // 创建Excel文件对象
    	$objPHPExcel = new PHPExcel();
    	
     $objPHPExcel->setActiveSheetIndex(0)
    	->setCellValue('A1', '委托日期')
    	->setCellValue('B1', '客户编号')
    	->setCellValue('C1', '付款人开户行行号')
    	->setCellValue('D1', '付款人账号')
    	->setCellValue('E1', '付款人名称')
    	->setCellValue('F1', '付款协议号')
    	->setCellValue('G1', '金额')    	
    	->setCellValue('H1', '附言（可选）');
    	
    	$objPHPExcel->getActiveSheet()->getStyle ('G')->getNumberFormat()->setFormatCode ("0.00");//设置格式
        $b=2;
        
    	for($i=0;$i<count($data);$i++){  		
    		if($data[$i][11]=="成功"){ 
    			
            $price=FindIdInExcel($referfile,$data[$i][4],"K","J"); 
//          $price="";         
    	  	$objPHPExcel->setActiveSheetIndex(0)
    	  	    ->setCellValue('A'.$b, '')
    	  	    ->setCellValue('B'.$b, $data[$i][4])
    	  	    ->setCellValue('C'.$b, $data[$i][8])  	    
    	  	  	->setCellValue('D'.$b, $data[$i][9] )
    	  	 	->setCellValue('E'.$b, $data[$i][10])
    	  	 	->setCellValue('F'.$b, $data[$i][6])
    	  	 	->setCellValue('G'.$b, $price)
    	  	 	->setCellValue('H'.$b, $data[$i][26]);
    	  	$b++;
    		}
    		
    	 }	
    	// 保存Excel 2007格式文件，保存路径为当前路径，名字为export.xlsx
    	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    	//选择默认保存路径
    	$FileName="	代收.xlsx";
    	    	
    	header("Content-Type: application/force-download"); 
        header("Content-Type: application/octet-stream"); 
        header("Content-Type: application/download"); 
        
        header('Content-Disposition:inline;filename="'.$FileName.'"'); 
        header("Content-Transfer-Encoding: binary"); 
        header("Expires: Mon, 26 Jul 1997 05:00:00 GMT"); 
        header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT"); 
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0"); 
        header("Pragma: no-cache"); 
        $objWriter->save('php://output'); 
        
//      deldir('/files/');
    	}

?>