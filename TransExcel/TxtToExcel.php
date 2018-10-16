<?php
	header("Content-Type:text/html;charset=gb2312");
	require_once 'ExcelFun.php';



  if(!empty($_FILES['file1']['tmp_name'])){
    echo "file1upload";
    $file1=getfile('file1');
    $result=TxtToArray($file1);
    $result=array_iconv("GBK","UTF-8",$result);//转化成utf-8格式，因为在代码中处理时不支持GBK
    createExcel($result,"file1","");
  }

   if(!empty($_FILES['file2']['tmp_name'])){
     echo "file2upload";
     $file2=getfile('file2');
     $result=TxtToArray($file2);
     print_r($result);
     $result=array_iconv("GBK","UTF-8",$result);//转化成utf-8格式，因为在代码中处理时不支持GBK
     createExcel($result,"file2","lib/bankid.xls");
  }
     
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

   exit;

 /**
  * //转化成utf-8格式，因为在代码中处理时不支持GBK
  */
     function array_iconv($in_charset,$out_charset,$arr){
       return eval('return '.iconv($in_charset,$out_charset,var_export($arr,true).';'));  
       } 
  
/**
 * 多维数组生成excel
 *  createExcel($data,$filename,$referfile);
 *  要生成的数组，判断上传的文件，是否有需要匹配的文件
 */
    function createExcel($data,$filename,$referfile){
        ob_end_clean();
    	error_reporting(E_ALL);
    	ini_set('display_errors', TRUE);
    	ini_set('display_startup_errors', TRUE);
    	date_default_timezone_set('PRC');
        // 创建Excel文件对象
    	$objPHPExcel = new PHPExcel();
    	
    	    	
    //
       if($filename=="file1"){ 
       	
       	//添加标题
    	$objPHPExcel->setActiveSheetIndex(0)
    	->setCellValue('A1', '类型')
    	->setCellValue('B1', '序号')
    	->setCellValue('C1', '客户名称')
    	->setCellValue('D1', '客户代码')
    	->setCellValue('E1', '委托单位')
    	->setCellValue('F1', '委托帐号')
    	->setCellValue('G1', '时间')
    	->setCellValue('H1', '详情')
    	->setCellValue('I1', '月租')
    	->setCellValue('J1', '付款金额')
    	->setCellValue('K1', '客户编号');    
       	   
    	for($i=0;$i<count($data);$i++){
    		$b=$i+2;    	
    		
    		$str9=$data[$i][7];
    		$str9=substr($str9,0,strlen($str9)-33);
    		$str9=substr_replace($str9,'.',strlen($str9)-2,0);
    		$str9=round($str9,2);
    		$objPHPExcel->getActiveSheet()->getStyle ('J')->getNumberFormat()->setFormatCode ("0.00");
    		
    		$userid=date('Ymd').sprintf("%04d", $b);
    		
    	  	$objPHPExcel->setActiveSheetIndex(0)
    	  	    ->setCellValue('A'.$b, $data[$i][0])
    	  	    ->setCellValue('B'.$b, $data[$i][1])
    	  	    ->setCellValue('C'.$b, $data[$i][2])  	    
    	  	  	->setCellValue('D'.$b, $data[$i][3] )
    	  	 	->setCellValue('E'.$b, $data[$i][4])
    	  	 	->setCellValue('F'.$b, $data[$i][5])
    	  	 	->setCellValue('G'.$b, $data[$i][6])
    	  	 	->setCellValue('H'.$b, $data[$i][7])	
    	  	 	->setCellValue('I'.$b, $data[$i][8])
    	  	 	->setCellValue('J'.$b, $str9)
    	  	 	->setCellValue('K'.$b, $userid);
    	 } 	
    	$FileName="系统文件".date('Ymd').".xlsx";	
    	}
    	
    //
       if($filename=="file2"){
       //添加标题
    	$objPHPExcel->setActiveSheetIndex(0)
    	->setCellValue('A1', '签约日期')
    	->setCellValue('B1', '业务种类')
    	->setCellValue('C1', '客户编号')
    	->setCellValue('D1', '客户名称')
    	->setCellValue('E1', '联系人名称')
    	->setCellValue('F1', '联系人地址')
    	->setCellValue('G1', '联系人邮编')
    	->setCellValue('H1', '联系人电话')
    	->setCellValue('I1', '委托单位代码')
    	->setCellValue('J1', '付款行行号')
    	->setCellValue('K1', '付款人账号')
    	->setCellValue('L1', '付款人证件类型（可选）')
    	->setCellValue('M1', '付款人证件号码（可选）')
    	->setCellValue('N1', '付款人手机号码（可选）')
    	->setCellValue('O1', '付款人电子邮箱（可选）')
    	->setCellValue('P1', '附言（可选）');
    	
    	for($i=0;$i<count($data);$i++){
    		$b=$i+2;
    		$userid=date('Ymd').sprintf("%04d", $b);
            $bankpayid=FindIdInExcel($referfile,$data[$i][4],"B","G");
    	  	$objPHPExcel->setActiveSheetIndex(0)
    	  	    ->setCellValue('A'.$b, '')
    	  	    ->setCellValue('B'.$b, '00500')
    	  	    ->setCellValue('C'.$b, $userid)  	    
    	  	  	->setCellValue('D'.$b, $data[$i][2] )
    	  	 	->setCellValue('E'.$b, '中移铁通')
    	  	 	->setCellValue('F'.$b, '广州越秀')
    	  	 	->setCellValue('G'.$b, '510000')
    	  	 	->setCellValue('H'.$b, '02061281614')	
    	  	 	->setCellValue('I'.$b, 'MA59B7XE8')
    	  	 	->setCellValue('J'.$b, $bankpayid)
    	  	 	->setCellValue('K'.$b, $data[$i][5])
    	  	 	->setCellValue('L'.$b, '')
    	  	 	->setCellValue('M'.$b, '')
    	  	 	->setCellValue('N'.$b, '')
    	  	 	->setCellValue('O'.$b, '')
    	  	 	->setCellValue('P'.$b, $data[$i][6].$data[$i][1]);
    	  }  
    	  $FileName="协议".date('Ymd').".xlsx"; 
    	}
    	
    	if($filename=="file4"){
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
    	}

    	// 保存Excel 2007格式文件，保存路径为当前路径,自动下载
    	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    	    	
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
    	}
    	

?>