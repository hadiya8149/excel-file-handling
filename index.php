<?php 
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet; 
function donwloadExcelFile(){
    $json_file = file_get_contents("./data.js");
    $json_data =  json_decode($json_file, true);
    
    $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    $spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(150, 'pt');

    $userWorksheet = $spreadsheet->getActiveSheet();
    $userWorksheet->setTitle('User data');

    $headerRow = ['username', 'email', 'amount', 'status', 'total amount for active status', 'total amount for inactive status', 'total words in username'];
    $usernameColumn = array_chunk(array_column($json_data, "username"), 1);
    $userEmailColumn =  array_chunk(array_column($json_data, "email"), 1);
    $userAmountColumn =  array_chunk(array_column($json_data, "amount"), 1);
    $userStatusColumn =  array_chunk(array_column($json_data, "status"), 1);
    $usernameArray = array_column($json_data, "username");
    function countWords($username){
        return str_word_count($username);
    }
    $totalWordsInUsernameColumn=array_map('countWords', $usernameArray);
    var_dump($totalWordsInUsernameColumn);
    $totalWordsInUsernameColumn = array_sum($totalWordsInUsernameColumn);
    
    
    $userWorksheet->fromArray($headerRow, NULL, 'A1');
    $userWorksheet->fromArray($usernameColumn,NULL, 'A2');
    $userWorksheet->fromArray($userEmailColumn,NULL, 'B2');
    $userWorksheet->fromArray($userAmountColumn,NULL, 'C2');
    $userWorksheet->fromArray($userStatusColumn,NULL, 'D2');

    $userWorksheet->setCellValue('E2', '=SUM(FILTER(C2:C20, D2:D20=D2, "no results"))');
    $userWorksheet->setCellValue('F2', '=SUM(FILTER(C2:C20, D2:D20=D3, "no results"))');
    $userWorksheet->setCellValue('G2', $totalWordsInUsernameColumn);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="myfile.xlsx"');
    header('Cache-Control: max-age=0');
    
    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');

    $writer->save('php://output');
}
if($_SERVER['REQUEST_METHOD']==='POST'){
    donwloadExcelFile();
}
?>
<div style="height:100%;margin:auto;display:flex;align-items:center;justify-content:center;">
<form action="index.php" style="height:500px;width:500px;text-align:center;margin:auto;padding:32px;" method="POST">
<h1>Download Converted php file</h1>
<button type="submit" class="download" style="height:100px;width:300px;background:#2196f3;border:none;color:white;font-size:32px;hover:cursor;">Downlaod</button></form>
<style>
    button:hover{
        cursor:pointer;
    }
    </style>
</div>
