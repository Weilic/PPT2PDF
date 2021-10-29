#$currentParentPath = Split-Path $MyInvocation.MyCommand.Definition -Parent #获取当前父路径
$pptFileList = Get-ChildItem | Where-Object{$_.Name -like "*.ppt*"}#提取出ppt或pptx文件列表
$pdfFileList = Get-ChildItem | Where-Object{$_.Name -like "*.pdf*"}#提取出pdf文件列表
$pdfFileList = $pdfFileList.baseName
$conFile= @()#转换文件
foreach($n in $pptFileList){
    if(!($n.BaseName -in $pdfFileList)){
        $conFile+= $n.FullName
    }
}#除去已转化的ppt or pptx文件
$PowerPoint = New-Object -ComObject Powerpoint.application
#$PowerPoint.Visible = $false
foreach($n in $conFile){
    $pdfName = $n.split('.')[0] + '.pdf'
    $presentation = $PowerPoint.presentations.open($n,1,0)
    $presentation.SaveAs($pdfName,32)
    #$presentation.SaveCopyAs($pdfName,32)
    $presentation.close()
}
$PowerPoint.Quit()#关闭PowerPoint实例