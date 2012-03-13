<?php

include '../PHPReport.php';

//which template to use
if(isset($_GET['template']))
	$template=$_GET['template'];
else
	$template='stats.xls';
	
//set absolute path to directory with template files
$templateDir='C:\Ampps\www\PHPReport\examples\template\\';

//we get some data, e.g. from database
//function generates some random data
function getData($n=1,$cols=array('A'),$rows=1,$c=array('Some country'))
{
	//data is an array with 34 elements for each row!
	$data=array();
	for($i=1;$i<=$n;$i++)
	{
		$d['country']=$c[$i-1];
		foreach($cols as $col)
		{
			for($r=1;$r<=$rows;$r++)
			{
				$d[$col.$r]=rand(0,20);
			}
		}
		$data[]=$d;
	}
	
	return $data;
}

$data=getData(3,array('A','B','C'),11,array('Italy','Germany','France'));

//set config for report
$config=array(
		'template'=>$template,
		'templateDir'=>$templateDir
	);

$R=new PHPReport($config);
$R->load(array(
				'id'=>'v',
				'repeat'=>true,
				'data'=>$data,
			)
        );
$R->setHeading('Report: Visitors in January');
echo $R->render('html');
exit();