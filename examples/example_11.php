<?php

include '../PHPReport.php';

//which template to use
if(isset($_GET['template']))
	$template=$_GET['template'];
else
	$template='invoice.xls';
	
//set absolute path to directory with template files
$templateDir='C:\Ampps\www\PHPReport\examples\template\\';

//we get some products, e.g. from database
$products=array(
	array('description'=>'Example product','qty'=>2,'price'=>4.5,'total'=>9),
	array('description'=>'Another product','qty'=>1,'price'=>13.9,'total'=>13.9),
	array('description'=>'Super product!','qty'=>3,'price'=>1.5,'total'=>4.5),
	array('description'=>'Yet another great product','qty'=>2,'price'=>10.8,'total'=>21.6),
	array('description'=>'Awesome','qty'=>1,'price'=>19.9,'total'=>19.9)
);

//set config for report
$config=array(
		'template'=>$template,
		'templateDir'=>$templateDir
	);

$R=new PHPReport($config);
$R->load(array(
			array(
				'id'=>'inv',
				'data'=>array('date'=>date('Y-m-d'),'number'=>312,'customerid'=>12,'orderid'=>517,'company'=>'Example Inc.','address'=>'Some address','city'=>'Some City, 1122','phone'=>'+111222333'),
				'format'=>array(
						'date'=>array('datetime'=>'d/m/Y')
					)
				),
			array(
				'id'=>'prod',
				'repeat'=>true,
				'data'=>$products,
				'minRows'=>2,
				'format'=>array(
						'price'=>array('number'=>array('prefix'=>'$','decimals'=>2)),
						'total'=>array('number'=>array('prefix'=>'$','decimals'=>2))
					)
				),
			array(
				'id'=>'total',
				'data'=>array('price'=>68.9),
				'format'=>array(
						'price'=>array('number'=>array('prefix'=>'$','decimals'=>2))
					)
				)
			)
        );

echo $R->render('html');
exit();