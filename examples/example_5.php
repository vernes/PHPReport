<?php

include '../PHPReport.php';

$R=new PHPReport();
$R->load(array(
            'id'=>'product',
			'header'=>array(
					'Product name','Price','Tax'
				),
			'footer'=>array(
					'',28.54,17.89
				),
            'data'=>array(
					array('Some product',23.99,12),
					array('Other product',5.25,2.25),
					array('Third product',0.20,3.5)
                ),
			'format'=>array(
					1=>array('number'=>array('prefix'=>'$','decimals'=>2)),
					2=>array('number'=>array('sufix'=>' EUR','decimals'=>1))
				)
            )
        );

echo $R->render();
exit();