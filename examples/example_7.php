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
			'config'=>array(
					'header'=>array(
						0=>array('width'=>120,'align'=>'center'),
						1=>array('width'=>80,'align'=>'center'),
						2=>array('width'=>80,'align'=>'center')
					),
					'data'=>array(
						0=>array('align'=>'left'),
						1=>array('align'=>'right'),
						2=>array('align'=>'right')
					),
					'footer'=>array(
						1=>array('align'=>'right'),
						2=>array('align'=>'right')
					)
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