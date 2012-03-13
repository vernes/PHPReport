<?php

include '../PHPReport.php';

$R=new PHPReport();
$R->load(array(
            'id'=>'product',
			'header'=>array(
					'product'=>'Product name','price'=>'Price','tax'=>'Tax'
				),
			'footer'=>array(
					'product'=>'','price'=>28.54,'tax'=>17.89
				),
			'config'=>array(
					'header'=>array(
						'product'=>array('width'=>120,'align'=>'center'),
						'price'=>array('width'=>80,'align'=>'center'),
						'tax'=>array('width'=>80,'align'=>'center')
					),
					'data'=>array(
						'product'=>array('align'=>'left'),
						'price'=>array('align'=>'right'),
						'tax'=>array('align'=>'right')
					),
					'footer'=>array(
						'price'=>array('align'=>'right'),
						'tax'=>array('align'=>'right')
					)
				),
            'data'=>array(
					array('product'=>'Some product','price'=>23.99,'tax'=>12),
					array('product'=>'Other product','price'=>5.25,'tax'=>2.25),
					array('product'=>'Third product','price'=>0.20,'tax'=>3.5)
                ),
			'group'=>array(
					'caption'=>array(
						'Category 1', 'Another category'
					),
					'rows'=>array(
						array(0),
						array(1,2)
					),
					'summary'=>array(
						array('product'=>'','price'=>23.99,'tax'=>12),
						array('product'=>'','price'=>5.45,'tax'=>5.75)
					)
				),
			'format'=>array(
					'price'=>array('number'=>array('prefix'=>'$','decimals'=>2)),
					'tax'=>array('number'=>array('sufix'=>' EUR','decimals'=>1))
				)
            )
        );

echo $R->render();
exit();