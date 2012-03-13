<?php

include '../PHPReport.php';

$R=new PHPReport();
$R->load(array(
            'id'=>'product',
            'data'=>array(
                        array('Some product',23.99,12),
                        array('Other product',5.25,2.25),
                        array('Third product',0.20,3.5)
                )
            )
        );

echo $R->render();
exit();