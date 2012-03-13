<?php

include '../PHPReport.php';


$R=new PHPReport();
$R->load(array(
        array(
            'id'=>'product',
            'data'=>array(
                        array('Some product',23.99,12),
                        array('Other product',5.25,2.25),
                        array('Third product',0.20,3.5)
                )
            ),
        array(
            'id'=>'product2',
            'data'=>array(
                        array('value1','value2','value3','value4'),
                        array('value5','value6','value7','value8')
                )
            )
        )
        );

echo $R->render();
exit();