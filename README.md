# PHPReport

PHPReport is a class for building and exporting reports from PHP. It's based on powerful [PHPExcel](http://www.phpexcel.net/) library and includes exporting to popular formats such are HTML, Excel or PDF.

## Features

* Simple for using
* Many ways to customize your input data
* Build reports based on predefined templates
* Export to HTML, Excel (xlsx and xls) and PDF


## Installation

Install [composer](https://getcomposer.org), and run `composer update` to install
the required `PHPExcel` dependencies

## Usage and examples

Basicly, there are two way to build a report.

### Building it "from scratch"

PHPReport does all the work, your just provide it with your data, like this:
<pre>
$R=new PHPReport();
$R->load(array(
            'id'=>'product',
            'data'=>array(
                        array('Some product',23.99),
                        array('Other product',5.25),
                        array('Third product',0.20)
                )
            )
        );

$R->render();
</pre>

It's great for exporting tabular data. Input data can be further formatted, grouped and customized. See the wiki.

### Building it from template

Template is usually some excel file, already formatted and with placeholders for data. There are two types of placeholders: static and dynamic.
Static placeholders are, for example, some data like date, city or customer name. Dynamic placeholders are, for example, list of products with variable number of rows.

<pre>
$R=new PHPReport(array('template'=>'invoice_template.xls'));
$R->load(array(
		array(
				'id'=>'invoice',
				'data'=>array(
					'date'=>date('Y-m-d'),
					'number'=>'000312',
					'customer_id'=>'512',
					'expiration_date'=>date('Y-m-d',strtotime('+30day')),
					'name'=>'John Doe',
					'company'=>'Example, inc',
					'address'=>'Some address',
					'city'=>'Gotham City',
					'zip'=>'0123',
					'phone'=>'+123456'
				),
				'format'=>array(
					'date'=>array('datetime'=>'d/m/Y'),
					'expiration_date'=>array('datetime'=>'d/m/Y')
				)
			),
		array(
				'id'=>'product',
				'data'=>array(
					array('description'=>'Some product','price'=>23.99,'total'=>23.99),
					array('description'=>'Other product','price'=>5.25,'total'=>2.25)
				),
				'repeat'=>true,
				'format'=>array(
					'price'=>array('number'=>array('prefix'=>'$','decimals'=>2)),
					'total'=>array('number'=>array('prefix'=>'$','decimals'=>2))
				)
			)
		)
	);

$R->render();
</pre>

These reports are great for complex exports like invoices. See the wiki for more examples.