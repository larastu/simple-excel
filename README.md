# simple-excel

```
Excel操作简化版
```

## Excel导出

- 导出功能根据“PHP_XLSXWriter”项目改写
- PHP_XLSXWriter项目地址(https://github.com/mk-j/PHP_XLSXWriter)

### 简单使用示例如下

```php
$writer = new ExcelWriter();
$headers = [
    '姓名' => 'string',
    '年龄' => '',
    '入职日期' => 'date',
];

$sheet1 = $writer->createSheet(' name1 ');
$sheet1->setColumnTypes(array_values($headers));
$sheet1->addHeader(['员工列表'], [
    'height' => 20
]);
$sheet1->addHeader(array_keys($headers));
$sheet1->merge('a1:c1');
$sheet1->addRow(['张三','19','2021-10-01']);
$sheet1->addRow(['李四',20,'2022-10-01']);

$writer->writeToFile('test.xlsx');
```

