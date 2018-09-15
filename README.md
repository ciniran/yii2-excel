yii2,excel
==========
Yii2  excel 导入，导出扩展，
本扩展基于phpexcel的升级版本 [phpspreadsheet](http://github.com/phpoffice/phpspreadsheet).
主要有以下功能：

## 导出
- 基于activeDataPrivoder生成Excel
- 基于ActiveRecrod生成Excel
- 基于二维数组生成Excel
- 模型中定义show(),可对输出值进行转换
- 指定导出字段
- 支持导出model的关联数据
- 自定义输出版本

## 导入
- 直接生成ActiveRecrod的模型
- 按表头生成数组
- 指定表头位置
- 可按模型attribute导入
- 按attributeLabel中文导入

## 安装
------------

您可以通过 [composer](http://getcomposer.org/download/) 安装.

本项目在github的地址是[http://github.com/ciniran/yii2-excel](http://github.com/ciniran/yii2-excel)

## 运行

```
php composer.phar require --prefer-dist ciniran/yii2-excel "*"
```

or add

```
"ciniran/yii2-excel": "*"
```

to the require section of your `composer.json` file.


## 用法
-----

下面是一些简单示例：

### DataPrivoder to excel:

```
      $dataPrivoder = new ActiveDataProvider([
            'query' => User::find(),
        ]);
        $excel = new SaveExcel([
            'dataProvider' => $dataPrivoder,
          //'show' => true,  //是否对值进行转换
            'fields' => 'email,username', //or ['email,username'] 限制导出的列
            'format' => SaveExcel::XLXS, 输出版本
            'all' => true,  //导出全部数据
            'relation' => false, //模型关系数据
        ]);
        $excel->dataProviderToExcel();

```

### ActiveRecord to excel:
```
        $models = User::find()->all();
        $excel = new SaveExcel([
            'models' => $models,
            // 'show' => true,
        ]);
        $excel->modelsToExcel();
```
### Array to excel

```
     $array = [
            [
                'name'=>'tom',
                'age'=>18,
            ],
            [
                'name'=>'jerry',
                'age'=>19,
            ],
        ];
        $excel = new SaveExcel([
            'array' => $array,
            'headerDataArray' => ['name', 'age'],
        ]);
        $excel->arrayToExcel();

```

### Read excel file to array
```
       $path = 'user.xlsx';
        $excel = new ReadExcel([
            'path' => $path,
            'head' => true,
            'headLine' => 1,
        ]);
        $data = $excel->getArray();
```
### Read excel file to models
```
        $path = 'user.xlsx';
        $excel = new ReadExcel([
            'path' => $path,
            'head' => true,
            'headLine' => 1,
            'class' => 'common\\models\\User',
            'useLabel' => true,
        ]);
        $models = $excel->getModels();
```

下方是英文说明
===========================================================


yii2,excel
==========
excel tools

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist ciniran/yii2-excel "*"
```

or add

```
"ciniran/yii2-excel": "*"
```

to the require section of your `composer.json` file.


Usage
-----

Once the extension is installed, simply use it in your code by  :

### DataPrivoder to excel:

```
      $dataPrivoder = new ActiveDataProvider([
            'query' => User::find(),
        ]);
        $excel = new SaveExcel([
            'dataProvider' => $dataPrivoder,
          //'show' => true,
            'fields' => 'email,username', //or ['email,username']
            'format' => SaveExcel::XLXS,
            'all' => true,
            'relation' => false, //模型关系数据
        ]);
        $excel->dataProviderToExcel();

```

### ActiveRecord to excel:
```
        $models = User::find()->all();
        $excel = new SaveExcel([
            'models' => $models,
            // 'show' => true,
        ]);
        $excel->modelsToExcel();
```
### Array to excel

```
     $array = [
            [
                'name'=>'tom',
                'age'=>18,
            ],
            [
                'name'=>'jerry',
                'age'=>19,
            ],
        ];
        $excel = new SaveExcel([
            'array' => $array,
            'headerDataArray' => ['name', 'age'],
        ]);
        $excel->arrayToExcel();

```

### Read excel file to array
```
       $path = 'user.xlsx';
        $excel = new ReadExcel([
            'path' => $path,
            'head' => true,
            'headLine' => 1,
        ]);
        $data = $excel->getArray();
```
### Read excel file to models
```
        $path = 'user.xlsx';
        $excel = new ReadExcel([
            'path' => $path,
            'head' => true,
            'headLine' => 1,
            'class' => 'common\\models\\User',
            'useLabel' => true,
        ]);
        $models = $excel->getModels();
```
