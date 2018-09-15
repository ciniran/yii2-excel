<?php
/**
 * User: xieo
 * Date: 2018/9/15
 * Time: 下午2:23
 */

namespace ciniran\excel;


use yii\base\Component;
use yii\base\Exception;
use yii\base\Model;

class ReadExcel extends Component
{
    /**
     * @var string $path 文件路径
     */
    public $path;

    /**
     * @var string|Model 类名或者模型
     */
    public $class;

    /**
     * @var bool $head 是否包含表头
     */
    public $head = true;

    /**
     * @var int $headLine 表头所在行数,不填默认为第一行
     */
    public $headLine = 1;

    /**
     * @var bool $useLabel 表头使用label名称，如果不使用，表头应对应模型中的attribute值
     */
    public $useLabel = true;

    /**
     * @var array $data 原始数据
     */
    private $data;

    /**
     * @var Model[] $models 模型数据
     */
    private $models;

    public function init()
    {
        parent::init();
        if (!$this->path) {
            throw new Exception('属性path,不能为空');
        }
    }

    /**
     * 取得数组
     */
    public function getArray()
    {
        $this->readExcel();
        if (!$this->head) {
            return $this->data;
        }
        $this->normaliseArrayKey();
        return $this->data;
    }

    /**
     * 取得模型对象
     */
    public function getModels()
    {
        if (!$this->class) {
            throw new Exception('class不能为空');
        }
        if (is_object($this->class)) {
            $this->class = get_class($this->class);
        }
        $this->getArray();

        $this->models = array_map(function ($value) {
            return $this->getInstanceModel($value);
        }, $this->data);
        return $this->models;
    }


    /**
     * 通过文件路径读取excel内容
     * @param $path
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */

    public function readExcel()
    {
        $objPHPExcel = \PhpOffice\PhpSpreadsheet\IOFactory::load($this->path);
        $sheetData = $objPHPExcel->getActiveSheet()->toArray(null, true, true, true);
        $this->data = $sheetData;
    }

    private function normaliseArrayKey()
    {
        if (!$this->data) {
            throw new Exception('数据为空');
        }

        if (!array_key_exists($this->headLine, $this->data)) {
            throw new Exception('指定的表头行headLine数不正确');
        }
        $keys = $this->data[$this->headLine];
        if ($this->headLine > 1) {
            $this->data = array_slice($this->data, $this->headLine - 1);
        } elseif ($this->headLine == 1) {
            array_shift($this->data);
        }
        $newData = array_map(function ($v) use ($keys) {
            return array_combine($keys, $v);
        }, $this->data);
        $this->data = $newData;
    }

    private function getInstanceModel($value)
    {
        /** var \yii\base\Model $modes **/
        $model = new $this->class();
        if ($this->useLabel) {
            $attributesLabels = $model->attributeLabels();
            $keys = array_keys($value);
            $labels = array_flip($attributesLabels);
            foreach ($value as $k => $v) {
                if (key_exists($k, $labels)) {
                    $attribute = $labels[$k];
                    $model->$attribute = $v;
                }
            }
        }else{
            foreach ($value as $k => $v) {
                if($model->hasAttribute($k)){
                    $model->$k = $v;
                }
            }
        }

        return $model;
    }
}