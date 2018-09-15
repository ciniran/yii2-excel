<?php
/**
 * Created by PhpStorm.
 * User: xiebo
 * Date: 2018/9/15
 * Time: 上午9:31
 */

namespace ciniran\excel;


use Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use yii\base\Component;
use yii\base\Model;
use yii\data\ActiveDataProvider;
use yii\db\ActiveRecord;

class SaveExcel extends Component
{
    /**
     * excel97格式
     */
    const XLS = 'xls';
    /**
     * excel2001格式
     */
    const XLXS = 'xlsx';
    /**
     * @var array|string $fields 需要输出的列
     */
    public $fields;
    /**
     * @var array $relation 需要输出的关联数据
     */
    public $relation = [];
    /**
     * @var bool $show 内容值转换
     * 如果需要对输出到excel表格的值进行转换
     * 在输出模型中定义一个名称为show的function即可
     */
    public $show = false;
    /**
     * @var string $fileName 指定输出名称,默认为当前时间值
     */
    public $fileName;
    /**
     * @var self::XLS | self::XLXS 指定输出格式
     */
    public $format;
    /**
     * @var array $array 需要输出的数组
     */
    public $array;

    /**
     * @var ActiveRecord[] $models 要导出的模型数据
     */
    public $models;
    /**
     * @var ActiveDataProvider $dataProvider
     */
    public $dataProvider;
    /**
     * @var bool $all 导出全部数据
     */
    public $all = true;
    /**
     * @var array $headerDataArray 列名数据
     */
    public $headerDataArray;
    private $bodyDataArray;


    public function init()
    {
        if (!$this->models && !$this->dataProvider && !$this->array) {
            throw new Exception('models,dataProvider,array,必需指定一个');
        }
        parent::init();
    }

    /**
     * 通过dataProvider生成excel文件
     * @throws Exception
     */
    public function dataProviderToExcel()
    {
        if ($this->all) {
            $this->dataProvider->pagination = false; //不使用分页，导出全部数据
        }
        $this->models = $this->dataProvider->getModels();
        if (!$this->models) {
            throw new Exception('没有数据无法导出');
        }
        $this->modelsToExcel();
    }
    /**
     * 通过模型生成excel文件
     * @throws Exception
     */
    public function modelsToExcel()
    {
        if (!$this->models) {
            throw new Exception('属性models,不能为空');
        }
        $this->checkFields();
        $this->getRelationModels();
        $this->initFields();
        $this->initHeaderData();
        $this->initBodyData();
        $this->arrayToExcel();
    }



    /**
     * 通过数组生成excel文件
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function arrayToExcel()
    {
        if (!$this->fileName) {
            $this->fileName = date('ymdhis');
        }
        if (!$this->bodyDataArray) {
            if (!$this->array) {
                throw new Exception('属性array不能为空');
            }
            $this->bodyDataArray = $this->array;
        }
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        if ($this->headerDataArray) {
            $sheet->fromArray($this->headerDataArray, null, 'A1');
            $sheet->fromArray($this->bodyDataArray, null, 'A2');
        }else{
            $sheet->fromArray($this->bodyDataArray, null, 'A1');
        }
        if ($this->format == self::XLS) {
            $writer = new Xls($spreadsheet);
            ob_start();
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename=' . $this->fileName . '.xls');
            header('Cache-Control: max-age=0');
            $writer->save('php://output');
        } else {
            $writer = new Xlsx($spreadsheet);
            ob_start();
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename=' . $this->fileName . '.xlsx');
            header('Cache-Control: max-age=0');
            $writer->save('php://output');
        }


    }






    private function getRelationModels()
    {
        if ($this->relation) {
            /** @var ActiveRecord $item */
            foreach ($this->models as $item) {
                foreach ($this->relation as $value) {
                    $rModel = $item->$value;
                    if (count($rModel) > 1) {
                        foreach ($rModel as $k => $v) {
                            if ($k > 0) {
                                $copyOrderData = clone $item;
                                $copyOrderData->populateRelation($value, [$v]);
                                $this->models[] = $copyOrderData;
                            }
                        }
                    }
                }
            }
            unset($item, $rModel, $copyOrderData, $v, $value);
            if ($this->models[0]->primaryKey) {
                //排序
                $pk = [];
                foreach ($this->models as $item) {
                    $pk[] = $item->primaryKey;
                }
                array_multisort($this->models, $pk);
            }
        }
    }

    /**
     * @throws Exception
     */
    protected function initHeaderData()
    {
        $res = [];
        if ($this->fields) {
            $model = $this->models[0];
            foreach ($this->fields as $item) {
                $label = $model->getAttributeLabel($item);
                if (!preg_match("/[\x7f-\xff]/", $label) && $this->relation) {
                    if (count($this->relation) > 1) {
                        foreach ($model->getRelatedRecords() as $value) {
                            $label = $value->getAttributeLabel($item);
                            break;
                        }
                    } else {
                        $temp = $this->relation[0];
                        $rModel = $model->$temp;
                        if (is_object($rModel)) {
                            $label = $rModel->getAttributeLabel($item);
                        } else {
                            $label = $rModel[0]->getAttributeLabel($item);
                        }
                    }
                }
                if (!$label) {
                    throw new Exception("请检查导出字段:$item,是否存在!");
                }
                $res[] = $label;
            }
        } else {
            //没有指定列时不导出关联字段
            /** @var ActiveRecord $item */
            $item = $this->models[0];
            $res = $item->attributeLabels();
            $this->fields = array_keys($res);
        }

        $this->headerDataArray = $res;
    }

    /**
     * @throws Exception
     */
    protected function initBodyData()
    {
        $res = [];
        /** @var ActiveRecord $item */
        foreach ($this->models as $key => $item) {

            if ($this->show) {
                if (!method_exists($item, 'show')) {
                    throw new Exception(get_class($item) . '中未定义show()方法,不能进行值的转换');
                }
                $item->show();
            }
            foreach ($this->fields as $attribute) {
                $res[$key][$attribute] = $this->getValue($item, $attribute);
            }

        }
        $this->bodyDataArray = $res;
    }

    /**
     * 检查要导出的列是否存在
     * @throws Exception
     */
    protected function checkFields()
    {
        if (!$this->fields) {
            return;
        }
        /** @var ActiveRecord $item */
        $item = $this->models[0];
        $modelFields = array_keys($item->getAttributes());
        if ($this->relation) {
            if (is_string($this->relation)) {
                $this->relation = [$this->relation];
            }
            foreach ($this->relation as $value) {
                try {
                    $item->$value;
                } catch (Exception $exception) {
                    throw new Exception('模型中不存在名为:' . $value . '的关联关系');
                }
                /** @var ActiveRecord $relateModel [$value] */
                $rModel = $item[$value];
                if (!$rModel) {
                    continue;
                }
                if (is_object($rModel)) {
                    $relationFields = array_keys($rModel->getAttributes());
                } else {
                    $relationFields = array_keys($item[$value][0]->getAttributes());
                }
                $modelFields = array_merge($modelFields, $relationFields);
                unset($relationFields);
            }
        }
        $fields = $this->fields;
        $temp = array_intersect($modelFields, $fields);
        if (sort($temp) != sort($fields)) {
            throw new Exception('导出的字段在数据中并不存在，请检查');
        }
    }

    /**
     * 取得相应的列值
     * @param ActiveRecord $model
     * @param string $attribute
     * @return mixed|null
     */
    protected function getValue($model, $attribute)
    {
        if ($model->hasProperty($attribute)) {
            return $model->getAttribute($attribute);
        } else {
            foreach ($this->relation as $value) {
                $rModel = $model->$value;
                if (!$rModel) {
                    return null;
                }
                if (is_object($rModel)) {
                    return $rModel->$attribute;
                } else {
                    return $rModel[0]->$attribute;
                }
            }
        }

    }


    private function initFields()
    {
        if (!$this->fields) {
            return;
        }
        if (is_string($this->fields)) {
            $this->fields = explode(",", $this->fields);
        }
    }


}