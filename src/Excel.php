<?php
/**
 * Created by PhpStorm.
 * User: boxie
 * Date: 2018/9/15
 * Time: 上午9:31
 */

namespace ciniran\excel;


use Exception;
use yii\base\Component;
use yii\base\Model;
use yii\data\ActiveDataProvider;
use yii\db\ActiveRecord;

class Excel extends Component
{
    public $fields = [];
    public $relation = [];
    public $show = true;
    public $fileName = "";
    /**
     * @var ActiveRecord[] $models 要导出的模型数据
     */
    public $models;
    /**
     * @var ActiveDataProvider $dataProvider
     */
    public $dataProvider;
    public $all = true;

    private $headerDataArray;
    private $bodyDataArray;


    public function init()
    {
        parent::init();
    }

    /**
     * @param ActiveDataProvider $dataProvider
     * @param                    $fileName
     * @param array              $fields
     * @param string             $relation 导出关联明细
     * @param bool               $show     是否显示值对应的文本
     * @throws Exception
     */
    public function dataProviderToExcel()
    {
        $this->initModels();

    }

    /**
     * 下载表格
     * @param      $headData
     * @param      $bodyData
     * @param null $fileName
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     * @throws \PHPExcel_Writer_Exception
     */
    public static function arrayToExcel($headData , $bodyData , $fileName = null)
    {
        $excel = new \PHPExcel();
        $excel->setActiveSheetIndex(0);
        /*
              $columnLength = count($headData);
                  if($columnLength > 26){
                      $lastColName = "A".chr($columnLength+64-26);
                 }else{
                      $lastColName = chr($columnLength+64);

                  }
              //报表头标题的输出
              $excel->getActiveSheet()->mergeCells('A1:'.$lastColName.'1');
              $excel->getActiveSheet()->setCellValue('A1',$fileName)->getStyle()->getFont()->setBold(true)->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
              $excel->getActiveSheet()->getCell('A1')->getStyle() ->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
              $excel->getActiveSheet()->getDefaultColumnDimension()->setAutoSize(true);
        */
        // 报表头列名的输出
        $excel->getActiveSheet()->fromArray($headData , NULL , 'A1');
        // 具体数据的输出
        $excel->getActiveSheet()->fromArray($bodyData , NULL , 'A2');
        $fileStr = $fileName ? $fileName : '未命名';
        ob_end_clean();
        ob_start();
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename=' . $fileStr . '.xls');
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($excel , 'Excel5');  //注意，要加“\”，否则会报错
        $objWriter->save('php://output');
    }

    /**
     * @param  UploadedFile $fileObj
     * @param string        $savePath
     * @param bool          $save
     * @return array|bool
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */

    public static function readExcel($fileObj , $savePath = '/' , $save = true)
    {
        if ($save) {
            /*以时间来命名上传的文件*/
            $str = date('Ymdhis');
            $file_name = $str . "." . $fileObj->extension;
            if (!is_dir($savePath)) {
                mkdir($savePath , 0777 , true);
            }
            if (!$fileObj->saveAs($savePath . $file_name)) {
                return false;
            }
            $file = $savePath . $file_name;
        } else {
            $file = $fileObj;
        }
        $objPHPExcel = PHPExcel_IOFactory::load($file);
        $sheetData = $objPHPExcel->getActiveSheet()->toArray(null , true , true , true);
        //删除文件
        @unlink($file);
        return $sheetData;
    }

    /**
     * @throws Exception
     */
    private function initModels()
    {
        if ($this->all) {
            $this->dataProvider->pagination = false; //不使用分页，导出全部数据
        }
        $this->allModels = $this->dataProvider->getModels();
        if (!$this->allModels) {
            throw new Exception('没有数据无法导出');
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
                                $copyOrderData->populateRelation($value , [$v]);
                                $this->models[] = $copyOrderData;
                            }
                        }
                    }
                }
            }
            unset($item , $rModel , $copyOrderData , $v , $value);
            if ($this->models[0]->primaryKey) {
                //排序
                $pk = [];
                foreach ($this->models as $item) {
                    $pk[] = $item->primaryKey;
                }
                array_multisort($this->models , $pk);
            }
        }
    }

    /**
     * @throws Exception
     */
    private function initHeaderData()
    {
        $res = [];
        if ($this->fields) {
            $model = $this->models[0];
            foreach ($this->fields as $item) {
                $label = $model->getAttributeLabel($item);
                if (!preg_match("/[\x7f-\xff]/" , $label) && $this->relation) {
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
    private function initBodyData()
    {
        $res = [];
        /** @var ActiveRecord $item */
        foreach ($this->models as $key => $item) {

            if ($this->show) {
                if (!method_exists($item , 'show')) {
                    throw new Exception(get_class($item) . '中未定义show()方法,不能进行值的转换');
                }
                $item->show();
            }
            foreach ($this->fields as $attribute) {
                $res[$key][$attribute] = $this->getValue($item , $attribute);
            }

        }
        $this->bodyDataArray = $res;
    }

    /**
     * 检查要导出的列是否存在
     * @throws Exception
     */
    private function checkFields()
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
                if(!$rModel){
                    continue;
                }
                if (is_object($rModel)) {
                    $relationFields = array_keys($rModel->getAttributes());
                } else {
                    $relationFields = array_keys($item[$value][0]->getAttributes());
                }
                $modelFields = array_merge($modelFields , $relationFields);
                unset($relationFields);
            }
        }
        $fields = $this->fields;
        $temp = array_intersect($modelFields , $fields);
        if (sort($temp) != sort($fields)) {
            throw new Exception('导出的字段在数据中并不存在，请检查');
        }
    }

    /**
     * 取得相应的列值
     * @param ActiveRecord $model
     * @param string       $attribute
     * @return mixed|null
     */
    private function getValue($model , $attribute)
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

    /**
     * @throws Exception
     */
    public function toExcel()
    {
        if ($this->models) {
            $this->modelsToExcel();
        }
    }

    /**
     * @throws Exception
     */
    private function modelsToExcel()
    {
        $this->checkFields();
        $this->getRelationModels();
        $this->initHeaderData();
        $this->initBodyData();
        self::arrayToExcel($this->headerDataArray , $this->bodyDataArray , $this->fileName);
    }

}