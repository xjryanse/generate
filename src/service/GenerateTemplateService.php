<?php

namespace xjryanse\generate\service;

use xjryanse\system\interfaces\MainModelInterface;
use xjryanse\system\service\SystemFileService;
use PhpOffice\PhpWord\PhpWord;
use xjryanse\logic\Arrays;
use xjryanse\logic\Debug;
use Exception;

/**
 * 生成文档模板
 */
class GenerateTemplateService extends Base implements MainModelInterface {

    use \xjryanse\traits\InstTrait;
    use \xjryanse\traits\MainModelTrait;
    use \xjryanse\traits\MainModelQueryTrait;

    protected static $mainModel;
    protected static $mainModelClass = '\\xjryanse\\generate\\model\\GenerateTemplate';

    /**
     * key  转id
     * @param type $key
     * @return type
     */
    public static function keyToId($key) {
        $con[] = ['file_key', '=', $key];
        $ids = self::ids($con);
        return $ids ? $ids[0] : '';
    }

    /*
     * 文档基本路径
     */

    protected static $docBasePath = "/Uploads/Download/CanDelete/";

    /**
     * 
     */
    public function generate($data = []) {
        $path = $this->get();
        Debug::debug('GenerateTemplateLogService 的 export 的 info', $path);
        // 20230211:当template_id所指文件不存在，会返回字符串，借此判断
        if (!$path || is_string($path['template_id'])) {
            throw new Exception('模板文件' . $this->uuid . '不存在');
        }

        $PHPWord = new PhpWord();
        $templateProcessor = $PHPWord->loadTemplate('./' . $path['template_id']['rawPath']);
        // 20230406:尝试加载远程
        // $templateProcessor = $PHPWord->loadTemplate($path['template_id']['file_path']);
        //获取模板字段
        $fields = GenerateTemplateFieldService::getTemplateFields($this->uuid);
        foreach ($fields as $key => $value) {
            //文本
            if ($value['field_type'] == 'text') {
                $tData = isset($data[$value['field_name']]) ? $data[$value['field_name']] : '';
                $templateProcessor->setValue($value['field_name'], $tData);
            }
            //签名
            if ($value['field_type'] == 'sign') {
                $tData = isset($data[$value['field_name']]) ? $data[$value['field_name']] : '';
                if ($tData) {
                    $templateProcessor->setImageValue($value['field_name'], ['path' => $tData, "width" => 200, "height" => 50]);
                } else {
                    $templateProcessor->setValue($value['field_name'], '');
                }
            }
            //2023-02-22图片
            if ($value['field_type'] == 'image') {
                $tData = isset($data[$value['field_name']]) ? $data[$value['field_name']] : '';
                if ($tData) {
                    $templateProcessor->setImageValue($value['field_name'], ['path' => $tData['rawPath'], "width" => $value['width'], "height" => $value['height']]);
                    // $templateProcessor->setImageValue($value['field_name'], ['path' => $tData['file_path'], "width" => 500, "height" => 800]);
                } else {
                    $templateProcessor->setValue($value['field_name'], '');
                }
            }
        }
        $fileName = self::getFileName($data);
        //基本路径
        $basePath = self::$docBasePath . $fileName;
        $path = $basePath . '.doc';
//               //保存文件
        $templateProcessor->saveAs('.' . $path);
        //20210309 生成word
        $saveData['file_path'] = str_replace("/Uploads", "Uploads", $path);
        $res = SystemFileService::save($saveData);
        return $res;
    }

    protected static function getFileName($data) {
        sort($data);
        $str = '';
        foreach ($data as $key => $value) {
            $str .= $value;
        }
        return md5($str) . sha1($str);
    }

    /**
     * 钩子-保存前
     */
    public static function extraPreSave(&$data, $uuid) {
        
    }

    /**
     * 钩子-保存后
     */
    public static function extraAfterSave(&$data, $uuid) {
        
    }

    /**
     * 钩子-更新前
     */
    public static function extraPreUpdate(&$data, $uuid) {
        
    }

    /**
     * 钩子-更新后
     */
    public static function extraAfterUpdate(&$data, $uuid) {
        
    }

    /**
     * 钩子-删除前
     */
    public function extraPreDelete() {
        
    }

    /**
     * 钩子-删除后
     */
    public function extraAfterDelete() {
        
    }

    /**
     * 20220907：测试
     * @param type $ids
     * @return type
     */
    public static function extraDetails($ids) {
        return self::commExtraDetails($ids, function($lists) use ($ids) {
                    $cond[] = ['template_id', 'in', $ids];
                    $templateFields = GenerateTemplateFieldService::mainModel()->where($cond)->group('template_id')->column('count( DISTINCT id ) AS number', 'template_id');
                    $templateLogs = GenerateTemplateLogService::mainModel()->where($cond)->group('template_id')->column('count( DISTINCT id ) AS number', 'template_id');

                    foreach ($lists as &$v) {
                        $v['templateFieldsCount'] = Arrays::value($templateFields, $v['id'], 0);
                        $v['templateLogsCount'] = Arrays::value($templateLogs, $v['id'], 0);
                    }
                    return $lists;
                });
    }

    /**
     *
     */
    public function fId() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     *
     */
    public function fCompanyId() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 培训id
     */
    public function fSafeId() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 驾驶员id
     */
    public function fDriverId() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 排序
     */
    public function fSort() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 状态(0禁用,1启用)
     */
    public function fStatus() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 有使用(0否,1是)
     */
    public function fHasUsed() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 锁定（0：未锁，1：已锁）
     */
    public function fIsLock() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 锁定（0：未删，1：已删）
     */
    public function fIsDelete() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 备注
     */
    public function fRemark() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 创建者，user表
     */
    public function fCreater() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 更新者，user表
     */
    public function fUpdater() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 创建时间
     */
    public function fCreateTime() {
        return $this->getFFieldValue(__FUNCTION__);
    }

    /**
     * 更新时间
     */
    public function fUpdateTime() {
        return $this->getFFieldValue(__FUNCTION__);
    }

}
