<?php

namespace xjryanse\generate\service;

use xjryanse\system\interfaces\MainModelInterface;
use xjryanse\curl\Query;
use xjryanse\system\service\SystemFileService;
use xjryanse\system\logic\FileLogic;
use xjryanse\system\logic\ExportLogic;
use xjryanse\system\service\columnlist\Dynenum;
use xjryanse\logic\Pdf;
use xjryanse\logic\Debug;
use xjryanse\logic\Arrays;
use think\facade\Request;
use Exception;

/**
 * 生成文档记录
 */
class GenerateTemplateLogService extends Base implements MainModelInterface {

    use \xjryanse\traits\InstTrait;
    use \xjryanse\traits\MainModelTrait;
    use \xjryanse\traits\MainModelRamTrait;
    use \xjryanse\traits\MainModelCacheTrait;
    use \xjryanse\traits\MainModelCheckTrait;
    use \xjryanse\traits\MainModelGroupTrait;
    use \xjryanse\traits\MainModelQueryTrait;


    protected static $mainModel;
    protected static $mainModelClass = '\\xjryanse\\generate\\model\\GenerateTemplateLog';

    /**
     * pdf转url
     * @var type 
     */
    protected static $pdfCovUrl = "http://office.xiesemi.cn/index.php/word/toPdf";

    /**
     * word文档，生成pdf
     * @param type $templateId
     * @param type $data
     * @return type
     */
    public static function generate($templateId, $data) {
        $identKey = self::getFileName($data);
        $con[] = ['ident_key', '=', $identKey];
        $con[] = ['template_id', '=', $templateId];
        $info = self::mainModel()->where($con)->find();
        if ($info) {
            return $info;
        } else {
            $res = GenerateTemplateService::getInstance($templateId)->generate($data);
            Debug::debug('$res', $res);
            $saveData['template_id'] = $templateId;
            $saveData['data'] = json_encode($data, JSON_UNESCAPED_UNICODE);
            $saveData['ident_key'] = $identKey;
            $saveData['file_raw'] = $res['id'];
            // 20230519：文件路径替换
            $urlFull = SystemFileService::mainModel()->getFullPath($res['file_path']);
            // 20231006调试
//            if(Debug::isDevIp()){
//                dump($res['file_path']);
//
//                $phpWord = new \PhpOffice\PhpWord\PhpWord();
//                // 加载 Word 文档
//                $phpWord->loadTemplate($res['file_path']);
//
//                // 创建 PDF 导出器
//                $pdfWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'PDF');
//
//                // 导出为 PDF 文件
//                $pdfWriter->save('./test20231006.pdf');
//                exit;
//                
//            }
            
            //word 转为pdf 
            $url = urlencode($urlFull);
            $pdfUrl = self::$pdfCovUrl . '?filePath=' . $url;
            Debug::debug('$pdfUrl', $pdfUrl);
            $pdf = Query::geturl($pdfUrl);
            Debug::debug('$pdf', $pdf);
            // 保存pdf文件到本地服务器
            $pdfFile = FileLogic::saveUrlFile($pdf['data']);

            $saveData['file_pdf'] = $pdfFile['id'];
            // pdf 转化成图片
            $pngs = Pdf::toPngMany($pdfFile['file_path']);
            $pngIds = [];
            foreach ($pngs as &$png) {
                $pngIds[] = SystemFileService::pathSaveGetId($png);
            }
            $saveData['file_image'] = implode(',', $pngIds);

            $log = self::save($saveData);

            return $log;
        }
    }

    /**
     * 自动识别，导出csv还是excel
     */
    public static function exportAuto($templateId, &$dataArr, $limit = 400) {
        if (count($dataArr) > $limit) {
            //500条以上用csv性能比较好
            $resp = self::exportCsv($templateId, $dataArr, []);
            $resp['url'] = $resp['file_path'];
        } else {
            //500条以下用这个可以直接处理excel
            $res = GenerateTemplateLogService::export($templateId, $dataArr, []);
            $resp['fileName'] = time() . '.xlsx';
            $resp['url'] = $res['file_path'];
        }
        return $resp;
    }

    /**
     * excel 数据导出
     * @param type $templateId      模板id
     * @param type $data            二维数组导出
     * @param type $replace        文档需要替换的信息
     * @param type $sumFields       求和字段（一维）
     * @return type
     */
    public static function export($templateId, $data, $replace = []) {
        $path = GenerateTemplateService::getInstance($templateId)->get();
        Debug::debug('GenerateTemplateLogService 的 export 的 info', $path);
        // 20230211:当template_id所指文件不存在，会返回字符串，借此判断
        if (!$path || is_string($path['template_id'])) {
            throw new Exception('模板文件' . $templateId . '不存在');
        }
        $dataImport = self::dataCov($templateId, $data, $path['with_title']);
        $templatePath = './' . $path['template_id']['rawPath'];
        // 20230211:从外部搬到内部
        if (Debug::isDebug()) {
            exit;
        }
        $startRow           = Arrays::value($path, 'excel_start_row', 0);

        $rr                 = ExportLogic::writeToExcel($dataImport, $startRow, $templatePath, '', $replace);
        $filePath           = str_replace('./', '/', $rr);
        // http会被谷歌浏览器拦截无法下载
        $res['file_path']   = Request::domain() . $filePath;
        $res['pathRaw']     = $rr;

        return $res;
    }

    /**
     * 20220823：导出csv，适用于数据量大的场景
     * @param type $templateId
     * @param type $data
     * @param type $replace
     * @return type
     */
    public static function exportCsv($templateId, $data) {
        // $path           = GenerateTemplateService::getInstance($templateId)->get();
        // csv一定要带标题
        $dataArr = self::dataCov($templateId, $data, false, true);
        // $templatePath   = './' . $path['template_id']['rawPath'];

        $fileName = ExportLogic::getInstance()->putIntoCsv($dataArr);
        $res['file_path'] = Request::domain() . '/Uploads/Download/CanDelete/' . $fileName;
        $res['fileName'] = $fileName;
        return $res;
    }

    /*
     * 数据转换
     */

    protected static function dataCov($templateId, $data, $withTitle = false, $withLabel = false) {
        // 20240406加
        $dynArrs        = GenerateTemplateFieldService::dynArrs($templateId); 
        // dump($dynArrs);exit;
        $dynDataList    = Dynenum::dynDataList($data, $dynArrs);

        foreach($dynArrs as $k=>$v){
            foreach($data as &$it){
                $it[$k] = Arrays::value($dynDataList[$k], $it[$k]) ?: $it[$k];
            }
        }
        
        $fields     = GenerateTemplateFieldService::getTemplateFields($templateId);
        $sumFields = GenerateTemplateFieldService::sumFields($templateId);

        $dataTitle = array_column($fields, 'field_name', 'label');

        $dataImport = [];
        /**
         * 导出带标签
         */
        if ($withLabel) {
            $dataImport[] = array_column($fields, 'label');
        }
        /**
         * 导出带字段名
         */
        if ($withTitle) {
            $dataImport[] = $dataTitle;
        }
        foreach ($data as $k=>&$v) {
            // 20240301
            $v['i'] = $k + 1;
            $tmp = [];
            // 20231229:序号专用
            // $tmp['i'] = $k + 1;
            
            foreach ($dataTitle as $k => &$title) {
                $tmp[$k] = Arrays::value($v, $title, '');
            }
            $dataImport[] = $tmp;
        }
        //20220815:求和字段
        if ($sumFields) {
            $sumArr = [];
            foreach ($dataTitle as $index => &$title) {
                if (!$sumArr) {
                    // 第一栏填合计
                    $sumArr[] = '合计';
                } else {
                    $sumArr[] = in_array($title, $sumFields) ? array_sum(array_column($dataImport, $index)) : '';
                }
            }
            $dataImport[] = $sumArr;
        }

        return $dataImport;
    }

    protected static function getFileName($data) {
        if (!$data) {
            return md5(time());
        }
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
