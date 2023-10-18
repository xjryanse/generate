<?php
namespace xjryanse\generate\model;

/**
 * 生成文档模板
 */
class GenerateTemplateLog extends Base
{
    /**
     * 原始word文档
     * @param type $value
     * @return type
     */
    public function getFileRawAttr($value) {
        return self::getImgVal($value);
    }

    /**
     * @param type $value
     * @throws \Exception
     */
    public function setFileRawAttr($value) {
        return self::setImgVal($value);
    }
    
    public function getFilePdfAttr($value) {
        return self::getImgVal($value);
    }

    /**
     * @param type $value
     * @throws \Exception
     */
    public function setFilePdfAttr($value) {
        return self::setImgVal($value);
    }
    
    /**
     * @param type $value
     * @throws \Exception
     */
    public function setFileImageAttr($value) {
        return self::setImgVal($value);
    }
    /**
     * 多张
     * @param type $value
     * @return type
     */
    public function getFileImageAttr($value) {
        return self::getImgVal($value,true);
    }
}