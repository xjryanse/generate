<?php
namespace xjryanse\generate\model;

/**
 * 生成文档模板
 */
class GenerateTemplate extends Base
{
    public static $picFields = ['template_id'];

    public function getTemplateIdAttr($value) {
        return self::getImgVal($value);
    }

    /**
     * @param type $value
     * @throws \Exception
     */
    public function setTemplateIdAttr($value) {
        return self::setImgVal($value);
    }

}