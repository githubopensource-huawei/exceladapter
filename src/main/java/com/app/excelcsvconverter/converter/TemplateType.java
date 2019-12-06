package com.app.excelcsvconverter.converter;

/**
 * 功能描述：
 *
 * @since 2019-09-23
 */
public enum TemplateType {
    /**
     * Summary表格
     * 模板Cover标题和版本信息动态刷新
     */
    SUMMARY() {
        @Override
        public boolean isSummary() {
            return true;
        }
        @Override
        public boolean isNegotiated() {
            return false;
        }

        @Override
        public boolean isControllerSummary() {
            return false;
        }
    },

    /**
     * 控制器表格
     * 需要删除模板cover之外所有页签才能写入数据
     * 保留模板Cover页签
     */
    NEGOTIATED() {
        @Override
        public boolean isSummary() {
            return false;
        }

        @Override
        public boolean isNegotiated() {
            return true;
        }

        @Override
        public boolean isControllerSummary() {
            return false;
        }

    },

    /**
     * ControllerSummary表格
     * 模板中Cover页签包含版本信息，位置固定（第二行第三列）
     * 用于取模板路径为{0}\cm\config\{1}\Model\summarytemplates\GSM\的表格
     * 处理逻辑同NEGOTIATED类型
     */
    CONTROLLERSUMMARY() {
        @Override
        public boolean isSummary() {
            return false;
        }
        @Override
        public boolean isNegotiated() {
            return false;
        }

        @Override
        public boolean isControllerSummary() {
            return true;
        }
    },

    /**
     * FileType没有类型填写
     */
    BLANK(){
        @Override
        public boolean isSummary() {
            return false;
        }
        @Override
        public boolean isNegotiated() {
            return false;
        }

        @Override
        public boolean isControllerSummary() {
            return false;
        }
    };

    TemplateType() {
    }

    public boolean isSummary() {
        return false;
    }

    public boolean isNegotiated() {
        return false;
    }

    public boolean isControllerSummary() {
        return false;
    }

    public static TemplateType of(String value) {
        if ("Summary".equalsIgnoreCase(value)) {
            return SUMMARY;
        } else if ("Negotiated".equalsIgnoreCase(value)) {
            return NEGOTIATED;
        } else if ("ControllerSummary".equalsIgnoreCase(value)) {
            return CONTROLLERSUMMARY;
        } else {
            return BLANK;
        }
    }
}
