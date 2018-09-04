package com.yeahpi.model;

/**
 * Class
 *
 * @author wgm
 * @date 2018/09/04
 */
public enum TestEnum {
    CREATE(1, "创建"),
    UPDATE(2, "更新")
    ;
    private int code;
    private String desc;

    TestEnum(int code, String desc){
        this.code = code;
        this.desc = desc;
    }

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    public static String getDescByCode(int code){
        for(TestEnum testEnum : TestEnum.values()){
            if(testEnum.code == code){
                return testEnum.desc;
            }
        }
        return "";
    }
}
