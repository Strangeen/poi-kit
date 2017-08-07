package online.dinghuiye.common;

/**
 * Created by Strangeen on 2017/7/10.
 *
 * 导出excel的写入模式，I - 插入，C - 覆盖
 */
public enum WriteMode {

    I("insert"), C("cover");

    private String modeName;

    WriteMode(String modeName) {
        this.modeName = modeName;
    }

    public String getModeName() {
        return modeName;
    }
}
