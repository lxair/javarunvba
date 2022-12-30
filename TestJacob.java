package hhhhhh;

/*
 * Copyright © 2019 bjfansr@cn.ibm.com Inc. All rights reserved
 * @package: com.ibm.jacob
 * @version V1.0
 */

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.junit.Test;

/**
 * @author Moses *
 * @Date 2019/4/1
 */
public class TestJacob {
    @Test
    public void testMacro() {
        JacobExcelTool tool = new JacobExcelTool();
        //打开
        tool.OpenExcel("E:/新desktop/python1.xlsm", false, false);
        Dispatch sheet = tool.getSheetByName("Sheet1");
//        for (int i = 2; i <= 7; i++) {
//            tool.setValue(sheet, "B" + i, "value", i * 1.2);
//        }
        //调用Excel宏
        tool.callMacro("test1");
        //调用Excel宏
        //tool.callMacro("样式设置");
//        Variant num = tool.getValue("A10", sheet);
//        System.out.println(num);
        //关闭并保存，释放对象
        tool.CloseExcel(true, true);
    }
}
