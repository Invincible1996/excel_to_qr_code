import bean.ClassBean;
import util.ExcelUtil;

import java.awt.image.BufferedImage;
import java.util.List;

public class QRCode {
    public static void main(String[] args) throws Exception {
        // 读取Excel文件
        List<ClassBean> beanList = ExcelUtil.readFromXLSX2007("/Users/kevin/Downloads/H_Room.xlsx");
        int count = 0;
        for (ClassBean classBean : beanList) {
            if (classBean.getAddress().contains("上海")) {
                count++;
                StringBuilder sb = new StringBuilder();
                sb.append("/Users/kevin/qrcode2/").append(classBean.getCode()).append(classBean.getClassName()).append(".png");
                BufferedImage image = QRCodeUtil.toBufferedImage(classBean.getCode(), 750, 750);
                QRCodeUtil.markImageByCode(image, "/Users/kevin/Downloads/bg.png", sb.toString(), 212, 587);
                QRCodeUtil.pressText(classBean.getClassName(), sb.toString(), sb.toString(), 238, 1600, 1.0f);
            }
        }
        System.out.println("生成" + count + "张二维码成功");
    }
}
