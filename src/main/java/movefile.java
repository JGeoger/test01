import java.io.*;
import java.nio.charset.Charset;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class movefile {
    static Date date = new Date();
    static SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
    public static String tmpdir = sdf.format(date) + Long.toString(date.getTime());
    @SuppressWarnings("rawtypes")
    public static void unZipFiles(File zipFile, String descDir) throws IOException {

        ZipFile zip = new ZipFile(zipFile, Charset.forName("GBK"));//解决中文文件夹乱码
        String name = zip.getName().substring(zip.getName().lastIndexOf('\\')+1, zip.getName().lastIndexOf('.'));
        //Date date = new Date();
        //SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        //String tmpdir = sdf.format(date) + Long.toString(date.getTime());
        System.out.println(tmpdir);
        File pathFile = new File(descDir+ tmpdir);
        //判断路径是否存在
        if (!pathFile.exists()) {
            pathFile.mkdirs();
        }
        for (Enumeration<? extends ZipEntry> entries = zip.entries(); entries.hasMoreElements();) {
            ZipEntry entry = (ZipEntry) entries.nextElement();
            String zipEntryName = entry.getName();
            InputStream in = zip.getInputStream(entry);
            String outPath = (descDir+ tmpdir +"/"+ zipEntryName).replaceAll("\\*", "/");

            // 判断路径是否存在,不存在则创建文件路径
            File file = new File(outPath.substring(0, outPath.lastIndexOf('/')));
            if (!file.exists()) {
                file.mkdirs();
            }
            // 判断文件全路径是否为文件夹,如果是上面已经上传,不需要解压
            if (new File(outPath).isDirectory()) {
                continue;
            }
            // 输出文件路径信息
            System.out.println(outPath);

            FileOutputStream out = new FileOutputStream(outPath);
            byte[] buf1 = new byte[1024];
            int len;
            while ((len = in.read(buf1)) > 0) {
                out.write(buf1, 0, len);
            }
            in.close();
            out.close();
        }
        System.out.println("******************解压完毕********************");
        return;
    }

    static void copy(String srcPathStr, String desPathStr)
    {
        //获取源文件的名称
        String newFileName = srcPathStr.substring(srcPathStr.lastIndexOf("\\")+1); //目标文件地址
        System.out.println("源文件:"+newFileName);
        desPathStr = desPathStr + File.separator + newFileName; //源文件地址
        System.out.println("目标文件地址:"+desPathStr);
        try
        {
            FileInputStream fis = new FileInputStream(srcPathStr);//创建输入流对象
            FileOutputStream fos = new FileOutputStream(desPathStr); //创建输出流对象
            byte datas[] = new byte[1024*8];//创建搬运工具
            int len = 0;//创建长度
            while((len = fis.read(datas))!=-1)//循环读取数据
            {
                fos.write(datas,0,len);
            }
            fis.close();//释放资源
            fis.close();//释放资源
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        movefile.unZipFiles(new File("D:\\test\\test1.zip"), "D:\\result1\\");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream("D:\\test\\testexcel.xlsx"));
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        //获取最后一行的num，即总行数。此处从0开始
        int maxRow = sheet.getLastRowNum();
        for (int row = 0; row <= maxRow; row++) {
            System.out.println(sheet.getRow(row).getCell(0));
            String srcPathStr = "D:\\result1\\" + tmpdir +String.valueOf(sheet.getRow(row).getCell(1));; //源文件地址
            String desPathStr = String.valueOf(sheet.getRow(row).getCell(0)); //目标文件地址
            copy(srcPathStr, desPathStr);
        }
        System.out.println();
    }
}
