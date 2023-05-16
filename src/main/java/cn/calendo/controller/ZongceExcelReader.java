package cn.calendo.controller;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.HashMap;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * @author Calendo
 * @version 1.0
 * @description zongce
 * @date 2023/5/15 17:41
 */

@CrossOrigin
@RestController
@RequestMapping("/zongce")
public class ZongceExcelReader {

    @PostMapping("/upload")
    public String upload(MultipartFile file, Integer majorStart) throws Exception {

        InputStream inputStream = file.getInputStream();
        ExcelReader excelReader = ExcelUtil.getReader(inputStream, 0);
        //阅读第一行，把学分信息记录下来
        HashMap<Integer, Double> creditMap = new HashMap<>();//index credit

        int i1 = majorStart;
        Pattern totalp = Pattern.compile("\\d+\\.\\d+"); // 匹配一个或多个数字，后面跟一个点，再跟一个或多个数字
        while (!"".equals(String.valueOf(excelReader.readCellValue(i1, 0)).trim()) && totalp.matcher(String.valueOf(excelReader.readCellValue(i1, 0))).find()) {
            String rawString = String.valueOf(excelReader.readCellValue(i1, 0));
            Pattern p = Pattern.compile("\\d+\\.\\d+"); // 匹配一个或多个数字，后面跟一个点，再跟一个或多个数字
            Matcher m = p.matcher(rawString);
            if (m.find()) { // 如果找到匹配
                Double matchString = Double.valueOf(m.group()); // 获取匹配的子字符串的数值
                System.out.println(matchString); // 输出0.3
                creditMap.put(i1, matchString);
            }
            i1++;
        }
        //for (int i = majorStart; String.valueOf(excelReader.readCellValue(i, 0)) != null && !"".equals(String.valueOf(excelReader.readCellValue(i, 0))); i++) {//第几列开始，啥时候结束
        //    String rawString = String.valueOf(excelReader.readCellValue(i, 0));
        //    Pattern p = Pattern.compile("\\d+\\.\\d+"); // 匹配一个或多个数字，后面跟一个点，再跟一个或多个数字
        //    Matcher m = p.matcher(rawString);
        //    if (m.find()) { // 如果找到匹配
        //        Double matchString = Double.valueOf(m.group()); // 获取匹配的子字符串的数值
        //        //System.out.println(matchString); // 输出0.3
        //        creditMap.put(i, matchString);
        //    }
        //}
        System.out.println("creditMap: " + creditMap);

        i1--;
        int endMajor = i1;

        TreeMap<Integer, Double> resMap = new TreeMap<>();
        HashMap<Integer, String> resultMap = new HashMap<>();
        //开始读行

        //新
        int i2 = 1;
        while (!"".equals(excelReader.readCellValue(0, i2)) && excelReader.readCellValue(0, i2) != null) {
            System.out.println("row: " + i2);
            Double personalSumCreditMultiScore = 0.0;//记录每行学生的加权成绩
            Double creditMapSumCredit = 0.0;//creditMap Sum Credit

            for (int j1 = majorStart; j1 < endMajor; j1++) {//开始每行学生的成绩
                System.out.println("col: " + j1);
                String value = String.valueOf(excelReader.readCellValue(j1, i2));//杂分
                if ("合格".equals(value.trim())) {
                    value = "70";
                }
                if ("不合格".equals(value.trim())) {
                    value = "50";
                }
                if ("优秀".equals(value.trim())) {
                    value = "95";
                }
                if ("良好".equals(value.trim())) {
                    value = "85";
                }
                if ("中等".equals(value.trim())) {
                    value = "75";
                }
                if ("及格".equals(value.trim())) {
                    value = "65";
                }
                if ("不及格".equals(value.trim())) {
                    value = "50";
                }
                if ("".equals(value.trim()) || value.trim() == null || " ".equals(value.trim()) || "  ".equals(value.trim())) {
                    value = "-1";
                }
                Double matchValue = Double.valueOf(value);//处理后的成绩，有就是正整数，无就是-1
                if (matchValue != -1) {
                    creditMapSumCredit += creditMap.get(j1);//渐进式加入当前有效学科
                    matchValue = matchValue / 10 - 5;
                    //System.out.print("matchValue: " + matchValue);
                    personalSumCreditMultiScore += matchValue * creditMap.get(j1);//乘加
                }
            }
            //System.out.println("totalMajor: " + totalMajor);
            personalSumCreditMultiScore /= creditMapSumCredit;
            resMap.put(i2, personalSumCreditMultiScore);
            //String result = String.valueOf(personalSumCreditMultiScore);
            String result = "id: " + excelReader.readCellValue(0, i2) + " name: " + excelReader.readCellValue(1, i2) + " result: " + personalSumCreditMultiScore;
            resultMap.put(i2, result);
            i2++;
        }
        //旧
        //for (int i = 1; "".equals(excelReader.readCellValue(0, i)) && excelReader.readCellValue(0, i) == null; i++) {//开始读行
        //    Double personalSumCreditMultiScore = 0.0;//记录每行学生的加权成绩
        //    Double creditMapSumCredit = 0.0;//creditMap Sum Credit
        //    for (int j = majorStart; j < majorEnd; j++) {//开始每行学生的成绩
        //        String value = String.valueOf(excelReader.readCellValue(j, i));//杂分
        //        if ("合格".equals(value.trim())) {
        //            value = "70";
        //        }
        //        if ("不合格".equals(value.trim())) {
        //            value = "50";
        //        }
        //        if ("优秀".equals(value.trim())) {
        //            value = "95";
        //        }
        //        if ("良好".equals(value.trim())) {
        //            value = "85";
        //        }
        //        if ("中等".equals(value.trim())) {
        //            value = "75";
        //        }
        //        if ("及格".equals(value.trim())) {
        //            value = "65";
        //        }
        //        if ("不及格".equals(value.trim())) {
        //            value = "50";
        //        }
        //        if ("".equals(value.trim()) || value.trim() == null || " ".equals(value.trim()) || "  ".equals(value.trim())) {
        //            value = "-1";
        //        }
        //        Double matchValue = Double.valueOf(value);//处理后的成绩，有就是正整数，无就是-1
        //        if (matchValue != -1) {
        //            creditMapSumCredit += creditMap.get(j);//渐进式加入当前有效学科
        //            matchValue = matchValue / 10 - 5;
        //            //System.out.print("matchValue: " + matchValue);
        //            personalSumCreditMultiScore += matchValue * creditMap.get(j);//乘加
        //        }
        //    }
        //    //System.out.println("totalMajor: " + totalMajor);
        //    personalSumCreditMultiScore /= creditMapSumCredit;
        //    resMap.put(i, personalSumCreditMultiScore);
        //    String result = String.valueOf(personalSumCreditMultiScore);
        //    //String result = "id: " + excelReader.readCellValue(0, i) + " name: " + excelReader.readCellValue(1, i) + " result: " + personalSumCreditMultiScore;
        //    resultMap.put(i, result);
        //}
        return String.valueOf(resultMap);
    }

}
