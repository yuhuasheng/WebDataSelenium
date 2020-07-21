package com.dayang.util;

import com.dayang.domain.ProteinsInfo;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * describe:
 * POI工具包
 *
 * @author 230257
 * @date 2020/07/20
 */
public class Utils {


    /**
     * 添加表头列
     *
     * @param sheet       Sheet表对象
     * @param columnStyle 列样式对象
     */
    public static void addTitleData(HSSFSheet sheet, HSSFCellStyle columnStyle) {
        String[] title = {"序号", "Cat", "Product Name", "Product Overview", "Description", "Source/Host", "Species", "Applications", "Tag",
                "Form", "Activity", "Formulation", "Molecular Mass", "Molecular Weight", "Purity", "Concentration", "Endotoxin",
                "Predicted N-terminus", "Bio-activity", "Unit Definition", "AA Sequence", "Protein length", "Storage",
                "Reconstitution", "Storage buffer", "Tissue specificity", "Shipping Condition", "Stability", "Usage",
                "Quality Control Test", "Preservative", "Sequence Similarities", "Gene Name", "Official Symbol", "Synonyms",
                "Gene ID", "mRNA Refseq", "Protein Refseq", "MIM", "UniProt ID", "Chromosome Location", "Function", "url"};
        HSSFCell cell = null;
        HSSFRow row = sheet.createRow(0);
        row.setHeightInPoints(50);
        for (int j = 0; j < title.length; j++) {
            cell = row.createCell(j);
            cell.setCellValue(new HSSFRichTextString(title[j]));
            cell.setCellStyle(columnStyle);
        }
    }


    /**
     * 设置每一列的宽度
     *
     * @param sheet Excel表对象
     */
    public static void setColumnWidth(HSSFSheet sheet) {
        for (int i = 0; i < 43; i++) {
            if (i == 0) {
                sheet.setColumnWidth(i, 3000);
            } else if (i == 1) {
                sheet.setColumnWidth(i, 10000);
            } else if (i == 3 || i == 4) {
                sheet.setColumnWidth(i, 35000);
            } else if (i == 9) {
                sheet.setColumnWidth(i, 25000);
            } else if (i == 11 || i == 14 || i == 15) {
                sheet.setColumnWidth(i, 8000);
            } else if (i == 42) {
                sheet.setColumnWidth(i, 15000);
            } else {
                sheet.setColumnWidth(i, 4500);
            }
        }
    }


    /**
     * 设置列样式
     *
     * @param workbook Excel表对象
     * @return 返回列样式对象
     */
    @SuppressWarnings("resource")
    public static HSSFCellStyle setColumnStyle(HSSFWorkbook workbook) {
        HSSFCellStyle cellStyle1 = workbook.createCellStyle();
        cellStyle1.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        cellStyle1.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle1.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle1.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        cellStyle1.setWrapText(true);
        return cellStyle1;
    }


    /**
     * 设置内容的样式
     *
     * @param workbook Excel表对象
     * @return 返回内容样式对象
     */
    public static HSSFCellStyle setContentStyle(HSSFWorkbook workbook) {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Times New Roman");
        cellStyle.setFont(font);
        cellStyle.setWrapText(true);
        return cellStyle;
    }


    /**
     * 冻结表头
     *
     * @param sheet 表对象
     */
    public static void freezeHeader(HSSFSheet sheet) {
        // 冻结第一行
        sheet.createFreezePane(0, 1, 0, 1);
    }


    public static void createLine(HSSFSheet sheet, HSSFCellStyle cellStyle, int number, ProteinsInfo proteinsInfo) {
        HSSFCell cell = null;
        HSSFRow row = sheet.createRow(number);
        row.setHeightInPoints(40);

        cell = row.createCell(0);
        cell.setCellValue(new HSSFRichTextString(String.valueOf(number)));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getCat()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getProductName()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(3);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getProductOverview()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(4);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getDescription()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(5);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getSourceHost()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(6);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getSpecies()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(7);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getApplications()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(8);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getTag()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(9);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getForm()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(10);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getActivity()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(11);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getFormulation()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(12);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getMolecularMass()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(13);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getMolecularWeight()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(14);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getPurity()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(15);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getConcentration()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(16);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getEndotoxin()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(17);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getPredictedNTerminus()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(18);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getBioActivity()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(19);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getUnitDefinition()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(20);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getAaSequence()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(21);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getProteinLength()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(22);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getStorage()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(23);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getReconstitution()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(24);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getStorageBuffer()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(25);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getTissueSpecificity()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(26);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getShippingCondition()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(27);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getStability()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(28);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getUsage()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(29);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getQualityControlTest()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(30);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getPreservative()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(31);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getSequenceSimilarities()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(32);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getGeneName()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(33);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getOfficialSymbol()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(34);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getSynonyms()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(35);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getGeneId()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(36);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getmRNARefseq()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(37);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getProteinRefseq()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(38);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getMIM()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(39);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getUniProtId()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(40);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getChromosomeLocation()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(41);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getFunction()));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(42);
        cell.setCellValue(new HSSFRichTextString(proteinsInfo.getUrl()));
        cell.setCellStyle(cellStyle);
    }
}
