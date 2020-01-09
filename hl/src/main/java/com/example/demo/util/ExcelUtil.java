package com.example.demo.util;

public class ExcelUtil {
    /**
     * 读取excel数据
     *
     * @param path
     */
    private void readExcelToObj(String path) {

        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(new File(path));
            readExcel(wb, 0, 0, 0);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取excel文件
     *
     * @param wb
     * @param sheetIndex    sheet页下标：从0开始
     * @param startReadLine 开始读取的行:从0开始
     * @param tailLine      去除最后读取的行
     */
    private void readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        Row row = null;

        for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) {
            row = sheet.getRow(i);
            for (Cell c : row) {
                boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
                // 判断是否具有合并单元格
                if (isMerge) {
                    String rs = getMergedRegionValue(sheet, row.getRowNum(), c.getColumnIndex());
                    System.out.print(rs + "");
                } else {
                    System.out.print(getCellValue(c) + "");
                }
            }
            System.out.println();

        }

    }

    /**
     * 获取合并单元格的值
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {

                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell);
                }
            }
        }

        return null;
    }

    /**
     * 判断合并了行
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    private boolean isMergedRow(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row == firstRow && row == lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param sheet
     * @param row    行下标
     * @param column 列下标
     * @return
     */
    private boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断sheet页中是否含有合并单元格
     *
     * @param sheet
     * @return
     */
    private boolean hasMerged(Sheet sheet) {
        return sheet.getNumMergedRegions() > 0 ? true : false;
    }

    /**
     * 合并单元格
     *
     * @param sheet
     * @param firstRow 开始行
     * @param lastRow  结束行
     * @param firstCol 开始列
     * @param lastCol  结束列
     */
    private void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public String getCellValue(Cell cell) {

        if (cell == null)
            return "";

        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

            return cell.getStringCellValue();

        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

            return String.valueOf(cell.getBooleanCellValue());

        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

            return cell.getCellFormula();

        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }

    public void readExcel() throws IOException {
        InputStream is = new FileInputStream("C:\\Users\\liu\\Desktop\\素质点 (1).xls");

        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);

        // 循环工作表Sheet
        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }
            Long previousId = null;// 上条记录的id
            Integer previousLevel = 1;// 上条记录的等级

            // 循环行Row
            // 第一行为说明，不进行读取
            for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                KeyQualityEntity kqe = new KeyQualityEntity();
                // 设置描述
                kqe.setDetail(hssfRow.getCell(0).getStringCellValue());
                // 设置提升建议
                kqe.setPromptAdvice(hssfRow.getCell(1).getStringCellValue());

                // 循环列Cell
                // 0：描述 1：提升意見 2：一級 3：二級 4：三級 5：四級 6：五級
                Integer level; // 标示当前行的层级
                for (int coluNum = 2; coluNum <= 6; coluNum++) {
                    HSSFCell hssfCell = hssfRow.getCell(coluNum);
                    // 如果该列为null则循环下一列
                    if (hssfCell == null) {
                        continue;
                    }
                    String name = hssfRow.getCell(coluNum).getStringCellValue();
                    if (StringUtils.isNotBlank(name)) {
                        level = coluNum - 1;
                        kqe.setName(name);
                        kqe.setLevel(level);
                        // 如果上条记录的等级比本条记录等级小，则说明上条记录是该记录的父亲，更新 previousId和
                        // previousLevel
                        if (previousLevel < level) {
                            /*
                             * 伪代码
                             * kqe.setParent(new
                             * KeyQualityEntity(previousId));   //设置上条记录的id为本记录父亲id
                             * kqe.save();       //保存进数据库并设置保存后id在kqe里
                             */
                            previousId = kqe.getId();
                            previousLevel = level;
                        } else {
                            /**
                             伪代码，查询按id倒序的等级等于当前记录等级level的第一条数据，如果结果为null，说明是第一个根
                             ，设置parentId为null，如果有结果，则设置当前记录parentId为查出来的记录的parentId
                             List<String> usePara = new ArrayList<>();
                             usePara.add("order by id desc");
                             ListResult<KeyQualityEntity> result = generalBeanDao.query(new KeyQualityEntity().setLevel(level), "keyQualityEntity",
                             null, usePara);
                             List<KeyQualityEntity> keyQualityEntities = result.getList();
                             if (CollectionUtils.isEmpty(keyQualityEntities)) {
                             kqe.setParent(new KeyQualityEntity(null));
                             } else {
                             kqe.setParent(new KeyQualityEntity(keyQualityEntities.get(0).getParentId()));
                             }
                             */
                            kqe.save(); //  保存kqe到数据库，得到保存后id
                            previousId = kqe.getId(); //更新    previousId和   previousLevel
                            previousLevel = level;
                        }
                    }
                }
            }
        }
}
