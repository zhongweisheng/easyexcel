package com.alibaba.easyexcel.test.demo.zws;

public class CityExcelSelectedImpl implements ExcelDynamicSelect{

    @Override
    public String[] getSource() {
//        CityMapper cityMapper = SpringContextUtil.getBean(CityMapper.class);
//
//        return cityMapper.selectAllCity().toArray(new String[]{});
        return new String[]{"A","B","C"};
    }

}
