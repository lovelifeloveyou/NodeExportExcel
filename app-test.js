var app = require('./app');

(function init()
{
    /*测试解析excel*/
    // parseExcelTest();

    /*测试导出excel*/
    exportToExcelTest();
})();

/***
 *测试解析excel
 ***/
function parseExcelTest()
{
    var path = './test.xls';
    app.parseExcel(path, function onParsedDone(err, excelData)
    {
        if (err)
        {
            console.log("解析出错---->", err);
            return;
        }
        console.log("解析后的test---->", excelData);
    });
}

/***
 *测试导出excel
 ***/
function exportToExcelTest()
{
    var sheetConfig = {};//表格配置对象
    sheetConfig.cols = //列
        [
            {caption: '名称', type: 'string', width: 20},
            {caption: '年龄', type: 'number', width: 30},
            {caption: '地址', type: 'string', width: 30},
        ];
    var rowDatas = //行数据
        [
            ['kk', 18, '莲花新村'],
			['kk', 18, '莲花新村'],
			['kk', 18, '莲花新村'],
			['kk', 18, '莲花新村'],
			['kk', 18, '莲花新村'],
			['kk', 18, '莲花新村'],//行数据
        ];
    sheetConfig.rows = rowDatas;//填充行数据
    app.exportToExcel(sheetConfig, 'expTest');
}