var xlsParser = require('node-xlrd');
var excelPort = require('excel-export');
var fs=require('fs');

/***
 *解析excel文件
 * @param path {String} 要解析的excel文件路径
 * @param onParseDone {Function} 解析后的回调函数
 ***/
function parseExcel(path, onParseDone)
{
    var excelDatas = [];
    xlsParser.open(path, function onParsed(err, workbook)//workbook解析成功的得到的工作簿对象
    {
        if (err)//解析出错
        {
            if (onParseDone)
            {
                onParseDone(err);
            }
            return;
        }

        var sheetCount = workbook.sheet.count;//工作表数量
        console.log("工作表数量---->", sheetCount);
        for (var sheetId = 0; sheetId < sheetCount; sheetId++)//遍历工作表
        {
            var sheetData = [];
            var sheet = workbook.sheets[sheetId],//获取当前工作表
                rowCount = sheet.row.count,//该工作表的行数
                colCount = sheet.column.count;//该工作表的列数
            for (var rowID = 0; rowID < rowCount; rowID++)//遍历行数据
            {    // rIdx：行数；cIdx：列数
                var data = [];
                for (var colID = 0; colID < colCount; colID++)//遍历列数据
                {
                    try
                    {
                        data[colID] = sheet.cell(rowID, colID);
                    } catch (e)
                    {
                        console.log(e.message);
                    }
                }
                sheetData[rowID] = data;
                console.log("表数据---->", data);
            }
            excelDatas[sheetId] = sheetData;
        }

        if (onParseDone)
        {
            onParseDone(null, excelDatas);
        }
    });
}

/***
 *导出excel
 ***/
function exportToExcel(conf,excelName)
{
    var result = excelPort.execute(conf);
    var filePath = excelName+ ".xlsx";

    fs.writeFile(filePath, result, 'binary', function (err)
    {
        if (err)
        {
            console.log(err);
            return;
        }
        console.log("----解析成功----");
    });
}

exports.parseExcel = parseExcel;
exports.exportToExcel = exportToExcel;