/**
 * Created by HP on 2019/7/8.
 */
function getExplorer() {
    var explorer = window.navigator.userAgent;
    //ie
    if(explorer.indexOf("MSIE") >= 0) {
        return 'ie';
    }
    //firefox
    else if(explorer.indexOf("Firefox") >= 0) {
        return 'Firefox';
    }
    //Chrome
    else if(explorer.indexOf("Chrome") >= 0) {
        return 'Chrome';
    }
    //Opera
    else if(explorer.indexOf("Opera") >= 0) {
        return 'Opera';
    }
    //Safari
    else if(explorer.indexOf("Safari") >= 0) {
        return 'Safari';
    }
}
function ExcelExportIE(tableid,name)     {
    var table = document.getElementById(tableid); //获取页面的table
    var excel = new ActiveXObject("Excel.Application"); //实例化Excel.Application对象
    var workB = excel.Workbooks.Add(); ////添加新的工作簿
    var sheet = workB.ActiveSheet;     //激活一个sheet
    /**将页面table写入到Excel中，具体复杂情况（合并单元格等）可在这里面具体操作**********/
    var LenRow = table .rows.length; //以下为循环遍历获取页面table的cell元素
    for (i = 0; i < LenRow ; i++)         {
        var lenCol = table.rows(i).cells.length-1;
        for (j = 0; j < lenCol ; j++)             {
            sheet.Cells(i + 1, j + 1).value = table.rows(i).cells(j).innerText; //通过该语句将table的每个  //cell赋予Excel 当前Active的sheet下的相应的cell下
        }
    }
    //excel.Visible = true;//设置excel为可见
    excel.UserControl = true; //将Excel交由用户控制

    try {
        var fname = excel.Application.GetSaveAsFilename(name+".xls",
            "Excel Spreadsheets (*.xls), *.xls");
    } catch(e) {
        // print("Nested catch caught " + e);
    } finally {
        sheet.SaveAs(fname);
        sheet.Close(savechanges = false);
        excel.Quit();
        excel = null;
    }
}
function ExcelExportOther(table, name) {
    var uri = 'data:application/vnd.ms-excel;base64,'
        , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->'+
        ' <style type="text/css">'+
        '.excelTable  {'+
        'border-collapse:collapse;'+
        ' border:thin solid #999; '+
        '}'+
        '   .excelTable  th {'+
        '   border: thin solid #999;'+
        '  padding:20px;'+
        '  text-align: center;'+
        '  border-top: thin solid #999;'+
        ' '+
        '  }'+
        ' .excelTable  td{'+
        ' border:thin solid #999;'+
        '  padding:2px 5px;'+
        '  text-align: center;'+
        ' }</style>'+'</head><body><table border="1">{table}</table></body></html>'
        , base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) }
        , format = function(s, c) { return s.replace(/{(\w+)}/g, function(m, p) { return c[p]; }) }
    if (!table.nodeType) table = document.getElementById(table)
    var ctx = {worksheet: name || 'Worksheet', table: table.innerHTML};
    var downloadLink = document.createElement("a");
    downloadLink.href = uri + base64(format(template, ctx));
    downloadLink.download = name+".xls";
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}
function tableToExcel(tableid) {
    if(getExplorer() == 'ie' || getExplorer() == undefined) {
        ExcelExportIE(tableid,'电子发票查询下载')
    }else {
        ExcelExportOther(tableid,'电子发票查询下载');
    }
}