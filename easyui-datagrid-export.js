(function($){
    function getRows(target){
        var state = $(target).data('datagrid');
        if (state.filterSource){
            return state.filterSource.rows;
        } else {
            return state.data.rows;
        }
    }
    
    function toHtml(target, rows){
        rows = rows || getRows(target);
        var dg = $(target);
        var data = ['<table border="1" rull="all" style="border-collapse:collapse">'];
        var fields = dg.datagrid('getColumnFields',true).concat(dg.datagrid('getColumnFields',false));
        var trStyle = 'height:32px';
        var tdStyle0 = 'vertical-align:middle;padding:0 4px';
        
        var frozen_columns_head = dg.datagrid("options").frozenColumns;
        var columns_head = dg.datagrid("options").columns;
        
        var getColumnMaxRowspan = function(columns_head) {
        	var max_rowspan = 0;
        	for(var i in columns_head) {
        		var columns = columns_head[i];
        			var rowspan = columns[0].rowspan;
        			if(rowspan > 1) {
        				max_rowspan += rowspan;
        			}else {
        				max_rowspan += 1;
        			}
        	}
        	return max_rowspan;
        }
        
//        var columnMaxRowspan = getColumnMaxRowspan(columns_head);
        
        var column_stylers = new Object();
        var setColumns = function(frozen_columns_head, columns_head) {
        	for(var i in columns_head) {
        		var frozen_columns = frozen_columns_head[i];
	        	var columns = columns_head[i];
	        	
	        	data.push('<tr style="'+trStyle+'">');
	        	var setColumn = function(columns, is_frozen) {
	        		for(var i in columns) {
		        		var column = columns[i];
		        		var h_rowsapn = column.rowspan;
		        		if(column.frowspan) {
		        			h_rowsapn = column.frowspan;
		        		}
//		        		else if(is_frozen) {
//		        			h_rowsapn = columnMaxRowspan;
//		        		}
		        		var tdStyle = tdStyle0 + ';width:'+column.boxWidth+'px;';
		        		data.push('<th rowspan="'+h_rowsapn+'" colspan="'+column.colspan+'" style="'+tdStyle+'">'
		        				+column.title+'</th>');
		        		column_stylers[column.field] = column.styler; // 取出调用者自定义的列style
	        		}
	        	}
	        	if(frozen_columns){setColumn(frozen_columns, true);}
	        	setColumn(columns);
	        	data.push('</tr>');
        	}
        }
        setColumns(frozen_columns_head, columns_head);
        
        var index = 0;
        $.map(rows, function(row){
            data.push('<tr style="'+trStyle+'">');
            for(var i=0; i<fields.length; i++){
                var field = fields[i];
                var value = row[field];
                if(value == null){
                	value = ""
                }
                var style = null;
                var style_fn = column_stylers[field];
                if(style_fn != null) {
                	style = style_fn(value,row,index); // 执行调用者自定义的列style
                }
                data.push('<td style="'+tdStyle0+';'+style+'">'+value+'</td>');
            }
            data.push('</tr>');
            index++;
        });
        data.push('</table>');
        return data.join('');
    }

    function toExcel(target, param){
        var filename = null;
        var rows = null;
        var worksheet = 'Worksheet';
        if (typeof param == 'string'){
            filename = param;
        } else {
            filename = param['filename'];
            rows = param['rows'];
            worksheet = param['worksheet'] || 'Worksheet';
        }
        var dg = $(target);
        var uri = 'data:application/vnd.ms-excel;base64,'
        , template = 
        	'<html xmlns:o="urn:schemas-microsoft-com:office:office" '
    			+ 'xmlns:x="urn:schemas-microsoft-com:office:excel" '
    			+ 'xmlns="http://www.w3.org/TR/REC-html40">'
				+ '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">'
				+ '<head><!--[if gte mso 9]>'
					+ '<xml>'
					+ '<x:ExcelWorkbook>'
						+ '<x:ExcelWorksheets>'
							+ '<x:ExcelWorksheet>'
							+ '<x:Name>{worksheet}</x:Name>'
							+ '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>'
							+ '</x:ExcelWorksheet>'
						+ '</x:ExcelWorksheets>'
					+ '</x:ExcelWorkbook>'
					+ '</xml><![endif]-->'
				+ '</head>'
				+ '<body>{table}</body>'
    		+ '</html>'
        , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
        , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }

        var table = toHtml(target, rows);
        var ctx = { worksheet: worksheet, table: table };
        var data = base64(format(template, ctx));
        if (window.navigator.msSaveBlob){
            var blob = b64toBlob(data);
            window.navigator.msSaveBlob(blob, filename);
        } else {
            var alink = $('<a style="display:none"></a>').appendTo('body');
            alink[0].href = uri + data;
            alink[0].download = filename;
            alink[0].click();
            alink.remove();
        }
    }
    
    function b64toBlob(data){
        var sliceSize = 512;
        var chars = atob(data);
        var byteArrays = [];
        for(var offset=0; offset<chars.length; offset+=sliceSize){
            var slice = chars.slice(offset, offset+sliceSize);
            var byteNumbers = new Array(slice.length);
            for(var i=0; i<slice.length; i++){
                byteNumbers[i] = slice.charCodeAt(i);
            }
            var byteArray = new Uint8Array(byteNumbers);
            byteArrays.push(byteArray);
        }
        return new Blob(byteArrays, {
            type: ''
        });
    }

    $.extend($.fn.datagrid.methods, {
        toHtml: function(jq, rows){
            return toHtml(jq[0], rows);
        },
        toExcel: function(jq, param){
            return jq.each(function(){
                toExcel(this, param);
            });
        }
    });
})(jQuery);
