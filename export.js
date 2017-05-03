/*
                                      __  ____      _______          
                                     |  \/  \ \    / /  __ \         
   ___ __ _ _ __ ___  _ __  _   _ ___| \  / |\ \  / /| |__) |__  ___ 
  / __/ _` | '_ ` _ \| '_ \| | | / __| |\/| | \ \/ / |  ___/ _ \/ __|
 | (_| (_| | | | | | | |_) | |_| \__ \ |  | |  \  /  | |_ |  __/\__ \
  \___\__,_|_| |_| |_| .__/ \__,_|___/_|  |_|   \/   |_(_) \___||___/
                     | |                                             
                     |_|                                             

Sample library created by Jose Alarcon (http://www.jasoft.org/)
*/

(function(w){

    //Helper class in order to define formats
    function Format(name, type) {
        this.name = name;
        this.mimeType = type;
    }

    ///// CONSTANTS
    const __formats = {
        'xls': new Format('Excel', 'data:application/vnd.ms-excel;base64'),
        'csv': new Format('CSV', 'data:text/plain'),
        'doc': new Format('Word', 'data:application/vnd.ms-word'),
        'html': new Format('HTML', 'data:text/html')
    };

    const __excelTemplate = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheetName}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>';    //The template to better export to Excel
    const __encodingMeta = '<meta charset="UTF-8">';  //The encoding to be able to work with non-english characters correctly
    ///// END CONSTANTS

    //// HELPER FUNCTIONS
    /**
     * Get the suitable format from the file extension
     * @param {string} fileName 
     */
    var __getFormatFromFileExt = function(fileName){
        var posDot = fileName.lastIndexOf('.');
        if (posDot <= 0)
            throw new Error('The file must have an extension');
        
        var ext = fileName.substring(posDot+1).toLowerCase();
        var formatInfo = __formats[ext];    //Get the format info from the proprties fo the format object
        if (!formatInfo)
            throw new Error('Unsupported format!!');
        else
            return formatInfo;
    };

    /**
     * Converto to base64
     * @param {string} str 
     */
    var __base64 = function(str) {
        return w.btoa(unescape(encodeURIComponent(str)))
    };

    /**
     * Replaces fields in the form {fieldName} in a template
     * @param {*} str The tempplate where field must be replaced
     * @param {*} fieldName The name of the field to be replaced (without the brackets)
     * @param {*} value The value to use as a replacement
     */
    var __replaceFields = function(template, fieldName, fieldValue) {
        return template.replace('{' + fieldName + '}', fieldValue);
    };

    /**
     * Converts a table element into a list using the specified separator, one row per line.
     * @param {DOMElement} eltTable 
     * @param {string} separator 
     */
    var __table2List = function (eltTable, separator) {
        if ( !(eltTable && eltTable.tagName == 'TABLE') )   //If it's not a table...
            throw new Error('Only tables can be converted to this format!!');
        
        //Hold the results in an array to optimize the process a little bit (instead of just using strings, althoug it could be obviously enhanced, it's not bad for a quick approach)
        var res = new Array(eltTable.rows.length);
        //loop through rows and cells to create each line
        for(var r=0; r<eltTable.rows.length; r++){
            var cells = eltTable.rows[r].cells;
            var line = new Array(cells.length);
            for(var c=0; c<cells.length; c++) {
                line[c] = cells[c].innerText;
            }
            res[r] = line.join(separator);
        }
        return res.join('\n');
    }
    //// END HELPER FUNCTIONS

    //// MAIN FUNCTION
    /**
     * Exports the table to the format specified by the filename
    var _export = function(DOMElmt, fileName, dataName) {
     * @param {element} objTable The DOM element to export (generally a table)
     * @param {string} fileName The name of the file to export the table to
     * @param {string} dataName Optional, the name of the worksheet when exporting to XLS
     */
    var _export = function(DOMElmt, fileName, dataName) {
        //Get the file format info from the file extension
        var format = __getFormatFromFileExt(fileName);

        //Get the info from the element
        var html = DOMElmt.outerHTML;

        //Generate file
        var rawData = '';   //For IE
        var resData = format.mimeType + ',';
        switch (format.name) {
            case 'Excel':
                dataName = dataName || 'ExportedData';    //Default name for the worksheet
                rawData = __replaceFields(__excelTemplate, 'table', html);    //The table data
                rawData = __replaceFields(rawData, 'worksheetName', dataName);  //The name of the Excel worksheet
                resData += __base64(__encodingMeta + rawData);
                break;
            case 'CSV':
                rawData = __table2List(DOMElmt);
                resData += '\ufeff' + encodeURIComponent(rawData);    //Adding the UTF-8 BOM to force Excel to interpret the data in UTF-8, so that non-English characters are correctly displayed
                break;
            case 'Word':
                rawData = html;
                resData += encodeURIComponent(__encodingMeta + html);
                break;
            case 'HTML':
                rawData = html;
                resData += __encodingMeta + html;
                break;
        }

        if (navigator.userAgent.indexOf('MSIE ') > 0 || navigator.userAgent.indexOf('Trident/') > 0)      // If Internet Explorer
        {
            //We need to append an iframe and write directly in it :-S Encodings don't work either: just the plain content.
            var ifr = w.document.createElement('iframe');
            ifr.style.display= 'none';
            w.document.body.appendChild(ifr);
            ifr.contentWindow.document.open("txt/html","replace");
            ifr.contentWindow.document.write(rawData);
            ifr.contentWindow.document.close();
            ifr.focus();
            ifr.contentWindow.document.execCommand("SaveAs",true,fileName);
            w.document.body.removeChild(ifr);
        }  
        else {  //Rest of the browsers (except MS Edge: no way to have it working there)
            //Create a link, not in the DOM, to get the resulting document
            var link = w.document.createElement('a');
            link.download = fileName;
            link.href = resData;
            //In firefox only links attached to the DOM will work so we need to attach it first
            link.style.display = 'none';
            w.document.body.appendChild(link);
            link.click();   //Simulation of a click on a real link
            w.document.body.removeChild(link);
        }
    };
    ////END MAIN FUNCTION

    //Export the funcionality through a global object
    if (!w.Exporter) {
        w.Exporter = {
            'export': _export
        };
    }
})(window);
