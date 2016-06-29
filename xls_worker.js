var config = {
    /*theWhat -> the number of sheets in the wb*/
    theWhat: {},
    sheetNames: '',
    range: '',
    newVal: '',
    workSheet: '',
    wb: '',
    f: '',
    clientNames: '',
    tasksNames: '',
    tasksNumber: '',
	OANames : '',
    IDs: '',
    IDs_plus_Tasks: [],
    fnArr: [function(el) {
        $(el).css('background-color') !== "rgba(0, 0, 0, 0)" ?
            $(el).css('background-color', '') :
            $(el).css('background-color', '#CCEEFF');
    }],
    defPreventer: function(e) {
        e.originalEvent.stopPropagation();
        e.originalEvent.preventDefault();
        config.fnArr.forEach(function(i, j) {
            if (typeof i == 'function') {
                i(e.target);
            }
        });
        config.fnArr = [];
    },

    init: function() {
        config.helper = [];
        $('#drag-and-drop').on(
            'dragover',
            config.defPreventer);

        $('#drag-and-drop').on(
            'dragenter',
            config.defPreventer);
    },
    table_template: '<table>' + '<thead>' + '<tr></tr>' + '</thead>' + '<tbody></tbody>' + '</table>',
    preview_template: '<div class="' + 'table-preview">' + '</div>',
    /*helper-function to construct html-ized xlsx table*/
    onloadHandlerSub: function(i, a, wb, the_number_of_rows, b) {
        if (wb.Sheets.hasOwnProperty(i)) {
            for (var n = 0; n < the_number_of_rows; n++) {
                var la = a + n;
                var row = n;
                if (wb.Sheets[i][la]) {
                    var dataSet = wb.Sheets[i][la];
                    var textValue = dataSet['w'] ? dataSet['w'] : dataSet['v'];
                    if (b == 0) {
                        var selector1 = '.table-preview[sheet=' + '"' + i + '"' + ']' + ' tbody';
                        var selector2 = '.table-preview[sheet=' + '"' + i + '"' + ']' + ' tbody tr';
                        $(selector1).append($("<tr></tr>", {
                            "row": row,
                            "column": a
                        }));
                        $(selector2).last().append($("<td></td>", {
                            "lineNum": n
                        }).text(row));
                        $(selector2).last().append($("<td></td>", {
                            "ref": a + row
                        }).text(textValue));
                        /* $('#table-preview tbody').append($("<tr></tr>", {"row": row, "column" : a}));
                        $('#table-preview tbody tr').last().append($("<td></td>", {"lineNum" : n}).text(row));
                        $('#table-preview tbody tr').last().append($("<td></td>", {"ref" : a+row}).text(textValue)); */
                    } else {
                        //var lookup = '.table-preview tbody tr[row="'+row+'"]';
                        var lookup = '.table-preview[sheet=' + '"' + i + '"' + ']' + ' tbody' + ' tr[row="' + row + '"]';
                        $(lookup).append($("<td></td>", {
                            "ref": a + row
                        }).text(textValue));
                    }
                } else if (parseInt(la.match(/\d+/)) !== 0) {
                    //var lookup = '.table-preview tbody tr[row="'+ parseInt(la.match(/\d+/)) +'"]';
                    var lookup1 = '.table-preview[sheet=' + '"' + i + '"' + ']' + ' tbody' + ' tr[row="' + parseInt(la.match(/\d+/)) + '"]';
                    var lookup2 = '.table-preview[sheet=' + '"' + i + '"' + ']' + ' tbody';
                    var lookup3 = '.table-preview[sheet=' + '"' + i + '"' + ']' + ' tbody tr';
                    if (!$(lookup1).length) {
                        $(lookup2).append($("<tr></tr>", {
                            "row": parseInt(la.match(/\d+/)),
                            "column": a
                        }));
                        $(lookup3).last().append($("<td></td>", {
                            "lineNum": parseInt(la.match(/\d+/))
                        }).text(parseInt(la.match(/\d+/))));
                    }
                    $(lookup1).append($("<td></td>", {
                        "ref": a + row
                    }).text(""));
                }
            }
        }
    },
    htmlize: function() {
        var subRoutine = function(i) {
            if (config.wb.Sheets.hasOwnProperty(i)) {
                /*appending the table-preview into the wrapper && the tableTab per each sheet*/
                var parser = new DOMParser(),
                    tableTab = parser.parseFromString(config.table_template, "text/html"),
                    tablePreview = parser.parseFromString(config.preview_template, "text/html");
                tablePreview = tablePreview.querySelector('.table-preview');
                tableTab = tableTab.querySelector('table');
                tablePreview.setAttribute('sheet', i);
                tableTab.setAttribute('sheet', i),
                    selector = '.table-preview[sheet=' + '"' + i + '"' + ']';
                document.querySelector('#wrapper').appendChild(tablePreview);
                document.querySelector(selector).appendChild(tableTab);
                var range = config.wb.Sheets[i]['!ref'];
                var the_number_of_rows = parseInt(range.split(':')[1].match(/\d+/)[0]);
                var letterRanges = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
                letterRanges.forEach(function(a, b) {
                    var row = a + '0',
                        selector = '.table-preview table[sheet="' + i + '"' + ']' + ' thead tr';
                    if (b == 0) {
                        $(selector).first().append($("<td></td>", {
                            "lineNum": row
                        }).text('#'));
                        $(selector).first().append($("<td></td>", {
                            "row": row
                        }).text(a));
                        config.onloadHandlerSub(i, a, config.wb, the_number_of_rows, b);
                    } else if (Object.keys(config.wb['Sheets'][i]).some(function(el, i, arr) {
                            var rx = new RegExp(a + "\\d?");
                            return el.match(rx);
                        })) {
                        $(selector).first().append($("<td></td>", {
                            "row": row
                        }).text(a));
                        config.onloadHandlerSub(i, a, config.wb, the_number_of_rows, b);
                    }
                });
            }
        };

        config.wb.SheetNames.forEach(function(i, j) {
            /*before calling the html-ize subRoutine, need to make sure each 
             sheet is fitted to its own html tab!*/
            subRoutine(i);
        });
    },
    row: '',
    lineNum: '',
	
    colorSubroutine: function(el) {
        if (el && $(el).css('backgroundColor') && ($(el).css('background-color') == "rgb(170, 170, 170)")) {
            $(el).css('backgroundColor', 'white');
        } else {
            $(el).css('backgroundColor', 'rgb(170, 170, 170)');
            /*get the range by color!*/
            if (el.getAttribute('ref')) {
                config.helper.push(el.getAttribute('ref'));
            }
        }
    },

    onclicker: function(e) {
        var trgt = e.target;
        if (trgt.tagName == "TD") {
            config.colorSubroutine(trgt);
            if (trgt.getAttribute('row') && $(trgt).closest('thead').length) {
                config.row = trgt.getAttribute('row').match(/\D+/);
                $('tbody tr').each(function() {
                    $(this.children).each(function() {
                        if ($(this).attr('ref') && ~config.row.indexOf($(this).attr('ref').match(/\D+?/g)[0])) {
                            config.colorSubroutine(this);
                        }
                    });
                });
                if (config.helper.length) {
                    Object.keys(config.theWhat).forEach(function(i, j) {
                        config.theWhat[i] = [];
                        config.helper.forEach(function(k, l) {
                            var t = {};
                            t[k] = '';
                            config.theWhat[i].push(t);
                        });
                    });
                }
            } else if (trgt.getAttribute('linenum') && $(trgt).closest('tbody').length) {
                config.lineNum = trgt.getAttribute('lineNum').match(/\d+/);
                $('tbody tr').each(function() {
                    $(this.children).each(function() {
                        if (~config.lineNum.indexOf($(this).closest('tr').attr('row'))) {
                            config.colorSubroutine(this);
                        }
                    });
                });
                if (config.helper.length) {
                    Object.keys(config.theWhat).forEach(function(i, j) {
                        config.theWhat[i] = [];
                        config.helper.forEach(function(k, l) {
                            var t = {};
                            t[k] = '';
                            config.theWhat[i].push(t);
                        });
                    });
                }
            }
        }
    },

    processWb: function() {
        /*processing the workbook here:*/
        var wopts = {
            bookType: 'xlsx',
            bookSST: false,
            type: 'binary'
        };
        var wbout = XLSX.write(config.wb, wopts);

        function s2ab(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }
        /* the saveAs call downloads a file on the local machine */
        saveAs(new Blob([s2ab(wbout)], {
            type: ""
        }), "MyExcel.xlsx");
    }, 
	
	getItemNamesByColumn : function(workSheet, columnName, unique) {
                    var workbook = config.wb.Workbook.Sheets;
                    var letterRanges = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
                    var returnable = [];
                    workbook.forEach(function(sheet) {
                        if (sheet['name'] == workSheet) {
                            var ref = config.wb.Sheets[sheet['name']]['!ref'];
                            var upperBound = parseInt(config.wb.Sheets[sheet['name']]['!ref'].split(':')[1].match(/\d+/));
                            letterRanges.forEach(function(letter) {
                                var _columnName = config.wb.Sheets[sheet['name']][letter + 1] ? config.wb.Sheets[sheet['name']][letter + 1]['v'] : '';
                                var theLetter = '';
                                if (_columnName == columnName) {
                                    theLetter = letter;
                                    while (upperBound > 1) {
                                        config.wb.Sheets[sheet['name']][theLetter + upperBound] &&
                                            config.wb.Sheets[sheet['name']][theLetter + upperBound]['v'] ?
                                            returnable.push(config.wb.Sheets[sheet['name']][theLetter + upperBound]['v']) :
                                            returnable;
                                        upperBound--;
                                    }
                                }
                            });
                        }
                    });
                    return returnable.reverse();
	}, 
	
	combine2Arrays : function(newArrayName, Array1, Array2) {
			if(!config[newArrayName]){
				config[newArrayName] = [];
			}
			for (var i = 0; i < Array1.length; i++) {
				Array1[i] ? config[newArrayName].push(Array1[i] + ' ' + Array2[i]) : '';
            }
			config[newArrayName];
	}, 
	
	complexCombine : function(newArrayName, Array1, Array2){ 		
				var router = function(z){
					Array2.forEach(function(what, i) {
						newArrayName.push(Array1[z] + ' ' + what);
					});
				}
				Array1.forEach(function(what, z) {
					router(z);
				});
	},
	
	rangeSeeker : function(workSheet /*Final List*/ , columnName /*Oracle Project Name*/ ) {
                    var workbook = config.wb['Workbook']['Sheets'];
                    var range;
                    var letterRanges = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
                    var ref;
                    var splitRefArrOf2;
                    var upperBoundNum;
                    var higherBoundNum;
                    var upperBoundLetter;
                    var lowerBoundLetter;
                    var columnNameLetter;
                    workbook.forEach(function(sheet) {
                        if (sheet['name'] == workSheet) {
                            ref = config.wb.Sheets[sheet['name']]['!ref'];
                            splitRefArrOf2 = ref.split(':');
                            upperBoundNum = parseInt(config.wb.Sheets[sheet['name']]['!ref'].split(':')[0].match(/\d+/));
                            lowerBoundNum = parseInt(config.wb.Sheets[sheet['name']]['!ref'].split(':')[1].match(/\d+/));
                            upperBoundLetter = ref.split(':')[0].match(/\D/)[0];
                            lowerBoundLetter = ref.split(':')[1].match(/\D/)[0];
                            for (var i = letterRanges.length; i--;) {
                                if (config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum] &&
                                    config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v']) {
                                    if (config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v'] == columnName ||
                                        config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v'].includes(columnName)) {
                                        range = letterRanges[i] + (upperBoundNum + 1) + ":" + letterRanges[i] + (upperBoundNum + 1);
                                    }
                                }
                            }
                        }
                    });
                    return range;
        }, 
		
		 writeable : function(workbook, range, data /*config.clientNames*/ ) {
                    Object.keys(config.theWhat).forEach(function(i, j) {
                        if (config.theWhat[i].forEach && i == workbook) {
                            config.theWhat[i].forEach(function(z, w) {
                                var splitRange = range.split(':');
                                var startRange = range.split(':')[0];
                                var startRangeLetter = splitRange[0].match(/\D+/)[0];
                                var startRangeNumber = parseInt(splitRange[1].match(/\d+/)[0]);
								if(data && data.length){
									for (var iter = 0; iter < data.length; iter++) {
										z[startRange + '-' + startRangeLetter + (startRangeNumber + iter)] = data[iter];
									}	
								}
                            });
                        }
                    });
		}, 
		
		 rangeSeeker : function(workSheet /*Final List*/ , columnName /*Oracle Project Name*/ ) {
                    var workbook = config.wb['Workbook']['Sheets'];
                    var range;
                    var letterRanges = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
                    var ref;
                    var splitRefArrOf2;
                    var upperBoundNum;
                    var higherBoundNum;
                    var upperBoundLetter;
                    var lowerBoundLetter;
                    var columnNameLetter;
                    workbook.forEach(function(sheet) {
                        if (sheet['name'] == workSheet) {
                            ref = config.wb.Sheets[sheet['name']]['!ref'];
                            splitRefArrOf2 = ref.split(':');
                            upperBoundNum = parseInt(config.wb.Sheets[sheet['name']]['!ref'].split(':')[0].match(/\d+/));
                            lowerBoundNum = parseInt(config.wb.Sheets[sheet['name']]['!ref'].split(':')[1].match(/\d+/));
                            upperBoundLetter = ref.split(':')[0].match(/\D/)[0];
                            lowerBoundLetter = ref.split(':')[1].match(/\D/)[0];
                            for (var i = letterRanges.length; i--;) {
                                if (config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum] &&
                                    config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v']) {
                                    if (config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v'] == columnName ||
                                        config.wb.Sheets[sheet['name']][letterRanges[i] + upperBoundNum]['v'].includes(columnName)) {
                                        range = letterRanges[i] + (upperBoundNum + 1) + ":" + letterRanges[i] + (upperBoundNum + 1);
                                    }
                                }
                            }
                        }
                    });
                    return range;
                }, 
				
				iterate_to_write_subRoutine : function(theWhat, auxArrName, length) {
					/*creates a sub-array by making copies of theWhat into auxArrName the assigned number of times per length	*/
					var returnable = {};
						returnable[auxArrName] = [];
						for (var i = 0; i < length; i++) {
							returnable[auxArrName].push(theWhat);
						}
					return returnable[auxArrName];	
				}, 
				rangeIncrementer : function(range, byNum) {
                    var interim = range.split(':'),
                        interimLowerBound = interim[1],
                        interimLowerBoundLetter = interim[1].match(/\D+/)[0],
                        interimLowerBoundNumber = parseInt(interim[1].match(/\d+/)[0]);
                    interimLowerBoundNumber = interimLowerBoundNumber + byNum - 1; //changed to skip the empty line in-between
                    interim[1] = interimLowerBoundLetter + interimLowerBoundNumber;
                    range = range.replace(/[A-Z]\d+$/, interim[1]);
                    return range;
                }
};

$(document).ready(function() {
    config.init();
    $('#drag-and-drop').on(
        'drop',
        function(e) {
            config.defPreventer(e);
            if (e.originalEvent.dataTransfer) {
                if (e.originalEvent.dataTransfer.files.length) {
                    var files = e.originalEvent.dataTransfer.files;
                    config.f = files[0];
                    var reader = new FileReader(),
                        name = config.f.name;
                    reader.onload = function(e) {
                        var data = e.target.result;
                        config.wb = XLSX.read(data, {
                            type: 'binary'
                        });
                        /*get the number of worksheets, i.e., the tabs:*/
                        config.sheetNames = config.wb.SheetNames;
                        config.sheetNames.forEach(function(i, j) {
                            config.theWhat[i] = [{}];
                        });
                        if (!config.sheetNames.length) {
                            function UserException(message) {
                                this.message = message;
                                this.name = "UserException";
                            }
                            throw new UserException("The Excel File Seems To Have No Sheets!");
							$('#drag-and-drop').addClass('failure');	
                        }
                        //make sure we have got only 1 sheet, because i do not have the multi-sheet representation:
                        /* config.htmlize(); */
						if(config.wb){
							$('#drag-and-drop').addClass('success');	
						}
                    };
                    reader.readAsBinaryString(config.f);
                    config.fnArr.push(function(el) {
                        $(el).css('background-color') !== "rgba(0, 0, 0, 0)" ?
                            $(el).css('background-color', '') :
                            $(el).css('background-color', '#CCEEFF');
                    });
                    config.fnArr.forEach(function(i, j) {
                        if (typeof i == 'function') {
                            i(e.target);
                        }
                    });
                }
            }
        }
    );
   
    $('#process_wb').on('click', function(e) {
            //automation specifically for Natalia
                for (var sheet in config.theWhat) {
                    if (config.theWhat[sheet] && ({}).toString.call(config.theWhat[sheet]) == '[object Array]') {
                        config.theWhat[sheet].forEach(function(range_value_pair, j) {
                            for (var cell in range_value_pair) {
                                if (cell.match(/[-]/)) {
                                    var column = cell.split('-')[0];
                                    var row = cell.split('-')[1];
                                    var rowNumber = parseInt(row.match(/\d+/)[0]);
                                    var rangeLetter = column.match(/\D+/g)[0];
                                    var val = range_value_pair[cell];
                                    var ref = config.wb.Sheets[sheet]['!ref'];
                                    var refLowerBound = parseInt(config.wb.Sheets[sheet]['!ref'].split(':')[1].match(/\d+/));
                                    if (rowNumber > refLowerBound) {
                                        config.wb.Sheets[sheet]['!ref'].replace(/\d+$/, rowNumber);
                                    }
                                    if (config.wb.Sheets[sheet][rangeLetter + rowNumber]) {
                                        config.wb.Sheets[sheet][rangeLetter + rowNumber]['v'] = val;
                                    } else {
                                        config.wb.Sheets[sheet][rangeLetter + rowNumber] = {
                                            t: "n",
                                            v: val,
                                            f: '',
                                            w: "0"
                                        };
                                    }
                                }
                            }
                        });
                    }
                }
                config.processWb();
    });
	
	/*onclick on the btn clear the workbook will clear the area with the html-ised workbook*/
	
    $('#clear_wb').on('click', function(e) {
        $('.table-preview').html('');
    });
	
	/*------------------------------------------------------------------------------------------
		This Code Section is Reponsible for Adding Getters and |or Combinators to the Page: 
	--------------------------------------------------------------------------------------------*/
	   
	$('#add_getter').on('click', function(e){
		var clone = $('.getter:first').clone(true);
			clone.find('#add_getter').remove();
			clone.find('input.items_collection_name.success').removeClass('success');
			clone.find('input.items_collection_name').val('Items Collection Name');
			$('.combinator:first').before(clone);
	});
	
	$('#add_combinator').on('click', function(e) {
		var clone = $('.combinator:first').clone(true);
			clone.find('#add_combinator').remove();
			clone.find('textarea.combinator.success').removeClass('success');
			$('.complex.combinator:first').before(clone)
	});
	
	$('#complex_add_combinator').on('click', function(e) {
		var clone = $('.complex.combinator:first').clone(true);
			clone.find('#complex_add_combinator').remove();
			clone.find('.success').removeClass('success');
			$('.simple-writeable_iterator:first').before(clone);
	});
	
	$('#add_simple_writable_iterator').on('click', function(e) {
		var clone = $('.simple-writeable_iterator:first').clone(true);
			clone.find('#add_simple_writable_iterator').remove();
			clone.find('.success').removeClass('success');
			$('.complex-writeable_iterator').before(clone);
	});
	
	$('#add_complex_writable_iterator').on('click', function(e){
		var clone = $('.complex-writeable_iterator:first').clone(true);
			clone.find('#add_complex_writable_iterator').remove();
			clone.find('.success').removeClass('success');
			$('.writable').before(clone);
	});
	
	/*-------------------------------------------------------------
			This Code Section is Reponsible for Getters: 
	---------------------------------------------------------------*/
	
	$('.go-getter').on('click', function(e) {
		if($('#drag-and-drop.success').length){
			var itemsCollectionName = $(e.target).closest('.getter').find('input').val().replace(/\s+/g, '');
			var paramsArray = $(e.target).closest('.getter').find('textarea.getter').val().replace('[', '').replace(']', '').split(', ');
			var WorkbookName = paramsArray[0];
			var ColumnName = paramsArray[1];
			config[itemsCollectionName] = config.getItemNamesByColumn(WorkbookName, ColumnName);
			$(e.target).closest('.getter').find('input').addClass('success');
			$(e.target).closest('.getter').find('input').val($(e.target).closest('.getter').find('input').val() + "-> length:" + config[itemsCollectionName].length)
			$('#drag-and-drop.success').removeClass('failure');
		}else{
			$(e.target).closest('.getter').find('input').addClass('failure');
		}
	});
	
	$('button.complex.combinator-do').on('click', function(e) {
		if($('#drag-and-drop.success').length) {
			var Array1 = $(e.target).closest('.complex.combinator').find('input.what1').val();
				Array1 = Array1.replace(/[\[|\]]/g, '');
				if(config[Array1]){
					$(e.target).closest('.complex.combinator').find('input.what1').addClass('success');
				}else{
					$(e.target).closest('.complex.combinator').find('input.what1').addClass('failure');
				}
			var Array2 = $(e.target).closest('.complex.combinator').find('input.what2').val();
				Array2 = Array2.replace(/[\[|\]]/g, '');
				if(config[Array2]){
					$(e.target).closest('.complex.combinator').find('input.what2').addClass('success');
				}else{
					$(e.target).closest('.complex.combinator').find('input.what2').addClass('failure');
				}
			var newArrayName = $(e.target).closest('.complex.combinator').find('input.aux_array').val();
				newArrayName = newArrayName.replace(/[\[|\]]/g, '');
				config[newArrayName] = [];
			config.complexCombine(config[newArrayName], config[Array1], config[Array2]);
			if(config[newArrayName].length == config[Array2].length * config[Array1].length) {
				$(e.target).closest('.complex.combinator').find('input.aux_array').addClass('success');
			} else {
				$(e.target).closest('.complex.combinator').find('input.aux_array').addClass('failure');
			}
			$('#drag-and-drop.success').removeClass('failure');
		}
	});
	
	$('.combinator-do:not(.complex)').on('click', function(e) {
		if($('#drag-and-drop.success').length) {
			var val = $(e.target).closest('.combinator').find('textarea.combinator').val();
			val = $(e.target).closest('.combinator').find('textarea.combinator').val().replace(/[\[|\]]/g, '');
			val = val.split(', ');
			var newArrayName = val[0];
			config[newArrayName] = [];
			var Array1 = val[1];
			var Array2 = val[2];
			config.combine2Arrays(newArrayName, config[Array1], config[Array2]);
			$(e.target).closest('.combinator').find('textarea.combinator').addClass('success');
			$(e.target).closest('.combinator').find('textarea.combinator').val($(e.target).closest('.combinator').find('textarea.combinator').val() + "length: " + config[newArrayName].length);
			$('#drag-and-drop.success').removeClass('failure');
		}else {
			$(e.target).closest('.combinator').find('textarea.combinator').addClass('success');
		}
	});
	
	/*-------------------------------------------------------------
			This Code Section is Reponsible for Writing: 
	---------------------------------------------------------------*/
	/*simple writable iterator do:*/
	$('.simple_write_do').on('click', function(e) {
		if($('#drag-and-drop.success').length) {
			debugger;
			var what = $(e.target).closest('.simple-writeable_iterator').find('input.what').val();
				what = what.replace(/[[|\]]/g, '');
			var destinationWorkbookName;
			var theWorksheet;
			//writableArrayRange_theWorksheet
			//writableArrayRange_theColumn
			var theColumn;
			var theRange;
			if(what && what !== '[Source Items Array Name]') {
				$(e.target).closest('.simple-writeable_iterator').find('input.what').addClass('success');
				theRange = $(e.target).closest('.simple-writeable_iterator').find('textarea.write_what_where').val();
				if(theRange && theRange !== '[Worksheet Name, Targeted Range [Worksheet Name, Column Name], Array Name]'){
					$(e.target).closest('.simple-writeable_iterator').find('textarea.write_what_where').addClass('success');
					theRange = theRange.replace(/^[[]/, '');
					theRange = theRange.replace(/[\]]$/, '');
					theRange = theRange.split(', ');
					destinationWorkbookName = theRange[0];
					theRange.splice(0, 1);
					theRange.splice(2, 1);
					theWorksheet = theRange[0].match(/[[](.*)/)[1];
					theColumn = theRange[1].replace(']', '');
					theRange = config.rangeSeeker(theWorksheet, theColumn);
				} else{
					$(e.target).closest('.simple-writeable_iterator').find('textarea.write_what_where').addClass('failure');
				}
			}else{
				$(e.target).closest('.simple-writeable_iterator').find('textarea.what').addClass('failure');
			}
			config.writeable(destinationWorkbookName, theRange, config[what]);
			theRange = config.rangeIncrementer(theRange, config[what].length+1);
			$('#drag-and-drop.success').removeClass('failure');
		}
	});
	
	/*complex writable iterator do:*/
	$('.iterate_to_write_do').on('click', function(e) {
		if($('#drag-and-drop.success').length) {
			debugger;
			var iterateTheWhat = $(e.target).closest('.complex-writeable_iterator').find('textarea.what').val();
				iterateTheWhat = iterateTheWhat.replace(/[[|\]]/g, '');
			var auxArrName = $(e.target).closest('.complex-writeable_iterator').find('.aux_array').val();
			var auxArrNameLength;
			var write_auxiliary_array_where;
			var destinationWorkbookName;
			var writeableArrayName;
			var writableArrayRange_theWorksheet;
			var writableArrayRange_theColumn;
			var theRange;
			if(auxArrName && auxArrName !== 'Auxiliary Array Name'){
				$(e.target).closest('.complex-writeable_iterator').find('.aux_array').addClass('success');
				auxArrNameLength = $(e.target).closest('.complex-writeable_iterator').find('.items_array_length').val();
				auxArrNameLength = auxArrNameLength.replace('.length', '');	
				write_auxiliary_array_where = $(e.target).closest('.complex-writeable_iterator').find('textarea.write_auxiliary_array').val();	
				if(write_auxiliary_array_where && write_auxiliary_array_where !== '[Worksheet Name, Targeted Range [Worksheet Name, Column Name], Array Name]'){
					$(e.target).closest('.complex-writeable_iterator').find('textarea.write_auxiliary_array').addClass('success');
					write_auxiliary_array_where = write_auxiliary_array_where.replace(/^[[]/, '');
					write_auxiliary_array_where = write_auxiliary_array_where.replace(/[\]]$/, '');
					write_auxiliary_array_where= write_auxiliary_array_where.split(', ');
					destinationWorkbookName = write_auxiliary_array_where[0];	
					writeableArrayName = write_auxiliary_array_where[3];
					write_auxiliary_array_where.splice(0, 1);
					write_auxiliary_array_where.splice(2, 1);
					writableArrayRange_theWorksheet = write_auxiliary_array_where[0].match(/[[](.*)/)[1];
					writableArrayRange_theColumn = write_auxiliary_array_where[1].replace(']', '');
					theRange = config.rangeSeeker(writableArrayRange_theWorksheet, writableArrayRange_theColumn);
				}else{
					$(e.target).closest('.complex-writeable_iterator').find('textarea.write_auxiliary_array').addClass('failure');
				}
			}else{
				$(e.target).closest('.complex-writeable_iterator').find('.aux_array').addClass('failure');
			}
			config[iterateTheWhat].forEach(function(theWhat) {
				var writeableArray = config.iterate_to_write_subRoutine(theWhat, auxArrName, config[auxArrNameLength].length);
				config.writeable(destinationWorkbookName, theRange, writeableArray);
				theRange = config.rangeIncrementer(theRange, writeableArray.length+1);
			});
			
			$('#drag-and-drop.success').removeClass('failure');
		}else {
			$(e.target).closest('.combinator').find('textarea.combinator').addClass('success');
		}
	});
});