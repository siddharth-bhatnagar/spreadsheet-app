// Scrollbar
const ps = new PerfectScrollbar("#cells", {
    wheelSpeed: 10,
    wheelPropagation: true,
    swipeEasing: true
});

// Sheet Column Title and Row Number
for (let i = 1; i <= 100; i++) {

    let columnNumber = i;
    let columnTitle = "";

    while (columnNumber != 0) {
        columnNumber--;
        let remainder = columnNumber % 26
        columnTitle = String.fromCharCode(65 + remainder) + columnTitle;
        columnNumber = Math.floor(columnNumber / 26);
    }
    $("#columns").append(`<div class="column-name">${columnTitle}</div>`);
    $("#rows").append(`<div class="row-name">${i}</div>`);
}

//cell data init
// new structure, only store values of those cells whose default properties have been changed
/*
cellData = {
    "Sheet1": {
        7: {
            1: {...defaultProperties},
            2: {...defaulProperties}
        },
        2: {
            5: {...}
        }
    },
    "Sheet2": {
        ...
    }
}
*/

let cellData = {
    "Sheet1": {}
};

// Keeping track of sheets
let selectedSheet = "Sheet1";
let totalSheets = 1;
let lastAddedSheet = 1;

// The default properties every cell has, applied in css
let defaultProperties = {
    "font-family": "Times New Roman",
    "font-size": "18",
    "text": "",
    "bold": false,
    "italic": false,
    "underlined": false,
    "alignment": "left",
    "color": "#000000",
    "bgcolor": "#ffffff"
};


// Generating cells
for (let i = 1; i <= 100; i++) {
    let row = $('<div class="cell-row"></div>');
    for (let j = 1; j <= 100; j++) {
        row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
    }
    $("#cells").append(row);
}

// Row Number & Column Titles scroll with cell scroll
$("#cells").scroll(function (e) {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
});

// get row and col number of current cell
const getRowColumn = (cell) => {
    let id = $(cell).attr("id");
    let idArray = id.split("-");
    let rowID = parseInt(idArray[1]);
    let colID = parseInt(idArray[3]);

    return [rowID, colID];
}

// get row and col number of adjacent cells
const getAdjacentCells = (row, col) => {
    let top = $(`#row-${row - 1}-col-${col}`);
    let left = $(`#row-${row}-col-${col - 1}`);
    let down = $(`#row-${row + 1}-col-${col}`);
    let right = $(`#row-${row}-col-${col + 1}`);

    return [top, left, down, right];
}


// Double clicking on a cell makes it editable and focuses on that
$(".input-cell").dblclick(function (e) {
    // Removes all selected cells first and selects current cell
    $(".input-cell.selected").removeClass("selected top-selected down-selected left-selected right-selected");
    
    $(this).addClass("selected");
    $(this).attr("contenteditable", "true");
    $(this).focus();
});

// Upon clicking elsewhere
$(".input-cell").blur(function (e) {
    // making the cells uneditable
    $(this).attr("contenteditable", "false");
    
    // storing text in cellData object
    updateCellData("text", $(this).text());
});

// Upon single click
$(".input-cell").click(function (e) {
    
    // getting row number and column number of current cell
    const [rowID, colID] = getRowColumn(this);
    
    // getting neighbouring cells wrapped in jquery object
    const [top, left, down, right] = getAdjacentCells(rowID, colID);
    
    // checking if currently clicked cell has selected class and control key is pressed 
    if ($(this).hasClass("selected") && e.ctrlKey) {
    
        // then unselect the current cell and remove classes from adjacent cells 
        unselectCell(this, e, top, left, down, right);
    }
    else {

        // select the current cell and if adjacent cells are selected, add relevant classes to them
        selectCell(this, e, top, left, down, right);
    }
});

// Handles selection of current cell
function selectCell(element, event, top, left, down, right) {
    // If control key was pressed simultaneously
    if (event.ctrlKey) {
        // if top cell exists
        let topSelect;
        if (top) {
            topSelect = top.hasClass("selected");
        }

        // if down cell exists
        let downSelect;
        if (down) {
            downSelect = down.hasClass("selected");
        }

        // if left cell exists
        let leftSelect;
        if (left) {
            leftSelect = left.hasClass("selected");
        }

        // if right cell exists
        let rightSelect;
        if (right) {
            rightSelect = right.hasClass("selected");
        }

        // if top cell is selected
        if (topSelect) {
            top.addClass("down-selected");
            $(element).addClass("top-selected");
        }

        // if left cell is selected
        if (leftSelect) {
            left.addClass("right-selected");
            $(element).addClass("left-selected");
        }

        // if down cell is selected
        if (downSelect) {
            down.addClass("top-selected");
            $(element).addClass("down-selected");
        }

        // if right cell is selected
        if (rightSelect) {
            right.addClass("left-selected");
            $(element).addClass("right-selected");
        }

    } else {
        // If control key was not pressed then remove classes from previous selections
        $(".input-cell.selected").removeClass("selected top-selected down-selected left-selected right-selected");
    }
    // Adding selected class to current selection
    $(element).addClass("selected");

    // Updating the header based on data in cellData object
    changeHeader(getRowColumn(element));
}

// Handles cell property display on menu-icon bar depending upon property of that cell in cellData object
function changeHeader([rowID, colID]) {

    let data;
    // if cell exists in cellData
    if (cellData[selectedSheet][rowID - 1] && cellData[selectedSheet][rowID - 1][colID - 1]) {
        data = cellData[selectedSheet][rowID - 1][colID - 1];
    }
    else {
        // if that row and column didn't exist in cellData
        data = defaultProperties;
    }

    // deselecting the previously selected icon
    $(".alignment.selected").removeClass("selected");
    // selecting the matching alignment icon
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");

    // Updating the font style icon selections
    updateFontStyleHeader(data, "bold");
    updateFontStyleHeader(data, "italic");
    updateFontStyleHeader(data, "underlined");

    // Matching the cell color and text color icon div selections to current cell's props
    $("#fill-color").css("border-bottom", `4px solid ${data.bgcolor}`);
    $("#text-color").css("border-bottom", `4px solid ${data.color}`);

    // Updates the value of font-size and family dropdowns
    $("#font-family").val(data["font-family"]);
    $("#font-size").val(data["font-size"]);

    // changing the font of selected dropdown item
    $("#font-family").css("font-family", data["font-family"]);
}

// Handles font-style icon changes based on currently selected property
function updateFontStyleHeader(data, property) {
    // if the value of property in data object is true
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    }
    else {
        $(`#${property}`).removeClass("selected");
    }
}

// Handles unselection of cells
function unselectCell(element, event, top, left, down, right) {
    // if the user was not writing data in cell, then only remove classes
    if ($(element).attr("contenteditable") == "false") {
        if ($(element).hasClass("top-selected")) {
            top.removeClass("down-selected");
        }

        if ($(element).hasClass("left-selected")) {
            left.removeClass("right-selected");
        }

        if ($(element).hasClass("down-selected")) {
            down.removeClass("top-selected");
        }

        if ($(element).hasClass("right-selected")) {
            right.removeClass("left-selected");
        }

        $(element).removeClass("selected top-selected down-selected left-selected right-selected");
    }
}

// For selecting cells with mouse click and drag
let startCellSelected = false;
let startCell = {};
let endCell = {};
let scrollXRStarted = false;
let scrollXLStarted = false;
// let scrollYTStarted = false;
// let scrollYDStarted = false;

// only used for scrolling/selecting start cell
$(".input-cell").mousemove(function (e) {

    if (e.buttons == 1) {
        e.preventDefault();
        if (e.pageX > ($(window).width() - 10) && !scrollXRStarted) {
            scrollXR();
        }
        else if (e.pageX < 10 && !scrollXLStarted) {
            scrollXL();
        }

        // if(e.pageY > ($(window).height() - 30) && !scrollYDStarted) {
        //     scrollYD();
        // } 
        // else if(e.pageY > 222 && e.pageY < 250 && !scrollYTStarted) {
        //     scrollYT();
        // }

        if (!startCellSelected) {
            let [rowID, colID] = getRowColumn(this);
            startCell = { "rowID": rowID, "colID": colID };
            startCellSelected = true;
            selectSubMatrix(startCell, startCell);
        }
    }
    else {
        startCellSelected = false;
    }
});

// selects end cell and stops scrolling
$(".input-cell").mouseenter(function (e) {
    if (e.buttons == 1) {
        if (e.pageX < ($(window).width() - 10) && scrollXRStarted) {
            clearInterval(scrollXRInterval);
            scrollXRStarted = false;
        }

        if (e.pageX > 10 && scrollXLStarted) {
            clearInterval(scrollXLInterval);
            scrollXLStarted = false;
        }

        // if(e.pageY > 250 || e.pageY < 222 && scrollYTStarted) {
        //     clearInterval(scrollYTInterval);
        //     scrollYTStarted = false;
        // }

        // if(e.pageY < $(window).height() - 30 && scrollYDStarted) {
        //     clearInterval(scrollYDInterval);
        //     scrollYDStarted = false;
        // }

        let [rowID, colID] = getRowColumn(this);
        endCell = { "rowID": rowID, "colID": colID };
        selectSubMatrix(startCell, endCell);
    }
});

// renders selection via mouse left btn click + drag
function selectSubMatrix(start, end) {
    $(".input-cell.selected").removeClass("selected top-selected down-selected left-selected right-selected");
    for (let i = Math.min(start.rowID, end.rowID); i <= Math.max(start.rowID, end.rowID); i++) {
        for (let j = Math.min(start.colID, end.colID); j <= Math.max(start.colID, end.colID); j++) {
            let [top, left, down, right] = getAdjacentCells(i, j);
            let currentCell = $(`#row-${i}-col-${j}`)[0];
            selectCell(currentCell, { ctrlKey: true }, top, left, down, right);
        }
    }
}

// selection + scrolling
// horizontal right
let scrollXRInterval;
function scrollXR() {
    scrollXRStarted = true;
    scrollXRInterval = setInterval(() => {
        $("#cells").scrollLeft($("#cells").scrollLeft() + 100);
    }, 100);
}

// horizontal left
let scrollXLInterval;
function scrollXL() {
    scrollXLStarted = true;
    scrollXLInterval = setInterval(() => {
        $("#cells").scrollLeft($("#cells").scrollLeft() - 100);
    }, 100);
}

// vertical up
// let scrollYTInterval;
// function scrollYT() {
//     scrollYTInterval = setInterval(() => {
//         $("#cells").scrollTop($("#cells").scrollTop() - 50);
//     }, 100);
// }

// // vertical down
// let scrollYDInterval;
// function scrollYD() {
//     scrollYDInterval = setInterval(() => {
//         $("#cells").scrollTop($("#cells").scrollTop() + 50);
//     }, 100);
// }

// handles selection + scrolling
$(".data-container").mousemove(function (e) {
    if (e.buttons == 1) {
        e.preventDefault();
        if (e.pageX > ($(window).width() - 10) && !scrollXRStarted) {
            scrollXR();
        }
        else if (e.pageX < 10 && !scrollXLStarted) {
            scrollXL();
        }

        // if(e.pageY > ($(window).height() - 30) && !scrollYDStarted) {
        //     scrollYD();
        // } 
        // else if(e.pageY > 222 && e.pageY < 250 && !scrollYTStarted) {
        //     scrollYT();
        // }
    }
});

// stops scrolling
$(".data-container").mouseup(function (e) {
    clearInterval(scrollXRInterval);
    clearInterval(scrollXLInterval);
    scrollXLStarted = false;
    scrollXRStarted = false;
    // clearInterval(scrollYTInterval);
    // clearInterval(scrollYDInterval);
    // scrollYTStarted = false;
    // scrollYDStarted = false;
    // console.log("mouseup");
});

// handling text-alignment
$(".alignment").click(function (e) {
    let alignment = $(this).attr("data-type");
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    $(".input-cell.selected").css("text-align", alignment);
    updateCellData("alignment", alignment);
});

// handling text styles
$("#bold").click(function (e) {
    setStyle(this, "bold", "font-weight", "bold");
});

$("#italic").click(function (e) {
    setStyle(this, "italic", "font-style", "italic");
});

$("#underlined").click(function (e) {
    setStyle(this, "underlined", "text-decoration", "underline");
});

// Handles click on text-style icons
function setStyle(element, property, key, value) {
    if ($(element).hasClass("selected")) {
        $(element).removeClass("selected");
        $(".input-cell.selected").css(key, "");
        updateCellData(property, false);
    }
    else {
        $(element).addClass("selected");
        $(".input-cell.selected").css(key, value);
        updateCellData(property, true);
    }
}

// renders color picker jquery
// onColorSelected is invoked twice on page load
$(".pick-color").colorPick({
    'initialColor': "#ABCD",
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': false,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function () {
        if (this.color != "#ABCD") {
            if ($(this.element.children()[1]).attr("id") == "fill-color") {
                $(".input-cell.selected").css("background-color", this.color);
                $("#fill-color").css("border-bottom", `4px solid ${this.color}`);
                updateCellData("bgcolor", this.color);
            }

            if ($(this.element.children()[1]).attr("id") == "text-color") {
                $(".input-cell.selected").css("color", this.color);
                $("#text-color").css("border-bottom", `4px solid ${this.color}`);
                updateCellData("color", this.color);
            }
        }
    }
});

// Handles click on image
// Clicks the div which has color palette
$("#fill-color").click(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

$("#text-color").click(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});

// Handles change in font-size/font-family
$(".menu-selector").change(function (e) {
    let value = $(this).val();
    let key = $(this).attr("id");

    if (key == "font-family") {
        $("#font-family").css(key, value);
    }

    if (!isNaN(value)) {
        value = parseInt(value);
    }

    $(".input-cell.selected").css(key, value);
    updateCellData(key, value);
});

// writing data to the cellData object
function updateCellData(property, value) {
    // if the value is not equal to default value
    if (value != defaultProperties[property]) {
        // iterate over each selected cell
        $(".input-cell.selected").each(function (index, data) {
            let [rowID, colID] = getRowColumn(data);
            // checking if row exists in DB or not
            // if it exists, that means some cell in that row was updated
            // if it doesnt, that means first time changes
            if (cellData[selectedSheet][rowID - 1] == undefined) {
                cellData[selectedSheet][rowID - 1] = {};
                cellData[selectedSheet][rowID - 1][colID - 1] = { ...defaultProperties };
                cellData[selectedSheet][rowID - 1][colID - 1][property] = value;
            }
            else {
                // checking if that particular cell in row exists in DB or not
                if (cellData[selectedSheet][rowID - 1][colID - 1] == undefined) {
                    cellData[selectedSheet][rowID - 1][colID - 1] = { ...defaultProperties };
                    cellData[selectedSheet][rowID - 1][colID - 1][property] = value;
                }
                else {
                    cellData[selectedSheet][rowID - 1][colID - 1][property] = value;
                }
            }
        });
    }
    // The value is being changed back to the default value
    else {
        $(".input-cell.selected").each(function (index, data) {
            let [rowID, colID] = getRowColumn(data);
            // checking if column exists, if it doesn't, that means default props are already set
            if (cellData[selectedSheet][rowID - 1] && cellData[selectedSheet][rowID - 1][colID - 1]) {
                cellData[selectedSheet][rowID - 1][colID - 1][property] = value;
                //    checking if the current object has become equal to default object
                if (JSON.stringify(cellData[selectedSheet][rowID - 1][colID - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][rowID - 1][colID - 1];
                    if (Object.keys(cellData[selectedSheet][rowID - 1]).length == 0) {
                        delete cellData[selectedSheet][rowID - 1];
                    }
                }
            }
        });
    }
}

// Handles click on add sheet button
function addSheetEvents() {
    $(".sheets-tab.selected").on("contextmenu", function (e) {
        e.preventDefault();
        $(".sheet-options-modal").remove();
        let modal = $(`<div class="sheet-options-modal">
                        <div class="option sheet-rename">Rename</div>
                        <div class="option sheet-delete">Delete</div>
                       </div>`);

        modal.css({ "left": e.pageX });
        $(".app-container").append(modal);
    });

    $(".sheets-tab.selected").click(function (e) {
        $(".sheet.selected").removeClass("selected");
        $(this).addClass("selected");
        selectSheet();
    });
}

addSheetEvents();

// removes previosly opened modals
// creates new modal
// appends that modal to the main app container
// happens on right click of the mouse
$(".sheets-tab").on("contextmenu", function (e) {
    e.preventDefault();
    $(".sheet-options-modal").remove();
    let modal = $(`<div class="sheet-options-modal">
                    <div class="option sheet-rename">Rename</div>
                    <div class="option sheet-delete">Delete</div>
                   </div>`);

    modal.css({ "left": e.pageX });
    $(".app-container").append(modal);
});

// adding a blank sheet
$(".add-sheet").click(function (e) {
    lastAddedSheet += 1;
    totalSheets += 1;
    cellData[`Sheet ${lastAddedSheet}`] = {};
    $(".sheet.selected").removeClass("selected");
    $(".sheets-tab").append(`<div class="sheet selected">Sheet ${lastAddedSheet}</div>`);
    selectSheet();
    addSheetEvents();
});

// Upon selecting a different sheet, this function is triggered
$(".sheets-tab").click(function (e) {
    $(".sheet.selected").removeClass("selected");
    $(this).addClass("selected");
    selectSheet();
});

function selectSheet() {
    emptyPreviousSheet();
    selectedSheet = $(".sheet.selected").text();
    loadCurrentSheet();
}

// traverse on cellData object and get the keys of rows and columns where changes were made
// Then change their css properties to all the default properties
function emptyPreviousSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let rowKey of rowKeys) {
        let columnKeys = Object.keys(data[rowKey]);
        let rowID = parseInt(rowKey);
        for (let columnKey of columnKeys) {
            let colID = parseInt(columnKey);
            let cell = $(`#row-${rowID + 1}-col-${colID + 1}`);
            cell.text("");
            cell.css({
                "font-family": "Times New Roman",
                "font-size": "18px",
                "font-weight": "normal",
                "font-style": "normal",
                "text-decoration": "none",
                "text-align": "left",
                "color": "#000000",
                "background-color": "#ffffff"
            });
        }
    }
}

// load the currently selected sheet
function loadCurrentSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let rowKey of rowKeys) {
        let columnKeys = Object.keys(data[rowKey]);
        let rowID = parseInt(rowKey);
        for (let columnKey of columnKeys) {
            let colID = parseInt(columnKey);
            let cell = $(`#row-${rowID + 1}-col-${colID + 1}`);
            cell.text(data[rowID][colID].text);
            cell.css({
                "font-family": data[rowID][colID]["font-family"],
                "font-size": data[rowID][colID]["font-size"],
                "font-weight": data[rowID][colID]["bold"] ? "bold" : "normal",
                "font-style": data[rowID][colID]["italic"] ? "italic" : "normal",
                "text-decoration": data[rowID][colID]["underlined"] ? "underline" : "none",
                "text-align": data[rowID][colID]["alignment"],
                "color": data[rowID][colID]["color"],
                "background-color": data[rowID][colID]["bgcolor"]
            });
        }
    }
}

