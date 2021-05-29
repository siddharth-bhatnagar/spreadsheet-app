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

// Generating cells and cell data init

let cellData = {
    "Sheet1": {}
};

let selectedSheet = "Sheet1";
let totalSheets = 1;

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

$(".input-cell").dblclick(function (e) {
    $(".input-cell.selected").removeClass("selected top-selected down-selected left-selected right-selected");
    $(this).addClass("selected");
    $(this).attr("contenteditable", "true");
    $(this).focus();
});

$(".input-cell").blur(function (e) {
    $(this).attr("contenteditable", "false");
    // storing text in cell data
    updateCellData("text", $(this).text());
});

$(".input-cell").click(function (e) {
    const [rowID, colID] = getRowColumn(this);
    const [top, left, down, right] = getAdjacentCells(rowID, colID);
    if ($(this).hasClass("selected") && e.ctrlKey) {
        unselectCell(this, e, top, left, down, right);
    }
    else {
        selectCell(this, e, top, left, down, right);
    }
});

function selectCell(element, event, top, left, down, right) {
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
        $(".input-cell.selected").removeClass("selected top-selected down-selected left-selected right-selected");
    }
    $(element).addClass("selected");
    changeHeader(getRowColumn(element));
}

// making changes two way
function changeHeader([rowID, colID]) {
    let data;
    if (cellData[selectedSheet][rowID - 1] && cellData[selectedSheet][rowID - 1][colID - 1]) {
        data = cellData[selectedSheet][rowID - 1][colID - 1];
    }
    else {
        data = defaultProperties;
    }
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");

    updateFontStyleHeader(data, "bold");
    updateFontStyleHeader(data, "italic");
    updateFontStyleHeader(data, "underlined");

    // changing the icon bar color
    $("#fill-color").css("border-bottom", `4px solid ${data.bgcolor}`);
    $("#text-color").css("border-bottom", `4px solid ${data.color}`);

    // changing the value in spinner
    $("#font-family").val(data["font-family"]);
    $("#font-size").val(data["font-size"]);
    // changing the font-family of spinner
    $("#font-family").css("font-family", data["font-family"]);
}

function updateFontStyleHeader(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    }
    else {
        $(`#${property}`).removeClass("selected");
    }
}

function unselectCell(element, event, top, left, down, right) {
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

// renders selection
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

// clicks on the main div
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

    modal.css({"left": e.pageX});
    $(".app-container").append(modal);
});