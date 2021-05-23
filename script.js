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

$(".input-cell").dblclick(function (e) {
    $(".input-cell.selected").removeClass("selected top-selected down-selected left-selected right-selected");
    $(this).attr("contenteditable", "true");
    $(this).focus();
});

$(".input-cell").blur(function (e) {
    $(this).attr("contenteditable", "false");
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




