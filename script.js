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
    $("#columns").append(`<div class="column-name column-${i}" id="${columnTitle}">${columnTitle}</div>`);
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
let save = true;

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
    "bgcolor": "#ffffff",
    "upStream": [],
    "downStream": []
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
            $(".input-cell.selected").attr("contenteditable", "false");
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
    let newCellData = cellData;
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
                cellData[selectedSheet][rowID - 1][colID - 1] = { ...defaultProperties, "upStream": [], "downStream": [] };
                cellData[selectedSheet][rowID - 1][colID - 1][property] = value;
            }
            else {
                // checking if that particular cell in row exists in DB or not
                if (cellData[selectedSheet][rowID - 1][colID - 1] == undefined) {
                    cellData[selectedSheet][rowID - 1][colID - 1] = { ...defaultProperties, "upStream": [], "downStream": [] };
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
                    // Deleting cell column when it becomes default
                    delete cellData[selectedSheet][rowID - 1][colID - 1];
                    // Deleting row if empty
                    if (Object.keys(cellData[selectedSheet][rowID - 1]).length == 0) {
                        delete cellData[selectedSheet][rowID - 1];
                    }
                }
            }
        });
    }

    if (save && JSON.stringify(newCellData) != JSON.stringify(cellData)) {
        save = false;
    }
}

// clicking anywhere on screen closes the sheet-options-modal
$(".app-container").click(function (e) {
    $(".sheet-options-modal").remove();
    // $(".file-options-modal").remove();
})

// Attaching event listeners to newly added sheet-tabs
function addSheetEvents() {
    // On adding new sheet, it gets selected
    // Adding event listener only on the selected sheet
    $(".sheet-tab.selected").on("contextmenu", function (e) {
        // prevents opening of context menu on right mouse btn click
        e.preventDefault();
        // if the clicked upon sheet is not selected
        if ($(this).hasClass("selected") == false) {
            // removing selection from current sheet
            $(".sheet-tab.selected").removeClass("selected");
            // Selecting the clicked upon sheet
            $(this).addClass("selected");
            selectSheet();
        }
        // Removing any previously opened modals
        $(".sheet-options-modal").remove();
        // Creating a new modal
        let modal = $(`<div class="sheet-options-modal">
                        <div class="option sheet-rename">Rename</div>
                        <div class="option sheet-delete">Delete</div>
                       </div>`);

        // x-coordinate of modal in UI
        modal.css({ "left": e.pageX });
        $(".app-container").append(modal);

        // Handles Rename click event
        $(".sheet-rename").click(function (e) {
            // Creating a new modal
            let renameModal = $(`<div class="sheet-modal-parent">
                                    <div class="sheet-rename-modal">
                                        <div class="sheet-modal-title">Rename Sheet</div>
                                        <div class="sheet-modal-input-container">
                                            <span class="sheet-modal-input-title">Rename Sheet To: </span>
                                            <input type="text" class="sheet-modal-input" />
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button yes-btn">OK</div>
                                            <div class="button no-btn">Cancel</div>
                                        </div>
                                    </div>
                                </div>`);

            $(".app-container").append(renameModal);

            // Cancel button closes the Modal
            $(".no-btn").click(function (e) {
                $(".sheet-modal-parent").remove();
            });

            // OK button calls renameSheet
            $('.yes-btn').click(function (e) {
                renameSheet();
            });

            // Enter press calls renameSheet
            $(".sheet-modal-input").keypress(function (e) {
                if (e.key == "Enter") {
                    renameSheet();
                }
            })
        });

        // Handles Delete option click
        $(".sheet-delete").click(function (e) {
            // Checking if it's possible to delete sheet
            if (totalSheets > 1) {
                // Creating a new modal
                let deleteModal = $(`<div class="sheet-modal-parent">
                                    <div class="sheet-delete-modal">
                                        <div class="sheet-modal-title">${selectedSheet}</div>
                                        <div class="sheet-modal-detail-container">
                                            <span class="sheet-modal-detail-title">Are you sure?</span>
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button yes-btn">
                                                <div class="material-icons delete-icon">delete</div>
                                                Delete
                                            </div>
                                            <div class="button no-btn">Cancel</div>
                                        </div>
                                    </div>
                                </div>`);

                $(".app-container").append(deleteModal);

                //  Cancel btn click
                $(".no-btn").click(function (e) {
                    $(".sheet-modal-parent").remove();
                });

                // Invokes delete sheet
                $(".yes-btn").click(function (e) {
                    deleteSheet();
                });
            }
            // When there is single sheet
            else {
                let alertBox = $(`<div class="sheet-modal-parent">
                                    <div class="alert-box">
                                        <span class="alert-text">Sorry! It is not possible to delete your only sheet!</span>
                                    </div>
                                </div>`);
                $(".app-container").append(alertBox);
                setTimeout(() => {
                    $(".sheet-modal-parent").remove();
                }, 1500);
            }
        });
    });

    // Attaching click event to sheet tab when it was created
    $(".sheet-tab.selected").click(function (e) {
        // If the clicked upon sheet is not already selected
        if (!$(this).hasClass("selected")) {
            $(".sheet-tab.selected").removeClass("selected");
            $(this).addClass("selected");
            selectSheet();
        }
    });
}

// Attaching events to default sheet on page load/reload
addSheetEvents();

// Rename Sheet
function renameSheet() {
    // Taking the value of name
    let name = $(".sheet-modal-input").val();
    // If the entered value is not empty or already present
    if (name != "" && !Object.keys(cellData).includes(name)) {
        // Changing the text of sheet tab
        save = false;
        $(".sheet-tab.selected").text(name);
        // Updating the key in cellData
        // This appraoch creates a dummy object and we will replace cellData with it 
        // These steps preserve the relative postioning of the keys in original object
        let newCellData = {};
        // Traversing on keys array of cellData
        for (let sheet of Object.keys(cellData)) {
            if (sheet == selectedSheet) {
                newCellData[name] = cellData[sheet];
            } else {
                newCellData[sheet] = cellData[sheet];
            }
        }

        cellData = newCellData;
        selectedSheet = name;
        $(".sheet-modal-parent").remove();
    }
    else {
        // If the sheet name is invalid
        $(".rename-error").remove();
        $(".sheet-modal-input-container").append(`<div class="rename-error">Entered name is invalid or already exists!</div>`);
    }
}

// Handles Sheet deletion
function deleteSheet() {
    // Removing the modal
    save = false;
    $(".sheet-modal-parent").remove();
    // index of the sheet being deleted
    let index = Object.keys(cellData).indexOf(selectedSheet);
    // Storing reference to current sheet
    let currentSheet = $(".sheet-tab.selected");
    // index = 0, select next sheet
    if (index == 0) {
        let nextSheet = currentSheet.next()[0];
        if ($(nextSheet).hasClass("selected") == false) {
            $(".sheet-tab.selected").removeClass("selected");
            $(nextSheet).addClass("selected");
            selectSheet();
        }
    }
    // index > 0, select previous sheet
    else {
        let prevSheet = currentSheet.prev()[0];
        if ($(prevSheet).hasClass("selected") == false) {
            $(".sheet-tab.selected").removeClass("selected");
            $(prevSheet).addClass("selected");
            selectSheet();
        }
    }
    // Removing sheet from UI
    currentSheet.remove();
    // Deleting the sheet from database
    delete cellData[currentSheet.text()];
    // Decrementing total Sheet Count
    totalSheets--;
    // console.log(cellData);
}

// Adding a blank sheet
$(".add-sheet").click(function (e) {
    save = false;
    lastAddedSheet += 1;
    totalSheets += 1;
    cellData[`Sheet${lastAddedSheet}`] = {};
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheets-tab-container").append(`<div class="sheet-tab selected">Sheet${lastAddedSheet}</div>`);
    selectSheet();
    addSheetEvents();
    $(".sheet-tab.selected")[0].scrollIntoView();
});


// selects sheet
function selectSheet() {
    emptyPreviousSheet();
    // Getting the name/key of current sheet
    selectedSheet = $(".sheet-tab.selected").text();
    loadCurrentSheet();
    // Selects first cell
    $("#row-1-col-1").click();
}

// traverse on cellData object and get the keys of rows and columns where changes were made
// Then change their css properties to all the default properties
// Removes content of previous sheet from UI
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

// load the currently selected sheet on UI
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

// Left Arrow, Right Arrow click

$(".scroll-left, .scroll-right").click(function (e) {
    // Storing keys in array
    let keys = Object.keys(cellData);
    // Finding the index of selected sheet
    let selectedSheetIndex = keys.indexOf(selectedSheet);
    // Storing ref to the current sheet
    let currentSheet = $(".sheet-tab.selected");

    // Sheets other than first sheet
    if (selectedSheetIndex != 0 && $(this).text() == "arrow_left") {
        let prevSheet = currentSheet.prev()[0];
        if ($(prevSheet).hasClass("selected") == false) {
            $(".sheet-tab.selected").removeClass("selected");
            $(prevSheet).addClass("selected");
            selectSheet();
        }
    }
    // Sheets other than last sheet
    else if (selectedSheetIndex != (keys.length - 1) && $(this).text() == "arrow_right") {
        let nextSheet = currentSheet.next()[0];
        if ($(nextSheet).hasClass("selected") == false) {
            $(".sheet-tab.selected").removeClass("selected");
            $(nextSheet).addClass("selected");
            selectSheet();
        }
    }
    // Adding bracket notation 0 because scrollIntoView is a JS method, not in jQuery
    // Javascript methods work only on native DOM objects, not DOM objects wrapped in jQuery
    // jQuery objects are arrays of DOM elements
    $(".sheet-tab.selected")[0].scrollIntoView();
});

// Upon clicking the file option in menu bar
$("#menu-file").click(function (e) {
    // Creating a modal
    let fileModal = $(`<div class="file-modal">
                            <div class="file-options-modal">
                                <div class="close">
                                    <div class="material-icons close-icon">arrow_circle_down</div>
                                    <div>Close</div>
                                </div>
                                <div class="new">
                                    <div class="material-icons new-icon">insert_drive_file</div>
                                    <div>New</div> 
                                </div>
                                <div class="open">
                                    <div class="material-icons open-icon">folder_open</div>
                                    <div>Open</div>
                                </div>
                                <div class="save">
                                    <div class="material-icons save-icon">save</div>
                                    <div>Save</div>
                                </div>
                            </div>
                            <div class="file-recent-modal"></div>
                            <div class="file-transparent"></div>
                        </div>`);

    $(".app-container").append(fileModal);

    // Animate modal opening
    fileModal.animate({
        width: "100vw"
    }, 200);

    // Handle modal closing
    $(".close, .file-transparent, .new, .save, .open").click(function (e) {
        fileModal.animate({
            width: "0"
        }, 200);
        setTimeout(() => {
            fileModal.remove();
        }, 150);
    });

    // Handle click on "NEW" option
    $(".new").click(function () {
        // If file is already saved
        if (save) {
            newFile();
        }
        // if file is not saved, ask user
        else {
            let modal = $(`<div class="sheet-modal-parent">
                                <div class="sheet-delete-modal">
                                    <div class="sheet-modal-title">${$(".title-bar").text()}</div>
                                    <div class="sheet-modal-detail-container">
                                        <span class="sheet-modal-detail-title">Do you want to save changes?</span>
                                    </div>
                                    <div class="sheet-modal-confirmation">
                                        <div class="button yes-btn">Yes</div>
                                        <div class="button no-btn">No</div>
                                    </div>
                                </div>
                            </div>`);


            $(".app-container").append(modal);

            // Dont save, open new file
            $(".no-btn").click(function (e) {
                newFile();
                modal.remove();
            });

            // save, then open new file
            $(".yes-btn").click(function (e) {
                // calls save file function
                saveFile();
                newFile();
                modal.remove();
            })
        }
    });

    // Handles click on save option
    $(".save").click(function (e) {
        saveFile();
    });

    // Handles open option click
    $(".open").click(function (e) {
        openFile();
    });
});

function newFile() {
    // Empty the sheet
    emptyPreviousSheet();
    // reset DB
    cellData = {
        "Sheet1": {}
    };
    // Set values to default
    totalSheets = 1;
    lastAddedSheet = 1;
    selectedSheet = "Sheet1";
    save = true;
    // Remove sheet tabs
    $(".sheet-tab").remove();
    // Add single sheet tab
    $(".sheets-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`);
    // Attach event listeners on new sheet
    addSheetEvents();
    // Change the file title to default title
    $(".title-bar").text("Excel - Book");
    // select first cell
    $("#row-1-col-1").click();
}

function saveFile() {
    // Create modal
    let saveModal = $(`<div class="sheet-modal-parent">
                            <div class="sheet-rename-modal">
                                <div class="sheet-modal-title">Save File</div>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title">File Name: </span>
                                    <input type="text" class="sheet-modal-input" value="${$('.title-bar').text()}"/>
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="button yes-btn">Save</div>
                                    <div class="button no-btn">Cancel</div>
                                </div>
                            </div>
                        </div>`);

    $(".app-container").append(saveModal);

    // handles cancel
    $(".no-btn").click(function (e) {
        saveModal.remove();
    });

    // Handles changes
    $(".yes-btn").click(function (e) {
        // Changing the file name
        $(".title-bar").text($(".sheet-modal-input").val());

        // Downloading file
        let anchorTag = document.createElement("a");
        $(".app-container").append(anchorTag);
        // The url 'data:' is used to download files
        // application/json is MIME type for json format
        // After comma, we append the data which we want to download in our file
        // encodeURIComponenent handles character encoding so that file downloads correctly
        anchorTag.href = `data:application/json,${encodeURIComponent(JSON.stringify(cellData))}`;
        anchorTag.download = $(".title-bar").text() + ".json";
        anchorTag.click();
        anchorTag.remove();
        saveModal.remove();
        save = true;
    });
}

function openFile() {
    // creating modal
    let openModal = $(`<div class="sheet-modal-parent">
                                <div class="sheet-open-modal">
                                    <div class="material-icons close-open-file">close</div>
                                    <div class="sheet-modal-title">Open File</div>
                                    <div class="sheet-modal-input-container">
                                        <input type="file" class="file-open" accept="application/json" />
                                    </div>
                                </div>
                            </div>`);

    // loading modal on UI
    $(".app-container").append(openModal);

    // Handling close icon click
    $(".close-open-file").click(function (e) {
        openModal.remove();
    })

    // Storing ref to the file input
    let inputFile = $(`input[type="file"]`);
    // On change event/loading file
    inputFile.change(function (e) {
        // selecting the file from event
        let file = e.target.files[0];
        // Changing the title of file
        $(".title-bar").text(file.name.split(".json")[0]);
        // Creating FileReader() object
        let reader = new FileReader();
        // Reading the content of file as text
        reader.readAsText(file);
        // Fired as soon as file is read
        reader.onload = () => {
            // closing the modal
            openModal.remove();
            // Emptying the current sheet
            emptyPreviousSheet();
            // Removing the sheet tabs from sheets-tab-container
            $(".sheet-tab").remove();
            // Parsing the data and updating cellData Object
            cellData = JSON.parse(reader.result);

            let keys = Object.keys(cellData);
            // takes this value if its not updated below
            lastAddedSheet = 1;
            // Iterting over keys and adding new sheet tabs in UI
            for (let key of keys) {
                // last added sheet to correct value to avoid bug
                if (key.includes("Sheet")) {
                    // Array of number possibly
                    let splittedSheetArray = key.split("Sheet");
                    if (splittedSheetArray.length == 2 && !isNaN(splittedSheetArray[1])) {
                        lastAddedSheet = parseInt(splittedSheetArray[1]);
                    }
                }
                $(".sheets-tab-container").append(`<div class="sheet-tab selected">${key}</div>`);
            }
            // Attaching event listeners on new tabs
            addSheetEvents();
            // Removing the selected class from sheet tab
            $(".sheet-tab.selected").removeClass("selected");
            // Selecting the first sheet tab
            $($(".sheet-tab")[0]).addClass("selected");
            // Setting the properties
            save = true;
            selectedSheet = keys[0];
            totalSheets = keys.length;
            // Loading the first sheet on UI
            loadCurrentSheet();
            // Selecting cell one
            $("#row-1-col-1").click();
        };
    });
}


// clipboard object stores the data from startCell to all selected cells in cellData
// init clipboard
let clipboard = { startCell: [], cellData: {} };
// flag to check if cut is clicked
let contentCut = false;
$("#cut,#copy").click(function (e) {
    // If the button clicked was cut
    if ($(this).text() == "content_cut") {
        contentCut = true;
    }
    // new clipboard
    clipboard = { startCell: [], cellData: {} };

    // getting the coordinates of start cell
    clipboard.startCell = getRowColumn($(".input-cell.selected")[0]);
    // Traversing on each selected cell
    $(".input-cell.selected").each(function (index, data) {
        // row and col number of cell
        let [rowId, colId] = getRowColumn(data);
        // if row and column exists in DB
        if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
            // if row doesn't exist in DB
            if (!clipboard.cellData[rowId]) {
                clipboard.cellData[rowId] = {};
            }
            // Making a copy of cell
            clipboard.cellData[rowId][colId] = { ...cellData[selectedSheet][rowId - 1][colId - 1] };
        }
    });
    // console.log(clipboard);
});

$("#paste").click(function (e) {
    // Clear cells if true
    if (contentCut) {
        emptyPreviousSheet();
    }
    // rows in clipboard
    let rows = Object.keys(clipboard.cellData);
    for (let i of rows) {
        // columns in clipboard
        let cols = Object.keys(clipboard.cellData[i]);
        for (let j of cols) {
            // Each cell
            // Deleting from database if content cut was true
            if (contentCut) {
                // deleting cell
                delete cellData[selectedSheet][i - 1][j - 1];
                // deleting row if empty
                if (Object.keys(cellData[selectedSheet][i - 1]).length == 0) {
                    delete cellData[selectedSheet][i - 1];
                }
            }
        }
    }
    // coordinates of destination start cell
    let startCell = getRowColumn($(".input-cell.selected")[0]);
    // traversal
    for (let i of rows) {
        let cols = Object.keys(clipboard.cellData[i]);
        for (let j of cols) {
            let rowDistance = parseInt(i) - parseInt(clipboard.startCell[0]);
            let colDistance = parseInt(j) - parseInt(clipboard.startCell[1]);
            // if row doesn't exist in DB
            if (!cellData[selectedSheet][startCell[0] + rowDistance - 1]) {
                cellData[selectedSheet][startCell[0] + rowDistance - 1] = {};
            }
            // copying data in DB from clipboard
            cellData[selectedSheet][startCell[0] + rowDistance - 1][startCell[1] + colDistance - 1] = { ...clipboard.cellData[i][j] };
        }
    }
    // Sheet reload
    loadCurrentSheet();
    // disabling paste
    if (contentCut) {
        contentCut = false;
        clipboard = { startCell: [], cellData: {} };
    }
});

$("#formula-input").blur(function (e) {
    if ($(".input-cell.selected").length > 0) {
        let formula = $(this).text();
        let tempElements = formula.split(" ");
        let elements = [];
        for (let i of tempElements) {
            if (i.length >= 2) {
                i = i.replace("(", "");
                i = i.replace(")", "");
                if (!elements.includes(i)) {
                    elements.push(i);
                }
            }
        }
        $(".input-cell.selected").each(function (index, data) {
            if (updateStreams(data, elements, false)) {
                let [rowId, colId] = getRowColumn(data);
                cellData[selectedSheet][rowId - 1][colId - 1].formula = formula;
                let selfColCode = $(`.column-${colId}`).attr("id");
                evalFormula(selfColCode + rowId);
            } else {
                alert("Formula is not valid");
            }
        })
    } else {
        alert("!Please select a cell First");
    }
});

function updateStreams(ele, elements, update, oldUpstream) {
    let [rowId, colId] = getRowColumn(ele);
    let selfColCode = $(`.column-${colId}`).attr("id");
    if (elements.includes(selfColCode + rowId)) {
        return false;
    }
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        for (let i of downStream) {
            if (elements.includes(i)) {
                return false;
            }
        }
        for (let i of downStream) {
            let [calRowId, calColId] = codeToValue(i);
            console.log(updateStreams($(`#row-${calRowId}-col-${calColId}`)[0], elements, true, upStream));
        }
    }

    if (!cellData[selectedSheet][rowId - 1]) {
        cellData[selectedSheet][rowId - 1] = {};
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else if (!cellData[selectedSheet][rowId - 1][colId - 1]) {
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else {

        let upStream = [...cellData[selectedSheet][rowId - 1][colId - 1].upStream];
        if (update) {
            for (let i of oldUpstream) {
                let [calRowId, calColId] = codeToValue(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
                index = cellData[selectedSheet][rowId - 1][colId - 1].upStream.indexOf(i);
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.splice(index, 1);
            }
            for (let i of elements) {
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.push(i);
            }
        } else {
            for (let i of upStream) {
                let [calRowId, calColId] = codeToValue(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
            }
            cellData[selectedSheet][rowId - 1][colId - 1].upStream = [...elements];
        }
    }

    for (let i of elements) {
        let [calRowId, calColId] = codeToValue(i);
        if (!cellData[selectedSheet][calRowId - 1]) {
            cellData[selectedSheet][calRowId - 1] = {};
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        } else if (!cellData[selectedSheet][calRowId - 1][calColId - 1]) {
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        } else {
            cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.push(selfColCode + rowId);
        }
    }
    console.log(cellData);
    return true;

}

function codeToValue(code) {
    let colCode = "";
    let rowCode = "";
    for (let i = 0; i < code.length; i++) {
        if (!isNaN(code.charAt(i))) {
            rowCode += code.charAt(i);
        } else {
            colCode += code.charAt(i);
        }
    }
    let colId = parseInt($(`#${colCode}`).attr("class").split(" ")[1].split("-")[1]);
    let rowId = parseInt(rowCode);
    return [rowId, colId];
}

function evalFormula(cell) {
    let [rowId, colId] = codeToValue(cell);
    let formula = cellData[selectedSheet][rowId - 1][colId - 1].formula;
    console.log(formula);
    if (formula != "") {
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        let upStreamValue = [];
        for (let i in upStream) {
            let [calRowId, calColId] = codeToValue(upStream[i]);
            let value;
            if (cellData[selectedSheet][calRowId - 1][calColId - 1].text == "") {
                value = "0";
            }
            else {
                value = cellData[selectedSheet][calRowId - 1][calColId - 1].text;
            }
            upStreamValue.push(value);
            console.log(upStreamValue);
            formula = formula.replace(upStream[i], upStreamValue[i]);
        }
        cellData[selectedSheet][rowId - 1][colId - 1].text = eval(formula);
        loadCurrentSheet();
    }
    let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
    for (let i = downStream.length - 1; i >= 0; i--) {
        evalFormula(downStream[i]);
    }
}
