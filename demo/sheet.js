const tfx = {};

window.tfx = tfx;

const data = [
    ['', 'Ford', 'Tesla', 'Toyota', 'Honda'],
    ['2017', 10, 11, '2017-01-01', 13],
    ['2018', 20, 11, '2018-12-12', 13],
    ['2018', 30, 11, '12-12-2018', 13],
    ['2019', 40, 15, '2019-30-30', 13]
];

window.onload = () => {
    const container = document.getElementById('sheet');
    const hot = new Handsontable(container, {
        data: data,
        rowHeaders: true,
        colHeaders: true,
        filters: true,
        dropdownMenu: true,
        columnSorting: true,
        outsideClickDeselects: false,
        licenseKey: 'non-commercial-and-evaluation'
    });
    window.Handsontable = Handsontable;
    window.hot = hot;

    window.addEventListener("message", function (event) {
        // We only accept messages from ourselves
        if (event.source !== window)
            return;

        if (event.data.type && (event.data.type === "FROM_SANAZ")) {
            console.log("App received: " + JSON.stringify(event.data));
            const [command, ...args] = event.data.command;
            if (tfx[command]) {
                tfx[command](...args)
            }
        }
    }, false);
};

const getLastRowNumber = () => window.hot.getData().length;
const getLastColumnLetter = () => String.fromCharCode('A'.charCodeAt(0) + window.hot.getData()[0].length - 1)

function getColumnIndex(columnLetter) {
    const charCode = columnLetter.charCodeAt(0);
    const charCodeForA = 'a'.charCodeAt(0);
    const charCodeForCapitalA = 'A'.charCodeAt(0);
    return charCode >= charCodeForA ? charCode - charCodeForA : charCode - charCodeForCapitalA;
}

function isCell(element) {
    return Number.isInteger(element[1]);
}

function getRowIndex(rowLetter) {
    return parseInt(rowLetter) - 1;
}

function parseCell(cell) {
    let [columnLetter, ...rowLetter] = cell;
    rowLetter = rowLetter.join('');
    return [getRowIndex(rowLetter), getColumnIndex(columnLetter)];
}

const rangeSeparator = ':';

function isRange(element) {
    return element.indexOf(rangeSeparator) > 0;
}

function parseRange(element) {
    let start, end;
    if (isRange(element)) {
        [start, end] = element.split(rangeSeparator);
    } else if (isCell(element)) {
        start = element;
        end = element;
    } else { // is row number or is columnLetter
        if (Number.isInteger(parseInt(element))) {
            element = parseInt(element);
            start = `A${element}`;
            end = `${getLastColumnLetter()}${element}`;
        } else {
            start = `${element}1`;
            end = `${element}${getLastRowNumber()}`
        }
    }
    start = parseCell(start);
    end = parseCell(end);
    return [...start, ...end];
}

function parseSelection(selection) {
    return selection.split(',').map(element => parseRange(element));
}

tfx.select = selection => {
    // e.g. tfx.select('A1:B2,B3,D2:E2')
    window.hot.selectCells(parseSelection(selection));
};

tfx.format = (selection, type) => {
    if (!selection) {
        return;
    }
    selection = selection === 'current' ? window.hot.getSelected() : parseSelection(selection);
    const typeConfigs = {
        'date': {
            correctFormat: true,
            dateFormat: 'MM/DD/YYYY'
        },
        'text': {
            validator: undefined
        }
    };
    for (const [rowStart, columnStart, rowEnd, columnEnd] of selection) {
        for (let row = rowStart; row <= rowEnd; row++) {
            for (let column = columnStart; column <= columnEnd; column++) {
                window.hot.setCellMetaObject(row, column, {
                    validator: type,
                    renderer: type,
                    editor: type,
                    ...typeConfigs[type]
                });
            }
        }
    }
    window.hot.validateCells();
    window.hot.render();
};

const alter = type => (element, position) => { 
    position = position === 'before' ? 0 : 1;
    let action, index;
    if (Number.isInteger(parseInt(element))) {
        action = `${type}_row`;
        index = parseInt(element) - 1 + position;
    } else {
        action = `${type}_col`;
        index = getColumnIndex(element) + position;
    }

    window.hot.alter(action, index);
};

tfx.insert = alter('insert');
tfx.delete = element => alter('remove')(element, 'before');

tfx.sort = (columnLetter, order='asc') => window.hot.getPlugin('columnSorting').sort({ column: getColumnIndex(columnLetter), sortOrder: order });
