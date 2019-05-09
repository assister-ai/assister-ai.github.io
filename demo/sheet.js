window.onload = () => {
    const data = [
        ['', 'Ford', 'Tesla', 'Toyota', 'Honda'],
        ['2017', 10, 11, '2017-01-01', 13],
        ['2018', 20, 11, '2018-12-12', 13],
        ['2018', 30, 11, '12-12-2018', 13],
        ['2019', 40, 15, '2019-30-30', 13]
    ];

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
};

const tfx = {};

window.tfx = tfx;

function getColumnIndex(columnLetter) {
    const charCode = columnLetter.charCodeAt(0);
    const charCodeForA = 'a'.charCodeAt(0);
    const charCodeForCapitalA = 'A'.charCodeAt(0);
    return charCode >= charCodeForA ? charCode - charCodeForA : charCode - charCodeForCapitalA;
}

function getRowIndex(rowLetter) {
    return parseInt(rowLetter) - 1;
}

function parseCell(cell) {
    let [columnLetter, ...rowLetter] = cell;
    rowLetter = rowLetter.join('');
    return [getRowIndex(rowLetter), getColumnIndex(columnLetter)];
}

function isRange(element) {
    return element.indexOf(':') > 0;
}

function parseRange(element) {
    let start, end;
    if (isRange(element)) {
        [start, end] = element.split(':');
    } else {
        start = element;
        end = element;
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
    if (Number.isInteger(element)) {
        action = `${type}_row`;
        index = element - 1 + position;
    } else {
        action = `${type}_col`;
        index = getColumnIndex(element) + position;
    }

    window.hot.alter(action, index);
};

tfx.insert = alter('insert');
tfx.delete = element => alter('remove')(element, 'before');

tfx.sort = (columnLetter, order='asc') => window.hot.getPlugin('columnSorting').sort({ column: getColumnIndex(columnLetter), sortOrder: order });
