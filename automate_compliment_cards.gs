// retrieve spreadsheet file id from url
// https://docs.google.com/spreadsheets/d/<SPREADSHEET FILE ID>/edit
const DATA_SPREADSHEET_ID = '<SPREADSHEET FILE ID>';
const SHEET_NAME = 'Form Responses 1';

// configure column mapping
const DATA_SPREADSHEET_COLUMNS = {
    'FROM': 1, // COLUMN B
    'TO': 2, // COLUMN C
    'MESSAGE': 3, // COLUMN D
};

// retrieve template slides file id from url
// https://docs.google.com/presentation/d/<TEMPLATE SLIDES FILE ID>/edit
const TEMPLATE_PRESENTATION_ID = '<TEMPLATE SLIDES FILE ID>';

// use function GET_LAYOUTS() to get the correct master slides
const LAYOUT_IDS = [
    'p12', 'g8a7f226702_0_11', 'g8a7f226702_2_160',
    'g8a7f226702_2_167', 'g8a7f226702_2_173', 'g8a7f226702_2_179',
    'g8a7f226702_2_185', 'g8a7f226702_2_207', 'g8a7f226702_2_213',
    'g8a7f226702_2_219', 'g8a7f226702_2_140', 'g8a7f226702_2_153'
];

function GET_LAYOUTS() {
    const layouts = SlidesApp.openById(TEMPLATE_PRESENTATION_ID).getLayouts();
    Logger.log(layouts.map(function (layout) {
        return layout.getLayoutName() + ': ' + layout.getObjectId();
    }).join('\n'));
}

function GENERATE() {
    // array of requests to be made to Slides API
    const REQUESTS = [];

    // load data from Sheets API
    const SPREADSHEET = SpreadsheetApp.openById(DATA_SPREADSHEET_ID);
    const SHEET = SPREADSHEET.getSheetByName(SHEET_NAME);

    let ROWS = SHEET
        .getDataRange()
        .getValues();

    // sort rows by last name then by first name
    ROWS.sort(function (a, b) {
        const [A_FIRST, A_LAST] = rsplit(a[DATA_SPREADSHEET_COLUMNS['TO']]);
        const A_NAME = (A_LAST + "\n" + A_FIRST);
        const [B_FIRST, B_LAST] = rsplit(b[DATA_SPREADSHEET_COLUMNS['TO']]);
        const B_NAME = (B_LAST + "\n" + B_FIRST);
        if (A_NAME < B_NAME) {
            return -1;
        }
        if (A_NAME > B_NAME) {
            return 1;
        }
        return 0;
    });

    // move teachers to the end
    const TEACHERS = ROWS
        .filter(ROW => ROW[DATA_SPREADSHEET_COLUMNS['TO']]
            .indexOf(['Mr.', 'Ms.', 'Mrs.']) > -1);
    ROWS = ROWS
        .filter(ROW => ROW[DATA_SPREADSHEET_COLUMNS['TO']]
            .indexOf(['Mr.', 'Ms.', 'Mrs.']) < 0);
    ROWS.push(...TEACHERS);

    // generate requests to Slides API
    for (let i = 0; i < ROWS.length; ++i) {
        const ROW = ROWS[i];
        const FROM = ROW[DATA_SPREADSHEET_COLUMNS['FROM']];
        const TO = ROW[DATA_SPREADSHEET_COLUMNS['TO']];
        const MESSAGE = ROW[DATA_SPREADSHEET_COLUMNS['MESSAGE']];

        REQUESTS.push({
            createSlide: {
                objectId: 'COMPLIMENT_CARD_' + i,
                slideLayoutReference: {
                    layoutId: LAYOUT_IDS[i % LAYOUT_IDS.length]
                },
                placeholderIdMappings: [{
                    layoutPlaceholder: {
                        "type": "TITLE",
                        "index": 0
                    },
                    "objectId": 'FROM_' + i,
                },
                {
                    layoutPlaceholder: {
                        "type": "SUBTITLE",
                        "index": 0
                    },
                    "objectId": 'TO__' + i,
                },
                {
                    layoutPlaceholder: {
                        "type": "BODY",
                        "index": 0
                    },
                    "objectId": 'MESSAGE_' + i,
                },
                ],
            }
        });

        REQUESTS.push({
            insertText: {
                objectId: 'FROM_' + i,
                text: FROM,
            }
        });
        REQUESTS.push({
            insertText: {
                objectId: 'TO__' + i,
                text: TO,
            }
        });
        REQUESTS.push({
            insertText: {
                objectId: 'MESSAGE_' + i,
                text: MESSAGE,
            }
        });
    }
    // duplicate template presentation with Drive API.
    const PRESENTATION_COPY = Drive.Files.copy({
        title: '19.3.03 Automated Compliment Cards',
        parents: [{
            id: 'root'
        }]
    }, TEMPLATE_PRESENTATION_ID);

    // execute the requests on the copied presentation
    Slides.Presentations.batchUpdate({
        requests: REQUESTS
    }, PRESENTATION_COPY.id);
}

function rsplit(string, separator = " ") {
    const index = string.lastIndexOf(separator);
    return [string.substring(0, index), string.substring(index + 1)];
}
