// RETRIEVE FILE IDs FROM URL AND PASTE HERE.
const DATA_SPREADSHEET_ID = '[REDACTED]';
const DATA_SPREADSHEET_COLUMNS = {
    'FROM': 1, //COLUMN B
    'TO': 2, //COLUMN C
    'MESSAGE': 3, //COLUMN D
};
const TEMPLATE_PRESENTATION_ID = '[REDACTED]';
const LAYOUT_IDS = ['p12', 'g8a7f226702_0_11', 'g8a7f226702_2_160', 'g8a7f226702_2_167', 'g8a7f226702_2_173', 'g8a7f226702_2_179', 'g8a7f226702_2_185', 'g8a7f226702_2_207', 'g8a7f226702_2_213', 'g8a7f226702_2_219', 'g8a7f226702_2_140', 'g8a7f226702_2_153'];

function GET_LAYOUTS() {
    var layouts = SlidesApp.openById(TEMPLATE_PRESENTATION_ID).getLayouts();
    Logger.log(layouts.map(function(layout) {
        return layout.getLayoutName() + '|' + layout.getObjectId();
    }).join('\n'));
}

function GENERATE() {
    // ARRAY of REQUESTS to be made to Slides API
    const REQUESTS = [];

    // LOAD ROWS from Sheets API
    let ROWS = SpreadsheetApp.openById(DATA_SPREADSHEET_ID).getRange('Form Responses 1!A192:D208')
        .getValues();

    //SORT ROWS BY LAST_NAME THEN BY FIRST_NAME
    ROWS.sort(function(a, b) {
        var A_NAME = (a[DATA_SPREADSHEET_COLUMNS['TO']].split(" ")[1] + "\n" + a[DATA_SPREADSHEET_COLUMNS['TO']].split(" ")[0]);
        var B_NAME = (b[DATA_SPREADSHEET_COLUMNS['TO']].split(" ")[1] + "\n" + b[DATA_SPREADSHEET_COLUMNS['TO']].split(" ")[0]);
        if (A_NAME < B_NAME) {
            return -1;
        }
        if (A_NAME > B_NAME) {
            return 1;
        }
        return 0;
    });

    const TEACHERS = ROWS.filter(ROW => ROW[DATA_SPREADSHEET_COLUMNS['TO']].indexOf(['Mr.', 'Ms.', 'Mrs.']) > -1);
    ROWS = ROWS.filter(ROW => ROW[DATA_SPREADSHEET_COLUMNS['TO']].indexOf(['Mr.', 'Ms.', 'Mrs.']) < 0);
    ROWS.push(...TEACHERS);

    for (var i = 0; i < ROWS.length; i++) {
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
                        "objectId": 'TOxx_' + i,
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
                objectId: 'TOxx_' + i,
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
    // DUPLICATE TEMPLATE_PRESENTATION to PRESENTATION_COPY with Drive API.
    const PRESENTATION_COPY = Drive.Files.copy({
        title: '19.3.03 Automated Compliment Cards',
        parents: [{
            id: 'root'
        }]
    }, TEMPLATE_PRESENTATION_ID);

    // EXECUTE the REQUESTS on PRESENTATION_COPY
    Slides.Presentations.batchUpdate({
        requests: REQUESTS
    }, PRESENTATION_COPY.id);
}