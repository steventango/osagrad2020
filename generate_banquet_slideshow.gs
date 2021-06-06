// retrieve spreadsheet file id from url
// https://docs.google.com/spreadsheets/d/<SPREADSHEET FILE ID>/edit
const DATA_SPREADSHEET_ID = '<SPREADSHEET FILE ID>';
const SHEET_NAME = 'Form Responses 1';

// configure column mapping
const DATA_SPREADSHEET_COLUMNS = {
    'LAST_NAME': 2, // COLUMN C
    'FIRST_NAME': 3, // COLUMN D
    'QUOTE': 4, // COLUMN E
    'BABY_PHOTO_URL': 5 // COLUMN F
};

// retrieve grad photos folder id from url
// https://drive.google.com/drive/u/0/folders/<GRAD PHOTOTS FOLDER ID>
const GRAD_PHOTOS_FOLDER_ID = '<GRAD PHOTOTS FOLDER ID>';

// retrieve template slides file id from url
// https://docs.google.com/presentation/d/<TEMPLATE SLIDES FILE ID>/edit
const TEMPLATE_PRESENTATION_ID = '<TEMPLATE SLIDES FILE ID>';
const CUSTOM_GRAD_SLIDE_LAYOUT = 'g511f3991ce_0_612';

function GET_PLACEHOLDERS() {
    const layouts = SlidesApp.openById(TEMPLATE_PRESENTATION_ID).getLayouts();
    Logger.log(layouts.map(function (layout) {
        return layout.getLayoutName() + ': ' + layout.getObjectId();
    }).join('\n'));

    const layout = layouts.filter(function (layout) {
        return layout.getObjectId() == 'g511f3991ce_0_612' ? 1 : 0;
    })[0];

    const placeholders = layout.getPlaceholders().map(function (placeholder) {
        placeholder.getTitle();
    }).join();
    Logger.log(placeholders);

    Logger.log(layout.getShapes()[1].getObjectId());
}

function GENERATE() {
    // ARRAY of REQUESTS to be made to Slides API
    const REQUESTS = [];

    // load data from Sheets API
    const SPREADSHEET = SpreadsheetApp.openById(DATA_SPREADSHEET_ID);
    const SHEET = SPREADSHEET.getSheetByName(SHEET_NAME);

    let ROWS = SHEET
        .getDataRange()
        .getValues()
        // ignore first row
        .slice(1)
        .filter(function (ROW) {
            return ROW[DATA_SPREADSHEET_COLUMNS['LAST_NAME']].length > 0;
        });

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

    for (let i = 0; i < ROWS.length; i++) {
        const ROW = ROWS[i];
        const LAST_NAME = ROW[DATA_SPREADSHEET_COLUMNS['LAST_NAME']];
        const FIRST_NAME = ROW[DATA_SPREADSHEET_COLUMNS['FIRST_NAME']];
        const NAME = (LAST_NAME + "\n" + FIRST_NAME).toUpperCase();
        const QUOTE = ROW[DATA_SPREADSHEET_COLUMNS['QUOTE']];
        const BABY_PHOTO_URL = ROW[DATA_SPREADSHEET_COLUMNS['BABY_PHOTO_URL']];

        // create baby slide
        REQUESTS.push({
            createSlide: {
                objectId: 'BABY_SLIDE' + i,
                slideLayoutReference: {
                    predefinedLayout: 'BLANK'
                }
            }
        });

        // insert baby photo url on baby slide
        if (BABY_PHOTO_URL) {
            const BABY_PHOTO_FILE = DriveApp.getFileById(BABY_PHOTO_URL.match(/[-\w]{25,}/));
            const BABY_PHOTO_FILE_TYPE = BABY_PHOTO_FILE.getMimeType();
            if (['image/jpeg', 'image/png'].indexOf(BABY_PHOTO_FILE_TYPE) < 0) {
                Logger.log('Invalid baby photo file type: ' + BABY_PHOTO_FILE_TYPE + ' | URL: ' + BABY_PHOTO_URL);
            } else if (BABY_PHOTO_FILE.getSize() > 10e6) {
                Logger.log('Baby photo file size larger than 10 MB: ' + ' | URL: ' + BABY_PHOTO_URL);
            } else {
                // temporarily set sharing to DriveApp.Access.ANYONE_WITH_LINK
                BABY_PHOTO_FILE.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                REQUESTS.push({
                    createImage: {
                        url: "https://drive.google.com/uc?export=download&id=" + BABY_PHOTO_FILE.getId(),
                        elementProperties: {
                            pageObjectId: 'BABY_SLIDE' + i,
                            size: {
                                width: {
                                    magnitude: 720,
                                    unit: "PT"
                                },
                                height: {
                                    magnitude: 405,
                                    unit: "PT"
                                }
                            },
                            transform: {
                                scaleX: 1,
                                scaleY: 1,
                                translateX: 0,
                                translateY: 0,
                                unit: "PT"
                            }
                        }
                    }
                });
            }
        }

        // create grad slide
        REQUESTS.push({
            createSlide: {
                objectId: 'GRAD_SLIDE' + i,
                slideLayoutReference: {
                    layoutId: CUSTOM_GRAD_SLIDE_LAYOUT
                },
                placeholderIdMappings: [{
                    layoutPlaceholder: {
                        "type": "TITLE",
                        "index": 0
                    },
                    "objectId": 'NAME' + i,
                },
                {
                    layoutPlaceholder: {
                        "type": "SUBTITLE",
                        "index": 0
                    },
                    "objectId": 'QUOTE' + i,
                },
                ],
            }
        });

        REQUESTS.push({
            insertText: {
                objectId: 'NAME' + i,
                text: NAME,
            }
        });

        REQUESTS.push({
            insertText: {
                objectId: 'QUOTE' + i,
                text: QUOTE,
            }
        });

        REQUESTS.push({
            createShape: {
                objectId: 'GRAD_PHOTO_TEMPLATE_OBJECT_' + i,
                elementProperties: {
                    pageObjectId: 'GRAD_SLIDE' + i,
                    size: {
                        width: {
                            magnitude: 360,
                            unit: "PT"
                        },
                        height: {
                            magnitude: 405,
                            unit: "PT"
                        }
                    },
                    transform: {
                        scaleX: 1,
                        scaleY: 1,
                        translateX: 360,
                        translateY: 0,
                        unit: "PT"
                    }
                },
                shapeType: 'RECTANGLE'
            }
        });

        REQUESTS.push({
            insertText: {
                objectId: 'GRAD_PHOTO_TEMPLATE_OBJECT_' + i,
                text: 'GRAD_PHOTO_TEMPLATE_OBJECT_' + i,
            }
        });

        // insert grad photo
        if (GRAD_PHOTOS_FOLDER_ID) {
            let GRAD_PHOTO_FILE;
            const GRAD_PHOTOS = DriveApp.searchFiles("'" + GRAD_PHOTOS_FOLDER_ID + "' in parents and title contains '" + LAST_NAME.toLowerCase() + '_' + FIRST_NAME.toLowerCase() + "' and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')");

            if (GRAD_PHOTOS.length > 1) {
                Logger.log('MULTIPLE GRAD PHOTOS: #GRAD_SLIDE' + i);
            } else if (GRAD_PHOTOS.length < 1) {
                Logger.log('NO GRAD PHOTO: #GRAD_SLIDE' + i);
            } else {
                if (GRAD_PHOTOS) {
                    while (GRAD_PHOTOS.hasNext()) {
                        GRAD_PHOTO_FILE = GRAD_PHOTOS.next();
                        if (GRAD_PHOTO_FILE.getSize() > 1e6) {
                            // Files larger than 1 MB need to be uploaded manually.
                            Logger.log('Grad photo file size larger than 1 MB: ' + GRAD_PHOTO_FILE.getname());
                        } else {
                            // temporarily set sharing to DriveApp.Access.DOMAIN_WITH_LINK
                            GRAD_PHOTO_FILE.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
                            REQUESTS.push({
                                replaceAllShapesWithImage: {
                                    imageUrl: "https://drive.google.com/uc?export=download&id=" + GRAD_PHOTO_FILE.getId(),
                                    imageReplaceMethod: 'CENTER_INSIDE',
                                    pageObjectIds: ['GRAD_SLIDE' + i],
                                    containsText: {
                                        text: 'GRAD_PHOTO_TEMPLATE_OBJECT_' + i,
                                        matchCase: false
                                    }
                                }
                            });
                            GRAD_PHOTO_FILE = null;
                        }
                    }
                }
            }
        }
    }

    // duplicate template presentation with Drive API.
    const PRESENTATION_COPY = Drive.Files.copy({
        title: '6.1.03 Automated Banquet Slideshow',
        parents: [{
            id: 'root'
        }]
    }, TEMPLATE_PRESENTATION_ID);

    // execute the requests on the copied presentation
    Slides.Presentations.batchUpdate({
        requests: REQUESTS
    }, PRESENTATION_COPY.id);
}

function UNSHARE() {
    const SPREADSHEET = SpreadsheetApp.openById(DATA_SPREADSHEET_ID);
    const SHEET = SPREADSHEET.getSheetByName(SHEET_NAME);

    let ROWS = SHEET
        .getDataRange()
        .getValues();

    // unshare baby photos
    for (let i = 0; i < ROWS.length; i++) {
        const ROW = ROWS[i];
        const BABY_PHOTO_URL = ROW[DATA_SPREADSHEET_COLUMNS['BABY_PHOTO_URL']];
        if (BABY_PHOTO_URL) {
            const BABY_PHOTO_FILE = DriveApp.getFileById(BABY_PHOTO_URL.match(/[-\w]{25,}/));
            BABY_PHOTO_FILE.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);
        }
    }

    // unshare grad photos
    const GRAD_PHOTOS = DriveApp.getFolderById(GRAD_PHOTOS_FOLDER_ID).getFiles();
    while (GRAD_PHOTOS.hasNext()) {
        GRAD_PHOTOS.next().setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    }
}
