// RETRIEVE FILE IDs FROM URL AND PASTE HERE.
const DATA_SPREADSHEET_ID = '[REDACTED]';
const DATA_SPREADSHEET_COLUMNS = {
    'LAST_NAME': 2, //COLUMN C
    'FIRST_NAME': 3, //COLUMN D
    'QUOTE': 4, //COLUMN E
    'BABY_PHOTO_URL': 5 //COLUMN F
}
const TEMPLATE_PRESENTATION_ID = '[REDACTED]';
const GRAD_PHOTOS_FOLDER_ID = null;
const CUSTOM_GRAD_SLIDE_LAYOUT = 'g511f3991ce_0_612';

function GET_PLACEHOLDERS() {
    var layouts = SlidesApp.openById(TEMPLATE_PRESENTATION_ID).getLayouts();
    Logger.log(layouts.map(function(layout) {
        return layout.getLayoutNAME() + '|' + layout.getObjectId()
    }).join('\n'));

    var layout = layouts.filter(function(layout) {
        return layout.getObjectId() == 'g511f3991ce_0_612' ? 1 : 0
    })[0];
    var placeholders = layout.getPlaceholders().map(function(placeholder) {
        placeholder.getTitle()
    }).join();
    Logger.log(placeholders);

    Logger.log(layout.getShapes()[1].getObjectId());
}

function GENERATE() {
    // ARRAY of REQUESTS to be made to Slides API
    const REQUESTS = [];

    // LOAD ROWS from Sheets API
    const ROWS = SpreadsheetApp.openById(DATA_SPREADSHEET_ID).getRange('Form Responses 1!A2:F120')
        .getValues()
        .filter(function(ROW) {
            return ROW[DATA_SPREADSHEET_COLUMNS['LAST_NAME']].length > 0
        });

    //SORT ROWS BY LAST_NAME THEN BY FIRST_NAME
    ROWS.sort(function(a, b) {
        var A_NAME = (a[DATA_SPREADSHEET_COLUMNS['LAST_NAME']] + "\n" + a[DATA_SPREADSHEET_COLUMNS['FIRST_NAME']]).toUpperCase();
        var B_NAME = (b[DATA_SPREADSHEET_COLUMNS['LAST_NAME']] + "\n" + b[DATA_SPREADSHEET_COLUMNS['FIRST_NAME']]).toUpperCase();
        if (A_NAME < B_NAME) {
            return -1;
        }
        if (A_NAME > B_NAME) {
            return 1;
        }
        return 0;
    });

    for (var i = 0; i < ROWS.length; i++) {
        const ROW = ROWS[i];
        const LAST_NAME = ROW[DATA_SPREADSHEET_COLUMNS['LAST_NAME']];
        const FIRST_NAME = ROW[DATA_SPREADSHEET_COLUMNS['FIRST_NAME']];
        const NAME = (LAST_NAME + "\n" + FIRST_NAME).toUpperCase();
        const QUOTE = ROW[DATA_SPREADSHEET_COLUMNS['QUOTE']];
        const BABY_PHOTO_URL = ROW[DATA_SPREADSHEET_COLUMNS['BABY_PHOTO_URL']];

        //CREATE BABY_SLIDE
        REQUESTS.push({
            createSlide: {
                objectId: 'BABY_SLIDE' + i,
                slideLayoutReference: {
                    predefinedLayout: 'BLANK'
                }
            }
        });

        //INSERT BABY_PHOTO ON BABY_SLIDE
        if (BABY_PHOTO_URL) {
            const BABY_PHOTO_FILE = DriveApp.getFileById(BABY_PHOTO_URL.match(/[-\w]{25,}/));
            const BABY_PHOTO_FILE_TYPE = BABY_PHOTO_FILE.getMimeType();
            if (['image/jpeg', 'image/png'].indexOf(BABY_PHOTO_FILE_TYPE) < 0) {
                Logger.log('Invalid baby photo file type: ' + BABY_PHOTO_FILE_TYPE + ' | URL: ' + BABY_PHOTO_URL);
            } else if (BABY_PHOTO_FILE.getSize() > 10e6) {
                Logger.log('Baby photo file size larger than 10 MB: ' + ' | URL: ' + BABY_PHOTO_URL);
            } else {
                // TEMPORARILY SET SHARING to DriveApp.Access.ANYONE_WITH_LINK
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

        // CREATE GRAD_SLIDE
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

        // INSERT NAME ON GRAD_SLIDE
        REQUESTS.push({
            insertText: {
                objectId: 'NAME' + i,
                text: NAME,
            }
        });

        //INSERT QUOTE ON GRAD_SLIDE
        REQUESTS.push({
            insertText: {
                objectId: 'QUOTE' + i,
                text: QUOTE,
            }
        });

        //CREATE GRAD_PHOTO_TEMPLATE_OBJECT
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

        // INSERT GRAD PHOTO
        if (GRAD_PHOTOS_FOLDER_ID) {
            var GRAD_PHOTO_FILE;
            var GRAD_PHOTOS = DriveApp.searchFiles("'" + GRAD_PHOTOS_FOLDER_ID + "' in parents and title contains '" + LAST_NAME.toLowerCase() + '_' + FIRST_NAME.toLowerCase() + "' and (mimeType contains 'image/jpeg' or mimeType contains 'image/png')");

            if (GRAD_PHOTOS.length > 1) {
                Logger.log('MULTIPLE GRAD PHOTOS: #GRAD_SLIDE' + i)
            } else if (GRAD_PHOTOS.length < 1) {
                Logger.log('NO GRAD PHOTO: #GRAD_SLIDE' + i)
            } else {
                if (GRAD_PHOTOS) {
                    while (GRAD_PHOTOS.hasNext()) {
                        GRAD_PHOTO_FILE = GRAD_PHOTOS.next();
                        if (GRAD_PHOTO_FILE.getSize() > 1e6) {
                            // Files larger than 1 MB need to be uploaded manually.
                            Logger.log('Grad photo file size larger than 1 MB: ' + GRAD_PHOTO_FILE.getname());
                        } else {
                            // TEMPORARILY SET SHARING to DriveApp.Access.DOMAIN_WITH_LINK
                            GRAD_PHOTO_FILE.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW)
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

    // DUPLICATE TEMPLATE_PRESENTATION to PRESENTATION_COPY with Drive API.
    const PRESENTATION_COPY = Drive.Files.copy({
        title: '6.1.03 Automated Banquet Slideshow',
        parents: [{
            id: 'root'
        }]
    }, TEMPLATE_PRESENTATION_ID);

    // EXECUTE the REQUESTS on PRESENTATION_COPY
    Slides.Presentations.batchUpdate({
        requests: REQUESTS
    }, PRESENTATION_COPY.id);
}

function UNSHARE() {
    const ROWS = SpreadsheetApp.openById(DATA_SPREADSHEET_ID).getRange('Form Responses 1!A2:F120').getValues();
    for (var i = 0; i < ROWS.length; i++) {
        const ROW = ROWS[i];
        const BABY_PHOTO_URL = ROW[DATA_SPREADSHEET_COLUMNS['BABY_PHOTO_URL']];
        if (BABY_PHOTO_URL) {
            const BABY_PHOTO_FILE = DriveApp.getFileById(BABY_PHOTO_URL.match(/[-\w]{25,}/));
            BABY_PHOTO_FILE.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);
        }
        const GRAD_PHOTOS = DriveApp.getFolderById(GRAD_PHOTOS_FOLDER_ID).getFiles();
        while (GRAD_PHOTOS.hasNext()) {
            GRAD_PHOTOS.next().setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);;
        }
    }
}