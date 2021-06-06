# osagrad2020
Google App Scripts used to generate the banquet slideshow and automate
compliment cards Google Slides for OSA's Class of 2020.

## Generate Banquet Slideshow
Used to generate the banquet Google Slides slideshow for the Class of 2020
from our Google Form spreadsheet data using a Google Slides template.
The original spreadsheet data and slides template are not altered.

### Usage
1. Create a Google Slide to be used as the template.
    1. Create at custom grad slide layout in `View > Theme Builder`
        1. Insert Placeholders via `Insert > Placeholder`
            1. Title placeholders are used for `Name`
            1. Subtitle placeholders are used for `Quote`
1. Edit `generate_banquet_slideshow.gs`
    1. Replace `<SPREADSHEET FILE ID>` and `<TEMPLATE SLIDES FILE ID>` with the
    file ids respectively.
    1. Replace `<GRAD PHOTOTS FOLDER ID>` with the folder id or null if there
    are no photos.
    1. Update `SHEET_NAME` and `DATA_SPREADSHEET_COLUMNS` as appropriate.
    1. Run the `GET_PLACEHOLDERS()` function and update the
    `CUSTOM_GRAD_SLIDE_LAYOUT` appropriately.
    1. Grant authorization to the script when prompted.
    1. Run the `GENERATE()` function.
1. The generated Google Slides banquet slideshow is outputted to your Google
Drive root folder.

## Automate Compliment Cards
Used to automatically convert compliment cards from our Google Form
spreadsheet data to printable fancy compliment cards using a Google Slides
template. The original spreadsheet data and slides template are not altered.

### Usage
1. Create a Google Slide to be used as the template.
    1. Create at least one layout in `View > Theme Builder`
        1. Insert Placeholders via `Insert > Placeholder`
            1. Title placeholders are used for `From`
            1. Subtitle placeholders are used for `To`
            1. Body placeholders are used for `Message`
1. Edit `automate_compliment_cards.gs`
    1. Replace `<SPREADSHEET FILE ID>` and `<TEMPLATE SLIDES FILE ID>` with the
    file ids respectively.
    1. Update `SHEET_NAME` and `DATA_SPREADSHEET_COLUMNS` as appropriate.
    1. Run the `GET_LAYOUTS()` function and update the `LAYOUTS` array
    appropriately.
    1. Grant authorization to the script when prompted.
    1. Run the `GENERATE()` function.
1. The generated Google Slides compliment card is outputted to your Google
Drive root folder.
